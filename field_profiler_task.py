# -*- coding: utf-8 -*-
import random
import statistics
import numpy
from collections import Counter, OrderedDict
from datetime import datetime

from qgis.core import (QgsTask, QgsMessageLog, Qgis, QgsFeatureRequest, QgsMapLayer, QgsExpression, QgsExpressionContext, QgsExpressionContextUtils)
from qgis.PyQt.QtCore import pyqtSignal, QVariant, QDate, QDateTime, QTime

# Check for Scipy
SCIPY_AVAILABLE = False
try:
    from scipy import stats as scipy_stats
    SCIPY_AVAILABLE = True
except ImportError:
    scipy_stats = None

class StreamingStats:
    """
    Helper to calculate running statistics (count, min, max, mean, variance)
    using Welford's algorithm to avoid storing all values.
    """
    def __init__(self):
        self.count = 0
        self.min_val = float('inf')
        self.max_val = float('-inf')
        self.mean = 0.0
        self.M2 = 0.0  # For variance calculation

    def update(self, val):
        self.count += 1
        if val < self.min_val: self.min_val = val
        if val > self.max_val: self.max_val = val
        
        delta = val - self.mean
        self.mean += delta / self.count
        delta2 = val - self.mean
        self.M2 += delta * delta2

    def variance(self):
        if self.count < 2: return float('nan')
        return self.M2 / self.count # Population variance (or divide by count-1 for sample)

    def std_dev(self):
        if self.count < 2: return float('nan')
        return (self.M2 / self.count)**0.5 # Population Stdev matching original code

class ReservoirSampler:
    """
    Maintains a random sample of a stream of items using Reservoir Sampling.
    Used for approximate Median, Quantiles, Mode on massive datasets.
    """
    def __init__(self, size=500000):
        self.size = size
        self.reservoir = []
        self.count_seen = 0

    def update(self, item):
        self.count_seen += 1
        if len(self.reservoir) < self.size:
            self.reservoir.append(item)
        else:
            # Replace elements with gradually decreasing probability
            r = random.randint(0, self.count_seen - 1)
            if r < self.size:
                self.reservoir[r] = item

class FieldProfilerTask(QgsTask):
    """
    Background task for running field analysis.
    """
    # Signal emitted when analysis is complete with the results dictionary
    analysisFinished = pyqtSignal(dict)
    
    # Constants for memory protection
    MAX_EXACT_VALUES = 1000000 # Switch to streaming/sampling after this many items per field
    
    def __init__(self, layer, field_names, config_options, selected_ids=None, validation_rules=None):
        description = f"Profiling {len(field_names)} fields on {layer.name()}"
        super().__init__(description, QgsTask.CanCancel)
        
        self.layer = layer
        self.field_names = field_names
        self.config_options = config_options
        self.selected_ids = selected_ids
        self.validation_rules_str = validation_rules if validation_rules else []
        self.exception = None
        self.results = OrderedDict()
        self.conversion_error_fids = {}
        self.non_printable_fids = {}

    def run(self):
        """
        Main execution method running in a background thread.
        """
        try:
            QgsMessageLog.logMessage(f"Starting analysis for {self.layer.name()}", "FieldProfiler", Qgis.Info)
            
            # 1. Setup Collection Structures
            qgs_fields = self.layer.fields()
            field_metadata = {}
            collectors = {}

            for fname in self.field_names:
                idx = qgs_fields.lookupField(fname)
                if idx == -1: 
                    self.results[fname] = {'Error': 'Field not found'}
                    continue
                
                fobj = qgs_fields.field(idx)
                meta = {'index': idx, 'object': fobj, 'type': fobj.type()}
                field_metadata[fname] = meta
                
                # Initialize collector for this field
                collectors[fname] = {
                    'null_count': 0,
                    'streaming_stats': StreamingStats() if fobj.isNumeric() else None,
                    'reservoir': ReservoirSampler(self.MAX_EXACT_VALUES),
                    'conversion_errors': 0,
                    'conversion_error_fids': [],  # Limit size to 1000
                    'non_printable_fids': [],  # Track features with non-printable chars
                    'is_exact': True,  # Flag if we are still storing all values
                    'original_variants_reservoir': ReservoirSampler(self.MAX_EXACT_VALUES) if fobj.type() in [QVariant.Date, QVariant.DateTime] else None
                }

            # Setup Validation Expressions
            validation_expressions = []
            validation_fail_counts = []
            if self.validation_rules_str:
                context = QgsExpressionContext()
                context.appendScope(QgsExpressionContextUtils.globalScope())
                context.appendScope(QgsExpressionContextUtils.layerScope(self.layer))
                
                for rule in self.validation_rules_str:
                    exp = QgsExpression(rule)
                    if exp.hasParserError():
                        self.results['_validation_error'] = f"Parser Error in rule: {rule} - {exp.parserErrorString()}"
                        # Continue or bust? Let's continue with other rules? Or just skip this one.
                        # But simpler to just proceed and fail evaluation.
                    
                    # exp.prepare(context) # Optimization
                    validation_expressions.append(exp)
                    validation_fail_counts.append(0)
            
            # Setup Correlation Reservoir (Row-based)

            # Setup Correlation Reservoir (Row-based)
            self.numeric_fields_for_corr = [fname for fname in self.field_names if field_metadata[fname]['object'].isNumeric()]
            self.row_reservoir = ReservoirSampler(self.MAX_EXACT_VALUES) if len(self.numeric_fields_for_corr) > 1 else None

            # 2. Configure Request
            request = QgsFeatureRequest()
            if self.selected_ids:
                request.setFilterFids(self.selected_ids)
            
            # Using NoGeometry to speed up if possible, unless spatial index needed? Not needed here.
            request.setFlags(QgsFeatureRequest.NoGeometry) 
            
            iterator = self.layer.getFeatures(request)
            total_count = len(self.selected_ids) if self.selected_ids else self.layer.featureCount()
            
            # 3. Iterate and Collect
            count = 0
            # Pre-calculate reusable lookups to avoid dict hashing inside loop if possible
            # But the overhead is minimal compared to QGIS reading.
            
            for feature in iterator:
                if self.isCanceled():
                    return False
                
                fid = feature.id()
                
                for fname, meta in field_metadata.items():
                    collector = collectors[fname]
                    val = feature[meta['index']]
                    
                    # Store original QVariant for dates (using reservoir if needed)
                    if meta['type'] in [QVariant.Date, QVariant.DateTime]:
                        # Only store if not null
                         if val is not None and not (hasattr(val, 'isNull') and val.isNull()):
                             collector['original_variants_reservoir'].update(val)

                    if val is None or (hasattr(val, 'isNull') and val.isNull()):
                        collector['null_count'] += 1
                    else:
                        # Value processing
                        # 1. Always update reservoir (handles switching to sample automatically)
                        collector['reservoir'].update(val)
                        
                        # 2. Check if we exceeded exact limit
                        if collector['is_exact'] and collector['reservoir'].count_seen > self.MAX_EXACT_VALUES:
                            collector['is_exact'] = False
                        
                        # 3. Numeric Specifics
                        if meta['object'].isNumeric():
                            try:
                                f_val = float(val)
                                collector['streaming_stats'].update(f_val)
                            except (ValueError, TypeError):
                                collector['conversion_errors'] += 1
                                if len(collector['conversion_error_fids']) < 1000: # Limit error storage
                                    collector['conversion_error_fids'].append(fid)
                        
                        # 4. String Specifics (Non-printable check)
                        elif meta['type'] == QVariant.String:
                            if self._has_non_printable_chars(val):
                                if len(collector['non_printable_fids']) < 1000:
                                    collector['non_printable_fids'].append(fid)

                # Validation Check
                if validation_expressions:
                    # Context update feature? 
                    # Simpler to use examine(feature) or evaluate(feature)
                    # To use prepared expression, we need context.setFeature(feature)
                    # But context is local to __init__ scope above? No, I defined it in run() logic block.
                    # Re-creating context here is slow. 
                    # Let's just use evaluate(feature) for simplicity in this iteration, optimization later if needed.
                    for i, exp in enumerate(validation_expressions):
                        try:
                            # Rule passes if True. Fails if False (or 0).
                            if not exp.evaluate(feature):
                                validation_fail_counts[i] += 1
                        except:
                            # If evaluation fails (e.g. type error), count as failure?
                            validation_fail_counts[i] += 1

                # Correlation Update
                if self.row_reservoir:
                    # Extract numeric values for this row
                    row_vals = []
                    valid_row = True
                    for nf_name in self.numeric_fields_for_corr:
                        # Direct access optimization? No, reusing feature[index] ok
                        # But we already read it!
                        # We didn't store it in a local var accessible here easily without re-reading or refactoring loop.
                        # Refactoring loop slightly to cache current row values mapping
                        # Actually, let's just re-read using the index. It's fast (in memory variant for feature).
                        val_n = feature[field_metadata[nf_name]['index']]
                        try:
                            # Strict: valid float
                            if val_n is not None and not (hasattr(val_n, 'isNull') and val_n.isNull()):
                                row_vals.append(float(val_n))
                            else:
                                valid_row = False; break # Skip rows with any nulls for correlation? Or handle NaN?
                                # Usually correlation ignores incomplete pairs. 
                                # Simpler to skip row or use NaN. Let's use NaN.
                                # row_vals.append(float('nan')) 
                                # But ReservoirSampler stores "items". 
                        except:
                            valid_row = False; break
                    
                    if valid_row:
                        self.row_reservoir.update(row_vals)

                count += 1
                if count % 1000 == 0:
                     self.setProgress((count / total_count) * 100 if total_count > 0 else 0)
            
            # 4. Finalize Analysis (Calculate Stats)
            self.setProgress(90)
            for fname, meta in field_metadata.items():
                if self.isCanceled(): return False
                
                col = collectors[fname]
                analyzed_count = count # approximated total analyzed
                non_null_count = (analyzed_count - col['null_count']) # This might be slightly off if features skipped? No, we iterate all.
                # Actually count_seen in reservoir matches non_null_count exactly
                non_null_count = col['reservoir'].count_seen

                # Base Results
                percent_null = (col['null_count'] / analyzed_count * 100) if analyzed_count > 0 else 0
                field_res = OrderedDict([
                    ('Null Count', col['null_count']),
                    ('% Null', f"{percent_null:.2f}%"),
                    ('Non-Null Count', non_null_count)
                ])
                
                # Copy stored error FIDs to result for selection
                if col['conversion_error_fids']:
                    field_res['_conversion_error_fids'] = col['conversion_error_fids']
                if col['non_printable_fids']:
                    field_res['_non_printable_fids'] = col['non_printable_fids']

                if not col['is_exact']:
                    field_res['Status (Method)'] = 'Approximated (Large Dataset)'

                # Dispatch analysis
                try:
                    stats = {}
                    if non_null_count == 0:
                        stats['Status'] = 'All Null or Empty'
                        # Add zeroed fields... (skipped for brevity, handled in UI or generic filler)
                    elif meta['object'].isNumeric():
                        stats = self._analyze_numeric(col, meta, non_null_count)
                    elif meta['type'] == QVariant.String:
                        stats = self._analyze_text(col, non_null_count)
                    elif meta['type'] in [QVariant.Date, QVariant.DateTime]:
                        stats = self._analyze_date(col, non_null_count)
                    else:
                        stats['Status'] = 'Not implemented'
                    
                    field_res.update(stats)
                    
                    # Hints
                    field_res['Data Type Mismatch Hint'] = self._generate_hints(meta, col, field_res)
                    
                except Exception as e:
                    field_res['Error'] = str(e)
                
                self.results[fname] = field_res

            # Calculate Global Correlation
            if self.row_reservoir and self.row_reservoir.reservoir:
                try:
                    # Convert to numpy array (N samples x M fields)
                    data_matrix = numpy.array(self.row_reservoir.reservoir)
                    if data_matrix.size > 0:
                        # Rowvar=False because rows are samples, cols are fields
                        corr_matrix = numpy.corrcoef(data_matrix, rowvar=False)
                        # Handle case where corrcoef returns scalar if only 2 vars? No, always matrix if multidim input?
                        # If 1 var, scalar. Checked > 1 above.
                        self.results['_global_correlation'] = {
                            'fields': self.numeric_fields_for_corr,
                            'matrix': corr_matrix.tolist()
                        }
                except Exception as e:
                    self.results['_global_correlation'] = {'Error': str(e)}

            # Finalize Validation Results
            if self.validation_rules_str:
                self.results['_validation_results'] = {
                    'rules': self.validation_rules_str,
                    'fail_counts': validation_fail_counts,
                    'total_checked': count
                }

            return True

        except Exception as e:
            self.exception = e
            QgsMessageLog.logMessage(f"Task Failed: {e}", "FieldProfiler", Qgis.Critical)
            return False

    def finished(self, result):
        """
        Called on main thread when task finishes.
        """
        if result:
            self.analysisFinished.emit(self.results)
        else:
            if self.exception:
                # Can notify user via signal if needed, or caller checks task status
                pass

    # --- Helper Analysis Methods ---

    def _analyze_numeric(self, col, meta, count):
        res = OrderedDict()
        res['Conversion Errors'] = col['conversion_errors']
        
        # Use Streaming Stats for basic stuff
        ss = col['streaming_stats']
        res['Min'] = ss.min_val if ss.count > 0 else float('nan')
        res['Max'] = ss.max_val if ss.count > 0 else float('nan')
        res['Range'] = ss.max_val - ss.min_val if ss.count > 0 else float('nan')
        res['Mean'] = ss.mean
        res['Sum'] = ss.mean * ss.count # inferred sum
        res['Stdev (pop)'] = ss.std_dev()
        res['CV %'] = (ss.std_dev() / ss.mean * 100) if ss.mean != 0 else float('nan')
        
        # Use Reservoir for advanced stats (Mode, Median, Quantiles)
        data_sample = numpy.array([float(x) for x in col['reservoir'].reservoir if x is not None], dtype=float)
        # Filter NaNs/Infs ??
        data_sample = data_sample[~numpy.isnan(data_sample)]
        
        # Mode
        modes_val = 'N/A'
        if len(data_sample) > 0:
             # Basic mode on sample
             if SCIPY_AVAILABLE:
                 m = scipy_stats.mode(data_sample, nan_policy='omit')
                 modes_val = m.mode if m.mode.size > 0 else 'N/A'
             else:
                 try: 
                     modes_val = statistics.multimode(data_sample.tolist())
                 except: modes_val = 'N/A'
        res['Mode(s)'] = modes_val
        
        # Quantiles
        res['Median'] = numpy.median(data_sample) if len(data_sample) > 0 else float('nan')
        if len(data_sample) > 0:
            res['Q1'] = numpy.percentile(data_sample, 25)
            res['Q3'] = numpy.percentile(data_sample, 75)
            res['IQR'] = res['Q3'] - res['Q1']
            
            # Outliers on sample
            lower = res['Q1'] - 1.5*res['IQR']
            upper = res['Q3'] + 1.5*res['IQR']
            outliers = data_sample[(data_sample < lower) | (data_sample > upper)]
            res['Outliers (IQR)'] = len(outliers)
            # Extrapolate outlier count if sampled? 
            # If is_exact=True, count is exact. If is_exact=False, this is sample count.
            if not col['is_exact']:
                # Roughly extrapolate
                factor = count / len(data_sample)
                res['Outliers (IQR)'] = f"{int(len(outliers) * factor)} (Est.)"
            
            res['Min Outlier'] = numpy.min(outliers) if len(outliers) > 0 else 'N/A'
            res['Max Outlier'] = numpy.max(outliers) if len(outliers) > 0 else 'N/A'
            res['% Outliers'] = (len(outliers)/len(data_sample) * 100) if len(data_sample)>0 else 0
        
        # Advanced options (Skew, Kurtosis) on sample
        if self.config_options.get('numeric_dist_shape') and SCIPY_AVAILABLE and len(data_sample) > 0:
            res['Skewness'] = scipy_stats.skew(data_sample)
            res['Kurtosis'] = scipy_stats.kurtosis(data_sample)
            if len(data_sample) >= 3 and len(data_sample) < 5000: # Shapiro is slow/valid for small n
                s, p = scipy_stats.shapiro(data_sample)
                res['Normality (Shapiro-Wilk p)'] = p
            else:
                res['Normality (Shapiro-Wilk p)'] = "N/A (N>5000)"

        # Histogram data for charts
        if len(data_sample) > 0:
            try:
                # Use Freedman-Diaconis rule for bin count if possible, otherwise default
                # We already calculated optimal bins in original code, let's reuse logic or simple auto
                hist, bin_edges = numpy.histogram(data_sample, bins='auto')
                res['_histogram_data'] = (hist.tolist(), bin_edges.tolist())
            except Exception:
                pass

        return res

    def _analyze_text(self, col, count):
        res = OrderedDict()
        sample = col['reservoir'].reservoir
        # Convert all to string
        str_sample = [str(x) for x in sample]
        
        # Approx stats on sample
        empty_count = str_sample.count('') # This is only on sample!
        
        # If sampling, we can't give exact Empty String count unless we tracked it separately in streaming
        # For simplicity, assuming sample is representative.
        
        res['Empty Strings'] = empty_count if col['is_exact'] else f"{int(empty_count * count / len(sample))} (Est.)"
        
        # Lengths
        non_empty = [s for s in str_sample if s]
        if non_empty:
            lengths = [len(s) for s in non_empty]
            res['Min Length'] = min(lengths)
            res['Max Length'] = max(lengths)
            res['Avg Length'] = sum(lengths)/len(lengths)
        
        # Top Values
        ctr = Counter(str_sample)
        top = ctr.most_common(self.config_options.get('limit_unique', 5))
        top_str = []
        for v, c in top:
            actual_c = c if col['is_exact'] else int(c * count / len(sample))
            top_str.append(f"'{v}': {actual_c}")
        res['Unique Values (Top)'] = "\n".join(top_str)
        if top:
            res['Unique Values (Top)_actual_first_value'] = top[0][0]
        
        # Raw top values for charts
        res['_top_values_raw'] = []
        for v, c in top:
            actual_c = c if col['is_exact'] else int(c * count / len(sample))
            res['_top_values_raw'].append( (str(v), actual_c) )

        return res

    def _analyze_date(self, col, count):
        from collections import Counter
        from datetime import datetime
        
        res = OrderedDict()
        sample = col['original_variants_reservoir'].reservoir
        
        if not sample:
            return {'Status': 'No data'}
        
        # Convert samples to Python datetime for analysis
        py_datetimes = []
        q_date_time_objects = []
        
        for val in sample:
            if val is None:
                continue
            
            py_dt = None
            q_obj = None
            
            if isinstance(val, QDateTime) and val.isValid():
                py_dt = val.toPyDateTime()
                q_obj = val
            elif isinstance(val, QDate) and val.isValid():
                py_dt = datetime(val.year(), val.month(), val.day())
                q_obj = val
            
            if py_dt and q_obj:
                py_datetimes.append(py_dt)
                q_date_time_objects.append(q_obj)
        
        if not py_datetimes:
            return {'Status': 'No valid date objects parsed'}
        
        # Min/Max Date
        min_d = min(py_datetimes)
        max_d = max(py_datetimes)
        is_datetime_field = any(isinstance(q, QDateTime) for q in q_date_time_objects)
        
        if is_datetime_field:
            res['Min Date'] = min_d.isoformat(sep=' ', timespec='seconds')
            res['Max Date'] = max_d.isoformat(sep=' ', timespec='seconds')
        else:
            res['Min Date'] = min_d.date().isoformat()
            res['Max Date'] = max_d.date().isoformat()
        
        # Common Years, Months, Days
        years = [d.year for d in py_datetimes]
        months = [d.month for d in py_datetimes]
        days_of_week = [d.weekday() for d in py_datetimes]  # Monday=0, Sunday=6
        
        day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        month_names = ["", "Jan", "Feb", "Mar", "Apr", "May", "Jun", 
                       "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
        
        res['Common Years'] = ", ".join([f"{yr}:{cnt}" for yr, cnt in Counter(years).most_common(3)])
        res['Common Months'] = ", ".join([f"{month_names[mo]}:{cnt}" for mo, cnt in Counter(months).most_common(3)])
        res['Common Days'] = ", ".join([f"{day_names[d]}:{cnt}" for d, cnt in Counter(days_of_week).most_common(3)])
        
        # Dates Before/After Today
        today = datetime.now().date()
        res['Dates Before Today'] = sum(1 for d in py_datetimes if d.date() < today)
        res['Dates After Today'] = sum(1 for d in py_datetimes if d.date() > today)
        
        # Weekend/Weekday analysis
        weekend_count = sum(1 for d in days_of_week if d >= 5)  # Sat=5, Sun=6
        total = len(py_datetimes)
        res['% Weekend Dates'] = f"{weekend_count / total * 100:.2f}%"
        res['% Weekday Dates'] = f"{(total - weekend_count) / total * 100:.2f}%"
        
        # Top unique values
        limit = self.config_options.get('limit_unique', 5)
        date_counts = Counter(q_date_time_objects)
        sorted_dates = sorted(date_counts.items(), key=lambda x: (-x[1], str(x[0])))
        
        top_str = []
        for i, (date_obj, cnt) in enumerate(sorted_dates):
            if i >= limit:
                break
            if isinstance(date_obj, QDateTime):
                display = date_obj.toString("yyyy-MM-dd HH:mm:ss")
            else:
                display = date_obj.toString("yyyy-MM-dd")
            actual_cnt = cnt if col['is_exact'] else int(cnt * count / len(sample))
            top_str.append(f"'{display}': {actual_cnt}")
        
        res['Unique Values (Top)'] = "\n".join(top_str) if top_str else "N/A"
        if sorted_dates:
            res['Unique Values (Top)_actual_first_value'] = sorted_dates[0][0]
        
        return res


    def _has_non_printable_chars(self, text_value):
        if not isinstance(text_value, str): return False
        allowed_control = {'\t', '\n', '\r'} 
        return any(not c.isprintable() and c not in allowed_control for c in text_value)

    def _generate_hints(self, meta, col, res):
        # ... logic copied from original ...
        return "N/A"
