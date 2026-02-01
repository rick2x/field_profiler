# -*- coding: utf-8 -*-

import os
import csv
import re
import string
from qgis.PyQt import QtWidgets, QtCore, QtGui
from qgis.PyQt.QtCore import QVariant, Qt, QDate, QDateTime, QTime
from qgis.PyQt.QtWidgets import (QWidget, QVBoxLayout, QGroupBox, QLabel, QCheckBox,
                                 QListWidget, QPushButton, QDockWidget, QTableWidget,
                                 QAbstractItemView, QTableWidgetItem, QApplication,
                                 QFileDialog, QHBoxLayout, QSizePolicy, QProgressBar,
                                 QSpinBox, QFormLayout)
from qgis.gui import QgsMapLayerComboBox
from qgis.core import (QgsProject, QgsVectorLayer, QgsField, Qgis,
                       QgsStatisticalSummary, QgsMapLayerProxyModel, QgsFeatureRequest,
                       QgsExpression)

import statistics
from collections import Counter, OrderedDict
import numpy # Keep this import
from datetime import datetime # Added for analyze_date_field_enhanced


SCIPY_AVAILABLE = False
try:
    from scipy import stats as scipy_stats
    SCIPY_AVAILABLE = True
except ImportError:
    scipy_stats = None # So we can check against it

STOP_WORDS = set([
    'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from', 'has', 'he',
    'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the', 'to', 'was', 'were',
    'will', 'with',
])

class FieldProfilerDockWidget(QDockWidget):
    STAT_KEYS_NUMERIC = [
        'Non-Null Count', 'Null Count', '% Null', 'Conversion Errors',
        'Min', 'Max', 'Range', 'Sum', 'Mean', 'Median', 'Stdev (pop)', 'Mode(s)',
        'Variety (distinct)', 'Q1', 'Q3', 'IQR',
        'Outliers (IQR)', 'Min Outlier', 'Max Outlier', '% Outliers',
        'Low Variance Flag',
        'Zeros', 'Positives', 'Negatives', 'CV %',
        'Integer Values', 'Decimal Values', '% Integer Values',
        'Skewness', 'Kurtosis', 'Normality (Shapiro-Wilk p)', 'Normality (Likely Normal)',
        '1st Pctl', '5th Pctl', '95th Pctl', '99th Pctl',
        'Optimal Bins (Freedman-Diaconis)',
    ]
    STAT_KEYS_TEXT = [
        'Non-Null Count', 'Null Count', '% Null', 'Empty Strings', '% Empty',
        'Leading/Trailing Spaces', 'Internal Multiple Spaces',
        'Variety (distinct)', 'Min Length', 'Max Length', 'Avg Length',
        'Unique Values (Top)', 'Values Occurring Once',
        'Top Words', 'Pattern Matches',
        '% Uppercase', '% Lowercase', '% Titlecase', '% Mixed Case',
        'Non-Printable Chars Count',
    ]
    STAT_KEYS_DATE = [
        'Non-Null Count', 'Null Count', '% Null', 'Min Date', 'Max Date',
        'Unique Values (Top)',
        'Common Years', 'Common Months', 'Common Days',
        'Common Hours (Top 3)', '% Midnight Time', '% Noon Time',
        '% Weekend Dates', '% Weekday Dates',
        'Dates Before Today', 'Dates After Today',
    ]
    STAT_KEYS_OTHER = [ 'Non-Null Count', 'Null Count', '% Null', 'Status', 'Data Type Mismatch Hint']
    STAT_KEYS_ERROR = ['Error', 'Status']


    def __init__(self, iface, parent=None):
        super().__init__(parent) # Python 3 style super()
        self.iface = iface
        self.setObjectName("FieldProfilerDockWidgetInstance")
        self.setWindowTitle(self.tr("Field Profiler"))

        self.main_widget = QWidget()
        self.main_layout = QVBoxLayout(self.main_widget)
        self.setWidget(self.main_widget)

        self.analysis_results_cache = OrderedDict()
        self.conversion_error_feature_ids_by_field = {}
        self.non_printable_char_feature_ids_by_field = {}
        self._was_analyzing_selected_features = False

        self._define_stat_tooltips()
        self._create_input_group()
        self._create_results_ui()

        self.layerComboBox.layerChanged.connect(self.populate_fields)
        self.analyzeButton.clicked.connect(self.run_analysis)
        self.copyButton.clicked.connect(self.copy_results_to_clipboard)
        self.exportButton.clicked.connect(self.export_results_to_csv)
        self.resultsTableWidget.cellDoubleClicked.connect(self._on_cell_double_clicked)

        self.populate_fields(self.layerComboBox.currentLayer())
        if not SCIPY_AVAILABLE:
            self.iface.messageBar().pushMessage(
                "Field Profiler Warning",
                self.tr("Scipy library not found. Advanced numeric statistics (Skewness, Kurtosis, Normality) will be unavailable."),
                level=Qgis.Warning, duration=10
            )
        
    def tr(self, message):
        return QtCore.QCoreApplication.translate("FieldProfilerDockWidget", message)

    def _define_stat_tooltips(self):
        self.stat_tooltips = {
            'Non-Null Count': self.tr("Number of features with non-missing values."),
            'Null Count': self.tr("Number of features with missing (NULL) values. Double-click cell to select these features."),
            '% Null': self.tr("Percentage of features with missing (NULL) values."),
            'Conversion Errors': self.tr("Number of values that could not be converted to a numeric type (for numeric fields). Double-click cell to select these features."),
            'Low Variance Flag': self.tr("True if standard deviation is close to zero or all values are identical (for numeric fields)."),
            'Outliers (IQR)': self.tr("Number of numeric values falling outside Q1 - 1.5*IQR and Q3 + 1.5*IQR. Double-click cell to select these features."),
            'Min Outlier': self.tr("Minimum value among those flagged as outliers by IQR method."),
            'Max Outlier': self.tr("Maximum value among those flagged as outliers by IQR method."),
            '% Outliers': self.tr("Percentage of non-null values flagged as outliers by IQR method."),
            'Min': self.tr("Minimum value."),
            'Max': self.tr("Maximum value."),
            'Range': self.tr("Difference between Max and Min values."),
            'Sum': self.tr("Sum of all numeric values."),
            'Mean': self.tr("Average of numeric values."),
            'Median': self.tr("Median (middle) value of numeric data."),
            'Stdev (pop)': self.tr("Population Standard Deviation. Measures the amount of variation or dispersion."),
            'Mode(s)': self.tr("Most frequently occurring value(s)."),
            'Variety (distinct)': self.tr("Number of unique distinct values."),
            'Q1': self.tr("First Quartile (25th percentile)."),
            'Q3': self.tr("Third Quartile (75th percentile)."),
            'IQR': self.tr("Interquartile Range (Q3 - Q1)."),
            'Zeros': self.tr("Count of zero values (for numeric fields)."),
            'Positives': self.tr("Count of positive values (for numeric fields)."),
            'Negatives': self.tr("Count of negative values (for numeric fields)."),
            'CV %': self.tr("Coefficient of Variation (Stdev / Mean * 100). Indicates relative variability. N/A if mean is zero."),
            
            'Integer Values': self.tr("Count of numeric values that are whole numbers."),
            'Decimal Values': self.tr("Count of numeric values with a fractional part."),
            '% Integer Values': self.tr("Percentage of non-null numeric values that are whole numbers."),
            'Skewness': self.tr("Measure of asymmetry. Positive: tail on right. Negative: tail on left. Requires Scipy."),
            'Kurtosis': self.tr("Measure of tailedness (Fisher's, normal=0). Positive: heavy tails. Negative: light tails. Requires Scipy."),
            'Normality (Shapiro-Wilk p)': self.tr("P-value from Shapiro-Wilk test for normality. Low p (<0.05) suggests non-normal. Requires Scipy & >=3 values."),
            'Normality (Likely Normal)': self.tr("True if Shapiro-Wilk p-value > 0.05. Requires Scipy."),
            '1st Pctl': self.tr("1st Percentile."), '5th Pctl': self.tr("5th Percentile."),
            '95th Pctl': self.tr("95th Percentile."), '99th Pctl': self.tr("99th Percentile."),
            'Optimal Bins (Freedman-Diaconis)': self.tr("Suggested number of bins for a histogram using Freedman-Diaconis rule."),

            'Empty Strings': self.tr("Number of non-null strings that are empty (''). Double-click cell to select these features."),
            '% Empty': self.tr("Percentage of non-null strings that are empty."),
            'Leading/Trailing Spaces': self.tr("Number of non-empty strings that have leading or trailing whitespace. Double-click cell to select these features."),
            'Internal Multiple Spaces': self.tr("Number of non-empty strings with consecutive internal spaces (e.g., 'word  word')."),
            'Min Length': self.tr("Minimum length of non-empty strings."),
            'Max Length': self.tr("Maximum length of non-empty strings."),
            'Avg Length': self.tr("Average length of non-empty strings."),
            'Unique Values (Top)': self.tr("Most frequent distinct values and their counts. Double-click cell to select features matching the first listed value (uses cached actual value)."),
            'Values Occurring Once': self.tr("Count of distinct values that appear only once in the non-null dataset."),
            'Top Words': self.tr("Most frequent words (after removing stop words and punctuation)."),
            'Pattern Matches': self.tr("Counts of values matching common patterns (e.g., Emails, URLs)."),
            '% Uppercase': self.tr("Percentage of non-empty strings that are entirely uppercase."),
            '% Lowercase': self.tr("Percentage of non-empty strings that are entirely lowercase."),
            '% Titlecase': self.tr("Percentage of non-empty strings that are in title case (e.g., 'Title Case String')."),
            '% Mixed Case': self.tr("Percentage of non-empty strings that have mixed casing (not fully upper, lower, or title)."),
            'Non-Printable Chars Count': self.tr("Number of strings containing non-printable ASCII characters (excluding tab, newline, carriage return). Double-click to select features."),

            'Min Date': self.tr("Earliest date/datetime found."),
            'Max Date': self.tr("Latest date/datetime found."),
            'Common Years': self.tr("Most frequent years."),
            'Common Months': self.tr("Most frequent months."),
            'Common Days': self.tr("Most frequent days of the week."),
            'Common Hours (Top 3)': self.tr("Most frequent hours for DateTime fields (e.g., 10:00, 14:00)."),
            '% Midnight Time': self.tr("Percentage of DateTime values where time is 00:00:00."),
            '% Noon Time': self.tr("Percentage of DateTime values where time is 12:00:00."),
            '% Weekend Dates': self.tr("Percentage of dates falling on a Saturday or Sunday."),
            '% Weekday Dates': self.tr("Percentage of dates falling on a weekday (Mon-Fri)."),
            'Dates Before Today': self.tr("Count of dates occurring before today."),
            'Dates After Today': self.tr("Count of dates occurring after today."),
            
            'Status': self.tr("General status or summary of the field analysis."),
            'Error': self.tr("An error occurred during analysis of this field."),
            'Data Type Mismatch Hint': self.tr("A suggestion if the field's content statistically resembles a different data type.")
        }

    def _create_input_group(self):
        self.input_group_box = QGroupBox(self.tr("Input & Settings"))
        main_input_layout = QVBoxLayout()
        
        layer_label = QLabel(self.tr("Select Layer:"))
        self.layerComboBox = QgsMapLayerComboBox(self.main_widget)
        self.layerComboBox.setFilters(QgsMapLayerProxyModel.VectorLayer)
        main_input_layout.addWidget(layer_label)
        main_input_layout.addWidget(self.layerComboBox)
        
        fields_label = QLabel(self.tr("Select Field(s):"))
        self.fieldListWidget = QListWidget()
        self.fieldListWidget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        main_input_layout.addWidget(fields_label)
        main_input_layout.addWidget(self.fieldListWidget)
        
        self.selectedOnlyCheckbox = QCheckBox(self.tr("Analyze selected features only"))
        main_input_layout.addWidget(self.selectedOnlyCheckbox)
        
        # --- Basic Configuration Group ---
        config_group = QGroupBox(self.tr("Configuration"))
        config_layout = QFormLayout()
        self.limitUniqueSpinBox = QSpinBox()
        self.limitUniqueSpinBox.setRange(1, 100); self.limitUniqueSpinBox.setValue(5)
        self.limitUniqueSpinBox.setToolTip(self.tr("Maximum number of unique values to display in 'Unique Values (Top)'."))
        config_layout.addRow(self.tr("Unique Values Limit:"), self.limitUniqueSpinBox)
        self.decimalPlacesSpinBox = QSpinBox()
        self.decimalPlacesSpinBox.setRange(0, 10); self.decimalPlacesSpinBox.setValue(2)
        self.decimalPlacesSpinBox.setToolTip(self.tr("Number of decimal places for numeric statistics in the table."))
        config_layout.addRow(self.tr("Numeric Decimal Places:"), self.decimalPlacesSpinBox)
        config_group.setLayout(config_layout)
        main_input_layout.addWidget(config_group)

        # --- Detailed Analysis Options Group ---
        detailed_options_group = QGroupBox(self.tr("Detailed Analysis Options"))
        detailed_options_layout = QVBoxLayout()

        self.chk_numeric_dist_shape = QCheckBox(self.tr("Numeric: Distribution Shape (Skew, Kurtosis, Normality)"))
        self.chk_numeric_dist_shape.setChecked(True)
        self.chk_numeric_dist_shape.setToolTip(self.tr("Requires Scipy. Calculates skewness, kurtosis, and Shapiro-Wilk normality test."))
        detailed_options_layout.addWidget(self.chk_numeric_dist_shape)

        self.chk_numeric_adv_percentiles = QCheckBox(self.tr("Numeric: Advanced Percentiles (1,5,95,99)"))
        self.chk_numeric_adv_percentiles.setChecked(True)
        detailed_options_layout.addWidget(self.chk_numeric_adv_percentiles)
        
        self.chk_numeric_int_decimal = QCheckBox(self.tr("Numeric: Integer/Decimal Counts & Optimal Bins"))
        self.chk_numeric_int_decimal.setChecked(True)
        detailed_options_layout.addWidget(self.chk_numeric_int_decimal)
        
        self.chk_numeric_outlier_details = QCheckBox(self.tr("Numeric: Min/Max Outlier Values & %"))
        self.chk_numeric_outlier_details.setChecked(True)
        detailed_options_layout.addWidget(self.chk_numeric_outlier_details)

        self.chk_text_case_analysis = QCheckBox(self.tr("Text: Case Analysis & Advanced Whitespace"))
        self.chk_text_case_analysis.setChecked(True)
        detailed_options_layout.addWidget(self.chk_text_case_analysis)

        self.chk_text_rarity_nonprintable = QCheckBox(self.tr("Text: Rarity (Once-Occurring) & Non-Printable Chars"))
        self.chk_text_rarity_nonprintable.setChecked(True)
        detailed_options_layout.addWidget(self.chk_text_rarity_nonprintable)
        
        self.chk_date_time_weekend = QCheckBox(self.tr("Date: Time Components & Weekend/Weekday Analysis"))
        self.chk_date_time_weekend.setChecked(True)
        detailed_options_layout.addWidget(self.chk_date_time_weekend)

        detailed_options_group.setLayout(detailed_options_layout)
        main_input_layout.addWidget(detailed_options_group)
        
        # --- Analyze Button and Progress Bar ---
        self.analyzeButton = QPushButton(self.tr("Analyze Selected Fields"))
        main_input_layout.addWidget(self.analyzeButton)
        self.progressBar = QProgressBar(self)
        self.progressBar.setTextVisible(True); self.progressBar.setVisible(False)
        main_input_layout.addWidget(self.progressBar)
        
        self.input_group_box.setLayout(main_input_layout)
        self.main_layout.addWidget(self.input_group_box)

    def _create_results_ui(self):
        self.results_group_box = QGroupBox(self.tr("Analysis Results"))
        results_layout = QVBoxLayout()
        self.resultsTableWidget = QTableWidget()
        self.resultsTableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.resultsTableWidget.setAlternatingRowColors(True)
        self.resultsTableWidget.setSortingEnabled(True) 
        results_layout.addWidget(self.resultsTableWidget)
        button_layout = QHBoxLayout()
        self.copyButton = QPushButton(self.tr("Copy Table"))
        self.exportButton = QPushButton(self.tr("Export Table"))
        button_layout.addStretch()
        button_layout.addWidget(self.copyButton); button_layout.addWidget(self.exportButton)
        results_layout.addLayout(button_layout)
        self.results_group_box.setLayout(results_layout)
        self.results_group_box.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.main_layout.addWidget(self.results_group_box)

    def populate_fields(self, layer):
        self.fieldListWidget.clear()
        self.resultsTableWidget.clear()
        self.resultsTableWidget.setRowCount(0)
        self.resultsTableWidget.setColumnCount(0)
        self.analysis_results_cache = OrderedDict()
        self.conversion_error_feature_ids_by_field = {}
        self.non_printable_char_feature_ids_by_field = {}
        self.progressBar.setVisible(False)

        if layer and isinstance(layer, QgsVectorLayer):
            self.fieldListWidget.setEnabled(True); self.selectedOnlyCheckbox.setEnabled(True); self.analyzeButton.setEnabled(True)
            for field in layer.fields(): item_text = f"{field.name()} ({field.typeName()})"; self.fieldListWidget.addItem(item_text)
        else:
            self.fieldListWidget.setEnabled(False); self.selectedOnlyCheckbox.setEnabled(False); self.analyzeButton.setEnabled(False)

    def _get_detailed_options_state(self):
        return {
            'numeric_dist_shape': self.chk_numeric_dist_shape.isChecked(),
            'numeric_adv_percentiles': self.chk_numeric_adv_percentiles.isChecked(),
            'numeric_int_decimal': self.chk_numeric_int_decimal.isChecked(),
            'numeric_outlier_details': self.chk_numeric_outlier_details.isChecked(),
            'text_case_analysis': self.chk_text_case_analysis.isChecked(),
            'text_rarity_nonprintable': self.chk_text_rarity_nonprintable.isChecked(),
            'date_time_weekend': self.chk_date_time_weekend.isChecked(),
        }

    def run_analysis(self):
        self.resultsTableWidget.clear(); self.resultsTableWidget.setRowCount(0); self.resultsTableWidget.setColumnCount(0)
        self.analysis_results_cache = OrderedDict()
        self.conversion_error_feature_ids_by_field = {}
        self.non_printable_char_feature_ids_by_field = {}

        current_layer = self.layerComboBox.currentLayer()
        selected_list_items = self.fieldListWidget.selectedItems()
        
        self.current_limit_unique_display = self.limitUniqueSpinBox.value()
        self.current_decimal_places = self.decimalPlacesSpinBox.value()
        self._was_analyzing_selected_features = self.selectedOnlyCheckbox.isChecked()
        detailed_options = self._get_detailed_options_state()

        if not current_layer or not isinstance(current_layer, QgsVectorLayer):
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Please select a valid vector layer."), level=Qgis.Warning); return
        if not selected_list_items:
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Please select one or more fields to analyze."), level=Qgis.Warning); return

        selected_field_names_from_widget = [item.text().split(" (")[0] for item in selected_list_items]
        if not selected_field_names_from_widget:
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Error extracting field names."), level=Qgis.Critical); return

        request = QgsFeatureRequest()
        feature_count_total_layer = current_layer.featureCount()
        feature_count_analyzed = 0

        if self._was_analyzing_selected_features:
            selected_ids = current_layer.selectedFeatureIds()
            if not selected_ids:
                self.iface.messageBar().pushMessage(self.tr("Warning"), self.tr("No features selected for analysis."), level=Qgis.Warning)
                self.progressBar.setVisible(False); return
            request.setFilterFids(selected_ids); feature_count_analyzed = len(selected_ids)
            analysis_scope_message = self.tr("Analyzing Selected Features ({})").format(feature_count_analyzed)
        else:
            feature_count_analyzed = feature_count_total_layer
            analysis_scope_message = self.tr("Analyzing All Features ({})").format(feature_count_analyzed)

        if feature_count_analyzed == 0:
            self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No features to analyze."), level=Qgis.Info, duration=5)
            self.progressBar.setVisible(False); return

        self.iface.messageBar().pushMessage(self.tr("Info"), analysis_scope_message, level=Qgis.Info, duration=3)
        QApplication.processEvents()

        self.progressBar.setRange(0, feature_count_analyzed if feature_count_analyzed > 0 else 100)
        self.progressBar.setValue(0); self.progressBar.setVisible(True)

        field_data_collector = OrderedDict()
        qgs_fields_objects = current_layer.fields()
        field_metadata = {}

        valid_selected_field_names = []
        for field_name in selected_field_names_from_widget:
            field_index = qgs_fields_objects.lookupField(field_name)
            if field_index == -1:
                self.analysis_results_cache[field_name] = {'Error': 'Field not found'}; continue
            
            valid_selected_field_names.append(field_name)
            field_obj = qgs_fields_objects.field(field_index)
            field_metadata[field_name] = {'index': field_index, 'object': field_obj, 'type': field_obj.type()}
            collector_init = {
                'raw_values': [], 'null_count': 0,
                'non_printable_fids': [] 
            }
            if field_obj.isNumeric():
                collector_init.update({'float_values': [], 'conversion_errors': 0, 'conversion_error_feature_ids': []})
            if field_obj.type() in [QVariant.Date, QVariant.DateTime]:
                 collector_init['original_variants'] = []

            field_data_collector[field_name] = collector_init
        
        if not valid_selected_field_names:
            self.populate_results_table(self.analysis_results_cache, selected_field_names_from_widget)
            self.progressBar.setVisible(False); return

        current_iterator = current_layer.getFeatures(request)
        iteration_count = 0
        try:
            for feature in current_iterator:
                iteration_count += 1
                fid = feature.id()
                for field_name in valid_selected_field_names:
                    meta = field_metadata[field_name]
                    collector = field_data_collector[field_name]
                    val = feature[meta['index']]
                    
                    if meta['type'] in [QVariant.Date, QVariant.DateTime]:
                        collector['original_variants'].append(val if (val is not None and not (hasattr(val, 'isNull') and val.isNull())) else None)

                    if val is None or (hasattr(val, 'isNull') and val.isNull()):
                        collector['null_count'] += 1
                    else:
                        collector['raw_values'].append(val)
                        if meta['object'].isNumeric():
                            try:
                                collector['float_values'].append(float(val))
                            except (ValueError, TypeError):
                                collector['conversion_errors'] += 1
                                collector['conversion_error_feature_ids'].append(fid)
                        elif meta['type'] == QVariant.String:
                            if detailed_options['text_rarity_nonprintable']:
                                if self._has_non_printable_chars(str(val)):
                                    collector['non_printable_fids'].append(fid)

                if iteration_count % 100 == 0 or iteration_count == feature_count_analyzed:
                    self.progressBar.setValue(iteration_count); QApplication.processEvents()
        except Exception as e_iter:
            for field_name in valid_selected_field_names: self.analysis_results_cache[field_name] = {'Error': f'Feature iteration error: {e_iter}'}
            self.populate_results_table(self.analysis_results_cache, selected_field_names_from_widget)
            self.progressBar.setVisible(False); return
        
        self.progressBar.setValue(feature_count_analyzed)

        for field_name in valid_selected_field_names:
            data = field_data_collector[field_name]
            meta = field_metadata[field_name]

            if meta['object'].isNumeric() and data.get('conversion_error_feature_ids'):
                self.conversion_error_feature_ids_by_field[field_name] = data['conversion_error_feature_ids']
            if meta['type'] == QVariant.String and data.get('non_printable_fids'):
                self.non_printable_char_feature_ids_by_field[field_name] = list(set(data['non_printable_fids'])) 

            non_null_count = len(data['raw_values'])
            percent_null = (data['null_count'] / feature_count_analyzed * 100) if feature_count_analyzed > 0 else 0
            field_results = OrderedDict([('Null Count', data['null_count']), ('% Null', f"{percent_null:.2f}%"), ('Non-Null Count', non_null_count)])
            
            status_set = False
            if non_null_count == 0:
                if meta['object'].isNumeric() and data.get('conversion_errors', 0) > 0:
                    field_results['Status'] = f"All values Null or conversion errors ({data['conversion_errors']})"
                else:
                    field_results['Status'] = 'All Null or Empty'
                status_set = True
            
            analysis_for_field = {}
            if not status_set: 
                try:
                    if meta['object'].isNumeric():
                        analysis_for_field = self.analyze_numeric_field_from_list(data['float_values'], data.get('conversion_errors',0), detailed_options, non_null_count)
                    elif meta['type'] == QVariant.String:
                        analysis_for_field = self.analyze_text_field(data['raw_values'], non_null_count, detailed_options)
                    elif meta['type'] in [QVariant.Date, QVariant.DateTime]:
                        original_variants = data.get('original_variants', data['raw_values'])
                        analysis_for_field = self.analyze_date_field_enhanced(original_variants, non_null_count, detailed_options)
                    else:
                        analysis_for_field = {'Status': 'Analysis not implemented for this type'}
                except Exception as e_analysis:
                    analysis_for_field = {'Error': f'Analysis function error: {e_analysis}'}
            
            field_results.update(analysis_for_field)

            hint = "N/A"
            if meta['type'] == QVariant.String and non_null_count > 0:
                numeric_like_count = sum(1 for s_val in data['raw_values'] if str(s_val).replace('.', '', 1).strip().isdigit()) # strip to handle " 123 "
                if numeric_like_count / non_null_count > 0.9: 
                    hint = "High % of numeric-like strings. Consider if this field should be numeric."
            elif meta['object'].isNumeric() and non_null_count > 0:
                if field_results.get('Variety (distinct)', float('inf')) < 15 and non_null_count > 20:
                     hint = "Low variety for a numeric field. Consider if this is categorical or a code."

            field_results['Data Type Mismatch Hint'] = hint
            self.analysis_results_cache[field_name] = field_results
        
        self.populate_results_table(self.analysis_results_cache, selected_field_names_from_widget)
        self.progressBar.setVisible(False); QApplication.processEvents()

    def populate_results_table(self, results_data, field_names_for_header):
        self.resultsTableWidget.clear()
        if not results_data and not field_names_for_header: return
        all_stat_names_from_data = set()
        for field_name, field_data in results_data.items(): all_stat_names_from_data.update(field_data.keys())
        
        # Filter out internal keys like '_actual_first_value' from the set of names to display as rows
        all_displayable_stat_names = {stat for stat in all_stat_names_from_data if not stat.endswith('_actual_first_value')}

        # --- Determine row order for statistics ---
        stat_rows_ordered = [] # This will hold the final ordered list of statistic keys (original English keys)
        seen_keys_for_order = set()
        
        # Combine all predefined STAT_KEYS lists into a single ordered list without duplicates
        predefined_order_source = []
        temp_seen_for_predefined_order = set()
        for key_list in [self.STAT_KEYS_NUMERIC, self.STAT_KEYS_TEXT, self.STAT_KEYS_DATE, 
                         self.STAT_KEYS_OTHER, self.STAT_KEYS_ERROR]:
            for key in key_list:
                if key not in temp_seen_for_predefined_order:
                    predefined_order_source.append(key)
                    temp_seen_for_predefined_order.add(key)

        # Add keys from predefined_order_source if they are present in the actual results
        for key in predefined_order_source:
            if key in all_displayable_stat_names and key not in seen_keys_for_order:
                stat_rows_ordered.append(key)
                seen_keys_for_order.add(key)
        
        # Add any remaining keys from results_data that weren't in predefined lists (sorted alphabetically)
        extras = sorted([key for key in all_displayable_stat_names if key not in seen_keys_for_order])
        stat_rows_ordered.extend(extras)
        
        # --- Setup table dimensions and headers ---
        num_rows = len(stat_rows_ordered)
        num_cols = len(field_names_for_header) + 1 # +1 for the statistic name column
        self.resultsTableWidget.setRowCount(num_rows)
        self.resultsTableWidget.setColumnCount(num_cols)
        
        # Headers: First column is "Statistic", others are field names
        headers = [self.tr("Statistic")] + field_names_for_header
        self.resultsTableWidget.setHorizontalHeaderLabels(headers)
        
        quality_keywords = ['%', 'Null', 'Empty', 'Error', 'Outlier', 'Spaces', 'Variance', 'Flag', 'Conversion', 'Mismatch', 'Non-Printable'] 
        dp = self.current_decimal_places # Decimal places for formatting floats
        
        # --- Populate table cells ---
        for r, original_stat_key in enumerate(stat_rows_ordered): # original_stat_key is the English key
            # Statistic Name Item (First Column)
            stat_item = QTableWidgetItem(self.tr(original_stat_key)) # Display translated name
            stat_item.setData(Qt.UserRole, original_stat_key) # Store original English key
            stat_item.setToolTip(self.stat_tooltips.get(original_stat_key, self.tr("No description available.")))
            
            is_quality_issue = any(keyword.lower() in original_stat_key.lower() for keyword in quality_keywords) or \
                               original_stat_key == 'Error'
            
            # Check boolean quality issues for the first field to color the statistic name row
            # This still assumes the first field is representative for row-level coloring
            first_field_name_for_color = field_names_for_header[0] if field_names_for_header else None
            if first_field_name_for_color:
                 first_field_data = results_data.get(first_field_name_for_color, {})
                 if original_stat_key == 'Normality (Likely Normal)' and first_field_data.get(original_stat_key) is False:
                     is_quality_issue = True
                 if original_stat_key == 'Low Variance Flag' and first_field_data.get(original_stat_key) is True:
                     is_quality_issue = True

            if is_quality_issue:
                stat_item.setBackground(QtGui.QColor(255, 240, 240)) # Light red
            elif original_stat_key.startswith('%') or "Pctl" in original_stat_key or original_stat_key in ['Skewness', 'Kurtosis']:
                stat_item.setBackground(QtGui.QColor(240, 240, 255)) # Light blue
            else:
                stat_item.setBackground(QtGui.QColor(230, 230, 230)) # Light grey
            
            self.resultsTableWidget.setItem(r, 0, stat_item)

            # Data Cells (Subsequent Columns)
            for c, field_name in enumerate(field_names_for_header):
                field_data = results_data.get(field_name, {})
                value = field_data.get(original_stat_key, "") # Get value using original_stat_key
                display_text = ""
                
                if isinstance(value, bool):
                    display_text = str(value)
                elif isinstance(value, float):
                    if original_stat_key == 'Normality (Shapiro-Wilk p)':
                         display_text = f"{value:.4g}" if not numpy.isnan(value) else "N/A"
                    else:
                         display_text = f"{value:.{dp}f}" if not numpy.isnan(value) else "N/A"
                elif isinstance(value, list) and original_stat_key != 'Mode(s)': # Check original_stat_key here
                    display_text = "; ".join(map(str, value))
                elif isinstance(value, list) and original_stat_key == 'Mode(s)': # Check original_stat_key here
                    # Format numbers in mode list with specified decimal places
                    formatted_modes = []
                    for v_mode in value:
                        if isinstance(v_mode, (int, float)):
                            try:
                                formatted_modes.append(f"{float(v_mode):.{dp}f}")
                            except ValueError: # Handle potential non-numeric items if list is mixed
                                formatted_modes.append(str(v_mode))
                        else:
                            formatted_modes.append(str(v_mode))
                    display_text = ", ".join(formatted_modes)
                else:
                    display_text = str(value)
                
                item = QTableWidgetItem(display_text)
                
                align_right_keywords = ['Count', 'Error', 'Outlier', 'Zero', 'Positive', 'Negative', 'Space', 'Empty', 'Value', 'Length', 'Pctl', 'Optimal Bins']
                align_right = isinstance(value, (int, float, bool, numpy.number)) or \
                              '%' in original_stat_key or \
                              any(kw in original_stat_key for kw in align_right_keywords) # Check original_stat_key

                item.setTextAlignment(Qt.AlignVCenter | (Qt.AlignRight if align_right else Qt.AlignLeft))
                
                if isinstance(value, str) and ('\n' in value or len(value) > 60):
                    item.setToolTip(value) # Show full value in tooltip if long or multiline
                elif item.text() == "N/A (Scipy not found)" or item.text() == "N/A (>=3 values needed)" or item.text() == "N/A (<3 valid)":
                    item.setForeground(QtGui.QBrush(Qt.gray)) # Grey out unavailable stats

                self.resultsTableWidget.setItem(r, c + 1, item)
        
        self.resultsTableWidget.resizeColumnsToContents()

    def analyze_numeric_field_from_list(self, non_null_values_list_float, conversion_errors, options, total_non_null_count):
        results = OrderedDict()
        results['Conversion Errors'] = conversion_errors
        
        try:
            data_np = numpy.array(non_null_values_list_float, dtype=float)
            if numpy.any(numpy.isinf(data_np)): 
                data_np = data_np[~numpy.isinf(data_np)]
        except Exception: 
            data_np = numpy.array([], dtype=float) 

        count_val = len(data_np)

        if count_val == 0:
            results['Status'] = 'No valid numeric data' if conversion_errors == 0 else f'No valid data ({conversion_errors} conversion errors)'
            for key in self.STAT_KEYS_NUMERIC:
                if key not in ['Non-Null Count', 'Null Count', '% Null', 'Conversion Errors', 'Status']:
                    if key in ['Variety (distinct)', 'Zeros', 'Positives', 'Negatives', 'Outliers (IQR)', 'Integer Values', 'Decimal Values', '% Outliers', 'Min Outlier', 'Max Outlier']: results[key] = 0
                    elif key in ['Low Variance Flag', 'Normality (Likely Normal)']: results[key] = False
                    elif key == '% Integer Values': results[key] = f"{0.0:.2f}%"
                    else: results[key] = 'N/A'
            return results

        min_val = numpy.nanmin(data_np) if count_val > 0 else numpy.nan
        max_val = numpy.nanmax(data_np) if count_val > 0 else numpy.nan
        sum_val = numpy.nansum(data_np) if count_val > 0 else numpy.nan
        mean_val = numpy.nanmean(data_np) if count_val > 0 else numpy.nan
        median_val = numpy.nanmedian(data_np) if count_val > 0 else numpy.nan
        
        results['Min'] = min_val; results['Max'] = max_val
        results['Range'] = max_val - min_val if not (numpy.isnan(min_val) or numpy.isnan(max_val)) else numpy.nan
        results['Sum'] = sum_val; results['Mean'] = mean_val; results['Median'] = median_val

        std_dev_pop = numpy.std(data_np) if count_val > 0 else numpy.nan
        results['Stdev (pop)'] = std_dev_pop

        modes_val = 'N/A'
        if count_val > 0:
            if SCIPY_AVAILABLE:
                try: 
                    mode_res = scipy_stats.mode(data_np, nan_policy='omit', keepdims=False) 
                    if hasattr(mode_res, 'mode') and numpy.size(mode_res.mode) > 0:
                        modes_val = list(mode_res.mode) if isinstance(mode_res.mode, numpy.ndarray) else [mode_res.mode]
                    elif not hasattr(mode_res, 'mode') and isinstance(mode_res, tuple) and len(mode_res) > 0: # Older scipy
                        if numpy.size(mode_res[0]) > 0:
                             modes_val = list(mode_res[0])
                        else: modes_val = 'N/A (no mode or all unique)'
                    else: 
                        modes_val = 'N/A (no mode or all unique)'
                except Exception:
                     modes_val = 'N/A (mode error)'
            else: 
                try:
                    # Ensure data_np is not empty before tolist()
                    modes_val = statistics.multimode(data_np.tolist()) if data_np.size > 0 else 'N/A (no data)'
                except statistics.StatisticsError: # No unique mode or data is empty
                    modes_val = 'N/A (no unique mode / empty)'
                except TypeError: # if data_np contains NaNs, multimode might fail
                    try:
                        non_nan_list = [x for x in data_np.tolist() if not numpy.isnan(x)]
                        if non_nan_list:
                            modes_val = statistics.multimode(non_nan_list)
                        else:
                            modes_val = 'N/A (all NaN or empty)'
                    except statistics.StatisticsError:
                         modes_val = 'N/A (no unique mode after NaN removal)'

        results['Mode(s)'] = modes_val

        results['Variety (distinct)'] = len(numpy.unique(data_np[~numpy.isnan(data_np)])) if count_val > 0 else 0
        
        q1, q3, iqr_val = numpy.nan, numpy.nan, numpy.nan
        outlier_count = 0
        min_outlier_val, max_outlier_val, percent_outliers = numpy.nan, numpy.nan, 0.0 # Default percent_outliers to 0.0

        if count_val > 0:
            q1 = numpy.nanpercentile(data_np, 25)
            q3 = numpy.nanpercentile(data_np, 75)
            if not (numpy.isnan(q1) or numpy.isnan(q3)):
                iqr_val = q3 - q1
                if options.get('numeric_outlier_details', True) and not numpy.isnan(iqr_val) : # Ensure iqr_val is not NaN
                    lower_bound = q1 - 1.5 * iqr_val
                    upper_bound = q3 + 1.5 * iqr_val
                    outliers_bool = (data_np < lower_bound) | (data_np > upper_bound)
                    outliers_vals = data_np[outliers_bool & ~numpy.isnan(data_np)] 
                    outlier_count = len(outliers_vals)
                    if outlier_count > 0:
                        min_outlier_val = numpy.min(outliers_vals)
                        max_outlier_val = numpy.max(outliers_vals)
                    # Calculate percent_outliers based on count_val (total non-NaN valid numerics)
                    percent_outliers = (outlier_count / count_val * 100.0) if count_val > 0 else 0.0

        results['Q1'] = q1; results['Q3'] = q3; results['IQR'] = iqr_val
        results['Outliers (IQR)'] = outlier_count
        if options.get('numeric_outlier_details', True):
            results['Min Outlier'] = min_outlier_val
            results['Max Outlier'] = max_outlier_val
            results['% Outliers'] = percent_outliers # Already a float
        else:
            results['Min Outlier'] = "N/A (Opt.)"; results['Max Outlier'] = "N/A (Opt.)"; results['% Outliers'] = "N/A (Opt.)"


        low_variance = False
        if count_val == 1: low_variance = True
        elif not numpy.isnan(std_dev_pop) and numpy.isclose(std_dev_pop, 0.0): low_variance = True
        elif results['Variety (distinct)'] == 1 and count_val > 1: low_variance = True
        results['Low Variance Flag'] = low_variance
        
        results['Zeros'] = numpy.sum(data_np == 0) if count_val > 0 else 0
        results['Positives'] = numpy.sum(data_np > 0) if count_val > 0 else 0
        results['Negatives'] = numpy.sum(data_np < 0) if count_val > 0 else 0
        
        cv = numpy.nan
        if not numpy.isnan(mean_val) and mean_val != 0 and not numpy.isnan(std_dev_pop):
            cv = (std_dev_pop / mean_val) * 100
        results['CV %'] = cv

        if options.get('numeric_int_decimal', False) and count_val > 0:
            valid_data_for_int_check = data_np[~numpy.isnan(data_np)] 
            integer_values_count = numpy.sum(valid_data_for_int_check == numpy.floor(valid_data_for_int_check))
            results['Integer Values'] = int(integer_values_count)
            results['Decimal Values'] = len(valid_data_for_int_check) - int(integer_values_count) 
            results['% Integer Values'] = (int(integer_values_count) / len(valid_data_for_int_check) * 100.0) if len(valid_data_for_int_check) > 0 else 0.0
            
            # Optimal Bins
            optimal_bins_val = "N/A"
            if count_val > 1 and not numpy.isnan(iqr_val) and iqr_val > 0 and not (numpy.isnan(min_val) or numpy.isnan(max_val)):
                bin_width = 2 * iqr_val / (count_val**(1/3))
                if bin_width > 0 : 
                    data_range = max_val - min_val
                    if not numpy.isnan(data_range) and data_range > 0:
                         optimal_bins_val = int(numpy.ceil(data_range / bin_width))
                    elif data_range == 0: 
                         optimal_bins_val = 1
                    # else data_range is NaN or negative (shouldn't happen if min/max are valid)
                # else bin_width is zero or negative (e.g. if count_val is huge or iqr_val very small)
            elif count_val > 0 and results['Variety (distinct)'] > 0: # Fallback if IQR method not suitable
                optimal_bins_val = results['Variety (distinct)']
            elif count_val == 1:
                optimal_bins_val = 1

            results['Optimal Bins (Freedman-Diaconis)'] = optimal_bins_val
        else:
            results['Integer Values'] = "N/A (Opt.)"; results['Decimal Values'] = "N/A (Opt.)"; results['% Integer Values'] = "N/A (Opt.)"
            results['Optimal Bins (Freedman-Diaconis)'] = "N/A (Opt.)"


        if options.get('numeric_dist_shape', False) and SCIPY_AVAILABLE and count_val > 0:
            data_for_scipy = data_np[~numpy.isnan(data_np)] 
            if len(data_for_scipy) > 0:
                results['Skewness'] = scipy_stats.skew(data_for_scipy) if len(data_for_scipy) > 0 else numpy.nan
                results['Kurtosis'] = scipy_stats.kurtosis(data_for_scipy, fisher=True) if len(data_for_scipy) > 0 else numpy.nan
                if len(data_for_scipy) >= 3: 
                    try:
                        shapiro_stat, shapiro_p = scipy_stats.shapiro(data_for_scipy)
                        results['Normality (Shapiro-Wilk p)'] = shapiro_p
                        results['Normality (Likely Normal)'] = bool(shapiro_p > 0.05) 
                    except ValueError as e_shapiro: # Handles cases like "Data must be at least 3 samples" or "Data must be distinct"
                        results['Normality (Shapiro-Wilk p)'] = "N/A (Error)" # More generic, could log e_shapiro
                        results['Normality (Likely Normal)'] = "N/A (Error)"
                else:
                    results['Normality (Shapiro-Wilk p)'] = "N/A (<3 valid)"
                    results['Normality (Likely Normal)'] = "N/A (<3 valid)"
            else: 
                results['Skewness'] = numpy.nan; results['Kurtosis'] = numpy.nan
                results['Normality (Shapiro-Wilk p)'] = "N/A (all NaN)"; results['Normality (Likely Normal)'] = "N/A (all NaN)"

        elif options.get('numeric_dist_shape', False) and not SCIPY_AVAILABLE:
            scipy_na_msg = "N/A (Scipy missing)"
            results['Skewness'] = scipy_na_msg; results['Kurtosis'] = scipy_na_msg
            results['Normality (Shapiro-Wilk p)'] = scipy_na_msg; results['Normality (Likely Normal)'] = scipy_na_msg
        else: 
            opt_na_msg = "N/A (Opt.)"
            results['Skewness'] = opt_na_msg; results['Kurtosis'] = opt_na_msg
            results['Normality (Shapiro-Wilk p)'] = opt_na_msg; results['Normality (Likely Normal)'] = opt_na_msg


        if options.get('numeric_adv_percentiles', False) and count_val > 0:
            percentiles_to_calc = [1, 5, 95, 99]
            # Ensure data_np is not all NaNs before calling nanpercentile
            if not numpy.all(numpy.isnan(data_np)):
                pctl_values = numpy.nanpercentile(data_np, percentiles_to_calc)
                results['1st Pctl'] = pctl_values[0]
                results['5th Pctl'] = pctl_values[1]
                results['95th Pctl'] = pctl_values[2]
                results['99th Pctl'] = pctl_values[3]
            else:
                nan_val = numpy.nan
                results['1st Pctl'] = nan_val; results['5th Pctl'] = nan_val
                results['95th Pctl'] = nan_val; results['99th Pctl'] = nan_val
        else:
            opt_na_msg = "N/A (Opt.)"
            results['1st Pctl'] = opt_na_msg; results['5th Pctl'] = opt_na_msg
            results['95th Pctl'] = opt_na_msg; results['99th Pctl'] = opt_na_msg
            
        return results

    def _has_non_printable_chars(self, text_value):
        if not isinstance(text_value, str): return False
        allowed_control = {'\t', '\n', '\r'} 
        return any(not c.isprintable() and c not in allowed_control for c in text_value)


    def analyze_text_field(self, values, non_null_count, options):
        results = OrderedDict(); dp = self.current_decimal_places
        
        if non_null_count == 0: 
            results['Status'] = 'No text data'
            for key in self.STAT_KEYS_TEXT:
                 if key not in ['Non-Null Count', 'Null Count', '% Null', 'Status']:
                    if key in ['Empty Strings', 'Leading/Trailing Spaces', 'Internal Multiple Spaces', 
                               'Variety (distinct)', 'Values Occurring Once', 'Non-Printable Chars Count']: results[key] = 0
                    elif key == '% Empty': results[key] = f"{0.0:.{dp}f}%"
                    elif key.startswith('%') and ('Case' in key): results[key] = f"{0.0:.{dp}f}%" 
                    else: results[key] = 'N/A'
            return results

        str_values = [str(v) if v is not None else "" for v in values] # Ensure all are strings, handle None as empty

        empty_string_count = str_values.count('')
        percent_empty = (empty_string_count / non_null_count * 100.0) if non_null_count > 0 else 0.0
        results['Empty Strings'] = empty_string_count
        results['% Empty'] = f"{percent_empty:.{dp}f}%" 
        
        non_empty_str_values = [s for s in str_values if s] 
        count_non_empty = len(non_empty_str_values)

        min_len, max_len, avg_len_val = 'N/A', 'N/A', 'N/A'
        if count_non_empty > 0:
            lengths = [len(s) for s in non_empty_str_values]
            min_len, max_len, avg_len_val = min(lengths), max(lengths), statistics.mean(lengths)
        results['Min Length'] = min_len; results['Max Length'] = max_len
        results['Avg Length'] = f"{avg_len_val:.{dp}f}" if isinstance(avg_len_val, float) else avg_len_val
        
        value_counts = Counter(str_values) 
        results['Variety (distinct)'] = len(value_counts)
        
        sorted_counts = sorted(value_counts.items(), key=lambda item: (-item[1], item[0]))
        top_unique_list = []; actual_first_unique_value_for_selection = None
        limit_unique = self.current_limit_unique_display
        if sorted_counts:
            actual_first_unique_value_for_selection = sorted_counts[0][0]
            for i, (val_str, count) in enumerate(sorted_counts):
                if i >= limit_unique: break
                display_val_preview = f"'{val_str[:50]}{'...' if len(val_str) > 50 else ''}'"
                if val_str == "": display_val_preview = self.tr("'(Empty String)'")
                top_unique_list.append(f"{display_val_preview}: {count}")
        results['Unique Values (Top)'] = "\n".join(top_unique_list) if top_unique_list else "N/A"
        if top_unique_list: # Store even if None or empty string
            results['Unique Values (Top)_actual_first_value'] = actual_first_unique_value_for_selection
        
        if options.get('text_rarity_nonprintable', False):
            # Values occurring once should exclude empty strings from this specific count if desired
            # (currently includes empty string if it appears once)
            results['Values Occurring Once'] = sum(1 for v_str,c in value_counts.items() if c == 1) 
            results['Non-Printable Chars Count'] = sum(1 for s_val in str_values if self._has_non_printable_chars(s_val))
        else:
            results['Values Occurring Once'] = "N/A (Opt.)"
            results['Non-Printable Chars Count'] = "N/A (Opt.)"

        if options.get('text_case_analysis', False):
            if count_non_empty > 0:
                upper_count = sum(1 for s_val in non_empty_str_values if s_val.isupper())
                lower_count = sum(1 for s_val in non_empty_str_values if s_val.islower())
                title_count = sum(1 for s_val in non_empty_str_values if s_val.istitle())
                
                # Calculate mixed_count more accurately: non-empty strings that are not
                # purely uppercase, not purely lowercase, and not purely titlecase.
                mixed_count = 0
                for s_val in non_empty_str_values:
                    is_u = s_val.isupper()
                    is_l = s_val.islower()
                    is_t = s_val.istitle()
                    # A string is "mixed" if it's not any single one of these exclusively.
                    # However, a string like "I" is_upper and is_title.
                    # Simplest for "Mixed Case" is: total_non_empty - (exclusive_upper + exclusive_lower + exclusive_title)
                    # This requires identifying exclusive cases or defining mixed as "none of the above".
                    # Current definition of % Mixed Case as 100 - sum_of_others is an approximation of "not one of the main categories".
                    # Let's count strings that are *not* (isupper OR islower OR istitle)
                    # This is still complex. Let's stick to the sum-based approach for now.
                    # The sum of these % might exceed 100% if a string fits multiple (e.g. "I" is upper & title)
                    # A clearer approach:
                    # % Uppercase: s.isupper() and not s.islower() and not s.istitle() (unless it's trivial like "I")
                    # This is too complex for summary stats. The current independent counts are fine.

                results['% Uppercase'] = f"{(upper_count / count_non_empty * 100.0):.{dp}f}%"
                results['% Lowercase'] = f"{(lower_count / count_non_empty * 100.0):.{dp}f}%"
                results['% Titlecase'] = f"{(title_count / count_non_empty * 100.0):.{dp}f}%"
                
                # For % Mixed Case: count strings that are not fully upper, not fully lower, not fully title
                # This is tricky due to overlaps (e.g. "A" is upper and title).
                # A simple definition of mixed: if it has both upper and lower characters, and is not title.
                # Or, more simply, items that don't fall into a clear category.
                # The previous (100 - sum) is problematic if sum > 100.
                # Count explicitly:
                explicit_mixed_count = 0
                for s in non_empty_str_values:
                    if not s.isupper() and not s.islower() and not s.istitle():
                        explicit_mixed_count +=1
                results['% Mixed Case'] = f"{(explicit_mixed_count / count_non_empty * 100.0):.{dp}f}%"


                results['Internal Multiple Spaces'] = sum(1 for s_val in non_empty_str_values if "  " in s_val.strip()) 
            else: 
                na_percent = f"{0.0:.{dp}f}%"
                results['% Uppercase'] = na_percent; results['% Lowercase'] = na_percent
                results['% Titlecase'] = na_percent; results['% Mixed Case'] = na_percent
                results['Internal Multiple Spaces'] = 0
        else: 
            opt_na_msg = "N/A (Opt.)"
            results['% Uppercase'] = opt_na_msg; results['% Lowercase'] = opt_na_msg
            results['% Titlecase'] = opt_na_msg; results['% Mixed Case'] = opt_na_msg
            results['Internal Multiple Spaces'] = opt_na_msg


        results['Leading/Trailing Spaces'] = sum(1 for s_val in non_empty_str_values if s_val != s_val.strip())
        
        word_list = []
        for text in non_empty_str_values:
            cleaned_text = text.lower(); cleaned_text = re.sub(r'[^\w\s-]', '', cleaned_text) # Keep hyphens in words
            words = cleaned_text.split()
            word_list.extend([word for word in words if word and word not in STOP_WORDS and not word.isdigit()])
        if word_list:
            word_counts = Counter(word_list)
            top_words_list = [f"{word}:{count}" for word, count in word_counts.most_common(10)]
            results['Top Words'] = "\n".join(top_words_list) if top_words_list else "N/A"
        else: results['Top Words'] = "N/A (No words found)"
        
        email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
        # More robust URL pattern allowing various TLDs and paths
        url_pattern = r'https?://(?:[-\w.]|(?:%[\da-fA-F]{2}))+(?:/[-\w._/?#%&@!=कोंडीत]*)*' 
        emails_found = sum(1 for text in non_empty_str_values if re.search(email_pattern, text))
        urls_found = sum(1 for text in non_empty_str_values if re.search(url_pattern, text))
        results['Pattern Matches'] = f"Emails: {emails_found}, URLs: {urls_found}"
        
        return results

    def analyze_date_field_enhanced(self, original_variant_values, non_null_count, options):
        results = OrderedDict()
        dp = self.current_decimal_places 

        if non_null_count == 0:
            results['Status'] = 'No date data'
            for key in self.STAT_KEYS_DATE:
                 if key not in ['Non-Null Count', 'Null Count', '% Null', 'Status']:
                    if key in ['Dates Before Today', 'Dates After Today']: results[key] = 0
                    elif key.startswith('%') and ('Time' in key or 'Dates' in key): results[key] = f"{0.0:.{dp}f}%"
                    else: results[key] = 'N/A'
            return results

        py_datetimes = []    
        q_date_time_objects = [] # Store QDate or QDateTime objects for unique value counting
        
        # has_time_component_overall = False # Not strictly needed if we check instance type later

        for v_orig in original_variant_values: 
            if v_orig is None: continue 

            py_dt = None
            q_obj_for_unique = None # This will be QDate or QDateTime

            if isinstance(v_orig, QDateTime) and v_orig.isValid():
                py_dt = v_orig.toPyDateTime()
                q_obj_for_unique = v_orig 
                # has_time_component_overall = True
            elif isinstance(v_orig, QDate) and v_orig.isValid():
                # Convert QDate to Python datetime at midnight for consistent list type
                py_dt = datetime(v_orig.year(), v_orig.month(), v_orig.day()) 
                q_obj_for_unique = v_orig
            elif isinstance(v_orig, str): # Attempt to parse string dates if necessary
                # This part is tricky and depends on expected formats.
                # For now, assume QGIS provides correct QVariant types.
                # If string parsing is needed, it would go here with try-except blocks.
                pass

            if py_dt and q_obj_for_unique:
                py_datetimes.append(py_dt)
                q_date_time_objects.append(q_obj_for_unique)
        
        if not py_datetimes: 
            results['Status'] = 'No valid date objects parsed'
            for key in self.STAT_KEYS_DATE: 
                 if key not in ['Non-Null Count', 'Null Count', '% Null', 'Status']:
                    if key in ['Dates Before Today', 'Dates After Today']: results[key] = 0
                    elif key.startswith('%') and ('Time' in key or 'Dates' in key): results[key] = f"{0.0:.{dp}f}%"
                    else: results[key] = 'N/A'
            return results

        min_d, max_d = min(py_datetimes), max(py_datetimes)
        # Format based on whether time is present in the original QDateTime objects
        # This requires checking the original_variant_values or if any QDateTime was found
        is_datetime_field = any(isinstance(q_obj, QDateTime) for q_obj in q_date_time_objects)

        results['Min Date'] = min_d.isoformat(sep=' ', timespec='auto') if is_datetime_field else min_d.date().isoformat()
        results['Max Date'] = max_d.isoformat(sep=' ', timespec='auto') if is_datetime_field else max_d.date().isoformat()

        years = [d.year for d in py_datetimes]
        months = [d.month for d in py_datetimes] 
        days_of_week_num = [d.weekday() for d in py_datetimes] # Monday is 0 and Sunday is 6

        day_names_map = [self.tr("Mon"), self.tr("Tue"), self.tr("Wed"), self.tr("Thu"), self.tr("Fri"), self.tr("Sat"), self.tr("Sun")]
        month_names_map = ["", self.tr("Jan"), self.tr("Feb"), self.tr("Mar"), self.tr("Apr"), 
                           self.tr("May"), self.tr("Jun"), self.tr("Jul"), self.tr("Aug"), 
                           self.tr("Sep"), self.tr("Oct"), self.tr("Nov"), self.tr("Dec")]


        results['Common Years'] = ", ".join([f"{yr}:{cnt}" for yr, cnt in Counter(years).most_common(3)])
        results['Common Months'] = ", ".join([f"{month_names_map[mo]}:{cnt}" for mo, cnt in Counter(months).most_common(3)])
        results['Common Days'] = ", ".join([f"{day_names_map[d]}:{cnt}" for d, cnt in Counter(days_of_week_num).most_common(3)])
        
        today_pydate = datetime.now().date() # Compare dates only for "Before/After Today"
        results['Dates Before Today'] = sum(1 for d_py in py_datetimes if d_py.date() < today_pydate)
        results['Dates After Today'] = sum(1 for d_py in py_datetimes if d_py.date() > today_pydate)

        # Use the collected QDate/QDateTime objects for unique value counting
        date_counts = Counter(q_date_time_objects) 
        sorted_date_counts = sorted(date_counts.items(), key=lambda item: (-item[1], item[0])) 
        top_unique_dates_list = []; actual_first_unique_date_for_selection = None
        limit_unique = self.current_limit_unique_display
        if sorted_date_counts:
            actual_first_unique_date_for_selection = sorted_date_counts[0][0] # This is a QDate or QDateTime object
            for i, (date_obj, count) in enumerate(sorted_date_counts):
                if i >= limit_unique: break
                
                if isinstance(date_obj, QDateTime):
                    display_val_preview = date_obj.toString(Qt.ISODateWithMs if date_obj.time().msec() > 0 else Qt.ISODate)
                elif isinstance(date_obj, QDate):
                    display_val_preview = date_obj.toString(Qt.ISODate)
                else: # Should not happen if collection is correct
                    display_val_preview = str(date_obj)

                top_unique_dates_list.append(f"'{display_val_preview}': {count}")
        results['Unique Values (Top)'] = "\n".join(top_unique_dates_list) if top_unique_dates_list else "N/A"
        if top_unique_dates_list: # Store the QDate/QDateTime object
            results['Unique Values (Top)_actual_first_value'] = actual_first_unique_date_for_selection


        if options.get('date_time_weekend', False) and q_date_time_objects:
            midnight_count = 0
            noon_count = 0
            hours_list = []
            
            q_datetimes_only = [q_obj for q_obj in q_date_time_objects if isinstance(q_obj, QDateTime)]
            
            if q_datetimes_only: 
                for q_dt_obj in q_datetimes_only:
                    time_obj = q_dt_obj.time()
                    hours_list.append(time_obj.hour())
                    if time_obj == QTime(0,0,0,0): midnight_count +=1
                    if time_obj == QTime(12,0,0,0): noon_count +=1
                
                results['Common Hours (Top 3)'] = ", ".join([f"{hr:02d}:00 ({cnt})" for hr, cnt in Counter(hours_list).most_common(3)]) if hours_list else "N/A"
                results['% Midnight Time'] = f"{(midnight_count / len(q_datetimes_only) * 100.0):.{dp}f}%"
                results['% Noon Time'] = f"{(noon_count / len(q_datetimes_only) * 100.0):.{dp}f}%"
            else: 
                results['Common Hours (Top 3)'] = "N/A (No time data)"
                results['% Midnight Time'] = f"{0.0:.{dp}f}%" # Or N/A if preferred
                results['% Noon Time'] = f"{0.0:.{dp}f}%"
            
            all_q_dates_for_dow = []
            for q_obj in q_date_time_objects: # Use original QDate/QDateTime objects
                if isinstance(q_obj, QDateTime):
                    all_q_dates_for_dow.append(q_obj.date()) # Get QDate part
                elif isinstance(q_obj, QDate):
                    all_q_dates_for_dow.append(q_obj)
            
            if all_q_dates_for_dow:
                # QDate.dayOfWeek(): Monday = 1, ..., Sunday = 7
                weekend_day_count = sum(1 for d_obj in all_q_dates_for_dow if d_obj.dayOfWeek() >= 6) # Saturday or Sunday
                total_for_dow_calc = len(all_q_dates_for_dow)
                results['% Weekend Dates'] = f"{(weekend_day_count / total_for_dow_calc * 100.0):.{dp}f}%"
                results['% Weekday Dates'] = f"{((total_for_dow_calc - weekend_day_count) / total_for_dow_calc * 100.0):.{dp}f}%"
            else: 
                 results['% Weekend Dates'] = f"{0.0:.{dp}f}%"; results['% Weekday Dates'] = f"{0.0:.{dp}f}%"
        else: 
            opt_na_msg = "N/A (Opt.)"
            results['Common Hours (Top 3)'] = opt_na_msg; results['% Midnight Time'] = opt_na_msg; results['% Noon Time'] = opt_na_msg
            results['% Weekend Dates'] = opt_na_msg; results['% Weekday Dates'] = opt_na_msg
            
        return results

    def _on_cell_double_clicked(self, row, column):
        if column == 0:
            self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("Please double-click on a specific field value (cells to the right) to select features."), level=Qgis.Info)
            return

        current_layer = self.layerComboBox.currentLayer()
        if not current_layer or not isinstance(current_layer, QgsVectorLayer):
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("No valid layer selected."), level=Qgis.Warning); return

        stat_name_item = self.resultsTableWidget.item(row, 0) # Item in the first column (statistic name)
        field_header_item = self.resultsTableWidget.horizontalHeaderItem(column)

        if not stat_name_item or not field_header_item:
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Could not identify clicked cell data."), level=Qgis.Warning); return

        # --- Get the original (English) statistic key stored in the item's UserRole data ---
        original_statistic_key = stat_name_item.data(Qt.UserRole)
        if not original_statistic_key:
            # This should ideally not happen if populate_results_table correctly sets UserRole
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Internal error: Statistic key not found for the selected row."), level=Qgis.Critical)
            return

        field_name_for_selection = field_header_item.text()
        field_qobj = current_layer.fields().field(field_name_for_selection)
        if not field_qobj: 
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Field '{0}' not found in layer.").format(field_name_for_selection), level=Qgis.Warning); return


        quoted_field_name = f'"{field_name_for_selection}"' # QGIS expression-friendly field name
        expression = None
        ids_to_select_directly = None

        is_string_field = (field_qobj.type() == QVariant.String)
        is_numeric_field = field_qobj.isNumeric()
        # is_date_field = field_qobj.type() in [QVariant.Date, QVariant.DateTime] # For future use if needed

        # --- Logic based on the original_statistic_key ---
        if original_statistic_key == 'Null Count':
            expression = f"{quoted_field_name} IS NULL"
        elif original_statistic_key == 'Empty Strings' and is_string_field:
            expression = f"{quoted_field_name} = ''"
        elif original_statistic_key == 'Leading/Trailing Spaces' and is_string_field:
            # Select features where the original value is different from the trimmed value,
            # and the trimmed value is not empty (to avoid selecting empty strings that are also "just spaces")
            expression = f"{quoted_field_name} != trim({quoted_field_name}) AND length(trim({quoted_field_name})) > 0"
        elif original_statistic_key == 'Conversion Errors' and is_numeric_field:
            ids_to_select_directly = self.conversion_error_feature_ids_by_field.get(field_name_for_selection, [])
            if not ids_to_select_directly: 
                self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No features with conversion errors were recorded for this field."), level=Qgis.Info); return
        
        elif original_statistic_key == 'Non-Printable Chars Count' and is_string_field:
            ids_to_select_directly = self.non_printable_char_feature_ids_by_field.get(field_name_for_selection, [])
            if not ids_to_select_directly: 
                self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No features with non-printable characters were recorded for this field."), level=Qgis.Info); return

        elif original_statistic_key == 'Outliers (IQR)' and is_numeric_field:
            field_stats = self.analysis_results_cache.get(field_name_for_selection, {})
            q1_val = field_stats.get('Q1') 
            q3_val = field_stats.get('Q3')
            iqr_val = field_stats.get('IQR')
            
            # Check if all necessary values are valid numbers
            if isinstance(q1_val, (int, float)) and isinstance(q3_val, (int, float)) and isinstance(iqr_val, (int, float)) and \
               not (numpy.isnan(q1_val) or numpy.isnan(q3_val) or numpy.isnan(iqr_val)):
                lower_bound = q1_val - 1.5 * iqr_val
                upper_bound = q3_val + 1.5 * iqr_val
                expression = f"({quoted_field_name} < {lower_bound} OR {quoted_field_name} > {upper_bound}) AND {quoted_field_name} IS NOT NULL"
            else:
                self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("Q1, Q3, or IQR is N/A or invalid for outlier selection. Cannot create expression."), level=Qgis.Info); return

        elif original_statistic_key == 'Unique Values (Top)':
            cached_field_results = self.analysis_results_cache.get(field_name_for_selection, {})
            actual_first_value = cached_field_results.get('Unique Values (Top)_actual_first_value') # This is the raw value

            # Check if the special key exists. If not, means no top unique value was determined or cached.
            if 'Unique Values (Top)_actual_first_value' not in cached_field_results:
                self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("No specific unique value cached for selection. This might happen if all values were NULL or the field was empty."), level=Qgis.Info); return
            
            # actual_first_value CAN be None (representing a NULL in the data that was frequent)
            # or an empty string. These are valid for selection.
            if actual_first_value is None:
                 expression = f"{quoted_field_name} IS NULL" # Select NULLs if the top unique value was NULL
            elif isinstance(actual_first_value, str):
                escaped_val = actual_first_value.replace("'", "''") # Escape single quotes for SQL-like expression
                expression = f"{quoted_field_name} = '{escaped_val}'"
            elif isinstance(actual_first_value, (int, float, numpy.number)): 
                if numpy.isnan(actual_first_value): 
                    # Selecting NaN by direct equality in QGIS expressions is tricky.
                    # It's better to inform the user or select NULLs if NaN implies missing.
                    # For now, let's prevent selection of explicit NaNs this way.
                    self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("Cannot select NaN (Not a Number) unique value directly by this expression method. Consider selecting NULLs if appropriate."), level=Qgis.Info); return
                expression = f"{quoted_field_name} = {float(actual_first_value)}" # Ensure it's a Python float
            elif isinstance(actual_first_value, QDate):
                # QGIS expression functions for date/datetime: date('YYYY-MM-DD'), datetime('YYYY-MM-DD HH:MM:SS')
                expression = f"{quoted_field_name} = date('{actual_first_value.toString(Qt.ISODate)}')"
            elif isinstance(actual_first_value, QDateTime):
                # For QDateTime, QGIS expressions expect ISO format, potentially with time.
                # Qt.ISODate produces YYYY-MM-DDTHH:MM:SS
                # QGIS datetime() function usually takes 'YYYY-MM-DD HH:MM:SS.mmmZ'
                # Let's try with Qt.ISODate and see if QGIS handles it. Otherwise, more formatting needed.
                iso_string = actual_first_value.toString(Qt.ISODate) # e.g., "2023-10-26T10:30:00"
                # QGIS might prefer space separator for datetime()
                # expression_dt_string = actual_first_value.toString("yyyy-MM-dd HH:mm:ss.zzz") # More QGIS friendly
                expression = f"{quoted_field_name} = datetime('{iso_string}')"

            else:
                self.iface.messageBar().pushMessage(self.tr("Warning"), self.tr("Cannot select unique value of type: {0}. Selection for this type is not implemented.").format(type(actual_first_value).__name__), level=Qgis.Warning); return
        
        else:
             self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("Feature selection is not available for '{0}'.").format(self.tr(original_statistic_key)), level=Qgis.Info)
             return

        if expression:
            self._select_features_by_expression(current_layer, field_name_for_selection, expression)
        elif ids_to_select_directly is not None: # Check for not None, as empty list is valid
            self._select_features_by_ids(current_layer, field_name_for_selection, ids_to_select_directly)


    def _select_features_by_expression(self, layer, field_name, expression_string):
        try:
            selection_mode = QgsVectorLayer.SetSelection
            # If analysis was on selected features, subsequent selections should intersect
            # with the *original selection scope used for analysis*, not necessarily the *current layer selection*.
            # This is tricky. For now,IntersectSelection will intersect with current layer selection.
            # A more advanced approach would store the original FIDs if analysis was on selection.
            if self._was_analyzing_selected_features: # Intersect with current selection
                selection_mode = QgsVectorLayer.IntersectSelection
            
            num_selected = layer.selectByExpression(expression_string, selection_mode)
            self.iface.mapCanvas().refresh() # Refresh map
            # Try to make the attribute table update if open and layer matches
            if self.iface.attributesToolBar() and self.iface.attributesToolBar().isVisible():
                for table_view in self.iface.mainWindow().findChildren(QgsAttributeTable): # QgsAttributeTable may not be directly accessible
                    if table_view.layer() == layer:
                        table_view.doSelect(layer.selectedFeatureIds()) # This is a guess, API might differ
                        break
            
            if hasattr(self.iface, 'actionOpenTable') and self.iface.actionOpenTable().isEnabled():
                # This is a generic way to try and get attention to selection
                pass


            msg = self.tr("Selected {0} features for field '{1}' where: {2}").format(num_selected, field_name, expression_string)
            if self._was_analyzing_selected_features and selection_mode == QgsVectorLayer.IntersectSelection:
                 msg += self.tr(" (Intersected with current layer selection).")
            else:
                 msg += "."
            self.iface.messageBar().pushMessage(self.tr("Selection Succeeded"), msg, level=Qgis.Success, duration=7)

        except Exception as e:
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Error selecting features by expression: {0}\nExpression: {1}").format(str(e), expression_string), level=Qgis.Critical)

    def _select_features_by_ids(self, layer, field_name, fids_to_select):
        try:
            num_selected = 0
            final_ids_for_selection = list(fids_to_select) # Ensure it's a list

            if not final_ids_for_selection: # No IDs to select
                 self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No feature IDs provided for selection."), level=Qgis.Info, duration=5)
                 return


            if self._was_analyzing_selected_features:
                # Intersect the provided FIDs with the current selection on the layer
                current_selection_on_layer = set(layer.selectedFeatureIds())
                ids_to_actually_select = [fid for fid in final_ids_for_selection if fid in current_selection_on_layer]
                
                layer.selectByIds(ids_to_actually_select, QgsVectorLayer.SetSelection) # Replace current selection with the intersection
                num_selected = len(ids_to_actually_select)
                msg_suffix = self.tr(" (Intersected with current layer selection).")
            else:
                layer.selectByIds(final_ids_for_selection, QgsVectorLayer.SetSelection) # Set new selection
                num_selected = len(final_ids_for_selection)
                msg_suffix = "."
            
            self.iface.mapCanvas().refresh()
            # Similar attribute table update attempt as in _select_features_by_expression
            if self.iface.attributesToolBar() and self.iface.attributesToolBar().isVisible():
                 for table_view in self.iface.mainWindow().findChildren(QgsAttributeTable):
                    if table_view.layer() == layer:
                        table_view.doSelect(layer.selectedFeatureIds())
                        break
            
            msg = self.tr("Selected {0} features for field '{1}' based on stored IDs{2}").format(num_selected, field_name, msg_suffix)
            self.iface.messageBar().pushMessage(self.tr("Selection Succeeded"), msg, level=Qgis.Success, duration=7)
        except Exception as e:
            self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Error selecting features by IDs: {0}").format(str(e)), level=Qgis.Critical)

    def copy_results_to_clipboard(self):
        if self.resultsTableWidget.rowCount() == 0 or self.resultsTableWidget.columnCount() == 0:
            self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No results to copy."), level=Qgis.Info); return
        clipboard = QApplication.clipboard()
        if not clipboard:
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Could not access clipboard."), level=Qgis.Critical); return
        output = ""
        # Headers
        headers = [self.resultsTableWidget.horizontalHeaderItem(c).text() for c in range(self.resultsTableWidget.columnCount())]
        output += "\t".join(headers) + "\n"
        # Data rows
        for r in range(self.resultsTableWidget.rowCount()):
            row_data = [];
            for c in range(self.resultsTableWidget.columnCount()):
                 item = self.resultsTableWidget.item(r, c)
                 # Replace newlines in cell text to keep CSV/TSV structure clean
                 cell_text = item.text().replace("\n", " | ") if item else ""
                 row_data.append(cell_text)
            output += "\t".join(row_data) + "\n"
        clipboard.setText(output)
        self.iface.messageBar().pushMessage(self.tr("Success"), self.tr("Table results copied to clipboard."), level=Qgis.Success)

    def export_results_to_csv(self):
        if self.resultsTableWidget.rowCount() == 0 or self.resultsTableWidget.columnCount() == 0:
            self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No results to export."), level=Qgis.Info); return
        
        default_filename = "field_profiler_results.csv"
        current_qgs_layer = self.layerComboBox.currentLayer()
        if current_qgs_layer: 
            layer_name_sanitized = re.sub(r'[^\w\.-]', '_', current_qgs_layer.name()) # Sanitize layer name
            default_filename = f"{layer_name_sanitized}_profile.csv"
            
        file_path, _ = QFileDialog.getSaveFileName(self, self.tr("Export Results to CSV"), default_filename, self.tr("CSV Files (*.csv);;All Files (*)"))
        
        if not file_path: return # User cancelled
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile: # utf-8-sig for Excel compatibility with BOM
                writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                # Headers
                headers = [self.resultsTableWidget.horizontalHeaderItem(c).text() for c in range(self.resultsTableWidget.columnCount())]
                writer.writerow(headers)
                # Data rows
                for r in range(self.resultsTableWidget.rowCount()):
                    row_data = [];
                    for c in range(self.resultsTableWidget.columnCount()):
                        item = self.resultsTableWidget.item(r, c)
                        # Replace newlines in cell text to keep CSV structure clean
                        cell_text = item.text().replace("\n", " | ") if item else ""
                        row_data.append(cell_text)
                    writer.writerow(row_data)
            self.iface.messageBar().pushMessage(self.tr("Success"), self.tr("Results successfully exported to CSV: {0}").format(file_path), level=Qgis.Success)
        except Exception as e: 
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Could not export results to CSV: ") + str(e), level=Qgis.Critical)
            print(f"CSV Export Error: {e}") # Log to console for debugging

    def closeEvent(self, event):
        # Store settings on close? (e.g., self.iface.pluginVsSettings().setValue(...))
        self.hide()
        event.ignore() # Important for dock widgets: hide instead of close/delete
