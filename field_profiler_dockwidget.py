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
                                 QSpinBox, QFormLayout, QPlainTextEdit, QHeaderView, QDialog, QComboBox)
from qgis.gui import QgsMapLayerComboBox
from qgis.core import (QgsProject, QgsVectorLayer, QgsField, Qgis,
                       QgsStatisticalSummary, QgsMapLayerProxyModel, QgsFeatureRequest,
                       QgsExpression, QgsTask, QgsApplication, QgsMessageLog)

import statistics
from collections import Counter, OrderedDict
import numpy # Keep this import
from datetime import datetime
from .field_profiler_task import FieldProfilerTask
from .report_generator import ReportGenerator


SCIPY_AVAILABLE = False
try:
    from scipy import stats as scipy_stats
    SCIPY_AVAILABLE = True
except ImportError:
    scipy_stats = None # So we can check against it

MATPLOTLIB_AVAILABLE = False
try:
    import matplotlib
    matplotlib.use('Qt5Agg') # Or QtAgg depending on environment, try safe default or check
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    try:
         # Try QtAgg for newer matplotlib/PyQt6 if Qt5Agg fails
        import matplotlib
        matplotlib.use('QtAgg')
        from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
        from matplotlib.figure import Figure
        import matplotlib.pyplot as plt
        MATPLOTLIB_AVAILABLE = True
    except ImportError:
        pass


STOP_WORDS = set([
    'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from', 'has', 'he',
    'in', 'is', 'it', 'its', 'of', 'on', 'that', 'the', 'to', 'was', 'were',
    'will', 'with',
])

class AnalysisResultsDialog(QDialog):
    """
    Floating dialog to display analysis results (Table, Charts, Correlation, Validation).
    """
    def __init__(self, parent=None, results_data=None, layer=None, detailed_options=None, was_analyzing_selection=False):
        super().__init__(parent)
        self.setWindowTitle(self.tr("Analysis Results"))
        self.resize(1000, 700) # Default size
        
        self.results_data = results_data
        self.layer = layer
        self.detailed_options = detailed_options or {}
        self.was_analyzing_selected_features = was_analyzing_selection
        self.analysis_results_cache = results_data if results_data else OrderedDict()
        
        # Helper dictionaries for selection (populated from results)
        self.conversion_error_feature_ids_by_field = {}
        self.non_printable_char_feature_ids_by_field = {}
        
        # Access iface from parent if possible
        self.iface = None
        if parent and hasattr(parent, 'iface'):
             self.iface = parent.iface
        
        self._init_ui()
        self._process_results()

    def tr(self, message):
         return QtCore.QCoreApplication.translate("AnalysisResultsDialog", message)

    def _init_ui(self):
        self.main_layout = QVBoxLayout(self)
        
        self.tabs = QtWidgets.QTabWidget()
        
        # --- Tab 1: Table ---
        self.tab_table = QWidget()
        table_layout = QVBoxLayout(self.tab_table)
        self.resultsTableWidget = QTableWidget()
        self.resultsTableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.resultsTableWidget.setAlternatingRowColors(True)
        self.resultsTableWidget.setSortingEnabled(True) 
        table_layout.addWidget(self.resultsTableWidget)
        
        button_layout = QHBoxLayout()
        self.copyButton = QPushButton(self.tr("Copy Table"))
        self.exportButton = QPushButton(self.tr("Export CSV"))
        self.exportHtmlButton = QPushButton(self.tr("Export HTML"))
        
        self.copyButton.clicked.connect(self.copy_results_to_clipboard)
        self.exportButton.clicked.connect(self.export_results_to_csv)
        self.exportHtmlButton.clicked.connect(self.export_results_to_html)

        button_layout.addStretch()
        button_layout.addWidget(self.copyButton)
        button_layout.addWidget(self.exportButton)
        button_layout.addWidget(self.exportHtmlButton)
        table_layout.addLayout(button_layout)
        
        self.tabs.addTab(self.tab_table, self.tr("Table"))
        
        # --- Tab 2: Charts ---
        self.tab_charts = QWidget()
        self.charts_layout = QVBoxLayout(self.tab_charts)
        
        if MATPLOTLIB_AVAILABLE:
            self.figure = Figure()
            self.canvas = FigureCanvas(self.figure)
            
            # Field Selector for Charts
            self.fieldSelector = QComboBox()
            self.fieldSelector.currentTextChanged.connect(self.update_charts_from_selector)
            self.charts_layout.addWidget(self.fieldSelector)
            
            self.charts_layout.addWidget(self.canvas)
            self.chart_info_label = QLabel(self.tr("Select a field column in the table to view charts."))
            self.chart_info_label.setAlignment(Qt.AlignCenter)
            self.charts_layout.addWidget(self.chart_info_label)
        else:
            self.charts_layout.addWidget(QLabel(self.tr("Matplotlib not installed. Charts unavailable.")))
            self.canvas = None

        self.tabs.addTab(self.tab_charts, self.tr("Charts"))
        
        # --- Tab 3: Correlation ---
        self.tab_correlation = QWidget()
        corr_layout = QVBoxLayout(self.tab_correlation)
        self.correlationTableWidget = QTableWidget()
        self.correlationTableWidget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.correlationTableWidget.setAlternatingRowColors(False)
        corr_layout.addWidget(self.correlationTableWidget)
        self.tabs.addTab(self.tab_correlation, self.tr("Correlation"))
        
        # --- Tab 4: Validation Results ---
        self.tab_validation = QWidget()
        val_res_layout = QVBoxLayout(self.tab_validation)
        self.validationResultsTable = QTableWidget()
        self.validationResultsTable.setColumnCount(3)
        self.validationResultsTable.setHorizontalHeaderLabels([self.tr("Rule"), self.tr("Fail Count"), self.tr("% Fail")])
        self.validationResultsTable.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        val_res_layout.addWidget(self.validationResultsTable)
        self.tabs.addTab(self.tab_validation, self.tr("Validation"))
        
        self.main_layout.addWidget(self.tabs)

        # Connect Selection
        self.resultsTableWidget.cellDoubleClicked.connect(self._on_cell_double_clicked)
        self.resultsTableWidget.itemSelectionChanged.connect(self.update_charts)
        self.resultsTableWidget.horizontalHeader().sectionClicked.connect(self.update_charts)

    def _process_results(self):
        if not self.results_data:
            return

        results = self.results_data
        
        # Update helper dictionaries
        for fname, field_data in results.items():
            if '_conversion_error_fids' in field_data:
                self.conversion_error_feature_ids_by_field[fname] = field_data['_conversion_error_fids']
                del field_data['_conversion_error_fids']
            if '_non_printable_fids' in field_data:
                self.non_printable_char_feature_ids_by_field[fname] = field_data['_non_printable_fids']
                del field_data['_non_printable_fids']
        
        # Populate Correlation
        if '_global_correlation' in results:
            self.latest_correlation_matrix = results['_global_correlation'] # Store for report
            self._populate_correlation_matrix(results['_global_correlation'])
            del results['_global_correlation']
        else:
            self.latest_correlation_matrix = None
            self.correlationTableWidget.clear()
            self.correlationTableWidget.setRowCount(0)
            self.correlationTableWidget.setColumnCount(0)
            
        # Handle Validation Results
        if '_validation_results' in results:
            self._populate_validation_results(results['_validation_results'])
            del results['_validation_results']
        else:
            self.validationResultsTable.clear()
            self.validationResultsTable.setRowCount(0)

        # Populate table
        field_names = list(results.keys())
        self.populate_results_table(self.analysis_results_cache, field_names)

        # Populate Chart Selector
        if self.canvas:
            self.fieldSelector.blockSignals(True)
            self.fieldSelector.clear()
            self.fieldSelector.addItems(field_names)
            self.fieldSelector.blockSignals(False)
            
            # Default to first numeric field if available
            for fname in field_names:
                f_data = results.get(fname, {})
                if '_histogram_data' in f_data:
                    self.fieldSelector.setCurrentText(fname)
                    self.update_charts_from_selector(fname) # Force update
                    break
            else:
                 if field_names: 
                     self.fieldSelector.setCurrentIndex(0)
                     self.update_charts_from_selector(field_names[0])

    def populate_results_table(self, results_data, field_names_for_header):
        self.resultsTableWidget.clear()
        if not results_data and not field_names_for_header: return
        all_stat_names_from_data = set()
        for field_name, field_data in results_data.items(): all_stat_names_from_data.update(field_data.keys())
        
        # Filter out internal keys like '_actual_first_value' AND '_histogram_data'
        all_displayable_stat_names = {stat for stat in all_stat_names_from_data 
                                      if not stat.endswith('_actual_first_value') 
                                      and stat != '_histogram_data'
                                      and stat != '_top_values_raw'}

        # --- Determine row order for statistics ---
        stat_rows_ordered = [] 
        seen_keys_for_order = set()
        
        STAT_KEYS_NUMERIC = FieldProfilerDockWidget.STAT_KEYS_NUMERIC
        STAT_KEYS_TEXT = FieldProfilerDockWidget.STAT_KEYS_TEXT
        STAT_KEYS_DATE = FieldProfilerDockWidget.STAT_KEYS_DATE
        STAT_KEYS_OTHER = FieldProfilerDockWidget.STAT_KEYS_OTHER
        STAT_KEYS_ERROR = FieldProfilerDockWidget.STAT_KEYS_ERROR
        
        predefined_order_source = []
        temp_seen_for_predefined_order = set()
        for key_list in [STAT_KEYS_NUMERIC, STAT_KEYS_TEXT, STAT_KEYS_DATE, 
                         STAT_KEYS_OTHER, STAT_KEYS_ERROR]:
            for key in key_list:
                if key not in temp_seen_for_predefined_order:
                    predefined_order_source.append(key)
                    temp_seen_for_predefined_order.add(key)

        for key in predefined_order_source:
            if key in all_displayable_stat_names and key not in seen_keys_for_order:
                stat_rows_ordered.append(key)
                seen_keys_for_order.add(key)
        
        extras = sorted([key for key in all_displayable_stat_names if key not in seen_keys_for_order])
        stat_rows_ordered.extend(extras)
        
        num_rows = len(stat_rows_ordered)
        num_cols = len(field_names_for_header) + 1 
        self.resultsTableWidget.setRowCount(num_rows)
        self.resultsTableWidget.setColumnCount(num_cols)
        
        headers = [self.tr("Statistic")] + field_names_for_header
        self.resultsTableWidget.setHorizontalHeaderLabels(headers)
        
        quality_keywords = ['%', 'Null', 'Empty', 'Error', 'Outlier', 'Spaces', 'Variance', 'Flag', 'Conversion', 'Mismatch', 'Non-Printable'] 
        dp = self.detailed_options.get('decimal_places', 2)
        
        for r, original_stat_key in enumerate(stat_rows_ordered):
            stat_item = QTableWidgetItem(self.tr(original_stat_key))
            stat_item.setData(Qt.UserRole, original_stat_key)
            stat_item.setToolTip(original_stat_key) 
            
            is_quality_issue = any(keyword.lower() in original_stat_key.lower() for keyword in quality_keywords) or \
                               original_stat_key == 'Error'
            
            first_field_name_for_color = field_names_for_header[0] if field_names_for_header else None
            if first_field_name_for_color:
                 first_field_data = results_data.get(first_field_name_for_color, {})
                 if original_stat_key == 'Normality (Likely Normal)' and first_field_data.get(original_stat_key) is False:
                     is_quality_issue = True
                 if original_stat_key == 'Low Variance Flag' and first_field_data.get(original_stat_key) is True:
                     is_quality_issue = True

            if is_quality_issue:
                stat_item.setBackground(QtGui.QColor(255, 240, 240)) 
            elif original_stat_key.startswith('%') or "Pctl" in original_stat_key or original_stat_key in ['Skewness', 'Kurtosis']:
                stat_item.setBackground(QtGui.QColor(240, 240, 255)) 
            else:
                stat_item.setBackground(QtGui.QColor(230, 230, 230))
            
            self.resultsTableWidget.setItem(r, 0, stat_item)

            for c, field_name in enumerate(field_names_for_header):
                field_data = results_data.get(field_name, {})
                value = field_data.get(original_stat_key, "")
                display_text = ""
                
                if isinstance(value, bool):
                    display_text = str(value)
                elif isinstance(value, float):
                    if original_stat_key == 'Normality (Shapiro-Wilk p)':
                         display_text = f"{value:.4g}" if not numpy.isnan(value) else "N/A"
                    else:
                         display_text = f"{value:.{dp}f}" if not numpy.isnan(value) else "N/A"
                elif isinstance(value, list) and original_stat_key != 'Mode(s)':
                    display_text = "; ".join(map(str, value))
                elif isinstance(value, list) and original_stat_key == 'Mode(s)':
                    formatted_modes = []
                    for v_mode in value:
                        if isinstance(v_mode, (int, float)):
                            try:
                                formatted_modes.append(f"{float(v_mode):.{dp}f}")
                            except ValueError:
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
                              any(kw in original_stat_key for kw in align_right_keywords) 

                item.setTextAlignment(Qt.AlignVCenter | (Qt.AlignRight if align_right else Qt.AlignLeft))
                
                if isinstance(value, str) and ('\n' in value or len(value) > 60):
                    item.setToolTip(value)
                elif item.text() in ["N/A (Scipy not found)", "N/A (>=3 values needed)", "N/A (<3 valid)"]:
                    item.setForeground(QtGui.QBrush(Qt.gray))

                self.resultsTableWidget.setItem(r, c + 1, item)
        
        self.resultsTableWidget.resizeColumnsToContents()

    def _populate_correlation_matrix(self, corr_data):
        self.correlationTableWidget.clear()
        if 'Error' in corr_data: 
            self.correlationTableWidget.setRowCount(1); self.correlationTableWidget.setColumnCount(1)
            item = QTableWidgetItem(corr_data['Error'])
            item.setForeground(QtGui.QBrush(Qt.red))
            self.correlationTableWidget.setItem(0, 0, item)
            self.correlationTableWidget.resizeColumnsToContents()
            return
            
        fields = corr_data.get('fields', [])
        matrix = corr_data.get('matrix', [])
        
        if not fields or not matrix: return
        
        n = len(fields)
        self.correlationTableWidget.setRowCount(n)
        self.correlationTableWidget.setColumnCount(n)
        self.correlationTableWidget.setHorizontalHeaderLabels(fields)
        self.correlationTableWidget.setVerticalHeaderLabels(fields)
        
        for r in range(n):
            for c in range(n):
                val = matrix[r][c]
                item = QTableWidgetItem(f"{val:.2f}")
                item.setTextAlignment(Qt.AlignCenter)
                
                bg_color = QtGui.QColor(255, 255, 255)
                if val > 0:
                    intensity = int(255 * (1 - abs(val)))
                    bg_color = QtGui.QColor(intensity, intensity, 255)
                elif val < 0:
                    intensity = int(255 * (1 - abs(val)))
                    bg_color = QtGui.QColor(255, intensity, intensity)
                
                item.setBackground(bg_color)
                if abs(val) > 0.6: item.setForeground(QtGui.QColor(255, 255, 255) if val > 0 else QtGui.QColor(0,0,0))
                
                self.correlationTableWidget.setItem(r, c, item)
        
        self.correlationTableWidget.resizeColumnsToContents()

    def _populate_validation_results(self, val_data):
        self.validationResultsTable.clearContents()
        self.validationResultsTable.setRowCount(0)
        
        if not val_data: return
        
        rules = val_data.get('rules', [])
        counts = val_data.get('fail_counts', [])
        total = val_data.get('total_checked', 0)
        
        self.validationResultsTable.setRowCount(len(rules))
        for i, rule in enumerate(rules):
            count = counts[i]
            pct = (count / total * 100) if total > 0 else 0
            
            item_rule = QTableWidgetItem(rule)
            item_count = QTableWidgetItem(str(count))
            item_pct = QTableWidgetItem(f"{pct:.2f}%")
            
            if count > 0:
                item_count.setForeground(QtGui.QColor(255, 0, 0)) 
            
            self.validationResultsTable.setItem(i, 0, item_rule)
            self.validationResultsTable.setItem(i, 1, item_count)
            self.validationResultsTable.setItem(i, 2, item_pct)

    def _on_cell_double_clicked(self, row, column):
        if column == 0:
            if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("Please double-click on a specific field value (cells to the right) to select features."), level=Qgis.Info)
            return

        current_layer = self.layer
        if not current_layer or not isinstance(current_layer, QgsVectorLayer):
             if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("No valid layer available."), level=Qgis.Warning)
             return

        stat_name_item = self.resultsTableWidget.item(row, 0)
        field_header_item = self.resultsTableWidget.horizontalHeaderItem(column)

        if not stat_name_item or not field_header_item: return

        original_statistic_key = stat_name_item.data(Qt.UserRole)
        field_name_for_selection = field_header_item.text()
        field_qobj = current_layer.fields().field(field_name_for_selection)
        if not field_qobj: return

        quoted_field_name = QgsExpression.quotedColumnRef(field_name_for_selection)
        expression = ""
        ids_to_select_directly = []

        is_string_field = (field_qobj.type() == QVariant.String)
        is_numeric_field = field_qobj.isNumeric()

        if original_statistic_key == 'Null Count':
            expression = f"{quoted_field_name} IS NULL"
        elif original_statistic_key == 'Empty Strings' and is_string_field:
            expression = f"{quoted_field_name} = ''"
        elif original_statistic_key == 'Leading/Trailing Spaces' and is_string_field:
            expression = f"{quoted_field_name} != trim({quoted_field_name}) AND length(trim({quoted_field_name})) > 0"
        elif original_statistic_key == 'Conversion Errors' and is_numeric_field:
            ids_to_select_directly = self.conversion_error_feature_ids_by_field.get(field_name_for_selection, [])
            if not ids_to_select_directly: 
                if self.iface: self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No features with conversion errors were recorded for this field."), level=Qgis.Info); return
        
        elif original_statistic_key == 'Non-Printable Chars Count' and is_string_field:
            ids_to_select_directly = self.non_printable_char_feature_ids_by_field.get(field_name_for_selection, [])
            if not ids_to_select_directly: 
                if self.iface: self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No features with non-printable characters were recorded for this field."), level=Qgis.Info); return

        elif original_statistic_key == 'Outliers (IQR)' and is_numeric_field:
             field_stats = self.analysis_results_cache.get(field_name_for_selection, {})
             q1_val = field_stats.get('Q1')
             q3_val = field_stats.get('Q3')
             iqr_val = field_stats.get('IQR')
             if all(isinstance(x, (int, float)) and not numpy.isnan(x) for x in [q1_val, q3_val, iqr_val]):
                 lower = q1_val - 1.5 * iqr_val
                 upper = q3_val + 1.5 * iqr_val
                 expression = f"({quoted_field_name} < {lower} OR {quoted_field_name} > {upper}) AND {quoted_field_name} IS NOT NULL"
             else:
                 if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("Q1, Q3, or IQR is N/A or invalid for outlier selection."), level=Qgis.Info); return

        elif original_statistic_key == 'Unique Values (Top)':
            cached_field_results = self.analysis_results_cache.get(field_name_for_selection, {})
            actual_first_value = cached_field_results.get('Unique Values (Top)_actual_first_value')

            if 'Unique Values (Top)_actual_first_value' not in cached_field_results:
                 if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("No specific unique value cached for selection."), level=Qgis.Info); return
            
            if actual_first_value is None:
                 expression = f"{quoted_field_name} IS NULL"
            elif isinstance(actual_first_value, str):
                if hasattr(QgsExpression, 'quotedValue'):
                    expression = f"{quoted_field_name} = {QgsExpression.quotedValue(actual_first_value)}"
                else: 
                     escaped_val = actual_first_value.replace("'", "''") 
                     expression = f"{quoted_field_name} = '{escaped_val}'"
            elif isinstance(actual_first_value, (int, float, numpy.number)): 
                if numpy.isnan(actual_first_value): 
                    if self.iface: self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("Cannot select NaN (Not a Number) directly."), level=Qgis.Info); return
                expression = f"{quoted_field_name} = {float(actual_first_value)}"
            elif isinstance(actual_first_value, QDate):
                if hasattr(QgsExpression, 'quotedValue'):
                    expression = f"{quoted_field_name} = {QgsExpression.quotedValue(actual_first_value)}"
                else:
                    expression = f"{quoted_field_name} = date('{actual_first_value.toString(Qt.ISODate)}')"
            elif isinstance(actual_first_value, QDateTime):
                if hasattr(QgsExpression, 'quotedValue'):
                    expression = f"{quoted_field_name} = {QgsExpression.quotedValue(actual_first_value)}"
                else:
                    iso_string = actual_first_value.toString(Qt.ISODate) 
                    expression = f"{quoted_field_name} = datetime('{iso_string}')"
            else:
                 if self.iface: self.iface.messageBar().pushMessage(self.tr("Warning"), self.tr("Cannot select unique value of type: {0}").format(type(actual_first_value).__name__), level=Qgis.Warning); return
        
        else:
              if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Info"), self.tr("Feature selection is not available for '{0}'.").format(self.tr(original_statistic_key)), level=Qgis.Info); return

        if expression:
            self._select_features_by_expression(current_layer, field_name_for_selection, expression)
        elif ids_to_select_directly is not None:
            self._select_features_by_ids(current_layer, field_name_for_selection, ids_to_select_directly)


    def _select_features_by_expression(self, layer, field_name, expression_string):
        try:
            selection_mode = QgsVectorLayer.SetSelection
            if self.was_analyzing_selected_features:
                selection_mode = QgsVectorLayer.IntersectSelection
            
            num_selected = layer.selectByExpression(expression_string, selection_mode)
            if self.iface: self.iface.mapCanvas().refresh()
            
            # Attributes table sync attempt
            if self.iface and self.iface.attributesToolBar() and self.iface.attributesToolBar().isVisible():
                 for table_view in self.iface.mainWindow().findChildren(QgsAttributeTable):
                    if table_view.layer() == layer:
                        table_view.doSelect(layer.selectedFeatureIds())
                        break
            
            msg = self.tr("Selected {0} features for field '{1}' where: {2}").format(num_selected, field_name, expression_string)
            if self.was_analyzing_selected_features and selection_mode == QgsVectorLayer.IntersectSelection:
                 msg += self.tr(" (Intersected with current layer selection).")
            else:
                 msg += "."
            if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Succeeded"), msg, level=Qgis.Success, duration=7)

        except Exception as e:
             if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Error selecting features by expression: {0}").format(str(e)), level=Qgis.Critical)

    def _select_features_by_ids(self, layer, field_name, fids_to_select):
        try:
            num_selected = 0
            final_ids_for_selection = list(fids_to_select)

            if not final_ids_for_selection:
                 if self.iface: self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No feature IDs provided for selection."), level=Qgis.Info, duration=5); return

            if self.was_analyzing_selected_features:
                current_selection_on_layer = set(layer.selectedFeatureIds())
                ids_to_actually_select = [fid for fid in final_ids_for_selection if fid in current_selection_on_layer]
                layer.selectByIds(ids_to_actually_select, QgsVectorLayer.SetSelection)
                num_selected = len(ids_to_actually_select)
                msg_suffix = self.tr(" (Intersected with current layer selection).")
            else:
                layer.selectByIds(final_ids_for_selection, QgsVectorLayer.SetSelection)
                num_selected = len(final_ids_for_selection)
                msg_suffix = "."
            
            if self.iface: self.iface.mapCanvas().refresh()
            
            if self.iface and self.iface.attributesToolBar() and self.iface.attributesToolBar().isVisible():
                 for table_view in self.iface.mainWindow().findChildren(QgsAttributeTable):
                    if table_view.layer() == layer:
                        table_view.doSelect(layer.selectedFeatureIds())
                        break
            
            msg = self.tr("Selected {0} features for field '{1}' based on stored IDs{2}").format(num_selected, field_name, msg_suffix)
            if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Succeeded"), msg, level=Qgis.Success, duration=7)
        except Exception as e:
            if self.iface: self.iface.messageBar().pushMessage(self.tr("Selection Error"), self.tr("Error selecting features by IDs: {0}").format(str(e)), level=Qgis.Critical)

    def copy_results_to_clipboard(self):
        if self.resultsTableWidget.rowCount() == 0 or self.resultsTableWidget.columnCount() == 0:
            if self.iface: self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No results to copy."), level=Qgis.Info); return
        clipboard = QApplication.clipboard()
        if not clipboard:
             if self.iface: self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Could not access clipboard."), level=Qgis.Critical); return
        output = ""
        headers = [self.resultsTableWidget.horizontalHeaderItem(c).text() for c in range(self.resultsTableWidget.columnCount())]
        output += "\t".join(headers) + "\n"
        for r in range(self.resultsTableWidget.rowCount()):
            row_data = []
            for c in range(self.resultsTableWidget.columnCount()):
                 item = self.resultsTableWidget.item(r, c)
                 cell_text = item.text().replace("\n", " | ") if item else ""
                 row_data.append(cell_text)
            output += "\t".join(row_data) + "\n"
        clipboard.setText(output)
        if self.iface: self.iface.messageBar().pushMessage(self.tr("Success"), self.tr("Table results copied to clipboard."), level=Qgis.Success)

    def export_results_to_csv(self):
        if self.resultsTableWidget.rowCount() == 0 or self.resultsTableWidget.columnCount() == 0:
             if self.iface: self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("No results to export."), level=Qgis.Info); return
        
        default_filename = "field_profiler_results.csv"
        current_qgs_layer = self.layer
        if current_qgs_layer: 
            layer_name_sanitized = re.sub(r'[^\w\.-]', '_', current_qgs_layer.name())
            default_filename = f"{layer_name_sanitized}_profile.csv"
            
        file_path, _ = QFileDialog.getSaveFileName(self, self.tr("Export Results to CSV"), default_filename, self.tr("CSV Files (*.csv);;All Files (*)"))
        if not file_path: return
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                writer = csv.writer(csvfile, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
                headers = [self.resultsTableWidget.horizontalHeaderItem(c).text() for c in range(self.resultsTableWidget.columnCount())]
                writer.writerow(headers)
                for r in range(self.resultsTableWidget.rowCount()):
                    row_data = []
                    for c in range(self.resultsTableWidget.columnCount()):
                        item = self.resultsTableWidget.item(r, c)
                        cell_text = item.text().replace("\n", " | ") if item else ""
                        row_data.append(cell_text)
                    writer.writerow(row_data)
            if self.iface: self.iface.messageBar().pushMessage(self.tr("Success"), self.tr("Results successfully exported to CSV: {0}").format(file_path), level=Qgis.Success)
        except Exception as e: 
             if self.iface: self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Could not export results to CSV: ") + str(e), level=Qgis.Critical)

    def export_results_to_html(self):
        if not self.analysis_results_cache:
            if self.iface: self.iface.messageBar().pushMessage(self.tr("Warning"), self.tr("No analysis results to export."), level=Qgis.Warning); return

        filename, filter = QFileDialog.getSaveFileName(self, self.tr("Export HTML Report"), "", self.tr("HTML Files (*.html)"))
        if not filename: return
        if not filename.lower().endswith('.html'): filename += '.html'
        
        try:
            current_layer = self.layer
            layer_name = current_layer.name() if current_layer else "Unknown Layer"
            
            generator = ReportGenerator(layer_name)
            html_content = generator.generate_report(self.analysis_results_cache, getattr(self, 'latest_correlation_matrix', None))
            
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(html_content)
                
            if self.iface: self.iface.messageBar().pushMessage(self.tr("Success"), self.tr("Report exported successfully."), level=Qgis.Success)
            
        except Exception as e:
             if self.iface: self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Could not export HTML: ") + str(e), level=Qgis.Critical)

    def update_charts_from_selector(self, field_name):
        """Wrapper to update charts when combo box changes."""
        self.update_charts(field_name_override=field_name)

    def update_charts(self, clicked_column_index=None, field_name_override=None):
        if not MATPLOTLIB_AVAILABLE or not self.canvas: return
        
        field_name = None
        
        # Priority 1: Direct Name Override (from Combo Box)
        if field_name_override:
            field_name = field_name_override
            # update selection in table to match if found? 
            # might be annoying if user just wants to browse charts. 
            # Let's just sync the combo box state to match.
        
        # Priority 2: Column Index (from Table Header Click)
        elif isinstance(clicked_column_index, int):
            col = clicked_column_index
            if col > 0:
                 header = self.resultsTableWidget.horizontalHeaderItem(col)
                 if header: field_name = header.text()
        
        # Priority 3: Selection (from Table Cell Click)
        else:
            selected_items = self.resultsTableWidget.selectedItems()
            if selected_items:
                col = selected_items[0].column()
                if col > 0:
                     header = self.resultsTableWidget.horizontalHeaderItem(col)
                     if header: field_name = header.text()

        if not field_name: return
        
        # Sync Combo Box if it didn't trigger this
        if self.fieldSelector.currentText() != field_name:
            self.fieldSelector.blockSignals(True)
            self.fieldSelector.setCurrentText(field_name)
            self.fieldSelector.blockSignals(False)

        field_data = self.analysis_results_cache.get(field_name)
        if not field_data: return
        
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        QgsMessageLog.logMessage(f"Updating chart for {field_name}. Keys: {list(field_data.keys())}", "FieldProfiler", Qgis.Info)

        if '_histogram_data' in field_data:
            hist_counts, bin_edges = field_data['_histogram_data']
            width = numpy.diff(bin_edges)
            ax.bar(bin_edges[:-1], hist_counts, width=width, align='edge', alpha=0.7)
            ax.set_title(f"Histogram: {field_name}")
            ax.set_xlabel("Value")
            ax.set_ylabel("Frequency")
            self.chart_info_label.setText(f"Displaying Histogram for numeric field: {field_name}")
            
        elif '_top_values_raw' in field_data:
            top_vals = field_data['_top_values_raw']
            if top_vals:
                labels = [str(v[0])[:15] for v in top_vals]
                counts = [v[1] for v in top_vals]
                x_pos = range(len(labels))
                ax.bar(x_pos, counts, align='center', alpha=0.7)
                ax.set_xticks(x_pos)
                ax.set_xticklabels(labels, rotation=45, ha='right')
                ax.set_title(f"Top Values: {field_name}")
                ax.set_ylabel("Count")
                self.figure.tight_layout()
                self.chart_info_label.setText(f"Displaying Bar Chart for text field: {field_name}")
            else:
                 ax.text(0.5, 0.5, "No data for chart", ha='center', va='center')
        else:
             ax.text(0.5, 0.5, "No chart available for this field type", ha='center', va='center')
             self.chart_info_label.setText(f"No specific chart for field: {field_name}")

        self.canvas.draw()


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
        super().__init__(parent)
        self.iface = iface
        self.setObjectName("FieldProfilerDockWidgetInstance")
        self.setWindowTitle(self.tr("Field Profiler"))

        self.main_widget = QWidget()
        self.main_layout = QVBoxLayout(self.main_widget)
        self.setWidget(self.main_widget)

        self._create_input_group()

        self.layerComboBox.layerChanged.connect(self.populate_fields)
        self.analyzeButton.clicked.connect(self.run_analysis)

        self.populate_fields(self.layerComboBox.currentLayer())
        if not SCIPY_AVAILABLE:
            self.iface.messageBar().pushMessage(
                "Field Profiler Warning",
                self.tr("Scipy library not found. Advanced numeric statistics (Skewness, Kurtosis, Normality) will be unavailable."),
                level=Qgis.Warning, duration=10
            )

        self.current_task = None
        self.analysis_results_dialog = None

        
    def tr(self, message):
        return QtCore.QCoreApplication.translate("FieldProfilerDockWidget", message)

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
        
        # Validation Rules UI
        self.validation_group = QGroupBox(self.tr("Validation Rules (Optional)"))
        self.validation_group.setCheckable(True)
        self.validation_group.setChecked(False)
        val_layout = QVBoxLayout()
        val_lbl = QLabel(self.tr("Enter QGIS Expressions (one per line). Example: \"age\" > 20"))
        val_layout.addWidget(val_lbl)
        self.validation_rules_edit = QPlainTextEdit()
        self.validation_rules_edit.setPlaceholderText("\"field_a\" > 0\nlength(\"name\") < 50")
        self.validation_rules_edit.setMaximumHeight(80) 
        val_layout.addWidget(self.validation_rules_edit)
        self.validation_group.setLayout(val_layout)
        main_input_layout.addWidget(self.validation_group)

        self.input_group_box.setLayout(main_input_layout)
        self.main_layout.addWidget(self.input_group_box)
        self.main_layout.addStretch()


    def populate_fields(self, layer):
        self.fieldListWidget.clear()
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
        # Check if task is already running
        # Check if task is already running
        if self.current_task:
            try:
                if self.current_task.status() == QgsTask.Running:
                    self.current_task.cancel()
                    self.analyzeButton.setText(self.tr("Analyze Selected Fields"))
                    return
            except RuntimeError:
                # Task object deleted (C++ side), but python ref exists. Treat as finished.
                self.current_task = None

        current_layer = self.layerComboBox.currentLayer()
        selected_list_items = self.fieldListWidget.selectedItems()
        
        self.current_limit_unique_display = self.limitUniqueSpinBox.value()
        self.current_decimal_places = self.decimalPlacesSpinBox.value()
        self._was_analyzing_selected_features = self.selectedOnlyCheckbox.isChecked()
        detailed_options = self._get_detailed_options_state()
        # Pass general config options too
        detailed_options['limit_unique'] = self.current_limit_unique_display
        detailed_options['decimal_places'] = self.current_decimal_places
        detailed_options['scipy_available'] = SCIPY_AVAILABLE

        if not current_layer or not isinstance(current_layer, QgsVectorLayer):
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Please select a valid vector layer."), level=Qgis.Warning); return
        if not selected_list_items:
            self.iface.messageBar().pushMessage(self.tr("Error"), self.tr("Please select one or more fields to analyze."), level=Qgis.Warning); return

        selected_field_names = [item.text().split(" (")[0] for item in selected_list_items]
        if not selected_field_names: return

        # Determine features to analyze
        selected_ids = []
        if self._was_analyzing_selected_features:
            selected_ids = current_layer.selectedFeatureIds()
            if not selected_ids:
                self.iface.messageBar().pushMessage(self.tr("Warning"), self.tr("No features selected for analysis."), level=Qgis.Warning); return
        
        # Prepare Validation Rules
        validation_rules = []
        if self.validation_group.isChecked():
            raw_rules = self.validation_rules_edit.toPlainText().split('\n')
            validation_rules = [r.strip() for r in raw_rules if r.strip()]

        # Setup Task
        self.current_task = FieldProfilerTask(
            current_layer, 
            selected_field_names, 
            detailed_options, 
            selected_ids=selected_ids if selected_ids else None,
            validation_rules=validation_rules
        )
        self.current_task.analysisFinished.connect(self.on_analysis_finished)
        self.current_task.progressChanged.connect(lambda p: self.progressBar.setValue(int(p)))
        
        # UI Updates
        self.analyzeButton.setText(self.tr("Cancel Analysis"))
        self.progressBar.setVisible(True)
        self.progressBar.setRange(0, 100)
        self.progressBar.setValue(0)
        
        QgsApplication.taskManager().addTask(self.current_task)

    def on_analysis_finished(self, results):
        self.analyzeButton.setText(self.tr("Analyze Selected Fields"))
        self.progressBar.setVisible(False)
        
        if not results:
            self.iface.messageBar().pushMessage(self.tr("Info"), self.tr("Analysis cancelled or failed."), level=Qgis.Warning)
            self.current_task = None
            return

        current_layer = self.layerComboBox.currentLayer()

        detailed_options = {
            'decimal_places': self.decimalPlacesSpinBox.value()
        }
        
        # Instantiate and show floating dialog
        self.analysis_results_dialog = AnalysisResultsDialog(
            parent=self, 
            results_data=results, 
            layer=current_layer, 
            detailed_options=detailed_options,
            was_analyzing_selection=self._was_analyzing_selected_features
        )
        self.analysis_results_dialog.show()

        self.iface.messageBar().pushMessage(self.tr("Success"), self.tr("Analysis complete. Results opened in new window."), level=Qgis.Success)
        self.current_task = None

    def closeEvent(self, event):
        self.hide()
        event.ignore()
