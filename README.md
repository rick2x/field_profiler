<p align="center">
  <img src="./icon2.png" alt="Field Profiler Logo" width="128"/>
</p>
QGIS Field Profiler Plugin

**Version:** 0.1.2
**QGIS Minimum Version:** 3.0
**Author:** ricks (rrcuario@gmail.com)

## Description

The Field Profiler is a QGIS plugin designed to provide an advanced and comprehensive analysis of attribute data for vector layers. It helps users understand data distributions, identify quality issues, and gain deep insights into their attribute fields.

This plugin goes beyond basic statistics, offering detailed metrics for numeric, text, and date/time fields, empowering users in data cleaning, validation, and exploration workflows.

![image](https://github.com/user-attachments/assets/fd0d0d3b-6de5-49a5-a584-8411295c80e2)

## Features

*   **Comprehensive Statistics:**
    *   **Numeric Fields:** Min, Max, Mean, Median, StdDev, Sum, Range, Quartiles (Q1, Q3), IQR, Null Count, Conversion Errors, Zeros, Positives, Negatives, Coefficient of Variation, Low Variance Flag.
        *   *Detailed Numeric:* Skewness, Kurtosis, Shapiro-Wilk Normality Test (p-value & flag), 1st/5th/95th/99th Percentiles, Integer/Decimal counts, Optimal Bins (Freedman-Diaconis), detailed Outlier counts (Min/Max Outlier, % Outliers). (Requires SciPy)
    *   **Text Fields:** Min/Max/Avg Length, Variety (distinct count), Empty String count, Leading/Trailing Space count.
        *   *Detailed Text:* Case analysis (% Upper/Lower/Title/Mixed), count of strings with internal multiple spaces, count of values occurring only once, count of strings with non-printable characters.
    *   **Date/Time Fields:** Min/Max Date, Common Years/Months/Days of Week, Dates Before/After Today, Unique Values.
        *   *Detailed Date/Time:* Common Hours, % Midnight/Noon times (for DateTime fields), % Weekend/Weekday dates.
*   **Data Type Mismatch Hints:** Provides suggestions if a field's content statistically resembles a different data type.
*   **Flexible Analysis Scope:**
    *   Analyze all features in a layer.
    *   Analyze only selected features.
*   **Interactive Results Table:**
    *   Clear presentation of statistics per field.
    *   Sortable columns.
    *   Visual cues for potential data quality issues.
*   **Feature Selection from Results:**
    *   Double-click on specific statistic cells (e.g., Null Count, Empty Strings, Conversion Errors, Outliers, Top Unique Value, Non-Printable Chars) to select the corresponding features on the map.
    *   Selection can intersect with existing selections if analysis was performed on selected features.
*   **Export & Copy:**
    *   Copy results table to clipboard.
    *   Export results table to CSV.
*   **Configurable Options:**
    *   Set the limit for "Top N" unique values displayed.
    *   Control the number of decimal places for numeric statistics.
    *   Enable/disable groups of detailed (potentially performance-intensive) statistics.
*   **User-Friendly Interface:** Dockable widget integrated into the QGIS environment.

## Installation

1.  **Download:**
    *   Download the latest release ZIP
2.  **QGIS Plugin Manager:**
    *   Open QGIS.
    *   Go to `Plugins` -> `Manage and Install Plugins...`.
    *   Select `Install from ZIP`.
    *   Browse to the downloaded ZIP file and click `Install Plugin`.
3.  **Enable the Plugin:**
    *   Find "Field Profiler" in the list of installed plugins and ensure its checkbox is ticked.

Alternatively, for development or manual installation:
1.  Clone or download this repository.
2.  Copy the `field_profiler` directory into your QGIS plugins directory:
    *   **Windows:** `C:\Users\<YourUser>\AppData\Roaming\QGIS\QGIS3\profiles\default\python\plugins\`
    *   **Linux:** `~/.local/share/QGIS/QGIS3[or QGIS]\profiles\default\python\plugins\`
    *   **macOS:** `~/Library/Application Support/QGIS/QGIS3/profiles/default/python/plugins\`
3.  Restart QGIS or use the Plugin Reloader plugin.
4.  Enable "Field Profiler" in the Plugin Manager.

## Dependencies

*   **QGIS:** Version 3.0 or higher.
*   **Python Libraries:**
    *   `numpy`: Usually bundled with QGIS.
    *   `scipy`: (Optional, but recommended for advanced numeric statistics like Skewness, Kurtosis, and Normality tests). If SciPy is not found, these specific statistics will be unavailable, and a warning will be shown.

## Usage

1.  Once installed and enabled, a toolbar icon for "Field Profiler" will appear, or you can find it under the `Vector` menu (or wherever you assigned it, e.g., `Plugins -> Field Profiler`). Click to open the dock widget.
2.  **Select Layer:** Choose a vector layer from the dropdown.
3.  **Select Field(s):** Select one or more fields from the list to analyze.
4.  **Settings:**
    *   Optionally, check "Analyze selected features only" to limit analysis to the currently selected features in the chosen layer.
    *   Adjust "Unique Values Limit" and "Numeric Decimal Places" as needed.
    *   Choose which "Detailed Analysis Options" to include. Disabling some may improve performance on very large datasets.
5.  **Analyze:** Click the "Analyze Selected Fields" button.
6.  **View Results:** The results table will populate with statistics.
    *   Double-click on certain cells (e.g., 'Null Count', 'Outliers (IQR)', 'Unique Values (Top)') to select the corresponding features on the map.
7.  **Export/Copy:** Use the buttons to copy the table to the clipboard or export it as a CSV file.

## Changelog

### Version 0.1.2 (2026-02-01)

This version contains critical bug fixes that resolve several issues preventing proper plugin operation.

**Critical Bug Fixes:**

1.  **Fixed Runtime Errors in Task Processing (`field_profiler_task.py`):**
    *   Added missing `non_printable_fids` key to collector dictionary initialization (was causing `KeyError`).
    *   Fixed `field_res` variable being referenced before definition (was causing `NameError`).
    *   Implemented proper non-printable character detection - now correctly calls `_has_non_printable_chars()`.

2.  **Completed Date Field Analysis (`_analyze_date`):**
    *   Replaced placeholder implementation with full date statistics.
    *   Now properly calculates Min/Max Date, Common Years/Months/Days, Weekend/Weekday percentages, and Before/After Today counts.

3.  **Fixed Control Flow in Results Handler (`field_profiler_dockwidget.py`):**
    *   Corrected mismatched `else` block in `on_analysis_finished()` that caused logic errors.
    *   Removed ~135 lines of orphaned dead code that was left from previous refactoring.

4.  **Added Missing Imports:**
    *   Added `QPlainTextEdit` and `QHeaderView` to widget imports.

**Code Quality Improvements:**

*   Removed duplicate imports in `field_profiler_plugin.py`.
*   Removed debug `print()` statements from production code.
*   Synchronized version number between `metadata.txt` and `README.md`.

---

### Version 0.1.1 (2025-05-14)

This version focuses on enhancing the robustness of internal logic, especially for feature selection from the results table, and includes several minor improvements across analysis functions.

**Key Changes (Internal Refactoring for Robustness):**

1.  **Improved Statistic Key Handling in Results Table (`populate_results_table`):**
    *   Statistic rows now internally use their original (English, non-translated) keys.
    *   The displayed statistic names in the table are translated for the user.
    *   The original key is stored with each statistic row item (`QTableWidgetItem.setData(Qt.UserRole, original_key)`).
    *   Internal logic (tooltips, coloring, formatting) consistently uses these original keys, making it independent of UI translations.

2.  **Enhanced Double-Click Feature Selection (`_on_cell_double_clicked`):**
    *   Retrieves the original statistic key directly from the table item's stored data.
    *   Removed the previous, more fragile method of matching translated text to find the original key.
    *   All subsequent logic for building feature selection expressions now reliably uses the original statistic key.

**Minor Improvements & Fixes:**

1.  **Numeric Field Analysis (`analyze_numeric_field_from_list`):**
    *   Improved handling of empty or NaN-only datasets for `statistics.multimode`.
    *   Ensured `% Outliers` correctly defaults to `0.0`.
    *   More robust calculation for "Optimal Bins (Freedman-Diaconis)", including fallbacks.
    *   Generalized `ValueError` handling for Shapiro-Wilk test.
    *   Added checks for all-NaN data before `numpy.nanpercentile` calculation.

2.  **Text Field Analysis (`analyze_text_field`):**
    *   Ensured `None` values are consistently treated as empty strings during analysis.
    *   Clarified calculation for `% Mixed Case`.
    *   Preserved hyphens within words for "Top Words" analysis.
    *   Slightly improved robustness of the URL detection regex.

3.  **Date Field Analysis (`analyze_date_field_enhanced`):**
    *   Standardized internal handling of `QDate` and `QDateTime` objects.
    *   Improved "Min/Max Date" formatting to reflect whether the field is primarily Date or DateTime.
    *   Translated day and month names for "Common Days/Months" statistics.
    *   Corrected "Dates Before/After Today" to compare only the date parts.
    *   Ensured "Unique Values (Top)" caches the actual `QDate` or `QDateTime` object for selection.
    *   Updated selection logic for `QDateTime` values to use the QGIS `datetime('...')` expression function with ISO-formatted strings.

4.  **Double-Click Feature Selection (`_on_cell_double_clicked` - Additional):**
    *   Now correctly handles selection for the top unique value when it is `NULL` (by creating an `IS NULL` expression).
    *   Improved user feedback messages for certain selection scenarios.
    *   Removed experimental attribute table update attempts; relies on QGIS's standard behavior.

5.  **CSV Export (`export_results_to_csv`):**
    *   Added `encoding='utf-8-sig'` to improve compatibility with Microsoft Excel (includes BOM).
    *   Sanitized layer names for use in default CSV filenames.

6.  **General Code Style:**
    *   Updated to use Python 3 style `super().__init__(parent)`.

## Future Enhancements

*   User-defined regex patterns for text analysis.
*   More granular character-level analysis for text.
*   Visualizations (e.g., small histograms/bar charts in the results).
*   Saving/loading analysis configurations.

## Contributing

Contributions, bug reports, and feature requests are welcome! Please open an issue or submit a pull request.

## License

This plugin is licensed under the GNU General Public License as published by the Free Software Foundation.
