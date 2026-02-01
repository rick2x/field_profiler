# -*- coding: utf-8 -*-
import datetime

class ReportGenerator:
    """
    Generates an HTML report from Field Profiler analysis results.
    """
    def __init__(self, layer_name):
        self.layer_name = layer_name
        self.timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def generate_report(self, results, correlation_matrix=None):
        """
        results: OrderedDict of field results
        correlation_matrix: Dict with 'fields' and 'matrix' (optional)
        """
        html = [
            "<!DOCTYPE html>",
            "<html>",
            "<head>",
            f"<title>Field Profile: {self.layer_name}</title>",
            "<style>",
            "body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; color: #333; }",
            "h1 { color: #2c3e50; }",
            "h2 { color: #34495e; margin-top: 30px; }",
            "table { border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 14px; }",
            "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }",
            "th { background-color: #f2f2f2; color: #333; font-weight: bold; }",
            "tr:nth-child(even) { background-color: #f9f9f9; }",
            "tr:hover { background-color: #f1f1f1; }",
            ".stat-name { font-weight: bold; color: #555; }",
            ".numeric { text-align: right; }",
            ".quality-issue { color: #d35400; font-weight: bold; }",
            ".section-info { background-color: #e8f6f3; padding: 10px; border-radius: 5px; border-left: 5px solid #1abc9c; }",
            ".heatmap-cell { text-align: center; }",
            "</style>",
            "</head>",
            "<body>",
            f"<h1>Field Profile Report</h1>",
            f"<div class='section-info'><p><strong>Layer:</strong> {self.layer_name}<br><strong>Generated:</strong> {self.timestamp}</p></div>",
            "<h2>Field Statistics</h2>",
            "<table>",
            "<thead><tr><th>Statistic</th>"
        ]

        # Table Header
        field_names = list(results.keys())
        for fname in field_names:
            html.append(f"<th>{fname}</th>")
        html.append("</tr></thead><tbody>")

        # Determine all unique stats
        all_stats = set()
        for res in results.values():
            all_stats.update(res.keys())
        
        # Sort stats (reuse roughly the logic from dockwidget if possible, or simple sort)
        # For simple report, alphabetical + some forced order is fine.
        start_stats = ['Status', 'Non-Null Count', 'Null Count', '% Null', 'Min', 'Max', 'Mean', 'Median', 'Mode(s)']
        sorted_stats = [s for s in start_stats if s in all_stats]
        sorted_stats.extend(sorted([s for s in all_stats if s not in start_stats and not s.startswith('_')])) # Skip internal keys

        for stat in sorted_stats:
            html.append(f"<tr><td class='stat-name'>{stat}</td>")
            for fname in field_names:
                val = results[fname].get(stat, "")
                # Simple formatting
                str_val = str(val)
                cls = "numeric" if isinstance(val, (int, float)) else ""
                html.append(f"<td class='{cls}'>{str_val}</td>")
            html.append("</tr>")
        
        html.append("</tbody></table>")

        # Correlation Matrix
        if correlation_matrix and 'fields' in correlation_matrix:
            c_fields = correlation_matrix['fields']
            matrix = correlation_matrix['matrix']
            html.append("<h2>Correlation Matrix</h2>")
            html.append("<table><thead><tr><th></th>")
            for f in c_fields: html.append(f"<th>{f}</th>")
            html.append("</tr></thead><tbody>")
            
            for i, row_f in enumerate(c_fields):
                html.append(f"<tr><td class='stat-name'>{row_f}</td>")
                for j, val in enumerate(matrix[i]):
                    color = "#ffffff"
                    text_color = "#000000"
                    # Simple color scale logic for HTML
                    if val > 0:
                        intensity = int(255 * (1 - abs(val)))
                        color = f"rgb({intensity}, {intensity}, 255)"
                    elif val < 0:
                        intensity = int(255 * (1 - abs(val)))
                        color = f"rgb(255, {intensity}, {intensity})"
                    
                    if abs(val) > 0.6: text_color = "#ffffff"
                    
                    html.append(f"<td class='heatmap-cell' style='background-color: {color}; color: {text_color}'>{val:.2f}</td>")
                html.append("</tr>")
            html.append("</tbody></table>")

        html.append("</body></html>")
        return "\n".join(html)
