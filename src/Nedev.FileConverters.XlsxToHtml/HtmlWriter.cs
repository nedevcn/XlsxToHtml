using System;
using System.IO;
using System.Text;

namespace Nedev.FileConverters.XlsxToHtml
{
    public class HtmlWriter : IHtmlWriter
    {
        /// <summary>
        /// When true, formulas will be evaluated using a simple built-in engine; the result
        /// replaces the cached value in the output HTML. Default is false.
        /// </summary>
        public bool EvaluateFormulas { get; set; }

        /// <summary>
        /// When true, column widths from the Excel file will be applied to the HTML table.
        /// Default is true.
        /// </summary>
        public bool ApplyColumnWidths { get; set; } = true;

        /// <summary>
        /// When true, row heights from the Excel file will be applied to the HTML table.
        /// Default is false (to allow natural row height).
        /// </summary>
        public bool ApplyRowHeights { get; set; } = false;

        public void Write(Workbook workbook, TextWriter output)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (output == null) throw new ArgumentNullException(nameof(output));

            var sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html><head><meta charset=\"utf-8\"/>");
            sb.AppendLine("<style>");
            sb.AppendLine("table { border-collapse: collapse; }");
            sb.AppendLine("td { box-sizing: border-box; }");
            sb.AppendLine("</style>");
            sb.AppendLine("</head><body>");

            for (int i = 0; i < workbook.Sheets.Count; i++)
            {
                var sheet = workbook.Sheets[i];
                sb.AppendLine($"<h1>{Escape(sheet.Name)}</h1>");
                
                // Generate colgroup for column widths
                if (ApplyColumnWidths && sheet.Columns.Count > 0)
                {
                    sb.AppendLine("<table cellspacing=\"0\" cellpadding=\"2\">");
                    sb.AppendLine("<colgroup>");
                    var maxCol = sheet.Columns.Keys.Max();
                    for (int c = 1; c <= maxCol; c++)
                    {
                        if (sheet.Columns.TryGetValue(c, out var colInfo) && colInfo.Width.HasValue)
                        {
                            // Convert Excel width (characters) to approximate pixels
                            // Standard: ~7 pixels per character + 10 pixels padding
                            var widthPx = (int)(colInfo.Width.Value * 7 + 10);
                            sb.AppendLine($"<col style=\"width:{widthPx}px;\"/>");
                        }
                        else
                        {
                            sb.AppendLine("<col/>");
                        }
                    }
                    sb.AppendLine("</colgroup>");
                }
                else
                {
                    sb.AppendLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"2\">");
                }
                
                foreach (var row in sheet.Rows)
                {
                    // Skip hidden rows
                    if (row.Hidden) continue;
                    
                    var rowStyle = new StringBuilder();
                    if (ApplyRowHeights && row.Height.HasValue)
                    {
                        rowStyle.Append($"height:{row.Height.Value}pt;");
                    }
                    
                    var rowAttr = rowStyle.Length > 0 ? $" style=\"{rowStyle}\"" : string.Empty;
                    sb.AppendLine($"<tr{rowAttr}>");
                    
                    foreach (var cell in row.Cells)
                    {
                        // Skip hidden columns
                        if (sheet.Columns.TryGetValue(cell.Column, out var col) && col.Hidden)
                            continue;
                        
                        // skip cells that are covered by a merge and are not top-left
                        var merge = sheet.Merges.FirstOrDefault(m => m.Covers(row.Number, cell.Column));
                        if (merge != null && !merge.IsTopLeft(row.Number, cell.Column))
                            continue;

                        var attrs = new List<string>();
                        if (merge != null)
                        {
                            int rowspan = merge.EndRow - merge.StartRow + 1;
                            int colspan = merge.EndCol - merge.StartCol + 1;
                            if (rowspan > 1) attrs.Add($"rowspan=\"{rowspan}\"");
                            if (colspan > 1) attrs.Add($"colspan=\"{colspan}\"");
                        }
                        var style = cell.Style?.ToCss();
                        if (!string.IsNullOrEmpty(style))
                            attrs.Add($"style=\"{style}\"");
                        if (!string.IsNullOrEmpty(cell.Formula))
                            attrs.Add($"title=\"={Escape(cell.Formula)}\"");

                        var attrText = attrs.Count > 0 ? " " + string.Join(" ", attrs) : string.Empty;
                        // decide display value: cached or evaluated
                        var displayValue = cell.Value;
                        if (EvaluateFormulas && !string.IsNullOrEmpty(cell.Formula))
                        {
                            var eval = FormulaEvaluator.Evaluate(cell.Formula, sheet);
                            displayValue = eval;
                        }
                        sb.AppendLine($"<td{attrText}>{Escape(displayValue)}</td>");
                    }
                    sb.AppendLine("</tr>");
                }
                sb.AppendLine("</table>");
            }

            sb.AppendLine("</body></html>");
            output.Write(sb.ToString());
        }

        public string Convert(Workbook workbook)
        {
            using var sw = new StringWriter();
            Write(workbook, sw);
            return sw.ToString();
        }

        private static string Escape(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;
            return System.Net.WebUtility.HtmlEncode(value);
        }
    }
}