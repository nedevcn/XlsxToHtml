using System;
using System.IO;
using System.Text;

namespace Nedev.XlsxToHtml
{
    public class HtmlWriter : IHtmlWriter
    {
        public void Write(Workbook workbook, TextWriter output)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));
            if (output == null) throw new ArgumentNullException(nameof(output));

            var sb = new StringBuilder();
            sb.AppendLine("<html><head><meta charset=\"utf-8\"/></head><body>");

            for (int i = 0; i < workbook.Sheets.Count; i++)
            {
                var sheet = workbook.Sheets[i];
                sb.AppendLine($"<h1>{Escape(sheet.Name)}</h1>");
                sb.AppendLine("<table border=\"1\" cellspacing=\"0\" cellpadding=\"2\">");
                foreach (var row in sheet.Rows)
                {
                    sb.AppendLine("<tr>");
                    foreach (var cell in row.Cells)
                    {
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

                        var attrText = attrs.Count > 0 ? " " + string.Join(" ", attrs) : string.Empty;
                        sb.AppendLine($"<td{attrText}>{Escape(cell.Value)}</td>");
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