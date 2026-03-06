using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace Nedev.FileConverters.XlsxToHtml
{
    public class XlsxReader : IXlsxReader
    {
        public Workbook Read(string path)
        {
            using var archive = ZipFile.OpenRead(path);
            return ReadArchive(archive);
        }

        public Workbook Read(Stream stream)
        {
            // allow stream to be used (e.g. coming from package infrastructure)
            using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
            return ReadArchive(archive);
        }

        private static Workbook ReadArchive(ZipArchive archive)
        {
            var workbook = new Workbook();

            var sharedStrings = ReadSharedStrings(archive);
            var styles = ReadStyles(archive);
            var sheetNames = ReadSheetNames(archive);

            for (int i = 0; i < sheetNames.Count; i++)
            {
                var sheet = ReadWorksheet(archive, i + 1, sheetNames[i], sharedStrings, styles);
                workbook.Sheets.Add(sheet);
            }

            return workbook;
        }

        private static List<string> ReadSharedStrings(ZipArchive archive)
        {
            var result = new List<string>();
            var entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
                return result;

            using var stream = entry.Open();
            // load with LINQ to XML for simplicity and correctness
            var doc = System.Xml.Linq.XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            foreach (var si in doc.Root.Elements(ns + "si"))
            {
                // concatenate all <t> descendants (covers rich text runs)
                var text = string.Concat(si.Descendants(ns + "t").Select(t => (string)t));
                result.Add(text);
            }
            return result;
        }

        private static Dictionary<int, CellStyle> ReadStyles(ZipArchive archive)
        {
            var styles = new Dictionary<int, CellStyle>();
            var entry = archive.GetEntry("xl/styles.xml");
            if (entry == null)
                return styles;

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream);

            // temporary lists for fonts and fills
            var fonts = new List<(bool Bold, bool Italic, bool Underline, string? Color)>();
            var fills = new List<string?>();
            var numFmts = new Dictionary<int, string?>();

            // We need two passes: collect numFmts/fonts/fills, then cellXfs
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name == "numFmt")
                    {
                        var idAttr = reader.GetAttribute("numFmtId");
                        var code = reader.GetAttribute("formatCode");
                        if (int.TryParse(idAttr, out int id))
                            numFmts[id] = code;
                    }
                    else if (reader.Name == "font")
                    {
                        // parse font element
                        bool bold = false, italic = false, underline = false;
                        string? color = null;
                        var depth = reader.Depth;
                        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "font" && reader.Depth == depth))
                        {
                            if (reader.NodeType == XmlNodeType.Element)
                            {
                                if (reader.Name == "b") bold = true;
                                else if (reader.Name == "i") italic = true;
                                else if (reader.Name == "u") underline = true;
                                else if (reader.Name == "color")
                                {
                                    var rgb = reader.GetAttribute("rgb");
                                    if (!string.IsNullOrEmpty(rgb) && rgb.Length == 8 && rgb.StartsWith("FF"))
                                        color = "#" + rgb.Substring(2);
                                }
                            }
                        }

                        fonts.Add((bold, italic, underline, color));
                    }
                    else if (reader.Name == "fill")
                    {
                        // parse fill element
                        string? bg = null;
                        var depth = reader.Depth;
                        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "fill" && reader.Depth == depth))
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Name == "fgColor")
                            {
                                var rgb = reader.GetAttribute("rgb");
                                if (!string.IsNullOrEmpty(rgb) && rgb.Length == 8 && rgb.StartsWith("FF"))
                                    bg = "#" + rgb.Substring(2);
                            }
                        }
                        fills.Add(bg);
                    }
                    else if (reader.Name == "cellXfs")
                    {
                        // parse xf entries sequentially
                        int xfIndex = -1;
                        var depth = reader.Depth;
                        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "cellXfs" && reader.Depth == depth))
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Name == "xf")
                            {
                                xfIndex++;
                                var style = new CellStyle();
                                if (int.TryParse(reader.GetAttribute("numFmtId"), out int nfmt))
                                {
                                    style.NumberFormatId = nfmt;
                                    if (numFmts.TryGetValue(nfmt, out var fmt))
                                        style.NumberFormat = fmt;
                                }
                                if (int.TryParse(reader.GetAttribute("fontId"), out int fontId) && fontId >= 0 && fontId < fonts.Count)
                                {
                                    var f = fonts[fontId];
                                    style.Bold = f.Bold;
                                    style.Italic = f.Italic;
                                    style.Underline = f.Underline;
                                    style.FontColor = f.Color;
                                }
                                if (int.TryParse(reader.GetAttribute("fillId"), out int fillId) && fillId >= 0 && fillId < fills.Count)
                                {
                                    style.BackgroundColor = fills[fillId];
                                }
                                styles[xfIndex] = style;
                            }
                        }
                    }
                }
            }

            return styles;
        }

        private static List<string> ReadSheetNames(ZipArchive archive)
        {
            var names = new List<string>();
            var entry = archive.GetEntry("xl/workbook.xml");
            if (entry == null)
                return names;

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "sheet")
                {
                    var name = reader.GetAttribute("name") ?? string.Empty;
                    names.Add(name);
                }
            }
            return names;
        }

        private static Worksheet ReadWorksheet(ZipArchive archive, int sheetNumber, string sheetName,
            List<string> sharedStrings, Dictionary<int, CellStyle> styles)
        {
            var path = $"xl/worksheets/sheet{sheetNumber}.xml";
            var entry = archive.GetEntry(path);
            var ws = new Worksheet { Name = sheetName, Index = sheetNumber - 1 };
            // marker so we can tell whether this version of ReadWorksheet ran
            if (entry == null)
                return ws;

            using var stream = entry.Open();
            using var reader = XmlReader.Create(stream);

            Row? currentRow = null;
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    if (reader.Name == "mergeCells")
                    {
                        // process mergeCell entries inside this container
                        var depth = reader.Depth;
                        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "mergeCells" && reader.Depth == depth))
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Name == "mergeCell")
                            {
                                var refAttr = reader.GetAttribute("ref");
                                if (!string.IsNullOrEmpty(refAttr))
                                {
                                    var parts = refAttr.Split(':');
                                    if (parts.Length == 2)
                                    {
                                        var start = DecodeCellRef(parts[0]);
                                        var end = DecodeCellRef(parts[1]);
                                        ws.Merges.Add(new MergeCell
                                        {
                                            StartRow = start.row,
                                            StartCol = start.col,
                                            EndRow = end.row,
                                            EndCol = end.col
                                        });
                                    }
                                }
                            }
                        }
                    }
                    else if (reader.Name == "row")
                    {
                        currentRow = new Row();
                        if (int.TryParse(reader.GetAttribute("r"), out int r))
                            currentRow.Number = r;
                        ws.Rows.Add(currentRow);
                    }
                    else if (reader.Name == "c" && currentRow != null)
                    {
                        var cell = new Cell();
                        var rref = reader.GetAttribute("r");
                        if (!string.IsNullOrEmpty(rref))
                        {
                            (cell.Row, cell.Column) = DecodeCellRef(rref);
                        }

                        var t = reader.GetAttribute("t");
                        switch (t)
                        {
                            case "s":
                                cell.Type = CellType.SharedString;
                                break;
                            case "b":
                                cell.Type = CellType.Boolean;
                                break;
                            case "str":
                            case "inlineStr":
                                cell.Type = CellType.InlineString;
                                break;
                            case "e":
                                cell.Type = CellType.Error;
                                break;
                            default:
                                cell.Type = CellType.Number;
                                break;
                        }

                        if (int.TryParse(reader.GetAttribute("s"), out int sidx))
                        {
                            if (styles.TryGetValue(sidx, out var style))
                                cell.Style = style;
                        }

                        // read value element inside <c>
                        if (!reader.IsEmptyElement)
                        {
                            int startDepth = reader.Depth;
                            while (reader.Read())
                            {
    
                                if (reader.NodeType == XmlNodeType.Element && reader.Name == "f")
                                {
                                    // formula element - content may span multiple nodes but we only care about the string
                                    cell.Formula = reader.ReadElementContentAsString();
                                }
                                else if (reader.NodeType == XmlNodeType.Element && reader.Name == "v")
                                {
                                    // read raw value manually instead of ReadElementContentAsString, so we don't consume </c>
                                    string raw = string.Empty;
                                    if (!reader.IsEmptyElement)
                                    {
                                        // move to the text inside <v>
                                        if (reader.Read() && (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.Whitespace || reader.NodeType == XmlNodeType.SignificantWhitespace))
                                        {
                                            raw = reader.Value;
                                        }
                                        // after reading text we should be positioned on the text node; advance until we hit the end of <v>
                                        while (reader.NodeType != XmlNodeType.EndElement || reader.Name != "v")
                                        {
                                            if (!reader.Read()) break;
                                        }
                                    }
                                    cell.Value = InterpretValue(raw, cell.Type, sharedStrings, cell.Style);
                                }
                                else if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == startDepth)
                                {
                                    // reached closing </c>; break out so outer loop can continue normally
                                    break;
                                }
                            }
                        }

                        currentRow.Cells.Add(cell);
                    }
                }
            }

            return ws;
        }

        private static (int row, int col) DecodeCellRef(string r)
        {
            int col = 0;
            int row = 0;
            foreach (char c in r)
            {
                if (char.IsLetter(c))
                {
                    col = col * 26 + (c - 'A' + 1);
                }
                else if (char.IsDigit(c))
                {
                    row = row * 10 + (c - '0');
                }
            }
            return (row, col);
        }

        private static string? InterpretValue(string raw, CellType type, List<string> shared, CellStyle? style)
        {
            switch (type)
            {
                case CellType.SharedString:
                    if (int.TryParse(raw, out int si) && si >= 0 && si < shared.Count)
                        return shared[si];
                    return raw;
                case CellType.Boolean:
                    return raw == "1" ? "TRUE" : "FALSE";
                case CellType.Number:
                    if (double.TryParse(raw, out double d))
                    {
                        if (style != null)
                        {
                            // custom format string takes priority
                            if (!string.IsNullOrEmpty(style.NumberFormat))
                            {
                                try
                                {
                                    // choose section based on value
                                    var fmt = PickSection(style.NumberFormat, d);
                                    // remove color codes like [Red] and maybe apply color
                                    var color = ExtractColor(ref fmt);
                                    if (style != null && !string.IsNullOrEmpty(color))
                                        style.FontColor = color;
                                    if (FormatIsDate(fmt))
                                    {
                                        var dt = DateTime.FromOADate(d);
                                        return dt.ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    // percent handling
                                    if (fmt.Contains("%"))
                                        return FormatPercent(d, fmt);
                                    // fraction handling
                                    if (fmt.Contains("?/"))
                                        return FormatFraction(d, fmt);
                                    // default numeric
                                    return d.ToString(fmt);
                                }
                                catch
                                {
                                    // fallback to raw if format fails
                                }
                            }
                            else if (style.NumberFormatId.HasValue && IsDateFormat(style.NumberFormatId.Value))
                            {
                                var dt = DateTime.FromOADate(d);
                                return dt.ToString("yyyy-MM-dd HH:mm:ss");
                            }
                        }
                    }
                    return raw;
                default:
                    return raw;
            }
        }

        private static string PickSection(string fmt, double value)
        {
            var parts = fmt.Split(';');
            if (parts.Length == 1) return fmt;
            if (parts.Length == 2)
                return value >= 0 ? parts[0] : parts[1];
            if (parts.Length == 3)
                return value > 0 ? parts[0] : value < 0 ? parts[1] : parts[2];
            return parts[0];
        }

        private static string StripColor(string fmt)
        {
            // remove all bracketed expressions
            return System.Text.RegularExpressions.Regex.Replace(fmt, @"\[[^\]]+\]", string.Empty);
        }

        private static string ExtractColor(ref string fmt)
        {
            // finds a color code in brackets and returns hex or named color
            var m = System.Text.RegularExpressions.Regex.Match(fmt, @"\[([^\]]+)\]");
            if (m.Success)
            {
                var text = m.Groups[1].Value.Trim();
                string? hex = null;
                // if it looks like a hex code (6 or 8 hex digits)
                if (System.Text.RegularExpressions.Regex.IsMatch(text, "^[0-9A-Fa-f]{6,8}$"))
                {
                    var h = text;
                    if (h.Length == 6) hex = "#" + h;
                    else if (h.Length == 8) hex = "#" + h.Substring(2); // drop alpha
                }
                else
                {
                    if (!ColorHelper.TryGetColor(text, out hex))
                        hex = null;
                }
                fmt = StripColor(fmt); // remove all bracket parts
                return hex ?? string.Empty;
            }
            return string.Empty;
        }
        private static string FormatPercent(double d, string fmt)
        {
            // Excel stores percent as value 0.5 -> 50%
            double p = d * 100;
            // strip % from format
            var f = fmt.Replace("%", string.Empty);
            try
            {
                return p.ToString(f) + "%";
            }
            catch
            {
                return p.ToString("0.##") + "%";
            }
        }

        private static string FormatFraction(double d, string fmt)
        {
            // very simple: convert to nearest denominator based on pattern like "# ?/?" or "# ??/??"
            var match = System.Text.RegularExpressions.Regex.Match(fmt, @"(\?+)/(\?+)");
            if (match.Success)
            {
                int denomDigits = match.Groups[2].Value.Length;
                int denom = (int)Math.Pow(10, denomDigits);
                int whole = (int)Math.Truncate(d);
                double frac = Math.Abs(d - whole);
                int num = (int)Math.Round(frac * denom);
                if (num == 0) return whole.ToString();
                return string.Format("{0} {1}/{2}", whole, num, denom);
            }
            return d.ToString();
        }
        private static bool IsDateFormat(int numFmtId)
        {
            // built-in date formats 14-22
            return numFmtId >= 14 && numFmtId <= 22;
        }

        private static bool FormatIsDate(string fmt)
        {
            if (string.IsNullOrEmpty(fmt))
                return false;
            var s = fmt.ToLowerInvariant();
            // very simplistic: if contains date/time tokens
            return s.Contains("yy") || s.Contains("mm") || s.Contains("dd") || s.Contains("h") || s.Contains("s");
        }
    }
}