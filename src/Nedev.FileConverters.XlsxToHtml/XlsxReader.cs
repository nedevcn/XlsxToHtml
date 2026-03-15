using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
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

        public async Task<Workbook> ReadAsync(string path, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            // Open file stream asynchronously
            await using var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
            return await ReadAsync(fileStream, cancellationToken);
        }

        public async Task<Workbook> ReadAsync(Stream stream, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            // For ZipArchive, we need to read into memory first as ZipArchive doesn't support async well
            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms, cancellationToken);
            ms.Position = 0;
            
            cancellationToken.ThrowIfCancellationRequested();
            
            using var archive = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);
            
            // Run the synchronous read on a background thread
            return await Task.Run(() => ReadArchive(archive, cancellationToken), cancellationToken);
        }

        private static Workbook ReadArchive(ZipArchive archive, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            
            var workbook = new Workbook();

            var sharedStrings = ReadSharedStrings(archive);
            cancellationToken.ThrowIfCancellationRequested();
            
            var styles = ReadStyles(archive);
            cancellationToken.ThrowIfCancellationRequested();
            
            var sheetNames = ReadSheetNames(archive);
            cancellationToken.ThrowIfCancellationRequested();
            
            var hyperlinks = ReadHyperlinks(archive);
            cancellationToken.ThrowIfCancellationRequested();

            for (int i = 0; i < sheetNames.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                
                var sheetHyperlinks = hyperlinks.TryGetValue(i + 1, out var sheetLinks) ? sheetLinks : new Dictionary<string, Hyperlink>();
                var sheet = ReadWorksheet(archive, i + 1, sheetNames[i], sharedStrings, styles, sheetHyperlinks);
                workbook.Sheets.Add(sheet);
            }

            return workbook;
        }

        private static Dictionary<int, Dictionary<string, Hyperlink>> ReadHyperlinks(ZipArchive archive)
        {
            var result = new Dictionary<int, Dictionary<string, Hyperlink>>();
            
            // Read workbook relationships to find hyperlink targets
            var relsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels");
            var hyperlinkTargets = new Dictionary<string, string>();
            
            if (relsEntry != null)
            {
                using var relsStream = relsEntry.Open();
                using var relsReader = XmlReader.Create(relsStream);
                while (relsReader.Read())
                {
                    if (relsReader.NodeType == XmlNodeType.Element && relsReader.Name == "Relationship")
                    {
                        var id = relsReader.GetAttribute("Id");
                        var type = relsReader.GetAttribute("Type");
                        var target = relsReader.GetAttribute("Target");
                        
                        if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target) && 
                            type != null && type.Contains("hyperlink"))
                        {
                            hyperlinkTargets[id] = target;
                        }
                    }
                }
            }
            
            // Read hyperlinks from each worksheet
            for (int sheetNum = 1; ; sheetNum++)
            {
                var entry = archive.GetEntry($"xl/worksheets/sheet{sheetNum}.xml");
                if (entry == null) break;
                
                var sheetHyperlinks = new Dictionary<string, Hyperlink>();
                using var stream = entry.Open();
                using var reader = XmlReader.Create(stream);
                
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.Name == "hyperlink")
                    {
                        var cellRef = reader.GetAttribute("ref");
                        var relId = reader.GetAttribute("r:id");
                        var tooltip = reader.GetAttribute("tooltip");
                        var display = reader.GetAttribute("display");
                        
                        if (!string.IsNullOrEmpty(cellRef))
                        {
                            var hyperlink = new Hyperlink
                            {
                                Tooltip = tooltip,
                                DisplayText = display
                            };
                            
                            // Try to get URL from relationships
                            if (!string.IsNullOrEmpty(relId) && hyperlinkTargets.TryGetValue(relId, out var url))
                            {
                                hyperlink.Url = url;
                            }
                            else
                            {
                                // External link stored directly
                                hyperlink.Url = reader.GetAttribute("location");
                            }
                            
                            sheetHyperlinks[cellRef] = hyperlink;
                        }
                    }
                }
                
                if (sheetHyperlinks.Count > 0)
                {
                    result[sheetNum] = sheetHyperlinks;
                }
            }
            
            return result;
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

            // temporary lists for fonts, fills, borders, and cellXfs alignment info
            var fonts = new List<FontInfo>();
            var fills = new List<string?>();
            var borders = new List<BorderInfo>();
            var numFmts = new Dictionary<int, string?>();
            var cellXfs = new List<CellXfInfo>();

            // First pass: collect all definitions
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.Name)
                    {
                        case "numFmt":
                            ReadNumFmt(reader, numFmts);
                            break;
                        case "font":
                            fonts.Add(ReadFont(reader));
                            break;
                        case "fill":
                            fills.Add(ReadFill(reader));
                            break;
                        case "border":
                            borders.Add(ReadBorder(reader));
                            break;
                        case "xf":
                            // Only process xf elements inside cellXfs
                            if (reader.Depth > 0)
                            {
                                cellXfs.Add(ReadCellXf(reader));
                            }
                            break;
                    }
                }
            }

            // Second pass: build CellStyle objects from cellXfs
            for (int i = 0; i < cellXfs.Count; i++)
            {
                var xf = cellXfs[i];
                var style = new CellStyle();

                // Number format
                if (xf.NumFmtId.HasValue)
                {
                    style.NumberFormatId = xf.NumFmtId.Value;
                    if (numFmts.TryGetValue(xf.NumFmtId.Value, out var fmt))
                        style.NumberFormat = fmt;
                }

                // Font
                if (xf.FontId.HasValue && xf.FontId.Value >= 0 && xf.FontId.Value < fonts.Count)
                {
                    var f = fonts[xf.FontId.Value];
                    style.Bold = f.Bold;
                    style.Italic = f.Italic;
                    style.Underline = f.Underline;
                    style.Strikethrough = f.Strikethrough;
                    style.FontSize = f.Size;
                    style.FontName = f.Name;
                    style.FontColor = f.Color;
                }

                // Fill
                if (xf.FillId.HasValue && xf.FillId.Value >= 0 && xf.FillId.Value < fills.Count)
                {
                    style.BackgroundColor = fills[xf.FillId.Value];
                }

                // Border
                if (xf.BorderId.HasValue && xf.BorderId.Value >= 0 && xf.BorderId.Value < borders.Count)
                {
                    var b = borders[xf.BorderId.Value];
                    style.BorderLeft = b.Left;
                    style.BorderRight = b.Right;
                    style.BorderTop = b.Top;
                    style.BorderBottom = b.Bottom;
                }

                // Alignment
                if (xf.ApplyAlignment && xf.Alignment != null)
                {
                    style.HorizontalAlign = xf.Alignment.Horizontal;
                    style.VerticalAlign = xf.Alignment.Vertical;
                    style.WrapText = xf.Alignment.WrapText;
                }

                styles[i] = style;
            }

            return styles;
        }

        private static void ReadNumFmt(XmlReader reader, Dictionary<int, string?> numFmts)
        {
            var idAttr = reader.GetAttribute("numFmtId");
            var code = reader.GetAttribute("formatCode");
            if (int.TryParse(idAttr, out int id))
                numFmts[id] = code;
        }

        private class FontInfo
        {
            public bool Bold { get; set; }
            public bool Italic { get; set; }
            public bool Underline { get; set; }
            public bool Strikethrough { get; set; }
            public double? Size { get; set; }
            public string? Name { get; set; }
            public string? Color { get; set; }
        }

        private static FontInfo ReadFont(XmlReader reader)
        {
            var font = new FontInfo();
            var depth = reader.Depth;
            
            while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "font" && reader.Depth == depth))
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    switch (reader.Name)
                    {
                        case "b":
                            font.Bold = true;
                            break;
                        case "i":
                            font.Italic = true;
                            break;
                        case "u":
                            font.Underline = true;
                            break;
                        case "strike":
                            font.Strikethrough = true;
                            break;
                        case "sz":
                            if (double.TryParse(reader.GetAttribute("val"), out double size))
                                font.Size = size;
                            break;
                        case "name":
                            font.Name = reader.GetAttribute("val");
                            break;
                        case "color":
                            font.Color = ReadColor(reader);
                            break;
                    }
                }
            }
            
            return font;
        }

        private static string? ReadFill(XmlReader reader)
        {
            string? bg = null;
            var depth = reader.Depth;
            
            while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "fill" && reader.Depth == depth))
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "fgColor")
                {
                    bg = ReadColor(reader);
                }
            }
            
            return bg;
        }

        private class BorderInfo
        {
            public BorderStyle? Left { get; set; }
            public BorderStyle? Right { get; set; }
            public BorderStyle? Top { get; set; }
            public BorderStyle? Bottom { get; set; }
        }

        private static BorderInfo ReadBorder(XmlReader reader)
        {
            var border = new BorderInfo();
            var depth = reader.Depth;
            
            while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "border" && reader.Depth == depth))
            {
                if (reader.NodeType == XmlNodeType.Element)
                {
                    BorderStyle? bs = null;
                    var style = reader.GetAttribute("style");
                    
                    // Read color if present
                    string? color = null;
                    var borderDepth = reader.Depth;
                    while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Depth == borderDepth))
                    {
                        if (reader.NodeType == XmlNodeType.Element && reader.Name == "color")
                        {
                            color = ReadColor(reader);
                        }
                    }
                    
                    if (!string.IsNullOrEmpty(style) && style != "none")
                    {
                        bs = new BorderStyle
                        {
                            Style = style,
                            Color = color,
                            Width = style.ToLowerInvariant() switch
                            {
                                "thin" => 1,
                                "medium" => 2,
                                "thick" => 3,
                                _ => 1
                            }
                        };
                    }
                    
                    switch (reader.Name)
                    {
                        case "left": border.Left = bs; break;
                        case "right": border.Right = bs; break;
                        case "top": border.Top = bs; break;
                        case "bottom": border.Bottom = bs; break;
                    }
                }
            }
            
            return border;
        }

        private class AlignmentInfo
        {
            public HorizontalAlignment Horizontal { get; set; } = HorizontalAlignment.Default;
            public VerticalAlignment Vertical { get; set; } = VerticalAlignment.Default;
            public bool WrapText { get; set; }
        }

        private class CellXfInfo
        {
            public int? NumFmtId { get; set; }
            public int? FontId { get; set; }
            public int? FillId { get; set; }
            public int? BorderId { get; set; }
            public bool ApplyAlignment { get; set; }
            public AlignmentInfo? Alignment { get; set; }
        }

        private static CellXfInfo ReadCellXf(XmlReader reader)
        {
            var xf = new CellXfInfo
            {
                NumFmtId = ParseInt(reader.GetAttribute("numFmtId")),
                FontId = ParseInt(reader.GetAttribute("fontId")),
                FillId = ParseInt(reader.GetAttribute("fillId")),
                BorderId = ParseInt(reader.GetAttribute("borderId"))
            };

            // Check if alignment is applied
            if (bool.TryParse(reader.GetAttribute("applyAlignment"), out bool applyAlign))
                xf.ApplyAlignment = applyAlign;

            // Read alignment element if present
            if (!reader.IsEmptyElement)
            {
                var depth = reader.Depth;
                while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "xf" && reader.Depth == depth))
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.Name == "alignment")
                    {
                        xf.Alignment = ReadAlignment(reader);
                    }
                }
            }

            return xf;
        }

        private static AlignmentInfo ReadAlignment(XmlReader reader)
        {
            var align = new AlignmentInfo();
            
            var horizontal = reader.GetAttribute("horizontal");
            align.Horizontal = horizontal?.ToLowerInvariant() switch
            {
                "left" => HorizontalAlignment.Left,
                "center" => HorizontalAlignment.Center,
                "right" => HorizontalAlignment.Right,
                "justify" => HorizontalAlignment.Justify,
                _ => HorizontalAlignment.Default
            };
            
            var vertical = reader.GetAttribute("vertical");
            align.Vertical = vertical?.ToLowerInvariant() switch
            {
                "top" => VerticalAlignment.Top,
                "center" => VerticalAlignment.Middle,
                "bottom" => VerticalAlignment.Bottom,
                _ => VerticalAlignment.Default
            };
            
            if (bool.TryParse(reader.GetAttribute("wrapText"), out bool wrap))
                align.WrapText = wrap;
            
            return align;
        }

        private static string? ReadColor(XmlReader reader)
        {
            var rgb = reader.GetAttribute("rgb");
            if (!string.IsNullOrEmpty(rgb) && rgb.Length == 8 && rgb.StartsWith("FF"))
                return "#" + rgb.Substring(2);
            
            // Handle indexed colors and theme colors (simplified)
            var indexed = reader.GetAttribute("indexed");
            if (!string.IsNullOrEmpty(indexed))
            {
                // Could look up in indexed color table
                return null;
            }
            
            return null;
        }

        private static int? ParseInt(string? value)
        {
            if (int.TryParse(value, out int result))
                return result;
            return null;
        }

        private static double? ParseDouble(string? value)
        {
            if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double result))
                return result;
            return null;
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
            List<string> sharedStrings, Dictionary<int, CellStyle> styles, Dictionary<string, Hyperlink> hyperlinks)
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
                    if (reader.Name == "cols")
                    {
                        // process column definitions
                        var depth = reader.Depth;
                        while (reader.Read() && !(reader.NodeType == XmlNodeType.EndElement && reader.Name == "cols" && reader.Depth == depth))
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.Name == "col")
                            {
                                var min = ParseInt(reader.GetAttribute("min"));
                                var max = ParseInt(reader.GetAttribute("max"));
                                var width = ParseDouble(reader.GetAttribute("width"));
                                var hidden = reader.GetAttribute("hidden") == "1";
                                
                                if (min.HasValue && max.HasValue)
                                {
                                    for (int col = min.Value; col <= max.Value; col++)
                                    {
                                        ws.Columns[col] = new ColumnInfo
                                        {
                                            Index = col,
                                            Width = width,
                                            Hidden = hidden
                                        };
                                    }
                                }
                            }
                        }
                    }
                    else if (reader.Name == "mergeCells")
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
                        if (double.TryParse(reader.GetAttribute("ht"), out double ht))
                            currentRow.Height = ht;
                        if (reader.GetAttribute("hidden") == "1")
                            currentRow.Hidden = true;
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

                        // Associate hyperlink if present
                        if (!string.IsNullOrEmpty(rref) && hyperlinks.TryGetValue(rref, out var hyperlink))
                        {
                            cell.Hyperlink = hyperlink;
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