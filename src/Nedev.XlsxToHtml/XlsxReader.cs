using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace Nedev.XlsxToHtml
{
    public class XlsxReader : IXlsxReader
    {
        public Workbook Read(string path)
        {
            var workbook = new Workbook();
            using var archive = ZipFile.OpenRead(path);

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
            using var reader = XmlReader.Create(stream);
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "si")
                {
                    result.Add(ReadStringItem(reader));
                }
            }
            return result;

            static string ReadStringItem(XmlReader r)
            {
                var sb = new System.Text.StringBuilder();
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.Element && r.Name == "t")
                    {
                        sb.Append(r.ReadElementContentAsString());
                    }
                    else if (r.NodeType == XmlNodeType.EndElement && r.Name == "si")
                    {
                        break;
                    }
                }
                return sb.ToString();
            }
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
                            while (reader.Read())
                            {
                                if (reader.NodeType == XmlNodeType.Element && reader.Name == "v")
                                {
                                    var raw = reader.ReadElementContentAsString();
                                    cell.Value = InterpretValue(raw, cell.Type, sharedStrings, cell.Style);
                                }
                                else if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "c")
                                {
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
                    return string.Empty;
                case CellType.Boolean:
                    return raw == "1" ? "TRUE" : "FALSE";
                case CellType.Number:
                    if (style?.NumberFormatId is int id && IsDateFormat(id))
                    {
                        if (double.TryParse(raw, out double oa))
                        {
                            var dt = DateTime.FromOADate(oa);
                            return dt.ToString("yyyy-MM-dd HH:mm:ss");
                        }
                    }
                    return raw;
                default:
                    return raw;
            }
        }

        private static bool IsDateFormat(int numFmtId)
        {
            // built-in date formats 14-22
            return numFmtId >= 14 && numFmtId <= 22;
        }
    }
}