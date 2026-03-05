using System;
using System.IO;
using System.IO.Compression;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Nedev.XlsxToHtml;

namespace Nedev.XlsxToHtml.Tests
{
    [TestClass]
    public class XlsxConversionTests
    {
        [TestMethod]
        public void SimpleWorkbook_IsConvertedToHtml()
        {
            string tempFile = Path.GetTempFileName();
            File.Delete(tempFile);
            tempFile = tempFile + ".xlsx";
            CreateMinimalWorkbook(tempFile);
            Console.WriteLine($"Workbook created at {tempFile}");
            using var za = ZipFile.OpenRead(tempFile);
            using var log = new StreamWriter("d:\\Project\\FileConverters\\XlsxToHtml\\entries.txt", false);
            log.WriteLine("workbookPath=" + tempFile);
            foreach (var e in za.Entries)
            {
                log.WriteLine(e.FullName);
            }
            log.Flush();
            // dump sharedStrings content for debugging
            var ssEntry = za.GetEntry("xl/sharedStrings.xml");
            if (ssEntry != null)
            {
                using var sr = new StreamReader(ssEntry.Open());
                File.WriteAllText("d:\\Project\\FileConverters\\XlsxToHtml\\sharedstrings.xml", sr.ReadToEnd());
            }
            // also dump worksheet XML
            var shEntry = za.GetEntry("xl/worksheets/sheet1.xml");
            if (shEntry != null)
            {
                using var sr2 = new StreamReader(shEntry.Open());
                File.WriteAllText("d:\\Project\\FileConverters\\XlsxToHtml\\sheet1.xml", sr2.ReadToEnd());
            }

            var reader = new XlsxReader();
            var wb = reader.Read(tempFile);
            Assert.AreEqual(1, wb.Sheets.Count);
            Assert.AreEqual("Sheet1", wb.Sheets[0].Name);
            // verify first cell value
            var first = wb.Sheets[0].Rows[0].Cells[0];
            // write info to same entries log
            log.WriteLine($"cell0 type={first.Type} value='{first.Value}'");
            Assert.AreEqual("Hello", first.Value);

            // ensure numeric cell isn't misclassified as shared string
            var third = wb.Sheets[0].Rows[2].Cells[0]; // B3 originally
            Assert.AreEqual(CellType.Number, third.Type);
            Assert.IsTrue(third.Value.StartsWith("1234"));

            var writer = new HtmlWriter();
            string html = writer.Convert(wb);
            // dump to disk for inspection
            File.WriteAllText(Path.Combine(Path.GetTempPath(), "debug.html"), html);
            Assert.IsTrue(html.Contains("Hello"));
            Assert.IsTrue(html.Contains("<table"));
            // style from fonts/fills should appear
            Assert.IsTrue(html.Contains("font-weight:bold"));
            Assert.IsTrue(html.Contains("background-color:#FFFF00"));
            // mergeCell should cause rowspan
            Assert.IsTrue(html.Contains("rowspan=\"2\""));
            // numeric cell should be formatted by custom format
            Assert.IsTrue(html.Contains("1,234.57"));
            // percent cell
            Assert.IsTrue(html.Contains("12.34%"));
            // fraction cell approximate
            Assert.IsTrue(html.Contains("3 14159/100000"));
            // positive cell is red
            Assert.IsTrue(html.Contains("color:#FF0000"));
            // negative cell should turn blue
            Assert.IsTrue(html.Contains("color:#0000FF"));
            Assert.IsTrue(html.Contains("-1,234.56"));
            // hex color formatting 00FF00 for positive and FF00FF for negative
            Assert.IsTrue(html.Contains("color:#00FF00"));
            Assert.IsTrue(html.Contains("color:#FF00FF"));
            // formula should be preserved in title
            Assert.IsTrue(html.Contains("title=\"=SUM(C4:C5)\""));

            // now evaluate formulas
            writer.EvaluateFormulas = true;
            string htmlEval = writer.Convert(wb);
            // original cached value should be replaced by computed sum of C4:C5 (0.1234)
            Assert.IsTrue(htmlEval.Contains(">0.1234<"));

            // custom named color mapping
            ColorHelper.AddOrUpdate("orchid", "#DA70D6");
            Assert.IsTrue(html.Contains("color:#DA70D6"));
        }

        private static void CreateMinimalWorkbook(string path)
        {
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            using var archive = new ZipArchive(fs, ZipArchiveMode.Create);

            void Add(string name, string content)
            {
                var entry = archive.CreateEntry(name, CompressionLevel.Optimal);
                using var sw = new StreamWriter(entry.Open());
                sw.Write(content);
            }

            Add("[Content_Types].xml",
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
"<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n" +
"  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n" +
"  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n" +
"  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n" +
"  <Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n" +
"  <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n" +
"</Types>");

            Add("_rels/.rels",
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
"  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\n" +
"</Relationships>");

            Add("xl/_rels/workbook.xml.rels",
"<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n" +
"<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n" +
"  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>\n" +
"</Relationships>");

            Add("xl/workbook.xml",
"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
"<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
"  <sheets>\n" +
"    <sheet name=\"Sheet1\" sheetId=\"1\"/>\n" +
"  </sheets>\n" +
"</workbook>");

            Add("xl/sharedStrings.xml",
"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
"<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"1\" uniqueCount=\"1\">\n" +
"  <si><t>Hello</t></si>\n" +
"</sst>");

            Add("xl/styles.xml",
"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
"<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
"  <numFmts count=\"5\">\n" +
"    <numFmt numFmtId=\"164\" formatCode=\"#,##0.00\"/>\n" +
"    <numFmt numFmtId=\"165\" formatCode=\"0.00%\"/>\n" +
"    <numFmt numFmtId=\"166\" formatCode=\"# ?/??\"/>\n" +
"    <numFmt numFmtId=\"167\" formatCode=\"[Red]#,##0.00;[Blue]-#,##0.00\"/>\n" +
"    <numFmt numFmtId=\"168\" formatCode=\"[00FF00]#,##0;[FF00FF]-#,##0\"/>\n" +
"  </numFmts>\n" +
"  <fonts count=\"1\">\n" +
"    <font><b/><color rgb=\"FFFF0000\"/></font>\n" +
"  </fonts>\n" +
"  <fills count=\"1\">\n" +
"    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFF00\"/></patternFill></fill>\n" +
"  </fills>\n" +
"  <cellXfs count=\"6\">\n" +
"    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\"/>\n" +
"    <xf numFmtId=\"164\" fontId=\"0\" fillId=\"0\"/>\n" +
"    <xf numFmtId=\"165\" fontId=\"0\" fillId=\"0\"/>\n" +
"    <xf numFmtId=\"166\" fontId=\"0\" fillId=\"0\"/>\n" +
"    <xf numFmtId=\"167\" fontId=\"0\" fillId=\"0\"/>\n" +
"    <xf numFmtId=\"168\" fontId=\"0\" fillId=\"0\"/>\n" +
"  </cellXfs>\n" +
"</styleSheet>");

            Add("xl/worksheets/sheet1.xml",
"<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
"<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n" +
"  <sheetData>\n" +
"    <row r=\"1\">\n" +
"      <c r=\"A1\" t=\"s\" s=\"0\"><v>0</v></c>\n" +
"    </row>\n" +
"    <row r=\"2\">\n" +
"      <c r=\"A2\" t=\"s\" s=\"0\"><v>0</v></c>\n" +
"    </row>\n" +
"    <row r=\"3\">\n" +
"      <c r=\"B3\" t=\"n\" s=\"1\" f=\"SUM(C4:C5)\"><v>1234.567</v></c>\n" +
"    </row>\n" +
"    <row r=\"4\">\n" +
"      <c r=\"C4\" t=\"n\" s=\"2\"><v>0.1234</v></c>\n" +
"    </row>\n" +
"    <row r=\"5\">\n" +
"      <c r=\"D5\" t=\"n\" s=\"3\"><v>3.14159</v></c>\n" +
"    </row>\n" +
"    <row r=\"6\">\n" +
"      <c r=\"E6\" t=\"n\" s=\"4\"><v>-1234.56</v></c>\n" +
"    </row>\n" +
"    <row r=\"7\">\n" +
"      <c r=\"F7\" t=\"n\" s=\"5\"><v>789</v></c>\n" +
"    </row>\n" +
"    <row r=\"8\">\n" +
"      <c r=\"G8\" t=\"n\" s=\"6\"><v>1</v></c>\n" +
"    </row>\n" +
"  </sheetData>\n" +
"  <mergeCells>\n" +
"    <mergeCell ref=\"A1:A2\"/>\n" +
"  </mergeCells>\n" +
"</worksheet>");
        }
    }
}
