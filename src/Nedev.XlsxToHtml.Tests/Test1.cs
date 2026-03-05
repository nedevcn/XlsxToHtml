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

            var reader = new XlsxReader();
            var wb = reader.Read(tempFile);
            Assert.AreEqual(1, wb.Sheets.Count);
            Assert.AreEqual("Sheet1", wb.Sheets[0].Name);

            var writer = new HtmlWriter();
            string html = writer.Convert(wb);
            Assert.IsTrue(html.Contains("Hello"));
            Assert.IsTrue(html.Contains("<table"));
            // style from fonts/fills should appear
            Assert.IsTrue(html.Contains("font-weight:bold"));
            Assert.IsTrue(html.Contains("background-color:#FFFF00"));
            // mergeCell should cause rowspan
            Assert.IsTrue(html.Contains("rowspan=\"2\""));
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
"  <fonts count=\"1\">\n" +
"    <font><b/><color rgb=\"FFFF0000\"/></font>\n" +
"  </fonts>\n" +
"  <fills count=\"1\">\n" +
"    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFF00\"/></patternFill></fill>\n" +
"  </fills>\n" +
"  <cellXfs count=\"1\">\n" +
"    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\"/>\n" +
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
"  </sheetData>\n" +
"  <mergeCells>\n" +
"    <mergeCell ref=\"A1:A2\"/>\n" +
"  </mergeCells>\n" +
"</worksheet>");
        }
    }
}
