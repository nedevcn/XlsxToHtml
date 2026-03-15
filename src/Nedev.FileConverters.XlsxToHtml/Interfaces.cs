namespace Nedev.FileConverters.XlsxToHtml
{
    public interface IXlsxReader
    {
        Workbook Read(string path);
        Workbook Read(System.IO.Stream stream);
    }

    public interface IHtmlWriter
    {
        void Write(Workbook workbook, System.IO.TextWriter output);
        string Convert(Workbook workbook);
    }
}
