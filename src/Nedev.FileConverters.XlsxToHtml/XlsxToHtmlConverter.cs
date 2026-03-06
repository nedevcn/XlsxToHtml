using System.IO;
using Nedev.FileConverters.Core;

namespace Nedev.FileConverters.XlsxToHtml
{
    [FileConverter("xlsx", "html")]
    public class XlsxToHtmlConverter : IFileConverter
    {
        public Stream Convert(Stream input)
        {
            // use reader and writer already defined in this library
            var reader = new XlsxReader();
            Workbook wb;
            // prefer stream overload when available
            if (input.CanSeek)
            {
                // ensure position at start
                input.Position = 0;
                wb = reader.Read(input);
            }
            else
            {
                // fallback: copy to MemoryStream
                using var ms = new MemoryStream();
                input.CopyTo(ms);
                ms.Position = 0;
                wb = reader.Read(ms);
            }

            var writer = new HtmlWriter();
            var outStream = new MemoryStream();
            using (var sw = new StreamWriter(outStream, System.Text.Encoding.UTF8, 1024, leaveOpen: true))
            {
                writer.Write(wb, sw);
            }
            outStream.Position = 0;
            return outStream;
        }
    }
}