using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Nedev.FileConverters.XlsxToHtml
{
    public interface IXlsxReader
    {
        Workbook Read(string path);
        Workbook Read(Stream stream);
        
        // Async methods
        Task<Workbook> ReadAsync(string path, CancellationToken cancellationToken = default);
        Task<Workbook> ReadAsync(Stream stream, CancellationToken cancellationToken = default);
    }

    public interface IHtmlWriter
    {
        void Write(Workbook workbook, TextWriter output);
        string Convert(Workbook workbook);
        
        // Async methods
        Task WriteAsync(Workbook workbook, TextWriter output, CancellationToken cancellationToken = default);
        Task<string> ConvertAsync(Workbook workbook, CancellationToken cancellationToken = default);
    }

    // Progress reporting for long operations
    public interface IConversionProgress
    {
        void Report(ConversionProgressInfo progress);
    }

    public class ConversionProgressInfo
    {
        public int TotalSheets { get; set; }
        public int CurrentSheet { get; set; }
        public string? CurrentSheetName { get; set; }
        public int TotalRows { get; set; }
        public int ProcessedRows { get; set; }
        public double PercentComplete => TotalRows > 0 ? (double)ProcessedRows / TotalRows * 100 : 0;
        public string? Status { get; set; }
    }
}
