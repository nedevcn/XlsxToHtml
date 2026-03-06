namespace Nedev.FileConverters.XlsxToHtml
{
    public enum CellType
    {
        Unknown,
        SharedString,
        Number,
        Boolean,
        InlineString,
        Date,
        Error,
        Formula // we'll emit cached value
    }

    public class Cell
    {
        public int Row { get; set; }
        public int Column { get; set; }
        public CellType Type { get; set; } = CellType.Unknown;
        public string? Value { get; set; }
        public CellStyle? Style { get; set; }
        public string? Formula { get; set; }
    }
}
