namespace Nedev.FileConverters.XlsxToHtml
{
    public class Row
    {
        public int Number { get; set; }
        public double? Height { get; set; } // in points
        public bool Hidden { get; set; }
        public List<Cell> Cells { get; } = new List<Cell>();
    }
}
