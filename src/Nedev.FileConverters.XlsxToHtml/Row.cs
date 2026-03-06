namespace Nedev.FileConverters.XlsxToHtml
{
    public class Row
    {
        public int Number { get; set; }
        public List<Cell> Cells { get; } = new List<Cell>();
    }
}
