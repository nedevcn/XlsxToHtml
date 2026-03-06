namespace Nedev.FileConverters.XlsxToHtml
{
    public class Worksheet
    {
        public string Name { get; set; } = string.Empty;
        public int Index { get; set; }
        public List<Row> Rows { get; } = new List<Row>();
        public List<MergeCell> Merges { get; } = new List<MergeCell>();

        public Cell? GetCell(int row, int col)
        {
            var r = Rows.Find(rr => rr.Number == row);
            if (r == null) return null;
            return r.Cells.Find(c => c.Column == col);
        }
    }
}
