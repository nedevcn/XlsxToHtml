namespace Nedev.XlsxToHtml
{
    public class MergeCell
    {
        public int StartRow { get; set; }
        public int StartCol { get; set; }
        public int EndRow { get; set; }
        public int EndCol { get; set; }

        public bool Covers(int row, int col)
        {
            return row >= StartRow && row <= EndRow && col >= StartCol && col <= EndCol;
        }

        public bool IsTopLeft(int row, int col)
        {
            return row == StartRow && col == StartCol;
        }
    }
}