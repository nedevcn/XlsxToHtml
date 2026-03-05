namespace Nedev.XlsxToHtml
{
    public class CellStyle
    {
        public int? NumberFormatId { get; set; }
        public string? NumberFormat { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public string? BackgroundColor { get; set; } // hex #RRGGBB
        public string? FontColor { get; set; }
        // other style properties as needed

        public string ToCss()
        {
            var sb = new System.Text.StringBuilder();
            if (Bold) sb.Append("font-weight:bold;");
            if (Italic) sb.Append("font-style:italic;");
            if (Underline) sb.Append("text-decoration:underline;");
            if (!string.IsNullOrEmpty(BackgroundColor)) sb.Append($"background-color:{BackgroundColor};");
            if (!string.IsNullOrEmpty(FontColor)) sb.Append($"color:{FontColor};");
            // number formats handled separately by value conversion
            return sb.ToString();
        }
    }
}