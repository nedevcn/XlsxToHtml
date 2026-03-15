namespace Nedev.FileConverters.XlsxToHtml
{
    public enum HorizontalAlignment
    {
        Default,
        Left,
        Center,
        Right,
        Justify
    }

    public enum VerticalAlignment
    {
        Default,
        Top,
        Middle,
        Bottom
    }

    public class BorderStyle
    {
        public string? Color { get; set; }
        public string? Style { get; set; } // thin, medium, thick, dashed, dotted, etc.
        public double? Width { get; set; } // in points
    }

    public class CellStyle
    {
        public int? NumberFormatId { get; set; }
        public string? NumberFormat { get; set; }
        
        // Font properties
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strikethrough { get; set; }
        public double? FontSize { get; set; } // in points
        public string? FontName { get; set; }
        public string? FontColor { get; set; } // hex #RRGGBB
        
        // Fill/Background
        public string? BackgroundColor { get; set; } // hex #RRGGBB
        
        // Alignment
        public HorizontalAlignment HorizontalAlign { get; set; }
        public VerticalAlignment VerticalAlign { get; set; }
        public bool WrapText { get; set; }
        
        // Borders
        public BorderStyle? BorderLeft { get; set; }
        public BorderStyle? BorderRight { get; set; }
        public BorderStyle? BorderTop { get; set; }
        public BorderStyle? BorderBottom { get; set; }
        public BorderStyle? BorderDiagonal { get; set; }

        public string ToCss()
        {
            var sb = new System.Text.StringBuilder();
            
            // Font styles
            if (Bold) sb.Append("font-weight:bold;");
            if (Italic) sb.Append("font-style:italic;");
            if (Underline) sb.Append("text-decoration:underline;");
            if (Strikethrough) sb.Append("text-decoration:line-through;");
            if (FontSize.HasValue) sb.Append($"font-size:{FontSize.Value}pt;");
            if (!string.IsNullOrEmpty(FontName)) sb.Append($"font-family:{FontName};");
            if (!string.IsNullOrEmpty(FontColor)) sb.Append($"color:{FontColor};");
            
            // Background
            if (!string.IsNullOrEmpty(BackgroundColor)) sb.Append($"background-color:{BackgroundColor};");
            
            // Alignment
            switch (HorizontalAlign)
            {
                case HorizontalAlignment.Left: sb.Append("text-align:left;"); break;
                case HorizontalAlignment.Center: sb.Append("text-align:center;"); break;
                case HorizontalAlignment.Right: sb.Append("text-align:right;"); break;
                case HorizontalAlignment.Justify: sb.Append("text-align:justify;"); break;
            }
            
            switch (VerticalAlign)
            {
                case VerticalAlignment.Top: sb.Append("vertical-align:top;"); break;
                case VerticalAlignment.Middle: sb.Append("vertical-align:middle;"); break;
                case VerticalAlignment.Bottom: sb.Append("vertical-align:bottom;"); break;
            }
            
            if (WrapText) sb.Append("white-space:normal;word-wrap:break-word;");
            else sb.Append("white-space:nowrap;");
            
            // Borders
            AppendBorderCss(sb, BorderLeft, "left");
            AppendBorderCss(sb, BorderRight, "right");
            AppendBorderCss(sb, BorderTop, "top");
            AppendBorderCss(sb, BorderBottom, "bottom");
            
            return sb.ToString();
        }
        
        private static void AppendBorderCss(System.Text.StringBuilder sb, BorderStyle? border, string side)
        {
            if (border == null) return;
            
            var style = border.Style?.ToLowerInvariant() ?? "solid";
            var width = border.Width ?? 1;
            var color = border.Color ?? "#000000";
            
            // Map Excel border styles to CSS
            string cssStyle = style switch
            {
                "thin" => "solid",
                "medium" => "solid",
                "thick" => "solid",
                "dashed" => "dashed",
                "dotted" => "dotted",
                "double" => "double",
                "hair" => "solid",
                _ => "solid"
            };
            
            sb.Append($"border-{side}:{width}px {cssStyle} {color};");
        }
    }
}