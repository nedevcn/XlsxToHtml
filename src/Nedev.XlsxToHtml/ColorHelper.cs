using System.Collections.Generic;
using System.Globalization;

namespace Nedev.XlsxToHtml
{
    public static class ColorHelper
    {
        private static readonly Dictionary<string, string> _map = new(StringComparer.OrdinalIgnoreCase)
        {
            {"black","#000000"},
            {"white","#FFFFFF"},
            {"red","#FF0000"},
            {"blue","#0000FF"},
            {"green","#008000"},
            {"yellow","#FFFF00"},
            {"magenta","#FF00FF"},
            {"cyan","#00FFFF"},
            {"gray","#808080"},
            {"grey","#808080"},
            {"orange","#FFA500"},
            {"brown","#A52A2A"},
            {"purple","#800080"},
            {"pink","#FFC0CB"},
            {"lime","#00FF00"},
            {"teal","#008080"},
            {"navy","#000080"},
            {"maroon","#800000"},
            {"olive","#808000"},
            {"silver","#C0C0C0"}
        };

        public static bool TryGetColor(string name, out string hex)
        {
            if (string.IsNullOrEmpty(name))
            {
                hex = string.Empty;
                return false;
            }
            return _map.TryGetValue(name.Trim(), out hex);
        }

        /// <summary>
        /// Adds or updates a named color mapping. Name comparisons are case-insensitive.
        /// </summary>
        public static void AddOrUpdate(string name, string hex)
        {
            if (string.IsNullOrEmpty(name) || string.IsNullOrEmpty(hex))
                return;
            // normalize hex
            if (!hex.StartsWith("#"))
                hex = "#" + hex;
            _map[name.Trim()] = hex;
        }
    }
}