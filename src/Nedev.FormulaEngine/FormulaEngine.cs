using System;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Nedev.FormulaEngine
{
    /// <summary>
    /// Simple, zero-dependency formula evaluation engine with extension points.
    /// Not meant to be Excel-perfect but provides arithmetic, cell references,
    /// ranges and a handful of functions (SUM, AVERAGE, MIN, MAX).
    /// </summary>
    public class FormulaEngine
    {
        /// <summary>
        /// Called to resolve a single cell reference (e.g. "A1"). Should return
        /// a numeric value or double.NaN if unavailable.
        /// </summary>
        public Func<string, double> CellResolver { get; set; } = _ => double.NaN;

        public Func<string, double> NameResolver { get; set; } = _ => double.NaN;

        /// <summary>
        /// Evaluate the given formula text (with or without leading '=').  The
        /// engine will call the resolvers for cell lookups when it encounters cell
        /// references or ranges.
        /// </summary>
        public string Evaluate(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                return string.Empty;
            if (formula.StartsWith("="))
                formula = formula.Substring(1);

            // Expand functions we know
            formula = Regex.Replace(formula, @"\bSUM\(([^)]*)\)", m => EvaluateSum(m.Groups[1].Value));
            formula = Regex.Replace(formula, @"\bAVERAGE\(([^)]*)\)", m => EvaluateAverage(m.Groups[1].Value));
            formula = Regex.Replace(formula, @"\bMIN\(([^)]*)\)", m => EvaluateMin(m.Groups[1].Value));
            formula = Regex.Replace(formula, @"\bMAX\(([^)]*)\)", m => EvaluateMax(m.Groups[1].Value));

            // replace cell refs like A1, B23 with numeric value
            formula = Regex.Replace(formula, @"([A-Z]+)(\d+)", m =>
            {
                var cell = m.Value;
                var val = CellResolver(cell);
                return val.ToString(CultureInfo.InvariantCulture);
            });

            // resolve named constants (not used by default)
            formula = Regex.Replace(formula, @"\b[A-Za-z_][A-Za-z0-9_]*\b", m =>
            {
                var name = m.Value;
                // avoid replacing if it's a numeric literal
                if (double.TryParse(name, out _))
                    return name;
                var val = NameResolver(name);
                return val.ToString(CultureInfo.InvariantCulture);
            });

            try
            {
                var dt = new DataTable();
                var result = dt.Compute(formula, null);
                return Convert.ToString(result, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch
            {
                return formula;
            }
        }

        private double SumRangeOrList(string arg)
        {
            arg = arg.Trim();
            if (arg.Contains(":"))
            {
                var seg = arg.Split(':');
                var start = DecodeCellRef(seg[0]);
                var end = DecodeCellRef(seg[1]);
                double sum = 0;
                for (int r = start.row; r <= end.row; r++)
                {
                    for (int c = start.col; c <= end.col; c++)
                    {
                        var cellRef = EncodeCellRef(r, c);
                        var d = CellResolver(cellRef);
                        if (!double.IsNaN(d)) sum += d;
                    }
                }
                return sum;
            }
            else
            {
                if (double.TryParse(arg, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                    return d;
                var cell = CellResolver(arg);
                return double.IsNaN(cell) ? 0 : cell;
            }
        }

        private string EvaluateSum(string arg) => SumRangeOrList(arg).ToString(CultureInfo.InvariantCulture);
        private string EvaluateAverage(string arg)
        {
            var parts = arg.Split(',');
            double total = 0;
            int count = 0;
            foreach (var p in parts)
            {
                total += SumRangeOrList(p);
                count++;
            }
            if (count == 0) return "0";
            return (total / count).ToString(CultureInfo.InvariantCulture);
        }
        private string EvaluateMin(string arg)
        {
            var parts = arg.Split(',');
            double? best = null;
            foreach (var p in parts)
            {
                // compute each individual value or inside range
                var vals = ExpandRange(p);
                foreach (var v in vals)
                {
                    if (best == null || v < best) best = v;
                }
            }
            return (best ?? 0).ToString(CultureInfo.InvariantCulture);
        }
        private string EvaluateMax(string arg)
        {
            var parts = arg.Split(',');
            double? best = null;
            foreach (var p in parts)
            {
                var vals = ExpandRange(p);
                foreach (var v in vals)
                {
                    if (best == null || v > best) best = v;
                }
            }
            return (best ?? 0).ToString(CultureInfo.InvariantCulture);
        }

        private double[] ExpandRange(string arg)
        {
            arg = arg.Trim();
            if (arg.Contains(":"))
            {
                var seg = arg.Split(':');
                var start = DecodeCellRef(seg[0]);
                var end = DecodeCellRef(seg[1]);
                var list = new System.Collections.Generic.List<double>();
                for (int r = start.row; r <= end.row; r++)
                {
                    for (int c = start.col; c <= end.col; c++)
                    {
                        var cellRef = EncodeCellRef(r, c);
                        var d = CellResolver(cellRef);
                        if (!double.IsNaN(d)) list.Add(d);
                    }
                }
                return list.ToArray();
            }
            if (double.TryParse(arg, NumberStyles.Any, CultureInfo.InvariantCulture, out var val))
                return new[] { val };
            var v2 = CellResolver(arg);
            return double.IsNaN(v2) ? Array.Empty<double>() : new[] { v2 };
        }

        private static (int row, int col) DecodeCellRef(string r)
        {
            int col = 0;
            int row = 0;
            foreach (char c in r)
            {
                if (char.IsLetter(c))
                {
                    col = col * 26 + (c - 'A' + 1);
                }
                else if (char.IsDigit(c))
                {
                    row = row * 10 + (c - '0');
                }
            }
            return (row, col);
        }

        private static string EncodeCellRef(int row, int col)
        {
            string s = string.Empty;
            while (col > 0)
            {
                col--;
                s = (char)('A' + (col % 26)) + s;
                col /= 26;
            }
            return s + row.ToString(CultureInfo.InvariantCulture);
        }
    }
}