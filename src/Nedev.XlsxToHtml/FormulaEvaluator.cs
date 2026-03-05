using System;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Nedev.XlsxToHtml
{
    public static class FormulaEvaluator
    {
        public static string Evaluate(string formula, Worksheet sheet)
        {
            if (string.IsNullOrEmpty(formula))
                return string.Empty;
            if (formula.StartsWith("="))
                formula = formula.Substring(1);
            // handle SUM and AVERAGE iteratively until none left
            formula = Regex.Replace(formula, @"\bSUM\(([^)]*)\)", m => EvaluateSum(m.Groups[1].Value, sheet));
            formula = Regex.Replace(formula, @"\bAVERAGE\(([^)]*)\)", m => EvaluateAverage(m.Groups[1].Value, sheet));
            // replace cell refs with numeric values
            formula = Regex.Replace(formula, @"([A-Z]+)(\d+)", m =>
            {
                var col = ColNameToNumber(m.Groups[1].Value);
                var row = int.Parse(m.Groups[2].Value);
                var cell = sheet.GetCell(row, col);
                if (cell != null && double.TryParse(cell.Value, out var d))
                    return d.ToString(CultureInfo.InvariantCulture);
                return "0";
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

        private static string EvaluateSum(string arg, Worksheet sheet)
        {
            double sum = SumRangeOrList(arg, sheet);
            return sum.ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateAverage(string arg, Worksheet sheet)
        {
            var parts = arg.Split(',');
            double total = 0;
            int count = 0;
            foreach (var p in parts)
            {
                var val = SumRangeOrList(p, sheet);
                total += val;
                count++;
            }
            if (count == 0) return "0";
            return (total / count).ToString(CultureInfo.InvariantCulture);
        }

        private static double SumRangeOrList(string arg, Worksheet sheet)
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
                        var cell = sheet.GetCell(r, c);
                        if (cell != null && double.TryParse(cell.Value, out var d))
                            sum += d;
                    }
                }
                return sum;
            }
            else
            {
                if (double.TryParse(arg, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                    return d;
                return 0;
            }
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

        private static int ColNameToNumber(string name)
        {
            int col = 0;
            foreach (char c in name)
            {
                col = col * 26 + (c - 'A' + 1);
            }
            return col;
        }
    }
}