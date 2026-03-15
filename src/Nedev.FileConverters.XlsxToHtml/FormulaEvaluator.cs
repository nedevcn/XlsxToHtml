using System;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace Nedev.FileConverters.XlsxToHtml
{
    public enum ExcelError
    {
        None,
        DivZero,    // #DIV/0!
        NA,         // #N/A
        Name,       // #NAME?
        Null,       // #NULL!
        Num,        // #NUM!
        Ref,        // #REF!
        Value       // #VALUE!
    }

    public static class FormulaEvaluator
    {
        public static string Evaluate(string formula, Worksheet sheet)
        {
            if (string.IsNullOrEmpty(formula))
                return string.Empty;
            if (formula.StartsWith("="))
                formula = formula.Substring(1);

            // Check for error values first
            var errorCheck = CheckForError(formula);
            if (errorCheck != ExcelError.None)
                return ErrorToString(errorCheck);

            try
            {
                // Handle functions iteratively until none left
                formula = EvaluateFunctions(formula, sheet);

                // Replace cell refs with numeric values
                formula = Regex.Replace(formula, @"([A-Z]+)(\d+)", m =>
                {
                    var col = ColNameToNumber(m.Groups[1].Value);
                    var row = int.Parse(m.Groups[2].Value);
                    var cell = sheet.GetCell(row, col);
                    if (cell != null)
                    {
                        // Check if cell contains an error
                        if (cell.Type == CellType.Error && !string.IsNullOrEmpty(cell.Value))
                        {
                            var cellError = ParseError(cell.Value);
                            if (cellError != ExcelError.None)
                                throw new ExcelFormulaException(cellError);
                        }
                        if (double.TryParse(cell.Value, out var d))
                            return d.ToString(CultureInfo.InvariantCulture);
                    }
                    return "0";
                });

                // Handle comparison operators for IF function results
                formula = HandleComparisons(formula);

                var dt = new DataTable();
                var result = dt.Compute(formula, null);
                return Convert.ToString(result, CultureInfo.InvariantCulture) ?? string.Empty;
            }
            catch (ExcelFormulaException ex)
            {
                return ErrorToString(ex.Error);
            }
            catch (DivideByZeroException)
            {
                return "#DIV/0!";
            }
            catch (OverflowException)
            {
                return "#NUM!";
            }
            catch
            {
                return "#VALUE!";
            }
        }

        private static string EvaluateFunctions(string formula, Worksheet sheet)
        {
            // Handle nested functions by processing innermost first
            int iterations = 0;
            const int maxIterations = 100;

            while (iterations < maxIterations)
            {
                var match = Regex.Match(formula, @"\b([A-Z]+)\(([^()]*(?:\([^()]*\)[^()]*)*)\)");
                if (!match.Success) break;

                var funcName = match.Groups[1].Value.ToUpperInvariant();
                var args = match.Groups[2].Value;
                string result;

                switch (funcName)
                {
                    case "SUM":
                        result = EvaluateSum(args, sheet);
                        break;
                    case "AVERAGE":
                        result = EvaluateAverage(args, sheet);
                        break;
                    case "COUNT":
                        result = EvaluateCount(args, sheet);
                        break;
                    case "COUNTA":
                        result = EvaluateCountA(args, sheet);
                        break;
                    case "COUNTIF":
                        result = EvaluateCountIf(args, sheet);
                        break;
                    case "MAX":
                        result = EvaluateMax(args, sheet);
                        break;
                    case "MIN":
                        result = EvaluateMin(args, sheet);
                        break;
                    case "IF":
                        result = EvaluateIf(args, sheet);
                        break;
                    case "VLOOKUP":
                        result = EvaluateVLookup(args, sheet);
                        break;
                    case "ABS":
                        result = EvaluateAbs(args, sheet);
                        break;
                    case "ROUND":
                        result = EvaluateRound(args, sheet);
                        break;
                    default:
                        result = match.Value; // Unknown function, leave as is
                        break;
                }

                formula = formula.Substring(0, match.Index) + result + formula.Substring(match.Index + match.Length);
                iterations++;
            }

            return formula;
        }

        private static string HandleComparisons(string formula)
        {
            // Convert Excel-style comparisons to DataTable-compatible format
            // Handle = comparison (equality)
            formula = Regex.Replace(formula, @"([^=<>!])=([^=])", "$1==$2");
            return formula;
        }

        private static ExcelError CheckForError(string value)
        {
            if (string.IsNullOrEmpty(value)) return ExcelError.None;
            var upper = value.ToUpperInvariant().Trim();
            return upper switch
            {
                "#DIV/0!" => ExcelError.DivZero,
                "#N/A" => ExcelError.NA,
                "#NAME?" => ExcelError.Name,
                "#NULL!" => ExcelError.Null,
                "#NUM!" => ExcelError.Num,
                "#REF!" => ExcelError.Ref,
                "#VALUE!" => ExcelError.Value,
                _ => ExcelError.None
            };
        }

        private static ExcelError ParseError(string value)
        {
            return CheckForError(value);
        }

        private static string ErrorToString(ExcelError error)
        {
            return error switch
            {
                ExcelError.DivZero => "#DIV/0!",
                ExcelError.NA => "#N/A",
                ExcelError.Name => "#NAME?",
                ExcelError.Null => "#NULL!",
                ExcelError.Num => "#NUM!",
                ExcelError.Ref => "#REF!",
                ExcelError.Value => "#VALUE!",
                _ => string.Empty
            };
        }

        private static string EvaluateSum(string arg, Worksheet sheet)
        {
            double sum = SumRangeOrList(arg, sheet);
            return sum.ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateAverage(string arg, Worksheet sheet)
        {
            var values = GetRangeValues(arg, sheet);
            if (values.Count == 0) return "0";
            var avg = values.Average();
            return avg.ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateCount(string arg, Worksheet sheet)
        {
            var values = GetRangeValues(arg, sheet);
            return values.Count.ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateCountA(string arg, Worksheet sheet)
        {
            var cells = GetRangeCells(arg, sheet);
            int count = 0;
            foreach (var cell in cells)
            {
                if (cell != null && !string.IsNullOrEmpty(cell.Value))
                    count++;
            }
            return count.ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateCountIf(string arg, Worksheet sheet)
        {
            var parts = SplitArgs(arg);
            if (parts.Count < 2) return "0";

            var range = parts[0];
            var criteria = parts[1].Trim('\'', '"', ' ');

            var cells = GetRangeCells(range, sheet);
            int count = 0;

            foreach (var cell in cells)
            {
                if (cell == null || string.IsNullOrEmpty(cell.Value)) continue;

                if (MatchesCriteria(cell.Value, criteria))
                    count++;
            }

            return count.ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateMax(string arg, Worksheet sheet)
        {
            var values = GetRangeValues(arg, sheet);
            if (values.Count == 0) return "0";
            return values.Max().ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateMin(string arg, Worksheet sheet)
        {
            var values = GetRangeValues(arg, sheet);
            if (values.Count == 0) return "0";
            return values.Min().ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateIf(string arg, Worksheet sheet)
        {
            var parts = SplitArgs(arg);
            if (parts.Count < 2) return "#VALUE!";

            var condition = parts[0];
            var trueValue = parts[1];
            var falseValue = parts.Count > 2 ? parts[2] : "FALSE";

            // Evaluate the condition
            bool result = EvaluateCondition(condition, sheet);

            return result ? trueValue : falseValue;
        }

        private static string EvaluateVLookup(string arg, Worksheet sheet)
        {
            var parts = SplitArgs(arg);
            if (parts.Count < 3) return "#N/A";

            var lookupValue = parts[0].Trim('\'', '"', ' ');
            var tableArray = parts[1];
            var colIndex = int.TryParse(parts[2], out var idx) ? idx : 1;
            var rangeLookup = parts.Count > 3 ? parts[3].Trim().ToUpperInvariant() == "TRUE" : true;

            // Parse table array range
            var rangeMatch = Regex.Match(tableArray, @"([A-Z]+)(\d+):([A-Z]+)(\d+)");
            if (!rangeMatch.Success) return "#REF!";

            var startCol = ColNameToNumber(rangeMatch.Groups[1].Value);
            var startRow = int.Parse(rangeMatch.Groups[2].Value);
            var endCol = ColNameToNumber(rangeMatch.Groups[3].Value);
            var endRow = int.Parse(rangeMatch.Groups[4].Value);

            var targetCol = startCol + colIndex - 1;
            if (targetCol > endCol) return "#REF!";

            // Search in first column
            for (int r = startRow; r <= endRow; r++)
            {
                var cell = sheet.GetCell(r, startCol);
                if (cell != null && cell.Value == lookupValue)
                {
                    var resultCell = sheet.GetCell(r, targetCol);
                    return resultCell?.Value ?? string.Empty;
                }
            }

            return "#N/A";
        }

        private static string EvaluateAbs(string arg, Worksheet sheet)
        {
            var value = EvaluateSingleValue(arg, sheet);
            return Math.Abs(value).ToString(CultureInfo.InvariantCulture);
        }

        private static string EvaluateRound(string arg, Worksheet sheet)
        {
            var parts = SplitArgs(arg);
            if (parts.Count < 2) return "#VALUE!";

            var number = EvaluateSingleValue(parts[0], sheet);
            var digits = (int)EvaluateSingleValue(parts[1], sheet);

            return Math.Round(number, digits).ToString(CultureInfo.InvariantCulture);
        }

        private static bool EvaluateCondition(string condition, Worksheet sheet)
        {
            // Handle comparison operators
            var match = Regex.Match(condition, @"(.+?)([=<>]+)(.+)");
            if (match.Success)
            {
                var left = match.Groups[1].Value.Trim();
                var op = match.Groups[2].Value.Trim();
                var right = match.Groups[3].Value.Trim();

                var leftVal = EvaluateSingleValue(left, sheet);
                var rightVal = EvaluateSingleValue(right, sheet);

                return op switch
                {
                    "=" or "==" => Math.Abs(leftVal - rightVal) < 0.0001,
                    "<>" or "!=" => Math.Abs(leftVal - rightVal) >= 0.0001,
                    ">" => leftVal > rightVal,
                    "<" => leftVal < rightVal,
                    ">=" => leftVal >= rightVal,
                    "<=" => leftVal <= rightVal,
                    _ => false
                };
            }

            // Simple boolean evaluation
            var result = EvaluateSingleValue(condition, sheet);
            return result != 0;
        }

        private static bool MatchesCriteria(string value, string criteria)
        {
            // Handle numeric comparisons in criteria
            if (criteria.StartsWith(">"))
            {
                if (double.TryParse(criteria.Substring(1), out var num) &&
                    double.TryParse(value, out var val))
                    return val > num;
            }
            else if (criteria.StartsWith("<"))
            {
                if (double.TryParse(criteria.Substring(1), out var num) &&
                    double.TryParse(value, out var val))
                    return val < num;
            }
            else if (criteria.StartsWith("="))
            {
                return value == criteria.Substring(1);
            }

            // Simple equality
            return value == criteria;
        }

        private static double EvaluateSingleValue(string arg, Worksheet sheet)
        {
            arg = arg.Trim();

            // Direct number
            if (double.TryParse(arg, NumberStyles.Any, CultureInfo.InvariantCulture, out var direct))
                return direct;

            // Cell reference
            var match = Regex.Match(arg, @"^([A-Z]+)(\d+)$");
            if (match.Success)
            {
                var col = ColNameToNumber(match.Groups[1].Value);
                var row = int.Parse(match.Groups[2].Value);
                var cell = sheet.GetCell(row, col);
                if (cell != null && double.TryParse(cell.Value, out var val))
                    return val;
            }

            return 0;
        }

        private static List<double> GetRangeValues(string arg, Worksheet sheet)
        {
            var values = new List<double>();
            var cells = GetRangeCells(arg, sheet);

            foreach (var cell in cells)
            {
                if (cell != null && double.TryParse(cell.Value, out var d))
                    values.Add(d);
            }

            return values;
        }

        private static List<Cell> GetRangeCells(string arg, Worksheet sheet)
        {
            var cells = new List<Cell>();
            arg = arg.Trim();

            if (arg.Contains(":"))
            {
                var seg = arg.Split(':');
                var start = DecodeCellRef(seg[0]);
                var end = DecodeCellRef(seg[1]);

                for (int r = start.row; r <= end.row; r++)
                {
                    for (int c = start.col; c <= end.col; c++)
                    {
                        var cell = sheet.GetCell(r, c);
                        if (cell != null)
                            cells.Add(cell);
                    }
                }
            }
            else
            {
                var cellRef = DecodeCellRef(arg);
                var cell = sheet.GetCell(cellRef.row, cellRef.col);
                if (cell != null)
                    cells.Add(cell);
            }

            return cells;
        }

        private static List<string> SplitArgs(string arg)
        {
            var parts = new List<string>();
            var current = new System.Text.StringBuilder();
            int depth = 0;

            foreach (char c in arg)
            {
                if (c == '(') depth++;
                else if (c == ')') depth--;

                if (c == ',' && depth == 0)
                {
                    parts.Add(current.ToString().Trim());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }

            if (current.Length > 0)
                parts.Add(current.ToString().Trim());

            return parts;
        }

        private static double SumRangeOrList(string arg, Worksheet sheet)
        {
            var values = GetRangeValues(arg, sheet);
            return values.Sum();
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

    public class ExcelFormulaException : Exception
    {
        public ExcelError Error { get; }

        public ExcelFormulaException(ExcelError error) : base($"Excel error: {error}")
        {
            Error = error;
        }
    }
}