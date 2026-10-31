#region Copyright (C) 2003-2026 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports  											}
{																	}
{	Copyright (C) 2003-2026 Stimulsoft     							}
{	ALL RIGHTS RESERVED												}
{																	}
{	The entire contents of this file is protected by U.S. and		}
{	International Copyright Laws. Unauthorized reproduction,		}
{	reverse-engineering, and distribution of all or any portion of	}
{	the code contained in this file is strictly prohibited and may	}
{	result in severe civil and criminal penalties and will be		}
{	prosecuted to the maximum extent possible under the law.		}
{																	}
{	RESTRICTIONS													}
{																	}
{	THIS SOURCE CODE AND ALL RESULTING INTERMEDIATE FILES			}
{	ARE CONFIDENTIAL AND PROPRIETARY								}
{	TRADE SECRETS OF Stimulsoft										}
{																	}
{	CONSULT THE END USER LICENSE AGREEMENT FOR INFORMATION ON		}
{	ADDITIONAL RESTRICTIONS.										}
{																	}
{*******************************************************************}
*/
#endregion Copyright (C) 2003-2026 Stimulsoft

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Import.CrystalReports.Formulas
{
    public partial class StiCrystalTranslator
    {
        #region Methods
        /// <summary>
        /// Translates a call of a function of Crystal or a word without parentheses like CurrentDate or PageNumber.
        /// The functions are expressed through the members of the base types and the functions of the dictionary of
        /// Stimulsoft, so the result works both in the compilation and in the interpretation mode.
        /// </summary>
        private StiCrystalExpression TranslateCall(string name, List<StiCrystalNode> argumentNodes, int position)
        {
            var lowerName = name.ToLowerInvariant();
            if (lowerName == "iif")
                return TranslateIif(argumentNodes, position);

            var args = argumentNodes.Select(TranslateNode).ToList();
            var result = TranslateStringFunction(lowerName, args, position)
                ?? TranslateDateFunction(lowerName, args, position)
                ?? TranslateMathFunction(lowerName, args, position)
                ?? TranslateDataFunction(lowerName, args, position)
                ?? TranslateFormatConstant(lowerName, args);

            if (result == null)
                throw new StiCrystalFormulaException($"The function {name} with {args.Count} arguments is not supported",
                    position);

            return result;
        }

        private StiCrystalExpression TranslateIif(List<StiCrystalNode> argumentNodes, int position)
        {
            if (argumentNodes.Count != 3)
                throw new StiCrystalFormulaException("IIF needs three arguments", position);

            var node = new StiCrystalNode(StiCrystalNodeKind.If, null, position, argumentNodes.ToArray());
            return TranslateIf(node);
        }

        #region Methods.Strings
        private StiCrystalExpression TranslateStringFunction(string name, List<StiCrystalExpression> args, int position)
        {
            switch (name)
            {
                case "totext":
                case "cstr":
                    return TranslateToText(args, position);

                case "length":
                case "len":
                    if (args.Count == 1)
                        return Member(args[0], "Length", typeof(int));

                    break;

                case "trim":
                    return Method(args, "Trim()", typeof(string));

                case "trimleft":
                case "ltrim":
                    return Method(args, "TrimStart()", typeof(string));

                case "trimright":
                case "rtrim":
                    return Method(args, "TrimEnd()", typeof(string));

                case "uppercase":
                case "ucase":
                    return Method(args, "ToUpper()", typeof(string));

                case "lowercase":
                case "lcase":
                    return Method(args, "ToLower()", typeof(string));

                case "propercase":
                    if (args.Count == 1)
                        return Call("ToProperCase", typeof(string), args[0]);

                    break;

                case "left":
                    if (args.Count == 2)
                        return Call("Left", typeof(string), args[0], Coerce(args[1], typeof(int)));

                    break;

                case "right":
                    if (args.Count == 2)
                        return Call("Right", typeof(string), args[0], Coerce(args[1], typeof(int)));

                    break;

                case "mid":
                    return TranslateMid(args);

                case "instr":
                    if (args.Count == 2)
                        return new StiCrystalExpression($"{args[0].ToOperand()}.IndexOf({args[1].Code}) + 1",
                            typeof(int));

                    break;

                case "replace":
                    if (args.Count == 3)
                        return new StiCrystalExpression($"{args[0].ToOperand()}.Replace({args[1].Code}, {args[2].Code})",
                            typeof(string), isAtomic: true);

                    break;

                case "space":
                    if (args.Count == 1)
                        return new StiCrystalExpression($"\"\".PadLeft({Coerce(args[0], typeof(int)).Code})",
                            typeof(string), isAtomic: true);

                    break;

                case "replicatestring":
                    if (args.Count == 2 && args[0].Literal != null && args[0].Literal.Length == 1)
                    {
                        var padding = args[0].Literal.Replace("\\", "\\\\").Replace("'", "\\'");
                        var code = $"\"\".PadLeft({Coerce(args[1], typeof(int)).Code}, '{padding}')";
                        return new StiCrystalExpression(code, typeof(string), isAtomic: true);
                    }

                    break;

                case "chr":
                case "chrw":
                    return TranslateChr(args);

                case "tonumber":
                case "val":
                case "cdbl":
                    return TranslateToNumber(args);

                case "ccur":
                    if (args.Count == 1 && IsNumber(args[0].Type))
                        return Coerce(args[0], typeof(decimal));

                    break;
            }

            return null;
        }

        /// <summary>
        /// ToText formats a value as Crystal does by default, or with the number of decimals, or with a format string.
        /// </summary>
        private StiCrystalExpression TranslateToText(List<StiCrystalExpression> args, int position)
        {
            if (args.Count == 0)
                return null;

            var value = args[0];
            if (args.Count == 1)
                return ToText(value);

            var second = args[1];
            if (second.Type == typeof(string) && second.Literal != null)
                return Format(EscapeFormat(second.Literal), value);

            if (IsNumber(value.Type) && IsNumber(second.Type))
            {
                var separator = args.Count > 2 ? args[2].Literal : null;
                var letter = separator == string.Empty ? "F" : "N";
                int decimals;
                if (int.TryParse(second.Code, NumberStyles.None, CultureInfo.InvariantCulture, out decimals))
                    return Format(letter + decimals, value);

                var code = $"Format(\"{{0:{letter}\" + {second.ToOperand()} + \"}}\", {value.Code})";
                return new StiCrystalExpression(code, typeof(string), isAtomic: true);
            }

            throw new StiCrystalFormulaException("This form of ToText is not supported", position);
        }

        /// <summary>
        /// The default conversion of a value to a string, used by ToText and by the operator &amp;.
        /// </summary>
        private static StiCrystalExpression ToText(StiCrystalExpression value)
        {
            if (value.Type == typeof(string))
                return value;

            if (value.Type == typeof(bool))
                return new StiCrystalExpression($"{value.ToOperand()} ? \"True\" : \"False\"", typeof(string));

            if (value.Type == typeof(decimal))
                return Format("C2", value);

            if (IsNumber(value.Type))
                return Format("N2", value);

            if (IsDate(value.Type))
                return Format("d", value);

            return new StiCrystalExpression($"{value.ToOperand()}.ToString()", typeof(string), isAtomic: true);
        }

        private static StiCrystalExpression Format(string format, StiCrystalExpression value)
        {
            return new StiCrystalExpression($"Format(\"{{0:{format}}}\", {value.Code})", typeof(string), isAtomic: true);
        }

        /// <summary>
        /// Prepares a format string of Crystal for a format item of .NET: the braces and the backslashes are escaped,
        /// the quotes are escaped for the C# string.
        /// </summary>
        private static string EscapeFormat(string format)
        {
            return format.Replace("\\", "\\\\\\\\").Replace("{", "{{").Replace("}", "}}").Replace("\"", "\\\"");
        }

        /// <summary>
        /// Mid of Crystal counts the characters from 1, Mid of Stimulsoft from 0; without the length it takes the rest.
        /// </summary>
        private static StiCrystalExpression TranslateMid(List<StiCrystalExpression> args)
        {
            if (args.Count < 2 || args.Count > 3)
                return null;

            var start = $"{Coerce(args[1], typeof(int)).ToOperand()} - 1";
            int startValue;
            if (int.TryParse(args[1].Code, NumberStyles.None, CultureInfo.InvariantCulture, out startValue))
                start = (startValue - 1).ToString(CultureInfo.InvariantCulture);

            var length = args.Count == 3 ? Coerce(args[2], typeof(int)).Code : "int.MaxValue";
            return new StiCrystalExpression($"Mid({args[0].Code}, {start}, {length})", typeof(string), isAtomic: true);
        }

        private static StiCrystalExpression TranslateChr(List<StiCrystalExpression> args)
        {
            if (args.Count != 1)
                return null;

            int code;
            if (int.TryParse(args[0].Code, NumberStyles.None, CultureInfo.InvariantCulture, out code))
                return new StiCrystalExpression(ToCSharpString(((char)code).ToString()), typeof(string), isAtomic: true);

            return new StiCrystalExpression($"((char){Coerce(args[0], typeof(int)).ToOperand()}).ToString()",
                typeof(string), isAtomic: true);
        }

        private static StiCrystalExpression TranslateToNumber(List<StiCrystalExpression> args)
        {
            if (args.Count != 1)
                return null;

            if (IsNumber(args[0].Type))
                return Coerce(args[0], typeof(double));

            if (args[0].Type == typeof(bool))
                return new StiCrystalExpression($"{args[0].ToOperand()} ? 1.0 : 0.0", typeof(double));

            return Call("ParseDouble", typeof(double), args[0]);
        }
        #endregion

        #region Methods.Dates
        private StiCrystalExpression TranslateDateFunction(string name, List<StiCrystalExpression> args, int position)
        {
            switch (name)
            {
                case "currentdate":
                case "today":
                case "datadate":
                case "printdate":
                    if (args.Count == 0)
                        return new StiCrystalExpression("Today", typeof(DateTime), isAtomic: true);

                    break;

                case "currentdatetime":
                case "now":
                case "currenttime":
                case "datatime":
                case "printtime":
                    if (args.Count == 0)
                        return new StiCrystalExpression("Time", typeof(DateTime), isAtomic: true);

                    break;

                case "year":
                case "month":
                case "day":
                case "hour":
                case "minute":
                case "second":
                    if (args.Count == 1 && IsDate(args[0].Type))
                        return Member(args[0], char.ToUpperInvariant(name[0]) + name.Substring(1), typeof(int));

                    break;

                case "dayofweek":
                case "weekday":
                    if (args.Count == 1 && IsDate(args[0].Type))
                        return new StiCrystalExpression($"(int){args[0].ToOperand()}.DayOfWeek + 1", typeof(int));

                    break;

                case "monthname":
                    if (args.Count >= 1 && args.Count <= 2)
                        return NameOfDate("MMMM", "MMM", $"DateSerial(2000, {args[0].Code}, 1)", args);

                    break;

                case "weekdayname":
                    if (args.Count >= 1 && args.Count <= 2)
                        return NameOfDate("dddd", "ddd", $"DateSerial(2000, 1, 1 + {args[0].ToOperand()})", args);

                    break;

                case "dateadd":
                    return TranslateDateAdd(args, position);

                case "datediff":
                    return TranslateDateDiff(args, position);

                case "date":
                case "cdate":
                case "datevalue":
                    return TranslateDate(args);

                case "datetime":
                case "cdatetime":
                case "datetimevalue":
                    return TranslateDateTime(args);
            }

            return null;
        }

        /// <summary>
        /// MonthName and WeekdayName: the name of a sample date in the current culture, abbreviated if asked.
        /// </summary>
        private static StiCrystalExpression NameOfDate(string fullFormat, string shortFormat, string date,
            List<StiCrystalExpression> args)
        {
            if (args.Count == 1 || args[1].Code == "false")
                return new StiCrystalExpression($"Format(\"{{0:{fullFormat}}}\", {date})",
                    typeof(string), isAtomic: true);

            if (args[1].Code == "true")
                return new StiCrystalExpression($"Format(\"{{0:{shortFormat}}}\", {date})",
                    typeof(string), isAtomic: true);

            var code = $"Format({args[1].ToOperand()} ? \"{{0:{shortFormat}}}\" : \"{{0:{fullFormat}}}\", {date})";
            return new StiCrystalExpression(code, typeof(string), isAtomic: true);
        }

        /// <summary>
        /// DateAdd(interval, number, date); the interval must be a string literal.
        /// </summary>
        private static StiCrystalExpression TranslateDateAdd(List<StiCrystalExpression> args, int position)
        {
            if (args.Count != 3 || args[0].Literal == null)
                throw new StiCrystalFormulaException("DateAdd needs a constant interval", position);

            var date = args[2].ToOperand();
            var number = args[1];
            string code;
            switch (args[0].Literal.ToLowerInvariant())
            {
                case "yyyy":
                    code = $"{date}.AddYears({Coerce(number, typeof(int)).Code})";
                    break;

                case "q":
                    code = $"{date}.AddMonths(3 * {Coerce(number, typeof(int)).ToOperand()})";
                    break;

                case "m":
                    code = $"{date}.AddMonths({Coerce(number, typeof(int)).Code})";
                    break;

                case "y":
                case "d":
                case "w":
                    code = $"{date}.AddDays({number.Code})";
                    break;

                case "ww":
                    code = $"{date}.AddDays(7 * {number.ToOperand()})";
                    break;

                case "h":
                    code = $"{date}.AddHours({number.Code})";
                    break;

                case "n":
                    code = $"{date}.AddMinutes({number.Code})";
                    break;

                case "s":
                    code = $"{date}.AddSeconds({number.Code})";
                    break;

                default:
                    throw new StiCrystalFormulaException($"Unknown interval {args[0].Literal}", position);
            }

            return new StiCrystalExpression(code, typeof(DateTime), isAtomic: true);
        }

        /// <summary>
        /// DateDiff(interval, start, end) counts the boundaries of the intervals between the dates, as Crystal does.
        /// </summary>
        private static StiCrystalExpression TranslateDateDiff(List<StiCrystalExpression> args, int position)
        {
            if (args.Count != 3 || args[0].Literal == null)
                throw new StiCrystalFormulaException("DateDiff needs a constant interval", position);

            var start = args[1].ToOperand();
            var end = args[2].ToOperand();
            string code;
            switch (args[0].Literal.ToLowerInvariant())
            {
                case "yyyy":
                    code = $"{end}.Year - {start}.Year";
                    break;

                case "q":
                    code = $"({end}.Year - {start}.Year) * 4 + ({end}.Month - 1) / 3 - ({start}.Month - 1) / 3";
                    break;

                case "m":
                    code = $"({end}.Year - {start}.Year) * 12 + {end}.Month - {start}.Month";
                    break;

                case "y":
                case "d":
                    code = $"({end}.Date - {start}.Date).Days";
                    break;

                case "w":
                case "ww":
                    code = $"({end}.Date - {start}.Date).Days / 7";
                    break;

                case "h":
                    code = $"(int)({end} - {start}).TotalHours";
                    break;

                case "n":
                    code = $"(int)({end} - {start}).TotalMinutes";
                    break;

                case "s":
                    code = $"(int)({end} - {start}).TotalSeconds";
                    break;

                default:
                    throw new StiCrystalFormulaException($"Unknown interval {args[0].Literal}", position);
            }

            return new StiCrystalExpression(code, typeof(int));
        }

        private static StiCrystalExpression TranslateDate(List<StiCrystalExpression> args)
        {
            if (args.Count == 3)
                return Call("DateSerial", typeof(DateTime), args.Select(a => Coerce(a, typeof(long))).ToArray());

            if (args.Count != 1)
                return null;

            if (IsDate(args[0].Type))
                return Member(args[0], "Date", typeof(DateTime));

            if (args[0].Type == typeof(string))
                return Call("StrToDateTime", typeof(DateTime), args[0]);

            return null;
        }

        private static StiCrystalExpression TranslateDateTime(List<StiCrystalExpression> args)
        {
            if (args.Count == 6)
            {
                var date = Call("DateSerial", typeof(DateTime),
                    args.Take(3).Select(a => Coerce(a, typeof(long))).ToArray());
                var code = $"{date.Code}.AddHours({args[3].Code}).AddMinutes({args[4].Code}).AddSeconds({args[5].Code})";
                return new StiCrystalExpression(code, typeof(DateTime), isAtomic: true);
            }

            return TranslateDate(args);
        }
        #endregion

        #region Methods.Math
        private static StiCrystalExpression TranslateMathFunction(string name, List<StiCrystalExpression> args,
            int position)
        {
            switch (name)
            {
                case "abs":
                    if (args.Count == 1)
                        return Call("Abs", args[0].Type, args[0]);

                    break;

                case "round":
                    if (args.Count == 1)
                        return Call("Round", RoundType(args[0]), Coerce(args[0], RoundType(args[0])));

                    if (args.Count == 2)
                        return Call("Round", RoundType(args[0]), Coerce(args[0], RoundType(args[0])),
                            Coerce(args[1], typeof(int)));

                    break;

                case "truncate":
                case "fix":
                    if (args.Count == 1)
                        return Call("Math.Truncate", RoundType(args[0]), Coerce(args[0], RoundType(args[0])));

                    break;

                case "int":
                    if (args.Count == 1)
                        return Call("Math.Floor", RoundType(args[0]), Coerce(args[0], RoundType(args[0])));

                    break;

                case "remainder":
                    if (args.Count == 2)
                        return TranslateArithmetic("%", args[0], args[1],
                            new StiCrystalNode(StiCrystalNodeKind.Call, name, position));

                    break;

                case "sqr":
                case "sqrt":
                    if (args.Count == 1)
                        return Call("Math.Sqrt", typeof(double), Coerce(args[0], typeof(double)));

                    break;

                case "sgn":
                    if (args.Count == 1)
                        return Call("Math.Sign", typeof(int), args[0]);

                    break;

                case "pi":
                    if (args.Count == 0)
                        return new StiCrystalExpression("Math.PI", typeof(double), isAtomic: true);

                    break;
            }

            return null;
        }

        private static Type RoundType(StiCrystalExpression value)
        {
            return value.Type == typeof(decimal) ? typeof(decimal) : typeof(double);
        }
        #endregion

        #region Methods.Data
        private StiCrystalExpression TranslateDataFunction(string name, List<StiCrystalExpression> args, int position)
        {
            switch (name)
            {
                case "isnull":
                    if (args.Count == 1)
                        return TranslateIsNull(args[0], position);

                    break;

                case "previous":
                case "next":
                case "previousisnull":
                case "nextisnull":
                    if (args.Count == 1)
                        return TranslateNeighbour(name, args[0], position);

                    break;

                case "haslowerbound":
                case "hasupperbound":
                    if (args.Count == 1 && args[0].IsRange)
                        return new StiCrystalExpression("true", typeof(bool), isAtomic: true);

                    break;

                case "minimum":
                case "maximum":
                    if (args.Count == 1 && args[0].IsRange)
                        return name == "minimum" ? args[0].From : args[0].To;

                    return TranslateSummary(name, args, position);

                case "sum":
                case "average":
                case "count":
                case "distinctcount":
                    return TranslateSummary(name, args, position);

                case "pagenumber":
                    return new StiCrystalExpression("PageNumber", typeof(int), isAtomic: true);

                case "totalpagecount":
                    return new StiCrystalExpression("TotalPageCount", typeof(int), isAtomic: true);

                case "pagenofm":
                    return new StiCrystalExpression("PageNofM", typeof(string), isAtomic: true);

                case "recordnumber":
                    return new StiCrystalExpression("Line", typeof(int), isAtomic: true);

                case "groupnumber":
                    return new StiCrystalExpression("GroupLine", typeof(int), isAtomic: true);

                case "onfirstrecord":
                    return new StiCrystalExpression("Line == 1", typeof(bool));

                case "currentfieldvalue":
                    if (args.Count == 0)
                    {
                        var reference = resolver.ResolveCurrentFieldValue();
                        if (reference == null)
                            throw new StiCrystalFormulaException(
                                "CurrentFieldValue is available only in conditional formatting",
                                position);

                        return CreateFromReference(reference);
                    }

                    break;
            }

            return null;
        }

        /// <summary>
        /// IsNull checks the column through the data source, because the typed value of a column is never null.
        /// The value of a formula, of a parameter or of a summary is never null in Stimulsoft.
        /// </summary>
        private static StiCrystalExpression TranslateIsNull(StiCrystalExpression value, int position)
        {
            var reference = value.Reference;
            if (reference == null || string.IsNullOrEmpty(reference.DataSourceName))
                return new StiCrystalExpression("false", typeof(bool), isAtomic: true);

            var code = $"IsNull({reference.DataSourceName}, {ToCSharpString(reference.ColumnName)})";
            return new StiCrystalExpression(code, typeof(bool), isAtomic: true);
        }

        private static StiCrystalExpression TranslateNeighbour(string name, StiCrystalExpression value, int position)
        {
            var reference = value.Reference;
            if (reference == null || string.IsNullOrEmpty(reference.DataSourceName))
                throw new StiCrystalFormulaException($"{name} needs a field of a database", position);

            var function = name == "previous" ? "Previous" : name == "next" ? "Next"
                : name == "previousisnull" ? "PreviousIsNull" : "NextIsNull";

            var type = function.EndsWith("IsNull") ? typeof(bool) : typeof(object);
            var code = $"{function}({reference.DataSourceName}, {ToCSharpString(reference.ColumnName)})";
            if (type == typeof(object) && value.Type != typeof(object))
                return new StiCrystalExpression($"({GetTypeName(value.Type)}){code}", value.Type);

            return new StiCrystalExpression(code, type, isAtomic: true);
        }

        /// <summary>
        /// The summaries over the whole report or over a group: Sum({field}) or Sum({field}, {group condition}).
        /// </summary>
        private StiCrystalExpression TranslateSummary(string name, List<StiCrystalExpression> args, int position)
        {
            if (args.Count < 1 || args.Count > 2)
                return null;

            var function = name == "sum" ? "Sum" : name == "average" ? "Avg" : name == "count" ? "Count"
                : name == "distinctcount" ? "CountDistinct" : name == "minimum" ? "Min" : "Max";

            // The numeric summaries of Stimulsoft return decimal in the compilation mode
            var type = function == "Count" || function == "CountDistinct" ? typeof(int) : typeof(decimal);
            if (function == "Min" || function == "Max")
            {
                if (IsDate(args[0].Type))
                {
                    type = typeof(DateTime);
                    function += "Date";
                }

                else if (args[0].Type == typeof(string))
                {
                    type = typeof(string);
                    function += "Str";
                }
            }

            var band = string.Empty;
            if (args.Count == 2)
            {
                var groupName = args[1].Reference?.Name;
                band = groupName != null ? resolver.ResolveGroupBand(groupName) : null;
                if (band == null)
                    throw new StiCrystalFormulaException("The group of the summary is unknown", position);
            }

            HasSummaries = true;

            // Count of Stimulsoft counts the rows, it takes only the band
            if (function == "Count")
                return new StiCrystalExpression($"Count({band})", type, isAtomic: true);

            var separator = band.Length > 0 ? ", " : string.Empty;
            return new StiCrystalExpression($"{function}({band}{separator}{args[0].Code})", type, isAtomic: true);
        }
        #endregion

        #region Methods.Formatting
        /// <summary>
        /// Returns the constant values which a formula of conditional formatting can give: the constants at the ends of
        /// the branches of If and Select, like crRed, crBold, Color(255, 0, 0), "Arial" or 12.
        /// </summary>
        public static List<object> GetResultConstants(StiCrystalNode node)
        {
            var constants = new List<object>();
            CollectResultConstants(node, constants);
            return constants;
        }

        private static void CollectResultConstants(StiCrystalNode node, List<object> constants)
        {
            switch (node.Kind)
            {
                case StiCrystalNodeKind.If:
                    foreach (var branch in node.Children.Skip(1))
                    {
                        CollectResultConstants(branch, constants);
                    }

                    return;

                case StiCrystalNodeKind.Select:
                    foreach (var branch in node.Children.Skip(1))
                    {
                        CollectResultConstants(branch.Children.Last(), constants);
                    }

                    return;

                case StiCrystalNodeKind.Sequence:
                    if (node.Children.Count > 0)
                        CollectResultConstants(node.Children.Last(), constants);

                    return;
            }

            object value;
            if (TryGetConstant(node, out value) && !constants.Any(c => AreEqual(c, value)))
                constants.Add(value);
        }

        /// <summary>
        /// Translates the condition "the formula gives the value": every end of a branch is compared with the value,
        /// so the condition is boolean and the constant ends become true or false.
        /// </summary>
        public StiCrystalExpression TranslateEquals(StiCrystalNode node, object value)
        {
            return Translate(ReplaceResults(node, value));
        }

        private static StiCrystalNode ReplaceResults(StiCrystalNode node, object value)
        {
            switch (node.Kind)
            {
                case StiCrystalNodeKind.If:
                    var result = new StiCrystalNode(StiCrystalNodeKind.If, null, node.Position, node.Children[0],
                        ReplaceResults(node.Children[1], value));

                    // If without Else gives the default value of the type, 0 or an empty string
                    var elseNode = node.Children.Count > 2
                        ? ReplaceResults(node.Children[2], value)
                        : CreateBoolean(AreEqual(value is string ? (object)string.Empty : 0, value), node.Position);

                    return result.Add(elseNode);

                case StiCrystalNodeKind.Select:
                    var select = new StiCrystalNode(StiCrystalNodeKind.Select, null, node.Position, node.Children[0]);
                    foreach (var branch in node.Children.Skip(1))
                    {
                        var replaced = new StiCrystalNode(StiCrystalNodeKind.Case, branch.Text, branch.Position,
                            branch.Children.Take(branch.Children.Count - 1).ToArray());

                        select.Add(replaced.Add(ReplaceResults(branch.Children.Last(), value)));
                    }

                    return select;

                case StiCrystalNodeKind.Sequence:
                    var sequence = new StiCrystalNode(StiCrystalNodeKind.Sequence, null, node.Position,
                        node.Children.Take(node.Children.Count - 1).ToArray());

                    return sequence.Add(ReplaceResults(node.Children.Last(), value));
            }

            object constant;
            if (TryGetConstant(node, out constant))
                return CreateBoolean(AreEqual(constant, value), node.Position);

            var constantNode = CreateConstant(value, node.Position);
            return new StiCrystalNode(StiCrystalNodeKind.Binary, "=", node.Position, node, constantNode);
        }

        private static StiCrystalNode CreateBoolean(bool value, int position)
        {
            return new StiCrystalNode(StiCrystalNodeKind.Boolean, value ? "true" : "false", position);
        }

        private static StiCrystalNode CreateConstant(object value, int position)
        {
            var text = value as string;
            if (text != null)
                return new StiCrystalNode(StiCrystalNodeKind.String, text, position);

            var number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            var digits = Math.Abs(number).ToString(CultureInfo.InvariantCulture);
            var numberNode = new StiCrystalNode(StiCrystalNodeKind.Number, digits, position);
            return number < 0 ? new StiCrystalNode(StiCrystalNodeKind.Unary, "-", position, numberNode) : numberNode;
        }

        /// <summary>
        /// Gets the value of a constant node: a number, a string, a constant of a color or a style, or Color(r, g, b).
        /// </summary>
        private static bool TryGetConstant(StiCrystalNode node, out object value)
        {
            value = null;
            switch (node.Kind)
            {
                case StiCrystalNodeKind.Number:
                    value = double.Parse(node.Text, CultureInfo.InvariantCulture);
                    return true;

                case StiCrystalNodeKind.String:
                    value = node.Text;
                    return true;

                case StiCrystalNodeKind.Unary:
                    object operand;
                    if (node.Text != "-" || !TryGetConstant(node.Children[0], out operand) || operand is string)
                        return false;

                    value = -(double)operand;
                    return true;

                case StiCrystalNodeKind.Identifier:
                    int constant;
                    if (!StiCrystalFormatConstants.TryGetValue(node.Text.ToLowerInvariant(), out constant))
                        return false;

                    value = (double)constant;
                    return true;

                case StiCrystalNodeKind.Call:
                    if (!string.Equals(node.Text, "color", StringComparison.OrdinalIgnoreCase) ||
                        node.Children.Count != 3 || node.Children.Any(c => c.Kind != StiCrystalNodeKind.Number))
                        return false;

                    var parts = node.Children.Select(c => int.Parse(c.Text, CultureInfo.InvariantCulture)).ToList();
                    value = (double)(parts[0] + parts[1] * 256 + parts[2] * 65536);
                    return true;
            }

            return false;
        }

        private static bool AreEqual(object first, object second)
        {
            if (first is string || second is string)
                return string.Equals(first as string, second as string, StringComparison.OrdinalIgnoreCase);

            return Convert.ToDouble(first, CultureInfo.InvariantCulture) ==
                Convert.ToDouble(second, CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// The constants of colors and font styles of the formulas of conditional formatting. A color is a number
        /// red + green * 256 + blue * 65536 as in Crystal, a style is a number from StiCrystalFormatConstants,
        /// DefaultAttribute is -1. The importer compares the formula with these numbers to make the conditions.
        /// </summary>
        private static StiCrystalExpression TranslateFormatConstant(string name, List<StiCrystalExpression> args)
        {
            if (name == "color" && args.Count == 3)
            {
                var code = $"{Coerce(args[0], typeof(int)).ToOperand()} + " +
                    $"{Coerce(args[1], typeof(int)).ToOperand()} * 256 + " +
                    $"{Coerce(args[2], typeof(int)).ToOperand()} * 65536";
                return new StiCrystalExpression(code, typeof(int));
            }

            if (args.Count != 0)
                return null;

            int value;
            if (!StiCrystalFormatConstants.TryGetValue(name, out value))
                return null;

            var number = value.ToString(CultureInfo.InvariantCulture);
            return new StiCrystalExpression(number, typeof(int), isAtomic: value >= 0);
        }
        #endregion

        private static StiCrystalExpression Member(StiCrystalExpression value, string member, Type type)
        {
            return new StiCrystalExpression($"{value.ToOperand()}.{member}", type, isAtomic: true);
        }

        private static StiCrystalExpression Method(List<StiCrystalExpression> args, string method, Type type)
        {
            if (args.Count != 1)
                return null;

            return new StiCrystalExpression($"{args[0].ToOperand()}.{method}", type, isAtomic: true);
        }

        private static StiCrystalExpression Call(string function, Type type, params StiCrystalExpression[] args)
        {
            var code = $"{function}({string.Join(", ", args.Select(a => a.Code))})";
            return new StiCrystalExpression(code, type, isAtomic: true);
        }
        #endregion
    }
}
