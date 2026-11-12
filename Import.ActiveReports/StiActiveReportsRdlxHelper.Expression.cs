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
using System.Text;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts the expressions of RDLX to the expressions of Stimulsoft. There are two syntaxes: the Visual
    /// Basic one of RDL, a value that starts with "=" (=Fields!Name.Value &amp; " " &amp; Sum(Fields!Total.Value)),
    /// and the template one of ActiveReportsJS, a text with expressions in braces ({Name} - {Sum(Total)}).
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Methods.Expression
        /// <summary>
        /// Converts a value of a report item to the text of a component, with the expressions in braces.
        /// </summary>
        private string ConvertValue(string value, StiRdlxScope scope, bool allowBandMarker)
        {
            if (string.IsNullOrEmpty(value))
                return "";

            if (value.StartsWith("="))
            {
                var expression = ConvertExpression(value.Substring(1), scope, isTemplate: false,
                    allowBandMarker: allowBandMarker);

                return IsStringLiteral(expression) ? GetLiteralText(expression) : "{" + expression + "}";
            }

            if (isJson && value.IndexOf('{') >= 0)
                return ConvertTemplate(value, scope, allowBandMarker);

            return value;
        }

        private string ConvertValue(string value, StiRdlxScope scope)
        {
            return ConvertValue(value, scope, allowBandMarker: false);
        }

        /// <summary>
        /// Converts a value to an expression without braces: a condition, a filter, an argument.
        /// A plain text gives a string literal.
        /// </summary>
        private string ConvertToExpression(string value, StiRdlxScope scope)
        {
            if (string.IsNullOrEmpty(value))
                return "\"\"";

            if (value.StartsWith("="))
                return ConvertExpression(value.Substring(1), scope, isTemplate: false, allowBandMarker: false);

            if (isJson && value.StartsWith("{") && value.EndsWith("}") && FindCloseBrace(value, 0) == value.Length - 1)
            {
                return ConvertExpression(value.Substring(1, value.Length - 2), scope, isTemplate: true,
                    allowBandMarker: false);
            }

            double number;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                return value;

            if (string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "false", StringComparison.OrdinalIgnoreCase))
                return value.ToLowerInvariant();

            return ToLiteral(value);
        }

        /// <summary>
        /// Converts a text of ActiveReportsJS with expressions in braces.
        /// </summary>
        private string ConvertTemplate(string value, StiRdlxScope scope, bool allowBandMarker)
        {
            var builder = new StringBuilder();
            var position = 0;

            while (position < value.Length)
            {
                var open = value.IndexOf('{', position);
                var close = open < 0 ? -1 : FindCloseBrace(value, open);
                if (close < 0)
                    break;

                builder.Append(value, position, open - position);

                var inner = value.Substring(open + 1, close - open - 1);
                builder.Append('{');
                builder.Append(ConvertExpression(inner, scope, isTemplate: true, allowBandMarker: allowBandMarker));
                builder.Append('}');

                position = close + 1;
            }

            builder.Append(value, position, value.Length - position);
            return builder.ToString();
        }

        private static int FindCloseBrace(string value, int open)
        {
            var depth = 0;
            var quote = '\0';

            for (var index = open; index < value.Length; index++)
            {
                var ch = value[index];

                if (quote != '\0')
                {
                    if (ch == quote)
                        quote = '\0';

                    continue;
                }

                if (ch == '"' || ch == '\'')
                    quote = ch;

                else if (ch == '{')
                    depth++;

                else if (ch == '}' && --depth == 0)
                    return index;
            }

            return -1;
        }

        private string ConvertExpression(string expression, StiRdlxScope scope, bool isTemplate, bool allowBandMarker)
        {
            var state = new StiRdlxExpressionState
            {
                Tokens = Tokenize(expression),
                Scope = scope ?? new StiRdlxScope(),
                IsTemplate = isTemplate,
                AllowBandMarker = allowBandMarker
            };

            var builder = new StringBuilder();

            while (!state.IsEnd)
            {
                var part = ConvertSequence(state);
                AppendPart(builder, part);

                if (state.IsEnd)
                    break;

                // A comma of the outer level is kept as it is, a closing parenthesis without its pair is
                // a slip of the author of the report and is dropped.
                if (state.Current.Type == StiRdlxTokenType.Comma)
                    AppendPart(builder, state.Current.Text);

                state.Index++;
            }

            return builder.ToString().Trim();
        }

        /// <summary>
        /// Converts the tokens up to a comma or a closing parenthesis of the outer level.
        /// </summary>
        private string ConvertSequence(StiRdlxExpressionState state)
        {
            var builder = new StringBuilder();

            while (!state.IsEnd)
            {
                var token = state.Current;
                if (token.Type == StiRdlxTokenType.CloseParen || token.Type == StiRdlxTokenType.Comma)
                    break;

                state.Index++;

                switch (token.Type)
                {
                    case StiRdlxTokenType.OpenParen:
                        AppendPart(builder, "(" + string.Join(", ", ConvertArguments(state)) + ")");
                        break;

                    case StiRdlxTokenType.String:
                        AppendPart(builder, ToLiteral(token.Text));
                        break;

                    case StiRdlxTokenType.Date:
                        AppendPart(builder, "Convert.ToDateTime(" + ToLiteral(token.Text) + ")");
                        break;

                    case StiRdlxTokenType.Operator:
                        AppendPart(builder, ConvertOperator(token.Text));
                        break;

                    case StiRdlxTokenType.Identifier:
                        AppendPart(builder, ConvertIdentifier(state, token.Text));
                        break;

                    default:
                        AppendPart(builder, token.Text);
                        break;
                }
            }

            return builder.ToString();
        }

        /// <summary>
        /// Converts the arguments after an opening parenthesis and passes the closing one.
        /// </summary>
        private List<string> ConvertArguments(StiRdlxExpressionState state)
        {
            var arguments = new List<string>();

            if (state.IsType(0, StiRdlxTokenType.CloseParen))
            {
                state.Index++;
                return arguments;
            }

            while (!state.IsEnd)
            {
                arguments.Add(ConvertSequence(state).Trim());

                if (state.IsEnd)
                    break;

                var token = state.Current;
                state.Index++;

                if (token.Type == StiRdlxTokenType.CloseParen)
                    break;
            }

            return arguments;
        }

        private static void AppendPart(StringBuilder builder, string part)
        {
            if (string.IsNullOrEmpty(part))
                return;

            var isMember = part == "." || part.StartsWith(".");
            var afterMember = builder.Length > 0 && (builder[builder.Length - 1] == '.' ||
                builder[builder.Length - 1] == '(' || builder[builder.Length - 1] == '!');

            if (builder.Length > 0 && !isMember && !afterMember)
                builder.Append(' ');

            builder.Append(part);
        }

        private static string ConvertOperator(string text)
        {
            switch (text)
            {
                case "=":
                    return "==";

                case "<>":
                    return "!=";

                case "&":
                    return "+";

                case "\\":
                    return "/";

                default:
                    return text;
            }
        }

        private string ConvertIdentifier(StiRdlxExpressionState state, string name)
        {
            var lower = name.ToLowerInvariant();

            switch (lower)
            {
                case "and":
                case "andalso":
                    return "&&";

                case "or":
                case "orelse":
                    return "||";

                case "not":
                    return "!";

                case "mod":
                    return "%";

                case "xor":
                    return "^";

                case "is":
                    return "==";

                case "isnot":
                    return "!=";

                case "true":
                case "false":
                    return lower;

                case "nothing":
                case "null":
                    return "null";

                case "fields":
                case "parameters":
                case "globals":
                case "reportitems":
                case "user":
                    string member;
                    if (TryReadMember(state, out member))
                        return ConvertMember(state, lower, member);

                    break;
            }

            // DateTime.Parse(value): a column of dates is a date already.
            if (lower == "datetime" && state.IsOperator(".") && state.IsType(1, StiRdlxTokenType.Identifier) &&
                state.IsType(2, StiRdlxTokenType.OpenParen) &&
                string.Equals(state.Tokens[state.Index + 1].Text, "Parse", StringComparison.OrdinalIgnoreCase))
            {
                state.Index += 3;

                var parsed = ConvertArguments(state);
                var value = parsed.Count > 0 ? parsed[0] : "";

                return ConvertToDate(value);
            }

            if (state.IsTemplate && name.StartsWith("@"))
                return GetParameterReference(name.Substring(1));

            var isMemberName = state.Index >= 2 && state.Tokens[state.Index - 2].Type == StiRdlxTokenType.Operator &&
                state.Tokens[state.Index - 2].Text == ".";

            if (state.IsType(0, StiRdlxTokenType.OpenParen))
            {
                state.Index++;
                var arguments = ConvertArguments(state);

                // A method of a value is kept as it is, only the functions are converted.
                return isMemberName
                    ? GetMemberName(name) + "(" + string.Join(", ", arguments) + ")"
                    : ConvertFunction(state, name, arguments);
            }

            if (isMemberName)
                return GetMemberName(name);

            var global = ConvertGlobal(lower);
            if (global != null)
                return global;

            if (state.IsTemplate)
            {
                if (FindFieldDataSet(name, state.Scope) != null)
                    return GetFieldReference(name, state.Scope);

                if (parameters.ContainsKey(name))
                    return parameters[name];
            }

            return name;
        }

        /// <summary>
        /// The name of a method or a property of a value in the case of .NET: the expressions of RDLX do
        /// not depend on the case of letters.
        /// </summary>
        private static string GetMemberName(string name)
        {
            var known = new[]
            {
                "ToString", "ToUpper", "ToLower", "Trim", "TrimStart", "TrimEnd", "Substring", "Replace", "Contains",
                "StartsWith", "EndsWith", "IndexOf", "Length", "PadLeft", "PadRight", "Split", "Year", "Month", "Day",
                "Hour", "Minute", "Second", "Date", "DayOfWeek", "AddDays", "AddMonths", "AddYears", "AddHours",
                "ToShortDateString", "ToLongDateString", "Count", "Equals"
            };

            foreach (var item in known)
            {
                if (string.Equals(item, name, StringComparison.OrdinalIgnoreCase))
                    return item;
            }

            return name;
        }

        /// <summary>
        /// Reads the name after a collection: Fields!Name.Value, Fields.Name, Fields("Name").Value,
        /// Fields.Item("Name").Value.
        /// </summary>
        private bool TryReadMember(StiRdlxExpressionState state, out string member)
        {
            member = null;

            var isBang = state.IsOperator("!");
            var isDot = state.IsOperator(".");

            if ((isBang || isDot) && state.IsType(1, StiRdlxTokenType.Identifier))
            {
                member = state.Tokens[state.Index + 1].Text;
                state.Index += 2;

                if (isDot && string.Equals(member, "Item", StringComparison.OrdinalIgnoreCase) &&
                    state.IsType(0, StiRdlxTokenType.OpenParen) && state.IsType(1, StiRdlxTokenType.String))
                {
                    member = state.Tokens[state.Index + 1].Text;
                    state.Index += 3;
                }
            }
            else if (state.IsType(0, StiRdlxTokenType.OpenParen) && state.IsType(1, StiRdlxTokenType.String) &&
                state.IsType(2, StiRdlxTokenType.CloseParen))
            {
                member = state.Tokens[state.Index + 1].Text;
                state.Index += 3;
            }

            if (member == null)
                return false;

            // A name with special characters is written in brackets and quotes: ['Encounter#'].
            if (member.StartsWith("['") && member.EndsWith("']") && member.Length > 4)
                member = member.Substring(2, member.Length - 4);

            // The property of the item of the collection is dropped: the item is the value itself.
            if (state.IsOperator(".") && state.IsType(1, StiRdlxTokenType.Identifier))
            {
                var property = state.Tokens[state.Index + 1].Text.ToLowerInvariant();
                if (property == "value" || property == "label")
                    state.Index += 2;
            }

            return true;
        }

        private string ConvertMember(StiRdlxExpressionState state, string collection, string member)
        {
            switch (collection)
            {
                case "fields":
                    return GetFieldReference(member, state.Scope);

                case "parameters":
                    return GetParameterReference(member);

                case "globals":
                    return ConvertGlobal(member.ToLowerInvariant()) ?? member;

                case "reportitems":
                    return member + ".Text";

                default:
                    return "\"\"";
            }
        }

        private static string ConvertGlobal(string lower)
        {
            switch (lower)
            {
                case "pagenumber":
                case "pagenumberinsection":
                case "cumulativepagenumber":
                    return "PageNumber";

                case "totalpages":
                case "totalpagesinsection":
                case "cumulativetotalpages":
                    return "TotalPageCount";

                case "executiontime":
                case "now":
                    return "Time";

                case "today":
                    return "Today";

                case "reportname":
                    return "ReportName";

                default:
                    return null;
            }
        }

        private string GetParameterReference(string name)
        {
            string variable;
            if (parameters.TryGetValue(name, out variable))
                return variable;

            // The report reads a parameter it does not declare, the application passes it at run time.
            Warn($"The parameter '{name}' is not declared in the report, a variable is added for it.");
            return AddVariable(name, typeof(string), name, null, isHidden: true);
        }

        private string GetFieldReference(string field, StiRdlxScope scope)
        {
            var dataSet = FindFieldDataSet(field, scope);
            if (dataSet == null)
            {
                Warn($"The field '{field}' is not found in the data sets of the report.");

                return scope != null && scope.DataSet != null
                    ? scope.DataSet.SourceName + "." + CreateColumnName(field)
                    : CreateColumnName(field);
            }

            return dataSet.SourceName + "." + dataSet.Columns[field];
        }

        /// <summary>
        /// Finds the data set of a field: the one of the scope or, when it has no such field, the first data
        /// set that has it.
        /// </summary>
        private StiRdlxDataSet FindFieldDataSet(string field, StiRdlxScope scope)
        {
            if (scope != null && scope.DataSet != null && scope.DataSet.Columns.ContainsKey(field))
                return scope.DataSet;

            foreach (var dataSet in dataSetList)
            {
                if (dataSet.Columns.ContainsKey(field))
                    return dataSet;
            }

            return null;
        }

        private string ConvertFunction(StiRdlxExpressionState state, string name, List<string> arguments)
        {
            var lower = name.ToLowerInvariant();
            var first = arguments.Count > 0 ? arguments[0] : "";

            switch (lower)
            {
                case "sum":
                case "avg":
                case "min":
                case "max":
                case "median":
                case "mode":
                case "last":
                case "countdistinct":
                    if (arguments.Count >= 2 && !IsStringLiteral(arguments[arguments.Count - 1]))
                        return ToTitle(lower) + "(" + string.Join(", ", arguments) + ")";

                    return ConvertTotal(state, ToTitle(lower), arguments);

                case "count":
                case "countrows":
                    return ConvertTotal(state, "Count", arguments);

                case "first":
                    return first.Length == 0 ? "\"\"" : first;

                case "runningvalue":
                    var function = arguments.Count > 1 ? ToTitle(arguments[1].ToLowerInvariant()) : "Sum";
                    return function + "Running(" + first + ")";

                case "rownumber":
                    return "Line";

                case "iif":
                    // ActiveReports allows to leave out the value of the false branch.
                    if (arguments.Count == 2)
                        arguments.Add("\"\"");

                    return "IIF(" + string.Join(", ", arguments) + ")";

                case "aggregateif":
                    var aggregate = ConvertAggregateIf(state, arguments);
                    if (aggregate != null)
                        return aggregate;

                    break;

                case "switch":
                    return "Switch(" + string.Join(", ", arguments) + ")";

                case "choose":
                    return "Choose(" + string.Join(", ", arguments) + ")";

                case "format":
                    if (arguments.Count == 2 && IsStringLiteral(arguments[1]))
                        return "Format(\"{0:" + GetLiteralText(arguments[1]) + "}\", " + first + ")";

                    return "ToString(" + first + ")";

                case "formatcurrency":
                    return "Format(\"{0:C" + GetDecimals(arguments) + "}\", " + first + ")";

                case "formatnumber":
                    return "Format(\"{0:N" + GetDecimals(arguments) + "}\", " + first + ")";

                case "formatpercent":
                    return "Format(\"{0:P" + GetDecimals(arguments) + "}\", " + first + ")";

                case "formatdatetime":
                    return "Format(\"{0:d}\", " + first + ")";

                case "cstr":
                case "tostring":
                case "str":
                    return "ToString(" + first + ")";

                case "cint":
                case "clng":
                case "cshort":
                case "int":
                case "fix":
                case "toint16":
                case "toint32":
                    return "Convert.ToInt32(" + first + ")";

                case "toint64":
                    return "Convert.ToInt64(" + first + ")";

                case "cdbl":
                case "csng":
                case "val":
                case "todouble":
                case "tosingle":
                    return "Convert.ToDouble(" + first + ")";

                case "cdec":
                case "todecimal":
                    return "Convert.ToDecimal(" + first + ")";

                case "cdate":
                case "todatetime":
                    return ConvertToDate(first);

                case "monthname":
                    if (first.StartsWith("Month(") && first.EndsWith(")") && first.IndexOf(')') == first.Length - 1)
                        return "MonthName(" + first.Substring(6, first.Length - 7) + ")";

                    return "MonthName(DateSerial(2000, Convert.ToInt32(" + first + "), 1))";

                case "cbool":
                case "toboolean":
                    return "Convert.ToBoolean(" + first + ")";

                case "datepart":
                    var part = ConvertDatePart(arguments);
                    if (part != null)
                        return part;

                    break;

                case "len":
                    return "Length(" + first + ")";

                case "ucase":
                case "toupper":
                case "upper":
                    return "ToUpperCase(" + first + ")";

                case "lcase":
                case "tolower":
                case "lower":
                    return "ToLowerCase(" + first + ")";

                case "trim":
                case "ltrim":
                case "rtrim":
                    return "Trim(" + first + ")";

                case "left":
                case "right":
                case "replace":
                case "abs":
                case "round":
                case "ceiling":
                case "floor":
                case "sqrt":
                case "sign":
                case "truncate":
                case "year":
                case "month":
                case "day":
                case "dateserial":
                    return ToTitle(lower) + "(" + string.Join(", ", arguments) + ")";

                case "mid":
                    if (arguments.Count >= 3)
                        return "Mid(" + first + ", (" + arguments[1] + ") - 1, " + arguments[2] + ")";

                    if (arguments.Count == 2)
                        return "ToString(" + first + ").Substring((" + arguments[1] + ") - 1)";

                    break;

                case "instr":
                    if (arguments.Count == 2)
                        return "(ToString(" + first + ").IndexOf(" + arguments[1] + ") + 1)";

                    break;

                case "now":
                    return "Time";

                case "today":
                    return "Today";

                case "datediff":
                    return ConvertDateDiff(arguments);

                case "lookup":
                    var lookup = ConvertLookup(arguments);
                    if (lookup != null)
                        return lookup;

                    break;

                case "isnothing":
                case "isdbnull":
                case "isnull":
                    return "(" + first + " == null)";
            }

            return name + "(" + string.Join(", ", arguments) + ")";
        }

        /// <summary>
        /// Converts a total function. The last argument of RDL is the scope: a data set, a data region or
        /// a group. A total of a data set that is placed out of the data regions gets the marker of the data
        /// band, it becomes the name of the band when all bands are made.
        /// </summary>
        private string ConvertTotal(StiRdlxExpressionState state, string function, List<string> arguments)
        {
            StiRdlxDataSet dataSet = null;
            var hasScope = arguments.Count >= 2 && IsStringLiteral(arguments[arguments.Count - 1]);

            if (hasScope)
            {
                var scopeName = GetLiteralText(arguments[arguments.Count - 1]);
                dataSet = FindDataSet(scopeName);
                arguments = arguments.GetRange(0, arguments.Count - 1);

                // Stimulsoft has no total of all the rows of a band that could stand on a band of a group.
                var isWholeRegion = regionNames.Contains(scopeName) ||
                    dataSet != null && dataSet == state.Scope.DataSet;

                if (state.Scope.IsRegion && isWholeRegion)
                {
                    Warn($"A total with the scope '{scopeName}' stands inside a data region. On a band of a group " +
                        "Stimulsoft counts it for the current group, not for all the rows of the region.");
                }
            }

            var expression = arguments.Count > 0 ? arguments[0] : "";

            if (dataSet == null)
                dataSet = FindDataSetOfExpression(expression) ?? state.Scope.DataSet;

            var marker = "";

            if (state.AllowBandMarker && dataSet != null && (!state.Scope.IsRegion || state.Scope.DataSet != dataSet))
                marker = BandMarkerStart + dataSet.Name + BandMarkerEnd;

            if (function == "Count")
                return "Count(" + marker + ")";

            return function + "(" + marker + expression + ")";
        }

        /// <summary>
        /// Finds the data set whose column is read by a converted expression.
        /// </summary>
        private StiRdlxDataSet FindDataSetOfExpression(string expression)
        {
            foreach (var dataSet in dataSetList)
            {
                var position = expression.IndexOf(dataSet.SourceName + ".", StringComparison.Ordinal);
                if (position == 0 || position > 0 && !IsNameChar(expression[position - 1]))
                    return dataSet;
            }

            return null;
        }

        /// <summary>
        /// Converts AggregateIf(condition, "Sum", value, scope): a total of the rows that pass the condition.
        /// It becomes SumIf, CountIf and the like of Stimulsoft. Returns null for an unknown function.
        /// </summary>
        private string ConvertAggregateIf(StiRdlxExpressionState state, List<string> arguments)
        {
            if (arguments.Count < 3)
                return null;

            var function = IsStringLiteral(arguments[1]) ? GetLiteralText(arguments[1]) : arguments[1];
            var name = ToTitle(function.ToLowerInvariant());

            if (name != "Sum" && name != "Count" && name != "Avg" && name != "Min" && name != "Max")
                return null;

            // The value and the scope go the way of a plain total, so the band of the total is found the same.
            var total = ConvertTotal(state, name, arguments.GetRange(2, arguments.Count - 2));
            var inner = total.Substring(name.Length + 1, total.Length - name.Length - 2);

            return name == "Count"
                ? "CountIf(" + inner + arguments[0] + ")"
                : name + "If(" + inner + ", " + arguments[0] + ")";
        }

        /// <summary>
        /// Converts DateDiff("m", start, end) of Visual Basic. DateDiff of Stimulsoft takes two dates and
        /// returns their difference as a time span.
        /// </summary>
        private static string ConvertDateDiff(List<string> arguments)
        {
            if (arguments.Count != 3 || !IsStringLiteral(arguments[0]))
                return "DateDiff(" + string.Join(", ", arguments) + ")";

            var start = arguments[1];
            var end = arguments[2];
            var span = "DateDiff(" + end + ", " + start + ")";

            switch (GetLiteralText(arguments[0]).ToLowerInvariant())
            {
                case "yyyy":
                case "y":
                    return "(Year(" + end + ") - Year(" + start + "))";

                case "m":
                    return "((Year(" + end + ") - Year(" + start + ")) * 12 + Month(" + end + ") - Month(" + start +
                        "))";

                case "w":
                case "ww":
                    return "(" + span + ".Days / 7)";

                case "h":
                    return "Convert.ToInt32(" + span + ".TotalHours)";

                case "n":
                    return "Convert.ToInt32(" + span + ".TotalMinutes)";

                case "s":
                    return "Convert.ToInt32(" + span + ".TotalSeconds)";

                default:
                    return span + ".Days";
            }
        }

        /// <summary>
        /// DatePart("m", date) and the like: the part of a date by its code.
        /// </summary>
        private static string ConvertDatePart(List<string> arguments)
        {
            if (arguments.Count < 2 || !IsStringLiteral(arguments[0]))
                return null;

            var date = arguments[1];

            switch (GetLiteralText(arguments[0]).ToLowerInvariant())
            {
                case "yyyy":
                case "y":
                    return "Year(" + date + ")";

                case "m":
                    return "Month(" + date + ")";

                case "d":
                    return "Day(" + date + ")";

                case "q":
                    return "((Month(" + date + ") + 2) / 3)";

                case "h":
                    return "(" + date + ").Hour";

                case "n":
                    return "(" + date + ").Minute";

                default:
                    return null;
            }
        }

        /// <summary>
        /// Makes a date of a converted expression. A column of dates is a date already; any other value is
        /// parsed as a text, the result of Convert.ToDateTime has no type for the parser of Stimulsoft.
        /// </summary>
        private string ConvertToDate(string expression)
        {
            return IsDateColumn(expression) ? expression : "DateTime.Parse(ToString(" + expression + "))";
        }

        /// <summary>
        /// The converted expression is a column of dates of a data source.
        /// </summary>
        private bool IsDateColumn(string expression)
        {
            var dot = expression.IndexOf('.');
            if (dot <= 0)
                return false;

            foreach (var dataSet in dataSetList)
            {
                if (dataSet.Source == null || dataSet.SourceName != expression.Substring(0, dot))
                    continue;

                var column = dataSet.Source.Columns[expression.Substring(dot + 1)];
                return column != null && column.Type == typeof(DateTime);
            }

            return false;
        }

        private static string GetDecimals(List<string> arguments)
        {
            int decimals;
            if (arguments.Count > 1 && int.TryParse(arguments[1], out decimals))
                return decimals.ToString(CultureInfo.InvariantCulture);

            return "2";
        }

        private static string ToTitle(string lower)
        {
            switch (lower)
            {
                case "countdistinct":
                    return "CountDistinct";

                case "monthname":
                    return "MonthName";

                case "dateserial":
                    return "DateSerial";

                case "datediff":
                    return "DateDiff";

                default:
                    return lower.Length == 0 ? lower : char.ToUpperInvariant(lower[0]) + lower.Substring(1);
            }
        }

        private static bool IsStringLiteral(string expression)
        {
            if (expression.Length < 2 || expression[0] != '"' || expression[expression.Length - 1] != '"')
                return false;

            for (var index = 1; index < expression.Length - 1; index++)
            {
                if (expression[index] == '\\')
                    index++;

                else if (expression[index] == '"')
                    return false;
            }

            return true;
        }

        private static string GetLiteralText(string literal)
        {
            var builder = new StringBuilder();

            for (var index = 1; index < literal.Length - 1; index++)
            {
                if (literal[index] == '\\' && index + 1 < literal.Length - 1)
                    index++;

                builder.Append(literal[index]);
            }

            return builder.ToString();
        }

        private static string ToLiteral(string text)
        {
            return "\"" + text.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", "\\r").Replace("\n", "\\n") +
                "\"";
        }
        #endregion

        #region Methods.Tokens
        private static List<StiRdlxToken> Tokenize(string expression)
        {
            var tokens = new List<StiRdlxToken>();
            var position = 0;

            while (position < expression.Length)
            {
                var ch = expression[position];

                // The sign of an interpolated string, $"text", is skipped: the string is read as a plain one.
                var isStringSign = ch == '$' && position + 1 < expression.Length && expression[position + 1] == '"';

                if (char.IsWhiteSpace(ch) || isStringSign)
                {
                    position++;
                }
                else if (char.IsLetter(ch) || ch == '_' || ch == '@' || ch == '$')
                {
                    var start = position++;

                    while (position < expression.Length && IsNameChar(expression[position]))
                    {
                        position++;
                    }

                    var name = expression.Substring(start, position - start);
                    tokens.Add(new StiRdlxToken(StiRdlxTokenType.Identifier, name));
                }
                else if (char.IsDigit(ch) || ch == '.' && position + 1 < expression.Length &&
                    char.IsDigit(expression[position + 1]))
                {
                    var start = position++;

                    while (position < expression.Length &&
                        (char.IsDigit(expression[position]) || expression[position] == '.'))
                    {
                        position++;
                    }

                    var number = expression.Substring(start, position - start);
                    tokens.Add(new StiRdlxToken(StiRdlxTokenType.Number, number));
                }
                else if (ch == '"' || ch == '\'')
                {
                    tokens.Add(new StiRdlxToken(StiRdlxTokenType.String, ReadString(expression, ref position)));
                }
                else if (ch == '#' && expression.IndexOf('#', position + 1) > position)
                {
                    var end = expression.IndexOf('#', position + 1);
                    var date = expression.Substring(position + 1, end - position - 1);
                    tokens.Add(new StiRdlxToken(StiRdlxTokenType.Date, date));
                    position = end + 1;
                }
                else if (ch == '(')
                {
                    tokens.Add(new StiRdlxToken(StiRdlxTokenType.OpenParen, "("));
                    position++;
                }
                else if (ch == ')')
                {
                    tokens.Add(new StiRdlxToken(StiRdlxTokenType.CloseParen, ")"));
                    position++;
                }
                else if (ch == ',')
                {
                    tokens.Add(new StiRdlxToken(StiRdlxTokenType.Comma, ","));
                    position++;
                }
                else
                {
                    var length = IsDoubleOperator(expression, position) ? 2 : 1;
                    tokens.Add(new StiRdlxToken(StiRdlxTokenType.Operator, expression.Substring(position, length)));
                    position += length;
                }
            }

            return tokens;
        }

        private static bool IsNameChar(char ch)
        {
            return char.IsLetterOrDigit(ch) || ch == '_';
        }

        private static bool IsDoubleOperator(string expression, int position)
        {
            if (position + 1 >= expression.Length)
                return false;

            switch (expression.Substring(position, 2))
            {
                case "<>":
                case "<=":
                case ">=":
                case "==":
                case "!=":
                case "&&":
                case "||":
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>
        /// Reads a string literal: a doubled quote inside is the quote itself.
        /// </summary>
        private static string ReadString(string expression, ref int position)
        {
            var quote = expression[position++];
            var builder = new StringBuilder();

            while (position < expression.Length)
            {
                var ch = expression[position++];

                if (ch != quote)
                {
                    builder.Append(ch);
                    continue;
                }

                if (position < expression.Length && expression[position] == quote)
                {
                    builder.Append(quote);
                    position++;
                    continue;
                }

                break;
            }

            return builder.ToString();
        }
        #endregion
    }
}
