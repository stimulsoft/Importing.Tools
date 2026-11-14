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

using Stimulsoft.Base.Helpers;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Import.ReportingServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts the expressions and the code of Reporting Services. The text is split into tokens of
    /// Visual Basic and only the known tokens are rewritten: the collections Fields, Parameters,
    /// Globals, User, Variables and ReportItems, the aggregate functions with their scope and a few
    /// functions that differ in Stimulsoft. A string, a comment and the member of an object are left
    /// as they are, and the arguments of a function are found by the matching parenthesis.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts
        private static readonly string[] Collections = { "Fields", "Parameters", "Globals", "User", "Variables",
            "ReportItems" };

        private static readonly string[] AggregateFunctions = { "Sum", "Avg", "Count", "CountDistinct", "CountRows",
            "Max", "Min", "First", "Last" };

        private static readonly string[] UnsupportedFunctions = { "Lookup", "LookupSet", "MultiLookup", "InScope",
            "Level", "Aggregate" };

        //the namespace of Visual Basic, a named constant can be written with it
        private const string VbNamespace = "Microsoft.VisualBasic.";

        //the named constants of Visual Basic the parser does not know, with their values
        private static readonly Dictionary<string, string> NamedValues =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "DateInterval.Year", "\"yyyy\"" },
            { "DateInterval.Quarter", "\"q\"" },
            { "DateInterval.Month", "\"m\"" },
            { "DateInterval.DayOfYear", "\"y\"" },
            { "DateInterval.Day", "\"d\"" },
            { "DateInterval.WeekOfYear", "\"ww\"" },
            { "DateInterval.Weekday", "\"w\"" },
            { "DateInterval.Hour", "\"h\"" },
            { "DateInterval.Minute", "\"n\"" },
            { "DateInterval.Second", "\"s\"" },
            { "DateFormat.GeneralDate", "0" },
            { "DateFormat.LongDate", "1" },
            { "DateFormat.ShortDate", "2" },
            { "DateFormat.LongTime", "3" },
            { "DateFormat.ShortTime", "4" },
            { "Environment.NewLine", "vbCrLf" },
            { "System.Environment.NewLine", "vbCrLf" }
        };
        #endregion

        #region Properties.Expressions
        private string NothingKeyword
        {
            get
            {
                return ConvertSyntaxToCSharp ? "null" : "Nothing";
            }
        }

        private string TrueKeyword
        {
            get
            {
                return ConvertSyntaxToCSharp ? "true" : "True";
            }
        }

        private string FalseKeyword
        {
            get
            {
                return ConvertSyntaxToCSharp ? "false" : "False";
            }
        }
        #endregion

        #region Expressions
        public string ConvertExpression(string baseExpression, string dataset)
        {
            if (baseExpression == null)
                return string.Empty;

            if (!IsExpression(baseExpression))
                return baseExpression;

            var tokens = StiVbTokenizer.Tokenize(baseExpression.Trim().Substring(1));
            var expression = ConvertTokens(tokens, 0, tokens.Count, dataset, isCode: false);

            return "{" + expression.Trim() + "}";
        }

        /// <summary>
        /// Converts the code of the report. The lines and the comments stay, only the references to
        /// the collections and the functions hidden by the methods of the report are changed.
        /// </summary>
        private string ConvertCode(string code)
        {
            var tokens = StiVbTokenizer.Tokenize(code);

            return ConvertTokens(tokens, 0, tokens.Count, null, isCode: true);
        }

        private string ConvertTokens(List<StiVbToken> tokens, int start, int end, string dataset, bool isCode)
        {
            var result = new StringBuilder();
            var index = start;

            while (index < end)
            {
                index = ConvertToken(tokens, index, end, dataset, isCode, result);
            }

            return result.ToString();
        }

        /// <summary>
        /// Converts the token at the index, with the tokens after it that belong to it, and returns
        /// the index of the next token to convert.
        /// </summary>
        private int ConvertToken(List<StiVbToken> tokens, int index, int end, string dataset, bool isCode,
            StringBuilder result)
        {
            var token = tokens[index];

            switch (token.Type)
            {
                case StiVbTokenType.NewLine:
                    //an expression is written in one line of the text of the component
                    if (isCode)
                        result.Append(token.Text);

                    else
                        AppendSpace(result);

                    return index + 1;

                case StiVbTokenType.Comment:
                    if (isCode)
                        result.Append(token.Text);

                    return index + 1;

                case StiVbTokenType.String:
                    result.Append(ConvertSyntaxToCSharp ? ToCSharpString(token.StringValue) : token.Text);
                    return index + 1;

                case StiVbTokenType.Operator:
                    ConvertOperator(token, result);
                    return index + 1;

                case StiVbTokenType.Identifier:
                    return ConvertIdentifier(tokens, index, end, dataset, isCode, result);

                default:
                    result.Append(token.Text);
                    return index + 1;
            }
        }

        /// <summary>
        /// Visual Basic reads a &amp; right after a name or a number as the type character of Long,
        /// so Fields!Room.Value&amp;", " does not compile, though Reporting Services takes it. A space
        /// is put before such a concatenation.
        /// </summary>
        private void ConvertOperator(StiVbToken token, StringBuilder result)
        {
            if (token.Text != "&")
            {
                result.Append(token.Text);
                return;
            }

            if (ConvertSyntaxToCSharp)
            {
                result.Append("+");
                return;
            }

            if (result.Length > 0)
            {
                var previous = result[result.Length - 1];
                if (char.IsLetterOrDigit(previous) || previous == '_' || previous == ')')
                    result.Append(' ');
            }

            result.Append(token.Text);
        }

        private int ConvertIdentifier(List<StiVbToken> tokens, int index, int end, string dataset, bool isCode,
            StringBuilder result)
        {
            var token = tokens[index];
            var previous = GetPreviousSignificant(tokens, index);
            var nextIndex = GetNextSignificant(tokens, index + 1, end);
            var next = nextIndex < end ? tokens[nextIndex] : null;

            //a member of an object is not a function or a collection of the report
            if (previous != null && previous.IsOperator("."))
            {
                result.Append(token.Text);
                return index + 1;
            }

            if (IsCollection(tokens, index, end))
                return ConvertCollection(tokens, index, end, dataset, result);

            //the code of the report reaches the collections through the Report object
            var isReportMember = token.IsName("Report") && next != null && next.IsOperator(".");
            if (isReportMember && IsCollection(tokens, nextIndex + 1, end))
                return ConvertCollection(tokens, nextIndex + 1, end, dataset, result);

            if (!isCode)
            {
                var isCall = next != null && next.IsOperator("(");

                //the functions of the code are the methods of the report
                if (token.IsName("Code") && next != null && next.IsOperator("."))
                {
                    usesReportCode = true;
                    return nextIndex + 1;
                }

                //Me is the current report item, so Me.Value in an expression of a text box, a color as a
                //rule, is the value of the text box itself
                if (token.IsName("Me") && next != null && next.IsOperator(".") && currentTextboxName != null)
                {
                    var propertyIndex = GetNextSignificant(tokens, nextIndex + 1, end);
                    if (propertyIndex < end && tokens[propertyIndex].IsName("Value"))
                    {
                        result.Append(ConvertReportItem(currentTextboxName, dataset));
                        return propertyIndex + 1;
                    }
                }

                if (token.IsName("DataSets"))
                {
                    var commandEnd = ConvertDataSetCommandText(tokens, index, end, result);
                    if (commandEnd != -1)
                        return commandEnd;
                }

                if (isCall && (token.IsName("Round") || token.IsName("Ceiling") || token.IsName("Floor")))
                    return ConvertRound(tokens, index, nextIndex, end, dataset, result);

                if (isCall && token.IsName("Join"))
                    return ConvertJoin(tokens, index, nextIndex, end, dataset, result);

                if (next != null && next.IsOperator(".") && !ConvertSyntaxToCSharp)
                {
                    var valueEnd = ConvertNamedValue(tokens, index, end, result);
                    if (valueEnd != -1)
                        return valueEnd;
                }

                if (isCall && FindName(AggregateFunctions, token.Name) != null)
                    return ConvertAggregate(tokens, index, nextIndex, end, dataset, result);

                if (isCall && FindName(UnsupportedFunctions, token.Name) != null)
                    return SkipUnsupportedFunction(tokens, index, nextIndex, end, result);

                if (isCall && token.IsName("RunningValue"))
                    return ConvertRunningValue(tokens, index, nextIndex, end, dataset, result);

                if (isCall && token.IsName("Previous"))
                    return ConvertPrevious(tokens, index, nextIndex, end, dataset, result);

                if (isCall && token.IsName("RowNumber"))
                    return ConvertRowNumber(tokens, index, nextIndex, end, result);

                if (token.IsName("Today"))
                    return AppendInsteadOfCall(tokens, index, isCall ? nextIndex : -1, end, "Today", result);

                if (token.IsName("Now"))
                    return AppendInsteadOfCall(tokens, index, isCall ? nextIndex : -1, end, "DateTime.Now", result);

                if (token.IsName("IIf"))
                {
                    result.Append("IIF");
                    return index + 1;
                }

                if (token.Text == "_" && (next == null || IsLineContinuation(tokens, index, end)))
                    return index + 1;

                if (ConvertSyntaxToCSharp && token.IsName("vbCrLf"))
                {
                    result.Append("\"\\r\\n\"");
                    return index + 1;
                }
            }

            if (ConvertSyntaxToCSharp && token.IsName("Nothing"))
            {
                result.Append("null");
                return index + 1;
            }

            if (!ConvertSyntaxToCSharp)
            {
                result.Append(QualifyHiddenFunction(token, previous, next));
                return index + 1;
            }

            result.Append(token.Text);
            return index + 1;
        }

        /// <summary>
        /// The code of the report is a member of the class of the report, so a method of the report
        /// hides the function of Visual Basic with the same name. StiReport.Convert hides
        /// System.Convert, and StiReport.Format(format, arg) takes the arguments of Format of Visual
        /// Basic in the reversed order and silently returns the value unformatted. MonthName of the
        /// functions of Stimulsoft takes a date and a flag of the localized name, MonthName of Visual
        /// Basic takes the number of the month and a flag of the abbreviated name. DateDiff of Stimulsoft
        /// takes two dates, DateDiff of Visual Basic takes the interval first. DateAdd of Visual Basic is
        /// called by the full name too, so the parser knows it as well.
        /// Such calls are given the full name, a declaration of an own function is left as it is.
        /// </summary>
        private static string QualifyHiddenFunction(StiVbToken token, StiVbToken previous, StiVbToken next)
        {
            if (token.IsName("Convert") && next != null && next.IsOperator("."))
                return "System.Convert";

            var isDeclaration = previous != null && (previous.IsName("Function") || previous.IsName("Sub"));
            var isCall = next != null && next.IsOperator("(") && !isDeclaration;

            if (token.IsName("Format") && isCall)
                return "Strings.Format";

            var isMember = previous != null && previous.IsOperator(".");

            if (token.IsName("MonthName") && isCall && !isMember)
                return "DateAndTime.MonthName";

            if (token.IsName("DateDiff") && isCall && !isMember)
                return "DateAndTime.DateDiff";

            if (token.IsName("DateAdd") && isCall && !isMember)
                return "DateAndTime.DateAdd";

            return token.Text;
        }
        #endregion

        #region Expressions.Collections
        /// <summary>
        /// Checks that a collection of the report starts at the index: Fields!Name, Parameters!Name
        /// and so on, the name of the collection is in any case.
        /// </summary>
        private static bool IsCollection(List<StiVbToken> tokens, int index, int end)
        {
            return index + 2 < end &&
                tokens[index].Type == StiVbTokenType.Identifier &&
                FindName(Collections, tokens[index].Name) != null &&
                tokens[index + 1].IsOperator("!") &&
                tokens[index + 2].Type == StiVbTokenType.Identifier;
        }

        private int ConvertCollection(List<StiVbToken> tokens, int index, int end, string dataset,
            StringBuilder result)
        {
            var collection = FindName(Collections, tokens[index].Name);
            var name = tokens[index + 2].Name;
            var next = index + 3;
            var property = GetProperty(tokens, next, end);

            switch (collection)
            {
                case "Fields":
                    return ConvertField(name, property, next, dataset, result);

                case "Parameters":
                    return ConvertParameter(tokens, name, property, next, end, result);

                case "Globals":
                    return ConvertGlobal(name, property, next, result);

                case "User":
                    result.Append(ConvertUser(name));
                    return next;

                case "Variables":
                    result.Append(ConvertVariable(name));
                    return string.Equals(property, "Value", StringComparison.OrdinalIgnoreCase) ? next + 2 : next;

                default:
                    result.Append(ConvertReportItem(name, dataset));
                    return string.Equals(property, "Value", StringComparison.OrdinalIgnoreCase) ? next + 2 : next;
            }
        }

        /// <summary>
        /// Converts a field. Reporting Services takes a field that is absent from the data set as
        /// Nothing and only warns about it, so such a field turns into Nothing and is named in the
        /// list of errors. IsMissing tells whether the field is absent.
        /// </summary>
        private int ConvertField(string name, string property, int next, string dataset, StringBuilder result)
        {
            var reference = FindFieldReference(name, dataset);

            if (string.Equals(property, "IsMissing", StringComparison.OrdinalIgnoreCase))
            {
                result.Append(reference == null ? TrueKeyword : FalseKeyword);
                return next + 2;
            }

            if (reference == null)
            {
                //a field is an Object, a bare Nothing fits every overload of Len or Convert.ToString and
                //the compiler cannot choose one
                ThrowError(null, null, $"Field not found: {name}");
                result.Append(ConvertSyntaxToCSharp ? NothingKeyword : $"CObj({NothingKeyword})");
            }
            else
            {
                result.Append(reference);
            }

            if (property == null)
                return next;

            if (string.Equals(property, "Value", StringComparison.OrdinalIgnoreCase))
                return next + 2;

            ThrowError(null, null, $"Field property not supported: {name}.{property}");
            return next + 2;
        }

        /// <summary>
        /// Converts a parameter to its variable. A parameter with many values is a list variable:
        /// its Count and its values by the index are the ones of the list. The label is taken with
        /// GetLabel, the viewer keeps the label of the chosen value.
        /// </summary>
        private int ConvertParameter(List<StiVbToken> tokens, string name, string property, int next, int end,
            StringBuilder result)
        {
            var variable = FindVariable(name);
            var isList = variable != null && IsListType(variable.Type);

            if (string.Equals(property, "IsMultiValue", StringComparison.OrdinalIgnoreCase))
            {
                result.Append(isList ? TrueKeyword : FalseKeyword);
                return next + 2;
            }

            if (variable == null)
            {
                ThrowError(null, null, $"Parameter not found: {name}");
                result.Append(NothingKeyword);
                return property != null ? next + 2 : next;
            }

            if (string.Equals(property, "Count", StringComparison.OrdinalIgnoreCase))
            {
                result.Append(isList ? $"{variable.Name}.Count" : "1");
                return next + 2;
            }

            if (string.Equals(property, "Label", StringComparison.OrdinalIgnoreCase))
            {
                result.Append($"GetLabel(\"{variable.Name}\")");
                return SkipIndex(tokens, next + 2, end);
            }

            result.Append(variable.Name);

            if (property == null)
                return next;

            if (!string.Equals(property, "Value", StringComparison.OrdinalIgnoreCase))
                return next;

            //the index of a value of the parameter, as in Parameters!P.Value(0)
            var open = next + 2;
            if (open < end && tokens[open].IsOperator("("))
            {
                var close = FindClosingParenthesis(tokens, open, end);
                if (close == -1)
                    return next + 2;

                if (isList)
                {
                    var index = ConvertTokens(tokens, open + 1, close, null, isCode: false).Trim();
                    result.Append(ConvertSyntaxToCSharp ? $"[{index}]" : $"({index})");
                }

                return close + 1;
            }

            return next + 2;
        }

        /// <summary>
        /// Skips the index after a property, as (0) in Parameters!P.Label(0): the label of the chosen
        /// values is one text.
        /// </summary>
        private static int SkipIndex(List<StiVbToken> tokens, int index, int end)
        {
            if (index >= end || !tokens[index].IsOperator("("))
                return index;

            var close = FindClosingParenthesis(tokens, index, end);

            return close == -1 ? index : close + 1;
        }

        private int ConvertGlobal(string name, string property, int next, StringBuilder result)
        {
            switch (name.ToLowerInvariant())
            {
                case "pagenumber":
                    result.Append("PageNumber");
                    return next;

                case "totalpages":
                    result.Append("TotalPageCount");
                    return next;

                case "overallpagenumber":
                    result.Append("PageNumberThrough");
                    return next;

                case "overalltotalpages":
                    result.Append("TotalPageCountThrough");
                    return next;

                case "reportname":
                    result.Append("ReportName");
                    return next;

                case "executiontime":
                    result.Append("Time");
                    return next;

                case "renderformat":
                    ThrowError(null, null, $"Global not supported: {name}");
                    result.Append(NothingKeyword);
                    return property != null ? next + 2 : next;

                default:
                    ThrowError(null, null, $"Global not supported: {name}");
                    result.Append(NothingKeyword);
                    return next;
            }
        }

        private string ConvertUser(string name)
        {
            if (string.Equals(name, "UserID", StringComparison.OrdinalIgnoreCase))
                return "System.Environment.UserName";

            if (string.Equals(name, "Language", StringComparison.OrdinalIgnoreCase))
                return "System.Globalization.CultureInfo.CurrentCulture.Name";

            ThrowError(null, null, $"User property not supported: {name}");
            return NothingKeyword;
        }

        private string ConvertVariable(string name)
        {
            var variable = FindVariable(name);
            if (variable != null)
                return variable.Name;

            ThrowError(null, null, $"Variable not found: {name}");
            return NothingKeyword;
        }

        /// <summary>
        /// Reporting Services takes ReportItems!Name.Value for the value of the text box with the name
        /// in the same scope, often to hide an item when a text box is empty. Stimulsoft has no such
        /// collection, so the value of the text box is put in its place: its expression in parentheses
        /// or its text as a string. A text box that is not found, has no value or refers to itself
        /// through other text boxes turns into Nothing and is named in the list of errors.
        /// </summary>
        private string ConvertReportItem(string name, string dataset)
        {
            var values = FindTextboxValues(name);

            if (values.Count == 0 || !inlinedReportItems.Add(name))
            {
                ThrowError(null, null, $"ReportItems not supported: {name}");
                return NothingKeyword;
            }

            try
            {
                var parts = values.Select(v => ConvertReportItemValue(v, dataset)).ToList();
                if (parts.Count == 1)
                    return parts[0];

                return "(" + string.Join(ConvertSyntaxToCSharp ? " + " : " & ", parts) + ")";
            }
            finally
            {
                inlinedReportItems.Remove(name);
            }
        }

        private string ConvertReportItemValue(string value, string dataset)
        {
            if (!IsExpression(value))
                return ConvertSyntaxToCSharp ? ToCSharpString(value) : "\"" + value.Replace("\"", "\"\"") + "\"";

            var tokens = StiVbTokenizer.Tokenize(value.Trim().Substring(1));

            return "(" + ConvertTokens(tokens, 0, tokens.Count, dataset, isCode: false).Trim() + ")";
        }

        /// <summary>
        /// Returns the values of the text runs of the text box with the given name, or the value of
        /// the text box itself in the schema of 2005.
        /// </summary>
        private List<string> FindTextboxValues(string name)
        {
            var values = new List<string>();
            var document = reportNode as XmlDocument ?? reportNode?.OwnerDocument;
            if (document == null)
                return values;

            var textbox = document.SelectNodes("//*[local-name()='Textbox']").Cast<XmlNode>()
                .FirstOrDefault(n => string.Equals(n.Attributes?["Name"]?.Value, name,
                    StringComparison.OrdinalIgnoreCase));

            if (textbox == null)
                return values;

            foreach (XmlNode node in textbox.SelectNodes(".//*[local-name()='Value']"))
            {
                var parent = node.ParentNode.LocalName;
                if (parent == "TextRun" || parent == "Textbox")
                    values.Add(GetNodeTextValue(node));
            }

            return values;
        }

        /// <summary>
        /// Returns the reference to a field in the code of the report. A field is looked for in the
        /// given data set, or among the fields of all data sets when the data set is not known.
        /// </summary>
        private string FindFieldReference(string name, string dataset)
        {
            var key = $"Fields!{name}.Value";

            if (dataset != null)
                key = $"{dataset}:{key}";

            return fieldsNames.TryGetValue(key, out var reference) ? reference : null;
        }

        /// <summary>
        /// Returns the variable of a parameter. The alias keeps the name of the parameter, the name of
        /// the variable differs when a data set has the name of the parameter.
        /// </summary>
        private StiVariable FindVariable(string name)
        {
            var variables = report.Dictionary.Variables.Cast<StiVariable>().ToList();

            return variables.FirstOrDefault(v => v.Category == ParametersCategory && IsSameName(v.Alias, name)) ??
                variables.FirstOrDefault(v => IsSameName(v.Name, name));
        }

        /// <summary>
        /// Returns the property after a member of a collection, as Value in Fields!Name.Value.
        /// </summary>
        private static string GetProperty(List<StiVbToken> tokens, int index, int end)
        {
            if (index + 1 < end && tokens[index].IsOperator(".") && tokens[index + 1].Type == StiVbTokenType.Identifier)
                return tokens[index + 1].Name;

            return null;
        }
        #endregion

        #region Expressions.Functions
        /// <summary>
        /// Converts an aggregate function to a function of Totals. The scope, the second argument, is
        /// kept when it names a data set: the function runs over the data source, and the fields of
        /// the argument are taken from that data set. The scope of a group of the data region is the
        /// header of the group, so an aggregate of the outer group in an inner one is not the total of
        /// the inner group; the scope of the data region itself is its data band, all its rows, see
        /// ResolveGroupScopes. Count and CountRows count the rows of the data source.
        /// </summary>
        private int ConvertAggregate(List<StiVbToken> tokens, int index, int open, int end, string dataset,
            StringBuilder result)
        {
            var close = FindClosingParenthesis(tokens, open, end);
            if (close == -1)
            {
                result.Append(tokens[index].Text);
                return index + 1;
            }

            var function = FindName(AggregateFunctions, tokens[index].Name);
            var arguments = SplitArguments(tokens, open + 1, close);
            var scope = GetScope(tokens, arguments, function == "CountRows" ? 0 : 1);
            var scopeDataSet = FindDataSetName(scope);
            var argumentDataSet = scopeDataSet ?? dataset;

            if (function == "Count" || function == "CountRows")
            {
                var dataSource = scopeDataSet ?? FindDataSetOfFields(tokens, arguments, argumentDataSet) ??
                    FindDataSetName(dataset);

                if (dataSource == null)
                {
                    ThrowError(null, null, $"Data set of the function not found: {function}");
                    result.Append("0");
                }
                else if (isInDataRegion && scopeDataSet == null && scope != null &&
                    IsGroupOrRegionName(scope))
                {
                    result.Append(MakeGroupScope(scope, $"Totals.Count({GroupScopeBand})",
                        MakeCountInScope(GetDataSourceReference(dataSource))));
                }

                else if (isInDataRegion && scopeDataSet == null)
                {
                    result.Append(MakeCountInScope(GetDataSourceReference(dataSource)));
                }
                else
                {
                    result.Append($"Totals.Count({GetDataSourceReference(dataSource)})");
                }

                return close + 1;
            }

            var argument = arguments.Count > 0
                ? ConvertTokens(tokens, arguments[0].Item1, arguments[0].Item2, argumentDataSet, isCode: false).Trim()
                : string.Empty;

            //Totals.Min and Totals.Max give a number, the dates and the strings have functions of their own
            if ((function == "Min" || function == "Max") && IsDateExpression(argument))
                function += "Date";

            else if ((function == "Min" || function == "Max") && IsStringExpression(argument))
                function += "Str";

            var total = MakeTotal(function, scopeDataSet != null ? GetDataSourceReference(scopeDataSet) : null,
                argument);

            //the scope of a group is its header, the band is known only when the bands of the region are made
            if (isInDataRegion && scopeDataSet == null && scope != null && IsGroupOrRegionName(scope))
                total = MakeGroupScope(scope, MakeTotal(function, GroupScopeBand, argument), total);

            result.Append(total);

            return close + 1;
        }

        /// <summary>
        /// Makes the function of Totals over the given band or data source, or over the band of the component.
        /// </summary>
        private string MakeTotal(string function, string scope, string argument)
        {
            var total = scope != null
                ? $"Totals.{function}({scope}, {argument})"
                : $"Totals.{function}({argument})";

            //Avg of Reporting Services is Double for any type but Decimal, Totals.Avg is always Decimal
            if (function == "Avg" && !IsDecimalExpression(argument))
                total = $"System.Convert.ToDouble({total})";

            return total;
        }

        private bool IsDateExpression(string expression)
        {
            var type = GetExpressionType(expression);

            return type == typeof(DateTime) || type == typeof(DateTime?);
        }

        private bool IsStringExpression(string expression)
        {
            return GetExpressionType(expression) == typeof(string);
        }

        private bool IsDecimalExpression(string expression)
        {
            var type = GetExpressionType(expression);

            return type == typeof(decimal) || type == typeof(decimal?);
        }

        /// <summary>
        /// Tells whether the converted expression is a number for sure: a number, a numeric field or
        /// variable, a sum, an average or a count, the first, the last, the least or the greatest of a
        /// numeric field.
        /// </summary>
        private bool IsNumericExpression(string expression)
        {
            if (Regex.IsMatch(expression, @"^-?[0-9]+(\.[0-9]+)?$"))
                return true;

            if (Regex.IsMatch(expression, @"^(System\.Convert\.ToDouble\()?Totals\.(Sum|Avg|Count|CountDistinct)\("))
                return true;

            var aggregate = Regex.Match(expression, @"^Totals\.(First|Last|Max|Min)\((\w+, )?(?<column>\w+\.\w+)\)$");
            if (aggregate.Success)
                expression = aggregate.Groups["column"].Value;

            var type = GetExpressionType(expression);

            return type != null && type.IsNumericType();
        }

        /// <summary>
        /// Converts Round, Ceiling and Floor. A field of Reporting Services is an Object, and the function
        /// picks the overload by the value when the report runs. The functions of Stimulsoft take a Double or
        /// a Decimal, so a value that is not a number for sure, as a String field of a local report or a
        /// division by First of a data set, is turned into a Double first. CDbl is known both to the compiler
        /// and to the parser, so the report works in both modes.
        /// </summary>
        private int ConvertRound(List<StiVbToken> tokens, int index, int open, int end, string dataset,
            StringBuilder result)
        {
            var close = FindClosingParenthesis(tokens, open, end);
            var arguments = close == -1 ? null : SplitArguments(tokens, open + 1, close);
            if (arguments == null || arguments.Count == 0)
            {
                result.Append(tokens[index].Text);
                return index + 1;
            }

            var values = arguments
                .Select(a => ConvertTokens(tokens, a.Item1, a.Item2, dataset, isCode: false).Trim())
                .ToList();

            if (!IsNumericExpression(values[0]))
                values[0] = ConvertSyntaxToCSharp ? $"System.Convert.ToDouble({values[0]})" : $"CDbl({values[0]})";

            result.Append($"{tokens[index].Text}({string.Join(", ", values)})");

            return close + 1;
        }

        /// <summary>
        /// Converts Join of Visual Basic, which joins the values of a parameter with many values, to
        /// String.Join, known both to the compiler and to the parser; the separator is a space by default.
        /// The labels of such a parameter are joined already, so Join of them gives them as they are.
        /// </summary>
        private int ConvertJoin(List<StiVbToken> tokens, int index, int open, int end, string dataset,
            StringBuilder result)
        {
            var close = FindClosingParenthesis(tokens, open, end);
            var arguments = close == -1 ? null : SplitArguments(tokens, open + 1, close);
            if (arguments == null || arguments.Count == 0 || arguments.Count > 2)
            {
                result.Append(tokens[index].Text);
                return index + 1;
            }

            var values = ConvertTokens(tokens, arguments[0].Item1, arguments[0].Item2, dataset, isCode: false).Trim();
            var separator = arguments.Count > 1
                ? ConvertTokens(tokens, arguments[1].Item1, arguments[1].Item2, dataset, isCode: false).Trim()
                : "\" \"";

            result.Append(values.StartsWith("GetLabel(") ? values : $"String.Join({separator}, {values})");

            return close + 1;
        }

        /// <summary>
        /// Puts the text of the query of a data set as a string instead of DataSets("X").CommandText or
        /// DataSets!X.CommandText; the text is known when the report is imported. When the form does not fit
        /// or the data set is not found, -1 is returned.
        /// </summary>
        private int ConvertDataSetCommandText(List<StiVbToken> tokens, int index, int end, StringBuilder result)
        {
            var position = GetNextSignificant(tokens, index + 1, end);
            string name = null;

            if (position + 1 < end && tokens[position].IsOperator("!") &&
                tokens[position + 1].Type == StiVbTokenType.Identifier)
            {
                name = tokens[position + 1].Name;
                position += 2;
            }
            else if (position + 2 < end && tokens[position].IsOperator("(") &&
                tokens[position + 1].Type == StiVbTokenType.String && tokens[position + 2].IsOperator(")"))
            {
                name = tokens[position + 1].StringValue;
                position += 3;
            }

            var property = name != null ? GetProperty(tokens, position, end) : null;
            if (!string.Equals(property, "CommandText", StringComparison.OrdinalIgnoreCase))
                return -1;

            var document = reportNode as XmlDocument ?? reportNode?.OwnerDocument;
            var command = document?.SelectNodes("//*[local-name()='DataSet']").Cast<XmlNode>()
                .FirstOrDefault(n => IsSameName(n.Attributes?["Name"]?.Value, name))?
                .SelectSingleNode("*[local-name()='Query']/*[local-name()='CommandText']");

            if (command == null)
                return -1;

            result.Append(ConvertSyntaxToCSharp ? ToCSharpString(command.InnerText) : ToVbString(command.InnerText));

            return position + 2;
        }

        /// <summary>
        /// Writes a text as a string of Visual Basic in an expression of the report. A string of the compiler
        /// takes no line break, and a brace would end the expression, so they are joined by their codes.
        /// </summary>
        private static string ToVbString(string text)
        {
            var parts = Regex.Split(text, @"(\r\n|\r|\n|\{|\})");
            var result = new StringBuilder();

            foreach (var part in parts)
            {
                if (part.Length == 0)
                    continue;

                if (result.Length > 0)
                    result.Append(" & ");

                switch (part)
                {
                    case "\r\n":
                    case "\r":
                    case "\n":
                        result.Append("vbCrLf");
                        break;

                    case "{":
                    case "}":
                        result.Append($"ChrW({(int)part[0]})");
                        break;

                    default:
                        result.Append("\"" + part.Replace("\"", "\"\"") + "\"");
                        break;
                }
            }

            return result.Length > 0 ? result.ToString() : "\"\"";
        }

        /// <summary>
        /// Puts the value of a named constant of Visual Basic instead of its name: the code of an interval
        /// instead of DateInterval.Day, the number of a format instead of DateFormat.ShortDate, vbCrLf instead
        /// of Environment.NewLine. The parser does not know these names, the functions of Visual Basic take the
        /// values as well. The longest name written is taken; when no name fits, -1 is returned.
        /// </summary>
        private static int ConvertNamedValue(List<StiVbToken> tokens, int index, int end, StringBuilder result)
        {
            var name = tokens[index].Name;
            var position = index;
            string value = null;
            var valueEnd = -1;

            while (true)
            {
                if (NamedValues.TryGetValue(name, out var found) ||
                    name.StartsWith(VbNamespace, StringComparison.OrdinalIgnoreCase) &&
                    NamedValues.TryGetValue(name.Substring(VbNamespace.Length), out found))
                {
                    value = found;
                    valueEnd = position + 1;
                }

                var dot = GetNextSignificant(tokens, position + 1, end);
                var next = dot < end ? GetNextSignificant(tokens, dot + 1, end) : end;
                if (dot >= end || !tokens[dot].IsOperator(".") || next >= end ||
                    tokens[next].Type != StiVbTokenType.Identifier)
                    break;

                name += "." + tokens[next].Name;
                position = next;
            }

            if (value == null)
                return -1;

            result.Append(value);

            return valueEnd;
        }

        /// <summary>
        /// A function with no counterpart in Stimulsoft turns into Nothing, so the report compiles,
        /// and the function is named in the list of errors.
        /// </summary>
        private int SkipUnsupportedFunction(List<StiVbToken> tokens, int index, int open, int end,
            StringBuilder result)
        {
            var close = FindClosingParenthesis(tokens, open, end);
            if (close == -1)
            {
                result.Append(tokens[index].Text);
                return index + 1;
            }

            ThrowError(null, null, $"Function not supported: {FindName(UnsupportedFunctions, tokens[index].Name)}");
            result.Append(NothingKeyword);

            return close + 1;
        }

        /// <summary>
        /// Puts the text instead of the name and its parentheses with the arguments. The open index
        /// is -1 for the name written without the parentheses.
        /// </summary>
        private static int AppendInsteadOfCall(List<StiVbToken> tokens, int index, int open, int end, string text,
            StringBuilder result)
        {
            result.Append(text);

            if (open == -1)
                return index + 1;

            var close = FindClosingParenthesis(tokens, open, end);

            return close == -1 ? index + 1 : close + 1;
        }

        private static string GetScope(List<StiVbToken> tokens, List<Tuple<int, int>> arguments, int position)
        {
            if (arguments.Count <= position)
                return null;

            var start = GetNextSignificant(tokens, arguments[position].Item1, arguments[position].Item2);
            if (start >= arguments[position].Item2 || tokens[start].Type != StiVbTokenType.String)
                return null;

            return tokens[start].StringValue;
        }

        /// <summary>
        /// Returns the name of the data set the first field of the arguments belongs to.
        /// </summary>
        private string FindDataSetOfFields(List<StiVbToken> tokens, List<Tuple<int, int>> arguments, string dataset)
        {
            if (arguments.Count == 0)
                return null;

            for (var index = arguments[0].Item1; index < arguments[0].Item2; index++)
            {
                if (!IsCollection(tokens, index, arguments[0].Item2) || !tokens[index].IsName("Fields"))
                    continue;

                var name = tokens[index + 2].Name;

                if (dataset != null)
                    return fieldsNames.ContainsKey($"{dataset}:Fields!{name}.Value") ? dataset : null;

                //a field out of a data set is taken from the last data set that has it, see FindFieldReference
                return report.Dictionary.DataSources.Cast<StiDataSource>()
                    .LastOrDefault(d => fieldsNames.ContainsKey($"{d.Name}:Fields!{name}.Value"))?.Name;
            }

            return null;
        }

        /// <summary>
        /// Returns the name of the data source made of the data set with the given name, or null
        /// when the name is a group or a data region.
        /// </summary>
        private string FindDataSetName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;

            return report.Dictionary.DataSources.Cast<StiDataSource>()
                .FirstOrDefault(d => string.Equals(d.Name, name, StringComparison.OrdinalIgnoreCase))?.Name;
        }

        private string GetDataSourceReference(string dataSourceName)
        {
            return StiNameValidator.CorrectName(dataSourceName, report);
        }
        #endregion

        #region Expressions.Tokens
        private static int GetNextSignificant(List<StiVbToken> tokens, int index, int end)
        {
            while (index < end && !tokens[index].IsSignificant)
            {
                index++;
            }

            return index;
        }

        private static StiVbToken GetPreviousSignificant(List<StiVbToken> tokens, int index)
        {
            for (var position = index - 1; position >= 0; position--)
            {
                if (tokens[position].IsSignificant)
                    return tokens[position];
            }

            return null;
        }

        /// <summary>
        /// Returns the index of the parenthesis that closes the one at the given index, or -1.
        /// </summary>
        private static int FindClosingParenthesis(List<StiVbToken> tokens, int open, int end)
        {
            var level = 0;

            for (var index = open; index < end; index++)
            {
                if (tokens[index].IsOperator("("))
                    level++;

                else if (tokens[index].IsOperator(")"))
                {
                    level--;

                    if (level == 0)
                        return index;
                }
            }

            return -1;
        }

        /// <summary>
        /// Splits the tokens between the parentheses of a call into the arguments, by the commas
        /// that are not inside nested parentheses. Every argument is the start and the end index.
        /// </summary>
        private static List<Tuple<int, int>> SplitArguments(List<StiVbToken> tokens, int start, int end)
        {
            var arguments = new List<Tuple<int, int>>();
            var level = 0;
            var argumentStart = start;

            for (var index = start; index < end; index++)
            {
                var token = tokens[index];

                if (token.IsOperator("("))
                    level++;

                else if (token.IsOperator(")"))
                    level--;

                else if (token.IsOperator(",") && level == 0)
                {
                    arguments.Add(Tuple.Create(argumentStart, index));
                    argumentStart = index + 1;
                }
            }

            if (GetNextSignificant(tokens, argumentStart, end) < end || arguments.Count > 0)
                arguments.Add(Tuple.Create(argumentStart, end));

            return arguments;
        }

        /// <summary>
        /// Checks that the underscore at the index continues the line: only spaces follow it before
        /// the end of the line.
        /// </summary>
        private static bool IsLineContinuation(List<StiVbToken> tokens, int index, int end)
        {
            var next = index + 1;

            while (next < end && tokens[next].Type == StiVbTokenType.Whitespace)
            {
                next++;
            }

            return next == end || tokens[next].Type == StiVbTokenType.NewLine;
        }

        /// <summary>
        /// Returns the name from the list as it is written there, the name is compared in any case.
        /// </summary>
        private static string FindName(string[] names, string name)
        {
            return names.FirstOrDefault(n => string.Equals(n, name, StringComparison.OrdinalIgnoreCase));
        }

        private static void AppendSpace(StringBuilder result)
        {
            if (result.Length > 0 && !char.IsWhiteSpace(result[result.Length - 1]))
                result.Append(' ');
        }

        private static string ToCSharpString(string value)
        {
            return "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }
        #endregion
    }
}
