#region Copyright (C) 2003-2026 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports												}
{	                         										}
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

using Stimulsoft.Report.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts the cells of a tablix with the groups of the columns. The columns are cross bands over
    /// a copy of the data source: the rows of a group of the columns are the rows of the copy, the bands
    /// of the rows go over the data source itself. A cell of Reporting Services aggregates the rows of
    /// its group of the rows and of its group of the columns together, so an aggregate of a cell runs
    /// over the rows of the copy in the group of the columns, which have the keys of the groups of the
    /// rows of the current row.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts
        private const string DoubleWrapper = "System.Convert.ToDouble(";

        private static readonly string[] MatrixAggregates =
            { "Sum", "Avg", "Min", "Max", "MinDate", "MaxDate", "Count" };
        #endregion

        #region Matrix
        /// <summary>
        /// Converts the cells of the groups of the columns. The row scopes give the conditions of the
        /// groups of the rows each band of the rows is printed in, from the outer one: none for the
        /// header and the footer of the rows, null for the details, which have no keys to filter by. The
        /// count scopes give the band whose rows the static columns aggregate, as the total of a row.
        /// </summary>
        private void ResolveMatrixCells(StiContainer container, Dictionary<StiBand, List<string>> rowScopes,
            Dictionary<StiBand, StiBand> countScopes, string dataSetName, string copyName)
        {
            var reference = GetDataSourceReference(dataSetName);
            var copyReference = GetDataSourceReference(copyName);

            var columnHeaders = container.Components[$"{container.Name}_Hdc"] as StiBand;
            if (columnHeaders != null)
            {
                foreach (var text in GetCrossTexts(columnHeaders, isGroup: true))
                {
                    //the headers of the columns show the values of the groups of the columns
                    var columnBand = GetColumnGroupHeaderName(text.Parent);
                    var value = ReplaceTotals(text.Text.Value, (function, arguments, isDouble) =>
                        MakeMatrixAggregate(function, arguments, isDouble, columnBand, new List<string>(), reference,
                            copyReference));

                    text.Text.Value = ToCopyReference(value, reference, copyReference);
                }
            }

            foreach (var pair in rowScopes)
            {
                foreach (var text in GetCrossTexts(pair.Key, isGroup: true))
                {
                    if (pair.Value == null)
                    {
                        if (text.Text.Value != null && text.Text.Value.Contains("Totals."))
                            ThrowError(null, null, $"Aggregate of details by groups of columns: {text.Name}");

                        continue;
                    }

                    var columnBand = GetColumnGroupHeaderName(text.Parent);
                    text.Text.Value = ReplaceTotals(text.Text.Value, (function, arguments, isDouble) =>
                        MakeMatrixAggregate(function, arguments, isDouble, columnBand, pair.Value, reference,
                            copyReference));
                }
            }

            foreach (var pair in countScopes)
            {
                foreach (var text in GetCrossTexts(pair.Key, isGroup: false))
                {
                    text.Text.Value = ReplaceTotals(text.Text.Value, (function, arguments, isDouble) =>
                        MakeScopedAggregate(function, arguments, isDouble, pair.Value.Name));
                }
            }
        }

        /// <summary>
        /// Returns the texts of the cross bands in a band of the rows: of the headers and the footers
        /// of the groups of the columns, or of the static columns.
        /// </summary>
        private static List<StiText> GetCrossTexts(StiBand band, bool isGroup)
        {
            return band.GetComponents().Cast<StiComponent>()
                .Where(c => isGroup
                    ? c is StiCrossGroupHeaderBand || c is StiCrossGroupFooterBand
                    : c is StiCrossHeaderBand || c is StiCrossFooterBand)
                .SelectMany(c => ((StiBand)c).GetComponents().Cast<StiComponent>().OfType<StiText>())
                .ToList();
        }

        /// <summary>
        /// The rows of a group of the columns are the rows of its header, the footer of the group of
        /// the same level counts them too.
        /// </summary>
        private static string GetColumnGroupHeaderName(StiContainer band)
        {
            while (band != null && !(band is StiCrossGroupHeaderBand) && !(band is StiCrossGroupFooterBand))
            {
                band = band.Parent;
            }

            if (band is StiCrossGroupFooterBand)
            {
                var index = band.Name.LastIndexOf("CGf", StringComparison.Ordinal);
                return band.Name.Substring(0, index) + "CGh" + band.Name.Substring(index + 3);
            }

            return band?.Name;
        }

        /// <summary>
        /// Replaces the functions of Totals in a text by the given maker, which takes the name of the
        /// function, its arguments and whether Avg is turned into a Double, and returns null to keep the
        /// function. The Double of Avg is replaced with the function, so the maker can keep a cell empty.
        /// </summary>
        private static string ReplaceTotals(string text, Func<string, List<string>, bool, string> make)
        {
            if (string.IsNullOrEmpty(text) || !text.Contains("Totals."))
                return text;

            var result = new StringBuilder();
            var index = 0;

            while (index < text.Length)
            {
                var start = text.IndexOf("Totals.", index, StringComparison.Ordinal);
                var open = start < 0 ? -1 : text.IndexOf('(', start);
                var close = open < 0 ? -1 : FindClosingParenthesis(text, open);
                if (close < 0)
                    break;

                var function = text.Substring(start + 7, open - start - 7).Trim();
                var arguments = SplitArguments(text.Substring(open + 1, close - open - 1));
                var isDouble = IsDoubleWrapped(text, index, start, close);
                var aggregate = make(function, arguments, isDouble);

                if (aggregate == null)
                {
                    result.Append(text, index, close + 1 - index);
                    index = close + 1;
                    continue;
                }

                if (isDouble)
                {
                    start -= DoubleWrapper.Length;
                    close++;
                }

                result.Append(text, index, start - index);
                result.Append(aggregate);
                index = close + 1;
            }

            result.Append(text, index, text.Length - index);

            return result.ToString();
        }

        private static bool IsDoubleWrapped(string text, int index, int start, int close)
        {
            return start - index >= DoubleWrapper.Length && close + 1 < text.Length && text[close + 1] == ')' &&
                string.CompareOrdinal(text, start - DoubleWrapper.Length, DoubleWrapper, 0, DoubleWrapper.Length) == 0;
        }

        /// <summary>
        /// Makes an aggregate of a cell of a group of the columns over the rows of the copy in the group,
        /// which have the keys of the groups of the rows of the current row.
        /// </summary>
        private string MakeMatrixAggregate(string function, List<string> arguments, bool isDouble,
            string columnBand, List<string> rowConditions, string reference, string copyReference)
        {
            if (arguments.Count != 1 || columnBand == null)
                return null;

            //Count of the band of the component has a marker, Count of a data source or of a band of a scope stays
            if (function == "Count" && !arguments[0].Trim().StartsWith(CountScopeMarker, StringComparison.Ordinal))
                return null;

            if (!MatrixAggregates.Contains(function))
            {
                ThrowError(null, null, $"Aggregate in a cell of a matrix not converted: {function}");
                return null;
            }

            var filter = string.Join(ConvertSyntaxToCSharp ? " && " : " AndAlso ", rowConditions.Select(c =>
                MakeEquals(ToCopyReference(c, reference, copyReference), c)));

            if (function == "Count")
            {
                return filter.Length == 0
                    ? $"Totals.Count({columnBand})"
                    : $"Totals.Sum({columnBand}, {MakeIif(filter, "1", "0")})";
            }

            var value = ToCopyReference(arguments[0].Trim(), reference, copyReference);
            if (filter.Length == 0)
                return MakeDouble(isDouble, $"Totals.{function}({columnBand}, {value})");

            //a cell without rows is empty in Reporting Services, Totals of no values give 0
            var aggregate = MakeDouble(isDouble,
                $"Totals.{function}({columnBand}, {MakeIif(filter, value, NothingKeyword)})");
            var count = $"Totals.Sum({columnBand}, {MakeIif(filter, "1", "0")})";
            var isEmpty = ConvertSyntaxToCSharp ? $"{count} == 0" : $"{count} = 0";

            return MakeIif(isEmpty, NothingKeyword, aggregate);
        }

        /// <summary>
        /// Gives an aggregate of a static column the band of its row: the group for a header or a footer
        /// of a group, the data band for the header and the footer of the rows. Count is resolved by
        /// ResolveCountScopes.
        /// </summary>
        private static string MakeScopedAggregate(string function, List<string> arguments, bool isDouble,
            string band)
        {
            if (arguments.Count != 1 || function == "Count")
                return null;

            return MakeDouble(isDouble, $"Totals.{function}({band}, {arguments[0].Trim()})");
        }

        private static string MakeDouble(bool isDouble, string value)
        {
            return isDouble ? $"{DoubleWrapper}{value})" : value;
        }

        private string MakeEquals(string left, string right)
        {
            return ConvertSyntaxToCSharp
                ? $"object.Equals({left}, {right})"
                : $"System.Object.Equals({left}, {right})";
        }

        private string MakeIif(string condition, string whenTrue, string whenFalse)
        {
            return ConvertSyntaxToCSharp
                ? $"(({condition}) ? (object)({whenTrue}) : {whenFalse})"
                : $"IIF({condition}, {whenTrue}, {whenFalse})";
        }

        /// <summary>
        /// Points the fields of the data source to its copy, DataSet1.Name turns into DataSet1Copy.Name.
        /// </summary>
        private static string ToCopyReference(string expression, string reference, string copyReference)
        {
            if (string.IsNullOrEmpty(expression) || reference == copyReference)
                return expression;

            return Regex.Replace(expression, @"(?<![\w.])" + Regex.Escape(reference) + @"\.", copyReference + ".");
        }

        /// <summary>
        /// Finds the parenthesis that closes the given one, the parentheses in a string do not count.
        /// </summary>
        private static int FindClosingParenthesis(string text, int open)
        {
            var level = 0;
            var inString = false;

            for (var index = open; index < text.Length; index++)
            {
                var ch = text[index];

                if (ch == '"')
                    inString = !inString;

                else if (!inString && ch == '(')
                    level++;

                else if (!inString && ch == ')' && --level == 0)
                    return index;
            }

            return -1;
        }

        /// <summary>
        /// Splits the arguments of a function by the commas outside the parentheses and the strings.
        /// </summary>
        private static List<string> SplitArguments(string text)
        {
            var arguments = new List<string>();
            var level = 0;
            var inString = false;
            var start = 0;

            for (var index = 0; index < text.Length; index++)
            {
                var ch = text[index];

                if (ch == '"')
                    inString = !inString;

                else if (!inString && ch == '(')
                    level++;

                else if (!inString && ch == ')')
                    level--;

                else if (!inString && ch == ',' && level == 0)
                {
                    arguments.Add(text.Substring(start, index - start));
                    start = index + 1;
                }
            }

            arguments.Add(text.Substring(start));

            return arguments;
        }
        #endregion
    }
}
