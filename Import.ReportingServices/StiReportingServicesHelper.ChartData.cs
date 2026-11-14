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

using Stimulsoft.Report.Chart;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Groups the data of a chart as Reporting Services does. A data point of Reporting Services is one per
    /// category and series group, and its value, as Sum(Fields!Amount.Value), is an aggregate of the rows of
    /// the group. A chart of Stimulsoft takes a point per row and has no aggregates of its own: Totals in a value
    /// gave 0 in the compilation and the total of all rows in the interpretation, and every row was a point.
    /// So a chart with aggregates of fields reads a virtual source grouped by the fields of its categories and
    /// series, where every aggregate is a column of totals; the expressions of the chart read these columns, and
    /// a field out of an aggregate reads the first value of the group. A group by an expression, as
    /// MonthName(Month(Fields!Date.Value)), is a calculated field of the data set.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Fields.ChartData
        private static readonly Regex ChartAggregateRegex = new Regex(
            @"\b(?<function>Sum|CountDistinct|Count|Avg|Min|Max|First|Last)\s*\(" +
            @"\s*Fields\s*!\s*(?<field>\w+)\s*\.\s*Value\s*\)",
            RegexOptions.IgnoreCase);
        private static readonly Regex ChartScopedAggregateRegex = new Regex(
            @"\b(Sum|CountDistinct|Count|Avg|Min|Max|First|Last)\s*\(\s*Fields\s*!\s*\w+\s*\.\s*Value\s*,",
            RegexOptions.IgnoreCase);
        private static readonly Regex ChartGroupFieldRegex = new Regex(
            @"^=\s*Fields\s*!\s*(?<field>\w+)\s*\.\s*Value\s*$",
            RegexOptions.IgnoreCase);
        private static readonly Regex ChartGroupNameRegex = new Regex(
            @"^=\s*(?<name>MonthName|WeekdayName)\s*\(\s*(?<inner>(Month|Weekday)\s*\(.*\))\s*\)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private Dictionary<string, string> chartAggregates;
        private Dictionary<string, string> chartGroupExpressions;
        private string chartAggregatesSource;

        /// <summary>
        /// The data sets of the charts with a source of their own in the cells of a data region.
        /// </summary>
        private Dictionary<StiChart, string> regionCharts = new Dictionary<StiChart, string>();
        #endregion

        #region Methods.ChartData
        /// <summary>
        /// Gives the charts of the cells of a data region the rows of the groups of their cells. A chart of
        /// Reporting Services in a cell aggregates the rows of the cell, the chart of Stimulsoft reads all the rows
        /// of its source. So the source of the chart is grouped by the keys of the groups of the rows too, and the
        /// chart takes the rows with the keys of the current row of the data set, which stands on the first row of
        /// the group when its header is printed; the source of the chart is another one, so the data set keeps its
        /// place. A row of the details has no keys, all its fields are the keys. A key that is not a field is
        /// named.
        /// </summary>
        private void ResolveChartScopes(Dictionary<StiBand, List<string>> rowScopes)
        {
            foreach (var pair in rowScopes)
            {
                foreach (var chart in pair.Key.GetComponents().Cast<StiComponent>().OfType<StiChart>())
                {
                    if (!regionCharts.TryGetValue(chart, out var dataset) || pair.Value?.Count == 0)
                        continue;

                    //a row of the details has no keys, it is the row with all the same fields
                    var conditions = pair.Value ?? report.Dictionary.DataSources[dataset].Columns
                        .Cast<StiDataColumn>().Select(c => GetColumnReference(dataset, c.Name)).ToList();

                    var keys = conditions.Select(c => GetChartKeyColumn(c, dataset)).ToList();
                    if (keys.Contains(null))
                    {
                        ThrowError(null, null,
                            $"Chart in a group by an expression shows the rows of the table: {chart.Name}");
                        continue;
                    }

                    var chartSource = (StiVirtualSource)report.Dictionary.DataSources[chart.DataSourceName];
                    AddChartKeyColumns(chartSource, keys);

                    for (var index = 0; index < keys.Count; index++)
                    {
                        var key = GetColumnReference(chartSource.Name, keys[index]);
                        chart.Filters.Add(new StiFilter(MakeEquals(key, conditions[index])));
                    }

                    chart.FilterOn = true;
                }
            }
        }

        /// <summary>
        /// Returns the column of the data set a condition of a group reads, or null for another expression.
        /// </summary>
        private string GetChartKeyColumn(string condition, string dataset)
        {
            var source = report.Dictionary.DataSources[dataset];
            if (source == null || condition == null)
                return null;

            return source.Columns.Cast<StiDataColumn>()
                .FirstOrDefault(c => GetColumnReference(dataset, c.Name) == condition)?.Name;
        }

        /// <summary>
        /// Groups the source of a chart by the keys first; a key is no more the first value of its group.
        /// </summary>
        private static void AddChartKeyColumns(StiVirtualSource chartSource, List<string> keys)
        {
            var newKeys = keys.Where(k => !chartSource.GroupColumns.Contains(k)).ToList();
            chartSource.GroupColumns = newKeys.Concat(chartSource.GroupColumns).ToArray();

            var results = new List<string>();
            for (var index = 0; index + 2 < chartSource.Results.Length; index += 3)
            {
                var isKey = newKeys.Contains(chartSource.Results[index]) &&
                    string.IsNullOrEmpty(chartSource.Results[index + 1]);
                if (!isKey)
                    results.AddRange(chartSource.Results.Skip(index).Take(3));
            }

            chartSource.Results = results.ToArray();
        }

        /// <summary>
        /// Creates the virtual source of the data of the chart and returns its name, or null when the chart has
        /// no aggregates of fields; such a chart keeps the data set. A chart in a cell of a data region always
        /// gets its source, which is grouped by the keys of the cell later; a field out of an aggregate is the
        /// first value of its group, as in Reporting Services.
        /// </summary>
        private string PrepareChartData(XmlNode baseNode, StiChart chart, string dataset, bool isInRegion)
        {
            chartAggregates = null;
            chartGroupExpressions = null;
            chartAggregatesSource = null;

            var masterSource = string.IsNullOrEmpty(dataset) ? null : report.Dictionary.DataSources[dataset];
            if (masterSource == null)
                return null;

            var groupColumns = new List<string>();
            var groupExpressions = new Dictionary<string, string>();
            var groupCalcColumns = new List<StiCalcDataColumn>();
            var dataNodes = new List<XmlNode>();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartCategoryHierarchy":
                    case "CategoryGroupings":
                    case "ChartSeriesHierarchy":
                    case "SeriesGroupings":
                        CollectChartGroupColumns(node, chart, masterSource, dataset, groupColumns, groupExpressions,
                            groupCalcColumns);
                        dataNodes.Add(node);
                        break;

                    case "ChartData":
                        dataNodes.Add(node);
                        break;
                }
            }

            var aggregates = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var results = new List<string>();
            var resultColumns = new List<StiDataColumn>();
            foreach (var node in dataNodes)
            {
                CollectChartAggregates(node, masterSource, dataset, aggregates, results, resultColumns);
            }

            if (aggregates.Count == 0 && !isInRegion)
                return null;

            foreach (var node in dataNodes)
            {
                CollectChartSortKeys(node, masterSource, dataset, groupColumns, aggregates, results, resultColumns);
            }

            //a group by an expression is a calculated field of the data set, the virtual source groups by it
            foreach (var column in groupCalcColumns)
            {
                masterSource.Columns.Add(column);
                fieldsNames[$"{dataset}:Fields!{column.Name}.Value"] = GetColumnReference(dataset, column.Name);
            }

            var name = StiNameCreation.CreateDataSourcesName(report, chart.Name + "Data");
            var chartSource = new StiVirtualSource(dataset, name);

            //a calculated field is taken as a plain field, the source of the chart reads its value
            foreach (StiDataColumn column in masterSource.Columns)
            {
                chartSource.Columns.Add(new StiDataColumn(column.Name, column.Type));
                if (groupColumns.Contains(column.Name))
                    continue;

                results.Add(column.Name);
                results.Add(string.Empty);
                results.Add(column.Name);
            }

            foreach (var column in resultColumns)
            {
                chartSource.Columns.Add(column);
            }

            chartSource.GroupColumns = groupColumns.ToArray();
            chartSource.Results = results.ToArray();
            report.Dictionary.DataSources.Add(chartSource);

            //the fields of the data set are the fields of the source of the chart
            var prefix = dataset + ":Fields!";
            foreach (var key in fieldsNames.Keys.Where(k => k.StartsWith(prefix, StringComparison.Ordinal)).ToList())
            {
                var reference = fieldsNames[key];
                var columnName = reference.Substring(reference.LastIndexOf('.') + 1);
                fieldsNames[name + ":" + key.Substring(dataset.Length + 1)] = GetColumnReference(name, columnName);
            }

            foreach (var column in resultColumns)
            {
                fieldsNames[$"{name}:Fields!{column.Name}.Value"] = GetColumnReference(name, column.Name);
            }

            chartAggregates = aggregates;
            chartGroupExpressions = groupExpressions;
            chartAggregatesSource = name;

            return name;
        }

        /// <summary>
        /// Adds the columns of the groups of a hierarchy, the categories first. A group by an expression gets a
        /// calculated field of the data set, and the expression of the group reads this field in the chart.
        /// </summary>
        private void CollectChartGroupColumns(XmlNode baseNode, StiChart chart, StiDataSource masterSource,
            string dataset, List<string> groupColumns, Dictionary<string, string> groupExpressions,
            List<StiCalcDataColumn> groupCalcColumns)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name != "GroupExpression")
                {
                    CollectChartGroupColumns(node, chart, masterSource, dataset, groupColumns, groupExpressions,
                        groupCalcColumns);
                    continue;
                }

                var expression = (GetNodeTextValue(node) ?? string.Empty).Trim();
                var match = ChartGroupFieldRegex.Match(expression);
                var column = match.Success
                    ? GetChartColumnName(match.Groups["field"].Value, masterSource, dataset)
                    : null;
                if (column == null && IsExpression(expression))
                {
                    column = GetChartGroupCalcColumn(expression, chart, masterSource, groupExpressions,
                        groupCalcColumns);
                }

                if (column != null && !groupColumns.Contains(column))
                    groupColumns.Add(column);
            }
        }

        /// <summary>
        /// Returns the calculated field of a group by an expression. The name of a month or a day of the week is
        /// a group of its number: the groups are the same, and they are sorted as in the calendar, not as text.
        /// </summary>
        private string GetChartGroupCalcColumn(string expression, StiChart chart, StiDataSource masterSource,
            Dictionary<string, string> groupExpressions, List<StiCalcDataColumn> groupCalcColumns)
        {
            var nameMatch = ChartGroupNameRegex.Match(expression);
            var value = nameMatch.Success && IsWholeCall(nameMatch.Groups["inner"].Value)
                ? "=" + nameMatch.Groups["inner"].Value
                : expression;

            var key = NormalizeChartGroupExpression(value);
            if (groupExpressions.TryGetValue(key, out var existing))
                return existing;

            var name = chart.Name + "Group";
            for (var index = 2; masterSource.Columns[name] != null ||
                groupCalcColumns.Any(c => c.Name == name); index++)
            {
                name = chart.Name + "Group" + index;
            }

            groupCalcColumns.Add(new StiCalcDataColumn(name, name, typeof(object), value));
            groupExpressions[key] = name;

            return name;
        }

        /// <summary>
        /// Returns true when the text is one call, as Month(x), and not a call followed by more, as Month(x), 1.
        /// </summary>
        private static bool IsWholeCall(string text)
        {
            var depth = 0;
            var opened = false;
            for (var index = 0; index < text.Length; index++)
            {
                if (text[index] == '(')
                {
                    depth++;
                    opened = true;
                }

                else if (text[index] == ')')
                {
                    depth--;
                    if (depth == 0 && opened && index < text.Length - 1)
                        return false;
                }
            }

            return depth == 0;
        }

        private static string NormalizeChartGroupExpression(string expression)
        {
            return Regex.Replace(expression, @"\s+", string.Empty).ToLowerInvariant();
        }

        /// <summary>
        /// Adds a column of totals for every aggregate of a field in the expressions of the node.
        /// </summary>
        private void CollectChartAggregates(XmlNode baseNode, StiDataSource masterSource, string dataset,
            Dictionary<string, string> aggregates, List<string> results, List<StiDataColumn> resultColumns)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.NodeType == XmlNodeType.Element)
                {
                    CollectChartAggregates(node, masterSource, dataset, aggregates, results, resultColumns);
                    continue;
                }

                if (node.NodeType != XmlNodeType.Text || !IsExpression(node.Value))
                    continue;

                foreach (Match match in ChartAggregateRegex.Matches(node.Value))
                {
                    var function = NormalizeChartAggregate(match.Groups["function"].Value);
                    var column = GetChartColumnName(match.Groups["field"].Value, masterSource, dataset);
                    var key = $"{function}|{match.Groups["field"].Value}";
                    if (column == null || aggregates.ContainsKey(key))
                        continue;

                    var resultName = function + column;
                    for (var index = 2; masterSource.Columns[resultName] != null ||
                        resultColumns.Any(c => c.Name == resultName); index++)
                    {
                        resultName = function + column + index;
                    }

                    var type = masterSource.Columns[column].Type;
                    if (function == "Sum" || function == "Avg")
                        type = typeof(decimal);

                    else if (function == "Count" || function == "CountDistinct")
                        type = typeof(long);

                    aggregates[key] = resultName;
                    resultColumns.Add(new StiDataColumn(resultName, type));
                    results.Add(column);
                    results.Add(function);
                    results.Add(resultName);
                }
            }
        }

        /// <summary>
        /// Adds a column of the least value for every field out of the groups a group of the chart is sorted by,
        /// as the months by the key of the month. Such a field is the first value of its group, but the grouping
        /// of the virtual source does not keep the order of the rows, so the first value was of any year and the
        /// months came in no order; the least value is the key the sorting means. See CollectChartGroupSort.
        /// </summary>
        private void CollectChartSortKeys(XmlNode baseNode, StiDataSource masterSource, string dataset,
            List<string> groupColumns, Dictionary<string, string> aggregates, List<string> results,
            List<StiDataColumn> resultColumns)
        {
            var values = baseNode.SelectNodes(
                ".//*[local-name()='SortExpressions']/*[local-name()='SortExpression']/*[local-name()='Value'] | " +
                ".//*[local-name()='Sorting']/*[local-name()='SortBy']/*[local-name()='SortExpression']");

            foreach (XmlNode node in values)
            {
                var match = ChartGroupFieldRegex.Match((GetNodeTextValue(node) ?? string.Empty).Trim());
                var column = match.Success
                    ? GetChartColumnName(match.Groups["field"].Value, masterSource, dataset)
                    : null;
                var key = $"Min|{match.Groups["field"].Value}";
                if (column == null || groupColumns.Contains(column) || aggregates.ContainsKey(key))
                    continue;

                var type = masterSource.Columns[column].Type;
                var function = type == typeof(string) ? "MinStr" : type == typeof(DateTime) ? "MinDate" : "Min";

                var resultName = "Min" + column;
                for (var index = 2; masterSource.Columns[resultName] != null ||
                    resultColumns.Any(c => c.Name == resultName); index++)
                {
                    resultName = "Min" + column + index;
                }

                aggregates[key] = resultName;
                resultColumns.Add(new StiDataColumn(resultName, type));
                results.Add(column);
                results.Add(function);
                results.Add(resultName);
            }
        }

        private static string NormalizeChartAggregate(string function)
        {
            foreach (var name in new[] { "Sum", "CountDistinct", "Count", "Avg", "Min", "Max", "First", "Last" })
            {
                if (name.Equals(function, StringComparison.OrdinalIgnoreCase))
                    return name;
            }

            return function;
        }

        private string GetChartColumnName(string field, StiDataSource masterSource, string dataset)
        {
            var reference = FindFieldReference(field, dataset);
            if (reference == null)
                return null;

            var column = reference.Substring(reference.LastIndexOf('.') + 1);
            if (masterSource.Columns[column] != null)
                return column;

            //the reference to a field named as a keyword, as Date, is escaped, Date_; the column keeps the name
            return masterSource.Columns.Cast<StiDataColumn>()
                .FirstOrDefault(c => GetColumnReference(dataset, c.Name) == reference)?.Name;
        }

        /// <summary>
        /// Replaces the aggregates of fields in an expression of the chart by the fields of their totals in the
        /// source of the chart, and an expression of a group by the field of the group.
        /// </summary>
        private string ReplaceChartAggregates(string expression, string dataset)
        {
            if (chartAggregates == null || dataset != chartAggregatesSource || !IsExpression(expression))
                return expression;

            if (chartGroupExpressions.TryGetValue(NormalizeChartGroupExpression(expression), out var column))
                return $"=Fields!{column}.Value";

            return ChartAggregateRegex.Replace(expression, ReplaceChartAggregate);
        }

        private string ReplaceChartAggregate(Match match)
        {
            var key = $"{NormalizeChartAggregate(match.Groups["function"].Value)}|{match.Groups["field"].Value}";

            return chartAggregates.TryGetValue(key, out var name) ? $"Fields!{name}.Value" : match.Value;
        }
        #endregion
    }
}
