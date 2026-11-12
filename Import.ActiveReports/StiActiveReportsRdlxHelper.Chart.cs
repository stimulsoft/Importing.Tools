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
using System.Text.RegularExpressions;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Base.Json.Linq;
using Stimulsoft.Report.Chart;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The charts. There are two models of a chart in RDLX: the current one, where a plot has DataSeries,
    /// CategoryGroupings and SeriesGroupings, and the older one (dvchart), where a plot has Encodings with
    /// Values, Category and Details. Both give a chart with a series per value of the plot.
    /// A chart of ActiveReports groups the rows by the category and shows a total of each group, a chart of
    /// Stimulsoft shows a point per row. So when the data of the chart is known at the import, the totals
    /// are counted here and written into a table of the report data, the chart reads this table.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Consts.Chart
        private const string ChartCategoryColumn = "Category";
        private const string ChartSeriesColumn = "Series";
        private const string ChartValueColumn = "Value";
        #endregion

        #region Methods.Chart
        private StiComponent CreateChart(StiRdlxNode node, RectangleD rect, StiRdlxScope scope)
        {
            var chart = new StiChart(rect);
            var dataSet = GetRegionDataSet(node, scope);
            var chartScope = new StiRdlxScope { DataSet = dataSet, IsRegion = true };
            Dictionary<string, string> columns = null;

            if (dataSet != null)
            {
                chart.DataSourceName = dataSet.SourceName;

                foreach (var filter in node.Items("Filters"))
                {
                    var expression = ConvertFilter(filter, chartScope);
                    if (expression == null)
                        continue;

                    chart.Filters.Add(new StiFilter("{" + expression + "}"));
                    chart.FilterOn = true;
                }

                columns = CreateChartTable(chart, node, dataSet, chartScope);

                // The table of the chart is counted of the filtered rows, the chart needs no filters.
                if (columns != null)
                {
                    chart.Filters.Clear();
                    chart.FilterOn = false;
                }
            }

            var style = node.Child("Style");
            if (style != null)
            {
                chart.Brush = GetBrush(style);
                chart.Border = GetBorder(style);
            }

            chart.Series.Clear();

            foreach (var plot in node.Items("Plots"))
            {
                AddPlotSeries(chart, node, plot, chartScope, columns);
            }

            if (chart.Series.Count == 0)
            {
                Warn($"The chart '{node.Get("Name")}' has no series that could be imported.");
            }
            else if (columns == null && dataSet != null)
            {
                Warn($"The values of the chart '{node.Get("Name")}' are not grouped by the category: its data " +
                    "is not known at the import or its expressions are not plain fields.");
            }

            var header = node.Child("Header");
            var title = header == null ? null : header.Get("Title");

            if (!string.IsNullOrWhiteSpace(title))
            {
                chart.Title.Text = ConvertValue(title, chartScope);
                chart.Title.Visible = true;
            }

            ApplyAxisTitles(chart, node);

            return chart;
        }

        private static bool IsPointPlot(StiRdlxNode plot)
        {
            var type = (plot.Get("PlotType") ?? plot.Get("PlotChartType") ?? "").ToLowerInvariant();
            return type == "scatter" || type == "bubble" || type == "radarscatter";
        }

        private void AddPlotSeries(StiChart chart, StiRdlxNode node, StiRdlxNode plot, StiRdlxScope scope,
            Dictionary<string, string> columns)
        {
            var encodings = plot.Child("Encodings");
            var category = GetPlotCategory(plot, encodings);
            var seriesKey = GetPlotSeriesKey(plot, encodings);
            var values = GetPlotValues(plot, encodings);
            var isPoints = IsPointPlot(plot);

            // A plot without categories, a pie as a rule, is split by its series only: they are its categories.
            if (string.IsNullOrEmpty(category) && !isPoints)
            {
                category = seriesKey;
                seriesKey = null;
            }

            var key = string.IsNullOrEmpty(seriesKey)
                ? ""
                : GetChartColumn(columns, StripTotal(ConvertToExpression(seriesKey, scope)));

            // A plot of points has no categories: its first value is the X of a point, the second is the Y
            // and the third, of a bubble, is the size.
            if (isPoints && values.Count >= 2 && string.IsNullOrEmpty(category))
            {
                var points = CreateSeries(node, plot);

                SetSeriesData(points, StripTotal(ConvertToExpression(values[0].Value, scope)), isValue: false);
                SetSeriesData(points, StripTotal(ConvertToExpression(values[1].Value, scope)), isValue: true);

                var bubble = points as StiBubbleSeries;
                var weight = values.Count > 2 ? StripTotal(ConvertToExpression(values[2].Value, scope)) : "";

                if (bubble != null && IsColumn(weight))
                    bubble.WeightDataColumn = weight;

                AddSeries(chart, points, values[1].Name, key);
                return;
            }

            foreach (var value in values)
            {
                var series = CreateSeries(node, plot);

                // Points without the X can not be drawn, a line shows the values by their order.
                if (isPoints && string.IsNullOrEmpty(category))
                    series = new StiLineSeries();

                var expression = ConvertToExpression(value.Value, scope);
                string column;

                if (columns != null && columns.TryGetValue(expression, out column))
                    SetSeriesData(series, column, isValue: true);

                else
                    SetSeriesData(series, StripTotal(expression), isValue: true);

                if (!string.IsNullOrEmpty(category))
                {
                    SetSeriesData(series, GetChartColumn(columns, StripTotal(ConvertToExpression(category, scope))),
                        isValue: false);
                }

                AddSeries(chart, series, value.Name, key);
            }
        }

        /// <summary>
        /// Returns the column of the table of the chart that holds the expression, else the expression.
        /// </summary>
        private static string GetChartColumn(Dictionary<string, string> columns, string expression)
        {
            string column;
            return columns != null && columns.TryGetValue(expression, out column) ? column : expression;
        }

        private void AddSeries(StiChart chart, StiSeries series, string name, string key)
        {
            if (IsColumn(key))
            {
                series.AutoSeriesKeyDataColumn = key;
                series.AutoSeriesTitleDataColumn = key;
            }

            if (!string.IsNullOrEmpty(name) && !name.StartsWith("=") && name.IndexOf('{') < 0)
                series.CoreTitle = name;

            if (chart.Series.Count == 0)
                chart.Area = (IStiArea)Activator.CreateInstance(series.GetDefaultAreaType());

            chart.Series.Add(series);
        }

        /// <summary>
        /// Counts the totals of the values of a chart by its categories and series and writes them into
        /// a table of the report data. Returns the columns of this table by the expressions they hold, null
        /// when the totals can not be counted: the rows are not known, the chart is filtered, its plots have
        /// different categories or an expression is not a plain field.
        /// </summary>
        private Dictionary<string, string> CreateChartTable(StiChart chart, StiRdlxNode node, StiRdlxDataSet dataSet,
            StiRdlxScope scope)
        {
            var table = embeddedData[dataSet.SourceName] as JArray;
            var filters = node.Items("Filters");

            if (table == null || dataSet.Node.Items("Filters").Count > 0 && !dataSet.IsFiltered ||
                filters.Count > 0 && !AreStaticFilters(filters, dataSet))
                return null;

            var tableRows = new List<JObject>();

            foreach (JObject row in table)
            {
                tableRows.Add(row);
            }

            var rows = FilterRows(tableRows, filters, dataSet, isTable: true);

            string category = null;
            var seriesKey = "";
            var values = new List<string>();

            foreach (var plot in node.Items("Plots"))
            {
                var encodings = plot.Child("Encodings");
                var plotCategory = GetPlotCategory(plot, encodings);
                var plotKey = GetPlotSeriesKey(plot, encodings);

                if (string.IsNullOrEmpty(plotCategory))
                {
                    plotCategory = plotKey;
                    plotKey = null;
                }

                if (IsPointPlot(plot) || string.IsNullOrEmpty(plotCategory))
                    return null;

                plotCategory = StripTotal(ConvertToExpression(plotCategory, scope));
                plotKey = string.IsNullOrEmpty(plotKey) ? "" : StripTotal(ConvertToExpression(plotKey, scope));

                if (category != null && (category != plotCategory || seriesKey != plotKey))
                    return null;

                category = plotCategory;
                seriesKey = plotKey;

                foreach (var value in GetPlotValues(plot, encodings))
                {
                    var expression = ConvertToExpression(value.Value, scope);
                    if (!values.Contains(expression))
                        values.Add(expression);
                }
            }

            var prefix = dataSet.SourceName + ".";
            var functions = new string[values.Count];
            var valueColumns = new string[values.Count];

            for (var index = 0; index < values.Count; index++)
            {
                var match = Regex.Match(values[index], @"^(?:(Sum|Avg|Count|Min|Max)\()?(\w+\.\w+)\)?$");
                if (!match.Success || !match.Groups[2].Value.StartsWith(prefix))
                    return null;

                functions[index] = match.Groups[1].Success ? match.Groups[1].Value : "Sum";
                valueColumns[index] = match.Groups[2].Value.Substring(prefix.Length);
            }

            if (category == null || values.Count == 0 || !IsColumn(category) || !category.StartsWith(prefix) ||
                seriesKey.Length > 0 && (!IsColumn(seriesKey) || !seriesKey.StartsWith(prefix)))
                return null;

            var categoryColumn = category.Substring(prefix.Length);
            var seriesColumn = seriesKey.Length > 0 ? seriesKey.Substring(prefix.Length) : null;
            var groups = new Dictionary<string, JObject>(StringComparer.Ordinal);
            var totals = new Dictionary<string, double[]>(StringComparer.Ordinal);
            var counts = new Dictionary<string, int[]>(StringComparer.Ordinal);
            var chartTable = new JArray();

            foreach (var row in rows)
            {
                var key = row[categoryColumn] + "|#|" + (seriesColumn == null ? null : row[seriesColumn]);

                if (!groups.ContainsKey(key))
                {
                    var group = new JObject();
                    group[ChartCategoryColumn] = row[categoryColumn] == null ? null : row[categoryColumn].DeepClone();

                    if (seriesColumn != null)
                        group[ChartSeriesColumn] = row[seriesColumn] == null ? null : row[seriesColumn].DeepClone();

                    groups[key] = group;
                    totals[key] = new double[values.Count];
                    counts[key] = new int[values.Count];
                    chartTable.Add(group);
                }

                for (var index = 0; index < values.Count; index++)
                {
                    AddChartValue(row[valueColumns[index]], functions[index], totals[key], counts[key], index);
                }
            }

            foreach (var pair in groups)
            {
                for (var index = 0; index < values.Count; index++)
                {
                    var count = counts[pair.Key][index];
                    var total = totals[pair.Key][index];

                    if (functions[index] == "Count")
                        total = count;

                    else if (functions[index] == "Avg" && count > 0)
                        total /= count;

                    pair.Value[ChartValueColumn + (index + 1)] = total;
                }
            }

            var name = CreateName(node.Get("Name", "Chart") + "Data");
            embeddedData[name] = chartTable;

            var source = new StiDataTableSource(EmbeddedDataName + "." + name, name, name);
            var columns = new Dictionary<string, string>(StringComparer.Ordinal);

            source.Columns.Add(new StiDataColumn(ChartCategoryColumn, ChartCategoryColumn, ChartCategoryColumn,
                GetColumnType(dataSet, categoryColumn)));
            columns[category] = name + "." + ChartCategoryColumn;

            if (seriesColumn != null)
            {
                source.Columns.Add(new StiDataColumn(ChartSeriesColumn, ChartSeriesColumn, ChartSeriesColumn,
                    GetColumnType(dataSet, seriesColumn)));
                columns[seriesKey] = name + "." + ChartSeriesColumn;
            }

            for (var index = 0; index < values.Count; index++)
            {
                var column = ChartValueColumn + (index + 1);
                source.Columns.Add(new StiDataColumn(column, column, column, typeof(decimal)));
                columns[values[index]] = name + "." + column;
            }

            report.Dictionary.DataSources.Add(source);
            chart.DataSourceName = name;

            return columns;
        }

        /// <summary>
        /// Adds a value of a row to the total of its group.
        /// </summary>
        private static void AddChartValue(JToken token, string function, double[] totals, int[] counts, int index)
        {
            if (token == null || token.Type == JTokenType.Null)
                return;

            var number = 0d;

            if (token.Type == JTokenType.Integer || token.Type == JTokenType.Float)
                number = token.Value<double>();

            else if (!double.TryParse(token.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out number) &&
                function != "Count")
                return;

            var isFirst = counts[index] == 0;
            counts[index]++;

            switch (function)
            {
                case "Min":
                    totals[index] = isFirst ? number : Math.Min(totals[index], number);
                    break;

                case "Max":
                    totals[index] = isFirst ? number : Math.Max(totals[index], number);
                    break;

                default:
                    totals[index] += number;
                    break;
            }
        }

        private static Type GetColumnType(StiRdlxDataSet dataSet, string column)
        {
            var dataColumn = dataSet.Source == null ? null : dataSet.Source.Columns[column];
            return dataColumn == null ? typeof(string) : dataColumn.Type;
        }

        /// <summary>
        /// The expression of the categories of a plot.
        /// </summary>
        private static string GetPlotCategory(StiRdlxNode plot, StiRdlxNode encodings)
        {
            foreach (var grouping in plot.Items("CategoryGroupings"))
            {
                var expression = GetGroupExpression(grouping.Child("Grouping", "Group"));
                if (expression != null)
                    return expression;
            }

            return encodings == null ? null : GetEncodingField(encodings.Child("Category"));
        }

        /// <summary>
        /// The expression that splits the values of a plot into series.
        /// </summary>
        private static string GetPlotSeriesKey(StiRdlxNode plot, StiRdlxNode encodings)
        {
            foreach (var grouping in plot.Items("SeriesGroupings"))
            {
                var expression = GetGroupExpression(grouping.Child("Grouping", "Group"));
                if (expression != null)
                    return expression;
            }

            if (encodings == null)
                return null;

            foreach (var detail in encodings.Items("Details"))
            {
                var expression = GetEncodingField(detail);
                if (expression != null)
                    return expression;
            }

            return GetEncodingField(encodings.Child("Color"));
        }

        private static string GetEncodingField(StiRdlxNode encoding)
        {
            var field = encoding == null ? null : encoding.Child("Field");
            if (field == null)
                return null;

            if (!string.IsNullOrEmpty(field.Value))
                return field.Value;

            foreach (var value in field.Items("Value"))
            {
                if (!string.IsNullOrEmpty(value.Value))
                    return value.Value;
            }

            return null;
        }

        /// <summary>
        /// The values of a plot: the name of a returned node is the title of the series, its value is the
        /// expression.
        /// </summary>
        private static List<StiRdlxNode> GetPlotValues(StiRdlxNode plot, StiRdlxNode encodings)
        {
            var values = new List<StiRdlxNode>();

            foreach (var series in plot.Items("DataSeries"))
            {
                foreach (var value in series.Items("Values"))
                {
                    var expression = value.Get("Value") ?? GetEncodingField(value);
                    if (!string.IsNullOrEmpty(expression))
                        values.Add(new StiRdlxNode(series.Get("Name", "")) { Value = expression });
                }
            }

            if (values.Count > 0 || encodings == null)
                return values;

            foreach (var value in encodings.Items("Values"))
            {
                var expression = GetEncodingField(value);
                if (!string.IsNullOrEmpty(expression))
                    values.Add(new StiRdlxNode("") { Value = expression });
            }

            return values;
        }

        private StiSeries CreateSeries(StiRdlxNode node, StiRdlxNode plot)
        {
            var config = plot.Child("Config") ?? new StiRdlxNode("Config");
            var type = plot.Get("PlotType") ?? plot.Get("PlotChartType") ?? "Column";
            var isOld = !plot.Has("PlotType");
            var subtype = plot.Get("PlotChartSubtype");
            var isStacked = string.Equals(subtype, "Stacked", StringComparison.OrdinalIgnoreCase) ||
                config.Get("SeriesArrangement", "").StartsWith("Stack", StringComparison.OrdinalIgnoreCase);
            var isSwapped = config.GetBool("SwapAxes");

            switch (type.ToLowerInvariant())
            {
                case "column":
                    if (isStacked)
                        return new StiStackedColumnSeries();

                    return new StiClusteredColumnSeries();

                case "bar":
                    var isHorizontal = isOld || isSwapped;

                    if (isHorizontal && isStacked)
                        return new StiStackedBarSeries();

                    if (isHorizontal)
                        return new StiClusteredBarSeries();

                    if (isStacked)
                        return new StiStackedColumnSeries();

                    return new StiClusteredColumnSeries();

                case "line":
                    if (string.Equals(config.Get("LineAspect"), "Spline", StringComparison.OrdinalIgnoreCase))
                        return new StiSplineSeries();

                    return new StiLineSeries();

                case "area":
                    if (isStacked)
                        return new StiStackedAreaSeries();

                    return new StiAreaSeries();

                case "pie":
                    double radius;
                    double.TryParse(config.Get("InnerRadius"), NumberStyles.Float, CultureInfo.InvariantCulture,
                        out radius);

                    if (radius > 0)
                        return new StiDoughnutSeries();

                    return new StiPieSeries();

                case "scatter":
                    return new StiScatterSeries();

                case "bubble":
                    return new StiBubbleSeries();

                case "radarline":
                    return new StiRadarLineSeries();

                case "radararea":
                    return new StiRadarAreaSeries();

                case "radarscatter":
                    return new StiRadarPointSeries();

                default:
                    Warn($"The chart '{node.Get("Name")}' has a plot of the type '{type}' that is not imported, " +
                        "columns are shown instead.");
                    return new StiClusteredColumnSeries();
            }
        }

        /// <summary>
        /// Sets the values or the arguments of a series: a column of the data source or an expression.
        /// </summary>
        private static void SetSeriesData(StiSeries series, string expression, bool isValue)
        {
            if (isValue && IsColumn(expression))
                series.ValueDataColumn = expression;

            else if (isValue)
                series.Value.Value = "{" + expression + "}";

            else if (IsColumn(expression))
                series.ArgumentDataColumn = expression;

            else
                series.Argument.Value = "{" + expression + "}";
        }

        private static bool IsColumn(string expression)
        {
            return Regex.IsMatch(expression ?? "", @"^\w+\.\w+$");
        }

        /// <summary>
        /// Takes the value out of a total function: a series gets the values row by row.
        /// </summary>
        private static string StripTotal(string expression)
        {
            var match = Regex.Match(expression, @"^(Sum|Avg|Min|Max|Count|Last|Median)\((.+)\)$");
            return match.Success && match.Groups[2].Value.IndexOf('(') < 0 ? match.Groups[2].Value : expression;
        }

        private void ApplyAxisTitles(StiChart chart, StiRdlxNode node)
        {
            var area = chart.Area as IStiAxisArea;
            var plotArea = node.Child("PlotArea");
            if (area == null || plotArea == null)
                return;

            foreach (var axis in plotArea.Items("Axes"))
            {
                var title = axis.Get("Title");
                if (string.IsNullOrWhiteSpace(title) || title.StartsWith("="))
                    continue;

                if (string.Equals(axis.Get("AxisType"), "Y", StringComparison.OrdinalIgnoreCase))
                    area.YAxis.Title.Text = title;

                else
                    area.XAxis.Title.Text = title;
            }
        }
        #endregion
    }
}
