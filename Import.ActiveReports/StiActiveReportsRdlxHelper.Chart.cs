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
    /// Stimulsoft shows a point per row. So the chart reads a data transformation that groups the rows of
    /// the data source by the category and the series and counts the totals when the report is built.
    /// The filters of the chart are the filters of a virtual source the transformation reads.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Consts.Chart
        private const string ChartCategoryColumn = "Category";
        private const string ChartSeriesColumn = "Series";
        private const string ChartValueColumn = "Value";
        private const string ChartKeyColumn = "Key";
        #endregion

        #region Fields.Chart
        private Dictionary<StiChart, StiRdlxChartData> chartData = new Dictionary<StiChart, StiRdlxChartData>();
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

                columns = CreateChartTransformation(chart, node, dataSet, chartScope);

                // The transformation counts the filtered rows, the chart needs no filters.
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
                Warn($"The values of the chart '{node.Get("Name")}' are not grouped by the category: its plots " +
                    "have different categories or its expressions are too complex.");
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

                if (series is StiPieSeries && !string.IsNullOrEmpty(category))
                    SetLegendArguments(chart, series, plot, value.Value);

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

        /// <summary>
        /// The legend of a pie names its slices: an item of the legend shows the category of the slice, and
        /// the value too when the text of the legend of the plot has the value in it.
        /// </summary>
        private static void SetLegendArguments(StiChart chart, StiSeries series, StiRdlxNode plot, string value)
        {
            var color = plot.Child("ColorSeries");
            var text = color == null ? "" : color.Get("Field", "");
            var field = Regex.Match(value ?? "", @"Fields!\w+\.Value|^\{\w+\}$").Value;
            var hasValue = field.Length > 0 && text.IndexOf(field, StringComparison.OrdinalIgnoreCase) >= 0;
            var type = hasValue ? StiSeriesLabelsValueType.ArgumentValue : StiSeriesLabelsValueType.Argument;

            foreach (var labels in new[] { chart.SeriesLabels, series.SeriesLabels })
            {
                if (labels == null)
                    continue;

                labels.LegendValueType = type;
                labels.ValueTypeSeparator = " ";
            }
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
        /// Collects the converted expressions of the category, of the series and of the values of a chart.
        /// Returns false when they can not be put into one table: a plot has no category or is a plot of
        /// points, or the plots have different categories.
        /// </summary>
        private bool CollectChartExpressions(StiRdlxNode node, StiRdlxScope scope, out string category,
            out string seriesKey, List<string> values)
        {
            category = null;
            seriesKey = "";

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
                    return false;

                plotCategory = StripTotal(ConvertToExpression(plotCategory, scope));
                plotKey = string.IsNullOrEmpty(plotKey) ? "" : StripTotal(ConvertToExpression(plotKey, scope));

                if (category != null && (category != plotCategory || seriesKey != plotKey))
                    return false;

                category = plotCategory;
                seriesKey = plotKey;

                foreach (var value in GetPlotValues(plot, encodings))
                {
                    var expression = ConvertToExpression(value.Value, scope);
                    if (!values.Contains(expression))
                        values.Add(expression);
                }
            }

            return category != null && values.Count > 0;
        }

        /// <summary>
        /// The totals of a chart are counted when the report is built: a data transformation groups the rows
        /// of the data source by the category and the series and counts the values. Returns the columns of
        /// the transformation by the expressions they hold, null when it can not be made.
        /// </summary>
        private Dictionary<string, string> CreateChartTransformation(StiChart chart, StiRdlxNode node,
            StiRdlxDataSet dataSet, StiRdlxScope scope)
        {
            if (dataSet.Source == null)
                return null;

            string category;
            string seriesKey;
            var values = new List<string>();

            if (!CollectChartExpressions(node, scope, out category, out seriesKey, values))
                return null;

            var prefix = dataSet.SourceName + ".";
            var categoryExpression = ToDataExpression(category, prefix);
            var seriesExpression = seriesKey.Length == 0 ? "" : ToDataExpression(seriesKey, prefix);

            if (categoryExpression == null || seriesExpression == null)
                return null;

            var measures = new List<string>();

            foreach (var value in values)
            {
                var measure = ToDataExpression(value == "Count()" ? "Count(" + category + ")" : value, prefix);
                if (measure == null)
                    return null;

                // A value without a total function is summed, as a chart of ActiveReports does.
                var hasTotal = Regex.IsMatch(measure, @"\b(Sum|Avg|Count|Min|Max)\(");
                measures.Add(hasTotal ? measure : "Sum(" + measure + ")");
            }

            // The filters of the chart are applied by a virtual source, the transformation reads its rows.
            var rowsName = CreateChartRows(node, dataSet, scope);

            if (rowsName != null)
            {
                var pattern = @"\b" + Regex.Escape(prefix);

                categoryExpression = Regex.Replace(categoryExpression, pattern, rowsName + ".");
                seriesExpression = Regex.Replace(seriesExpression, pattern, rowsName + ".");

                for (var index = 0; index < measures.Count; index++)
                {
                    measures[index] = Regex.Replace(measures[index], pattern, rowsName + ".");
                }
            }

            var name = CreateName(node.Get("Name", "Chart") + "Data");
            var transformation = new StiDataTransformation(name, name);
            var columns = new Dictionary<string, string>(StringComparer.Ordinal);

            AddTransformationColumn(transformation, ChartCategoryColumn, categoryExpression,
                StiDataTransformationMode.Dimension);
            columns[category] = name + "." + ChartCategoryColumn;

            if (seriesExpression.Length > 0)
            {
                AddTransformationColumn(transformation, ChartSeriesColumn, seriesExpression,
                    StiDataTransformationMode.Dimension);
                columns[seriesKey] = name + "." + ChartSeriesColumn;
            }

            for (var index = 0; index < values.Count; index++)
            {
                var column = ChartValueColumn + (index + 1);

                AddTransformationColumn(transformation, column, measures[index], StiDataTransformationMode.Measure);
                columns[values[index]] = name + "." + column;
            }

            report.Dictionary.DataSources.Add(transformation);
            chart.DataSourceName = name;
            chartData[chart] = new StiRdlxChartData(transformation, rowsName ?? dataSet.SourceName, dataSet);

            return columns;
        }

        /// <summary>
        /// Gives the charts that stand in a group the rows of the group. A chart of ActiveReports placed in
        /// a group of a data region counts the rows of the current group only; a data transformation counts
        /// all the rows of its data source. So the key of the group becomes one more dimension of the
        /// transformation, and the chart takes the rows with the key of the current row of the data set,
        /// which stands on the first row of the group when the band of the group is printed. It is called
        /// when all the bands stand on their places.
        /// </summary>
        private void ResolveChartScopes()
        {
            foreach (var pair in chartData)
            {
                var chart = pair.Key;
                var data = pair.Value;
                var keys = new List<string>();

                for (var parent = chart.Parent; parent != null; parent = parent.Parent)
                {
                    var band = parent as StiBand;
                    if (band != null)
                        CollectGroupKeys(chart, band, data.DataSet, keys);
                }

                var columns = new List<string>();

                // A key is the column of the data source the group reads: the data set or its copy that
                // shows the rows of an outer group. The nearest group of a column is taken.
                foreach (var key in keys)
                {
                    var keyColumn = key.Substring(key.IndexOf('.') + 1);
                    if (columns.Contains(keyColumn))
                        continue;

                    columns.Add(keyColumn);
                    var column = ChartKeyColumn + columns.Count;

                    data.Transformation.Columns.Add(new StiDataTransformationColumn(column, column,
                        GetColumnType(data.DataSet, keyColumn), data.RowsName + "." + keyColumn)
                    {
                        Mode = StiDataTransformationMode.Dimension
                    });

                    chart.Filters.Add(new StiFilter("{" + data.Transformation.Name + "." + column + " == " + key + "}"));
                    chart.FilterOn = true;
                }
            }
        }

        /// <summary>
        /// Adds the columns of the conditions of the groups a band is printed in: of the group of the band
        /// itself and of the groups above it. A data band is printed in all the groups that stand before it.
        /// </summary>
        private void CollectGroupKeys(StiChart chart, StiBand band, StiRdlxDataSet dataSet, List<string> keys)
        {
            var isGroupBand = band is StiGroupHeaderBand || band is StiGroupFooterBand || band is StiDataBand;
            if (band.Parent == null || !isGroupBand)
                return;

            var bands = new List<StiBand>();

            foreach (StiComponent component in band.Parent.Components)
            {
                if (component is StiBand)
                    bands.Add((StiBand)component);
            }

            bands.Sort(CompareBandTops);

            // The headers of the groups stand one after another before the data band, the outer group first;
            // the footers stand after it in the reverse order.
            var position = bands.IndexOf(band);
            var footers = 0;

            while (position >= 0 && bands[position] is StiGroupFooterBand)
            {
                footers++;
                position--;
            }

            if (footers > 0 && position >= 0 && bands[position] is StiDataBand)
                position--;

            else if (footers == 0 && band is StiDataBand)
                position--;

            var headers = new List<StiGroupHeaderBand>();

            while (position >= 0 && bands[position] is StiGroupHeaderBand)
            {
                headers.Insert(0, (StiGroupHeaderBand)bands[position]);
                position--;
            }

            // The first footer after the data band closes the last header.
            var count = footers > 0 ? headers.Count - footers + 1 : headers.Count;

            for (var index = 0; index < count && index < headers.Count; index++)
            {
                var condition = headers[index].Condition.Value ?? "";
                var match = Regex.Match(condition, @"^\{(\w+)\.(\w+)\}$");

                // A group without an expression holds all the rows, it has no key.
                if (condition == "{\"\"}")
                    continue;

                if (!match.Success || !IsSourceOfDataSet(match.Groups[1].Value, dataSet))
                {
                    Warn($"The chart '{chart.Name}' stands in a group whose condition is not a field of its data " +
                        "set, it shows the rows of all the groups.");
                    continue;
                }

                var key = match.Groups[1].Value + "." + match.Groups[2].Value;
                if (!keys.Contains(key))
                    keys.Add(key);
            }
        }

        /// <summary>
        /// The data source is the one of the data set or of its copy that shows the rows of a group.
        /// </summary>
        private bool IsSourceOfDataSet(string sourceName, StiRdlxDataSet dataSet)
        {
            if (sourceName == dataSet.SourceName)
                return true;

            foreach (var copy in groupRowSets)
            {
                if (copy.SourceName == sourceName && copy.Original == (dataSet.Original ?? dataSet))
                    return true;
            }

            return false;
        }

        private static int CompareBandTops(StiBand first, StiBand second)
        {
            return first.Top.CompareTo(second.Top);
        }

        /// <summary>
        /// Makes a virtual source that leaves the rows passing the filters of a chart. Returns its name,
        /// null when the chart has no filters.
        /// </summary>
        private string CreateChartRows(StiRdlxNode node, StiRdlxDataSet dataSet, StiRdlxScope scope)
        {
            var filters = new List<string>();

            foreach (var filter in node.Items("Filters"))
            {
                var expression = ConvertFilter(filter, scope);
                if (expression != null)
                    filters.Add(expression);
            }

            if (filters.Count == 0)
                return null;

            var name = CreateName(node.Get("Name", "Chart") + "Rows");
            var source = new StiVirtualSource(dataSet.SourceName, name);
            var results = new List<string>();

            foreach (StiDataColumn column in dataSet.Source.Columns)
            {
                source.Columns.Add(new StiDataColumn(column.Name, column.Name, column.Alias, column.Type));
                results.Add(column.Name);
                results.Add(string.Empty);
                results.Add(column.Name);
            }

            foreach (var filter in filters)
            {
                source.Filters.Add(new StiFilter("{" + filter + "}"));
            }

            source.Results = results.ToArray();
            source.FilterOn = true;
            report.Dictionary.DataSources.Add(source);

            return name;
        }

        private static void AddTransformationColumn(StiDataTransformation transformation, string name,
            string expression, StiDataTransformationMode mode)
        {
            var type = mode == StiDataTransformationMode.Measure ? typeof(decimal) : typeof(object);

            transformation.Columns.Add(new StiDataTransformationColumn(name, name, type, expression)
            {
                Mode = mode
            });
        }

        /// <summary>
        /// Makes an expression of a data transformation of a converted expression. Returns null when the
        /// expression has more than the columns of the data source, numbers, texts, arithmetic and the
        /// functions that the transformations have too.
        /// </summary>
        private static string ToDataExpression(string expression, string prefix)
        {
            // The functions of the transformations take a date written as a text.
            var result = Regex.Replace(expression, @"DateTime\.Parse\(ToString\((\w+\.\w+)\)\)", "$1");

            var rest = Regex.Replace(result, @"""[^""]*""", "");
            var columns = Regex.Replace(rest, Regex.Escape(prefix) + @"\w+", "");

            if (columns == rest)
                return null;

            rest = Regex.Replace(columns, @"\b(Sum|Avg|Count|Min|Max|Year|Month|Day|MonthName|Round|Abs)\(", "(");

            return Regex.IsMatch(rest, @"^[\d\s.+\-*/(),]*$") ? result : null;
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
