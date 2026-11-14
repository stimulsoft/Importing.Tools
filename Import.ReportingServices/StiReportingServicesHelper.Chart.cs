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

using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Chart;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.TextFormats;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Units;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
        #region Fields
        private HashSet<string> secondaryValueAxisNames;
        #endregion

        #region Chart
        #region Fields.Chart
        private static readonly Regex ChartLabelKeywordRegex =
            new Regex(@"#(?<keyword>[A-Za-z]+)(\{(?<format>[^}]*)\})?");
        private static readonly Regex ChartAxisExtremeRegex = new Regex(
            @"^=\s*(?<function>Max|Min)\s*\(\s*(?<inner>.+)\s*,\s*""(?<scope>[^""]+)""\s*\)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);
        private string defaultChartType;
        private string defaultChartSubtype;
        private List<string> chartCategoryLevels;

        /// <summary>
        /// The limits of the axes of the chart being converted given by expressions: the axis, the limit and the
        /// expression, see ApplyChartAxisLimits.
        /// </summary>
        private List<Tuple<IStiAxis, string, string>> chartAxisLimits = new List<Tuple<IStiAxis, string, string>>();
        #endregion

        private void ProcessChartType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            var component = new StiChart();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;

            //a chart of Reporting Services has no spacing around its areas, the areas start at its edges
            component.HorSpacing = 0;
            component.VertSpacing = 0;

            //a chart without a legend node has no legend, the node turns it on, see ProcessChartLegendType
            component.Legend.Visible = false;

            var datasetName = dataset;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSetName":
                        datasetName = MakeFullDataSetName(dataset, GetNodeTextValue(node));
                        break;
                }
            }
            ApplyChartPalette(baseNode, component);

            //the points of a chart are the groups of its data, see PrepareChartData; in a cell of a data region
            //only of the rows of the groups of the cell, see ResolveChartScopes
            var chartDataName = PrepareChartData(baseNode, component, datasetName, !string.IsNullOrEmpty(dataset));
            if (chartDataName != null && !string.IsNullOrEmpty(dataset))
                regionCharts[component] = datasetName;

            var dataName = chartDataName ?? datasetName;
            component.DataSourceName = dataName;

            var categoryExpression = string.Empty;
            var seriesExpression = string.Empty;
            var categorySort = new List<string>();
            var seriesSort = new List<string>();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartCategoryHierarchy":
                    case "CategoryGroupings":
                        categoryExpression = ProcessChartHierarchyType(node, dataName);
                        chartCategoryLevels = new List<string>();
                        CollectChartCategoryLevels(node, dataName, chartCategoryLevels);
                        CollectChartGroupSort(node, dataName, categorySort);
                        break;

                    case "ChartSeriesHierarchy":
                    case "SeriesGroupings":
                        seriesExpression = ProcessChartSeriesHierarchyType(node, dataName);
                        CollectChartGroupSort(node, dataName, seriesSort);
                        break;
                }
            }

            //the series groups are sorted first, so the auto series come in their order with the points of each
            //in the order of the categories
            if (categorySort.Count > 0 || seriesSort.Count > 0)
                component.Sort = seriesSort.Concat(categorySort).ToArray();

            var chartAreas3D = GetChartAreas3D(baseNode);
            secondaryValueAxisNames = GetSecondaryValueAxisNames(baseNode);

            //the schema of 2005 gives the type to the whole chart, its series have none
            defaultChartType = GetChildText(baseNode, "Type");
            defaultChartSubtype = GetChildText(baseNode, "Subtype");

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
                        component.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
                        component.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
                        component.Width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        component.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, component, datasetName);
                        break;

                    case "ChartData":
                        ProcessChartDataType(node, component, dataName, categoryExpression, seriesExpression,
                            chartAreas3D);
                        break;

                    //the filters of a grouped chart take the rows of the data set before the groups
                    case "Filters":
                        if (chartDataName != null)
                        {
                            var chartSource = (StiVirtualSource)report.Dictionary.DataSources[chartDataName];
                            chartSource.Filters = ProcessFiltersType(node, datasetName);
                        }
                        else
                        {
                            component.Filters = ProcessFiltersType(node, datasetName);
                            component.FilterOn = component.Filters != null && component.Filters.Count > 0;
                        }
                        break;

                    case "ChartAreas":
                        ProcessChartAreasType(node, component);
                        break;

                    case "ChartLegends":
                        ProcessChartLegendsType(node, component);
                        break;

                    case "ChartTitles":
                        ProcessChartTitlesType(node, component, datasetName);
                        break;

                    case "Visibility":
                        ProcessVisibility(node, component, datasetName);
                        break;

                    case "CustomProperties":
                        ProcessChartCustomProperties(node);
                        break;

                    case "DataSetName":
                    case "ChartCategoryHierarchy":
                    case "CategoryGroupings":
                    case "ChartSeriesHierarchy":
                    case "SeriesGroupings":
                    case "Palette":
                    case "ChartCustomPaletteColors":
                    case "ChartBorderSkin":
                    case "ChartNoDataMessage":
                    case "NoRowsMessage":
                    case "DynamicHeight":
                    case "DynamicWidth":
                    case "DataElementName":
                    case "DataElementOutput":
                    case "ZIndex":
                        break;

                    case "Type":
                    case "Subtype":
                    case "Legend":
                    case "Title":
                    case "CategoryAxis":
                    case "ValueAxis":
                    case "PlotArea":
                    case "ThreeDProperties":
                        //the chart of 2005, processed after the data, see ProcessChart2005Nodes
                        break;

                    default:
                        if (!ProcessInteractivityNode(node, component, dataset))
                            ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            ProcessChart2005Nodes(baseNode, component, datasetName);
            ApplyChartAxisLimits(component, datasetName);
            chartAggregates = null;
            chartCategoryLevels = null;

            container.Components.Add(component);
        }

        /// <summary>
        /// Gives the chart the palette of Reporting Services. The colors go to a chart style of the report, and the
        /// chart takes it as its custom style, so the series without an own color are drawn in the colors of the
        /// palette. A named palette has one style for all the charts, the palette Custom a style per chart; the
        /// custom colors are used only with the palette Custom, and without them the chart is painted white.
        /// No palette is the palette Default.
        /// </summary>
        private void ApplyChartPalette(XmlNode baseNode, StiChart chart)
        {
            var palette = GetChildText(baseNode, "Palette")?.Trim();
            if (string.IsNullOrEmpty(palette))
                palette = "Default";

            if (IsExpression(palette))
            {
                ThrowError(baseNode.Name, "Palette");
                return;
            }

            var report = chart.Report;
            string styleName;
            Color[] colors;
            if (palette.Equals("Custom", StringComparison.OrdinalIgnoreCase))
            {
                colors = GetChartCustomPaletteColors(baseNode);
                styleName = $"{chart.Name}Palette";
                for (var index = 2; report.Styles[styleName] != null; index++)
                {
                    styleName = $"{chart.Name}Palette{index}";
                }
            }
            else
            {
                colors = GetChartPaletteColors(palette);
                if (colors == null)
                {
                    ThrowError(baseNode.Name, "Palette");
                    return;
                }

                styleName = $"Palette{palette}";
            }

            if (!(report.Styles[styleName] is StiChartStyle))
            {
                //the other values of the style are the ones of the chart, only the colors are of the palette;
                //a line of Reporting Services has no markers unless its series or points set them
                var style = new StiChartStyle(styleName);
                style.GetStyleFromComponent(chart, StiStyleElements.All);
                style.StyleColors = colors;
                style.MarkerVisible = false;
                report.Styles.Add(style);
            }

            chart.CustomStyleName = styleName;
            chart.Style = new StiCustomStyle(styleName);
        }

        /// <summary>
        /// Checks the custom properties of a chart. MaxDataCount is written by Bold Reports and the mark of
        /// the upgrade by the designer of Reporting Services, the rectangle is the shape of a legend item by
        /// default; none of them changes the look of the chart.
        /// </summary>
        private void ProcessChartCustomProperties(XmlNode baseNode)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name != "CustomProperty")
                    continue;

                var name = GetChildText(node, "Name");
                var isRectangleShape = name == "LegendShape" &&
                    string.Equals(GetChildText(node, "Value"), "Rectangle", StringComparison.OrdinalIgnoreCase);
                if (name != "MaxDataCount" && name != "__Upgraded2005__" && !isRectangleShape)
                    ThrowError(null, null, $"Chart custom property not supported: {name}");
            }
        }

        private Color[] GetChartCustomPaletteColors(XmlNode baseNode)
        {
            var colors = new List<Color>();
            var paletteNode = GetChildNode(baseNode, "ChartCustomPaletteColors");
            if (paletteNode != null)
            {
                foreach (XmlNode node in paletteNode.ChildNodes)
                {
                    if (node.Name != "ChartCustomPaletteColor")
                        continue;

                    var value = GetNodeTextValue(node);
                    if (IsExpression(value))
                    {
                        ThrowError(paletteNode.Name, node.Name);
                        continue;
                    }

                    colors.Add(ParseColor(value));
                }
            }

            if (colors.Count == 0)
                colors.Add(Color.White);

            return colors.ToArray();
        }

        /// <summary>
        /// Converts the legend, the title, the axes and the plot area of a chart of the schema of
        /// 2005. A series added to the chart replaces the area that does not accept its type, so
        /// these nodes are read after the data, else the settings of the area are lost.
        /// </summary>
        private void ProcessChart2005Nodes(XmlNode baseNode, StiChart chart, string dataset)
        {
            var axisArea = chart.Area as IStiAxisArea;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Legend":
                        ProcessChartLegendType(node, chart.Legend);
                        break;

                    case "Title":
                        ProcessChartTitleType(node, chart.Title, dataset);
                        break;

                    case "CategoryAxis":
                        if (axisArea != null)
                        {
                            var axis = GetChartTargetAxis(axisArea, isCategory: true, index: 0);
                            ProcessChartAxisWrapperType(node, axis);
                        }
                        break;

                    case "ValueAxis":
                        if (axisArea != null)
                        {
                            var axis = GetChartTargetAxis(axisArea, isCategory: false, index: 0);
                            ProcessChartAxisWrapperType(node, axis);
                        }
                        break;

                    case "PlotArea":
                        var style = GetChildNode(node, "Style");
                        if (style != null)
                            ProcessChartAreaStyleType(style, chart.Area);
                        break;

                    case "ThreeDProperties":
                        if (IsTrue(GetChildText(node, "Enabled")))
                            ThrowError(null, null, $"3D chart of 2005 not imported: {chart.Name}");
                        break;
                }
            }
        }

        private static string GetChildText(XmlNode node, string name)
        {
            return GetChildNode(node, name)?.InnerText.Trim() ?? string.Empty;
        }

        /// <summary>
        /// Collects the sorting of the groups of a hierarchy of a chart, the outer group first, as the direction and
        /// the expression of the sorting of the chart. The groups of Reporting Services are sorted by their sort
        /// expressions, as the months by the key of the month; the chart took the points in the order of the
        /// groups of its source, the names of the months as text. A sort expression of the schema of 2008 is in
        /// SortExpressions, of 2005 in Sorting.
        /// </summary>
        private void CollectChartGroupSort(XmlNode baseNode, string dataset, List<string> sort)
        {
            var nodes = baseNode.SelectNodes(
                ".//*[local-name()='SortExpressions']/*[local-name()='SortExpression'] | " +
                ".//*[local-name()='Sorting']/*[local-name()='SortBy']");

            foreach (XmlNode node in nodes)
            {
                var valueNode = GetChildNode(node, "Value") ?? GetChildNode(node, "SortExpression");
                var value = valueNode != null ? GetNodeTextValue(valueNode)?.Trim() : null;
                //a total of another scope, as the whole data set, is the same for every group and orders nothing
                if (string.IsNullOrEmpty(value) || ChartScopedAggregateRegex.IsMatch(value))
                    continue;

                //a field out of the groups is sorted by its least value in the group, see CollectChartSortKeys
                var field = ChartGroupFieldRegex.Match(value).Groups["field"].Value;
                if (chartAggregates != null && dataset == chartAggregatesSource &&
                    chartAggregates.ContainsKey($"Min|{field}"))
                {
                    value = $"=Min(Fields!{field}.Value)";
                }

                var expression = ConvertChartExpression(value, dataset);
                if (string.IsNullOrWhiteSpace(expression))
                    continue;

                sort.Add(GetChildText(node, "Direction") == "Descending" ? "DESC" : "ASC");
                sort.Add(expression);
            }
        }

        private string ProcessChartHierarchyType(XmlNode baseNode, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupExpression":
                    case "Label":
                        var expression = ConvertChartExpression(GetNodeTextValue(node), dataset);
                        if (!string.IsNullOrWhiteSpace(expression))
                            return expression;
                        break;

                    default:
                        expression = ProcessChartHierarchyType(node, dataset);
                        if (!string.IsNullOrWhiteSpace(expression))
                            return expression;
                        break;
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Collects the expressions of the nested category groups, from the outer group to the inner one: the
        /// first group expression of the first grouped member on each level. The levels are the rings of a
        /// sunburst, see AddSunburstLevels. An empty group of Reporting Services is a segment under its parent,
        /// so an empty value of an inner level is the value of the level above it: the segment goes on through
        /// the rings, and the empty values under different parents are not one segment.
        /// </summary>
        private void CollectChartCategoryLevels(XmlNode baseNode, string dataset, List<string> levels,
            string parentLevel = null)
        {
            var members = GetChildNode(baseNode, "ChartMembers");
            if (members == null)
                return;

            foreach (XmlNode member in members.ChildNodes)
            {
                var expression = member.SelectSingleNode(
                    "*[local-name()='Group']/*[local-name()='GroupExpressions']/*[local-name()='GroupExpression']");
                if (expression == null)
                    continue;

                var text = GetNodeTextValue(expression)?.Trim();
                if (string.IsNullOrEmpty(text))
                    continue;

                var isExpression = text.StartsWith("=", StringComparison.Ordinal);
                var rdlLevel = isExpression ? $"({text.Substring(1)})" : $"\"{text}\"";
                if (parentLevel != null)
                {
                    rdlLevel = $"IIF(IsNothing({rdlLevel}) OrElse CStr({rdlLevel}) = \"\", {parentLevel}, " +
                        $"{rdlLevel})";
                }

                var level = ConvertChartExpression(parentLevel != null ? "=" + rdlLevel : text, dataset);
                if (string.IsNullOrWhiteSpace(level))
                    continue;

                levels.Add(level);
                CollectChartCategoryLevels(member, dataset, levels, rdlLevel);
                return;
            }
        }

        //A real series grouping is defined only by Group.GroupExpressions.GroupExpression.
        //A static ChartMember with only a Label is not a grouping and must not create an auto series.
        private string ProcessChartSeriesHierarchyType(XmlNode baseNode, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupExpression":
                        var expression = ConvertChartExpression(GetNodeTextValue(node), dataset);
                        if (!string.IsNullOrWhiteSpace(expression))
                            return expression;
                        break;

                    case "Label":
                    case "SortExpressions":
                        break;

                    default:
                        expression = ProcessChartSeriesHierarchyType(node, dataset);
                        if (!string.IsNullOrWhiteSpace(expression))
                            return expression;
                        break;
                }
            }

            return string.Empty;
        }

        private void ProcessChartDataType(XmlNode baseNode, StiChart chart, string dataset, string categoryExpression, string seriesExpression, Dictionary<string, bool> chartAreas3D)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartSeriesCollection":
                        ProcessChartSeriesCollectionType(node, chart, dataset, categoryExpression, seriesExpression, chartAreas3D);
                        break;

                    case "ChartSeries":
                        ProcessChartSeriesType(node, chart, dataset, categoryExpression, seriesExpression, chartAreas3D);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartSeriesCollectionType(XmlNode baseNode, StiChart chart, string dataset, string categoryExpression, string seriesExpression, Dictionary<string, bool> chartAreas3D)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartSeries":
                        ProcessChartSeriesType(node, chart, dataset, categoryExpression, seriesExpression, chartAreas3D);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartSeriesType(XmlNode baseNode, StiChart chart, string dataset, string categoryExpression, string seriesExpression, Dictionary<string, bool> chartAreas3D)
        {
            var type = string.Empty;
            var subtype = string.Empty;
            var chartAreaName = string.Empty;
            var valueAxisName = string.Empty;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Type":
                        type = GetNodeTextValue(node);
                        break;

                    case "Subtype":
                        subtype = GetNodeTextValue(node);
                        break;

                    case "ChartAreaName":
                        chartAreaName = GetNodeTextValue(node);
                        break;

                    case "ValueAxisName":
                        valueAxisName = GetNodeTextValue(node);
                        break;
                }
            }

            if (string.IsNullOrWhiteSpace(type))
            {
                type = defaultChartType;
                subtype = string.IsNullOrWhiteSpace(subtype) ? defaultChartSubtype : subtype;
            }

            var series = CreateChartSeries(baseNode, type, subtype, IsChartSeries3D(chartAreaName, chartAreas3D));
            var name = baseNode.Attributes["Name"]?.Value;
            if (!string.IsNullOrWhiteSpace(name))
                series.CoreTitle = name;

            SetChartSeriesValueAxis(chart, series, valueAxisName);

            //the type of a marker of Reporting Services is None by default, a marker of a line is shown by default
            //and comes into view as soon as the series drops the chart style
            var marker = GetSeriesMarker(series, "Marker");
            if (marker != null)
                marker.Visible = false;

            //the labels of Reporting Services are hidden unless a data label shows them; the labels of the chart
            //would show the values of every series, and a pie can not show the labels of an axis
            series.SeriesLabels = CreateChartSeriesLabels(baseNode, series) ?? series.SeriesLabels;
            series.ShowSeriesLabels = StiShowSeriesLabels.FromSeries;
            if (series.SeriesLabels != null)
                series.SeriesLabels.Visible = false;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartDataPoints":
                    case "DataPoints":
                        ProcessChartDataPointsType(node, series, dataset);
                        break;

                    case "CustomProperties":
                        ProcessChartSeriesCustomProperties(node, series);
                        break;

                    case "ChartDataLabel":
                    case "DataLabel":
                        ProcessChartDataLabelType(node, series, dataset);
                        break;

                    case "Style":
                        ProcessChartSeriesStyleType(node, series);
                        break;

                    case "ChartEmptyPoints":
                    case "EmptyPoints":
                        ProcessChartEmptyPointsType(node, series);
                        break;

                    case "ChartSmartLabel":
                        //smart labels of Reporting Services move the labels apart unless disabled
                        if (series.SeriesLabels != null)
                            series.SeriesLabels.PreventIntersection = !IsTrue(GetChildText(node, "Disabled"));
                        break;

                    //the marker of the series is the marker of its points
                    case "ChartMarker":
                    case "Marker":
                        if (ProcessChartMarkerType(node, GetSeriesMarker(series, "Marker")))
                            series.AllowApplyStyle = false;
                        break;

                    case "ChartItemInLegend":
                        ProcessChartItemInLegendType(node, series, dataset, seriesExpression);
                        break;

                    case "Hidden":
                        var hidden = GetNodeTextValue(node);
                        if (IsExpression(hidden))
                            ThrowError(baseNode.Name, node.Name);

                        else if (IsTrue(hidden))
                            return;
                        break;

                    case "Type":
                    case "Subtype":
                    case "LegendName":
                    case "ChartAreaName":
                    case "ValueAxisName":
                    case "CategoryAxisName":
                    case "DataElementName":
                    case "DataElementOutput":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (string.IsNullOrWhiteSpace(series.Argument.Value) && !string.IsNullOrWhiteSpace(categoryExpression))
                series.Argument.Value = categoryExpression;

            //a series of a sunburst is a ring, the groups of the series of Reporting Services only color it
            var isSunburst = series is StiSunburstSeries;
            if (!isSunburst && !string.IsNullOrWhiteSpace(seriesExpression))
            {
                var autoSeriesDataColumn = UnwrapChartExpression(seriesExpression);
                if (string.IsNullOrWhiteSpace(series.AutoSeriesKeyDataColumn))
                    series.AutoSeriesKeyDataColumn = autoSeriesDataColumn;
                if (string.IsNullOrWhiteSpace(series.AutoSeriesTitleDataColumn))
                    series.AutoSeriesTitleDataColumn = autoSeriesDataColumn;
            }

            chart.Series.Add(series);
            MoveLabelsToChart(chart, series);

            if (isSunburst)
                AddSunburstLevels(chart, series);
        }

        /// <summary>
        /// Adds a ring to a sunburst for each inner category group. A sunburst of Reporting Services draws the
        /// nested category groups as rings of one series; a sunburst of ours takes a ring from each series, the
        /// argument of the series is the group of the ring, and the values from the first series. With one series
        /// for all the levels every ring showed the outer group with the same values.
        /// </summary>
        private void AddSunburstLevels(StiChart chart, StiSeries series)
        {
            if (chartCategoryLevels == null || chartCategoryLevels.Count < 2 || chart.Series.Count != 1)
                return;

            series.Argument.Value = chartCategoryLevels[0];

            foreach (var level in chartCategoryLevels.Skip(1))
            {
                var ring = new StiSunburstSeries
                {
                    CoreTitle = series.CoreTitle,
                    Argument = new StiArgumentExpression(level),
                    Value = new StiExpression(series.Value.Value),
                    ShowSeriesLabels = series.ShowSeriesLabels,
                    ShowInLegend = series.ShowInLegend,
                    AllowApplyStyle = series.AllowApplyStyle
                };
                ring.SeriesLabels.Visible = series.SeriesLabels.Visible;

                chart.Series.Add(ring);
            }
        }

        /// <summary>
        /// Moves the shown labels of a tree map or a sunburst to the chart, the series shows the labels of the
        /// chart. A tree map reads only the labels of the chart; a sunburst reads the labels of the series, but
        /// draws nothing with them. The labels move after the series is added: the chart then takes the area
        /// of the series and the labels of that area, None for a tree map.
        /// </summary>
        private static void MoveLabelsToChart(StiChart chart, StiSeries series)
        {
            if (!(series is StiTreemapSeries) && !(series is StiSunburstSeries))
                return;

            if (series.SeriesLabels == null || !series.SeriesLabels.Visible)
                return;

            var labels = (IStiSeriesLabels)series.SeriesLabels.Clone();

            //a tile of a tree map of Reporting Services shows its category above the data label, hidden
            //data labels hide the category too
            if (series is StiTreemapSeries && labels.ValueType == StiSeriesLabelsValueType.Value)
                labels.ValueType = StiSeriesLabelsValueType.ArgumentValue;

            chart.SeriesLabels = labels;
            series.ShowSeriesLabels = StiShowSeriesLabels.FromChart;
        }

        /// <summary>
        /// Returns the labels that fit the series, or null to keep its own. A pie shows its labels inside unless
        /// PieLabelStyle or the position of the data label puts them outside; a funnel the same with
        /// PyramidLabelStyle or FunnelLabelStyle. The labels of a column or a bar are out of its end by
        /// default; BarLabelStyle or the position of the data label move them.
        /// </summary>
        private IStiSeriesLabels CreateChartSeriesLabels(XmlNode baseNode, StiSeries series)
        {
            var position = GetChartDataLabelPosition(baseNode);

            //a tree map and a sunburst draw only the labels of the chart, see MoveLabelsToChart; a tile of a
            //tree map shows its category above the value, a sunburst turns its labels along the radius
            if (series is StiTreemapSeries)
                return InitHierarchyChartLabels(new StiCenterTreemapLabels { ValueTypeSeparator = "\n" });

            if (series is StiSunburstSeries)
                return InitHierarchyChartLabels(new StiCenterPieLabels { AutoRotate = true });

            if (series is StiPie3dSeries)
                return new StiCenterPie3dLabels();

            if (series is StiPieSeries)
            {
                var outside = position == "outside" ||
                    GetChartSeriesCustomProperty(baseNode, "PieLabelStyle") == "outside";

                return outside ? (IStiSeriesLabels)new StiOutsidePieLabels() : new StiCenterPieLabels();
            }

            if (series is StiFunnelSeries)
            {
                var outside = position == "outside" ||
                    GetChartSeriesCustomProperty(baseNode, "PyramidLabelStyle") == "outside" ||
                    GetChartSeriesCustomProperty(baseNode, "FunnelLabelStyle") == "outside";

                return outside ? (IStiSeriesLabels)new StiOutsideRightFunnelLabels() : new StiCenterFunnelLabels();
            }

            if (series.GetType() != typeof(StiClusteredColumnSeries) &&
                series.GetType() != typeof(StiClusteredBarSeries))
            {
                return null;
            }

            if (string.IsNullOrEmpty(position) || position == "auto")
            {
                switch (GetChartSeriesCustomProperty(baseNode, "BarLabelStyle"))
                {
                    case "center":
                        return new StiCenterAxisLabels();

                    case "left":
                        return new StiInsideBaseAxisLabels();

                    case "right":
                        return new StiInsideEndAxisLabels();

                    default:
                        return new StiOutsideEndAxisLabels();
                }
            }

            if (position == "center")
                return new StiCenterAxisLabels();

            //the base of a bar is on the left, of a column on the bottom
            var baseSide = series is StiClusteredBarSeries ? "left" : "bottom";

            return position.StartsWith(baseSide, StringComparison.Ordinal)
                ? (IStiSeriesLabels)new StiInsideBaseAxisLabels()
                : new StiOutsideEndAxisLabels();
        }

        /// <summary>
        /// Gives the labels of a tree map or a sunburst the look of a data label of Reporting Services: black text
        /// of 10pt without a background and a border. The chart style would draw the labels on white plates, so
        /// the labels do not take it; the style of the data label sets the font and the color then.
        /// </summary>
        private IStiSeriesLabels InitHierarchyChartLabels(IStiSeriesLabels labels)
        {
            labels.AllowApplyStyle = false;
            labels.Brush = new StiSolidBrush(Color.Transparent);
            labels.BorderColor = Color.Transparent;
            labels.DrawBorder = false;
            labels.LabelColor = Color.Black;
            labels.Font = new Font(DefaultFontFamily, 10);

            return labels;
        }

        /// <summary>
        /// Returns the position of the first data label of the series or of its points in lower case, or null.
        /// </summary>
        private string GetChartDataLabelPosition(XmlNode baseNode)
        {
            var labels = baseNode.SelectNodes(".//*[local-name()='ChartDataLabel' or local-name()='DataLabel']");
            foreach (XmlNode node in labels)
            {
                var position = GetChildText(node, "Position");
                if (!string.IsNullOrEmpty(position) && !IsExpression(position))
                    return position.ToLowerInvariant();
            }

            return null;
        }

        /// <summary>
        /// Returns the value of a custom property of the series in lower case, or null.
        /// </summary>
        private string GetChartSeriesCustomProperty(XmlNode baseNode, string name)
        {
            var properties = GetChildNode(baseNode, "CustomProperties");
            if (properties == null)
                return null;

            foreach (XmlNode node in properties.ChildNodes)
            {
                if (node.Name == "CustomProperty" && GetChildText(node, "Name") == name)
                    return GetChildText(node, "Value").ToLowerInvariant();
            }

            return null;
        }

        /// <summary>
        /// Converts the custom properties of a series. The styles of the labels are read when the labels are
        /// created, see CreateChartSeriesLabels; the width of a point is the width of a column.
        /// </summary>
        private void ProcessChartSeriesCustomProperties(XmlNode baseNode, StiSeries series)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name != "CustomProperty")
                    continue;

                var name = GetChildText(node, "Name");
                var value = GetChildText(node, "Value");
                switch (name)
                {
                    case "PointWidth":
                        if (series is IStiClusteredColumnSeries column &&
                            double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var width) &&
                            width > 0 && width <= 1)
                        {
                            column.Width = (float)width;
                        }
                        break;

                    //the legend of Stimulsoft shows the series by the shape of its markers
                    case "LegendShape":
                        if (!string.Equals(value, "Rectangle", StringComparison.OrdinalIgnoreCase))
                            ThrowError(null, null, $"Chart series custom property not supported: {name}");
                        break;

                    //Bold Reports writes the name None or no name
                    case "":
                    case "LabelStyle":
                    case "PieLabelStyle":
                    case "BarLabelStyle":
                    case "PyramidLabelStyle":
                    case "FunnelLabelStyle":
                    case "ShowColumnAs":
                    case "RadarDrawingStyle":
                    case "MaxDataCount":
                    case "None":
                        break;

                    default:
                        ThrowError(null, null, $"Chart series custom property not supported: {name}");
                        break;
                }
            }
        }

        /// <summary>
        /// Converts the item of the series in the legend: the text is the title of the series, a hidden item is
        /// not shown in the legend. A series of a series group takes its titles from the group, its text stays.
        /// </summary>
        private void ProcessChartItemInLegendType(XmlNode baseNode, StiSeries series, string dataset,
            string seriesExpression)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "LegendText":
                        var text = GetNodeTextValue(node)?.Trim();
                        if (string.IsNullOrEmpty(text) || !string.IsNullOrWhiteSpace(seriesExpression))
                            break;

                        if (IsExpression(text))
                            series.Title.Value = ConvertChartExpression(text, dataset);

                        else
                            series.CoreTitle = text;
                        break;

                    case "Hidden":
                        var hidden = GetNodeTextValue(node);
                        if (IsExpression(hidden))
                            ThrowError(baseNode.Name, node.Name);

                        else
                            series.ShowInLegend = !IsTrue(hidden);
                        break;

                    case "ToolTip":
                    case "ActionInfo":
                    case "CustomProperties":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void SetChartSeriesValueAxis(StiChart chart, StiSeries series, string valueAxisName)
        {
            if (string.IsNullOrWhiteSpace(valueAxisName) || secondaryValueAxisNames == null ||
                !secondaryValueAxisNames.Contains(valueAxisName))
                return;

            series.YAxis = StiSeriesYAxis.RightYAxis;

            var axisArea = chart.Area as IStiAxisArea;
            if (axisArea != null)
                axisArea.YRightAxis.Visible = true;
        }

        private HashSet<string> GetSecondaryValueAxisNames(XmlNode baseNode)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartAreas":
                    case "ChartAreaCollection":
                        foreach (XmlNode areaNode in node.ChildNodes)
                        {
                            if (areaNode.Name != "ChartArea")
                                continue;

                            CollectSecondaryValueAxisNames(areaNode, names);
                        }
                        break;
                }
            }

            return names;
        }

        private void CollectSecondaryValueAxisNames(XmlNode areaNode, HashSet<string> names)
        {
            foreach (XmlNode node in areaNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartValueAxes":
                    case "ValueAxes":
                        var index = 0;
                        foreach (XmlNode axisNode in node.ChildNodes)
                        {
                            if (axisNode.Name != "ChartValueAxis" && axisNode.Name != "ChartAxis")
                                continue;

                            if (index > 0)
                            {
                                var name = GetChartAxisName(axisNode);
                                if (!string.IsNullOrWhiteSpace(name))
                                    names.Add(name);
                            }
                            index++;
                        }
                        break;
                }
            }
        }

        private static string GetChartAxisName(XmlNode axisNode)
        {
            var name = axisNode.Attributes["Name"]?.Value;
            if (!string.IsNullOrWhiteSpace(name))
                return name;

            foreach (XmlNode node in axisNode.ChildNodes)
            {
                if (node.Name == "ChartAxis")
                    return node.Attributes["Name"]?.Value;
            }

            return null;
        }

        private Dictionary<string, bool> GetChartAreas3D(XmlNode baseNode)
        {
            var chartAreas3D = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartAreas":
                    case "ChartAreaCollection":
                        foreach (XmlNode areaNode in node.ChildNodes)
                        {
                            if (areaNode.Name != "ChartArea")
                                continue;

                            var name = areaNode.Attributes["Name"]?.Value;
                            if (string.IsNullOrWhiteSpace(name))
                                name = string.Empty;

                            chartAreas3D[name] = IsChartArea3D(areaNode);
                        }
                        break;
                }
            }

            return chartAreas3D;
        }

        private bool IsChartArea3D(XmlNode baseNode)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartThreeDProperties":
                    case "ThreeDProperties":
                        return IsChartThreeDPropertiesEnabled(node);
                }
            }

            return false;
        }

        private bool IsChartThreeDPropertiesEnabled(XmlNode baseNode)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Enabled":
                        return IsTrue(GetNodeTextValue(node));
                }
            }

            return false;
        }

        private static bool IsChartSeries3D(string chartAreaName, Dictionary<string, bool> chartAreas3D)
        {
            if (chartAreas3D == null || chartAreas3D.Count == 0)
                return false;

            bool is3D;
            if (!string.IsNullOrWhiteSpace(chartAreaName) && chartAreas3D.TryGetValue(chartAreaName, out is3D))
                return is3D;

            if (chartAreas3D.TryGetValue("Default", out is3D))
                return is3D;

            if (chartAreas3D.TryGetValue(string.Empty, out is3D))
                return is3D;

            if (chartAreas3D.Count == 1)
            {
                foreach (var pair in chartAreas3D)
                {
                    return pair.Value;
                }
            }

            return false;
        }

        /// <summary>
        /// Creates the series of the type and the subtype of Reporting Services. The types of RDL 2008 and later
        /// are Area, Bar, Column, Line, Polar, Range, Scatter and Shape; the types of RDL 2005 Pie, Doughnut,
        /// Bubble and Stock are kept too. A box plot and an error bar take precomputed values that no series
        /// reads, they stay a range.
        /// </summary>
        private StiSeries CreateChartSeries(XmlNode baseNode, string type, string subtype, bool chartArea3D)
        {
            var chartType = NormalizeChartType(type);
            var is3D = chartArea3D || IsChart3DType(type) || IsChart3DSubtype(subtype);

            switch (chartType)
            {
                case "bar":
                    if (IsFullStackedChartSubtype(subtype))
                        return new StiFullStackedBarSeries();

                    if (IsStackedChartSubtype(subtype))
                        return new StiStackedBarSeries();

                    return new StiClusteredBarSeries();

                case "line":
                    if (IsFullStackedChartSubtype(subtype))
                        return new StiFullStackedLineSeries();
                    if (IsStackedChartSubtype(subtype))
                        return new StiStackedLineSeries();
                    if (is3D)
                        return new StiLineSeries3D();
                    if (string.Equals(subtype, "Smooth", StringComparison.OrdinalIgnoreCase))
                        return new StiSplineSeries();
                    if (string.Equals(subtype, "Stepped", StringComparison.OrdinalIgnoreCase))
                        return new StiSteppedLineSeries();
                    return new StiLineSeries();

                case "area":
                    if (IsFullStackedChartSubtype(subtype))
                        return new StiFullStackedAreaSeries();
                    if (IsStackedChartSubtype(subtype))
                        return new StiStackedAreaSeries();
                    if (is3D)
                        return new StiAreaSeries3D();
                    if (string.Equals(subtype, "Smooth", StringComparison.OrdinalIgnoreCase))
                        return new StiSplineAreaSeries();
                    if (string.Equals(subtype, "Stepped", StringComparison.OrdinalIgnoreCase))
                        return new StiSteppedAreaSeries();
                    return new StiAreaSeries();

                case "shape":
                case "pie":
                case "doughnut":
                    if (IsChartSubtype(subtype, "TreeMap"))
                        return new StiTreemapSeries();

                    if (IsChartSubtype(subtype, "Sunburst"))
                        return new StiSunburstSeries();

                    if (chartType == "doughnut" || IsChartSubtype(subtype, "Doughnut") ||
                        IsChartSubtype(subtype, "ExplodedDoughnut"))
                        return new StiDoughnutSeries();

                    if (IsChartSubtype(subtype, "Funnel") || IsChartSubtype(subtype, "Pyramid"))
                        return new StiFunnelSeries();

                    if (is3D)
                        return new StiPie3dSeries();

                    return new StiPieSeries();

                case "scatter":
                    if (IsChartSubtype(subtype, "Bubble"))
                        return new StiBubbleSeries();

                    if (IsChartSubtype(subtype, "Smooth") || IsChartSubtype(subtype, "SmoothLine"))
                        return new StiScatterSplineSeries();

                    if (IsChartSubtype(subtype, "Line"))
                        return new StiScatterLineSeries();

                    return new StiScatterSeries();

                case "bubble":
                    return new StiBubbleSeries();

                case "polar":
                    if (IsChartSubtype(subtype, "Radar"))
                        return CreateRadarSeries(GetChartSeriesCustomProperty(baseNode, "RadarDrawingStyle"));

                    return new StiPolarSeries();

                case "radar":
                    return CreateRadarSeries(string.IsNullOrWhiteSpace(subtype) ? "line" : subtype.ToLowerInvariant());

                case "range":
                    if (IsChartSubtype(subtype, "Smooth"))
                        return new StiSplineRangeSeries();

                    //a range column of Reporting Services stands upright as a range bar of ours,
                    //a range bar lies as a gantt
                    if (IsChartSubtype(subtype, "Column"))
                        return new StiRangeBarSeries();

                    if (IsChartSubtype(subtype, "Bar"))
                        return new StiGanttSeries();

                    if (IsChartSubtype(subtype, "Candlestick"))
                        return new StiCandlestickSeries();

                    if (IsChartSubtype(subtype, "Stock"))
                        return new StiStockSeries();

                    return new StiRangeSeries();

                case "stock":
                    if (IsChartSubtype(subtype, "Candlestick"))
                        return new StiCandlestickSeries();

                    return new StiStockSeries();

                case "candlestick":
                    return new StiCandlestickSeries();

                case "surface":
                case "wireframesurface":
                    return new StiWireframeSurfaceSeries3D();

                case "column":
                default:
                    if (IsFullStackedChartSubtype(subtype))
                    {
                        if (is3D)
                            return new StiFullStackedColumnSeries3D();

                        return new StiFullStackedColumnSeries();
                    }
                    if (IsStackedChartSubtype(subtype))
                    {
                        if (is3D)
                            return new StiStackedColumnSeries3D();

                        return new StiStackedColumnSeries();
                    }
                    if (is3D)
                        return new StiClusteredColumnSeries3D();

                    //a plain column of Reporting Services turns into a histogram or a pareto by ShowColumnAs
                    switch (GetChartSeriesCustomProperty(baseNode, "ShowColumnAs"))
                    {
                        case "histogram":
                            return new StiHistogramSeries();

                        case "pareto":
                            return new StiParetoSeries();
                    }

                    return new StiClusteredColumnSeries();
            }
        }

        /// <summary>
        /// Creates the radar series of the drawing style in lower case: Area, Line or Marker of
        /// RadarDrawingStyle. The area is the default style of Reporting Services.
        /// </summary>
        private static StiSeries CreateRadarSeries(string style)
        {
            switch (style)
            {
                case "line":
                    return new StiRadarLineSeries();

                case "marker":
                case "point":
                    return new StiRadarPointSeries();

                default:
                    return new StiRadarAreaSeries();
            }
        }

        private static bool IsChartSubtype(string subtype, string name)
        {
            return string.Equals(subtype, name, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeChartType(string type)
        {
            var chartType = (type ?? string.Empty).ToLowerInvariant();

            if (chartType.EndsWith("3d", StringComparison.OrdinalIgnoreCase))
                chartType = chartType.Substring(0, chartType.Length - 2);

            if (chartType.StartsWith("3d", StringComparison.OrdinalIgnoreCase))
                chartType = chartType.Substring(2);

            return chartType;
        }

        private static bool IsChart3DType(string type)
        {
            return HasChart3DMarker(type) ||
                string.Equals(type, "Surface", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "WireframeSurface", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "Cylinder", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(type, "Cone", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsChart3DSubtype(string subtype)
        {
            return HasChart3DMarker(subtype) ||
                string.Equals(subtype, "Cylinder", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(subtype, "Cone", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasChart3DMarker(string value)
        {
            return !string.IsNullOrWhiteSpace(value) && (
                value.IndexOf("3D", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("3-D", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("ThreeD", StringComparison.OrdinalIgnoreCase) >= 0 ||
                value.IndexOf("ThreeDimensional", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static bool IsStackedChartSubtype(string subtype)
        {
            return string.Equals(subtype, "Stacked", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(subtype, "StackedPercent", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(subtype, "PercentStacked", StringComparison.OrdinalIgnoreCase) ||
                (!string.IsNullOrWhiteSpace(subtype) &&
                    subtype.IndexOf("Stacked", StringComparison.OrdinalIgnoreCase) >= 0);
        }
        private static bool IsFullStackedChartSubtype(string subtype)
        {
            return string.Equals(subtype, "StackedPercent", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(subtype, "PercentStacked", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(subtype, "FullStacked", StringComparison.OrdinalIgnoreCase) ||
                (!string.IsNullOrWhiteSpace(subtype) &&
                    (subtype.IndexOf("StackedPercent", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    subtype.IndexOf("PercentStacked", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    subtype.IndexOf("FullStacked", StringComparison.OrdinalIgnoreCase) >= 0));
        }

        private void ProcessChartDataPointsType(XmlNode baseNode, StiSeries series, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartDataPoint":
                    case "DataPoint":
                        ProcessChartDataPointType(node, series, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartDataPointType(XmlNode baseNode, StiSeries series, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartDataPointValues":
                    case "DataValues":
                        ProcessChartDataPointValuesType(node, series, dataset);
                        break;

                    case "ChartDataLabel":
                    case "DataLabel":
                        ProcessChartDataLabelType(node, series, dataset);
                        break;

                    case "Style":
                        ProcessChartSeriesStyleType(node, series);
                        break;

                    case "ChartMarker":
                        if (ProcessChartMarkerType(node, GetSeriesMarker(series, "Marker")))
                            series.AllowApplyStyle = false;
                        break;

                    case "ToolTip":
                        series.ToolTip = new StiToolTipExpression(ConvertInteractiveValue(node, dataset));
                        break;

                    case "Action":
                    case "ActionInfo":
                        ProcessChartDataPointAction(node, series, dataset);
                        break;

                    //the item in the legend of the point of a series is the item of the series
                    case "ChartItemInLegend":
                        ProcessChartItemInLegendType(node, series, dataset, series.AutoSeriesKeyDataColumn);
                        break;

                    case "Marker":
                        if (ProcessChartMarkerType(node, GetSeriesMarker(series, "Marker")))
                            series.AllowApplyStyle = false;
                        break;

                    case "DataElementName":
                    case "DataElementOutput":
                    case "CustomProperties":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartEmptyPointsType(XmlNode baseNode, StiSeries series)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartMarker":
                    case "Marker":
                        var lineMarker = GetSeriesMarker(series, "LineMarker");
                        ProcessChartMarkerType(node, lineMarker);
                        if (lineMarker != null && IsChartMarkerVisible(node))
                            SetSeriesShowNulls(series, true);
                        break;

                    case "Style":
                    case "ChartDataLabel":
                    case "DataLabel":
                    case "AxisLabel":
                    case "CustomProperties":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        //Returns true when the marker is shown and was explicitly configured (type, visibility or color): the
        //style of the chart would hide it. An empty or a hidden marker leaves the series style untouched so the
        //palette still applies, the style of the palette hides the markers as Reporting Services does.
        private bool ProcessChartMarkerType(XmlNode baseNode, IStiMarker marker)
        {
            if (marker == null)
                return false;

            var configured = false;
            var shown = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Type":
                        SetChartMarkerType(marker, GetNodeTextValue(node));
                        configured = true;
                        shown = marker.Visible;
                        break;

                    case "Size":
                        try
                        {
                            var size = (float)ToPt(GetNodeTextValue(node), true);
                            if (size > 0)
                                marker.Size = size;
                        }
                        catch
                        {
                        }
                        break;

                    case "Visible":
                        marker.Visible = IsTrue(GetNodeTextValue(node));
                        configured = true;
                        shown = marker.Visible;
                        break;

                    case "Hidden":
                        marker.Visible = !IsTrue(GetNodeTextValue(node));
                        configured = true;
                        shown = marker.Visible;
                        break;

                    case "Style":
                        if (ProcessChartMarkerStyleType(node, marker))
                            configured = true;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            //the type of a marker of Reporting Services is None by default
            return configured && shown;
        }

        private bool ProcessChartMarkerStyleType(XmlNode baseNode, IStiMarker marker)
        {
            var configured = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                    case "BackgroundColor":
                        marker.Brush = new StiSolidBrush(ParseColor(GetChartStyleValue(node)));
                        configured = true;
                        break;

                    case "BorderColor":
                        marker.BorderColor = ParseColor(GetChartStyleValue(node));
                        configured = true;
                        break;

                    case "Border":
                        var markerBorder = ReadChartBorder(node);
                        if (markerBorder.Item1.HasValue)
                        {
                            marker.BorderColor = markerBorder.Item1.Value;
                            configured = true;
                        }
                        break;

                    //a marker has no text
                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                    case "TextDecoration":
                    case "TextAlign":
                    case "VerticalAlign":
                    case "Format":
                    case "BorderStyle":
                    case "BorderWidth":
                    case "BackgroundGradientType":
                    case "BackgroundGradientEndColor":
                    case "BackgroundHatchType":
                    case "ShadowColor":
                    case "ShadowOffset":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return configured;
        }

        private static void SetChartMarkerType(IStiMarker marker, string type)
        {
            if (string.IsNullOrWhiteSpace(type))
                return;

            if (string.Equals(type, "None", StringComparison.OrdinalIgnoreCase))
            {
                marker.Visible = false;
                return;
            }

            marker.Visible = true;

            if (string.Equals(type, "Auto", StringComparison.OrdinalIgnoreCase))
                return;

            if (string.Equals(type, "Square", StringComparison.OrdinalIgnoreCase))
                marker.Type = StiMarkerType.Rectangle;

            else if (string.Equals(type, "Circle", StringComparison.OrdinalIgnoreCase))
                marker.Type = StiMarkerType.Circle;

            else if (string.Equals(type, "Triangle", StringComparison.OrdinalIgnoreCase))
                marker.Type = StiMarkerType.Triangle;

            else if (type.IndexOf("Star", StringComparison.OrdinalIgnoreCase) >= 0)
                marker.Type = StiMarkerType.Star5;
        }

        private bool IsChartMarkerVisible(XmlNode baseNode)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Type":
                        var type = GetNodeTextValue(node);
                        return !string.IsNullOrWhiteSpace(type) &&
                            !string.Equals(type, "None", StringComparison.OrdinalIgnoreCase);

                    case "Visible":
                        return IsTrue(GetNodeTextValue(node));
                }
            }

            return false;
        }

        private static IStiMarker GetSeriesMarker(StiSeries series, string propertyName)
        {
            var property = series.GetType().GetProperty(propertyName);
            if (property == null)
                return null;

            return property.GetValue(series, null) as IStiMarker;
        }

        private static void SetSeriesShowNulls(StiSeries series, bool value)
        {
            var property = series.GetType().GetProperty("ShowNulls");
            if (property != null && property.PropertyType == typeof(bool))
                property.SetValue(series, value, null);
        }

        private void ProcessChartDataPointValuesType(XmlNode baseNode, StiSeries series, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataValue":
                        ProcessChartDataPointValuesType(node, series, dataset);
                        break;

                    case "Value":
                    case "Y":
                        SetChartSeriesValue(series, ConvertChartExpression(GetNodeTextValue(node), dataset));
                        break;

                    case "X":
                    case "Argument":
                        SetChartSeriesArgument(series, ConvertChartExpression(GetNodeTextValue(node), dataset));
                        break;

                    case "Size":
                        SetChartSeriesExpression(series, "Weight", ConvertChartExpression(GetNodeTextValue(node), dataset));
                        break;

                    case "High":
                    case "End":
                        SetChartSeriesHigh(series, ConvertChartExpression(GetNodeTextValue(node), dataset));
                        break;

                    case "Low":
                    case "Start":
                        SetChartSeriesLow(series, ConvertChartExpression(GetNodeTextValue(node), dataset));
                        break;

                    case "Open":
                        SetChartSeriesExpression(series, "ValueOpen", ConvertChartExpression(GetNodeTextValue(node), dataset));
                        break;

                    case "Close":
                        SetChartSeriesExpression(series, "ValueClose", ConvertChartExpression(GetNodeTextValue(node), dataset));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private static void SetChartSeriesValue(StiSeries series, string expression)
        {
            if (!string.IsNullOrWhiteSpace(expression) && string.IsNullOrWhiteSpace(series.Value.Value))
                series.Value.Value = expression;
        }
        private static void SetChartSeriesArgument(StiSeries series, string expression)
        {
            if (!string.IsNullOrWhiteSpace(expression) && string.IsNullOrWhiteSpace(series.Argument.Value))
                series.Argument.Value = expression;
        }

        private static void SetChartSeriesHigh(StiSeries series, string expression)
        {
            if (SetChartSeriesExpression(series, "ValueHigh", expression))
                return;

            SetChartSeriesExpression(series, "ValueEnd", expression);
        }

        private static void SetChartSeriesLow(StiSeries series, string expression)
        {
            if (SetChartSeriesExpression(series, "ValueLow", expression))
                return;

            SetChartSeriesValue(series, expression);
        }

        private static bool SetChartSeriesExpression(StiSeries series, string propertyName, string expression)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return false;

            var property = series.GetType().GetProperty(propertyName);
            if (property == null)
                return false;

            var target = property.GetValue(series, null) as StiExpression;
            if (target == null)
                return false;

            if (string.IsNullOrWhiteSpace(target.Value))
                target.Value = expression;

            return true;
        }

        private void ProcessChartDataLabelType(XmlNode baseNode, StiSeries series, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Visible":
                        series.ShowSeriesLabels = StiShowSeriesLabels.FromSeries;
                        series.SeriesLabels.Visible = IsTrue(GetNodeTextValue(node));
                        break;

                    case "Hidden":
                        series.ShowSeriesLabels = StiShowSeriesLabels.FromSeries;
                        series.SeriesLabels.Visible = !IsTrue(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessChartSeriesLabelsStyleType(node, series.SeriesLabels);
                        break;

                    case "Value":
                        series.SeriesLabels.TextAfter = ConvertChartExpression(GetNodeTextValue(node), dataset);
                        break;

                    case "Label":
                        if (!IsTrue(GetChildText(baseNode, "UseValueAsLabel")))
                            ProcessChartDataLabelText(GetNodeTextValue(node), series, dataset);
                        break;

                    case "Position":
                    case "Rotation":
                    case "UseValueAsLabel":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        /// <summary>
        /// Converts the text of the labels of a series. A keyword of Reporting Services with an optional format,
        /// as #VALY or #PERCENT{P1}, and the text around it become the value type, the format and the text before
        /// and after of the labels. Any other text or expression goes to the tag of the series, which the labels
        /// then show.
        /// </summary>
        private void ProcessChartDataLabelText(string label, StiSeries series, string dataset)
        {
            if (string.IsNullOrWhiteSpace(label))
                return;

            var labels = series.SeriesLabels;
            var match = ChartLabelKeywordRegex.Match(label);
            if (!IsExpression(label) && match.Success && ChartLabelKeywordRegex.Matches(label).Count == 1)
            {
                var keyword = match.Groups["keyword"].Value.ToUpperInvariant();
                var format = match.Groups["format"].Value;
                switch (keyword)
                {
                    case "VALX":
                    case "AXISLABEL":
                        labels.ValueType = StiSeriesLabelsValueType.Argument;
                        break;

                    case "SERIESNAME":
                    case "SER":
                        labels.ValueType = StiSeriesLabelsValueType.SeriesTitle;
                        break;

                    case "TOTAL":
                        labels.ValueType = StiSeriesLabelsValueType.Total;
                        break;

                    //the format P of the labels shows the share of the value in the total
                    case "PERCENT":
                        labels.ValueType = StiSeriesLabelsValueType.Value;
                        format = format.StartsWith("P", StringComparison.OrdinalIgnoreCase) ? format : "P2";
                        break;

                    case "VAL":
                    case "VALY":
                        labels.ValueType = StiSeriesLabelsValueType.Value;
                        break;

                    default:
                        ThrowError(null, null, $"Chart label keyword not supported: #{keyword}");
                        return;
                }

                if (!string.IsNullOrEmpty(format))
                    labels.Format = format;

                labels.TextBefore = label.Substring(0, match.Index);
                labels.TextAfter = label.Substring(match.Index + match.Length);
                return;
            }

            //the text with several keywords would be shown as is, the labels keep their value
            if (!IsExpression(label) && match.Success)
            {
                ThrowError(null, null, $"Chart label not supported: {label}");
                return;
            }

            series.Tag.Value = IsExpression(label) ? ConvertChartExpression(label, dataset) : label;
            labels.ValueType = StiSeriesLabelsValueType.Tag;
        }

        private void ProcessChartAreasType(XmlNode baseNode, StiChart chart)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartArea":
                        ProcessChartAreaType(node, chart);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartAreaType(XmlNode baseNode, StiChart chart)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Style":
                        ProcessChartAreaStyleType(node, chart.Area);
                        break;

                    case "ChartCategoryAxes":
                    case "CategoryAxes":
                        ProcessChartAxesType(node, chart, true);
                        break;

                    case "ChartValueAxes":
                    case "ValueAxes":
                        ProcessChartAxesType(node, chart, false);
                        break;

                    case "ChartElementPosition":
                    case "ChartInnerPlotPosition":
                    case "ChartThreeDProperties":
                    case "ThreeDProperties":
                    case "CustomProperties":
                    case "Hidden":
                    case "AlignOrientation":
                    case "EquallySizedAxesFont":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartAxesType(XmlNode baseNode, StiChart chart, bool isCategory)
        {
            var axisArea = chart.Area as IStiAxisArea;
            if (axisArea == null)
                return;

            var index = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartCategoryAxis":
                    case "ChartValueAxis":
                        ProcessChartAxisWrapperType(node, GetChartTargetAxis(axisArea, isCategory, index));
                        index++;
                        break;

                    case "ChartAxis":
                        ProcessChartAxisType(node, GetChartTargetAxis(axisArea, isCategory, index));
                        index++;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartAxisWrapperType(XmlNode baseNode, IStiAxis axis)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartAxis":
                    case "Axis":
                        ProcessChartAxisType(node, axis);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        /// <summary>
        /// Returns the axis of Stimulsoft for an axis of categories or of values. A bar area of Stimulsoft keeps
        /// its values on the axis X and its categories on the axis Y, as the bars lie.
        /// </summary>
        private static IStiAxis GetChartTargetAxis(IStiAxisArea axisArea, bool isCategory, int index)
        {
            var isBar = (axisArea as IStiArea)?.Core is StiClusteredBarAreaCoreXF;
            if (isCategory != isBar)
                return index <= 0 ? (IStiAxis)axisArea.XAxis : axisArea.XTopAxis;

            return index <= 0 ? (IStiAxis)axisArea.YAxis : axisArea.YRightAxis;
        }

        /// <summary>
        /// Converts an axis. The labels of Reporting Services fit the axis by themselves unless the auto fit is
        /// disabled: the rotation they may take is the automatic rotation of Stimulsoft; the range of the font
        /// of the auto fit has no counterpart, the font of the labels stays.
        /// </summary>
        private void ProcessChartAxisType(XmlNode baseNode, IStiAxis axis)
        {
            var autoFit = true;
            var allowRotation = false;
            var hideLabels = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Visible":
                        axis.Visible = IsTrue(GetNodeTextValue(node));
                        break;

                    case "Hidden":
                        axis.Visible = !IsTrue(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessChartAxisStyleType(node, axis);
                        break;

                    case "ChartAxisTitle":
                    case "Title":
                        ProcessChartAxisTitleType(node, axis.Title);
                        break;

                    case "LogScale":
                        axis.LogarithmicScale = IsTrue(GetNodeTextValue(node));
                        break;

                    case "Minimum":
                    case "Min":
                        SetChartAxisMinimum(axis, GetNodeTextValue(node));
                        break;

                    case "Maximum":
                    case "Max":
                        SetChartAxisMaximum(axis, GetNodeTextValue(node));
                        break;

                    case "ChartMajorGridLines":
                        ProcessChartGridLinesType(node, axis, isMinor: false);
                        break;

                    case "ChartMinorGridLines":
                        ProcessChartGridLinesType(node, axis, isMinor: true);
                        break;

                    case "MajorInterval":
                    case "Interval":
                        SetChartAxisStep(axis, GetNodeTextValue(node));
                        break;

                    case "IncludeZero":
                        axis.StartFromZero = IsTrue(GetNodeTextValue(node));
                        break;

                    //an angle of Reporting Services turns clockwise, of Stimulsoft counterclockwise
                    case "Angle":
                        if (TryGetChartAxisNumber(GetNodeTextValue(node), out var angle) && angle != 0)
                            axis.Labels.Angle = (float)-angle;
                        break;

                    case "HideLabels":
                        hideLabels = IsTrue(GetNodeTextValue(node));
                        break;

                    case "LabelsAutoFitDisabled":
                        autoFit = !IsTrue(GetNodeTextValue(node));
                        break;

                    case "AllowLabelRotation":
                        allowRotation = !string.Equals(GetNodeTextValue(node), "None",
                            StringComparison.OrdinalIgnoreCase);
                        break;

                    case "Scalar":
                    case "Margin":
                    case "MinorInterval":
                    case "IntervalType":
                    case "Location":
                    case "MinFontSize":
                    case "MaxFontSize":
                    case "ChartMajorTickMarks":
                    case "ChartMinorTickMarks":
                    case "ChartAxisScaleBreak":
                    case "CrossAt":
                    case "Interlaced":
                    case "InterlacedColor":
                    case "ChartStripLines":
                    case "MajorGridLines":
                    case "MinorGridLines":
                    case "MajorTickMarks":
                    case "MinorTickMarks":
                    case "Reverse":
                    case "VariableAutoInterval":
                    case "PreventFontShrink":
                    case "PreventFontGrow":
                    case "PreventLabelOffset":
                    case "PreventWordWrap":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (hideLabels)
                axis.Labels.Placement = StiLabelsPlacement.None;

            else if (autoFit && allowRotation && axis.Labels.Angle == 0)
                axis.Labels.Placement = StiLabelsPlacement.AutoRotation;
        }

        private void SetChartAxisMinimum(IStiAxis axis, string text)
        {
            double value;
            if (!TryGetChartAxisNumber(text, out value))
            {
                AddChartAxisLimit(axis, "Minimum", text);
                return;
            }

            axis.Range.Minimum = value;
            axis.Range.Auto = false;
        }

        private void SetChartAxisMaximum(IStiAxis axis, string text)
        {
            double value;
            if (!TryGetChartAxisNumber(text, out value))
            {
                AddChartAxisLimit(axis, "Maximum", text);
                return;
            }

            axis.Range.Maximum = value;
            axis.Range.Auto = false;
        }

        private void AddChartAxisLimit(IStiAxis axis, string limit, string text)
        {
            if (IsExpression(text))
                chartAxisLimits.Add(Tuple.Create(axis, limit, text));
        }

        /// <summary>
        /// Gives the axes the limits of the expressions, as Max(Sum(Fields!Sales.Value), "Tablix1") of the data bars
        /// of a table, which draws all the bars in one scale. The range of an axis of Stimulsoft is a number, so the
        /// limit is set when the chart is printed by a script, and the script runs in the compilation only. An
        /// aggregate of an aggregate over the region or the data set runs over the source of the chart, where the
        /// inner aggregate is a column of a group; an aggregate of a field runs over the data set.
        /// </summary>
        private void ApplyChartAxisLimits(StiChart chart, string dataset)
        {
            var limits = chartAxisLimits.ToList();
            chartAxisLimits.Clear();

            var axisArea = chart.Area as IStiAxisArea;
            if (axisArea == null || limits.Count == 0)
                return;

            var statements = new List<string>();
            foreach (var limit in limits)
            {
                var axisName = GetChartAxisName(axisArea, limit.Item1);
                var value = axisName != null ? ConvertChartAxisLimit(limit.Item3, dataset) : null;
                if (value == null)
                {
                    ThrowError(null, null, $"Chart axis limit given by an expression is not converted: {chart.Name}");
                    continue;
                }

                var range = ConvertSyntaxToCSharp
                    ? $"((Stimulsoft.Report.Chart.IStiAxisArea)((Stimulsoft.Report.Chart.StiChart)sender).Area)" +
                        $".{axisName}.Range"
                    : "CType(CType(sender, Stimulsoft.Report.Chart.StiChart).Area, " +
                        $"Stimulsoft.Report.Chart.IStiAxisArea).{axisName}.Range";
                var end = ConvertSyntaxToCSharp ? ";" : string.Empty;

                statements.Add($"{range}.{limit.Item2} = System.Convert.ToDouble({value}){end}");
                statements.Add($"{range}.Auto = {(ConvertSyntaxToCSharp ? "false" : "False")}{end}");
            }

            if (statements.Count == 0)
                return;

            var script = chart.BeforePrintEvent.Script;
            if (!string.IsNullOrEmpty(script))
                statements.Insert(0, script);

            chart.BeforePrintEvent.Script = string.Join(Environment.NewLine, statements);
            ThrowError(null, null,
                $"Chart axis limit given by an expression works in the compilation only: {chart.Name}");
        }

        private static string GetChartAxisName(IStiAxisArea axisArea, IStiAxis axis)
        {
            if (axis == axisArea.XAxis)
                return "XAxis";

            if (axis == axisArea.XTopAxis)
                return "XTopAxis";

            if (axis == axisArea.YAxis)
                return "YAxis";

            return axis == axisArea.YRightAxis ? "YRightAxis" : null;
        }

        /// <summary>
        /// Converts a limit of an axis to an expression without braces. Max or Min of an aggregate over the region
        /// or the data set is the extreme of the column of the aggregate in the source of the chart, of a field the
        /// extreme of the field in the data set; the scope of a group has no counterpart.
        /// </summary>
        private string ConvertChartAxisLimit(string text, string dataset)
        {
            var match = ChartAxisExtremeRegex.Match(text);
            if (match.Success)
            {
                var scope = match.Groups["scope"].Value;
                if (IsGroupName(scope))
                    return null;

                var function = match.Groups["function"].Value.Equals("Max", StringComparison.OrdinalIgnoreCase)
                    ? "Max"
                    : "Min";
                var inner = "=" + match.Groups["inner"].Value;
                var hasAggregate = ChartAggregateRegex.IsMatch(inner) && chartAggregatesSource != null;
                var source = hasAggregate ? chartAggregatesSource : dataset;
                var value = ToIndicatorOperand(
                    ConvertExpression(hasAggregate ? ReplaceChartAggregates(inner, source) : inner, source).Trim());

                return value != null ? $"Totals.{function}({GetDataSourceReference(source)}, {value})" : null;
            }

            return ToIndicatorOperand(ConvertExpression(text, dataset).Trim());
        }

        private static void SetChartAxisStep(IStiAxis axis, string text)
        {
            double value;
            if (TryGetChartAxisNumber(text, out value) && value > 0)
                axis.Step = (float)value;
        }

        private static bool TryGetChartAxisNumber(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = text.Trim();
            if (text.StartsWith("="))
                text = text.Substring(1).Trim();

            if (!double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value) &&
                !double.TryParse(text, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
                return false;

            return !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private void ProcessChartAxisTitleType(XmlNode baseNode, IStiAxisTitle title)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Caption":
                    case "Title":
                        title.Text = GetNodeTextValue(node);
                        break;

                    case "Style":
                        ProcessChartAxisTitleStyleType(node, title);
                        break;

                    case "Position":
                    case "Hidden":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartLegendsType(XmlNode baseNode, StiChart chart)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartLegend":
                        ProcessChartLegendType(node, chart.Legend);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        /// <summary>
        /// Converts the legend. A legend of Reporting Services is out of the plot area unless it is docked to
        /// a chart area and not outside it; the position and the layout are applied when all nodes are read.
        /// </summary>
        private void ProcessChartLegendType(XmlNode baseNode, IStiLegend legend)
        {
            legend.Visible = true;

            var position = string.Empty;
            var layout = string.Empty;
            var dockToChartArea = false;
            var dockOutsideChartArea = false;
            var insidePlotArea = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Hidden":
                        legend.Visible = !IsTrue(GetNodeTextValue(node));
                        break;

                    case "Position":
                        position = GetNodeTextValue(node);
                        break;

                    case "Layout":
                        layout = GetNodeTextValue(node);
                        break;

                    case "DockToChartArea":
                        dockToChartArea = !string.IsNullOrWhiteSpace(GetNodeTextValue(node));
                        break;

                    case "DockOutsideChartArea":
                        dockOutsideChartArea = IsTrue(GetNodeTextValue(node));
                        break;

                    case "Title":
                        legend.Title = GetNodeTextValue(node);
                        break;

                    case "ChartLegendTitle":
                        ProcessChartLegendTitleType(node, legend);
                        break;

                    case "Visible":
                        legend.Visible = IsTrue(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessChartLegendStyleType(node, legend);
                        break;

                    case "InsidePlotArea":
                        insidePlotArea = IsTrue(GetNodeTextValue(node));
                        break;

                    case "ColumnSeparatorColor":
                    case "HeaderSeparatorColor":
                        //the legend of Stimulsoft has no separators
                        break;

                    case "ChartElementPosition":
                    case "ColumnSpacing":
                    case "AutoFitTextDisabled":
                    case "MinFontSize":
                    case "HeaderSeparator":
                    case "ColumnSeparator":
                    case "InterlacedRows":
                    case "InterlacedRowsColor":
                    case "EquallySpacedItems":
                    case "Reversed":
                    case "MaxAutoSize":
                    case "TextWrapThreshold":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            var outside = !insidePlotArea && (!dockToChartArea || dockOutsideChartArea);
            var side = SetChartLegendPosition(legend, position, outside);
            legend.Direction = GetChartLegendDirection(layout, side);
        }

        /// <summary>
        /// Sets the position of the legend and returns the side it is on. The first word of the position of
        /// Reporting Services is the side, the second the alignment along the side: BottomCenter is under the
        /// plot area in the middle. The position by default is RightTop.
        /// </summary>
        private static string SetChartLegendPosition(IStiLegend legend, string position, bool outside)
        {
            var match = Regex.Match(position ?? string.Empty,
                "^(Top|Bottom|Left|Right)(Top|Bottom|Left|Right|Center)$", RegexOptions.IgnoreCase);
            var side = match.Success ? match.Groups[1].Value.ToLowerInvariant() : "right";
            var alignment = match.Success ? match.Groups[2].Value.ToLowerInvariant() : "top";

            switch (side)
            {
                case "top":
                case "bottom":
                    legend.VertAlignment = side == "top"
                        ? outside ? StiLegendVertAlignment.TopOutside : StiLegendVertAlignment.Top
                        : outside ? StiLegendVertAlignment.BottomOutside : StiLegendVertAlignment.Bottom;
                    legend.HorAlignment = alignment == "left"
                        ? StiLegendHorAlignment.Left
                        : alignment == "right" ? StiLegendHorAlignment.Right : StiLegendHorAlignment.Center;
                    break;

                default:
                    legend.HorAlignment = side == "left"
                        ? outside ? StiLegendHorAlignment.LeftOutside : StiLegendHorAlignment.Left
                        : outside ? StiLegendHorAlignment.RightOutside : StiLegendHorAlignment.Right;
                    legend.VertAlignment = alignment == "top"
                        ? StiLegendVertAlignment.Top
                        : alignment == "bottom" ? StiLegendVertAlignment.Bottom : StiLegendVertAlignment.Center;
                    break;
            }

            return side;
        }

        /// <summary>
        /// Returns the direction of the items of the legend. A row and a wide table go across, a column and a
        /// tall table down; the automatic table goes across on the top and the bottom, down on the sides.
        /// </summary>
        private static StiLegendDirection GetChartLegendDirection(string layout, string side)
        {
            //the legend of Stimulsoft with the columns counted by itself lays the items across in rows for
            //TopToBottom and down in columns for LeftToRight
            switch ((layout ?? string.Empty).ToLowerInvariant())
            {
                case "row":
                case "widetable":
                    return StiLegendDirection.TopToBottom;

                case "column":
                case "talltable":
                    return StiLegendDirection.LeftToRight;

                default:
                    return side == "top" || side == "bottom"
                        ? StiLegendDirection.TopToBottom
                        : StiLegendDirection.LeftToRight;
            }
        }

        private void ProcessChartTitlesType(XmlNode baseNode, StiChart chart, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartTitle":
                        ProcessChartTitleType(node, chart.Title, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartTitleType(XmlNode baseNode, IStiChartTitle title, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Caption":
                    case "Title":
                        title.Text = ConvertChartExpression(GetNodeTextValue(node), dataset);
                        title.Visible = !string.IsNullOrWhiteSpace(title.Text);
                        break;

                    case "Hidden":
                        title.Visible = !IsTrue(GetNodeTextValue(node));
                        break;

                    case "Position":
                        SetChartTitlePosition(title, GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessChartTitleStyleType(node, title);
                        break;

                    case "TextOrientation":
                        if (!string.Equals(GetNodeTextValue(node), "Auto", StringComparison.OrdinalIgnoreCase) &&
                            !string.Equals(GetNodeTextValue(node), "Horizontal", StringComparison.OrdinalIgnoreCase))
                        {
                            ThrowError(baseNode.Name, node.Name);
                        }
                        break;

                    //the title of Stimulsoft is always out of the plot area
                    case "ChartElementPosition":
                    case "DockToChartArea":
                    case "DockOutsideChartArea":
                    case "DockOffset":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        /// <summary>
        /// Sets the position of the title: the first word is the side, the second the alignment along the side,
        /// so TopLeft is a title over the chart at the left. The position by default is TopCenter.
        /// </summary>
        private static void SetChartTitlePosition(IStiChartTitle title, string position)
        {
            var match = Regex.Match(position ?? string.Empty,
                "^(Top|Bottom|Left|Right)(Top|Bottom|Left|Right|Center)$", RegexOptions.IgnoreCase);
            if (!match.Success)
                return;

            switch (match.Groups[1].Value.ToLowerInvariant())
            {
                case "left":
                    title.Dock = StiChartTitleDock.Left;
                    break;

                case "right":
                    title.Dock = StiChartTitleDock.Right;
                    break;

                case "bottom":
                    title.Dock = StiChartTitleDock.Bottom;
                    break;

                default:
                    title.Dock = StiChartTitleDock.Top;
                    break;
            }

            switch (match.Groups[2].Value.ToLowerInvariant())
            {
                case "left":
                case "top":
                    title.Alignment = StringAlignment.Near;
                    break;

                case "right":
                case "bottom":
                    title.Alignment = StringAlignment.Far;
                    break;

                default:
                    title.Alignment = StringAlignment.Center;
                    break;
            }
        }

        /// <summary>
        /// Reads the Border node of a style of a chart element of the schema of 2008: the color and
        /// the width, null when not given. The style None makes the color transparent.
        /// </summary>
        private Tuple<Color?, float?> ReadChartBorder(XmlNode borderNode)
        {
            Color? color = null;
            float? width = null;

            var colorText = GetChildText(borderNode, "Color");
            if (!string.IsNullOrEmpty(colorText) && !IsExpression(colorText))
                color = ParseColor(colorText);

            var widthText = GetChildText(borderNode, "Width");
            if (!string.IsNullOrEmpty(widthText) && !IsExpression(widthText))
                width = (float)ToHi(widthText);

            if (GetChildText(borderNode, "Style") == "None")
                color = Color.Transparent;

            return Tuple.Create(color, width);
        }

        /// <summary>
        /// Converts the grid lines of an axis. The lines of the category axis are vertical, of the
        /// value axis horizontal. Enabled Auto keeps the default of Stimulsoft, a color set here is
        /// kept from the style of the chart.
        /// </summary>
        private void ProcessChartGridLinesType(XmlNode baseNode, IStiAxis axis, bool isMinor)
        {
            var axisArea = axis.Area;
            if (axisArea == null)
                return;

            IStiGridLines gridLines = null;
            if (axis == axisArea.XAxis)
                gridLines = axisArea.GridLinesVert;

            else if (axis == axisArea.YAxis)
                gridLines = axisArea.GridLinesHor;

            else if (axis == axisArea.YRightAxis)
                gridLines = axisArea.GridLinesHorRight;

            if (gridLines == null)
                return;

            var enabled = GetChildText(baseNode, "Enabled");
            if (string.Equals(enabled, "True", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(enabled, "False", StringComparison.OrdinalIgnoreCase))
            {
                if (isMinor)
                    gridLines.MinorVisible = IsTrue(enabled);

                else
                    gridLines.Visible = IsTrue(enabled);
            }

            var border = GetChildNode(GetChildNode(baseNode, "Style"), "Border");
            var color = border != null ? ReadChartBorder(border).Item1 : null;
            if (!color.HasValue)
                return;

            if (isMinor)
                gridLines.MinorColor = color.Value;

            else
                gridLines.Color = color.Value;

            gridLines.AllowApplyStyle = false;
        }

        /// <summary>
        /// The caption of the title of a legend and its font and color.
        /// </summary>
        private void ProcessChartLegendTitleType(XmlNode baseNode, IStiLegend legend)
        {
            legend.Title = GetChildText(baseNode, "Caption");

            var style = GetChildNode(baseNode, "Style");
            if (style == null)
                return;

            legend.TitleFont = ProcessChartFontStyleType(style, legend.TitleFont);

            var color = GetChildText(style, "Color");
            if (!string.IsNullOrEmpty(color) && !IsExpression(color))
                legend.TitleColor = ParseColor(color);
        }

        /// <summary>
        /// The hyperlink of the action of a data point is the hyperlink of the series, a link to a bookmark a
        /// hyperlink that starts with the number sign. A drillthrough opens the other report from the series;
        /// a series of Stimulsoft gives that report the argument, the value and the tag of the point, not the
        /// parameters of Reporting Services, so the parameters are named in the list of errors.
        /// </summary>
        private void ProcessChartDataPointAction(XmlNode baseNode, StiSeries series, string dataset)
        {
            var hyperlink = baseNode.SelectSingleNode(".//*[local-name()='Hyperlink']");
            if (hyperlink != null)
            {
                series.Hyperlink = new StiHyperlinkExpression(ConvertInteractiveValue(hyperlink, dataset));
                return;
            }

            var bookmarkLink = baseNode.SelectSingleNode(".//*[local-name()='BookmarkLink']");
            if (bookmarkLink != null)
            {
                series.Hyperlink = new StiHyperlinkExpression("#" + ConvertInteractiveValue(bookmarkLink, dataset));
                return;
            }

            var reportNode = baseNode.SelectSingleNode(
                ".//*[local-name()='Drillthrough']/*[local-name()='ReportName']");
            var reportName = reportNode != null ? GetNodeTextValue(reportNode) : null;
            if (string.IsNullOrWhiteSpace(reportName) || IsExpression(reportName))
            {
                ThrowError(null, null, $"Action of a chart data point not imported: {series.CoreTitle}");
                return;
            }

            series.DrillDownEnabled = true;
            series.DrillDownReport = reportName.TrimStart('/').Split('/').Last() + ".mrt";
            ThrowError(null, null, $"Drillthrough report has to be imported too: {reportName}");

            var parameters = reportNode.ParentNode
                .SelectNodes("*[local-name()='Parameters']/*[local-name()='Parameter']")
                .Cast<XmlNode>()
                .Select(p => p.Attributes?["Name"]?.Value)
                .ToList();

            if (parameters.Count > 0)
            {
                ThrowError(null, null, "Drillthrough parameters of a chart point not imported, the report gets " +
                    $"SeriesArgument, SeriesValue and SeriesTag: {string.Join(", ", parameters)}");
            }
        }

        private void ProcessChartAreaStyleType(XmlNode baseNode, IStiArea area)
        {
            if (area == null)
                return;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "BackgroundColor":
                        area.Brush = new StiSolidBrush(ParseColor(GetChartStyleValue(node)));
                        break;

                    case "BorderColor":
                        area.BorderColor = ParseColor(GetChartStyleValue(node));
                        break;

                    case "BorderWidth":
                        area.BorderThickness = (float)ToHi(GetChartStyleValue(node));
                        break;

                    case "Border":
                        var areaBorder = ReadChartBorder(node);
                        if (areaBorder.Item1.HasValue)
                            area.BorderColor = areaBorder.Item1.Value;

                        if (areaBorder.Item2.HasValue)
                            area.BorderThickness = areaBorder.Item2.Value;
                        break;

                    //an area has no text, the fonts of the axes are in their own styles
                    case "Color":
                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                    case "TextDecoration":
                    case "TextAlign":
                    case "VerticalAlign":
                    case "Format":
                    case "BorderStyle":
                    case "BackgroundGradientType":
                    case "BackgroundGradientEndColor":
                    case "BackgroundHatchType":
                    case "ShadowColor":
                    case "ShadowOffset":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (area.Brush is StiSolidBrush solidBackground)
                area.Brush = GetBackgroundBrush(baseNode, solidBackground.Color) ?? area.Brush;
        }

        private void ProcessChartSeriesStyleType(XmlNode baseNode, StiSeries series)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                    case "BackgroundColor":
                        SetChartSeriesBrush(series, ParseColor(GetChartStyleValue(node)));
                        break;

                    case "BorderColor":
                        SetChartSeriesColorProperty(series, "BorderColor", ParseColor(GetChartStyleValue(node)));
                        break;

                    case "BorderWidth":
                        SetChartSeriesIntProperty(series, "BorderThickness", (int)Math.Round(ToHi(GetChartStyleValue(node))));
                        SetChartLineWidth(series, GetChartStyleValue(node));
                        break;

                    case "Border":
                        var seriesBorder = ReadChartBorder(node);
                        if (seriesBorder.Item1.HasValue)
                            SetChartSeriesColorProperty(series, "BorderColor", seriesBorder.Item1.Value);

                        if (seriesBorder.Item2.HasValue)
                            SetChartSeriesIntProperty(series, "BorderThickness",
                                (int)Math.Round(seriesBorder.Item2.Value));

                        SetChartLineWidth(series, GetChildText(node, "Width"));
                        SetChartLineStyle(series, GetChildText(node, "Style"));
                        break;

                    case "BorderStyle":
                        SetChartLineStyle(series, GetChartStyleValue(node));
                        break;

                    //a series and a data point have no font, the font of a label is in ChartDataLabel
                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                    case "TextDecoration":
                    case "Format":
                        break;

                    case "BackgroundGradientType":
                    case "BackgroundGradientEndColor":
                    case "BackgroundHatchType":
                    case "ShadowColor":
                    case "ShadowOffset":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            //the gradient and the hatch need the color of the series, a series of the style has none
            var brushProperty = series.GetType().GetProperty("Brush");
            if (brushProperty?.GetValue(series, null) is StiSolidBrush solidBackground && !series.AllowApplyStyle)
            {
                var backgroundBrush = GetBackgroundBrush(baseNode, solidBackground.Color);
                if (backgroundBrush != null)
                    brushProperty.SetValue(series, backgroundBrush, null);
            }

            if (GetShadow(baseNode, out _, out _))
                series.ShowShadow = true;
        }

        /// <summary>
        /// Gives the series the color of its style: the fill of a series with a fill, the color of the line of a
        /// line or a spline, which have no fill. The line kept the color of the chart style.
        /// </summary>
        private static void SetChartSeriesBrush(StiSeries series, Color color)
        {
            var property = series.GetType().GetProperty("Brush");
            if (property != null && property.PropertyType == typeof(StiBrush))
            {
                property.SetValue(series, new StiSolidBrush(color), null);
                series.AllowApplyStyle = false;
            }

            else if (property == null && series is IStiBaseLineSeries lineSeries)
            {
                lineSeries.LineColor = color;
                lineSeries.Lighting = false;
                series.AllowApplyStyle = false;
            }
        }

        /// <summary>
        /// The width of the border of a line or a spline without a fill is the width of its line, in points.
        /// </summary>
        private void SetChartLineWidth(StiSeries series, string width)
        {
            if (series.GetType().GetProperty("Brush") != null || !(series is IStiBaseLineSeries lineSeries) ||
                string.IsNullOrWhiteSpace(width) || IsExpression(width))
                return;

            lineSeries.LineWidth = (float)ToPt(width);
        }

        /// <summary>
        /// The style of the border of a series is the dash of its line: of a line, a spline, a stepped line, the
        /// outline of an area, a radar line. None, Double and Default keep the solid line, None of a line of
        /// Reporting Services still draws the line and a chart line has no double stroke.
        /// </summary>
        private static void SetChartLineStyle(StiSeries series, string style)
        {
            var property = series.GetType().GetProperty("LineStyle");
            if (property == null || property.PropertyType != typeof(StiPenStyle))
                return;

            switch (style)
            {
                case "Dotted":
                    property.SetValue(series, StiPenStyle.Dot, null);
                    break;

                case "Dashed":
                    property.SetValue(series, StiPenStyle.Dash, null);
                    break;

                case "DashDot":
                    property.SetValue(series, StiPenStyle.DashDot, null);
                    break;

                case "DashDotDot":
                    property.SetValue(series, StiPenStyle.DashDotDot, null);
                    break;

                case "Solid":
                    property.SetValue(series, StiPenStyle.Solid, null);
                    break;
            }
        }

        private static void SetChartSeriesColorProperty(StiSeries series, string propertyName, Color color)
        {
            var property = series.GetType().GetProperty(propertyName);
            if (property != null && property.PropertyType == typeof(Color))
                property.SetValue(series, color, null);
        }

        private static void SetChartSeriesIntProperty(StiSeries series, string propertyName, int value)
        {
            var property = series.GetType().GetProperty(propertyName);
            if (property != null && property.PropertyType == typeof(int))
                property.SetValue(series, value, null);
        }

        private void ProcessChartSeriesLabelsStyleType(XmlNode baseNode, IStiSeriesLabels labels)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                        labels.LabelColor = ParseColor(GetChartStyleValue(node));
                        break;

                    case "BackgroundColor":
                        labels.Brush = new StiSolidBrush(ParseColor(GetChartStyleValue(node)));
                        break;

                    case "BorderColor":
                        labels.BorderColor = ParseColor(GetChartStyleValue(node));
                        break;

                    case "Border":
                        var labelsBorder = ReadChartBorder(node);
                        if (labelsBorder.Item1 == Color.Transparent)
                        {
                            labels.DrawBorder = false;
                        }
                        else if (labelsBorder.Item1.HasValue)
                        {
                            labels.BorderColor = labelsBorder.Item1.Value;
                            labels.DrawBorder = true;
                        }
                        break;

                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                        labels.Font = ProcessChartFontStyleType(baseNode, labels.Font);
                        break;

                    case "Format":
                        var format = GetNodeTextValue(node)?.Trim();
                        if (IsExpression(format))
                            ThrowError(baseNode.Name, node.Name);

                        else if (!string.IsNullOrEmpty(format))
                            labels.Format = ToChartLabelsFormat(format);
                        break;

                    case "BorderStyle":
                    case "BorderWidth":
                    case "TextAlign":
                    case "VerticalAlign":
                    case "BackgroundGradientType":
                    case "BackgroundGradientEndColor":
                    case "BackgroundHatchType":
                    case "ShadowColor":
                    case "ShadowOffset":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (labels.Brush is StiSolidBrush solidBackground)
                labels.Brush = GetBackgroundBrush(baseNode, solidBackground.Color) ?? labels.Brush;
        }

        /// <summary>
        /// Returns the format of the labels for the format of a data label or of the labels of an axis. The format
        /// P of the labels shows the share of the value in the total, of an axis the value with the percent sign
        /// without multiplying; the format P of Reporting Services multiplies the value by 100, so P with its
        /// digits becomes the custom format of the percent.
        /// </summary>
        private static string ToChartLabelsFormat(string format)
        {
            if (format.Length > 3 || !format.StartsWith("P", StringComparison.OrdinalIgnoreCase))
                return format;

            var digits = 2;
            if (format.Length > 1 && !int.TryParse(format.Substring(1), out digits))
                return format;

            return digits > 0 ? "#,0." + new string('0', digits) + "%" : "#,0%";
        }

        private void ProcessChartAxisStyleType(XmlNode baseNode, IStiAxis axis)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                        axis.Labels.Color = ParseColor(GetChartStyleValue(node));
                        break;

                    case "BorderColor":
                        axis.LineColor = ParseColor(GetChartStyleValue(node));
                        break;

                    case "BorderWidth":
                        axis.LineWidth = (float)ToHi(GetChartStyleValue(node));
                        break;

                    case "Border":
                        var axisBorder = ReadChartBorder(node);
                        if (axisBorder.Item1.HasValue)
                            axis.LineColor = axisBorder.Item1.Value;

                        if (axisBorder.Item2.HasValue)
                            axis.LineWidth = axisBorder.Item2.Value;
                        break;

                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                        axis.Labels.Font = ProcessChartFontStyleType(baseNode, axis.Labels.Font);
                        break;

                    case "Format":
                        var format = GetChartStyleValue(node)?.Trim();
                        if (IsExpression(format))
                            ThrowError(baseNode.Name, node.Name);

                        else if (!string.IsNullOrEmpty(format))
                            axis.Labels.Format = ToChartLabelsFormat(format);
                        break;

                    case "BorderStyle":
                    case "BackgroundColor":
                    case "BackgroundGradientType":
                    case "BackgroundGradientEndColor":
                    case "BackgroundHatchType":
                    case "ShadowColor":
                    case "ShadowOffset":
                    case "TextAlign":
                    case "VerticalAlign":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartAxisTitleStyleType(XmlNode baseNode, IStiAxisTitle title)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                        title.Color = ParseColor(GetChartStyleValue(node));
                        break;

                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                        title.Font = ProcessChartFontStyleType(baseNode, title.Font);
                        break;

                    //the title of an axis has no frame
                    case "Border":
                    case "BorderColor":
                    case "BorderStyle":
                    case "BorderWidth":
                    case "TextAlign":
                    case "VerticalAlign":
                    case "BackgroundColor":
                    case "BackgroundGradientType":
                    case "BackgroundGradientEndColor":
                    case "BackgroundHatchType":
                    case "ShadowColor":
                    case "ShadowOffset":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessChartLegendStyleType(XmlNode baseNode, IStiLegend legend)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                        legend.LabelsColor = ParseColor(GetChartStyleValue(node));
                        legend.TitleColor = legend.LabelsColor;
                        break;

                    case "BackgroundColor":
                        legend.Brush = new StiSolidBrush(ParseColor(GetChartStyleValue(node)));
                        break;

                    case "BorderColor":
                        legend.BorderColor = ParseColor(GetChartStyleValue(node));
                        break;

                    case "Border":
                        var legendBorder = ReadChartBorder(node);
                        if (legendBorder.Item1.HasValue)
                            legend.BorderColor = legendBorder.Item1.Value;
                        break;

                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                        legend.Font = ProcessChartFontStyleType(baseNode, legend.Font);
                        break;

                    case "BorderStyle":
                    case "BorderWidth":
                    case "TextAlign":
                    case "VerticalAlign":
                    case "BackgroundGradientType":
                    case "BackgroundGradientEndColor":
                    case "BackgroundHatchType":
                    case "ShadowColor":
                    case "ShadowOffset":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (legend.Brush is StiSolidBrush solidBackground)
                legend.Brush = GetBackgroundBrush(baseNode, solidBackground.Color) ?? legend.Brush;
        }

        private void ProcessChartTitleStyleType(XmlNode baseNode, IStiChartTitle title)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                        title.Brush = new StiSolidBrush(ParseColor(GetChartStyleValue(node)));
                        break;

                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                        title.Font = ProcessChartFontStyleType(baseNode, title.Font);
                        break;

                    case "BackgroundGradientType":
                        CheckBackgroundGradientType(baseNode, node);
                        break;

                    //the title of Stimulsoft has no frame
                    case "Border":
                    case "BorderColor":
                    case "BorderStyle":
                    case "BorderWidth":
                    case "TextAlign":
                    case "VerticalAlign":
                    case "BackgroundGradientEndColor":
                    case "BackgroundHatchType":
                    case "ShadowColor":
                    case "ShadowOffset":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private Font ProcessChartFontStyleType(XmlNode baseNode, Font font)
        {
            if (font == null)
                font = new Font(DefaultFontFamily, 8);

            var fontFamily = font.Name;
            var fontSize = font.Size;
            var fontStyle = font.Style;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "FontFamily":
                        fontFamily = GetNodeTextValue(node);
                        break;

                    case "FontSize":
                        try
                        {
                            fontSize = (float)ToPt(GetNodeTextValue(node), true);
                        }
                        catch
                        {
                        }
                        break;

                    case "FontStyle":
                        if (GetNodeTextValue(node) == "Italic")
                            fontStyle |= FontStyle.Italic;
                        break;

                    case "FontWeight":
                        if (IsBoldFontWeight(GetNodeTextValue(node)))
                            fontStyle |= FontStyle.Bold;
                        break;
                }
            }

            return new Font(fontFamily, fontSize, fontStyle);
        }

        private string GetChartStyleValue(XmlNode baseNode)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == "Default")
                    return GetNodeTextValue(node);
            }

            return GetNodeTextValue(baseNode);
        }

        private string ConvertChartExpression(string expression, string dataset)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return string.Empty;

            return ConvertExpression(ReplaceChartAggregates(expression, dataset), dataset);
        }

        private static string UnwrapChartExpression(string expression)
        {
            if (expression != null && expression.StartsWith("{") && expression.EndsWith("}"))
                return expression.Substring(1, expression.Length - 2);

            return expression;
        }
        #endregion
    }
}
