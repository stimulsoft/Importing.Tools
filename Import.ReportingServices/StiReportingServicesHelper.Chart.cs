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
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
        #region Fields
        private HashSet<string> secondaryValueAxisNames;
        #endregion

        #region Chart
        private void ProcessChartType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            var component = new StiChart();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;

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
            component.DataSourceName = datasetName;

            var categoryExpression = string.Empty;
            var seriesExpression = string.Empty;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartCategoryHierarchy":
                    case "CategoryGroupings":
                        categoryExpression = ProcessChartHierarchyType(node, datasetName);
                        break;

                    case "ChartSeriesHierarchy":
                    case "SeriesGroupings":
                        seriesExpression = ProcessChartSeriesHierarchyType(node, datasetName);
                        break;
                }
            }

            var chartAreas3D = GetChartAreas3D(baseNode);
            secondaryValueAxisNames = GetSecondaryValueAxisNames(baseNode);
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
                        ProcessChartDataType(node, component, datasetName, categoryExpression, seriesExpression, chartAreas3D);
                        break;

                    case "Filters":
                        component.Filters = ProcessFiltersType(node, datasetName);
                        component.FilterOn = component.Filters != null && component.Filters.Count > 0;
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

                    case "DataSetName":
                    case "ChartCategoryHierarchy":
                    case "CategoryGroupings":
                    case "ChartSeriesHierarchy":
                    case "SeriesGroupings":
                    case "Palette":
                    case "ChartBorderSkin":
                    case "ChartNoDataMessage":
                    case "NoRowsMessage":
                    case "DynamicHeight":
                    case "DynamicWidth":
                    case "DataElementName":
                    case "DataElementOutput":
                    case "ZIndex":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            container.Components.Add(component);
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

            var series = CreateChartSeries(type, subtype, IsChartSeries3D(chartAreaName, chartAreas3D));
            var name = baseNode.Attributes["Name"]?.Value;
            if (!string.IsNullOrWhiteSpace(name))
                series.CoreTitle = name;

            SetChartSeriesValueAxis(chart, series, valueAxisName);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ChartDataPoints":
                    case "DataPoints":
                        ProcessChartDataPointsType(node, series, dataset);
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

                    case "Type":
                    case "Subtype":
                    case "LegendName":
                    case "ChartAreaName":
                    case "ValueAxisName":
                    case "CategoryAxisName":
                    case "CustomProperties":
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

            if (!string.IsNullOrWhiteSpace(seriesExpression))
            {
                var autoSeriesDataColumn = UnwrapChartExpression(seriesExpression);
                if (string.IsNullOrWhiteSpace(series.AutoSeriesKeyDataColumn))
                    series.AutoSeriesKeyDataColumn = autoSeriesDataColumn;
                if (string.IsNullOrWhiteSpace(series.AutoSeriesTitleDataColumn))
                    series.AutoSeriesTitleDataColumn = autoSeriesDataColumn;
            }

            series.ShowSeriesLabels = StiShowSeriesLabels.FromChart;

            chart.Series.Add(series);
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

        private StiSeries CreateChartSeries(string type, string subtype, bool chartArea3D)
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
                    if (string.Equals(subtype, "Doughnut", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(subtype, "ExplodedDoughnut", StringComparison.OrdinalIgnoreCase))
                        return new StiDoughnutSeries();
                    if (string.Equals(subtype, "Funnel", StringComparison.OrdinalIgnoreCase) || string.Equals(subtype, "Pyramid", StringComparison.OrdinalIgnoreCase))
                        return new StiFunnelSeries();
                    if (is3D)
                        return new StiPie3dSeries();
                    return new StiPieSeries();

                case "scatter":
                    if (string.Equals(subtype, "Smooth", StringComparison.OrdinalIgnoreCase))
                        return new StiScatterSplineSeries();
                    if (string.Equals(subtype, "Line", StringComparison.OrdinalIgnoreCase))
                        return new StiScatterLineSeries();
                    return new StiScatterSeries();

                case "bubble":
                    return new StiBubbleSeries();

                case "radar":
                    if (string.Equals(subtype, "Area", StringComparison.OrdinalIgnoreCase))
                        return new StiRadarAreaSeries();
                    if (string.Equals(subtype, "Point", StringComparison.OrdinalIgnoreCase))
                        return new StiRadarPointSeries();
                    return new StiRadarLineSeries();

                case "range":
                    if (string.Equals(subtype, "Bar", StringComparison.OrdinalIgnoreCase))
                        return new StiRangeBarSeries();
                    return new StiRangeSeries();

                case "stock":
                    if (string.Equals(subtype, "Candlestick", StringComparison.OrdinalIgnoreCase))
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
                    return new StiClusteredColumnSeries();
            }
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

                    case "DataElementName":
                    case "DataElementOutput":
                    case "Action":
                    case "ToolTip":
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

        //Returns true when the marker was explicitly configured (type, visibility or color).
        //An empty marker leaves the series style untouched so the palette still applies.
        private bool ProcessChartMarkerType(XmlNode baseNode, IStiMarker marker)
        {
            if (marker == null)
                return false;

            var configured = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Type":
                        SetChartMarkerType(marker, GetNodeTextValue(node));
                        configured = true;
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
                        break;

                    case "Hidden":
                        marker.Visible = !IsTrue(GetNodeTextValue(node));
                        configured = true;
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

            return configured;
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

                    case "BorderStyle":
                    case "BorderWidth":
                    case "BackgroundGradientType":
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
                        ProcessChartAxisType(node, axis);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private static IStiAxis GetChartTargetAxis(IStiAxisArea axisArea, bool isCategory, int index)
        {
            if (isCategory)
                return index <= 0 ? (IStiAxis)axisArea.XAxis : axisArea.XTopAxis;

            return index <= 0 ? (IStiAxis)axisArea.YAxis : axisArea.YRightAxis;
        }

        private void ProcessChartAxisType(XmlNode baseNode, IStiAxis axis)
        {
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
                        SetChartAxisMinimum(axis, GetNodeTextValue(node));
                        break;

                    case "Maximum":
                        SetChartAxisMaximum(axis, GetNodeTextValue(node));
                        break;

                    case "MajorInterval":
                    case "Interval":
                        SetChartAxisStep(axis, GetNodeTextValue(node));
                        break;

                    case "IncludeZero":
                        axis.StartFromZero = IsTrue(GetNodeTextValue(node));
                        break;

                    case "Scalar":
                    case "Margin":
                    case "MinorInterval":
                    case "IntervalType":
                    case "Location":
                    case "LabelsAutoFitDisabled":
                    case "ChartMajorTickMarks":
                    case "ChartMinorTickMarks":
                    case "ChartAxisScaleBreak":
                    case "CrossAt":
                    case "Interlaced":
                    case "InterlacedColor":
                    case "ChartMajorGridLines":
                    case "ChartMinorGridLines":
                    case "ChartStripLines":
                    case "VariableAutoInterval":
                    case "PreventFontShrink":
                    case "PreventFontGrow":
                    case "PreventLabelOffset":
                    case "PreventWordWrap":
                    case "HideLabels":
                    case "Angle":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private static void SetChartAxisMinimum(IStiAxis axis, string text)
        {
            double value;
            if (!TryGetChartAxisNumber(text, out value))
                return;

            axis.Range.Minimum = value;
            axis.Range.Auto = false;
        }

        private static void SetChartAxisMaximum(IStiAxis axis, string text)
        {
            double value;
            if (!TryGetChartAxisNumber(text, out value))
                return;

            axis.Range.Maximum = value;
            axis.Range.Auto = false;
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

        private void ProcessChartLegendType(XmlNode baseNode, IStiLegend legend)
        {
            legend.Visible = true;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Hidden":
                        legend.Visible = !IsTrue(GetNodeTextValue(node));
                        break;

                    case "Position":
                        SetChartLegendPosition(legend, GetNodeTextValue(node));
                        break;

                    case "Layout":
                        legend.Direction = string.Equals(GetNodeTextValue(node), "Row", StringComparison.OrdinalIgnoreCase)
                            ? StiLegendDirection.LeftToRight
                            : StiLegendDirection.TopToBottom;
                        break;

                    case "Title":
                        legend.Title = GetNodeTextValue(node);
                        break;

                    case "Style":
                        ProcessChartLegendStyleType(node, legend);
                        break;

                    case "ChartElementPosition":
                    case "DockToChartArea":
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
        }

        private static void SetChartLegendPosition(IStiLegend legend, string position)
        {
            if (position.IndexOf("Left", StringComparison.OrdinalIgnoreCase) >= 0)
                legend.HorAlignment = StiLegendHorAlignment.Left;
            else if (position.IndexOf("Right", StringComparison.OrdinalIgnoreCase) >= 0)
                legend.HorAlignment = StiLegendHorAlignment.Right;
            else
                legend.HorAlignment = StiLegendHorAlignment.Center;

            if (position.IndexOf("Top", StringComparison.OrdinalIgnoreCase) >= 0)
                legend.VertAlignment = StiLegendVertAlignment.Top;
            else if (position.IndexOf("Bottom", StringComparison.OrdinalIgnoreCase) >= 0)
                legend.VertAlignment = StiLegendVertAlignment.Bottom;
            else
                legend.VertAlignment = StiLegendVertAlignment.Center;
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

                    case "ChartElementPosition":
                    case "DockToChartArea":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private static void SetChartTitlePosition(IStiChartTitle title, string position)
        {
            if (position.IndexOf("Left", StringComparison.OrdinalIgnoreCase) >= 0)
                title.Dock = StiChartTitleDock.Left;
            else if (position.IndexOf("Right", StringComparison.OrdinalIgnoreCase) >= 0)
                title.Dock = StiChartTitleDock.Right;
            else if (position.IndexOf("Bottom", StringComparison.OrdinalIgnoreCase) >= 0)
                title.Dock = StiChartTitleDock.Bottom;
            else
                title.Dock = StiChartTitleDock.Top;
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

                    case "BorderStyle":
                    case "BackgroundGradientType":
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
                        break;

                    case "BorderStyle":
                    case "BackgroundGradientType":
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

        private static void SetChartSeriesBrush(StiSeries series, Color color)
        {
            var property = series.GetType().GetProperty("Brush");
            if (property == null || property.PropertyType != typeof(StiBrush))
                return;

            property.SetValue(series, new StiSolidBrush(color), null);
            series.AllowApplyStyle = false;
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

                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                        labels.Font = ProcessChartFontStyleType(baseNode, labels.Font);
                        break;

                    case "Format":
                    case "BorderStyle":
                    case "BorderWidth":
                    case "TextAlign":
                    case "VerticalAlign":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
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

                    case "FontFamily":
                    case "FontSize":
                    case "FontStyle":
                    case "FontWeight":
                        axis.Labels.Font = ProcessChartFontStyleType(baseNode, axis.Labels.Font);
                        break;

                    case "BorderStyle":
                    case "BackgroundColor":
                    case "Format":
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

                    case "TextAlign":
                    case "VerticalAlign":
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
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
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

                    case "TextAlign":
                    case "VerticalAlign":
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
                font = new Font("Arial", 8);

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

            return ConvertExpression(expression, dataset);
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
