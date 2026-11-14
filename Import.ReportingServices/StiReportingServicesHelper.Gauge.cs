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

using Stimulsoft.Base.Drawing;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.Gauge;
using Stimulsoft.Report.Components.Gauge.Primitives;
using Stimulsoft.Report.Gauge;
using Stimulsoft.Report.Gauge.Helpers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts a gauge panel of Reporting Services to a gauge of Stimulsoft. Only the first gauge of
    /// the panel is converted, in its simple form: the kind of the gauge, the range of its scale, its
    /// colored ranges, the value of its first pointer, the look of the scale and of the pointer and the
    /// labels of the panel. The frames and the other gauges and indicators of the panel are lost and
    /// named in the list of errors.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts.Gauge
        private static readonly Regex ParameterValueRegex =
            new Regex(@"^=\s*Parameters!(?<name>\w+)\.Value\s*$", RegexOptions.IgnoreCase);
        #endregion

        #region Fields.Gauge
        /// <summary>
        /// The places of the items of the panels of a gauge and an indicator, in the percents of the panel.
        /// </summary>
        private Dictionary<StiComponent, RectangleD> gaugePanelLayouts = new Dictionary<StiComponent, RectangleD>();

        /// <summary>
        /// The nodes of the gauges, their type is given when their size is known.
        /// </summary>
        private Dictionary<StiGauge, XmlNode> gaugeNodesByItem = new Dictionary<StiGauge, XmlNode>();
        #endregion

        #region Gauge
        /// <summary>
        /// Converts the first gauge and the first state indicator of a panel. A panel of one of them is that one
        /// item of the size of the panel; a panel of both is a panel with both items, laid out in the percents
        /// of the panel when its size is known, see ApplyGaugePanelLayouts.
        /// </summary>
        private void ProcessGaugePanelType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            var gaugeNodes = FindNodes(baseNode, "RadialGauge").Concat(FindNodes(baseNode, "LinearGauge")).ToList();
            var indicatorNodes = FindNodes(baseNode, "StateIndicator").ToList();
            var labelNodes = FindNodes(baseNode, "GaugeLabel")
                .Where(n => !IsTrue(GetChildText(n, "Hidden")))
                .ToList();
            var name = baseNode.Attributes["Name"]?.Value ?? baseNode.Name;

            if (gaugeNodes.Count == 0 && indicatorNodes.Count == 0)
            {
                ProcessUnsupportedItem(baseNode, container, page);
                return;
            }

            if (gaugeNodes.Count > 1 || indicatorNodes.Count > 1 || FindNodes(baseNode, "NumericIndicator").Any())
                ThrowError(null, null, $"Gauge converted in a simple form, the first gauge only: {name}");

            if ((gaugeNodes.Count == 0 || indicatorNodes.Count == 0) && labelNodes.Count == 0)
            {
                var item = gaugeNodes.Count > 0
                    ? (StiComponent)CreateGauge(baseNode, gaugeNodes[0], page, dataset, name)
                    : CreateStateIndicator(baseNode, indicatorNodes[0], page, dataset, name);

                item.Left = ToHi(GetChildText(baseNode, "Left"));
                item.Top = ToHi(GetChildText(baseNode, "Top"));
                item.Width = ToHi(GetChildText(baseNode, "Width"));
                item.Height = ToHi(GetChildText(baseNode, "Height"));

                container.Components.Add(item);
                return;
            }

            var panel = new StiPanel
            {
                Name = CheckComponentName(page.Report, name),
                Page = page,
                Left = ToHi(GetChildText(baseNode, "Left")),
                Top = ToHi(GetChildText(baseNode, "Top")),
                Width = ToHi(GetChildText(baseNode, "Width")),
                Height = ToHi(GetChildText(baseNode, "Height"))
            };

            if (gaugeNodes.Count > 0)
            {
                var gaugeItem = CreateGauge(baseNode, gaugeNodes[0], page, dataset, GetItemName(gaugeNodes[0], name));
                panel.Components.Add(gaugeItem);
                gaugePanelLayouts[gaugeItem] = GetGaugeItemRect(gaugeNodes[0]);
            }

            if (indicatorNodes.Count > 0)
            {
                var indicatorItem = CreateStateIndicator(baseNode, indicatorNodes[0], page, dataset,
                    GetItemName(indicatorNodes[0], name));
                panel.Components.Add(indicatorItem);
                gaugePanelLayouts[indicatorItem] = GetGaugeItemRect(indicatorNodes[0]);
            }

            foreach (var labelNode in labelNodes)
            {
                AddGaugeLabel(labelNode, baseNode, panel, page, dataset);
            }

            container.Components.Add(panel);
        }

        /// <summary>
        /// Converts a label of a gauge panel to a text of the panel. The style of the label is converted as of a
        /// text box of the version of 2005 at the place of the panel, the color given by an expression too; the
        /// text is converted as the value of the gauge, an aggregate runs over the data set of the panel. The
        /// place of a label is in the percents of its parent item, of the panel without it; the label is cut by
        /// the panel.
        /// </summary>
        private void AddGaugeLabel(XmlNode labelNode, XmlNode panelNode, StiPanel panel, StiPage page, string dataset)
        {
            var document = labelNode.OwnerDocument;
            var textbox = document.CreateElement("Textbox", labelNode.NamespaceURI);
            textbox.SetAttribute("Name", GetItemName(labelNode, "GaugeLabel"));

            var style = GetChildNode(labelNode, "Style");
            if (style != null)
                textbox.AppendChild(style.CloneNode(true));

            //the text box stands in the label for the time of the conversion, a text box has a parent in a report
            labelNode.AppendChild(textbox);
            var count = panel.Components.Count;
            ProcessTextboxType(textbox, panel, page, dataset);
            labelNode.RemoveChild(textbox);

            if (panel.Components.Count == count)
                return;

            var text = (StiText)panel.Components[panel.Components.Count - 1];
            text.CanGrow = false;
            gaugePanelLayouts[text] = GetGaugeLabelRect(labelNode, panelNode);

            var value = GetChildText(labelNode, "Text");
            text.Text = IsExpression(value) ? ConvertGaugeExpression(value, panelNode, dataset) : value;
        }

        /// <summary>
        /// Returns the place of a label in the percents of the panel: its own place is in the percents of its
        /// parent item, the name of the parent follows the name of the collection of the item.
        /// </summary>
        private static RectangleD GetGaugeLabelRect(XmlNode labelNode, XmlNode panelNode)
        {
            var rect = GetGaugeItemRect(labelNode);

            var parentName = GetChildText(labelNode, "ParentItem")?.Split('.').Last();
            var parentNode = string.IsNullOrEmpty(parentName)
                ? null
                : panelNode.SelectNodes(".//*[@Name]").Cast<XmlNode>()
                    .FirstOrDefault(n => n.Attributes["Name"].Value == parentName && n.LocalName != "GaugeLabel");

            if (parentNode != null)
            {
                var parentRect = GetGaugeItemRect(parentNode);
                rect = new RectangleD(
                    parentRect.Left + parentRect.Width * rect.Left / 100,
                    parentRect.Top + parentRect.Height * rect.Top / 100,
                    parentRect.Width * rect.Width / 100,
                    parentRect.Height * rect.Height / 100);
            }

            return new RectangleD(rect.Left, rect.Top,
                Math.Max(0, Math.Min(rect.Width, 100 - rect.Left)),
                Math.Max(0, Math.Min(rect.Height, 100 - rect.Top)));
        }

        private static string GetItemName(XmlNode itemNode, string panelName)
        {
            return itemNode.Attributes["Name"]?.Value ?? panelName;
        }

        /// <summary>
        /// Returns the place of a gauge or an indicator in the percents of its panel, the whole panel by default.
        /// </summary>
        private static RectangleD GetGaugeItemRect(XmlNode itemNode)
        {
            var left = ParseGaugeNumber(GetChildText(itemNode, "Left")) ?? 0;
            var top = ParseGaugeNumber(GetChildText(itemNode, "Top")) ?? 0;
            var width = ParseGaugeNumber(GetChildText(itemNode, "Width")) ?? 100 - left;
            var height = ParseGaugeNumber(GetChildText(itemNode, "Height")) ?? 100 - top;

            return new RectangleD(left, top, width, height);
        }

        /// <summary>
        /// Lays out the items of the gauge panels in the percents of their panels, then gives every gauge its type.
        /// The size of a panel in a cell of a tablix is the size of the cell, it is known only when the bands of the
        /// tablix are made, and a linear gauge of the orientation Auto lies by its own shape.
        /// </summary>
        private void ApplyGaugePanelLayouts()
        {
            foreach (var pair in gaugePanelLayouts)
            {
                var item = pair.Key;
                var rect = pair.Value;
                var panel = item.Parent;
                if (panel == null)
                    continue;

                item.Left = panel.Width * rect.Left / 100;
                item.Top = panel.Height * rect.Top / 100;
                item.Width = panel.Width * rect.Width / 100;
                item.Height = panel.Height * rect.Height / 100;
            }

            foreach (var pair in gaugeNodesByItem)
            {
                var type = GetGaugeType(pair.Value, pair.Key);
                if (type == pair.Key.Type)
                    continue;

                SetGaugeType(pair.Key, type);
                ApplyGaugeLook(pair.Key, pair.Value);
            }
        }

        /// <summary>
        /// Makes the scale of the type, the scale keeps the limits of the gauge, its ranges and the value of its
        /// indicator. The look of the scale and of the pointer is given again by ApplyGaugeLook.
        /// </summary>
        private static void SetGaugeType(StiGauge gauge, StiGaugeType type)
        {
            var oldItems = gauge.Scales.Count > 0 ? gauge.Scales[0].Items.ToArray() : new List<StiGaugeElement>();
            var rangeLists = oldItems.OfType<StiScaleRangeList>().ToList();
            var value = oldItems.OfType<StiIndicatorBase>().FirstOrDefault()?.Value?.Value;

            gauge.Type = type;
            StiGaugeV2InitHelper.Init(gauge, type, addRange: false, skipText: false);

            var scale = gauge.Scales[0];
            scale.Minimum = (float)gauge.Minimum;
            scale.Maximum = (float)gauge.Maximum;

            foreach (var rangeList in rangeLists)
            {
                InsertFirst(scale, rangeList);
            }

            if (value == null)
                return;

            foreach (var indicator in scale.Items.ToArray().OfType<StiIndicatorBase>())
            {
                indicator.Value = new StiValueExpression(value);
            }
        }

        /// <summary>
        /// Gives the scale and the first pointer of a gauge the look of Reporting Services: the kind of the
        /// pointer, their fills, widths and borders, and hides the tick marks and the labels hidden there. What is
        /// not given in the gauge keeps the look of the style of Stimulsoft.
        /// </summary>
        private void ApplyGaugeLook(StiGauge gauge, XmlNode gaugeNode)
        {
            var isRadial = gaugeNode.LocalName == "RadialGauge";
            var scaleNode = FindNodes(gaugeNode, isRadial ? "RadialScale" : "LinearScale").FirstOrDefault();
            if (scaleNode == null || gauge.Scales.Count == 0)
                return;

            var scale = gauge.Scales[0];
            ApplyScaleLook(scaleNode, scale, isRadial);

            var pointerNode = FindNodes(scaleNode, isRadial ? "RadialPointer" : "LinearPointer").FirstOrDefault();
            if (pointerNode != null)
                ApplyPointerLook(pointerNode, scale, isRadial);
        }

        /// <summary>
        /// The fill, the border and the width of a scale, and its hidden tick marks and labels. The width of a
        /// radial scale of Reporting Services is in the percents of its radius, of Stimulsoft a part of its
        /// diameter.
        /// </summary>
        private void ApplyScaleLook(XmlNode scaleNode, StiScaleBase scale, bool isRadial)
        {
            var styleNode = GetChildNode(scaleNode, "Style");
            var brush = GetGaugeBrush(styleNode);
            if (brush != null)
            {
                scale.Brush = brush;
                scale.AllowApplyStyle = false;
            }

            var borderBrush = GetGaugeBorderBrush(styleNode);
            if (borderBrush != null)
                scale.BorderBrush = borderBrush;

            var width = ParseGaugeNumber(GetChildText(scaleNode, "Width"));
            if (isRadial && width != null && !float.IsNaN(width.Value))
            {
                scale.StartWidth = width.Value / 200;
                scale.EndWidth = width.Value / 200;
            }

            if (scale is StiRadialScale radialScale)
                ApplyRadialScaleGeometry(scaleNode, radialScale);

            if (IsGaugeElementHidden(scaleNode, "GaugeMajorTickMarks"))
                RemoveScaleItems<StiRadialTickMarkMajor, StiLinearTickMarkMajor>(scale);

            if (IsGaugeElementHidden(scaleNode, "GaugeMinorTickMarks"))
                RemoveScaleItems<StiRadialTickMarkMinor, StiLinearTickMarkMinor>(scale);

            if (IsGaugeElementHidden(scaleNode, "ScaleLabels"))
            {
                RemoveScaleItems<StiRadialTickLabelMajor, StiLinearTickLabelMajor>(scale);
                RemoveScaleItems<StiRadialTickLabelMinor, StiLinearTickLabelMinor>(scale);
            }
        }

        /// <summary>
        /// The kind of the first pointer and its fill, border, width and placement. A radial pointer is a needle
        /// by default in Reporting Services; the bar of a radial gauge of Stimulsoft stands for the bar, the
        /// needle for the needle and the marker, and the other indicator is removed. The width of a bar is in the
        /// percents of the radius, as the width of the scale.
        /// </summary>
        private void ApplyPointerLook(XmlNode pointerNode, StiScaleBase scale, bool isRadial)
        {
            if (isRadial)
            {
                var isBar = string.Equals(GetChildText(pointerNode, "Type"), "Bar", StringComparison.OrdinalIgnoreCase);
                if (isBar)
                    RemoveScaleItems<StiNeedle, StiNeedle>(scale);

                else
                    RemoveScaleItems<StiRadialBar, StiRadialBar>(scale);
            }

            var indicator = scale.Items.ToArray().OfType<StiIndicatorBase>().FirstOrDefault();
            if (indicator == null)
                return;

            var styleNode = GetChildNode(pointerNode, "Style");
            var brush = GetGaugeBrush(styleNode);
            if (brush != null)
            {
                indicator.Brush = brush;
                indicator.AllowApplyStyle = false;

                if (indicator is StiNeedle needle)
                {
                    needle.CapBrush = GetGaugeBrush(GetChildNode(GetChildNode(pointerNode, "PointerCap"), "Style")) ??
                        needle.CapBrush;
                    needle.CapBorderBrush = brush;
                }
            }

            var borderBrush = GetGaugeBorderBrush(styleNode);
            if (borderBrush != null)
                indicator.BorderBrush = borderBrush;

            if (!(indicator is StiBarBase bar))
                return;

            bar.Placement = GetRangePlacement(GetChildText(pointerNode, "Placement"));

            var distance = ParseGaugeNumber(GetChildText(pointerNode, "DistanceFromScale"));
            if (distance != null && !float.IsNaN(distance.Value))
                bar.Offset = distance.Value / 100;

            var width = ParseGaugeNumber(GetChildText(pointerNode, "Width"));
            if (isRadial && width != null && !float.IsNaN(width.Value))
            {
                bar.StartWidth = width.Value / 200;
                bar.EndWidth = width.Value / 200;
            }
        }

        /// <summary>
        /// The radius, the center and the angles of a radial scale given in the gauge. The radius of Reporting
        /// Services is in the percents of the smaller side of the gauge, the radius of Stimulsoft is the part of
        /// that side taken by the diameter. The center is the pivot of the gauge in the percents. An angle of
        /// Reporting Services starts at the bottom, of Stimulsoft at the right, both go clockwise. The aspect ratio
        /// of the gauge is not kept, the gauge takes the whole place.
        /// </summary>
        private static void ApplyRadialScaleGeometry(XmlNode scaleNode, StiRadialScale scale)
        {
            var radius = ParseGaugeNumber(GetChildText(scaleNode, "Radius"));
            if (radius != null && radius > 0)
            {
                scale.RadiusMode = StiRadiusMode.Auto;
                scale.Radius = radius.Value * 2 / 100;
            }

            var gaugeNode = scaleNode.ParentNode?.ParentNode;
            var pivotX = ParseGaugeNumber(GetChildText(gaugeNode, "PivotX"));
            var pivotY = ParseGaugeNumber(GetChildText(gaugeNode, "PivotY"));
            if (pivotX != null || pivotY != null)
                scale.Center = new PointF((pivotX ?? 50) / 100, (pivotY ?? 50) / 100);

            var startAngle = ParseGaugeNumber(GetChildText(scaleNode, "StartAngle"));
            if (startAngle != null)
                scale.StartAngle = (startAngle.Value + 90) % 360;

            var sweepAngle = ParseGaugeNumber(GetChildText(scaleNode, "SweepAngle"));
            if (sweepAngle != null && sweepAngle > 0)
                scale.SweepAngle = sweepAngle.Value;
        }

        private static bool IsGaugeElementHidden(XmlNode scaleNode, string nodeName)
        {
            return IsTrue(GetChildText(GetChildNode(scaleNode, nodeName), "Hidden"));
        }

        private static void RemoveScaleItems<TRadial, TLinear>(StiScaleBase scale)
            where TRadial : StiGaugeElement
            where TLinear : StiGaugeElement
        {
            foreach (var item in scale.Items.ToArray().Where(i => i is TRadial || i is TLinear))
            {
                scale.Items.Remove(item);
            }
        }

        /// <summary>
        /// Returns the fill of an element of a gauge by its style, a gradient or a hatch over its background
        /// color; null when the style gives no color.
        /// </summary>
        private StiBrush GetGaugeBrush(XmlNode styleNode)
        {
            var colorText = styleNode != null ? GetStyleValue(styleNode, "BackgroundColor") : null;
            if (string.IsNullOrWhiteSpace(colorText))
                return null;

            if (IsExpression(colorText))
            {
                ThrowError(styleNode.Name, "BackgroundColor");
                return null;
            }

            var color = ParseColor(colorText.Trim());

            return GetBackgroundBrush(styleNode, color) ?? new StiSolidBrush(color);
        }

        /// <summary>
        /// Returns the border of an element of a gauge by its style: none for the style None, the color of the
        /// border otherwise; null when the style gives neither.
        /// </summary>
        private StiBrush GetGaugeBorderBrush(XmlNode styleNode)
        {
            var borderNode = GetChildNode(styleNode, "Border");
            if (borderNode == null)
                return null;

            var border = ReadChartBorder(borderNode);
            if (border.Item1 == null)
                return null;

            return border.Item1 == Color.Transparent
                ? (StiBrush)new StiEmptyBrush()
                : new StiSolidBrush(border.Item1.Value);
        }

        /// <summary>
        /// Puts the item first, under the other items of the scale. Insert of the collection of the items adds an
        /// item to the end, whatever the index is, so the items are added again after it.
        /// </summary>
        private static void InsertFirst(StiScaleBase scale, StiGaugeElement element)
        {
            var items = scale.Items.ToArray();
            scale.Items.Clear();
            scale.Items.Add(element);

            foreach (var item in items)
            {
                scale.Items.Add(item);
            }
        }

        /// <summary>
        /// Converts the ranges of a scale to a list of ranges under the other items of the scale. A range keeps its
        /// bounds, its widths and its distance from the scale in the percents, its placement and its fill. A bound
        /// given by a parameter takes the default value of the parameter, which a hidden parameter always has;
        /// another expression has no counterpart, the range is lost and named. A scale with ranges is drawn as a
        /// gauge of its own items, so the ranges are not spread evenly as the simple gauge spreads its own.
        /// </summary>
        private void AddScaleRanges(XmlNode scaleNode, StiGauge gauge, bool isRadial, string name)
        {
            var rangeNodes = scaleNode != null ? FindNodes(scaleNode, "ScaleRange").ToList() : new List<XmlNode>();
            var rangeList = isRadial ? (StiScaleRangeList)new StiRadialRangeList() : new StiLinearRangeList();

            foreach (var rangeNode in rangeNodes)
            {
                if (IsTrue(GetChildText(rangeNode, "Hidden")))
                    continue;

                var start = GetRangeBound(rangeNode, "StartValue", name);
                var end = GetRangeBound(rangeNode, "EndValue", name);
                if (start == null || end == null)
                    continue;

                var range = isRadial ? (StiRangeBase)new StiRadialRange() : new StiLinearRange();
                range.StartValue = start.Value;
                range.EndValue = end.Value;
                range.StartWidth = (ParseGaugeNumber(GetChildText(rangeNode, "StartWidth")) ?? 10) / 100;
                range.EndWidth = (ParseGaugeNumber(GetChildText(rangeNode, "EndWidth")) ?? 10) / 100;
                range.Offset = (ParseGaugeNumber(GetChildText(rangeNode, "DistanceFromScale")) ?? 0) / 100;
                range.Placement = GetRangePlacement(GetChildText(rangeNode, "Placement"));
                range.Brush = GetRangeBrush(GetChildNode(rangeNode, "Style"));
                range.BorderBrush = new StiEmptyBrush();

                rangeList.Ranges.Add(range);
            }

            if (rangeList.Ranges.Count > 0)
                InsertFirst(gauge.Scales[0], rangeList);
        }

        /// <summary>
        /// Returns a bound of a range: a number, or the default value of the parameter it names.
        /// </summary>
        private float? GetRangeBound(XmlNode rangeNode, string nodeName, string gaugeName)
        {
            var text = GetChildText(GetChildNode(rangeNode, nodeName), "Value")?.Trim();
            var number = ParseGaugeNumber(text);
            if (number != null && !float.IsNaN(number.Value))
                return number;

            var match = text != null ? ParameterValueRegex.Match(text) : Match.Empty;
            var parameterNode = match.Success ? FindReportParameter(match.Groups["name"].Value) : null;
            var defaultNode = parameterNode?.SelectSingleNode(
                ".//*[local-name()='DefaultValue']//*[local-name()='Value']");
            var defaultValue = defaultNode != null ? ParseGaugeNumber(defaultNode.InnerText) : null;

            if (defaultValue != null)
            {
                if (!IsTrue(GetChildText(parameterNode, "Hidden")))
                    ThrowError(null, null, $"Gauge range bound takes the default value of the parameter: {gaugeName}");

                return defaultValue;
            }

            ThrowError(null, null, $"Gauge range with a bound given by an expression is not converted: {gaugeName}");
            return null;
        }

        /// <summary>
        /// Finds a parameter in the whole document, the report node of the schema of 2010 is a report section and
        /// the parameters are out of it.
        /// </summary>
        private XmlNode FindReportParameter(string parameterName)
        {
            var documentNode = (XmlNode)reportNode.OwnerDocument ?? reportNode;

            return documentNode.SelectNodes("//*[local-name()='ReportParameter']").Cast<XmlNode>()
                .FirstOrDefault(n => string.Equals(n.Attributes?["Name"]?.Value, parameterName,
                    StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Cross of Reporting Services lies over the scale, Inside and Outside are the same.
        /// </summary>
        private static StiPlacement GetRangePlacement(string placement)
        {
            if (string.Equals(placement, "Inside", StringComparison.OrdinalIgnoreCase))
                return StiPlacement.Inside;

            if (string.Equals(placement, "Outside", StringComparison.OrdinalIgnoreCase))
                return StiPlacement.Outside;

            return StiPlacement.Overlay;
        }

        /// <summary>
        /// Returns the fill of a range, a gradient or a hatch over its background color, light gray by default.
        /// </summary>
        private StiBrush GetRangeBrush(XmlNode styleNode)
        {
            var colorText = styleNode != null ? GetStyleValue(styleNode, "BackgroundColor") : null;
            if (IsExpression(colorText))
            {
                ThrowError(styleNode.Name, "BackgroundColor");
                colorText = null;
            }

            var color = string.IsNullOrWhiteSpace(colorText) ? Color.LightGray : ParseColor(colorText.Trim());

            return (styleNode != null ? GetBackgroundBrush(styleNode, color) : null) ?? new StiSolidBrush(color);
        }

        /// <summary>
        /// Converts a gauge: the kind, the range of its scale and the value of its first pointer.
        /// </summary>
        private StiGauge CreateGauge(XmlNode baseNode, XmlNode gaugeNode, StiPage page, string dataset, string name)
        {
            var isRadial = gaugeNode.LocalName == "RadialGauge";
            var scaleNode = FindNodes(gaugeNode, isRadial ? "RadialScale" : "LinearScale").FirstOrDefault();
            var pointerNode = scaleNode != null
                ? FindNodes(scaleNode, isRadial ? "RadialPointer" : "LinearPointer").FirstOrDefault()
                : null;

            var gauge = new StiGauge
            {
                Name = CheckComponentName(page.Report, name),
                Page = page
            };
            gaugeNodesByItem[gauge] = gaugeNode;

            var value = pointerNode != null ? GetGaugeValue(pointerNode, baseNode, dataset) : null;
            var relativeValue = GetRelativeGaugeValue(scaleNode, baseNode, dataset, value);

            //the gauge sets the range of its scale when it renders, from its own limits in the custom mode
            gauge.CalculationMode = StiGaugeCalculationMode.Custom;
            gauge.Minimum = relativeValue != null ? 0 : (decimal)GetScaleLimit(scaleNode, "MinimumValue", 0, name);
            gauge.Maximum = relativeValue != null ? 100 : (decimal)GetScaleLimit(scaleNode, "MaximumValue", 100, name);
            SetGaugeType(gauge, GetGaugeType(gaugeNode, gauge));
            AddScaleRanges(scaleNode, gauge, isRadial, name);
            ApplyGaugeLook(gauge, gaugeNode);

            foreach (var indicator in gauge.Scales[0].Items.ToArray().OfType<StiIndicatorBase>())
            {
                indicator.Value = new StiValueExpression(relativeValue ?? value ?? "0");
            }

            return gauge;
        }

        /// <summary>
        /// A scale from zero to a maximum given by an expression, as the quota of the sales, shows the share of
        /// the value in the maximum: the limits of a gauge of Stimulsoft are numbers, so the scale goes from 0 to
        /// 100 and the value is in the percents of the maximum. The labels of the scale and its ranges would show
        /// the percents, not the values, so such a scale keeps the default limits, see GetScaleLimit. Returns null
        /// when the scale is not such.
        /// </summary>
        private string GetRelativeGaugeValue(XmlNode scaleNode, XmlNode panelNode, string dataset, string value)
        {
            if (scaleNode == null || !IsSingleExpression(value) || !IsGaugeElementHidden(scaleNode, "ScaleLabels") ||
                FindNodes(scaleNode, "ScaleRange").Any(n => !IsTrue(GetChildText(n, "Hidden"))))
                return null;

            var minimum = GetChildText(GetChildNode(scaleNode, "MinimumValue"), "Value");
            if (!string.IsNullOrEmpty(minimum) && ParseGaugeNumber(minimum) != 0)
                return null;

            var maximum = GetChildText(GetChildNode(scaleNode, "MaximumValue"), "Value");
            if (!IsExpression(maximum))
                return null;

            var maximumExpression = ConvertGaugeExpression(maximum, panelNode, dataset);
            if (!IsSingleExpression(maximumExpression))
                return null;

            return $"{{{DoubleWrapper}{value.Substring(1, value.Length - 2)}) * 100 / " +
                $"{DoubleWrapper}{maximumExpression.Substring(1, maximumExpression.Length - 2)})}}";
        }

        /// <summary>
        /// A radial gauge of a half circle or less is half circular. A linear gauge lies by its
        /// orientation, Auto is horizontal when the gauge is wider than high.
        /// </summary>
        private static StiGaugeType GetGaugeType(XmlNode gaugeNode, StiGauge gauge)
        {
            if (gaugeNode.LocalName == "RadialGauge")
            {
                var scaleNode = FindNodes(gaugeNode, "RadialScale").FirstOrDefault();
                var sweepAngle = scaleNode != null ? ParseGaugeNumber(GetChildText(scaleNode, "SweepAngle")) : null;
                return sweepAngle != null && sweepAngle <= 180 ? StiGaugeType.HalfCircular : StiGaugeType.FullCircular;
            }

            var orientation = GetChildText(gaugeNode, "Orientation");
            if (string.Equals(orientation, "Horizontal", StringComparison.OrdinalIgnoreCase))
                return StiGaugeType.HorizontalLinear;

            if (string.Equals(orientation, "Vertical", StringComparison.OrdinalIgnoreCase))
                return StiGaugeType.Linear;

            return gauge.Width > gauge.Height ? StiGaugeType.HorizontalLinear : StiGaugeType.Linear;
        }

        /// <summary>
        /// Returns the minimum or the maximum of a scale. A limit given by an expression has no
        /// counterpart, the default of Reporting Services is used and the loss is named.
        /// </summary>
        private float GetScaleLimit(XmlNode scaleNode, string nodeName, float defaultValue, string gaugeName)
        {
            var limitNode = scaleNode?.ChildNodes.Cast<XmlNode>().FirstOrDefault(n => n.LocalName == nodeName);
            var text = limitNode != null ? GetChildText(limitNode, "Value") : null;
            if (string.IsNullOrWhiteSpace(text))
                return defaultValue;

            var number = ParseGaugeNumber(text);
            if (number != null)
                return (float)number.Value;

            ThrowError(null, null, $"Gauge scale limit given by an expression is not used: {gaugeName}");
            return defaultValue;
        }

        /// <summary>
        /// Converts the value of a pointer. The value keeps its braces, a value without them is a string
        /// to the gauge. An aggregate without a scope runs over the data set of the panel, a gauge stands
        /// out of any band.
        /// </summary>
        private string GetGaugeValue(XmlNode pointerNode, XmlNode panelNode, string dataset)
        {
            var inputNode = FindNodes(pointerNode, "GaugeInputValue").FirstOrDefault();
            var value = inputNode != null ? GetChildText(inputNode, "Value") : null;
            if (string.IsNullOrWhiteSpace(value))
                return null;

            return ConvertGaugeExpression(value, panelNode, dataset);
        }

        /// <summary>
        /// Converts an expression of a gauge, an aggregate without a scope runs over the data set of the panel.
        /// </summary>
        private string ConvertGaugeExpression(string value, XmlNode panelNode, string dataset)
        {
            var dataSetName = MakeFullDataSetName(dataset, GetChildText(panelNode, "DataSetName"));

            var wasInDataRegion = isInDataRegion;
            isInDataRegion = false;
            var expression = ConvertExpression(value, dataSetName).Trim();
            isInDataRegion = wasInDataRegion;

            var dataSource = FindDataSetName(dataSetName);
            if (dataSource == null)
                return expression;

            var reference = GetDataSourceReference(dataSource);

            return ReplaceTotals(expression, (function, arguments, isDouble) =>
                arguments.Count == 1 && function != "Count"
                    ? MakeDouble(isDouble, $"Totals.{function}({reference}, {arguments[0].Trim()})")
                    : null);
        }

        private static float? ParseGaugeNumber(string text)
        {
            return double.TryParse(text?.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var number)
                ? (float)number
                : (float?)null;
        }

        private static IEnumerable<XmlNode> FindNodes(XmlNode node, string localName)
        {
            return node.SelectNodes($".//*[local-name()='{localName}']").Cast<XmlNode>();
        }
        #endregion
    }
}
