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
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Helpers;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts a state indicator of Reporting Services to an image. The indicator shows the icon of the state
    /// whose range holds its value; the image takes the icon of the state from the resources of the report by
    /// an expression of its URL, so it works in the compilation and in the interpretation. An icon of the state
    /// is the icon of Stimulsoft of the same shape and color; a shape painted by the color of the state, as a
    /// circle, takes the icon of the nearest color.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts.Indicator
        private const string IndicatorResourcePrefix = "Indicator";

        private static readonly Dictionary<string, StiIcon> IndicatorIcons =
            new Dictionary<string, StiIcon>(StringComparer.OrdinalIgnoreCase)
            {
                ["ThreeSignsCircle"] = StiIcon.CircleGreen,
                ["ThreeSignsTriangle"] = StiIcon.TriangleYellow,
                ["ThreeSignsDiamond"] = StiIcon.RhombRed,
                ["ThreeSymbolCheck"] = StiIcon.CircleCheckGreen,
                ["ThreeSymbolExclamation"] = StiIcon.CircleExclamationYellow,
                ["ThreeSymbolCross"] = StiIcon.CircleCrossRed,
                ["ThreeSymbolUnCircledCheck"] = StiIcon.CheckGreen,
                ["ThreeSymbolUnCircledExclamation"] = StiIcon.ExclamationYellow,
                ["ThreeSymbolUnCircledCross"] = StiIcon.CrossRed
            };
        #endregion

        #region Methods.Indicator
        /// <summary>
        /// Converts a state indicator to an image. In a data region the value aggregates the rows of the cell, as
        /// Reporting Services does; out of it the rows of the data set of the panel, as the value of a gauge.
        /// </summary>
        private StiImage CreateStateIndicator(XmlNode baseNode, XmlNode indicatorNode, StiPage page, string dataset,
            string name)
        {
            var image = new StiImage
            {
                Name = CheckComponentName(page.Report, name),
                Page = page,
                Stretch = true,
                AspectRatio = true,
                HorAlignment = StiHorAlignment.Center,
                VertAlignment = StiVertAlignment.Center
            };

            var value = GetIndicatorValue(indicatorNode, baseNode, dataset);
            var url = value != null ? MakeIndicatorUrl(indicatorNode, value, dataset, name, page.Report) : null;
            if (url != null)
                image.ImageURL = new StiImageURLExpression($"{{{url}}}");

            else
                ThrowError(null, null, $"State indicator without a value or states: {name}");

            return image;
        }

        /// <summary>
        /// Returns the value of the indicator as an expression without braces.
        /// </summary>
        private string GetIndicatorValue(XmlNode indicatorNode, XmlNode panelNode, string dataset)
        {
            var inputNode = FindNodes(indicatorNode, "GaugeInputValue").FirstOrDefault();
            var value = inputNode != null ? GetChildText(inputNode, "Value") : null;
            if (string.IsNullOrWhiteSpace(value))
                return null;

            var expression = isInDataRegion
                ? ConvertExpression(value, dataset).Trim()
                : GetGaugeValue(indicatorNode, panelNode, dataset);

            //Reporting Services takes the value as a number, True is 1; True of VB is -1
            var operand = ToIndicatorOperand(expression);

            return operand != null ? $"System.Convert.ToDouble({operand})" : null;
        }

        /// <summary>
        /// Makes the URL of the icon of the state whose range holds the value, the first one in the order of the
        /// states, or an empty URL when no range holds it. A state with a bound that is not a number or an
        /// expression is skipped and named.
        /// </summary>
        private string MakeIndicatorUrl(XmlNode indicatorNode, string value, string dataset, string name,
            StiReport report)
        {
            var states = FindNodes(indicatorNode, "IndicatorState").ToList();
            var url = "\"\"";

            for (var index = states.Count - 1; index >= 0; index--)
            {
                var icon = GetIndicatorIcon(states[index], name);
                if (icon == StiIcon.None)
                    continue;

                var start = GetIndicatorBound(GetChildNode(states[index], "StartValue"), dataset, out var isStartValid);
                var end = GetIndicatorBound(GetChildNode(states[index], "EndValue"), dataset, out var isEndValid);
                if (!isStartValid || !isEndValid)
                {
                    ThrowError(null, null, $"State indicator bound not converted: {name}");
                    continue;
                }

                var conditions = new List<string>();
                if (start != null)
                    conditions.Add($"{value} >= {start}");

                if (end != null)
                    conditions.Add($"{value} <= {end}");

                var iconUrl = $"\"{StiHyperlinkProcessor.ResourceIdent}{AddIndicatorIcon(report, icon)}\"";
                url = conditions.Count == 0
                    ? iconUrl
                    : MakeIif(string.Join(ConvertSyntaxToCSharp ? " && " : " AndAlso ", conditions), iconUrl, url);
            }

            return states.Count > 0 ? url : null;
        }

        /// <summary>
        /// Returns a bound of a state: a number, an expression without braces, or null when it is not given.
        /// </summary>
        private string GetIndicatorBound(XmlNode boundNode, string dataset, out bool isValid)
        {
            isValid = true;

            var text = boundNode != null ? GetChildText(boundNode, "Value")?.Trim() : null;
            if (string.IsNullOrEmpty(text) || string.Equals(text, "NaN", StringComparison.OrdinalIgnoreCase))
                return null;

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                return text;

            var bound = text.StartsWith("=") ? ToIndicatorOperand(ConvertExpression(text, dataset).Trim()) : null;
            isValid = bound != null;

            return bound;
        }

        /// <summary>
        /// Makes an operand of a converted expression: the expression in braces, or a number.
        /// </summary>
        private static string ToIndicatorOperand(string expression)
        {
            if (IsSingleExpression(expression))
                return expression.Substring(1, expression.Length - 2);

            return double.TryParse(expression, NumberStyles.Float, CultureInfo.InvariantCulture, out _)
                ? expression
                : null;
        }

        /// <summary>
        /// Returns the icon of a state by its style; a style painted by the color of the state takes the icon of
        /// the nearest color, and the loss of the shape is named.
        /// </summary>
        private StiIcon GetIndicatorIcon(XmlNode stateNode, string name)
        {
            var style = GetChildText(stateNode, "IndicatorStyle")?.Trim();
            if (style != null && IndicatorIcons.TryGetValue(style, out var icon))
                return icon;

            if (!string.Equals(style, "Circle", StringComparison.OrdinalIgnoreCase))
                ThrowError(null, null, $"State indicator style shown as a circle of its color: {style}, {name}");

            return GetCircleIcon(ParseColor(GetChildText(stateNode, "Color")?.Trim()));
        }

        /// <summary>
        /// Returns the circle of the nearest color: red, yellow, green, or black for a gray or another color.
        /// </summary>
        private static StiIcon GetCircleIcon(Color color)
        {
            if (color.A == 0 || color.GetSaturation() < 0.25f)
                return StiIcon.CircleBlack;

            var hue = color.GetHue();
            if (hue < 20 || hue >= 330)
                return StiIcon.CircleRed;

            if (hue < 75)
                return StiIcon.CircleYellow;

            return hue < 170 ? StiIcon.CircleGreen : StiIcon.CircleBlack;
        }

        /// <summary>
        /// Adds the image of the icon to the resources of the report once and returns the name of the resource.
        /// </summary>
        private static string AddIndicatorIcon(StiReport report, StiIcon icon)
        {
            var resourceName = IndicatorResourcePrefix + icon;
            if (report.Dictionary.Resources[resourceName] == null)
            {
                var bytes = StiImageConverter.ImageToBytes(StiIconSetHelper.GetIcon(icon));
                report.Dictionary.Resources.Add(new StiResource(resourceName, StiResourceType.Image, bytes));
            }

            return resourceName;
        }
        #endregion
    }
}
