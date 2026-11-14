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
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Makes the fill and the shadow of a style of Reporting Services. The background of an item or of a part of
    /// a chart may be a gradient from BackgroundColor to BackgroundGradientEndColor, or a hatch of the two
    /// colors; the shadow is drawn only with an offset and a visible color. The designer writes the gradient
    /// None, the end color White and the shadow Transparent for every part of a chart, so these values are
    /// taken silently, and only a fill that cannot be made goes to the list of errors.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Methods.Fill
        /// <summary>
        /// Returns the brush of the gradient or the hatch of the style over the given background color, or null
        /// when the background is solid.
        /// </summary>
        private StiBrush GetBackgroundBrush(XmlNode styleNode, Color backColor)
        {
            var gradientType = GetStyleValue(styleNode, "BackgroundGradientType");
            var hatchType = GetStyleValue(styleNode, "BackgroundHatchType");
            var hasGradient = !string.IsNullOrWhiteSpace(gradientType) && !IsNone(gradientType);
            var hasHatch = !string.IsNullOrWhiteSpace(hatchType) && !IsNone(hatchType);
            if (!hasGradient && !hasHatch)
                return null;

            var endValue = GetStyleValue(styleNode, "BackgroundGradientEndColor");
            var expressionName = IsExpression(gradientType) ? "BackgroundGradientType" :
                IsExpression(hatchType) ? "BackgroundHatchType" :
                IsExpression(endValue) ? "BackgroundGradientEndColor" : null;
            if (expressionName != null)
            {
                ThrowError(styleNode.Name, expressionName);
                return null;
            }

            var endColor = string.IsNullOrWhiteSpace(endValue) ? Color.White : ParseColor(endValue);

            //the hatch draws its lines with the end color over the background color
            if (hasHatch)
            {
                if (Enum.TryParse(hatchType.Trim(), true, out HatchStyle hatchStyle))
                    return new StiHatchBrush(hatchStyle, endColor, backColor);

                ThrowError(styleNode.Name, "BackgroundHatchType");
                return null;
            }

            switch (gradientType.Trim().ToLowerInvariant())
            {
                case "leftright":
                    return new StiGradientBrush(backColor, endColor, 0);

                case "topbottom":
                    return new StiGradientBrush(backColor, endColor, 90);

                case "diagonalleft":
                    return new StiGradientBrush(backColor, endColor, 45);

                case "diagonalright":
                    return new StiGradientBrush(backColor, endColor, 135);

                //the end color is in the middle and the background color on both sides
                case "horizontalcenter":
                    return new StiGlareBrush(backColor, endColor, 90);

                case "verticalcenter":
                case "center":
                    return new StiGlareBrush(backColor, endColor, 0);

                default:
                    ThrowError(styleNode.Name, "BackgroundGradientType");
                    return null;
            }
        }

        /// <summary>
        /// Tells whether the style draws a shadow: a visible color and an offset above zero. Returns the color
        /// and the offset in hundredths of an inch.
        /// </summary>
        private bool GetShadow(XmlNode styleNode, out Color color, out double offset)
        {
            color = Color.Transparent;
            offset = 0;

            var colorValue = GetStyleValue(styleNode, "ShadowColor");
            var offsetValue = GetStyleValue(styleNode, "ShadowOffset");
            if (string.IsNullOrWhiteSpace(offsetValue) || IsExpression(colorValue) || IsExpression(offsetValue))
                return false;

            offset = ToHi(offsetValue);
            color = string.IsNullOrWhiteSpace(colorValue) ? Color.Black : ParseColor(colorValue);

            return offset > 0 && color.A > 0;
        }

        /// <summary>
        /// Draws the shadow of the style under the border of an item.
        /// </summary>
        private void ApplyComponentShadow(XmlNode styleNode, StiComponent component)
        {
            if (!(component is IStiBorder border) || !GetShadow(styleNode, out var color, out var offset))
                return;

            border.Border.DropShadow = true;
            border.Border.ShadowBrush = new StiSolidBrush(color);
            border.Border.ShadowSize = offset;
        }

        /// <summary>
        /// Returns the value of a property of a style, the value itself or the Default of a property given for
        /// the parts of a chart of the schema of 2005.
        /// </summary>
        private string GetStyleValue(XmlNode styleNode, string name)
        {
            var node = GetChildNode(styleNode, name);
            if (node == null)
                return null;

            var defaultNode = GetChildNode(node, "Default");

            return GetNodeTextValue(defaultNode ?? node);
        }

        private static bool IsNone(string value)
        {
            return string.Equals(value.Trim(), "None", StringComparison.OrdinalIgnoreCase);
        }
        #endregion
    }
}
