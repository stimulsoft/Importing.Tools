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
using System.Drawing;
using System.Globalization;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.TextFormats;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The sizes, colors, fonts and borders of the report items.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Consts.Style
        private const string DefaultFontName = "Arial";
        private const double DefaultFontSize = 10;
        #endregion

        #region Methods.Units
        /// <summary>
        /// Converts a size with a unit (in, cm, mm, pt, pc, px) to hundredths of an inch.
        /// </summary>
        private double ToHi(string value, double defaultValue)
        {
            double points;
            if (!TryParseSize(value, out points))
                return defaultValue;

            return Math.Round(points / 72 * 100, 2);
        }

        private double ToHi(string value)
        {
            return ToHi(value, 0);
        }

        /// <summary>
        /// Parses a size with a unit to points. A size without a unit is taken as inches.
        /// </summary>
        private static bool TryParseSize(string value, out double points)
        {
            points = 0;

            if (string.IsNullOrWhiteSpace(value) || value.StartsWith("="))
                return false;

            value = value.Trim().ToLowerInvariant();
            var factor = 72d;
            var units = new[] { "in", "cm", "mm", "pt", "pc", "px" };
            var factors = new[] { 72d, 72 / 2.54, 72 / 25.4, 1d, 12d, 0.75 };

            for (var index = 0; index < units.Length; index++)
            {
                if (!value.EndsWith(units[index]))
                    continue;

                factor = factors[index];
                value = value.Substring(0, value.Length - 2).Trim();
                break;
            }

            double number;
            if (!double.TryParse(value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                return false;

            points = number * factor;
            return true;
        }

        private RectangleD GetRect(StiRdlxNode node)
        {
            if (node.IsType("Line") && node.Has("StartPoint"))
            {
                var start = node.Child("StartPoint");
                var end = node.Child("EndPoint") ?? start;
                var x1 = ToHi(start.Get("X"));
                var y1 = ToHi(start.Get("Y"));
                var x2 = ToHi(end.Get("X"));
                var y2 = ToHi(end.Get("Y"));

                return new RectangleD(Math.Min(x1, x2), Math.Min(y1, y2), Math.Abs(x2 - x1), Math.Abs(y2 - y1));
            }

            var left = ToHi(node.Get("Left"));
            var top = ToHi(node.Get("Top"));
            var width = ToHi(node.Get("Width"));
            var height = ToHi(node.Get("Height"));

            if (width < 0)
            {
                left += width;
                width = -width;
            }

            if (height < 0)
            {
                top += height;
                height = -height;
            }

            return new RectangleD(left, top, width, height);
        }
        #endregion

        #region Methods.Style
        private void ApplyTextStyle(StiRdlxNode node, StiText text)
        {
            var style = node.Child("Style") ?? new StiRdlxNode("Style");

            text.Font = GetFont(style);
            text.TextBrush = new StiSolidBrush(GetColor(style.Get("Color"), Color.Black));
            text.Brush = GetBrush(style);
            text.Border = GetBorder(style);
            text.HorAlignment = GetHorAlignment(style.Get("TextAlign"));
            text.VertAlignment = GetVertAlignment(style.Get("VerticalAlign"));
            text.Margins = new StiMargins(
                ToHi(style.Get("PaddingLeft")),
                ToHi(style.Get("PaddingRight")),
                ToHi(style.Get("PaddingTop")),
                ToHi(style.Get("PaddingBottom")));
            text.WordWrap = !string.Equals(style.Get("WrapMode"), "NoWrap", StringComparison.OrdinalIgnoreCase);

            if (style.GetBool("ShrinkToFit"))
                text.ShrinkFontToFit = true;

            var format = style.Get("Format");
            if (!string.IsNullOrEmpty(format) && !format.StartsWith("="))
                text.TextFormat = new StiCustomFormatService(format);

            double angle;
            if (double.TryParse(style.Get("Angle"), NumberStyles.Float, CultureInfo.InvariantCulture, out angle))
                text.Angle = (float)-angle;
        }

        /// <summary>
        /// The background and the border of a container: a rectangle, a band, a data region.
        /// </summary>
        private void ApplyContainerStyle(StiRdlxNode node, StiContainer container)
        {
            var style = node.Child("Style");
            if (style == null)
                return;

            container.Brush = GetBrush(style);
            container.Border = GetBorder(style);
        }

        private Font GetFont(StiRdlxNode style)
        {
            var name = style.Get("FontFamily");
            if (string.IsNullOrEmpty(name) || name.StartsWith("="))
                name = DefaultFontName;

            double size;
            if (!TryParseSize(style.Get("FontSize"), out size) || size <= 0)
                size = DefaultFontSize;

            var fontStyle = FontStyle.Regular;

            if (IsBold(style.Get("FontWeight")))
                fontStyle |= FontStyle.Bold;

            if (string.Equals(style.Get("FontStyle"), "Italic", StringComparison.OrdinalIgnoreCase))
                fontStyle |= FontStyle.Italic;

            var decoration = style.Get("TextDecoration", "");

            if (decoration.IndexOf("Underline", StringComparison.OrdinalIgnoreCase) >= 0)
                fontStyle |= FontStyle.Underline;

            if (decoration.IndexOf("LineThrough", StringComparison.OrdinalIgnoreCase) >= 0)
                fontStyle |= FontStyle.Strikeout;

            try
            {
                return new Font(name, (float)size, fontStyle);
            }
            catch (ArgumentException)
            {
                return new Font(DefaultFontName, (float)size, fontStyle);
            }
        }

        private static bool IsBold(string weight)
        {
            if (string.IsNullOrEmpty(weight))
                return false;

            int number;
            if (int.TryParse(weight, out number))
                return number >= 600;

            switch (weight.ToLowerInvariant())
            {
                case "bold":
                case "bolder":
                case "semibold":
                case "extrabold":
                case "heavy":
                    return true;

                default:
                    return false;
            }
        }

        private StiBrush GetBrush(StiRdlxNode style)
        {
            var color = GetColor(style.Get("BackgroundColor"), Color.Transparent);
            return new StiSolidBrush(color);
        }

        private static StiTextHorAlignment GetHorAlignment(string value)
        {
            switch ((value ?? "").ToLowerInvariant())
            {
                case "center":
                    return StiTextHorAlignment.Center;

                case "right":
                    return StiTextHorAlignment.Right;

                case "justify":
                    return StiTextHorAlignment.Width;

                default:
                    return StiTextHorAlignment.Left;
            }
        }

        private static StiVertAlignment GetVertAlignment(string value)
        {
            switch ((value ?? "").ToLowerInvariant())
            {
                case "middle":
                    return StiVertAlignment.Center;

                case "bottom":
                    return StiVertAlignment.Bottom;

                default:
                    return StiVertAlignment.Top;
            }
        }

        /// <summary>
        /// Reads the border of a style. JSON keeps it as Border and TopBorder, LeftBorder and so on, each with
        /// Color, Style and Width; XML keeps it as BorderColor, BorderStyle and BorderWidth, each with Default,
        /// Top, Left and so on.
        /// </summary>
        private StiBorder GetBorder(StiRdlxNode style)
        {
            var sideNames = new[] { "Top", "Bottom", "Left", "Right" };
            var sides = new StiBorderSide[4];
            var isSame = true;
            var visible = StiBorderSides.None;
            var flags = new[] { StiBorderSides.Top, StiBorderSides.Bottom, StiBorderSides.Left, StiBorderSides.Right };
            StiBorderSide first = null;

            for (var index = 0; index < sideNames.Length; index++)
            {
                sides[index] = GetBorderSide(style, sideNames[index]);
                if (sides[index].Style == StiPenStyle.None)
                    continue;

                visible |= flags[index];

                if (first == null)
                    first = sides[index];

                else if (first.Style != sides[index].Style || first.Color != sides[index].Color ||
                    first.Size != sides[index].Size)
                    isSame = false;
            }

            if (first == null)
                return new StiBorder(StiBorderSides.None, Color.Black, 1, StiPenStyle.Solid);

            if (isSame)
                return new StiBorder(visible, first.Color, first.Size, first.Style);

            var border = new StiAdvancedBorder();
            CopySide(sides[0], border.TopSide);
            CopySide(sides[1], border.BottomSide);
            CopySide(sides[2], border.LeftSide);
            CopySide(sides[3], border.RightSide);
            return border;
        }

        private static void CopySide(StiBorderSide source, StiBorderSide target)
        {
            target.Color = source.Color;
            target.Size = source.Size;
            target.Style = source.Style;
        }

        private StiBorderSide GetBorderSide(StiRdlxNode style, string side)
        {
            var common = style.Child("Border");
            var own = style.Child(side + "Border");

            var penStyle = GetSideValue(style, common, own, side, "Style");
            var color = GetSideValue(style, common, own, side, "Color");
            var width = GetSideValue(style, common, own, side, "Width");

            double points;
            if (!TryParseSize(width, out points))
                points = 1;

            return new StiBorderSide(
                GetColor(color, Color.Black),
                Math.Max(1, Math.Round(points * 96 / 72, 1)),
                GetPenStyle(penStyle));
        }

        private static string GetSideValue(StiRdlxNode style, StiRdlxNode common, StiRdlxNode own, string side,
            string property)
        {
            var value = own != null ? own.Get(property) : null;

            if (string.IsNullOrEmpty(value) && common != null)
                value = common.Get(property);

            if (!string.IsNullOrEmpty(value))
                return value;

            var group = style.Child("Border" + property);
            if (group == null)
                return null;

            return group.Get(side, group.Get("Default"));
        }

        private static StiPenStyle GetPenStyle(string value)
        {
            switch ((value ?? "").ToLowerInvariant())
            {
                case "":
                case "none":
                    return StiPenStyle.None;

                case "dashed":
                case "dash":
                    return StiPenStyle.Dash;

                case "dotted":
                case "dot":
                    return StiPenStyle.Dot;

                case "dashdot":
                    return StiPenStyle.DashDot;

                case "dashdotdot":
                    return StiPenStyle.DashDotDot;

                case "double":
                    return StiPenStyle.Double;

                default:
                    return value.StartsWith("=") ? StiPenStyle.None : StiPenStyle.Solid;
            }
        }

        /// <summary>
        /// Parses a color: a name, #rgb, #rrggbb, #rrggbbaa, rgb() or rgba(). An expression gives the default.
        /// </summary>
        private static Color GetColor(string value, Color defaultColor)
        {
            if (string.IsNullOrWhiteSpace(value) || value.StartsWith("="))
                return defaultColor;

            value = value.Trim();

            try
            {
                if (value.StartsWith("#"))
                    return ParseHexColor(value.Substring(1), defaultColor);

                if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase))
                    return ParseRgbColor(value, defaultColor);

                if (string.Equals(value, "LightGrey", StringComparison.OrdinalIgnoreCase))
                    return Color.LightGray;

                var color = Color.FromName(value);
                return color.IsKnownColor ? color : defaultColor;
            }
            catch (FormatException)
            {
                return defaultColor;
            }
        }

        private static Color ParseHexColor(string hex, Color defaultColor)
        {
            if (hex.Length == 3 || hex.Length == 4)
            {
                var wide = "";

                foreach (var ch in hex)
                {
                    wide += new string(ch, 2);
                }

                hex = wide;
            }

            if (hex.Length != 6 && hex.Length != 8)
                return defaultColor;

            var red = Convert.ToInt32(hex.Substring(0, 2), 16);
            var green = Convert.ToInt32(hex.Substring(2, 2), 16);
            var blue = Convert.ToInt32(hex.Substring(4, 2), 16);
            var alpha = hex.Length == 8 ? Convert.ToInt32(hex.Substring(6, 2), 16) : 255;

            return Color.FromArgb(alpha, red, green, blue);
        }

        private static Color ParseRgbColor(string value, Color defaultColor)
        {
            var open = value.IndexOf('(');
            var close = value.IndexOf(')');
            if (open < 0 || close < open)
                return defaultColor;

            var parts = value.Substring(open + 1, close - open - 1).Split(',');
            if (parts.Length < 3)
                return defaultColor;

            var red = (int)double.Parse(parts[0], CultureInfo.InvariantCulture);
            var green = (int)double.Parse(parts[1], CultureInfo.InvariantCulture);
            var blue = (int)double.Parse(parts[2], CultureInfo.InvariantCulture);
            var alpha = 255;

            if (parts.Length > 3)
                alpha = (int)Math.Round(double.Parse(parts[3], CultureInfo.InvariantCulture) * 255);

            return Color.FromArgb(alpha, red, green, blue);
        }
        #endregion
    }
}
