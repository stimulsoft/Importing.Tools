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
        #region Style
        private void ProcessStyleType(XmlNode baseNode, StiComponent component, string dataSet = null)
        {
            IStiBorder border = component as IStiBorder;
            IStiBrush brush = component as IStiBrush;
            IStiFont font = component as IStiFont;
            StiText text = component as StiText;
            StiImage image = component as StiImage;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Border":
                        ProcessBorder(node, component);
                        break;

                    case "BorderColor":
                        string[] colors = ProcessBorderColorStyleWidthType(node);
                        if (border != null)
                        {
                            if (border.Border is StiAdvancedBorder)
                            {
                                StiAdvancedBorder advBorder = border.Border as StiAdvancedBorder;
                                advBorder.LeftSide.Color = ParseColor(colors[0]);
                                advBorder.RightSide.Color = ParseColor(colors[1]);
                                advBorder.TopSide.Color = ParseColor(colors[2]);
                                advBorder.BottomSide.Color = ParseColor(colors[3]);
                            }
                            else
                            {
                                border.Border.Color = ParseColor(colors[0]);
                            }
                        }
                        break;

                    case "BorderWidth":
                        string[] widths = ProcessBorderColorStyleWidthType(node);
                        if (border != null)
                        {
                            if (border.Border is StiAdvancedBorder)
                            {
                                StiAdvancedBorder advBorder = border.Border as StiAdvancedBorder;
                                if (widths[0] != null) advBorder.LeftSide.Size = ToHi(widths[0]);
                                if (widths[1] != null) advBorder.RightSide.Size = ToHi(widths[1]);
                                if (widths[2] != null) advBorder.TopSide.Size = ToHi(widths[2]);
                                if (widths[3] != null) advBorder.BottomSide.Size = ToHi(widths[3]);
                            }
                            else
                            {
                                if (widths[0] != null) border.Border.Size = ToHi(widths[0]);
                            }
                        }
                        break;

                    case "BorderStyle":
                        string[] styles = ProcessBorderColorStyleWidthType(node);
                        if (border != null)
                        {
                            if (border.Border is StiAdvancedBorder)
                            {
                                StiAdvancedBorder advBorder = border.Border as StiAdvancedBorder;
                                advBorder.LeftSide.Style = ParseStyle(styles[0]);
                                advBorder.RightSide.Style = ParseStyle(styles[1]);
                                advBorder.TopSide.Style = ParseStyle(styles[2]);
                                advBorder.BottomSide.Style = ParseStyle(styles[3]);
                            }
                            else
                            {
                                border.Border.Style = ParseStyle(styles[0]);
                            }
                        }
                        break;

                    case "PaddingLeft":
                        if (text != null)
                            text.Margins.Left = ToHi(GetNodeTextValue(node));
                        if (image != null)
                            image.Margins.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "PaddingRight":
                        if (text != null)
                            text.Margins.Right = ToHi(GetNodeTextValue(node));
                        if (image != null)
                            image.Margins.Right = ToHi(GetNodeTextValue(node));
                        break;
                    case "PaddingTop":
                        if (text != null)
                            text.Margins.Top = ToHi(GetNodeTextValue(node));
                        if (image != null)
                            image.Margins.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "PaddingBottom":
                        if (text != null)
                            text.Margins.Bottom = ToHi(GetNodeTextValue(node));
                        if (image != null)
                            image.Margins.Bottom = ToHi(GetNodeTextValue(node));
                        break;

                    case "FontSize":
                        if (font != null)
                        {
                            try
                            {
                                var nodeText = GetNodeTextValue(node);
                                if (!string.IsNullOrEmpty(nodeText) && !IsExpression(nodeText))
                                    font.Font = new Font(font.Font.Name, (float)ToPt(nodeText, true), font.Font.Style);
                                else
                                    font.Font = new Font(font.Font.Name, 8, font.Font.Style);
                            }
                            catch
                            {
                                font.Font = new Font(font.Font.Name, 8, font.Font.Style);
                            }
                        }
                        break;
                    case "FontFamily":
                        if (font != null) font.Font = new Font(GetNodeTextValue(node), font.Font.Size, font.Font.Style);
                        break;
                    case "FontStyle":
                        if (font != null)
                        {
                            if (GetNodeTextValue(node) == "Italic")
                                font.Font = new Font(font.Font.Name, font.Font.Size, font.Font.Style | FontStyle.Italic);
                        }
                        break;
                    case "FontWeight":
                        string fontWeight = GetNodeTextValue(node);
                        if ((font != null) && !string.IsNullOrEmpty(fontWeight))
                        {
                            bool needBold = IsBoldFontWeight(fontWeight);
                            if (needBold)
                            {
                                font.Font = new Font(font.Font.Name, font.Font.Size, font.Font.Style | FontStyle.Bold);
                            }
                        }
                        break;

                    case "BackgroundColor":
                        if (brush != null)
                        {
                            var stColor = GetNodeTextValue(node);
                            if (IsExpression(stColor))
                            {
                                var expr = ConvertExpression(stColor, dataSet);
                                expr = expr.Substring(1, expr.Length - 2);
                                component.Expressions.Add(new StiAppExpression("Brush", $"SolidBrushValue({expr})"));
                            }
                            else
                            {
                                var solidBrush = brush.Brush as StiSolidBrush;
                                if (solidBrush != null)
                                    solidBrush.Color = ParseColor(stColor);
                            }
                        }
                        break;

                    case "Color":
                        if (text != null)
                        {
                            var stColor = GetNodeTextValue(node);
                            if (IsExpression(stColor))
                            {
                                var expr = ConvertExpression(stColor, dataSet);
                                expr = expr.Substring(1, expr.Length - 2);
                                component.Expressions.Add(new StiAppExpression("TextBrush", $"SolidBrushValue({expr})"));
                            }
                            else
                            {
                                var solidBrush = text.TextBrush as StiSolidBrush;
                                if (solidBrush != null)
                                    solidBrush.Color = ParseColor(GetNodeTextValue(node));
                            }
                        }
                        break;

                    case "TextAlign":
                        if (text != null) text.HorAlignment = ParseTextAlign(GetNodeTextValue(node));
                        break;

                    case "VerticalAlign":
                        if (text != null) text.VertAlignment = ParseVerticalAlign(GetNodeTextValue(node));
                        break;

                    case "TextDecoration":
                        if (font != null)
                        {
                            if (GetNodeTextValue(node) == "Underline")
                                font.Font = new Font(font.Font.Name, font.Font.Size, font.Font.Style | FontStyle.Underline);
                            if (GetNodeTextValue(node) == "LineThrough")
                                font.Font = new Font(font.Font.Name, font.Font.Size, font.Font.Style | FontStyle.Strikeout);
                        }
                        break;

                    case "Direction":
                        if (text != null)
                        {
                            if (GetNodeTextValue(node) == "RTL") text.TextOptions.RightToLeft = true;
                        }
                        break;

                    case "Format":
                        string format = GetNodeTextValue(node);
                        if (!string.IsNullOrWhiteSpace(format) && text != null)
                        {
                            if (format == "P") text.TextFormat = new StiPercentageFormatService();
                            else if (format == "C") text.TextFormat = new StiCurrencyFormatService();
                            else text.TextFormat = new StiCustomFormatService(format);
                        }
                        break;

                    case "LeftBorder":
                        ProcessBorderSide(node, component, StiBorderSides.Left);
                        break;
                    case "RightBorder":
                        ProcessBorderSide(node, component, StiBorderSides.Right);
                        break;
                    case "TopBorder":
                        ProcessBorderSide(node, component, StiBorderSides.Top);
                        break;
                    case "BottomBorder":
                        ProcessBorderSide(node, component, StiBorderSides.Bottom);
                        break;

                    case "BackgroundImage":
                        ProcessBackgroundImage(node, component);
                        break;

                    //case "InteractiveHeight":
                    //case "InteractiveWidth":
                    //ignored or not implemented yet
                    //break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessBorderSide(XmlNode baseNode, StiComponent component, StiBorderSides borderSide)
        {
            IStiBorder border = component as IStiBorder;
            if (border == null) return;

            if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
            var adv = border.Border as StiAdvancedBorder;
            StiBorderSide side = new StiBorderSide();
            if (borderSide == StiBorderSides.Left) side = adv.LeftSide;
            if (borderSide == StiBorderSides.Right) side = adv.RightSide;
            if (borderSide == StiBorderSides.Top) side = adv.TopSide;
            if (borderSide == StiBorderSides.Bottom) side = adv.BottomSide;

            var tempText = new StiText();
            ProcessBorder(baseNode, tempText);

            side.Color = tempText.Border.Color;
            side.Size = tempText.Border.Size;
            side.Style = tempText.Border.Style;
        }

        private void ProcessBackgroundImage(XmlNode baseNode, StiComponent component)
        {
            StiImage image = ProcessImageType(baseNode, null, null, null);
            backgroundImages[image] = component;
        }

        private void ProcessBorder(XmlNode baseNode, StiComponent component)
        {
            IStiBorder border = component as IStiBorder;
            if (border == null) return;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                        border.Border.Color = ParseColor(GetNodeTextValue(node));
                        break;

                    case "Width":
                        string borderWidth = GetNodeTextValue(node);
                        if (!string.IsNullOrEmpty(borderWidth)) border.Border.Size = ToHi(borderWidth);
                        break;

                    case "Style":
                        border.Border.Style = ParseStyle(GetNodeTextValue(node));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private string[] ProcessBorderColorStyleWidthType(XmlNode baseNode)
        {
            string[] arrayResult = new string[4];
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Default":
                        arrayResult[0] = GetNodeTextValue(node);
                        arrayResult[1] = arrayResult[0];
                        arrayResult[2] = arrayResult[0];
                        arrayResult[3] = arrayResult[0];
                        break;
                    case "Left":
                        arrayResult[0] = GetNodeTextValue(node);
                        break;
                    case "Right":
                        arrayResult[1] = GetNodeTextValue(node);
                        break;
                    case "Top":
                        arrayResult[2] = GetNodeTextValue(node);
                        break;
                    case "Bottom":
                        arrayResult[3] = GetNodeTextValue(node);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return arrayResult;
        }

        private static StiPenStyle ParseStyle(string styleAttribute)
        {
            switch (styleAttribute)
            {
                case null:
                case "None":
                    return StiPenStyle.None;

                case "Dotted":
                    return StiPenStyle.Dot;

                case "Dashed":
                    return StiPenStyle.Dash;

                case "Double":
                    return StiPenStyle.Double;
            }
            return StiPenStyle.Solid;
        }

        private static StiTextHorAlignment ParseTextAlign(string alignAttribute)
        {
            switch (alignAttribute)
            {
                case "Center":
                    return StiTextHorAlignment.Center;

                case "Right":
                    return StiTextHorAlignment.Right;
            }
            return StiTextHorAlignment.Left;
        }

        private static StiVertAlignment ParseVerticalAlign(string alignAttribute)
        {
            switch (alignAttribute)
            {
                case "Middle":
                    return StiVertAlignment.Center;

                case "Bottom":
                    return StiVertAlignment.Bottom;
            }
            return StiVertAlignment.Top;
        }

        private static Hashtable HtmlNameToColor = null;

        private static Color ParseColor(string colorAttribute)
        {
            Color color = Color.Transparent;
            if (colorAttribute != null && colorAttribute.Length > 1)
            {
                if (colorAttribute[0] == '#')
                {
                    #region Parse RGB value in hexadecimal notation
                    string colorSt = colorAttribute.Substring(1).ToLowerInvariant();
                    StringBuilder sbc = new StringBuilder();
                    foreach (char ch in colorSt)
                    {
                        if (ch == '0' || ch == '1' || ch == '2' || ch == '3' || ch == '4' || ch == '5' || ch == '6' || ch == '7' ||
                            ch == '8' || ch == '9' || ch == 'a' || ch == 'b' || ch == 'c' || ch == 'd' || ch == 'e' || ch == 'f') sbc.Append(ch);
                    }
                    if (sbc.Length == 3)
                    {
                        colorSt = string.Format("{0}{0}{1}{1}{2}{2}", sbc[0], sbc[1], sbc[2]);
                    }
                    else
                    {
                        colorSt = sbc.ToString();
                    }
                    if (colorSt.Length == 6)
                    {
                        int colorInt = Convert.ToInt32(colorSt, 16);
                        color = Color.FromArgb(0xFF, (colorInt >> 16) & 0xFF, (colorInt >> 8) & 0xFF, colorInt & 0xFF);
                    }
                    #endregion
                }
                else
                {
                    #region Parse color keywords
                    if (HtmlNameToColor == null)
                    {
                        #region Init hashtable
                        string[,] initData = {
                        {"AliceBlue",       "#F0F8FF"},
                        {"AntiqueWhite",    "#FAEBD7"},
                        {"Aqua",        "#00FFFF"},
                        {"Aquamarine",  "#7FFFD4"},
                        {"Azure",       "#F0FFFF"},
                        {"Beige",       "#F5F5DC"},
                        {"Bisque",      "#FFE4C4"},
                        {"Black",       "#000000"},
                        {"BlanchedAlmond",  "#FFEBCD"},
                        {"Blue",        "#0000FF"},
                        {"BlueViolet",  "#8A2BE2"},
                        {"Brown",       "#A52A2A"},
                        {"BurlyWood",   "#DEB887"},
                        {"CadetBlue",   "#5F9EA0"},
                        {"Chartreuse",  "#7FFF00"},
                        {"Chocolate",   "#D2691E"},
                        {"Coral",       "#FF7F50"},
                        {"CornflowerBlue",  "#6495ED"},
                        {"Cornsilk",    "#FFF8DC"},
                        {"Crimson",     "#DC143C"},
                        {"Cyan",        "#00FFFF"},
                        {"DarkBlue",    "#00008B"},
                        {"DarkCyan",    "#008B8B"},
                        {"DarkGoldenRod",   "#B8860B"},
                        {"DarkGray",    "#A9A9A9"},
                        {"DarkGrey",    "#A9A9A9"},
                        {"DarkGreen",   "#006400"},
                        {"DarkKhaki",   "#BDB76B"},
                        {"DarkMagenta", "#8B008B"},
                        {"DarkOliveGreen",  "#556B2F"},
                        {"Darkorange",  "#FF8C00"},
                        {"DarkOrchid",  "#9932CC"},
                        {"DarkRed",     "#8B0000"},
                        {"DarkSalmon",  "#E9967A"},
                        {"DarkSeaGreen",    "#8FBC8F"},
                        {"DarkSlateBlue",   "#483D8B"},
                        {"DarkSlateGray",   "#2F4F4F"},
                        {"DarkSlateGrey",   "#2F4F4F"},
                        {"DarkTurquoise",   "#00CED1"},
                        {"DarkViolet",  "#9400D3"},
                        {"DeepPink",    "#FF1493"},
                        {"DeepSkyBlue", "#00BFFF"},
                        {"DimGray",     "#696969"},
                        {"DimGrey",     "#696969"},
                        {"DodgerBlue",  "#1E90FF"},
                        {"FireBrick",   "#B22222"},
                        {"FloralWhite", "#FFFAF0"},
                        {"ForestGreen", "#228B22"},
                        {"Fuchsia",     "#FF00FF"},
                        {"Gainsboro",   "#DCDCDC"},
                        {"GhostWhite",  "#F8F8FF"},
                        {"Gold",        "#FFD700"},
                        {"GoldenRod",   "#DAA520"},
                        {"Gray",        "#808080"},
                        {"Grey",        "#808080"},
                        {"Green",       "#008000"},
                        {"GreenYellow", "#ADFF2F"},
                        {"HoneyDew",    "#F0FFF0"},
                        {"HotPink",     "#FF69B4"},
                        {"IndianRed",   "#CD5C5C"},
                        {"Indigo",      "#4B0082"},
                        {"Ivory",       "#FFFFF0"},
                        {"Khaki",       "#F0E68C"},
                        {"Lavender",    "#E6E6FA"},
                        {"LavenderBlush",   "#FFF0F5"},
                        {"LawnGreen",   "#7CFC00"},
                        {"LemonChiffon",    "#FFFACD"},
                        {"LightBlue",   "#ADD8E6"},
                        {"LightCoral",  "#F08080"},
                        {"LightCyan",   "#E0FFFF"},
                        {"LightGoldenRodYellow",    "#FAFAD2"},
                        {"LightGray",   "#D3D3D3"},
                        {"LightGrey",   "#D3D3D3"},
                        {"LightGreen",  "#90EE90"},
                        {"LightPink",   "#FFB6C1"},
                        {"LightSalmon", "#FFA07A"},
                        {"LightSeaGreen",   "#20B2AA"},
                        {"LightSkyBlue",    "#87CEFA"},
                        {"LightSlateGray",  "#778899"},
                        {"LightSlateGrey",  "#778899"},
                        {"LightSteelBlue",  "#B0C4DE"},
                        {"LightYellow", "#FFFFE0"},
                        {"Lime",        "#00FF00"},
                        {"LimeGreen",   "#32CD32"},
                        {"Linen",       "#FAF0E6"},
                        {"Magenta",     "#FF00FF"},
                        {"Maroon",      "#800000"},
                        {"MediumAquaMarine",    "#66CDAA"},
                        {"MediumBlue",  "#0000CD"},
                        {"MediumOrchid",    "#BA55D3"},
                        {"MediumPurple",    "#9370D8"},
                        {"MediumSeaGreen",  "#3CB371"},
                        {"MediumSlateBlue", "#7B68EE"},
                        {"MediumSpringGreen",   "#00FA9A"},
                        {"MediumTurquoise", "#48D1CC"},
                        {"MediumVioletRed", "#C71585"},
                        {"MidnightBlue",    "#191970"},
                        {"MintCream",   "#F5FFFA"},
                        {"MistyRose",   "#FFE4E1"},
                        {"Moccasin",    "#FFE4B5"},
                        {"NavajoWhite", "#FFDEAD"},
                        {"Navy",        "#000080"},
                        {"OldLace",     "#FDF5E6"},
                        {"Olive",       "#808000"},
                        {"OliveDrab",   "#6B8E23"},
                        {"Orange",      "#FFA500"},
                        {"OrangeRed",   "#FF4500"},
                        {"Orchid",      "#DA70D6"},
                        {"PaleGoldenRod",   "#EEE8AA"},
                        {"PaleGreen",   "#98FB98"},
                        {"PaleTurquoise",   "#AFEEEE"},
                        {"PaleVioletRed",   "#D87093"},
                        {"PapayaWhip",  "#FFEFD5"},
                        {"PeachPuff",   "#FFDAB9"},
                        {"Peru",        "#CD853F"},
                        {"Pink",        "#FFC0CB"},
                        {"Plum",        "#DDA0DD"},
                        {"PowderBlue",  "#B0E0E6"},
                        {"Purple",      "#800080"},
                        {"Red",         "#FF0000"},
                        {"RosyBrown",   "#BC8F8F"},
                        {"RoyalBlue",   "#4169E1"},
                        {"SaddleBrown", "#8B4513"},
                        {"Salmon",      "#FA8072"},
                        {"SandyBrown",  "#F4A460"},
                        {"SeaGreen",    "#2E8B57"},
                        {"SeaShell",    "#FFF5EE"},
                        {"Sienna",      "#A0522D"},
                        {"Silver",      "#C0C0C0"},
                        {"SkyBlue",     "#87CEEB"},
                        {"SlateBlue",   "#6A5ACD"},
                        {"SlateGray",   "#708090"},
                        {"SlateGrey",   "#708090"},
                        {"Snow",        "#FFFAFA"},
                        {"SpringGreen", "#00FF7F"},
                        {"SteelBlue",   "#4682B4"},
                        {"Tan",         "#D2B48C"},
                        {"Teal",        "#008080"},
                        {"Thistle",     "#D8BFD8"},
                        {"Tomato",      "#FF6347"},
                        {"Turquoise",   "#40E0D0"},
                        {"Violet",      "#EE82EE"},
                        {"Wheat",       "#F5DEB3"},
                        {"White",       "#FFFFFF"},
                        {"WhiteSmoke",  "#F5F5F5"},
                        {"Yellow",      "#FFFF00"},
                        {"YellowGreen", "#9ACD32"}};

                        HtmlNameToColor = new Hashtable();
                        for (int index = 0; index < initData.GetLength(0); index++)
                        {
                            string key = initData[index, 0].ToLowerInvariant();
                            int colorInt = Convert.ToInt32(initData[index, 1].Substring(1), 16);
                            Color value = Color.FromArgb(0xFF, (colorInt >> 16) & 0xFF, (colorInt >> 8) & 0xFF, colorInt & 0xFF);
                            HtmlNameToColor[key] = value;
                        }
                        #endregion
                    }
                    string colorSt = colorAttribute.ToLowerInvariant();
                    if (HtmlNameToColor.ContainsKey(colorSt))
                    {
                        color = (Color)HtmlNameToColor[colorSt];
                    }
                    #endregion
                }
            }
            return color;
        }

        #endregion
    }
}
