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
using System.Drawing;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report.BarCodes;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.ShapeTypes;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The report items: text boxes, images, lines, shapes, check boxes, bar codes and containers.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Methods.Items
        /// <summary>
        /// Adds the items to a container. The coordinates of the items are moved by the offset, it is the
        /// corner of the container in the coordinates of the items.
        /// </summary>
        private void AddItems(List<StiRdlxNode> items, StiContainer container, double offsetX, double offsetY,
            StiRdlxScope scope)
        {
            foreach (var item in items)
            {
                var rect = GetRect(item);
                rect.X -= offsetX;
                rect.Y -= offsetY;

                var component = CreateItem(item, rect, scope);
                if (component != null)
                    container.Components.Add(component);
            }
        }

        /// <summary>
        /// Makes a component of a report item. Returns null for the items that are not imported.
        /// </summary>
        private StiComponent CreateItem(StiRdlxNode node, RectangleD rect, StiRdlxScope scope)
        {
            StiComponent component;

            switch ((node.ItemType ?? "").ToLowerInvariant())
            {
                case "textbox":
                    component = CreateText(node, rect, scope, isHtml: false);
                    break;

                case "richtext":
                case "formattedtext":
                    component = CreateText(node, rect, scope, isHtml: true);
                    break;

                case "inputfield":
                    component = CreateInputField(node, rect, scope);
                    break;

                case "checkbox":
                    component = CreateCheckBox(node, rect, scope);
                    break;

                case "image":
                    component = CreateImage(node, rect, scope);
                    break;

                case "line":
                    component = CreateLine(node, rect);
                    break;

                case "shape":
                    component = CreateShape(node, rect);
                    break;

                case "barcode":
                    component = CreateBarCode(node, rect, scope);
                    break;

                case "rectangle":
                case "container":
                    component = CreatePanel(node, rect, scope);
                    break;

                case "table":
                case "list":
                case "bandedlist":
                case "tablix":
                    component = CreateRegionPanel(node, rect, scope);
                    break;

                case "chart":
                case "dvchart":
                    component = CreateChart(node, rect, scope);
                    break;

                case "subreport":
                    Warn($"The subreport '{node.Get("ReportName")}' is not imported, a panel is left in its place.");
                    component = new StiPanel(rect);
                    break;

                default:
                    Warn($"The report item of the type '{node.ItemType}' is not imported.");
                    return null;
            }

            if (component == null)
                return null;

            component.Name = CreateName(node.Get("Name", "Item"));
            ApplyVisibility(node, component, scope);
            ApplyAction(node, component, scope);

            return component;
        }

        private StiText CreateText(StiRdlxNode node, RectangleD rect, StiRdlxScope scope, bool isHtml)
        {
            var text = new StiText(rect);
            text.Text.Value = ConvertValue(node.Get("Value"), scope, allowBandMarker: true);
            text.CanGrow = node.GetBool("CanGrow");
            text.CanShrink = node.GetBool("CanShrink");
            text.AllowHtmlTags = isHtml;

            ApplyTextStyle(node, text);

            return text;
        }

        private StiComponent CreateInputField(StiRdlxNode node, RectangleD rect, StiRdlxScope scope)
        {
            if (string.Equals(node.Get("InputType"), "CheckBox", StringComparison.OrdinalIgnoreCase))
            {
                var checkBox = new StiCheckBox(rect);
                checkBox.Checked.Value = "{" + ConvertToExpression(node.Get("Value", "false"), scope) + "}";
                checkBox.Editable = true;
                return checkBox;
            }

            var text = CreateText(node, rect, scope, isHtml: false);
            text.Editable = true;
            return text;
        }

        /// <summary>
        /// A check box with a text is a panel with the check mark at the left and the text beside it.
        /// </summary>
        private StiComponent CreateCheckBox(StiRdlxNode node, RectangleD rect, StiRdlxScope scope)
        {
            var size = Math.Min(rect.Height, rect.Width);
            var checkBox = new StiCheckBox(new RectangleD(rect.X, rect.Y, size, size));
            checkBox.Checked.Value = "{" + ConvertToExpression(node.Get("Checked", "false"), scope) + "}";

            var caption = node.Get("Text");
            if (string.IsNullOrEmpty(caption))
            {
                checkBox.ClientRectangle = new RectangleD(rect.X + (rect.Width - size) / 2, rect.Y, size, size);
                return checkBox;
            }

            var panel = new StiPanel(rect);
            checkBox.ClientRectangle = new RectangleD(0, 0, size, size);
            checkBox.Name = CreateName(node.Get("Name", "Item") + "Check");
            panel.Components.Add(checkBox);

            var text = new StiText(new RectangleD(size, 0, Math.Max(0, rect.Width - size), rect.Height));
            text.Name = CreateName(node.Get("Name", "Item") + "Text");
            text.Text.Value = ConvertValue(caption, scope);
            ApplyTextStyle(node, text);
            text.VertAlignment = StiVertAlignment.Center;
            panel.Components.Add(text);

            return panel;
        }

        private StiComponent CreateImage(StiRdlxNode node, RectangleD rect, StiRdlxScope scope)
        {
            var image = new StiImage(rect);
            var value = node.Get("Value", "");
            var source = node.Get("Source", "").ToLowerInvariant();

            StiRdlxNode embedded;

            if (value.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
            {
                image.ImageBytes = ParseBase64(value.Substring(value.IndexOf(',') + 1));
            }
            else if (source != "database" && embeddedImages.TryGetValue(value, out embedded))
            {
                image.ImageBytes = ParseBase64(embedded.Get("ImageData", ""));
            }
            else if (value.StartsWith("=") || isJson && value.StartsWith("{"))
            {
                var expression = "{" + ConvertToExpression(value, scope) + "}";

                if (source == "external")
                    image.ImageURL.Value = expression;

                else
                    image.ImageData.Value = expression;
            }
            else if (value.Length > 0)
            {
                image.ImageURL.Value = value;
            }

            switch (node.Get("Sizing", "").ToLowerInvariant())
            {
                case "fit":
                    image.Stretch = true;
                    break;

                case "fitproportional":
                    image.Stretch = true;
                    image.AspectRatio = true;
                    break;

                case "autosize":
                    image.CanGrow = true;
                    break;
            }

            switch (node.Get("HorizontalAlignment", "").ToLowerInvariant())
            {
                case "center":
                    image.HorAlignment = StiHorAlignment.Center;
                    break;

                case "right":
                    image.HorAlignment = StiHorAlignment.Right;
                    break;
            }

            switch (node.Get("VerticalAlignment", "").ToLowerInvariant())
            {
                case "middle":
                    image.VertAlignment = StiVertAlignment.Center;
                    break;

                case "bottom":
                    image.VertAlignment = StiVertAlignment.Bottom;
                    break;
            }

            var style = node.Child("Style");
            if (style != null)
                image.Border = GetBorder(style);

            return image;
        }

        private byte[] ParseBase64(string text)
        {
            try
            {
                return Convert.FromBase64String(text.Trim());
            }
            catch (FormatException)
            {
                Warn("An embedded image is not valid Base64 and is skipped.");
                return null;
            }
        }

        /// <summary>
        /// A line is a shape: horizontal, vertical or diagonal. JSON keeps the two points of the line,
        /// XML keeps a rectangle whose width or height may be negative.
        /// </summary>
        private StiComponent CreateLine(StiRdlxNode node, RectangleD rect)
        {
            var isDown = true;

            if (node.Has("StartPoint"))
            {
                var start = node.Child("StartPoint");
                var end = node.Child("EndPoint") ?? start;
                var dx = ToHi(end.Get("X")) - ToHi(start.Get("X"));
                var dy = ToHi(end.Get("Y")) - ToHi(start.Get("Y"));
                isDown = dx * dy >= 0;
            }
            else
            {
                isDown = ToHi(node.Get("Width")) * ToHi(node.Get("Height")) >= 0;
            }

            var style = node.Child("Style") ?? new StiRdlxNode("Style");
            var side = GetBorderSide(style, "Top");

            double points;
            var size = TryParseSize(node.Get("LineWidth"), out points) ? Math.Max(1, points * 96 / 72) : side.Size;
            var penStyle = node.Has("LineStyle") ? GetPenStyle(node.Get("LineStyle")) : side.Style;

            var shape = new StiShape();
            shape.BorderColor = GetColor(node.Get("LineColor"), side.Color);
            shape.Size = (float)size;
            shape.Style = penStyle == StiPenStyle.None ? StiPenStyle.Solid : penStyle;

            var thickness = Math.Max(1, size);

            if (rect.Height < Tolerance)
            {
                shape.ShapeType = new StiHorizontalLineShapeType();
                rect = new RectangleD(rect.X, rect.Y - thickness / 2, rect.Width, thickness);
            }
            else if (rect.Width < Tolerance)
            {
                shape.ShapeType = new StiVerticalLineShapeType();
                rect = new RectangleD(rect.X - thickness / 2, rect.Y, thickness, rect.Height);
            }
            else if (isDown)
            {
                shape.ShapeType = new StiDiagonalDownLineShapeType();
            }
            else
            {
                shape.ShapeType = new StiDiagonalUpLineShapeType();
            }

            shape.ClientRectangle = rect;
            return shape;
        }

        private StiComponent CreateShape(StiRdlxNode node, RectangleD rect)
        {
            var style = node.Child("Style") ?? new StiRdlxNode("Style");
            var side = GetBorderSide(style, "Top");

            var shape = new StiShape(rect);
            shape.Brush = GetBrush(style);
            shape.BorderColor = side.Color;
            shape.Size = (float)side.Size;
            shape.Style = side.Style;

            switch (node.Get("ShapeStyle", "").ToLowerInvariant())
            {
                case "roundrect":
                    shape.ShapeType = new StiRoundedRectangleShapeType();
                    break;

                case "ellipse":
                    shape.ShapeType = new StiOvalShapeType();
                    break;

                default:
                    shape.ShapeType = new StiRectangleShapeType();
                    break;
            }

            return shape;
        }

        private StiComponent CreateBarCode(StiRdlxNode node, RectangleD rect, StiRdlxScope scope)
        {
            var barCode = new StiBarCode(rect);
            barCode.Code.Value = ConvertValue(node.Get("Value"), scope);
            barCode.BarCodeType = GetBarCodeType(node.Get("Symbology", ""), node.Get("Name"));
            barCode.AutoScale = true;
            barCode.ShowLabelText = node.Has("CaptionLocation");

            var style = node.Child("Style");
            if (style != null)
            {
                barCode.ForeColor = GetColor(style.Get("Color"), Color.Black);
                barCode.BackColor = GetColor(style.Get("BackgroundColor"), Color.Transparent);
            }

            return barCode;
        }

        private StiBarCodeTypeService GetBarCodeType(string symbology, string name)
        {
            switch (symbology.ToLowerInvariant())
            {
                case "qrcode":
                    return new StiQRCodeBarCodeType();

                case "gs1qrcode":
                    return new StiGS1QRCodeBarCodeType();

                case "datamatrix":
                    return new StiDataMatrixBarCodeType();

                case "gs1datamatrix":
                    return new StiGS1DataMatrixBarCodeType();

                case "pdf417":
                    return new StiPdf417BarCodeType();

                case "aztec":
                    return new StiAztecBarCodeType();

                case "maxicode":
                    return new StiMaxicodeBarCodeType();

                case "code_128_a":
                    return new StiCode128aBarCodeType();

                case "code_128_b":
                    return new StiCode128bBarCodeType();

                case "code_128_c":
                    return new StiCode128cBarCodeType();

                case "":
                case "code_128auto":
                    return new StiCode128AutoBarCodeType();

                case "ean128fnc1":
                case "ucceean128":
                case "ucc_ean128":
                case "uccean128":
                    return new StiEAN128AutoBarCodeType();

                case "code39":
                case "ansi39":
                    return new StiCode39BarCodeType();

                case "code39x":
                case "ansi39x":
                    return new StiCode39ExtBarCodeType();

                case "code_93":
                    return new StiCode93BarCodeType();

                case "code93x":
                    return new StiCode93ExtBarCodeType();

                case "codabar":
                    return new StiCodabarBarCodeType();

                case "code25intlv":
                    return new StiInterleaved2of5BarCodeType();

                case "code_2_of_5":
                    return new StiStandard2of5BarCodeType();

                case "itf14":
                    return new StiITF14BarCodeType();

                case "ean_13":
                    return new StiEAN13BarCodeType();

                case "ean_8":
                    return new StiEAN8BarCodeType();

                case "jan13":
                    return new StiJan13BarCodeType();

                case "jan8":
                    return new StiJan8BarCodeType();

                case "upc_a":
                    return new StiUpcABarCodeType();

                case "upc_e0":
                case "upc_e1":
                    return new StiUpcEBarCodeType();

                case "msi":
                    return new StiMsiBarCodeType();

                case "postnet":
                    return new StiPostnetBarCodeType();

                case "intelligentmail":
                    return new StiIntelligentMail4StateBarCodeType();

                case "rm4scc":
                    return new StiRoyalMail4StateBarCodeType();

                default:
                    Warn($"The bar code '{name}' has the symbology '{symbology}' that Stimulsoft does not have, " +
                        "Code 128 is used instead.");
                    return new StiCode128AutoBarCodeType();
            }
        }

        /// <summary>
        /// A rectangle is a panel with its items.
        /// </summary>
        private StiComponent CreatePanel(StiRdlxNode node, RectangleD rect, StiRdlxScope scope)
        {
            var panel = new StiPanel(rect);
            panel.CanGrow = true;
            panel.CanBreak = true;
            ApplyContainerStyle(node, panel);
            AddItems(node.Items("ReportItems"), panel, 0, 0, scope);

            return panel;
        }

        /// <summary>
        /// A data region that shares its rows with other items is a panel with the bands of the region.
        /// </summary>
        private StiComponent CreateRegionPanel(StiRdlxNode node, RectangleD rect, StiRdlxScope scope)
        {
            var panel = new StiPanel(rect);
            panel.CanGrow = true;
            panel.CanShrink = true;
            panel.CanBreak = true;
            BuildDataRegion(node, panel, 0, scope);

            return panel;
        }

        /// <summary>
        /// A hidden item is disabled. A condition of visibility becomes a condition that disables the
        /// component. An item that is shown by a click on another one is left visible.
        /// </summary>
        private void ApplyVisibility(StiRdlxNode node, StiComponent component, StiRdlxScope scope)
        {
            var visibility = node.Child("Visibility");
            if (visibility == null || visibility.Has("ToggleItem"))
                return;

            var hidden = visibility.Get("Hidden", "");

            if (string.Equals(hidden, "true", StringComparison.OrdinalIgnoreCase))
            {
                component.Enabled = false;
            }
            else if (hidden.StartsWith("=") || isJson && hidden.StartsWith("{"))
            {
                var expression = ConvertToExpression(hidden, scope);

                component.Conditions.Add(new StiCondition(expression, Color.Black, Color.Transparent,
                    new Font(DefaultFontName, 8), false));
            }
        }

        private void ApplyAction(StiRdlxNode node, StiComponent component, StiRdlxScope scope)
        {
            var action = node.Child("Action");
            if (action != null && action.Has("Hyperlink"))
                component.Hyperlink.Value = ConvertValue(action.Get("Hyperlink"), scope);

            var bookmark = node.Get("Bookmark");
            if (!string.IsNullOrEmpty(bookmark))
                component.Bookmark.Value = ConvertValue(bookmark, scope);
        }
        #endregion
    }
}
