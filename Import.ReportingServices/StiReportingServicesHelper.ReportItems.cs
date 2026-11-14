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
using Stimulsoft.Base.Helpers;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.TextFormats;
using Stimulsoft.Report.Dictionary;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
        #region Fields.Rectangles
        /// <summary>
        /// The panels of the rectangles.
        /// </summary>
        private List<StiPanel> rectanglePanels = new List<StiPanel>();
        #endregion

        #region Report items
        private void ProcessReportItemsType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            int colspan = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ColSpan":
                        colspan = Convert.ToInt32(GetNodeTextValue(node));
                        break;
                }
            }
            if (colspan > 0)
            {
                container.TagValue = colspan;
            }
            else
            {
                container.TagValue = null;
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Textbox":
                        ProcessTextboxType(node, container, page, dataset);
                        break;
                    case "Rectangle":
                        ProcessRectangleType(node, container, page, dataset);
                        break;
                    case "Table":
                        ProcessTableType(node, container, page, dataset);
                        break;
                    case "Tablix":
                        ProcessTablixType(node, container, page, dataset);
                        break;
                    case "Image":
                        ProcessImageType(node, container, page, dataset);
                        break;
                    case "List":
                        ProcessListType(node, container, page, dataset);
                        break;
                    case "Line":
                        ProcessLineType(node, container, page, dataset);
                        break;
                    case "Subreport":
                        ProcessSubreportType(node, container, page, dataset);
                        break;
                    case "Chart":
                        {
                            //a chart counts by its own categories, Count in it keeps the data source
                            var wasInDataRegion = isInDataRegion;
                            isInDataRegion = false;
                            ProcessChartType(node, container, page, dataset);
                            isInDataRegion = wasInDataRegion;
                        }
                        break;
                    case "FixedPage":
                        ProcessFixedPageType(node, container, page);
                        break;
                    case "CustomReportItem":
                        ProcessCustomReportItemType(node, container, page, dataset);
                        break;

                    case "GaugePanel":
                        ProcessGaugePanelType(node, container, page, dataset);
                        break;

                    case "Map":
                        ProcessUnsupportedItem(node, container, page);
                        break;

                    case "RowSpan":
                        //the rows of a tablix are bands of their own, a cell cannot cross them
                        if (Convert.ToInt32(GetNodeTextValue(node)) > 1)
                            ThrowError(null, null, $"Cell merged across rows is split: {GetNodeTextValue(node)} rows");
                        break;

                    case "ColSpan":
                        //already processed
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }

                if ((container.TagValue != null) && (container.Components.Count > 0))
                {
                    container.Components[container.Components.Count - 1].TagValue = container.TagValue;
                }
            }

            container.TagValue = null;
        }
        #endregion

        /// <summary>
        /// A gauge panel, with the indicators in it, and a map have no counterpart that is filled
        /// from their definition. An empty panel of the same name and size keeps their place in the
        /// layout, so the items around them stay where they were, and the item is named in the list
        /// of errors.
        /// </summary>
        private void ProcessUnsupportedItem(XmlNode baseNode, StiContainer container, StiPage page)
        {
            var name = baseNode.Attributes["Name"]?.Value ?? baseNode.Name;
            var component = new StiPanel
            {
                Name = CheckComponentName(page.Report, name),
                Page = page,
                Left = ToHi(GetChildText(baseNode, "Left")),
                Top = ToHi(GetChildText(baseNode, "Top")),
                Width = ToHi(GetChildText(baseNode, "Width")),
                Height = ToHi(GetChildText(baseNode, "Height"))
            };

            ThrowError(null, null, $"{baseNode.Name} not supported, an empty panel keeps its place: {name}");

            container.Components.Add(component);
        }

        #region Image
        private StiImage ProcessImageType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiImage component = new StiImage();
            string name = "Image";
            if (baseNode.Attributes["Name"] != null)
            {
                name = baseNode.Attributes["Name"].Value;
            }
            component.Name = CheckComponentName(report, name);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            string source = null;
            string value = null;
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
                        ProcessStyleType(node, component, dataset);
                        break;

                    case "Sizing":
                        string sizing = Convert.ToString(GetNodeTextValue(node));
                        if (sizing == "Fit")
                        {
                            component.Stretch = true;
                        }
                        else if (sizing == "FitProportional")
                        {
                            component.Stretch = true;
                            component.AspectRatio = true;
                        }
                        else if (sizing == "AutoSize")
                        {
                            component.CanGrow = true;
                            component.CanShrink = true;
                        }
                        break;

                    case "Source":
                        source = Convert.ToString(GetNodeTextValue(node));
                        break;
                    case "Value":
                        value = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "Visibility":
                        ProcessVisibility(node, component, dataset);
                        break;

                    case "MIMEType":
                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        if (!ProcessInteractivityNode(node, component, dataset))
                            ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            if (source != null && value != null)
            {
                if (source == "Embedded")
                {
                    var image = (byte[])embeddedImages[value];
                    if (image != null) component.ImageBytes = image;
                }
                else if (source == "External")
                {
                    component.ImageURL = new StiImageURLExpression(ConvertExpression(value, dataset));
                }
                else if (source == "Database")
                {
                    string datacolumn = ConvertExpression(value, dataset);
                    if (datacolumn.StartsWith("{")) datacolumn = datacolumn.Substring(1, datacolumn.Length - 2);
                    component.DataColumn = datacolumn;
                }
            }
            if (container != null)
            {
                container.Components.Add(component);
            }

            return component;
        }
        #endregion

        #region Textbox
        private void ProcessTextboxType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiText component = new StiText();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Font = new Font(DefaultFontFamily, 10);
            component.Page = page;
            component.WordWrap = true;

            lastTextComponent = component;
            currentTextboxName = baseNode.Attributes["Name"].Value;

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

                    case "CanGrow":
                        component.CanGrow = Convert.ToBoolean(GetNodeTextValue(node));
                        break;
                    case "CanShrink":
                        component.CanShrink = Convert.ToBoolean(GetNodeTextValue(node));
                        break;

                    case "Value":
                        component.Text = GetNodeTextValue(node);
                        break;

                    case "Paragraphs":
                        component.Text = ProcessParagraphsType(node, dataset);
                        break;

                    case "Style":
                        ProcessTextboxStyle(node, baseNode, component, dataset);
                        break;

                    case "Visibility":
                        ProcessVisibility(node, component, dataset);
                        break;

                    case "KeepTogether":
                        string keepTogether = GetNodeTextValue(node);
                        component.CanBreak = !IsTrue(keepTogether);
                        break;

                    case "HideDuplicates":
                        //The value is the name of the scope (a group or a dataset) the duplicates are counted in.
                        //Only the text is hidden, the border and the background stay visible, so RemoveText is used.
                        //The scope itself is not mapped, in the banded model the duplicates are reset by the band the component belongs to.
                        if (!string.IsNullOrWhiteSpace(GetNodeTextValue(node)))
                            component.ProcessingDuplicates = StiProcessingDuplicatesType.RemoveText;
                        break;

                    case "dd:ShrinkToFit":
                        if (IsTrue(GetNodeTextValue(node)))
                            component.ShrinkFontToFit = true;
                        break;

                    //case "InteractiveHeight":
                    //case "InteractiveWidth":
                    case "ZIndex":
                    case "rd:DefaultName":
                    case "rd:WatermarkTextbox":
                    case "DataElementName":
                    case "DataElementOutput":
                        //ignored or not implemented yet
                        break;

                    //the icon of a toggle; a collapsing group of Stimulsoft draws its own marker
                    case "ToggleImage":
                        break;

                    default:
                        if (!ProcessInteractivityNode(node, component, dataset))
                            ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            component.Text = ConvertExpression(component.Text, dataset);
            component.Text = ApplyFormatExpression(component.Text, baseNode, dataset);

            if (generalAlignedTexts.Remove(component))
                ResolveGeneralAlignment(component);

            if (container.TagValue != null)
            {
                component.TagValue = container.TagValue;
            }

            container.Components.Add(component);

            lastTextComponent = null;
            currentTextboxName = null;
        }

        /// <summary>
        /// A text format cannot be an expression, so a format of Reporting Services given by an
        /// expression is applied to the value itself: {value} turns into {Strings.Format(value, format)}.
        /// The owner is the text box or the text run whose style holds the format. A text made of
        /// several parts is left unformatted, the loss is named in the list of errors.
        /// </summary>
        private string ApplyFormatExpression(string text, XmlNode owner, string dataset)
        {
            var format = owner["Style"]?["Format"];
            if (format == null || ConvertSyntaxToCSharp)
                return text;

            var formatValue = GetNodeTextValue(format);
            if (!IsExpression(formatValue))
                return text;

            //a constant text is not formatted by Reporting Services either
            if (!IsSingleExpression(text))
            {
                if (text != null && text.Contains("{"))
                    ThrowError(null, null, $"Format expression not applied: {formatValue}");

                return text;
            }

            var formatExpression = ConvertExpression(formatValue, dataset);

            return $"{{Strings.Format({text.Substring(1, text.Length - 2)}, " +
                $"{formatExpression.Substring(1, formatExpression.Length - 2)})}}";
        }

        /// <summary>
        /// Checks whether the whole text is one expression in braces. A brace inside a string of
        /// the expression does not count.
        /// </summary>
        private static bool IsSingleExpression(string text)
        {
            if (string.IsNullOrEmpty(text) || !text.StartsWith("{") || !text.EndsWith("}"))
                return false;

            var level = 0;
            var inString = false;

            for (var index = 0; index < text.Length; index++)
            {
                var ch = text[index];

                if (ch == '"')
                    inString = !inString;

                if (inString)
                    continue;

                if (ch == '{')
                    level++;

                if (ch == '}')
                {
                    level--;
                    if (level == 0 && index < text.Length - 1)
                        return false;
                }
            }

            return level == 0;
        }

        /// <summary>
        /// TextAlign General of Reporting Services puts a number or a date to the right and any other
        /// value to the left. The value is known only at run time, so the type is taken from the
        /// expression: a field, an aggregate, a variable or a number format of the text.
        /// </summary>
        private void ResolveGeneralAlignment(StiText text)
        {
            text.HorAlignment = IsNumericOrDateText(text) ? StiTextHorAlignment.Right : StiTextHorAlignment.Left;
        }

        private bool IsNumericOrDateText(StiText text)
        {
            if (text.TextFormat is StiNumberFormatService || text.TextFormat is StiDateFormatService)
                return true;

            var value = text.Text.Value;
            if (!IsSingleExpression(value))
                return false;

            var expression = value.Substring(1, value.Length - 2).Trim();

            if (Regex.IsMatch(expression, @"^Totals\.(Sum|Avg|Count|CountDistinct)\("))
                return true;

            if (Regex.IsMatch(expression, @"^-?[0-9]+(\.[0-9]+)?$"))
                return true;

            var aggregate = Regex.Match(expression, @"^Totals\.(First|Last|Max|Min)\((\w+, )?(?<column>\w+\.\w+)\)$");
            if (aggregate.Success)
                expression = aggregate.Groups["column"].Value;

            var type = GetExpressionType(expression);

            return type != null && (type.IsNumericType() || type == typeof(DateTime) || type == typeof(DateTime?));
        }

        /// <summary>
        /// Returns the type of a column written as DataSource.Column in the code of the report, or
        /// of a variable, or null. The names are compared as the code generator corrects them.
        /// </summary>
        private Type GetExpressionType(string expression)
        {
            var variable = FindVariable(expression);
            if (variable != null)
                return variable.Type;

            var dot = expression.IndexOf('.');
            if (dot == -1)
                return null;

            var sourceName = expression.Substring(0, dot);
            var columnName = expression.Substring(dot + 1);

            var dataSource = report.Dictionary.DataSources.Cast<StiDataSource>()
                .FirstOrDefault(d => StiNameValidator.CorrectName(d.Name, report) == sourceName);

            var column = dataSource?.Columns.Cast<StiDataColumn>()
                .FirstOrDefault(c => StiNameValidator.CorrectName(c.Name, report) == columnName);

            return column?.Type;
        }

        /// <summary>
        /// Checks whether the text runs of a text box differ in the font or in the color. One style
        /// of a component cannot hold them, so such a text is written with HTML tags.
        /// </summary>
        private static bool IsRichText(List<XmlNode> runs)
        {
            if (runs.Count < 2)
                return false;

            var first = GetRunStyle(runs[0]);

            return runs.Skip(1).Any(r => !GetRunStyle(r).SequenceEqual(first));
        }

        /// <summary>
        /// Returns the properties of the style of a text run that change the look of the text, sorted
        /// by the name of the property.
        /// </summary>
        private static SortedDictionary<string, string> GetRunStyle(XmlNode run)
        {
            var style = new SortedDictionary<string, string>();
            var styleNode = run.ChildNodes.Cast<XmlNode>().FirstOrDefault(n => n.Name == "Style");
            if (styleNode == null)
                return style;

            foreach (XmlNode node in styleNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "FontFamily":
                    case "FontSize":
                    case "FontWeight":
                    case "FontStyle":
                    case "TextDecoration":
                    case "Color":
                        style[node.Name] = node.InnerText.Trim();
                        break;
                }
            }

            return style;
        }

        /// <summary>
        /// Converts the paragraphs of a text box whose text runs differ in the font or the color.
        /// The style of the first run is the style of the component, without the bold, italic,
        /// underline and strikeout that not all the runs have; every run adds the rest with HTML tags.
        /// The format of a run is applied to the value of that run only, the component has no format.
        /// </summary>
        private string ProcessRichParagraphs(XmlNode baseNode, List<XmlNode> runs, StiText component, string dataset)
        {
            var firstStyle = runs[0].ChildNodes.Cast<XmlNode>().FirstOrDefault(n => n.Name == "Style");
            if (firstStyle != null)
                ProcessStyleType(firstStyle, component, dataset);

            var commonFlags = runs.Select(GetRunFontFlags).Aggregate((a, b) => a & b);
            component.Font = new Font(component.Font.FontFamily, component.Font.Size, commonFlags);
            component.TextFormat = new StiGeneralFormatService();
            component.AllowHtmlTags = true;

            var baseStyle = GetRunStyle(runs[0]);
            var text = new StringBuilder();

            foreach (XmlNode paragraph in baseNode.ChildNodes)
            {
                if (paragraph.Name != "Paragraph")
                {
                    ThrowError(baseNode.Name, paragraph.Name);
                    continue;
                }

                if (text.Length > 0)
                    text.Append("<br>");

                var paragraphStyle = paragraph.ChildNodes.Cast<XmlNode>().FirstOrDefault(n => n.Name == "Style");
                if (paragraphStyle != null)
                    ProcessStyleType(paragraphStyle, component, dataset);

                foreach (XmlNode run in paragraph.SelectNodes(".//*[local-name()='TextRun']"))
                {
                    text.Append(ConvertRichRun(run, baseStyle, commonFlags, dataset));
                }
            }

            return text.ToString();
        }

        /// <summary>
        /// Converts the value of a text run and wraps it in the HTML tags of its own font and color.
        /// A plain text is escaped, so a less-than sign is not taken for a tag.
        /// </summary>
        private string ConvertRichRun(XmlNode run, SortedDictionary<string, string> baseStyle, FontStyle commonFlags,
            string dataset)
        {
            var valueNode = run.ChildNodes.Cast<XmlNode>().FirstOrDefault(n => n.Name == "Value");
            var value = valueNode != null ? GetNodeTextValue(valueNode) : string.Empty;

            string text;
            if (IsExpression(value))
            {
                text = ApplyFormatExpression(ConvertExpression(value, dataset), run, dataset);
                text = ApplyRunFormat(text, run);
            }
            else
            {
                text = WebUtility.HtmlEncode(value ?? string.Empty);
            }

            var style = GetRunStyle(run);
            var flags = GetRunFontFlags(run) & ~commonFlags;
            var open = new StringBuilder();
            var close = new StringBuilder();

            var fontAttributes = GetFontTagAttributes(style, baseStyle);
            if (fontAttributes.Length > 0)
                AddTag(open, close, $"font{fontAttributes}", "font");

            if ((flags & FontStyle.Bold) != 0)
                AddTag(open, close, "b", "b");

            if ((flags & FontStyle.Italic) != 0)
                AddTag(open, close, "i", "i");

            if ((flags & FontStyle.Underline) != 0)
                AddTag(open, close, "u", "u");

            if ((flags & FontStyle.Strikeout) != 0)
                AddTag(open, close, "s", "s");

            return open + text + close;
        }

        /// <summary>
        /// A static format of a run is applied to its value with Format of Visual Basic, the same
        /// way a format given by an expression is, see ApplyFormatExpression.
        /// </summary>
        private string ApplyRunFormat(string text, XmlNode run)
        {
            var format = run["Style"]?["Format"];
            if (format == null || ConvertSyntaxToCSharp || !IsSingleExpression(text))
                return text;

            var formatValue = GetNodeTextValue(format);
            if (string.IsNullOrEmpty(formatValue) || IsExpression(formatValue))
                return text;

            var formatString = "\"" + formatValue.Replace("\"", "\"\"") + "\"";

            return $"{{Strings.Format({text.Substring(1, text.Length - 2)}, {formatString})}}";
        }

        private static string GetFontTagAttributes(SortedDictionary<string, string> style,
            SortedDictionary<string, string> baseStyle)
        {
            var attributes = new StringBuilder();

            if (IsChanged(style, baseStyle, "FontFamily", out var family))
                attributes.Append($" face=\"{family}\"");

            if (IsChanged(style, baseStyle, "FontSize", out var size))
            {
                var points = size.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
                    ? size.Substring(0, size.Length - 2)
                    : size;
                attributes.Append($" size=\"{points}\"");
            }

            if (IsChanged(style, baseStyle, "Color", out var color))
                attributes.Append($" color=\"{color}\"");

            return attributes.ToString();
        }

        /// <summary>
        /// Checks that the run sets a static value of the property other than the first run does.
        /// A value given by an expression cannot be put into a tag, the run keeps the base style.
        /// </summary>
        private static bool IsChanged(SortedDictionary<string, string> style, SortedDictionary<string, string> baseStyle,
            string name, out string value)
        {
            if (!style.TryGetValue(name, out value) || IsExpression(value))
                return false;

            return !baseStyle.TryGetValue(name, out var baseValue) || baseValue != value;
        }

        private static FontStyle GetRunFontFlags(XmlNode run)
        {
            var style = GetRunStyle(run);
            var flags = FontStyle.Regular;

            if (style.TryGetValue("FontWeight", out var weight) && IsBoldFontWeight(weight))
                flags |= FontStyle.Bold;

            if (style.TryGetValue("FontStyle", out var fontStyle) && fontStyle == "Italic")
                flags |= FontStyle.Italic;

            if (style.TryGetValue("TextDecoration", out var decoration))
            {
                if (decoration == "Underline")
                    flags |= FontStyle.Underline;

                else if (decoration == "LineThrough")
                    flags |= FontStyle.Strikeout;
            }

            return flags;
        }

        private static void AddTag(StringBuilder open, StringBuilder close, string openTag, string closeTag)
        {
            open.Append($"<{openTag}>");
            close.Insert(0, $"</{closeTag}>");
        }

        private string ProcessParagraphsType(XmlNode baseNode, string dataset)
        {
            var runs = baseNode.SelectNodes(".//*[local-name()='TextRun']").Cast<XmlNode>().ToList();
            ApplyTextRunActions(runs, dataset);

            if (lastTextComponent != null && IsRichText(runs))
                return ProcessRichParagraphs(baseNode, runs, lastTextComponent, dataset);

            StringBuilder text = new StringBuilder();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Paragraph":
                        if (text.Length > 0) text.Append("\r\n");
                        ProcessParagraphType(node, text, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return text.ToString();
        }

        private void ProcessParagraphType(XmlNode baseNode, StringBuilder text, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TextRuns":
                        ProcessTextRunsType(node, text, dataset);
                        break;

                    case "Style":
                        if (lastTextComponent != null)
                        {
                            ProcessStyleType(node, lastTextComponent, dataset);
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessTextRunsType(XmlNode baseNode, StringBuilder text, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TextRun":
                        ProcessTextRunType(node, text, dataset);
                        break;

                    //case "Style":
                    //    ProcessStyleType(node, component);
                    //    break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessTextRunType(XmlNode baseNode, StringBuilder text, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        string expr = GetNodeTextValue(node);
                        text.Append(ApplyFormatExpression(ConvertExpression(expr, dataset), baseNode, dataset));
                        break;

                    case "Style":
                        if (lastTextComponent != null)
                        {
                            ProcessStyleType(node, lastTextComponent, dataset);
                        }
                        break;

                    case "Label":
                        //ignore
                        break;

                    //applied to the whole text box, see ApplyTextRunActions
                    case "ActionInfo":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        /// <summary>
        /// A text run of Reporting Services may have its own action, a text of Stimulsoft has one for the whole
        /// text: the action of the first run that has one goes to the text box, the others are named in the list
        /// of errors. An action of the text box itself is read after its paragraphs and wins.
        /// </summary>
        private void ApplyTextRunActions(List<XmlNode> runs, string dataset)
        {
            if (lastTextComponent == null)
                return;

            var actions = runs
                .Select(r => r.SelectSingleNode("*[local-name()='ActionInfo']"))
                .Where(a => a != null)
                .ToList();
            if (actions.Count == 0)
                return;

            ProcessActionInfo(actions[0], lastTextComponent, dataset);

            if (actions.Count > 1)
                ThrowError(null, null, $"Only the action of the first text run is imported: {lastTextComponent.Name}");
        }
        #endregion

        #region Rectangle
        private void ProcessRectangleType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiPanel component = new StiPanel();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;
            rectanglePanels.Add(component);

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
                        ProcessStyleType(node, component, dataset);
                        break;

                    case "ReportItems":
                        ProcessReportItemsType(node, component, page, dataset);
                        break;

                    case "Visibility":
                        ProcessVisibility(node, component, dataset);
                        break;

                    case "KeepTogether":
                        string keepTogether = GetNodeTextValue(node);
                        component.CanBreak = !IsTrue(keepTogether);
                        if (IsTrue(keepTogether))
                            keptTogetherRectangles.Add(component);
                        break;

                    case "PageBreak":
                        ProcessItemPageBreak(node, component);
                        break;

                    //ignored or not implemented yet
                    //break;

                    default:
                        if (!ProcessInteractivityNode(node, component, dataset))
                            ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            container.Components.Add(component);
        }

        /// <summary>
        /// The items of Reporting Services below an item that shrinks move up, a component of Stimulsoft moves
        /// up only with the DecreasingSize shift mode. It is given to every component of a container that lies
        /// below a component that can shrink. The containers around a component that can shrink or can be hidden
        /// shrink too, up to the page, so the place it gives up is not kept by them, and so does a container of
        /// the rows docked to its top; they are marked first, the shift modes are given to the components below
        /// them then. A hidden component is not printed at all, so the components below it do not move up, only
        /// the container shrinks when nothing is below it.
        /// </summary>
        private void ApplyShiftModes()
        {
            foreach (StiPage page in report.Pages)
            {
                var components = page.GetComponents().Cast<StiComponent>().ToList();

                //the rows docked to the top of a container take the place of a hidden row, it shrinks then;
                //a band is docked to the top always, it is not a row
                foreach (var container in components.OfType<StiContainer>())
                {
                    if (container.Components.Cast<StiComponent>()
                        .Any(c => c.DockStyle == StiDockStyle.Top && !(c is StiBand)))
                        container.CanShrink = true;
                }

                //a hidden item gives up its place as a shrinking one does
                foreach (var component in components.Where(c => c.CanShrink || CanBeHidden(c)).ToList())
                {
                    SetParentsCanShrink(component);
                }

                ApplyShiftModes(page);

                foreach (var container in page.GetComponents().OfType<StiContainer>())
                {
                    ApplyShiftModes(container);
                }
            }
        }

        private void ApplyShiftModes(StiContainer container)
        {
            var shrinkComponents = container.Components.Cast<StiComponent>()
                .Where(c => c.CanShrink && !(c is StiBand))
                .ToList();

            if (shrinkComponents.Count == 0)
                return;

            foreach (StiComponent component in container.Components)
            {
                if (component is StiBand)
                    continue;

                var top = Math.Round(component.Top, 2);
                if (shrinkComponents.Any(c => c != component && Math.Round(c.Bottom, 2) <= top))
                    component.ShiftMode |= StiShiftMode.DecreasingSize;
            }
        }

        /// <summary>
        /// Returns true when the component is hidden always or by the expression of its visibility, which goes
        /// to the expression of Enabled or, when it calls the code of the report, to the BeforePrint event.
        /// </summary>
        private static bool CanBeHidden(StiComponent component)
        {
            if (!component.Enabled)
                return true;

            if (component.Expressions.Cast<StiAppExpression>().Any(e => e.Name == "Enabled"))
                return true;

            var script = component.BeforePrintEvent.Script;
            return !string.IsNullOrEmpty(script) && script.Contains(").Enabled = ");
        }

        private static void SetParentsCanShrink(StiComponent component)
        {
            var parent = component.Parent;

            while (parent != null && !(parent is StiPage))
            {
                parent.CanShrink = true;
                parent = parent.Parent;
            }
        }
        #endregion

        #region Line
        private void ProcessLineType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiHorizontalLinePrimitive component = new StiHorizontalLinePrimitive();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Page = page;
            StiText text = new StiText();

            //double left = 0;
            //double top = 0;
            double width = 0;
            double height = 0;

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
                        width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, text, dataset);
                        break;

                    case "Visibility":
                        ProcessVisibility(node, component, dataset);
                        break;

                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        if (!ProcessInteractivityNode(node, component, dataset))
                            ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (height < 0)
            {
                component.Top += height;
                height = Math.Abs(height);
            }
            if (width < 0)
            {
                component.Left += width;
                width = Math.Abs(width);
            }

            if (width > height)
            {
                component.Width = width;
                component.Height = height;
                component.Style = text.Border.Style;
                component.Color = text.Border.Color;
                component.Size = (float)text.Border.Size;
                container.Components.Add(component);
            }
            else
            {
                StiVerticalLinePrimitive line = new StiVerticalLinePrimitive();
                line.Name = component.Name;
                line.Page = component.Page;
                line.Left = component.Left;
                line.Top = component.Top;
                line.Width = width;
                line.Height = height;
                line.Style = text.Border.Style;
                line.Color = text.Border.Color;
                line.Size = (float)text.Border.Size;
                page.Components.Add(line);

                StiStartPointPrimitive start = new StiStartPointPrimitive();
                start.Left = line.Left;
                start.Top = line.Top;
                start.ReferenceToGuid = line.Guid;
                start.Page = page;
                container.Components.Add(start);

                StiEndPointPrimitive end = new StiEndPointPrimitive();
                end.Left = line.Right;
                end.Top = line.Bottom;
                end.ReferenceToGuid = line.Guid;
                end.Page = page;
                container.Components.Add(end);
            }
        }
        #endregion

        #region List
        private void ProcessListType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiPanel component = new StiPanel();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value + "Panel");
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;

            StiDataBand band = new StiDataBand();
            band.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            band.Page = page;
            band.CanGrow = true;

            //first pass - get dataset name
            string datasetName = null;
            string fullDatasetName = null;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSetName":
                        datasetName = Convert.ToString(GetNodeTextValue(node));
                        break;
                }
            }
            fullDatasetName = MakeFullDataSetName(dataset, datasetName);

            //second pass
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
                        ProcessStyleType(node, component, dataset);
                        break;

                    case "ReportItems":
                        ProcessReportItemsType(node, band, page, fullDatasetName);
                        break;

                    case "Visibility":
                        ProcessVisibility(node, component, dataset);
                        break;

                    case "DataSetName":
                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        if (!ProcessInteractivityNode(node, component, dataset))
                            ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            band.Height = component.Height;
            band.DataSourceName = datasetName;
            component.Components.Add(band);
            container.Components.Add(component);
        }
        #endregion

		#region Subreport
		private void ProcessSubreportType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiSubReport component = new StiSubReport();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;

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
                        ProcessStyleType(node, component, dataset);
                        break;

                    case "Parameters":
                        ProcessSubreportParametersType(node, component, dataset);
                        break;

                    case "ReportName":
                        component.SubReportUrl = GetNodeTextValue(node) + ".mrt";
                        break;

                    default:
                        if (!ProcessInteractivityNode(node, component, dataset))
                            ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            container.Components.Add(component);
        }

        private void ProcessSubreportParametersType(XmlNode baseNode, StiSubReport component, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Parameter":
                        ProcessSubreportParameterType(node, component, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessSubreportParameterType(XmlNode baseNode, StiSubReport component, string dataset)
        {
            var name = baseNode.Attributes["Name"]?.Value;
            if (string.IsNullOrWhiteSpace(name))
                return;

            var parameter = new StiParameter();
            parameter.Name = name;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        var expr = ConvertExpression(GetNodeTextValue(node), dataset).Trim();
                        if (expr.StartsWith("{") && expr.EndsWith("}"))
                            expr = expr.Substring(1, expr.Length - 2);
                        parameter.Expression.Value = expr;
                        break;

                    case "Omit":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            component.Parameters.Add(parameter);
        }
        #endregion
    }
}
