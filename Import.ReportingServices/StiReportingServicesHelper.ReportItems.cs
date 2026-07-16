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
using Stimulsoft.Report.Components;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
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
                        ProcessChartType(node, container, page, dataset);
                        break;
                    case "FixedPage":
                        ProcessFixedPageType(node, container, page);
                        break;
                    case "CustomReportItem":
                        ProcessCustomReportItemType(node, container, page, dataset);
                        break;

                    case "RowSpan":     //todo
                        //ignored or not implemented yet
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
            component.Font = new Font("Arial", 10);
            component.Page = page;
            component.WordWrap = true;

            lastTextComponent = component;

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
                        ProcessStyleType(node, component, dataset);
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

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            component.Text = ConvertExpression(component.Text, dataset);
            if (container.TagValue != null)
            {
                component.TagValue = container.TagValue;
            }

            container.Components.Add(component);

            lastTextComponent = null;
        }

        private string ProcessParagraphsType(XmlNode baseNode, string dataset)
        {
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
                        text.Append(ConvertExpression(expr, dataset));
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

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
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
                        break;

                    //ignored or not implemented yet
                    //break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            container.Components.Add(component);
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
