#region Copyright (C) 2003-2025 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports  											}
{																	}
{	Copyright (C) 2003-2025 Stimulsoft     							}
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
#endregion Copyright (C) 2003-2025 Stimulsoft

using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Net.Mime;
using System.Text;
using System.Threading;
using System.Xml;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components.TextFormats;
using Stimulsoft.Base;

namespace Stimulsoft.Report.Import
{
    public class StiActiveReportsHelper
    {
        #region Fields
        List<string> errorList = null;
        ArrayList fields = new ArrayList();
        Hashtable fieldsNames = new Hashtable();
        #endregion

        #region Root node
        public void ProcessRootNode(XmlNode rootNode, StiReport report, List<string> errorList)
        {
            this.errorList = errorList;
            this.fields.Clear();
            this.fieldsNames = new Hashtable();
            report.ReportUnit = StiReportUnitType.HundredthsOfInch;
            StiPage page = report.Pages[0];

            if (rootNode.Attributes["DocumentName"] != null)
            {
                report.ReportName = rootNode.Attributes["DocumentName"].Value;
            }
            if (rootNode.Attributes["Version"] != null)
            {
                report.ReportDescription = "ActiveReports Version: " + rootNode.Attributes["Version"].Value;
            }

            //ScriptLang="C#"
            //MasterReport="0"

            bool pageSettingsUsed = false;
            bool dataSourceUsed = false;

            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {

                    case "Sections":
                        ProcessSections(node, report, page);
                        break;

                    case "DataSources":
                        dataSourceUsed = ProcessDataSources(node, report);
                        break;

                    case "ReportComponentTray":
                        ProcessReportComponentTray(node, report);
                        break;
                    case "PageSettings":
                        pageSettingsUsed = ProcessPageSettings(node, report, page);
                        break;
                    case "Parameters":
                        ProcessParameters(node, report);
                        break;

                    case "StyleSheet":
                        ProcessStyleSheet(node, report);
                        break;

                    //case "EmbeddedImages":
                    //    ProcessEmbeddedImagesType(node);
                    //    break;

                    //case "wwwwWidth":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(rootNode.Name, node.Name);
                        break;
                }
            }

            #region Check PageSetting
            double printWidth = ToHi(rootNode.Attributes["PrintWidth"].Value);
            if (!pageSettingsUsed)
            {
                double marginLR = (page.PageWidth - printWidth) / 2;
                if (marginLR < 0) marginLR = 0;
                page.Margins.Left = marginLR;
                page.Margins.Right = marginLR;
            }
            if (page.Margins.Left + page.Margins.Right + printWidth > page.PageWidth * 1.1)
            {
                page.PaperSize = System.Drawing.Printing.PaperKind.Custom;
                page.PageWidth = page.Margins.Left + page.Margins.Right + printWidth;
            }
            #endregion

            #region Check DataSource
            if (!dataSourceUsed)
            {
                StiDataTableSource dataSource = new StiDataTableSource();
                dataSource.Name = "ds";
                dataSource.Alias = "ds";
                dataSource.NameInSource = "";

                foreach (string st in fields)
                {
                    if (st.Length > 0)
                    {
                        dataSource.Columns.Add(new StiDataColumn(st));
                    }
                }

                report.Dictionary.DataSources.Add(dataSource);
            }
            #endregion

        }
        #endregion

        #region Sections
        private void ProcessSections(XmlNode baseNode, StiReport report, StiPage page)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Section":
                        ProcessSection(node, report, page);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessSection(XmlNode baseNode, StiReport report, StiPage page)
        {
            string sectionType = baseNode.Attributes["Type"].Value;
            StiBand band = new StiDataBand();
            if (sectionType == "ReportHeader") band = new StiReportTitleBand();
            if (sectionType == "PageHeader") band = new StiPageHeaderBand();
            if (sectionType == "GroupHeader") band = new StiGroupHeaderBand();
            if (sectionType == "Detail") band = new StiDataBand();
            if (sectionType == "GroupFooter") band = new StiGroupFooterBand();
            if (sectionType == "PageFooter") band = new StiPageFooterBand();
            if (sectionType == "ReportFooter") band = new StiReportSummaryBand();

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        band.Name = node.Value;
                        break;
                    case "Height":
                        band.Height = ToHi(node.Value);
                        break;
                    case "BackColor":
                        band.Brush = new StiSolidBrush(ParseColor(node));
                        break;

                    case "CanGrow":
                        if (node.Value == "0") band.CanGrow = false;
                        break;
                    case "CanShrink":
                        if (node.Value == "1") band.CanShrink = true;
                        break;

                    case "Type":
                        //ignored or not implemented yet
                        break;

                    case "DataField":
                        if (band is StiGroupHeaderBand) (band as StiGroupHeaderBand).Condition = new StiGroupConditionExpression("{" + ConvertExpression(node.Value, report) + "}");
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Control":
                        ProcessControl(node, report, page, band);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            page.Components.Add(band);
        }
        #endregion

        #region Controls
        private void ProcessControl(XmlNode baseNode, StiReport report, StiPage page, StiBand band)
        {
            string sectionType = baseNode.Attributes["Type"].Value;
            StiComponent component = null;

            switch (sectionType)
            {
                case "AR.Field":
                    component = ProcessARField(baseNode, report, page, band);
                    break;
                case "AR.Label":
                    component = ProcessARLabel(baseNode, report, page, band);
                    break;
                case "AR.Line":
                    component = ProcessARLine(baseNode, report, page, band);
                    break;
                case "AR.Subreport":
                    component = ProcessARSubreport(baseNode, report, page, band);
                    break;
                case "AR.CheckBox":
                    component = ProcessARCheckBox(baseNode, report, page, band);
                    break;
                case "AR.RTF":
                    component = ProcessARRTF(baseNode, report, page, band);
                    break;
                case "AR.Image":
                    component = ProcessARImage(baseNode, report, page, band);
                    break;

                //case "Type":
                //    //ignored or not implemented yet
                //    break;

                default:
                    ThrowError(baseNode.Name, sectionType);
                    break;
            }

            if (component != null)
            {
                //comp.Name = CheckComponentsNames(obj.Name);

                if ((component is StiVerticalLinePrimitive) || (component is StiRectanglePrimitive))
                {
                    StiStartPointPrimitive start = new StiStartPointPrimitive();
                    start.Name = StiNameCreation.CreateName(report, "StartPoint");
                    start.Left = component.Left;
                    start.Top = component.Top;
                    start.ReferenceToGuid = component.Guid;
                    start.Parent = band;
                    start.Page = page;
                    band.Components.Add(start);

                    StiEndPointPrimitive end = new StiEndPointPrimitive();
                    end.Name = StiNameCreation.CreateName(report, "EndPoint");
                    end.Left = component.Right;
                    end.Top = component.Bottom;
                    end.ReferenceToGuid = component.Guid;
                    end.Parent = band;
                    end.Page = page;
                    band.Components.Add(end);

                    component.Top += band.Top;
                    page.Components.Add(component);
                }
                else
                {
                    band.Components.Add(component);
                }

                component.Page = page;
                component.Report = report;
            }
        }

        private bool ProcessCommonControlProperties(XmlNode node, StiComponent component)
        {
            IStiBorder border = component as IStiBorder;

            switch (node.Name)
            {
                case "Name":
                    component.Name = node.Value;
                    return true;
                case "Left":
                    component.Left = ToHi(node.Value);
                    return true;
                case "Top":
                    component.Top = ToHi(node.Value);
                    return true;
                case "Width":
                    component.Width = ToHi(node.Value);
                    return true;
                case "Height":
                    component.Height = ToHi(node.Value);
                    return true;
                case "Visible":
                    if (node.Value == "0") component.Enabled = false;
                    return true;


                case "BorderLeftStyle":
                    if (border != null)
                    {
                        if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
                        (border.Border as StiAdvancedBorder).LeftSide.Style = ParseLineStyle(node.Value);
                    }
                    return true;
                case "BorderLeftColor":
                    if (border != null)
                    {
                        if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
                        (border.Border as StiAdvancedBorder).LeftSide.Color = ParseColor(node);
                    }
                    return true;

                case "BorderRightStyle":
                    if (border != null)
                    {
                        if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
                        (border.Border as StiAdvancedBorder).RightSide.Style = ParseLineStyle(node.Value);
                    }
                    return true;
                case "BorderRightColor":
                    if (border != null)
                    {
                        if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
                        (border.Border as StiAdvancedBorder).RightSide.Color = ParseColor(node);
                    }
                    return true;

                case "BorderTopStyle":
                    if (border != null)
                    {
                        if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
                        (border.Border as StiAdvancedBorder).TopSide.Style = ParseLineStyle(node.Value);
                    }
                    return true;
                case "BorderTopColor":
                    if (border != null)
                    {
                        if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
                        (border.Border as StiAdvancedBorder).TopSide.Color = ParseColor(node);
                    }
                    return true;

                case "BorderBottomStyle":
                    if (border != null)
                    {
                        if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
                        (border.Border as StiAdvancedBorder).BottomSide.Style = ParseLineStyle(node.Value);
                    }
                    return true;
                case "BorderBottomColor":
                    if (border != null)
                    {
                        if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
                        (border.Border as StiAdvancedBorder).BottomSide.Color = ParseColor(node);
                    }
                    return true;

                case "ClassName":
                    component.ComponentStyle = node.Value;
                    return true;

                case "Type":
                    //ignored
                    return true;

            }
            return false;
        }

        #region AR.Field
        private StiComponent ProcessARField(XmlNode baseNode, StiReport report, StiPage page, StiBand band)
        {
            StiText component = new StiText();
            component.CanGrow = true;
            component.CanShrink = false;
            component.WordWrap = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "DataField":
                        component.Text = "{" + ConvertExpression(node.Value, report) + "}";
                        break;
                    case "Text":
                        if (baseNode.Attributes["DataField"] == null)
                        {
                            component.Text = node.Value;
                        }
                        break;

                    case "CanGrow":
                        if (node.Value == "0") component.CanGrow = false;
                        break;
                    case "CanShrink":
                        if (node.Value == "1") component.CanShrink = true;
                        break;

                    case "Style":
                        StiStyle style = ParseStringToStyle(node.Value, component.TextOptions);
                        style.SetStyleToComponent(component);
                        break;

                    case "OutputFormat":
                        component.TextFormat = new StiCustomFormatService(node.Value);
                        break;

                    //case "Type":
                    //ignored or not implemented yet
                    //break;

                    default:
                        if (!ProcessCommonControlProperties(node, component)) ThrowError(baseNode.Attributes["Type"].Value, "#" + node.Name);
                        break;
                }
            }

            return component;
        }
        #endregion

        #region AR.Label
        private StiComponent ProcessARLabel(XmlNode baseNode, StiReport report, StiPage page, StiBand band)
        {
            StiText component = new StiText();
            component.CanGrow = true;
            component.CanShrink = false;
            component.WordWrap = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Caption":
                        component.Text = node.Value;
                        break;

                    case "CanGrow":
                        if (node.Value == "0") component.CanGrow = false;
                        break;
                    case "CanShrink":
                        if (node.Value == "1") component.CanShrink = true;
                        break;

                    case "Style":
                        StiStyle style = ParseStringToStyle(node.Value, component.TextOptions);
                        style.SetStyleToComponent(component);
                        break;

                    case "HyperLink":
                        if (node.Value.Length > 0) component.Hyperlink.Value = node.Value;
                        break;

                    //case "Type":
                    //ignored or not implemented yet
                    //break;

                    default:
                        if (!ProcessCommonControlProperties(node, component)) ThrowError(baseNode.Attributes["Type"].Value, "#" + node.Name);
                        break;
                }
            }

            return component;
        }
        #endregion

        #region AR.Line
        private StiComponent ProcessARLine(XmlNode baseNode, StiReport report, StiPage page, StiBand band)
        {
            double x1 = 0;
            double x2 = 0;
            double y1 = 0;
            double y2 = 0;
            string name = null;
            bool enabled = true;
            Color lineColor = Color.Black;
            StiPenStyle lineStyle = StiPenStyle.Solid;
            float lineWeight = 1;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        break;
                    case "X1":
                        x1 = ToHi(node.Value);
                        break;
                    case "X2":
                        x2 = ToHi(node.Value);
                        break;
                    case "Y1":
                        y1 = ToHi(node.Value);
                        break;
                    case "Y2":
                        y2 = ToHi(node.Value);
                        break;

                    case "LineColor":
                        lineColor = ParseColor(node);
                        break;
                    case "LineStyle":
                        lineStyle = ParseLineStyle(node.Value);
                        break;
                    case "LineWeight":
                        lineWeight = (float)ParseDouble(node.Value);
                        break;

                    case "Visible":
                        if (node.Value == "0") enabled = false;
                        break;

                    case "Type":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Attributes["Type"].Value, "#" + node.Name);
                        break;
                }
            }

            StiShape shape = new StiShape();
            shape.Left = Math.Min(x1, x2);
            shape.Width = Math.Abs(x2 - x1);
            shape.Top = Math.Min(y1, y2);
            shape.Height = Math.Abs(y2 - y1);
            if ((x1 < x2) && (y1 < y2) || (x1 > x2) && (y1 > y2))
            {
                shape.ShapeType = new Stimulsoft.Report.Components.ShapeTypes.StiDiagonalDownLineShapeType();
            }
            else
            {
                shape.ShapeType = new Stimulsoft.Report.Components.ShapeTypes.StiDiagonalUpLineShapeType();
            }

            if (shape.Width == 0 || shape.Height == 0)
            {
                StiLinePrimitive line = null;
                if (shape.Width == 0)
                {
                    line = new StiVerticalLinePrimitive();
                    line.Height = shape.Height;
                }
                else
                {
                    line = new StiHorizontalLinePrimitive();
                    line.Width = shape.Width;
                }
                line.Size = lineWeight;
                line.Style = lineStyle;
                line.Color = lineColor;
                line.Left = shape.Left;
                line.Top = shape.Top;
                line.Enabled = enabled;
                line.Name = name;

                return line;
            }
            else
            {
                shape.Style = lineStyle;
                shape.BorderColor = lineColor;
                shape.Size = lineWeight;
                shape.Enabled = enabled;
                shape.Name = name;

                return shape;
            }
        }
        #endregion

        #region AR.Subreport
        private StiComponent ProcessARSubreport(XmlNode baseNode, StiReport report, StiPage page, StiBand band)
        {
            StiSubReport component = new StiSubReport();
            component.CanGrow = true;
            component.CanShrink = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "CanGrow":
                        if (node.Value == "0") component.CanGrow = false;
                        break;
                    case "CanShrink":
                        if (node.Value == "0") component.CanShrink = false;
                        break;

                    case "ReportName":
                        component.SubReportPage = report.Pages[node.Value];
                        break;

                    //case "Type":
                    //ignored or not implemented yet
                    //break;

                    default:
                        if (!ProcessCommonControlProperties(node, component)) ThrowError(baseNode.Attributes["Type"].Value, "#" + node.Name);
                        break;
                }
            }

            return component;
        }
        #endregion

        #region AR.CheckBox
        private StiComponent ProcessARCheckBox(XmlNode baseNode, StiReport report, StiPage page, StiBand band)
        {
            StiCheckBox component = new StiCheckBox();
            component.Checked = new StiCheckedExpression("{true}");

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "DataField":
                        component.Checked = new StiCheckedExpression("{" + ConvertExpression(node.Value, report) + "}");
                        break;

                    case "Style":
                        StiStyle style = ParseStringToStyle(node.Value, new StiTextOptions());
                        style.SetStyleToComponent(component);
                        break;

                    case "Value":
                        if (baseNode.Attributes["DataField"] == null)
                        {
                            string stValue = node.Value == "1" ? "{true}" : "{false}";
                            component.Checked = new StiCheckedExpression(stValue);
                        }
                        break;

                    case "Caption":
                        //not implemented yet
                        break;

                    //case "Type":
                    //  //ignored or not implemented yet
                    //  //break;

                    default:
                        if (!ProcessCommonControlProperties(node, component)) ThrowError(baseNode.Attributes["Type"].Value, "#" + node.Name);
                        break;
                }
            }

            return component;
        }
        #endregion

        #region AR.RTF
        private StiComponent ProcessARRTF(XmlNode baseNode, StiReport report, StiPage page, StiBand band)
        {
            StiRichText component = new StiRichText();
            component.CanGrow = true;
            component.CanShrink = false;
            component.WordWrap = true;
            component.RtfText = @"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Microsoft Sans Serif;}}\viewkind4\uc1\pard\f0\fs17\par}";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "DataField":
                        component.RtfText = @"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Microsoft Sans Serif;}{\f1\fnil\fcharset0 Microsoft Sans Serif;}}\viewkind4\uc1\pard\f0\fs17\lang1033\{" + ConvertExpression(node.Value, report) + @"\}\f1\par}";
                        break;

                    case "CanGrow":
                        if (node.Value == "0") component.CanGrow = false;
                        break;
                    case "CanShrink":
                        if (node.Value == "1") component.CanShrink = true;
                        break;

                    case "BackColor":
                        component.BackColor = ParseColor(node);
                        break;

                    //case "Type":
                    //ignored or not implemented yet
                    //break;

                    default:
                        if (!ProcessCommonControlProperties(node, component)) ThrowError(baseNode.Attributes["Type"].Value, "#" + node.Name);
                        break;
                }
            }

            return component;
        }
        #endregion

        #region AR.Image
        private StiComponent ProcessARImage(XmlNode baseNode, StiReport report, StiPage page, StiBand band)
        {
            StiImage component = new StiImage();
            component.CanGrow = false;
            component.CanShrink = false;
            component.Border = new StiBorder(StiBorderSides.None, Color.Transparent, 0, StiPenStyle.None);

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "DataField":
                        component.DataColumn = "{" + ConvertExpression(node.Value, report) + "}";
                        break;

                    case "BackColor":
                        component.Brush = new StiSolidBrush(ParseColor(node));
                        break;

                    case "SizeMode":
                        if (node.Value == "1")
                        {
                            component.Stretch = true;
                            component.AspectRatio = false;
                        }
                        else if (node.Value == "2")
                        {
                            component.Stretch = true;
                            component.AspectRatio = true;
                        }
                        else
                        {
                            component.Stretch = false;
                            component.AspectRatio = true;
                        }
                        break;

                    case "PictureAlignment":
                        if (node.Value == "0")
                        {
                            component.HorAlignment = StiHorAlignment.Left;
                            component.VertAlignment = StiVertAlignment.Top;
                        }
                        else if (node.Value == "1")
                        {
                            component.HorAlignment = StiHorAlignment.Right;
                            component.VertAlignment = StiVertAlignment.Top;
                        }
                        else if (node.Value == "3")
                        {
                            component.HorAlignment = StiHorAlignment.Left;
                            component.VertAlignment = StiVertAlignment.Bottom;
                        }
                        else if (node.Value == "4")
                        {
                            component.HorAlignment = StiHorAlignment.Right;
                            component.VertAlignment = StiVertAlignment.Bottom;
                        }
                        else
                        {
                            component.HorAlignment = StiHorAlignment.Center;
                            component.VertAlignment = StiVertAlignment.Center;
                        }
                        break;

                    case "LineColor":
                        component.Border.Color = ParseColor(node);
                        component.Border.Side = StiBorderSides.All;
                        break;
                    case "LineStyle":
                        component.Border.Style = ParseLineStyle(node.Value);
                        component.Border.Side = StiBorderSides.All;
                        break;
                    case "LineWeight":
                        component.Border.Size = (float)ParseDouble(node.Value);
                        component.Border.Side = StiBorderSides.All;
                        break;

                    //case "Type":
                    //  //ignored
                    //  //break;

                    default:
                        if (!ProcessCommonControlProperties(node, component)) ThrowError(baseNode.Attributes["Type"].Value, "#" + node.Name);
                        break;
                }
            }

            if (baseNode.HasChildNodes && baseNode.FirstChild.NodeType == XmlNodeType.Text)
            {
                string st = Convert.ToString(baseNode.FirstChild.Value);

                #region read image data
                int newPos2 = 9 * 2;
                MemoryStream ms = new MemoryStream();
                while (newPos2 < st.Length - 1)
                {
                    ms.WriteByte((byte)ParseHexTwoCharToInt(st[newPos2], st[newPos2 + 1]));
                    newPos2 += 2;
                }
                ms.Seek(0, SeekOrigin.Begin);
                #endregion

                component.ImageBytes = ms.ToArray();
            }

            return component;
        }
        #endregion

        #endregion

        #region DataSources
        private bool ProcessDataSources(XmlNode baseNode, StiReport report)
        {
            bool dataSourceUsed = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "OleDbDataSource":
                        ProcessOleDbDataSource(node, report);
                        dataSourceUsed = true;
                        break;
                    case "SqlDbDataSource":
                        ProcessSqlDbDataSource(node, report);
                        dataSourceUsed = true;
                        break;
                    case "XmlDataSource":
                        ProcessXmlDataSource(node, report);
                        dataSourceUsed = true;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return dataSourceUsed;
        }

        private void ProcessOleDbDataSource(XmlNode baseNode, StiReport report)
        {
            StiOleDbDatabase dataBase = new StiOleDbDatabase();
            dataBase.Name = "OleDbConnection";
            dataBase.Alias = "OleDbConnection";
            StiOleDbSource dataSource = new StiOleDbSource();
            dataSource.Name = "ds";
            dataSource.Alias = "ds";
            dataSource.NameInSource = "OleDbConnection";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Connect":
                        dataBase.ConnectionString = node.Value;
                        break;
                    case "ConnectF":
                        dataBase.ConnectionString = "";
                        break;
                    case "SQL":
                        dataSource.SqlCommand = node.Value;
                        break;

                    //case "Type":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (string st in fields)
            {
                if (st.Length > 0)
                {
                    dataSource.Columns.Add(new StiDataColumn(st));
                }
            }

            report.Dictionary.Databases.Add(dataBase);
            report.Dictionary.DataSources.Add(dataSource);
        }

        private void ProcessSqlDbDataSource(XmlNode baseNode, StiReport report)
        {
            StiSqlDatabase dataBase = new StiSqlDatabase();
            dataBase.Name = "SqlDbConnection";
            dataBase.Alias = "SqlDbConnection";
            StiSqlSource dataSource = new StiSqlSource();
            dataSource.Name = "ds";
            dataSource.Alias = "ds";
            dataSource.NameInSource = "SqlDbConnection";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Connect":
                        dataBase.ConnectionString = node.Value;
                        break;
                    case "ConnectF":
                        dataBase.ConnectionString = "";
                        break;
                    case "SQL":
                        dataSource.SqlCommand = node.Value;
                        break;

                    //case "Type":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (string st in fields)
            {
                if (st.Length > 0)
                {
                    dataSource.Columns.Add(new StiDataColumn(st));
                }
            }

            report.Dictionary.Databases.Add(dataBase);
            report.Dictionary.DataSources.Add(dataSource);
        }

        private void ProcessXmlDataSource(XmlNode baseNode, StiReport report)
        {
            StiXmlDatabase dataBase = new StiXmlDatabase();
            dataBase.Name = "XmlConnection";
            dataBase.Alias = "XmlConnection";
            StiDataTableSource dataSource = new StiDataTableSource();
            dataSource.Name = "ds";
            dataSource.Alias = "ds";
            dataSource.NameInSource = "XmlConnection";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "FileURL":
                        dataBase.PathData = node.Value;
                        break;

                    case "Pattern":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (string st in fields)
            {
                if (st.Length > 0)
                {
                    dataSource.Columns.Add(new StiDataColumn(st));
                }
            }

            report.Dictionary.Databases.Add(dataBase);
            report.Dictionary.DataSources.Add(dataSource);
        }
        #endregion

        #region ReportComponentTray
        private void ProcessReportComponentTray(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {

                    //case "":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region PageSettings
        private bool ProcessPageSettings(XmlNode baseNode, StiReport report, StiPage page)
        {
            if (baseNode.Attributes["Orientation"] != null)
            {
                if (baseNode.Attributes["Orientation"].Value == "2") page.Orientation = StiPageOrientation.Landscape;
            }

            bool pageSettingsUsed = false;
            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {

                    case "LeftMargin":
                        page.Margins.Left = ToHi(node.Value);
                        pageSettingsUsed = true;
                        break;
                    case "RightMargin":
                        page.Margins.Right = ToHi(node.Value);
                        pageSettingsUsed = true;
                        break;
                    case "TopMargin":
                        page.Margins.Top = ToHi(node.Value);
                        pageSettingsUsed = true;
                        break;
                    case "BottomMargin":
                        page.Margins.Bottom = ToHi(node.Value);
                        pageSettingsUsed = true;
                        break;


                    case "Orientation":
                        //ignored
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return pageSettingsUsed;
        }
        #endregion

        #region StyleSheet
        private void ProcessStyleSheet(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Style":
                        ProcessStyle(node, report);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessStyle(XmlNode baseNode, StiReport report)
        {
            StiStyle style = null;
            string styleName = null;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        styleName = node.Value;
                        break;

                    case "Value":
                        style = ParseStringToStyle(node.Value, new StiTextOptions());
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }
            if (style == null) style = ParseStringToStyle(string.Empty, new StiTextOptions());
            if (styleName != null)
            {
                style.Name = styleName;
                report.Styles.Add(style);
            }
        }
        #endregion

        #region Parameters
        private void ProcessParameters(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {

                    case "Parameter":
                        ProcessParameter(node, report);
                        break;

                    //case "":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessParameter(XmlNode baseNode, StiReport report)
        {
            string name = null;
            string defaultValue = null;
            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {

                    case "Key":
                        name = node.Value;
                        break;
                    case "DefaultValue":
                        defaultValue = node.Value;
                        break;


                    //case "":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (name != null)
            {
                if (!report.Dictionary.Variables.Contains(name))
                {
                    StiVariable var = new StiVariable("Parameters", name, typeof(string));
                    report.Dictionary.Variables.Add(var);
                }
                report.Dictionary.Variables[name].Value = defaultValue;
            }
        }
        #endregion

        #region Style

        private static StiTextHorAlignment ParseTextAlign(string alignAttribute)
        {
            switch (alignAttribute)
            {
                case "center":
                    return StiTextHorAlignment.Center;

                case "right":
                    return StiTextHorAlignment.Right;

                case "justify":
                    return StiTextHorAlignment.Width;
            }
            return StiTextHorAlignment.Left;
        }

        private static StiVertAlignment ParseVerticalAlign(string alignAttribute)
        {
            switch (alignAttribute)
            {
                case "middle":
                    return StiVertAlignment.Center;

                case "bottom":
                    return StiVertAlignment.Bottom;
            }
            return StiVertAlignment.Top;
        }

        private StiStyle ParseStringToStyle(string styleString, StiTextOptions textOptions)
        {
            StiStyle style = new StiStyle();
            style.Font = new Font("Arial", 10);
            style.AllowUseBorderFormatting = false;
            style.AllowUseBorderSides = false;
            style.AllowUseBrush = true;
            style.AllowUseFont = true;
            style.AllowUseHorAlignment = true;
            style.AllowUseImage = false;
            style.AllowUseTextBrush = true;
            style.AllowUseVertAlignment = true;

            string[] parts = styleString.Split(new char[] { ';' });
            foreach (string part in parts)
            {
                int indexPair = part.IndexOf(":");
                if (indexPair != -1)
                {
                    string attributeName = part.Substring(0, indexPair).Trim();
                    string attributeValue = part.Substring(indexPair + 1).Trim();
                    switch (attributeName)
                    {
                        case "background-color":
                            style.Brush = new StiSolidBrush(ParseStyleColor(attributeValue));
                            break;
                        case "color":
                            style.TextBrush = new StiSolidBrush(ParseStyleColor(attributeValue));
                            break;

                        case "font-family":
                            if (attributeValue.StartsWith("'") && attributeValue.EndsWith("'")) attributeValue = attributeValue.Substring(1, attributeValue.Length - 2);
                            style.Font = StiFontUtils.ChangeFontName(style.Font, attributeValue);
                            break;
                        case "font-size":
                            style.Font = StiFontUtils.ChangeFontSize(style.Font, (float)ParseDouble(attributeValue.Substring(0, attributeValue.Length - 2)));
                            break;
                        case "font-style":
                            if (attributeValue.Contains("italic")) style.Font = StiFontUtils.ChangeFontStyleItalic(style.Font, true);
                            break;
                        case "font-weight":
                            if (attributeValue.Contains("bold")) style.Font = StiFontUtils.ChangeFontStyleBold(style.Font, true);
                            break;
                        case "text-decoration":
                            if (attributeValue.Contains("underline")) style.Font = StiFontUtils.ChangeFontStyleUnderline(style.Font, true);
                            if (attributeValue.Contains("line-through")) style.Font = StiFontUtils.ChangeFontStyleStrikeout(style.Font, true);
                            break;

                        case "text-align":
                            style.HorAlignment = ParseTextAlign(attributeValue);
                            break;
                        case "vertical-align":
                            style.VertAlignment = ParseVerticalAlign(attributeValue);
                            break;

                        case "white-space":
                            if (attributeValue.Contains("nowrap"))
                            {
                                textOptions.WordWrap = false;
                            }
                            break;

                        case "ddo-font-vertical":
                        case "ddo-char-set":
                            //not implemented yet
                            break;

                        //case "":
                        //    //ignored
                        //    break;

                        default:
                            ThrowError("Style", attributeName);
                            break;
                    }
                }
            }

            return style;
        }

        private static Hashtable HtmlNameToColor = null;

        private static Color ParseStyleColor(string colorAttribute)
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
                else if (colorAttribute.StartsWith("rgb"))
                {
                    #region Parse RGB function
                    string[] colors = colorAttribute.Substring(4, colorAttribute.Length - 5).Split(new char[] { ',' });
                    color = Color.FromArgb(0xFF, int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]));
                    #endregion
                }
                else
                {
                    #region Parse color keywords
                    if (HtmlNameToColor == null)
                    {
                        #region Init hashtable
                        string[,] initData = {
                        {"AliceBlue",	    "#F0F8FF"},
                        {"AntiqueWhite",	"#FAEBD7"},
                        {"Aqua",	    "#00FFFF"},
                        {"Aquamarine",	"#7FFFD4"},
                        {"Azure",	    "#F0FFFF"},
                        {"Beige",	    "#F5F5DC"},
                        {"Bisque",	    "#FFE4C4"},
                        {"Black",	    "#000000"},
                        {"BlanchedAlmond",	"#FFEBCD"},
                        {"Blue",	    "#0000FF"},
                        {"BlueViolet",	"#8A2BE2"},
                        {"Brown",	    "#A52A2A"},
                        {"BurlyWood",	"#DEB887"},
                        {"CadetBlue",	"#5F9EA0"},
                        {"Chartreuse",	"#7FFF00"},
                        {"Chocolate",	"#D2691E"},
                        {"Coral",	    "#FF7F50"},
                        {"CornflowerBlue",	"#6495ED"},
                        {"Cornsilk",	"#FFF8DC"},
                        {"Crimson",	    "#DC143C"},
                        {"Cyan",	    "#00FFFF"},
                        {"DarkBlue",	"#00008B"},
                        {"DarkCyan",	"#008B8B"},
                        {"DarkGoldenRod",	"#B8860B"},
                        {"DarkGray",	"#A9A9A9"},
                        {"DarkGrey",	"#A9A9A9"},
                        {"DarkGreen",	"#006400"},
                        {"DarkKhaki",	"#BDB76B"},
                        {"DarkMagenta",	"#8B008B"},
                        {"DarkOliveGreen",	"#556B2F"},
                        {"Darkorange",	"#FF8C00"},
                        {"DarkOrchid",	"#9932CC"},
                        {"DarkRed",	    "#8B0000"},
                        {"DarkSalmon",	"#E9967A"},
                        {"DarkSeaGreen",	"#8FBC8F"},
                        {"DarkSlateBlue",	"#483D8B"},
                        {"DarkSlateGray",	"#2F4F4F"},
                        {"DarkSlateGrey",	"#2F4F4F"},
                        {"DarkTurquoise",	"#00CED1"},
                        {"DarkViolet",	"#9400D3"},
                        {"DeepPink",	"#FF1493"},
                        {"DeepSkyBlue",	"#00BFFF"},
                        {"DimGray",	    "#696969"},
                        {"DimGrey",	    "#696969"},
                        {"DodgerBlue",	"#1E90FF"},
                        {"FireBrick",	"#B22222"},
                        {"FloralWhite",	"#FFFAF0"},
                        {"ForestGreen",	"#228B22"},
                        {"Fuchsia",	    "#FF00FF"},
                        {"Gainsboro",	"#DCDCDC"},
                        {"GhostWhite",	"#F8F8FF"},
                        {"Gold",	    "#FFD700"},
                        {"GoldenRod",	"#DAA520"},
                        {"Gray",	    "#808080"},
                        {"Grey",	    "#808080"},
                        {"Green",	    "#008000"},
                        {"GreenYellow",	"#ADFF2F"},
                        {"HoneyDew",	"#F0FFF0"},
                        {"HotPink",	    "#FF69B4"},
                        {"IndianRed",	"#CD5C5C"},
                        {"Indigo",	    "#4B0082"},
                        {"Ivory",	    "#FFFFF0"},
                        {"Khaki",	    "#F0E68C"},
                        {"Lavender",	"#E6E6FA"},
                        {"LavenderBlush",	"#FFF0F5"},
                        {"LawnGreen",	"#7CFC00"},
                        {"LemonChiffon",	"#FFFACD"},
                        {"LightBlue",	"#ADD8E6"},
                        {"LightCoral",	"#F08080"},
                        {"LightCyan",	"#E0FFFF"},
                        {"LightGoldenRodYellow",	"#FAFAD2"},
                        {"LightGray",	"#D3D3D3"},
                        {"LightGrey",	"#D3D3D3"},
                        {"LightGreen",	"#90EE90"},
                        {"LightPink",	"#FFB6C1"},
                        {"LightSalmon",	"#FFA07A"},
                        {"LightSeaGreen",	"#20B2AA"},
                        {"LightSkyBlue",	"#87CEFA"},
                        {"LightSlateGray",	"#778899"},
                        {"LightSlateGrey",	"#778899"},
                        {"LightSteelBlue",	"#B0C4DE"},
                        {"LightYellow",	"#FFFFE0"},
                        {"Lime",	    "#00FF00"},
                        {"LimeGreen",	"#32CD32"},
                        {"Linen",	    "#FAF0E6"},
                        {"Magenta",	    "#FF00FF"},
                        {"Maroon",	    "#800000"},
                        {"MediumAquaMarine",	"#66CDAA"},
                        {"MediumBlue",	"#0000CD"},
                        {"MediumOrchid",	"#BA55D3"},
                        {"MediumPurple",	"#9370D8"},
                        {"MediumSeaGreen",	"#3CB371"},
                        {"MediumSlateBlue",	"#7B68EE"},
                        {"MediumSpringGreen",	"#00FA9A"},
                        {"MediumTurquoise",	"#48D1CC"},
                        {"MediumVioletRed",	"#C71585"},
                        {"MidnightBlue",	"#191970"},
                        {"MintCream",	"#F5FFFA"},
                        {"MistyRose",	"#FFE4E1"},
                        {"Moccasin",	"#FFE4B5"},
                        {"NavajoWhite",	"#FFDEAD"},
                        {"Navy",	    "#000080"},
                        {"OldLace",	    "#FDF5E6"},
                        {"Olive",	    "#808000"},
                        {"OliveDrab",	"#6B8E23"},
                        {"Orange",	    "#FFA500"},
                        {"OrangeRed",	"#FF4500"},
                        {"Orchid",	    "#DA70D6"},
                        {"PaleGoldenRod",	"#EEE8AA"},
                        {"PaleGreen",	"#98FB98"},
                        {"PaleTurquoise",	"#AFEEEE"},
                        {"PaleVioletRed",	"#D87093"},
                        {"PapayaWhip",	"#FFEFD5"},
                        {"PeachPuff",	"#FFDAB9"},
                        {"Peru",	    "#CD853F"},
                        {"Pink",	    "#FFC0CB"},
                        {"Plum",	    "#DDA0DD"},
                        {"PowderBlue",	"#B0E0E6"},
                        {"Purple",	    "#800080"},
                        {"Red",	        "#FF0000"},
                        {"RosyBrown",	"#BC8F8F"},
                        {"RoyalBlue",	"#4169E1"},
                        {"SaddleBrown",	"#8B4513"},
                        {"Salmon",	    "#FA8072"},
                        {"SandyBrown",	"#F4A460"},
                        {"SeaGreen",	"#2E8B57"},
                        {"SeaShell",	"#FFF5EE"},
                        {"Sienna",	    "#A0522D"},
                        {"Silver",	    "#C0C0C0"},
                        {"SkyBlue",	    "#87CEEB"},
                        {"SlateBlue",	"#6A5ACD"},
                        {"SlateGray",	"#708090"},
                        {"SlateGrey",	"#708090"},
                        {"Snow",	    "#FFFAFA"},
                        {"SpringGreen",	"#00FF7F"},
                        {"SteelBlue",	"#4682B4"},
                        {"Tan",	        "#D2B48C"},
                        {"Teal",	    "#008080"},
                        {"Thistle",	    "#D8BFD8"},
                        {"Tomato",	    "#FF6347"},
                        {"Turquoise",	"#40E0D0"},
                        {"Violet",	    "#EE82EE"},
                        {"Wheat",	    "#F5DEB3"},
                        {"White",	    "#FFFFFF"},
                        {"WhiteSmoke",	"#F5F5F5"},
                        {"Yellow",	    "#FFFF00"},
                        {"YellowGreen",	"#9ACD32"}};

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

        #region Utils
        private static string ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

        private double ToHi(string strValue)
        {
            string stSize = strValue.Trim();
            double factor = 1 / 14.4;
            //if (strValue.EndsWith("cm"))
            //{
            //    stSize = strValue.Substring(0, strValue.Length - 2);
            //    factor = 100 / 2.54;
            //}
            //if (strValue.EndsWith("in"))
            //{
            //    stSize = strValue.Substring(0, strValue.Length - 2);
            //    factor = 100;
            //}
            //if (strValue.EndsWith("pt"))
            //{
            //    stSize = strValue.Substring(0, strValue.Length - 2);
            //    //factor = 100 / 72f;
            //}
            if (stSize.Length == 0)
            {
                ThrowError(null, null, "Expression in SizeType: " + strValue);
                return 0;
            }
            double size = Convert.ToDouble(stSize.Replace(",", ".").Replace(".", ApplicationDecimalSeparator)) * factor;
            return Math.Round(size, 2);
        }

        private static double ParseDouble(string strValue)
        {
            return Convert.ToDouble(strValue.Trim().Replace(",", ".").Replace(".", ApplicationDecimalSeparator));
        }


        private static Color ParseColor(XmlNode node)
        {
            uint colorValue = (uint)int.Parse(node.Value);
            return Color.FromArgb(0xFF, (int)(colorValue & 0xFF), (int)((colorValue >> 8) & 0xFF), (int)((colorValue >> 16) & 0xFF));
        }


        private static StiPenStyle ParseLineStyle(string styleNumber)
        {
            switch (styleNumber)
            {
                case "0":
                    return StiPenStyle.None;
                case "2":
                    return StiPenStyle.Dash;
                case "3":
                    return StiPenStyle.Dot;
                case "4":
                    return StiPenStyle.DashDot;
                case "5":
                    return StiPenStyle.DashDotDot;
            }
            return StiPenStyle.Solid;
        }

        private static int ParseHexTwoCharToInt(char c1, char c2)
        {
            return ParseHexCharToInt(char.ToLowerInvariant(c1)) * 16 + ParseHexCharToInt(char.ToLowerInvariant(c2));
        }
        private static int ParseHexCharToInt(char ch)
        {
            if (ch <= '9') return (int)ch - (int)'0';
            if (ch == 'a') return 10;
            if (ch == 'b') return 11;
            if (ch == 'c') return 12;
            if (ch == 'd') return 13;
            if (ch == 'e') return 14;
            if (ch == 'f') return 15;
            return 0;
        }

        private void ThrowError(string baseNodeName, string nodeName)
        {
            ThrowError(baseNodeName, nodeName, null);
        }

        private void ThrowError(string baseNodeName, string nodeName, string message1)
        {
            string message = null;
            if (message1 == null)
            {
                message = string.Format("Node not found: {0}.{1}", baseNodeName, nodeName);
            }
            else
            {
                message = string.Format("{0}", message1);
            }
            errorList.Add(message);
        }
        

        public string ConvertExpression(string baseExpression, StiReport report)
        {
            if (baseExpression.StartsWith("param:"))
            {
                string name = baseExpression.Substring(6);
                if (!report.Dictionary.Variables.Contains(name))
                {
                    StiVariable var = new StiVariable("Parameters", name, typeof(string));
                    report.Dictionary.Variables.Add(var);
                }
                return name;
            }
            if (!fieldsNames.ContainsKey(baseExpression))
            {
                fields.Add(baseExpression);
                fieldsNames[baseExpression] = null;
            }
            return "ds." + baseExpression;
        }
        #endregion

        #region Methods.Import
        public static StiImportResult Import(byte[] bytes)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = StiCultureInfo.GetEN(false);

                var report = new StiReport();
                var helper = new StiActiveReportsHelper();
                var errors = new List<string>();
                var doc = new XmlDocument();

                using (var stream = new MemoryStream(bytes))
                {
                    doc.Load(stream);
                }

                helper.ProcessRootNode(doc.DocumentElement, report, errors);

                return new StiImportResult(report, errors);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }
        #endregion

    }
}
