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
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Units;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.ShapeTypes;
using Stimulsoft.Report.Dictionary;

namespace Stimulsoft.Report.Import
{
    public class StiFastReportsHelper
    {
        #region Fields
        List<string> errorList = null;
        StiReport report = null;
        bool isFr3 = false;
        #endregion

        #region Methods
        public StiReport Convert(byte[] bytes, List<string> errorList = null)
        {
            this.errorList = errorList;

            report = new StiReport();
            report.Pages.Clear();

            isFr3 = bytes[39] == 'T' || bytes[40] == 'T' || bytes[41] == 'T';
            if (isFr3)
            {
                bytes = RemakeFR3(bytes);
            }

            XmlDocument doc = new XmlDocument();
            using (var stream = new MemoryStream(bytes))
            {
                doc.Load(stream);
            }

            XmlNodeList documentList = doc.GetElementsByTagName("Report");
            if (documentList.Count == 0)
                documentList = doc.GetElementsByTagName("TfrxReport");
            if (documentList.Count > 0)
            {
                var reportNode = documentList[0];
                if (reportNode != null)
                {
                    report.ReportName = reportNode.Attributes["ReportInfo.Name"]?.Value ?? "Report";
                    report.ReportDescription = reportNode.Attributes["ReportInfo.Description"]?.Value ?? "";
                    report.ReportAuthor = reportNode.Attributes["ReportInfo.Author"]?.Value ?? "";
                }

                //first pass, assemble dictionary
                foreach (XmlNode documentNode in documentList)
                {
                    foreach (XmlNode elementNode in documentNode.ChildNodes)
                    {
                        switch (elementNode.Name)
                        {
                            case "Dictionary":
                                ReadDictionary(elementNode, report);
                                break;

                            case "ReportPage":
                            case "TfrxReportPage":
                                ReadPage(elementNode, report);
                                break;
                        }
                    }
                }
                report.Pages.Clear();

                //second pass
                foreach (XmlNode documentNode in documentList)
                {
                    foreach (XmlNode elementNode in documentNode.ChildNodes)
                    {
                        switch (elementNode.Name)
                        {
                            case "Dictionary":
                                //ReadDictionary(elementNode, report);
                                break;

                            case "ReportPage":
                            case "TfrxReportPage":
                                ReadPage(elementNode, report);
                                break;

                            default:
                                ThrowError(documentNode.Name, elementNode.Name);
                                break;
                        }
                    }
                }
            }

            report.Info.ShowHeaders = true;
            foreach (StiPage page in report.Pages)
            {
                page.DockToContainer();
                page.Correct();
            }

            return report;
        }
        #endregion

        #region Methods.Read.Dictionary
        private void ReadDictionary(XmlNode dictNode, StiReport report)
        {
            foreach (XmlNode elementNode in dictNode.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "MsAccessDataConnection":
                        ReadMsAccessDataConnection(elementNode, report);
                        break;

                    case "XmlDataConnection":
                        ReadXmlDataConnection(elementNode, report);
                        break;

                    case "OleDbDataConnection":
                        ReadOleDbDataConnection(elementNode, report);
                        break;

                    case "OdbcDbDataConnection":
                        ReadOdbcDataConnection(elementNode, report);
                        break;

                    case "TableDataSource":
                        ReadTableDataSource(elementNode, report, null);
                        break;

                    case "Relation":
                        ReadRelation(elementNode, report);
                        break;

                    default:
                        ThrowError(dictNode.Name, elementNode.Name);
                        break;
                }
            }
        }

        private void ReadMsAccessDataConnection(XmlNode node, StiReport report)
        {
            StiMSAccessDatabase database = new StiMSAccessDatabase();
            report.Dictionary.Databases.Add(database);

            database.Name = ReadString(node, "Name", database.Name);

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "TableDataSource":
                        ReadTableDataSource(elementNode, report, database.Name);
                        break;
                }
            }
        }

        private void ReadXmlDataConnection(XmlNode node, StiReport report)
        {
            StiXmlDatabase database = new StiXmlDatabase();
            report.Dictionary.Databases.Add(database);

            database.Name = ReadString(node, "Name", database.Name);

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "TableDataSource":
                        ReadTableDataSource(elementNode, report, database.Name);
                        break;
                }
            }
        }

        private void ReadOleDbDataConnection(XmlNode node, StiReport report)
        {
            StiOleDbDatabase database = new StiOleDbDatabase();
            report.Dictionary.Databases.Add(database);

            database.Name = ReadString(node, "Name", database.Name);

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "TableDataSource":
                        ReadTableDataSource(elementNode, report, database.Name);
                        break;
                }
            }
        }

        private void ReadOdbcDataConnection(XmlNode node, StiReport report)
        {
            StiOdbcDatabase database = new StiOdbcDatabase();
            report.Dictionary.Databases.Add(database);

            database.Name = ReadString(node, "Name", database.Name);

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "TableDataSource":
                        ReadTableDataSource(elementNode, report, database.Name);
                        break;
                }
            }
        }

        private void ReadTableDataSource(XmlNode node, StiReport report, string connectionName)
        {
            StiDataTableSource source = new StiDataTableSource();
            report.DataSources.Add(source);

            if (connectionName == null)
                source.NameInSource = ReadString(node, "ReferenceName", source.Name);
            else
                source.NameInSource = connectionName;

            source.Name = source.Alias = ReadString(node, "Name", source.Name);

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "Column":
                        ReadColumn(elementNode, source);
                        break;
                }
            }
        }

        private void ReadColumn(XmlNode node, StiDataTableSource source)
        {
            StiDataColumn column = new StiDataColumn();
            source.Columns.Add(column);

            column.Name = column.Alias = ReadString(node, "Name", column.Name);
            string type = ReadString(node, "DataType", string.Empty);
            column.Type = Type.GetType(type);

        }

        private void ReadRelation(XmlNode node, StiReport report)
        {
            StiDataRelation relation = new StiDataRelation();
            report.Dictionary.Relations.Add(relation);
            relation.NameInSource = ReadString(node, "ReferenceName", relation.Name);
            relation.Name = relation.Alias = ReadString(node, "Name", relation.Name);

            string parentDataSource = ReadString(node, "ParentDataSource", "");
            string childDataSource = ReadString(node, "ChildDataSource", "");

            string parentColumnsString = ReadString(node, "ParentColumns", "");
            string childColumnsString = ReadString(node, "ChildColumns", "");

            string[] parentColumns = parentColumnsString.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            string[] childColumns = parentColumnsString.Split(new string[] { "\r\n" }, StringSplitOptions.None);

            relation.ParentSource = report.DataSources[parentDataSource];
            relation.ChildSource = report.DataSources[childDataSource];
            relation.ParentColumns = parentColumns;
            relation.ChildColumns = childColumns;

        }
        #endregion

        #region Methods.Read.Values
        private static string ReadString(XmlNode node, string name, string defaultValue)
        {
            XmlAttribute attrName = node.Attributes[name];
            if (attrName != null) return attrName.Value;
            else return defaultValue;
        }

        private static bool ReadBool(XmlNode node, string name, StiComponent comp, bool defaultValue)
        {
            XmlAttribute attr = node.Attributes[name];
            if (attr != null)
                return attr.Value.ToLowerInvariant() == "true";
            else
                return defaultValue;//default value
        }

        private static double ReadDouble(XmlNode node, string name, StiComponent comp, double defaultValue)
        {
            XmlAttribute attr = node.Attributes[name];
            if (attr != null)
                return double.Parse(attr.Value.Replace(".", ",").Replace(",", Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator));
            else
                return defaultValue;
        }

        private static int ReadInt(XmlNode node, string name, StiComponent comp, int defaultValue)
        {
            XmlAttribute attr = node.Attributes[name];
            if (attr != null)
                return int.Parse(attr.Value);
            else
                return defaultValue;
        }

        private static double ReadValueFromMillimeters(XmlNode node, string name, StiPage page, double defaultValue)
        {
            XmlAttribute attr = node.Attributes[name];
            if (attr != null)
                return ConvertFromMillimeters(double.Parse(attr.Value), page);
            else
                return ConvertFromMillimeters(defaultValue, page);//default value
        }
        #endregion

        #region Methods.Convert
        private static double ConvertFromMillimeters(double value, StiComponent comp)
        {
            StiMillimetersUnit mmUnit = new StiMillimetersUnit();
            value = mmUnit.ConvertToHInches(value);
            return comp.Report.Unit.ConvertFromHInches(value);
        }

        private static double ConvertFromFRUnit(double value, StiComponent comp)
        {
            value = value / 0.96;
            return comp.Report.Unit.ConvertFromHInches(value);
        }
        #endregion

        #region Methods.Read.Properties
        private Color ReadColor(XmlNode node, string name)
        {
            XmlAttribute attr = node.Attributes[name];
            if (attr == null) return Color.Black;

            ColorConverter colorConverter = new ColorConverter();
            try
            {
                return (Color)colorConverter.ConvertFromString(attr.Value);
            }
            catch
            {
            }
            return Color.Black;
        }

        private void ReadBorder(XmlNode node, StiComponent comp)
        {
            try
            {
                if (comp is IStiBorder)
                {
                    StiBorder border = ((IStiBorder)comp).Border;

                    #region Border.Color
                    border.Color = ReadColor(node, "Border.Color");
                    #endregion

                    #region Border.Lines
                    XmlAttribute attr = node.Attributes["Border.Lines"];
                    if (attr != null)
                    {
                        if (attr.Value == "All") border.Side = StiBorderSides.All;
                        else
                        {
                            if (attr.Value.Contains("Left")) border.Side |= StiBorderSides.Left;
                            if (attr.Value.Contains("Right")) border.Side |= StiBorderSides.Right;
                            if (attr.Value.Contains("Top")) border.Side |= StiBorderSides.Top;
                            if (attr.Value.Contains("Bottom")) border.Side |= StiBorderSides.Bottom;
                        }
                    }
                    #endregion

                    #region Border.Width
                    border.Size = ReadDouble(node, "Border.Width", comp, 1d);
                    #endregion

                    #region Border.Style
                    attr = node.Attributes["Border.Style"];
                    if (attr != null)
                    {
                        switch (attr.Value)
                        {
                            case "Dash":
                                border.Style = StiPenStyle.Dash;
                                break;

                            case "DashDot":
                                border.Style = StiPenStyle.DashDot;
                                break;

                            case "DashDotDot":
                                border.Style = StiPenStyle.DashDotDot;
                                break;

                            case "Dot":
                                border.Style = StiPenStyle.Dot;
                                break;

                            case "Double":
                                border.Style = StiPenStyle.Double;
                                break;

                            case "Solid":
                                border.Style = StiPenStyle.Solid;
                                break;
                        }
                    }
                    #endregion

                    #region Border.Shadow
                    border.DropShadow = ReadBool(node, "Border.Shadow", comp, false);
                    #endregion

                    #region Border.ShadowWidth
                    border.ShadowSize = ReadDouble(node, "Border.ShadowWidth", comp, 4d);
                    #endregion

                    #region Border.ShadowColor
                    attr = node.Attributes["Border.ShadowColor"];
                    if (attr != null) border.ShadowBrush = new StiSolidBrush(ReadColor(node, "Border.ShadowColor"));
                    #endregion

                    int sides = ReadInt(node, "Frame.Typ", comp, 0);
                    if ((sides & 0x01) > 0) border.Side |= StiBorderSides.Left;
                    if ((sides & 0x02) > 0) border.Side |= StiBorderSides.Top;
                    if ((sides & 0x04) > 0) border.Side |= StiBorderSides.Right;
                    if ((sides & 0x08) > 0) border.Side |= StiBorderSides.Bottom;
                }
            }
            catch
            {
            }
        }

        private void ReadBrush(XmlNode node, StiComponent comp)
        {
            if (comp is IStiBrush)
            {
                IStiBrush brushComp = comp as IStiBrush;
                brushComp.Brush = GetBrush(node, comp, "Fill");
            }
        }

        private void ReadTextBrush(XmlNode node, StiComponent comp)
        {
            if (comp is IStiTextBrush)
            {
                IStiTextBrush brushComp = comp as IStiTextBrush;
                brushComp.TextBrush = GetBrush(node, comp, "TextFill");
            }
        }

        private StiBrush GetBrush(XmlNode node, StiComponent comp, string key)
        {
            try
            {
                XmlAttribute attr = node.Attributes[key];
                if (attr != null)
                {
                    switch (attr.Value)
                    {
                        case "LinearGradient":
                            StiGradientBrush gradientBrush = new StiGradientBrush();
                            gradientBrush.StartColor = ReadColor(node, key + ".StartColor");
                            gradientBrush.EndColor = ReadColor(node, key + ".EndColor");
                            gradientBrush.Angle = (float)ReadDouble(node, key + ".Angle", comp, 0d);
                            return gradientBrush;

                        case "PathGradient":
                            StiGlareBrush glareBrush = new StiGlareBrush();
                            glareBrush.StartColor = ReadColor(node, key + ".CenterColor");
                            glareBrush.EndColor = ReadColor(node, key + ".EdgeColor");
                            return glareBrush;

                        case "Hatch":
                            StiHatchBrush hatchBrush = new StiHatchBrush();
                            hatchBrush.ForeColor = ReadColor(node, key + ".ForeColor");
                            hatchBrush.BackColor = ReadColor(node, key + ".BackColor");
                            string str = ReadString(node, key + ".Style", "BackwardDiagonal");
                            hatchBrush.Style = (HatchStyle)Enum.Parse(typeof(HatchStyle), str, false);
                            return hatchBrush;

                        case "Glass":
                            StiGlassBrush glassBrush = new StiGlassBrush();
                            glassBrush.Color = ReadColor(node, key + ".Color");
                            glassBrush.Blend = (float)ReadDouble(node, key + ".Blend", comp, 0.2d);
                            glassBrush.DrawHatch = ReadBool(node, key + ".Hatch", comp, false);
                            return glassBrush;
                    }
                }
                else
                {
                    XmlAttribute attr2 = node.Attributes[key + ".Color"];
                    if (attr2 != null)
                    {
                        StiSolidBrush solidBrush = new StiSolidBrush();
                        solidBrush.Color = ReadColor(node, key + ".Color");
                        return solidBrush;
                    }
                }

            }
            catch
            {
            }
            if (key == "Fill")
                return new StiEmptyBrush();
            else
                return new StiSolidBrush(Color.Black);
        }
        #endregion

        #region Methods.Read.Bands
        private void ReadBand(XmlNode node, StiBand band)
        {
            ReadComp(node, band);

            band.CanBreak = ReadBool(node, "CanBreak", band, true);

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "TextObject":
                    case "TfrxMemoView":
                        ReadTextObject(elementNode, band);
                        break;

                    case "PictureObject":
                    case "TfrxPictureView":
                        ReadPictureObject(elementNode, band);
                        break;

                    case "CheckBoxObject":
                        ReadCheckBoxObject(elementNode, band);
                        break;

                    case "ShapeObject":
                    case "TfrxShapeView":
                        ReadShapeObject(elementNode, band);
                        break;

                    case "LineObject":
                        ReadLineObject(elementNode, band);
                        break;

                    case "ChildBand":
                        ReadChildBand(elementNode, band.Page);
                        break;

                    case "DataHeaderBand":
                        ReadDataHeaderBand(elementNode, band.Page, band);
                        break;

                    case "DataBand":
                        ReadDataBand(elementNode, band.Page, band);
                        break;

                    case "DataFooterBand":
                        ReadDataFooterBand(elementNode, band.Page, band);
                        break;

                    case "GroupHeaderBand":
                        ReadGroupHeaderBand(elementNode, band.Page);
                        break;

                    case "GroupFooterBand":
                        ReadGroupFooterBand(elementNode, band.Page);
                        break;

                    case "SubreportObject":
                        ReadSubreportObject(elementNode, band);
                        break;

                    default:
                        ThrowError(node.Name, elementNode.Name);
                        break;
                }
            }

            foreach (StiComponent comp in band.Components)
            {
                if (comp.Left > band.Width) comp.Linked = true;
            }
        }

        private void ReadReportTitleBand(XmlNode node, StiPage page)
        {
            StiReportTitleBand band = new StiReportTitleBand();
            page.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadReportSummaryBand(XmlNode node, StiPage page)
        {
            StiReportSummaryBand band = new StiReportSummaryBand();
            page.Components.Add(band);

            ReadBand(node, band);

            band.KeepReportSummaryTogether = ReadBool(node, "KeepWithData", band, false);
        }

        private void ReadChildBand(XmlNode node, StiPage page)
        {
            StiChildBand band = new StiChildBand();
            page.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadDataBand(XmlNode node, StiPage page, StiBand masterBand)
        {
            StiDataBand band = new StiDataBand();
            page.Components.Add(band);

            band.MasterComponent = masterBand;

            string dataSource = ReadString(node, "DataSource", "");
            if (string.IsNullOrEmpty(dataSource))
                dataSource = ReadString(node, "DataSetName", "");
            if (string.IsNullOrEmpty(dataSource))
                dataSource = ReadString(node, "DataSet", "");
            band.DataSourceName = dataSource;

            AddDataSourceName(dataSource);

            ReadBand(node, band);
        }

        private void ReadDataHeaderBand(XmlNode node, StiPage page, StiBand masterBand)
        {
            StiHeaderBand band = new StiHeaderBand();
            if (masterBand != null)
                page.Components.Insert(page.Components.IndexOf(masterBand), band);
            else
                page.Components.Add(band);

            ReadBand(node, band);

            band.KeepHeaderTogether = ReadBool(node, "KeepWithData", band, false);
        }

        private void ReadDataFooterBand(XmlNode node, StiPage page, StiBand masterBand)
        {
            StiFooterBand band = new StiFooterBand();
            page.Components.Add(band);

            ReadBand(node, band);

            band.KeepFooterTogether = ReadBool(node, "KeepWithData", band, false);
        }

        private void ReadPageHeaderBand(XmlNode node, StiPage page)
        {
            StiPageHeaderBand band = new StiPageHeaderBand();
            page.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadPageFooterBand(XmlNode node, StiPage page)
        {
            StiPageFooterBand band = new StiPageFooterBand();
            page.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadColumnHeaderBand(XmlNode node, StiPage page)
        {
            StiColumnHeaderBand band = new StiColumnHeaderBand();
            page.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadColumnFooterBand(XmlNode node, StiPage page)
        {
            StiColumnFooterBand band = new StiColumnFooterBand();
            page.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadGroupHeaderBand(XmlNode node, StiPage page)
        {
            StiGroupHeaderBand band = new StiGroupHeaderBand();
            page.Components.Add(band);

            ReadBand(node, band);

            string condition = ReadString(node, "Condition", band.Name);
            band.Condition.Value = ProcessExpression(condition);
        }

        private void ReadGroupFooterBand(XmlNode node, StiPage page)
        {
            StiGroupFooterBand band = new StiGroupFooterBand();
            page.Components.Add(band);

            ReadBand(node, band);
        }
        #endregion

        #region Methods.Read.Components
        private void ReadPage(XmlNode node, StiReport report)
        {
            StiPage page = new StiPage(report);
            report.Pages.Add(page);

            page.Name = ReadString(node, "Name", page.Name);
            page.Orientation = ReadBool(node, "Landscape", page, false) ? StiPageOrientation.Landscape : StiPageOrientation.Portrait;
            page.PageWidth = ReadValueFromMillimeters(node, "PaperWidth", page, 210);
            page.PageHeight = ReadValueFromMillimeters(node, "PaperHeight", page, 297);

            page.Margins.Left = ReadValueFromMillimeters(node, "LeftMargin", page, 10d);
            page.Margins.Top = ReadValueFromMillimeters(node, "TopMargin", page, 10d);
            page.Margins.Right = ReadValueFromMillimeters(node, "RightMargin", page, 10d);
            page.Margins.Bottom = ReadValueFromMillimeters(node, "BottomMargin", page, 10d);

            page.PrintOnPreviousPage = ReadBool(node, "PrintOnPreviousPage", page, false);
            page.ResetPageNumber = ReadBool(node, "ResetPageNumber", page, false);
            page.TitleBeforeHeader = ReadBool(node, "TitleBeforeHeader", page, true);
            page.Enabled = ReadBool(node, "Visible", page, true);

            page.Columns = ReadInt(node, "Columns.Count", page, 0);
            page.ColumnWidth = ReadValueFromMillimeters(node, "Columns.Width", page, page.ColumnWidth);

            #region Sort bands by Top
            List<XmlNode> nodes = node.ChildNodes.Cast<XmlNode>().ToList();
            nodes.Sort((x, y) =>
            {
                double xTop = double.Parse(x.Attributes["Top"]?.Value ?? "0", CultureInfo.InvariantCulture);
                double yTop = double.Parse(y.Attributes["Top"]?.Value ?? "0", CultureInfo.InvariantCulture);
                return xTop.CompareTo(yTop);
            });
            foreach (XmlNode node2 in node.ChildNodes)
            {
                node.RemoveChild(node2);
            }
            foreach (XmlNode node2 in nodes)
            {
                node.AppendChild(node2);
            }
            #endregion

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "ReportTitleBand":
                    case "TfrxReportTitle":
                        ReadReportTitleBand(elementNode, page);
                        break;

                    case "ReportSummaryBand":
                    case "TfrxReportSummary":
                        ReadReportSummaryBand(elementNode, page);
                        break;

                    case "DataBand":
                    case "TfrxMasterData":
                        ReadDataBand(elementNode, page, null);
                        break;

                    case "TfrxHeader":
                        ReadDataHeaderBand(elementNode, page, null);
                        break;

                    case "TfrxFooter":
                        ReadDataFooterBand(elementNode, page, null);
                        break;

                    case "GroupHeaderBand":
                    case "TfrxGroupHeader":
                        ReadGroupHeaderBand(elementNode, page);
                        break;

                    case "PageHeaderBand":
                    case "TfrxPageHeader":
                        ReadPageHeaderBand(elementNode, page);
                        break;

                    case "PageFooterBand":
                    case "TfrxPageFooter":
                        ReadPageFooterBand(elementNode, page);
                        break;

                    case "ColumnHeaderBand":
                        ReadColumnHeaderBand(elementNode, page);
                        break;

                    case "ColumnFooterBand":
                        ReadColumnFooterBand(elementNode, page);
                        break;

                    case "TfrxMemoView":
                        ReadTextObject(elementNode, page);
                        break;

                    case "PictureObject":
                    case "TfrxPictureView":
                        ReadPictureObject(elementNode, page);
                        break;

                    case "CheckBoxObject":
                        ReadCheckBoxObject(elementNode, page);
                        break;

                    case "ShapeObject":
                    case "TfrxShapeView":
                        ReadShapeObject(elementNode, page);
                        break;

                    default:
                        ThrowError(node.Name, elementNode.Name);
                        break;
                }
            }
        }

        private void ReadComp(XmlNode node, StiComponent comp)
        {
            comp.Name = ReadString(node, "Name", comp.Name);

            comp.Left = ConvertFromFRUnit(ReadDouble(node, "Left", comp, comp.Left), comp);
            comp.Top = ConvertFromFRUnit(ReadDouble(node, "Top", comp, comp.Top), comp);
            comp.Width = ConvertFromFRUnit(ReadDouble(node, "Width", comp, comp.Width), comp);
            comp.Height = ConvertFromFRUnit(ReadDouble(node, "Height", comp, comp.Height), comp);

            comp.CanGrow = ReadBool(node, "CanGrow", comp, true);
            comp.CanShrink = ReadBool(node, "CanShrink", comp, true);
            comp.Printable = ReadBool(node, "Printable", comp, true);
            comp.Enabled = ReadBool(node, "Enabled", comp, true);
            comp.GrowToHeight = ReadBool(node, "GrowToBottom", comp, false);
            comp.ComponentStyle = ReadString(node, "Style", "");
            comp.Bookmark.Value = ReadString(node, "Bookmark", "");
            comp.Hyperlink.Value = ReadString(node, "Hyperlink.Expression", "");

            if (comp.Left < 0) comp.Linked = true;

            #region Dock
            switch (ReadString(node, "Dock", "None"))
            {
                case "Bottom":
                    comp.DockStyle = StiDockStyle.Bottom;
                    break;

                case "Fill":
                    comp.DockStyle = StiDockStyle.Fill;
                    break;

                case "Left":
                    comp.DockStyle = StiDockStyle.Left;
                    break;

                case "None":
                    comp.DockStyle = StiDockStyle.None;
                    break;

                case "Right":
                    comp.DockStyle = StiDockStyle.Right;
                    break;

                case "Top":
                    comp.DockStyle = StiDockStyle.Top;
                    break;
            }
            #endregion

            #region ShiftMode
            switch (ReadString(node, "ShiftMode", "Always"))
            {
                case "Always":
                    comp.ShiftMode = StiShiftMode.IncreasingSize | StiShiftMode.DecreasingSize;
                    break;

                case "WhenOverlapped":
                    comp.ShiftMode = StiShiftMode.IncreasingSize;
                    break;

                case "Never":
                    comp.ShiftMode = (StiShiftMode)0;
                    break;
            }
            #endregion

            #region Restrictions
            XmlAttribute attr = node.Attributes["Restrictions"];
            if (attr != null)
            {
                if (attr.Value.Contains("DontModify")) comp.Restrictions &= Stimulsoft.Report.Components.StiRestrictions.AllowChange;
                if (attr.Value.Contains("DontDelete")) comp.Restrictions &= Stimulsoft.Report.Components.StiRestrictions.AllowDelete;
                if (attr.Value.Contains("DontMove")) comp.Restrictions &= Stimulsoft.Report.Components.StiRestrictions.AllowMove;
                if (attr.Value.Contains("DontResize")) comp.Restrictions &= Stimulsoft.Report.Components.StiRestrictions.AllowResize;
            }
            #endregion

            #region PrintOn
            attr = node.Attributes["PrintOn"];
            if (attr != null)
            {
                if (!attr.Value.Contains("FirstPage") && !attr.Value.Contains("LastPage"))
                    comp.PrintOn = StiPrintOnType.ExceptFirstAndLastPage;

                else if (!attr.Value.Contains("FirstPage") && attr.Value.Contains("LastPage"))
                    comp.PrintOn = StiPrintOnType.ExceptFirstPage;

                else if (attr.Value.Contains("FirstPage") && !attr.Value.Contains("LastPage"))
                    comp.PrintOn = StiPrintOnType.ExceptLastPage;

                IStiPrintOnEvenOddPages printOn = comp as IStiPrintOnEvenOddPages;
                if (printOn != null)
                {
                    if (attr.Value.Contains("OddPages") && attr.Value.Contains("EvenPages"))
                        printOn.PrintOnEvenOddPages = StiPrintOnEvenOddPagesType.Ignore;

                    else if (!attr.Value.Contains("OddPages") && attr.Value.Contains("EvenPages"))
                        printOn.PrintOnEvenOddPages = StiPrintOnEvenOddPagesType.PrintOnEvenPages;

                    else if (attr.Value.Contains("OddPages") && !attr.Value.Contains("EvenPages"))
                        printOn.PrintOnEvenOddPages = StiPrintOnEvenOddPagesType.PrintOnOddPages;
                }
            }
            #endregion

            #region PrintOn
            attr = node.Attributes["PrintOn"];
            if (attr != null)
            {
                if (!attr.Value.Contains("FirstPage") && !attr.Value.Contains("LastPage"))
                    comp.PrintOn = StiPrintOnType.ExceptFirstAndLastPage;

                else if (!attr.Value.Contains("FirstPage") && attr.Value.Contains("LastPage"))
                    comp.PrintOn = StiPrintOnType.ExceptFirstPage;

                else if (attr.Value.Contains("FirstPage") && !attr.Value.Contains("LastPage"))
                    comp.PrintOn = StiPrintOnType.ExceptLastPage;

                IStiPrintOnEvenOddPages printOn = comp as IStiPrintOnEvenOddPages;
                if (printOn != null)
                {
                    if (attr.Value.Contains("OddPages") && attr.Value.Contains("EvenPages"))
                        printOn.PrintOnEvenOddPages = StiPrintOnEvenOddPagesType.Ignore;

                    else if (!attr.Value.Contains("OddPages") && attr.Value.Contains("EvenPages"))
                        printOn.PrintOnEvenOddPages = StiPrintOnEvenOddPagesType.PrintOnEvenPages;

                    else if (attr.Value.Contains("OddPages") && !attr.Value.Contains("EvenPages"))
                        printOn.PrintOnEvenOddPages = StiPrintOnEvenOddPagesType.PrintOnOddPages;
                }
            }
            #endregion

            #region PrintOnBottom
            attr = node.Attributes["PrintOnBottom"];
            if (attr != null && comp is IStiPrintAtBottom)
            {
                ((IStiPrintAtBottom)comp).PrintAtBottom = ReadBool(node, "PrintOnBottom", comp, false);
            }
            #endregion

            #region StartNewPage
            attr = node.Attributes["StartNewPage"];
            if (attr != null && comp is IStiPageBreak)
            {
                ((IStiPageBreak)comp).NewPageBefore = ReadBool(node, "StartNewPage", comp, false);
            }
            #endregion

            #region RepeatOnEveryPage
            attr = node.Attributes["RepeatOnEveryPage"];
            if (attr != null && comp is IStiPrintOnAllPages)
            {
                ((IStiPrintOnAllPages)comp).PrintOnAllPages = ReadBool(node, "RepeatOnEveryPage", comp, false);
            }
            #endregion

            ReadBorder(node, comp);
            ReadBrush(node, comp);
            ReadTextBrush(node, comp);
        }

        private void ReadTextObject(XmlNode node, StiContainer container)
        {
            StiText text = new StiText();
            container.Components.Add(text);

            ReadComp(node, text);

            string st = ReadString(node, "Text", "");
            text.Text.Value = ParseText(st);

            text.Angle = -(float)ReadDouble(node, "Angle", text, 0d);
            if (isFr3)
                text.Angle = (float)ReadDouble(node, "Rotation", text, text.Angle);

            text.LinesOfUnderline = ReadBool(node, "Underlines", text, false) ? StiPenStyle.Solid : StiPenStyle.None;
            text.AutoWidth = ReadBool(node, "AutoWidth", text, false);
            text.RenderTo = ReadString(node, "BreakTo", "");
            text.HideZeros = ReadBool(node, "HideZeros", text, false);
            text.AllowHtmlTags = ReadBool(node, "HtmlTags", text, false);
            text.TextOptions.RightToLeft = ReadBool(node, "RightToLeft", text, false);
            text.WordWrap = ReadBool(node, "WordWrap", text, true);
            text.TextQuality = ReadBool(node, "Wysiwyg", text, false) ? StiTextQuality.Wysiwyg : StiTextQuality.Standard;
            text.OnlyText = ReadBool(node, "AllowExpressions", text, true) ? false : true;

            #region Trimming
            switch (ReadString(node, "Trimming", "None"))
            {
                case "None":
                    text.TextOptions.Trimming = StringTrimming.None;
                    break;

                case "Character":
                    text.TextOptions.Trimming = StringTrimming.Character;
                    break;

                case "EllipsisCharacter":
                    text.TextOptions.Trimming = StringTrimming.EllipsisCharacter;
                    break;

                case "EllipsisPath":
                    text.TextOptions.Trimming = StringTrimming.EllipsisPath;
                    break;

                case "EllipsisWord":
                    text.TextOptions.Trimming = StringTrimming.EllipsisWord;
                    break;

                case "Word":
                    text.TextOptions.Trimming = StringTrimming.Word;
                    break;

            }
            #endregion

            #region Duplicates
            switch (ReadString(node, "Duplicates", "Show"))
            {
                case "Show":
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.None;
                    break;

                case "Hide":
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.Hide;
                    break;

                case "Clear":
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.RemoveText;
                    break;

                case "Merge":
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.Merge;
                    break;
            }
            #endregion

            #region HorzAlign
            switch (ReadString(node, isFr3 ? "HAlign" : "HorzAlign", ""))
            {
                case "Center":
                case "haCenter":
                    text.HorAlignment = StiTextHorAlignment.Center;
                    break;

                case "Right":
                case "haRight":
                    text.HorAlignment = StiTextHorAlignment.Right;
                    break;

                case "Justify":
                case "haJustify":
                    text.HorAlignment = StiTextHorAlignment.Width;
                    break;

                default:
                    text.HorAlignment = StiTextHorAlignment.Left;
                    break;
            }
            #endregion

            #region VertAlign
            switch (ReadString(node, isFr3 ? "VAlign" : "VertAlign", ""))
            {
                case "Center":
                case "vaCenter":
                    text.VertAlignment = StiVertAlignment.Center;
                    break;

                case "Bottom":
                case "vaBottom":
                    text.VertAlignment = StiVertAlignment.Bottom;
                    break;

                default:
                    text.VertAlignment = StiVertAlignment.Top;
                    break;
            }
            #endregion

            #region Font
            try
            {
                XmlAttribute attrName = node.Attributes["Font"];
                if (attrName != null)
                {
                    //FontConverter fontConverter = new FontConverter();
                    //text.Font = fontConverter.ConvertFromString(attrName.Value) as Font;
                    text.Font = ParseFont(attrName.Value);
                }
                if (isFr3)
                {
                    var fontName = ReadString(node, "Font.Name", "");
                    var fontSize = Math.Abs(ReadDouble(node, "Font.Height", text, 0));
                    var fontStyleInt = ReadInt(node, "Font.Style", text, 0);
                    if (!string.IsNullOrWhiteSpace(fontName) && fontSize > 0)
                    {
                        FontStyle fs = FontStyle.Regular;
                        if ((fontStyleInt & 0x01) > 0) fs |= FontStyle.Bold;
                        if ((fontStyleInt & 0x02) > 0) fs |= FontStyle.Italic;

                        text.Font = new Font(fontName, (float)fontSize, fs);
                    }
                }
            }
            catch
            {
            }
            #endregion
        }

        private void ReadPictureObject(XmlNode node, StiContainer container)
        {
            StiImage image = new StiImage();
            container.Components.Add(image);

            ReadComp(node, image);

            image.DataColumn = ReadString(node, "DataColumn", "");

            string hex = ReadString(node, "Picture.PropData", "");
            if (!string.IsNullOrWhiteSpace(hex))
            {
                int length = hex.Length;
                byte[] bytes = new byte[length / 2];
                for (int i = 0; i < length; i += 2)
                {
                    bytes[i / 2] = System.Convert.ToByte(hex.Substring(i, 2), 16);
                }

                try
                {
                    var image2 = System.Drawing.Image.FromStream(new MemoryStream(bytes));
                    //if image is loaded - then assign it to component
                    image.ImageBytes = bytes;
                }
                catch { }
            }
        }

        private void ReadCheckBoxObject(XmlNode node, StiContainer container)
        {
            StiCheckBox checkBox = new StiCheckBox();
            container.Components.Add(checkBox);

            ReadComp(node, checkBox);

            checkBox.Checked.Value = "{" + ReadString(node, "DataColumn", "") + "}";
        }

        private void ReadLineObject(XmlNode node, StiContainer container)
        {
            StiHorizontalLinePrimitive line = new StiHorizontalLinePrimitive();
            container.Components.Add(line);

            ReadComp(node, line);
        }

        private void ReadShapeObject(XmlNode node, StiContainer container)
        {
            var shape = new StiShape();
            container.Components.Add(shape);

            ReadComp(node, shape);

            var stType = ReadString(node, "Shape", "");
            if (stType == "skTriangle") shape.ShapeType = new StiTriangleShapeType();
            if (stType == "skRoundRectangle") shape.ShapeType = new StiRoundedRectangleShapeType();
            if (stType == "skEllipse") shape.ShapeType = new StiOvalShapeType();
            if (stType == "skRectangle") shape.ShapeType = new StiRectangleShapeType();
            if (stType == "skDiamond") shape.ShapeType = new StiFlowchartDecisionShapeType();
            if (stType == "skCircle") shape.ShapeType = new StiOvalShapeType();
        }

        private void ReadSubreportObject(XmlNode node, StiContainer parent)
        {
            StiSubReport subReport = new StiSubReport();
            parent.Components.Add(subReport);

            ReadComp(node, subReport);
        }
        #endregion

        #region Methods.Utils
        private StiDataSource AddDataSourceName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;

            var ds = report.Dictionary.DataSources[name];
            if (ds != null) return ds;

            ds = new StiDataTableSource(name, name);
            report.Dictionary.DataSources.Add(ds);
            return ds;
        }

        private void AddDataSourceField(string dataSourceName, string dataSourceField)
        {
            var ds = AddDataSourceName(dataSourceName);
            if (ds == null) return;

            var field = ds.Columns[dataSourceField];
            if (field != null) return;

            field = new StiDataColumn(dataSourceField);
            ds.Columns.Add(field);
        }

        private string ParseText(string inputText)
        {
            StringBuilder output = new StringBuilder();
            StringBuilder sb = new StringBuilder();

            int level = 0;
            for (int index = 0; index < inputText.Length; index++)
            {
                char c = inputText[index];
                if (level > 0)
                {
                    if (c == '[')
                    {
                        level++;
                    }
                    if (c == ']')
                    {
                        level--;
                    }
                    if (level == 0)
                    {
                        output.Append("{" + ProcessExpression(sb.ToString()) + "}");
                        sb.Clear();
                    }
                    else
                    {
                        sb.Append(c);
                    }
                }
                else
                {
                    if (c == '[')
                    {
                        level++;
                    }
                    else
                    {
                        output.Append(c);
                    }
                }
            }

            if (sb.Length > 0) 
                output.Append("[" + sb);

            return output.ToString();
        }

        private string ProcessExpression(string baseInput)
        {
            string input = baseInput.Trim();

            if (input == "Page") input = "PageNumber";
            if (input == "TotalPages") input = "TotalPagesCount";
            if (input == "PageN") input = "\"Page \" + PageNumber";
            if (input == "Page#") input = "PageNumberThrough";
            if (input == "TotalPages#") input = "TotalPagesCountThrough";
            if (input == "Line#") input = "Line";

            input = input.Replace("Report.ReportInfo.Name", "ReportName");
            input = input.Replace("Report.ReportInfo.Description", "ReportDescription");
            input = input.Replace("Report.ReportInfo.Author", "ReportAuthor");

            foreach(StiDataSource ds in report.Dictionary.DataSources)
            {
                int pos = 0;
                while ((pos = input.IndexOf(ds.Name, pos)) != -1) 
                {
                    int pos2 = pos + ds.Name.Length;
                    if ((pos2 < input.Length) && (input.Substring(pos2, 2) == ".\""))
                    {
                        pos2 += 2;
                        int pos3 = input.IndexOf("\"", pos2);
                        if (pos3 != -1)
                        {
                            var fieldName = input.Substring(pos2, pos3 - pos2);
                            AddDataSourceField(ds.Name, fieldName);

                            var newName = ds.Name + "." + fieldName;
                            pos3++;
                            if ((pos > 0) && (input[pos - 1] == '<')) pos--;
                            if ((pos3 < input.Length) && (input[pos3] == '>')) pos3++;

                            input = input.Substring(0, pos) + newName + input.Substring(pos3);
                            pos += newName.Length;
                        }
                    }
                    else
                        pos++;
                }
            }

            return input;
        }

        private Font ParseFont(string fontString)
        {
            if (string.IsNullOrWhiteSpace(fontString))
                return new Font("Arial", 8);

            string[] parts = fontString.Split(',');

            string fontName = parts[0].Trim();
            float fontSize = 8;
            FontStyle fontStyle = FontStyle.Regular;
            GraphicsUnit fontUnit = GraphicsUnit.Point;

            if (parts.Length > 1)
            {
                string sizePart = parts[1].Trim();
                string numberPart = Regex.Replace(sizePart, @"^(?<number>\d+(\.\d+)?)[a-zA-Z]*$", "${number}");
                if (sizePart.EndsWith("pt", StringComparison.OrdinalIgnoreCase)) fontUnit = GraphicsUnit.Point;
                if (sizePart.EndsWith("px", StringComparison.OrdinalIgnoreCase)) fontUnit = GraphicsUnit.Pixel;
                fontSize = float.Parse(numberPart, CultureInfo.InvariantCulture);
            }

            if (parts.Length > 2)
            {
                string stylePart = parts[2].Trim();
                if (stylePart.StartsWith("style=", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (stylePart.IndexOf("bold", StringComparison.InvariantCultureIgnoreCase) != -1) fontStyle |= FontStyle.Bold;
                    if (stylePart.IndexOf("italic", StringComparison.InvariantCultureIgnoreCase) != -1) fontStyle |= FontStyle.Bold;
                    if (stylePart.IndexOf("underline", StringComparison.InvariantCultureIgnoreCase) != -1) fontStyle |= FontStyle.Bold;
                    if (stylePart.IndexOf("strikeout", StringComparison.InvariantCultureIgnoreCase) != -1) fontStyle |= FontStyle.Bold;
                }
            }

            return new Font(fontName, fontSize, fontStyle, fontUnit);
        }

        private void ThrowError(string baseNodeName, string nodeName)
        {
            ThrowError(string.Format("Node not found: {0}.{1}", baseNodeName, nodeName));
        }

        private void ThrowError(string message)
        {
            errorList.Add(message);
        }

        private byte[] RemakeFR3(byte[] data)
        {
            StringBuilder sb = new StringBuilder();

            int last = 0;
            for (int index = 0; index < data.Length; index++)
            {
                byte b = data[index];
                if (b == '<') last = index + 1;
                if (b == '>')
                {
                    sb.Append(ParseOneLine(data, last, index - last) + "\r\n");
                }
            }

            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private string ParseOneLine(byte[] data, int start, int len)
        {
            int dataLen = start + len;
            var name = new StringBuilder();
            int index = start;
            while (index < dataLen && data[index] != ' ')
            {
                name.Append((char)data[index]);
                index++;
            }

            while (index < dataLen && data[index] == ' ') index++;

            List<KeyValuePair<string, string>> attrs = new List<KeyValuePair<string, string>>();
            while (index < dataLen)
            {
                var attrName = new StringBuilder();
                while (index < dataLen && (data[index] != ' ' && data[index] != '='))
                {
                    attrName.Append((char)data[index]);
                    index++;
                }

                while (index < dataLen && data[index] == ' ') index++;

                string attrValue = null;
                if (data[index] == '=')
                {
                    index++;
                    while (index < dataLen && data[index] == ' ') index++;
                    if (data[index] != '\"') throw new Exception("No quotes!");
                    index++;
                    int last = index;
                    while (index < dataLen && data[index] != '\"') index++;

                    attrValue = GetStringEncoding(data, last, index - last, attrName.ToString());
                    index++;
                    while (index < dataLen && data[index] == ' ') index++;
                }

                attrs.Add(new KeyValuePair<string, string>(attrName.ToString(), attrValue));
            }

            StringBuilder sb = new StringBuilder();
            sb.Append("<");
            sb.Append(name + " ");
            foreach (KeyValuePair<string, string> attr in attrs)
            {
                sb.Append(attr.Key);
                if (attr.Value != null)
                {
                    sb.Append("=\"" + attr.Value + "\"");
                }
                sb.Append(" ");
            }            

            string st = sb.ToString().Trim() + ">";

            return st;
        }

        private string GetStringEncoding(byte[] data, int start, int len, string attrName)
        {
            string st = null;
            if (attrName == "Text")
            {
                st = Encoding.UTF8.GetString(data, start, len);
            }
            else
            {
                st = Encoding.GetEncoding(1251).GetString(data, start, len);
            }
            return st;
        }
        #endregion

        #region Methods.Import
        public static StiImportResult Import(byte[] bytes)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = StiCultureInfo.GetEN(false);

                var errorList = new List<string>();
                var helper = new StiFastReportsHelper();
                var report = helper.Convert(bytes, errorList);

                return new StiImportResult(report, errorList);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }
        #endregion
    }
}
