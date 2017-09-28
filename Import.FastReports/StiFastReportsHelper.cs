#region Copyright (C) 2003-2017 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports  											}
{																	}
{	Copyright (C) 2003-2017 Stimulsoft     							}
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
#endregion Copyright (C) 2003-2017 Stimulsoft

using System;
using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Net.Mime;
using System.Text;
using System.Threading;
using System.Xml;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Units;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;

namespace Stimulsoft.Report.Import
{
    public class StiFastReportsHelper
    {
        #region Methods
        public StiReport Convert(byte[] bytes)
        {

            StiReport report = new StiReport();
            report.Pages.Clear();

            XmlDocument doc = new XmlDocument();
            using (var stream = new MemoryStream(bytes))
            {
                doc.Load(stream);
            }

            XmlNodeList documentList = doc.GetElementsByTagName("Report");
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
                            ReadPage(elementNode, report);
                            break;
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
                        ReadTextObject(elementNode, band);
                        break;

                    case "PictureObject":
                        ReadPictureObject(elementNode, band);
                        break;

                    case "CheckBoxObject":
                        ReadCheckBoxObject(elementNode, band);
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
                }
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

            ReadBand(node, band);

            band.MasterComponent = masterBand;

            string dataSource = ReadString(node, "DataSource", band.Name);
            band.DataSourceName = dataSource;
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
            band.Condition.Value = condition;
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

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "ReportTitleBand":
                        ReadReportTitleBand(elementNode, page);
                        break;

                    case "ReportSummaryBand":
                        ReadReportSummaryBand(elementNode, page);
                        break;

                    case "DataBand":
                        ReadDataBand(elementNode, page, null);
                        break;

                    case "GroupHeaderBand":
                        ReadGroupHeaderBand(elementNode, page);
                        break;

                    case "PageHeaderBand":
                        ReadPageHeaderBand(elementNode, page);
                        break;

                    case "PageFooterBand":
                        ReadPageFooterBand(elementNode, page);
                        break;

                    case "ColumnHeaderBand":
                        ReadColumnHeaderBand(elementNode, page);
                        break;

                    case "ColumnFooterBand":
                        ReadColumnFooterBand(elementNode, page);
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

        private void ReadTextObject(XmlNode node, StiBand band)
        {
            StiText text = new StiText();
            band.Components.Add(text);

            ReadComp(node, text);

            text.Text.Value = ReadString(node, "Text", "");
            /*if (text.Text.Value.StartsWith("[") && text.Text.Value.EndsWith("]"))
            {
                text.Text.Value = "{" + text.Text.Value.Substring(1, text.Text.Value.Length - 2) + "}";
            }*/

            text.Angle = -(float)ReadDouble(node, "Angle", text, 0d);

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
            switch (ReadString(node, "HorzAlign", ""))
            {
                case "Center":
                    text.HorAlignment = StiTextHorAlignment.Center;
                    break;

                case "Right":
                    text.HorAlignment = StiTextHorAlignment.Right;
                    break;

                case "Justify":
                    text.HorAlignment = StiTextHorAlignment.Width;
                    break;

                default:
                    text.HorAlignment = StiTextHorAlignment.Left;
                    break;
            }
            #endregion

            #region VertAlign
            switch (ReadString(node, "VertAlign", ""))
            {
                case "Center":
                    text.VertAlignment = StiVertAlignment.Center;
                    break;

                case "Bottom":
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
                    FontConverter fontConverter = new FontConverter();
                    text.Font = fontConverter.ConvertFromString(attrName.Value) as Font;
                }
            }
            catch
            {
            }
            #endregion
        }

        private void ReadPictureObject(XmlNode node, StiBand band)
        {
            StiImage image = new StiImage();
            band.Components.Add(image);

            ReadComp(node, image);

            image.DataColumn = ReadString(node, "DataColumn", "");
        }

        private void ReadCheckBoxObject(XmlNode node, StiBand band)
        {
            StiCheckBox checkBox = new StiCheckBox();
            band.Components.Add(checkBox);

            ReadComp(node, checkBox);

            checkBox.Checked.Value = "{" + ReadString(node, "DataColumn", "") + "}";
        }

        private void ReadLineObject(XmlNode node, StiBand band)
        {
            StiHorizontalLinePrimitive line = new StiHorizontalLinePrimitive();
            band.Components.Add(line);

            ReadComp(node, line);
        }

        private void ReadSubreportObject(XmlNode node, StiContainer parent)
        {
            StiSubReport subReport = new StiSubReport();
            parent.Components.Add(subReport);

            ReadComp(node, subReport);
        }
        #endregion

        #region Methods.Import
        public static StiImportResult Import(byte[] bytes)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US", false);

                var helper = new StiFastReportsHelper();
                var report = helper.Convert(bytes);

                return new StiImportResult(report, null);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }
        #endregion
    }
}
