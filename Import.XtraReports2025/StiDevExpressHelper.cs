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
using System.Text;
using System.Threading;
using System.Xml;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report.BarCodes;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.CrossTab;
using Stimulsoft.Base;
using Stimulsoft.Report.Components.ShapeTypes;
using Stimulsoft.Report.Chart;

namespace Stimulsoft.Report.Import
{
    public class StiDevExpressHelper
    {
        #region Fields
        private List<string> errorList = null;
        private ArrayList fields = new ArrayList();
        private string currentDataSourceName = "ds";
        private string mainDataSourceName = "ds";
        private List<string> summaryTypeCollection = new List<string>();
        private List<string> clonnedDataSourceNames = new List<string>();
        private Font defaultFont = new Font("Arial", 8);
        private string reportUnit = "HundredthsOfAnInch";
        private Hashtable dataSourceReferenceCollection = new Hashtable();
        private string dataSourceReference = "";
        private int bandCountForDifferentName = 1;
        private double posX = 0d;
        private Hashtable queryCollection = new Hashtable();
        private Hashtable dataSourseCorrectNameCollection = new Hashtable();
        private double allRowCellsWeight = 0d;
        private Hashtable styles = new Hashtable();
        private List<StiOverlayBand> topMarginBandsCollection = new List<StiOverlayBand>();
        private List<StiReportTitleBand> reportHeaderBandsCollection = new List<StiReportTitleBand>();
        private List<StiPageHeaderBand> pageHeaderBandsCollection = new List<StiPageHeaderBand>();
        private List<StiDataBand> detailBandsCollection = new List<StiDataBand>();
        private List<StiPanel> detailReportBandsCollection = new List<StiPanel>();
        private List<StiOverlayBand> bottomMarginBandsCollection = new List<StiOverlayBand>();
        private List<StiGroupHeaderBand> groupHeaderBandsCollection = new List<StiGroupHeaderBand>();
        private List<StiGroupFooterBand> groupFooterBandsCollection = new List<StiGroupFooterBand>();
        private List<StiReportSummaryBand> reportFooterBandsCollection = new List<StiReportSummaryBand>();
        private List<StiPageFooterBand> pageFooterBandsCollection = new List<StiPageFooterBand>();
        private List<StiDataBand> verticalDetailBandsCollection = new List<StiDataBand>();
        private List<StiDataBand> detailReportDetailBandsCollection = new List<StiDataBand>();
        private List<StiReportTitleBand> detailReportReportHeaderBandsCollection = new List<StiReportTitleBand>();
        private List<StiGroupHeaderBand> detailReportGroupHeaderBandsCollection = new List<StiGroupHeaderBand>();
        private List<StiPanel> detailReportDetailReportBandsCollection = new List<StiPanel>();
        private List<StiDataBand> detailReportVerticalDetailBandsCollection = new List<StiDataBand>();
        private List<StiGroupFooterBand> detailReportGroupFooterBandsCollection = new List<StiGroupFooterBand>();
        private List<StiReportSummaryBand> detailReportReportFooterBandsCollection = new List<StiReportSummaryBand>();
        #endregion

        #region Root node
        public void ProcessFile(byte[] bytes, StiReport report, List<string> errorList)
        {
            var doc = new XmlDocument();

            using (var stream = new MemoryStream(bytes))
            {
                doc.Load(stream);
            }
            var rootNode = doc.DocumentElement;

            this.errorList = errorList;
            fields.Clear();
            report.ReportUnit = StiReportUnitType.HundredthsOfInch;
            StiPage page = report.Pages[0];
            page.UnlimitedBreakable = false;
            page.TitleBeforeHeader = true;

            var pageWidth = "827";
            var pageHeight = "1169";
            var margins = "100, 100, 100, 100";
            var font = "";
            var dataMember = "";

            foreach (XmlNode node in rootNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        report.ReportName = node.Value;
                        break;

                    case "PageWidth":
                        pageWidth = node.Value;
                        break;

                    case "PageHeight":
                        pageHeight = node.Value;
                        break;

                    case "ReportUnit":
                        reportUnit = node.Value;
                        break;

                    case "Margins":
                        margins = node.Value;
                        break;

                    case "Font":
                        font = node.Value;
                        break;

                    case "DataSource":
                        dataSourceReference = node.Value.Replace("#Ref-", "");
                        break;

                    case "DataMember":
                        dataMember = node.Value;
                        break;
                }
            }

            if (!string.IsNullOrEmpty(margins))
            {
                var marginsArray = margins.Split(' ');

                if (marginsArray.Length == 4)
                {
                    page.Margins = new StiMargins(
                    ToHi(reportUnit, marginsArray[0].TrimEnd(',')),
                    ToHi(reportUnit, marginsArray[1].TrimEnd(',')),
                    ToHi(reportUnit, marginsArray[2].TrimEnd(',')),
                    ToHi(reportUnit, marginsArray[3].TrimEnd(',')));
                }
            }

            if (!string.IsNullOrEmpty(pageWidth))
                page.PageWidth = ToHi(reportUnit, pageWidth);

            if (!string.IsNullOrEmpty(pageHeight))
                page.PageHeight = ToHi(reportUnit, pageHeight);


            if (!string.IsNullOrEmpty(font))
                defaultFont = ParseFont(font);

            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ComponentStorage":
                        ProcessComponentStorage(node, report, page);
                        break;

                    case "StyleSheet":
                        ProcessStyleSheet(node, report, page);
                        report.Styles.Clear();

                        var keys = styles.Keys;

                        foreach (var key in keys)
                        {
                            var style = styles[key];

                            report.Styles.Add((StiStyle)style);
                        }
                        break;
                }
            }

            if (dataSourceReferenceCollection.Count > 0 && !string.IsNullOrEmpty(dataSourceReference) && string.IsNullOrEmpty(dataMember))
                mainDataSourceName = (string)dataSourceReferenceCollection[dataSourceReference];

            else if (dataSourceReferenceCollection.Count > 0 && !string.IsNullOrEmpty(dataSourceReference)
                && !string.IsNullOrEmpty(dataMember))
                mainDataSourceName = dataMember;

            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Bands":
                        ProcessBands(node, report, page);
                        break;
                }
            }

            foreach (StiComponent comp in report.GetComponents())
            {
                comp.Page = page;
            }

            page.DockToContainer();
            page.SortByPriority();
        }
        #endregion

        #region Items
        private void ProcessBands(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    ProcessBandsItem(node, report, container);

                    itemIndex++;
                }
            }

            if (container.GetType() == typeof(StiPage))
            {
                foreach (var band in topMarginBandsCollection)
                    container.Components.Add(band);

                foreach (var band in reportHeaderBandsCollection)
                    container.Components.Add(band);

                foreach (var band in pageHeaderBandsCollection)
                    container.Components.Add(band);

                foreach (var band in groupHeaderBandsCollection)
                    container.Components.Add(band);

                foreach (var band in detailBandsCollection)
                    container.Components.Add(band);

                foreach (var panel in detailReportBandsCollection)
                {
                    foreach (StiComponent item in panel.Components)
                    {
                        container.Components.Add(item);
                    }
                }

                foreach (var band in verticalDetailBandsCollection)
                    container.Components.Add(band);

                foreach (var band in groupFooterBandsCollection)
                    container.Components.Add(band);

                foreach (var band in pageFooterBandsCollection)
                    container.Components.Add(band);

                foreach (var band in reportFooterBandsCollection)
                    container.Components.Add(band);

                foreach (var band in bottomMarginBandsCollection)
                    container.Components.Add(band);
            }

            if (container.GetType() == typeof(StiPanel))
            {
                foreach (var band in detailReportReportHeaderBandsCollection)
                    container.Components.Add(band);

                foreach (var band in detailReportGroupHeaderBandsCollection)
                    container.Components.Add(band);

                foreach (var band in detailReportDetailReportBandsCollection)
                    container.Components.Add(band);

                foreach (var band in detailReportVerticalDetailBandsCollection)
                    container.Components.Add(band);

                foreach (var band in detailReportGroupFooterBandsCollection)
                    container.Components.Add(band);

                foreach (var band in detailReportReportFooterBandsCollection)
                    container.Components.Add(band);

                foreach (var band in detailReportDetailBandsCollection)
                    container.Components.Add(band);

                detailReportReportHeaderBandsCollection.Clear();
                detailReportGroupHeaderBandsCollection.Clear();
                detailReportDetailReportBandsCollection.Clear();
                detailReportVerticalDetailBandsCollection.Clear();
                detailReportGroupFooterBandsCollection.Clear();
                detailReportReportFooterBandsCollection.Clear();
                detailReportDetailBandsCollection.Clear();
            }
        }

        private void ProcessStyleSheet(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    var component = new StiText();

                    ProcessStyle(node, component);

                    var style = new StiStyle
                    {
                        Name = component.Name,
                        Brush = component.Brush,
                        Font = component.Font,
                        TextBrush = component.TextBrush,
                        HorAlignment = component.HorAlignment,
                        VertAlignment = component.VertAlignment,
                        AllowUseVertAlignment = true,
                        AllowUseHorAlignment = true,
                        Border = component.Border
                    };

                    if (!string.IsNullOrEmpty(style.Name))
                        styles.Add(style.Name, style);

                    itemIndex++;
                }
            }
        }

        private void ProcessStyle(XmlNode baseNode, StiText component)
        {
            var borderSide = StiBorderSides.All;
            var borderColor = Color.Black;
            var borderWidth = 1f;
            var borderDashStyle = StiPenStyle.Solid;
            var newFont = new Font("Arial", 8);

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        component.Name = node.Value;
                        break;

                    case "Visible":
                        if (node.Value == "True") component.Enabled = true;
                        if (node.Value == "False") component.Enabled = false;
                        break;

                    case "BackColor":
                        component.Brush = new StiSolidBrush(ParseStyleColor(node.Value));
                        break;

                    case "ForeColor":
                        component.TextBrush = new StiSolidBrush(ParseStyleColor(node.Value));
                        break;

                    case "TextAlignment":
                        TextHorizontalAlignment(component, node.Value);
                        VerticalAlignment(component, node.Value);
                        break;

                    case "Font":
                        newFont = ParseFont(node.Value);
                        break;

                    case "Borders":
                        borderSide = BorderSide(node.Value);
                        break;

                    case "Sides":
                        borderSide = BorderSide(node.Value);
                        break;

                    case "BorderColor":
                        borderColor = ParseStyleColor(node.Value);
                        break;

                    case "BorderWidth":
                        borderWidth = float.Parse(node.Value);
                        break;

                    case "BorderDashStyle":
                        borderDashStyle = ParseBorderStyle(node.Value);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            component.Font = newFont;
            component.Border = new StiBorder(borderSide, borderColor, borderWidth, borderDashStyle);
        }

        private void ProcessComponentStorage(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    ProcessComponentStorageItem(node, report, container);
                    itemIndex++;
                }
            }
        }

        private void ProcessBandsItem(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var bandItemType = "";
            var bandItemName = "";
            var bandItemHeight = "100";
            var bandItemLayout = "Across";
            var backColor = Color.Transparent;
            var foreColor = Color.Black;
            var borderSide = StiBorderSides.All;
            var borderColor = Color.Black;
            var borderWidth = 1f;
            var borderDashStyle = StiPenStyle.Solid;
            var isBorderInside = false;
            var font = defaultFont;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "ControlType":
                        bandItemType = node.Value;
                        break;

                    case "Name":
                        bandItemName = node.Value;
                        break;

                    case "HeightF":
                        bandItemHeight = node.Value;
                        break;

                    case "BackColor":
                        backColor = ParseStyleColor(node.Value);
                        break;

                    case "ForeColor":
                        foreColor = ParseStyleColor(node.Value);
                        break;

                    case "Borders":
                        borderSide = BorderSide(node.Value);
                        isBorderInside = true;
                        break;

                    case "BorderColor":
                        borderColor = ParseStyleColor(node.Value);
                        isBorderInside = true;
                        break;

                    case "BorderWidth":
                        borderWidth = float.Parse(node.Value);
                        isBorderInside = true;
                        break;

                    case "BorderDashStyle":
                        borderDashStyle = ParseBorderStyle(node.Value);
                        isBorderInside = true;
                        break;

                    case "Font":
                        font = ParseFont(node.Value);
                        break;
                }
            }

            var border = new StiBorder(borderSide, borderColor, borderWidth, borderDashStyle);

            var panel = new StiPanel();
            panel.CanGrow = true;
            panel.CanShrink = true;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Controls":

                        if (isBorderInside)
                            ProcessControls(node, panel, foreColor, report, border, font);

                        else
                            ProcessControls(node, panel, foreColor, report, null, font);

                        break;

                    case "MultiColumn":
                        foreach (XmlNode nodee in node.Attributes)
                        {
                            switch (nodee.Name)
                            {
                                case "Layout":
                                    bandItemLayout = nodee.Value;
                                    break;

                                case "Mode":

                                    break;
                            }
                        }
                        break;
                }
            }

            if (bandItemType == "TopMarginBand")
            {
                var topMarginBand = new StiOverlayBand();
                topMarginBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                topMarginBand.Height = ToHi(reportUnit, bandItemHeight);
                topMarginBand.MaxHeight = ToHi(reportUnit, bandItemHeight);
                topMarginBand.MinHeight = ToHi(reportUnit, bandItemHeight);
                topMarginBand.Brush = new StiSolidBrush(backColor);
                topMarginBand.VertAlignment = StiVertAlignment.Top;

                if (container is StiPage)
                {
                    var page = container as StiPage;

                    if (page.Margins.Top < topMarginBand.Height)
                        page.Margins.Top = topMarginBand.Height;
                }

                foreach (var comp in panel.Components)
                {
                    topMarginBand.Components.Add(comp as StiComponent);
                }

                if (container.GetType() == typeof(StiPage) && topMarginBand.Components.Count > 0)
                    topMarginBandsCollection.Add(topMarginBand);
            }
            else if (bandItemType == "ReportHeaderBand")
            {
                var reportTitleBand = new StiReportTitleBand();
                reportTitleBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                reportTitleBand.Height = ToHi(reportUnit, bandItemHeight);
                reportTitleBand.Brush = new StiSolidBrush(backColor);
                reportTitleBand.CanShrink = false;

                foreach (var comp in panel.Components)
                {
                    reportTitleBand.Components.Add(comp as StiComponent);
                }

                if (container.GetType() == typeof(StiPage))
                    reportHeaderBandsCollection.Add(reportTitleBand);

                if (container.GetType() == typeof(StiPanel))
                    detailReportReportHeaderBandsCollection.Add(reportTitleBand);
            }
            else if (bandItemType == "PageHeaderBand")
            {
                var pageHeaderBand = new StiPageHeaderBand();
                pageHeaderBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                pageHeaderBand.Height = ToHi(reportUnit, bandItemHeight);
                pageHeaderBand.Brush = new StiSolidBrush(backColor);

                foreach (var comp in panel.Components)
                {
                    pageHeaderBand.Components.Add(comp as StiComponent);
                }

                if (container.GetType() == typeof(StiPage))
                    pageHeaderBandsCollection.Add(pageHeaderBand);
            }
            else if (bandItemType == "DetailBand")
            {
                var dataBand = new StiDataBand();
                dataBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                dataBand.Height = ToHi(reportUnit, bandItemHeight);

                if (container.GetType() == typeof(StiPage))
                    dataBand.DataSourceName = mainDataSourceName;

                dataBand.Brush = new StiSolidBrush(backColor);

                foreach (StiComponent comp in panel.Components)
                {
                    if (comp.ClientRectangle.Bottom > dataBand.Height)
                        dataBand.Height = comp.ClientRectangle.Bottom;

                    dataBand.Components.Add(comp);
                }

                if (container.GetType() == typeof(StiPage))
                    detailBandsCollection.Add(dataBand);

                if (container.GetType() == typeof(StiPanel))
                {
                    if (detailBandsCollection.Count > 0)
                        dataBand.MasterComponent = detailBandsCollection[0];

                    detailReportDetailBandsCollection.Add(dataBand);
                }
            }
            else if (bandItemType == "DetailReportBand")
            {
                var panelDetail = new StiPanel();
                panelDetail.CanGrow = true;
                panelDetail.CanShrink = true;
                panelDetail.Name = $"{bandItemName}{bandCountForDifferentName}Panel";

                foreach (XmlNode node in baseNode.ChildNodes)
                {
                    switch (node.Name)
                    {
                        case "Bands":
                            ProcessBands(node, report, panelDetail);
                            break;
                    }
                }

                var allPanelBandsHeight = 0d;
                var maxPanelBandsWidth = 0d;

                foreach (StiBand band in panelDetail.Components)
                {
                    allPanelBandsHeight += band.ClientRectangle.Height;

                    foreach (StiComponent comp in band.Components)
                    {
                        if (maxPanelBandsWidth < comp.ClientRectangle.Width)
                            maxPanelBandsWidth = comp.ClientRectangle.Right;
                    }
                }

                panelDetail.Height = allPanelBandsHeight;
                panelDetail.Width = container.Width;

                if (container.GetType() == typeof(StiPage))
                    detailReportBandsCollection.Add(panelDetail);

                if (container.GetType() == typeof(StiPanel))
                    detailReportDetailReportBandsCollection.Add(panelDetail);
            }
            else if (bandItemType == "BottomMarginBand")
            {
                var bottomMargin = new StiOverlayBand();
                bottomMargin.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                bottomMargin.Height = ToHi(reportUnit, bandItemHeight);
                bottomMargin.MaxHeight = ToHi(reportUnit, bandItemHeight);
                bottomMargin.MinHeight = ToHi(reportUnit, bandItemHeight);
                bottomMargin.Brush = new StiSolidBrush(backColor);
                bottomMargin.VertAlignment = StiVertAlignment.Bottom;

                if (container is StiPage)
                {
                    var page = container as StiPage;

                    if (page.Margins.Bottom < bottomMargin.Height)
                        page.Margins.Bottom = bottomMargin.Height;
                }

                foreach (StiComponent comp in panel.Components)
                {
                    comp.Linked = true;
                    bottomMargin.Components.Add(comp);
                }

                if (container.GetType() == typeof(StiPage) && bottomMargin.Components.Count > 0)
                    bottomMarginBandsCollection.Add(bottomMargin);
            }
            else if (bandItemType == "GroupHeaderBand")
            {
                var groupHeaderBand = new StiGroupHeaderBand();
                groupHeaderBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                groupHeaderBand.Height = ToHi(reportUnit, bandItemHeight);
                groupHeaderBand.Brush = new StiSolidBrush(backColor);

                foreach (var comp in panel.Components)
                    groupHeaderBand.Components.Add(comp as StiComponent);

                if (container.GetType() == typeof(StiPage))
                    groupHeaderBandsCollection.Add(groupHeaderBand);

                if (container.GetType() == typeof(StiPanel))
                    detailReportGroupHeaderBandsCollection.Add(groupHeaderBand);
            }
            else if (bandItemType == "GroupFooterBand")
            {
                var groupFooterBand = new StiGroupFooterBand();
                groupFooterBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                groupFooterBand.Height = ToHi(reportUnit, bandItemHeight);
                groupFooterBand.Brush = new StiSolidBrush(backColor);

                foreach (var comp in panel.Components)
                    groupFooterBand.Components.Add(comp as StiComponent);

                if (container.GetType() == typeof(StiPage))
                    groupFooterBandsCollection.Add(groupFooterBand);

                if (container.GetType() == typeof(StiPanel))
                    detailReportGroupFooterBandsCollection.Add(groupFooterBand);
            }
            else if (bandItemType == "ReportFooterBand")
            {
                var reportSummaryBand = new StiReportSummaryBand();
                reportSummaryBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                reportSummaryBand.Height = ToHi(reportUnit, bandItemHeight);
                reportSummaryBand.Brush = new StiSolidBrush(backColor);

                foreach (var comp in panel.Components)
                    reportSummaryBand.Components.Add(comp as StiComponent);

                if (container.GetType() == typeof(StiPage))
                    reportFooterBandsCollection.Add(reportSummaryBand);

                if (container.GetType() == typeof(StiPanel))
                    detailReportReportFooterBandsCollection.Add(reportSummaryBand);
            }
            else if (bandItemType == "PageFooterBand")
            {
                var pageFooterBand = new StiPageFooterBand();
                pageFooterBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                pageFooterBand.Height = ToHi(reportUnit, bandItemHeight);
                pageFooterBand.Brush = new StiSolidBrush(backColor);

                foreach (var comp in panel.Components)
                    pageFooterBand.Components.Add(comp as StiComponent);

                if (container.GetType() == typeof(StiPage))
                    pageFooterBandsCollection.Add(pageFooterBand);
            }
            else if (bandItemType == "VerticalDetailBand")
            {
                var dataBand = new StiDataBand();
                dataBand.Name = $"{bandItemName}{bandCountForDifferentName}";
                bandCountForDifferentName++;
                dataBand.Height = ToHi(reportUnit, bandItemHeight);
                dataBand.DataSourceName = mainDataSourceName;
                dataBand.Brush = new StiSolidBrush(backColor);

                foreach (var comp in panel.Components)
                {
                    var crossDataBand = new StiCrossDataBand();
                    crossDataBand.Name = $"cdb{dataBand.Name}";
                    crossDataBand.DataSourceName = mainDataSourceName;
                    crossDataBand.Components.Add(comp as StiComponent);

                    dataBand.Components.Add(crossDataBand);
                }

                if (container.GetType() == typeof(StiPage))
                    verticalDetailBandsCollection.Add(dataBand);

                if (container.GetType() == typeof(StiPanel))
                    detailReportVerticalDetailBandsCollection.Add(dataBand);
            }
            else if (bandItemType == "VerticalHeaderBand")
            {

            }
            else if (bandItemType == "VerticalTotalBand")
            {

            }
        }

        private void ProcessComponentStorageItem(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var dataType = "";
            var dataName = "";
            var base64 = "";
            var reff = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "ObjectType":
                        dataType = node.Value;
                        break;

                    case "Name":
                        dataName = node.Value;
                        break;

                    case "Base64":
                        base64 = node.Value;
                        break;

                    case "Ref":
                        reff = node.Value;
                        break;
                }
            }

            if (!string.IsNullOrEmpty(dataType))
            {
                var newDataType = dataType.Split(',')[0];
                dataType = newDataType;
            }

            if (dataType == "DevExpress.DataAccess.Excel.ExcelDataSource")
            {
                if (!string.IsNullOrEmpty(base64))
                {
                    var dataBytes = Convert.FromBase64String(base64);
                    var dataXML = new XmlDocument();

                    using (var stream = new MemoryStream(dataBytes))
                    {
                        dataXML.Load(stream);
                    }

                    var rootNode = dataXML.DocumentElement;

                    ProcessCsvDataSource(rootNode, report, reff);
                }
            }

            if (dataType == "DevExpress.DataAccess.Sql.SqlDataSource")
            {
                if (!string.IsNullOrEmpty(base64))
                {
                    var dataBytes = Convert.FromBase64String(base64);
                    var dataXML = new XmlDocument();

                    using (var stream = new MemoryStream(dataBytes))
                    {
                        dataXML.Load(stream);
                    }

                    var rootNode = dataXML.DocumentElement;

                    ProcessSqlDataSource(rootNode, report, reff);
                }
            }
        }

        private void ProcessCsvDataSource(XmlElement rootNode, StiReport report, string reference)
        {
            var name = "";
            var path = "";
            var valueSeparator = "";

            foreach (XmlNode node in rootNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        break;

                    case "FileName":
                        path = node.Value;
                        break;

                    case "ValueSeparator":
                        valueSeparator = node.Value;
                        break;
                }
            }

            if (!string.IsNullOrEmpty(reference) && !string.IsNullOrEmpty(name))
                dataSourceReferenceCollection.Add(reference, name);

            StiCsvDatabase dataBase = new StiCsvDatabase();
            StiCsvSource dataSource = new StiCsvSource();

            dataSource.Name = name;
            dataSource.Alias = name;

            if (!string.IsNullOrEmpty(valueSeparator))
                dataSource.Separator = valueSeparator;

            dataSource.NameInSource = "db" + name;
            dataBase.Name = "db" + name;
            dataBase.Alias = "db" + name;
            currentDataSourceName = name;

            dataBase.PathData = path;
            dataSource.NameInSource = dataBase.PathData;

            if (dataSource.NameInSource.StartsWith("file:///"))
                dataSource.NameInSource = dataSource.NameInSource.Substring(8);
            if (dataSource.NameInSource.StartsWith("file://"))
                dataSource.NameInSource = dataSource.NameInSource.Substring(7);

            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Schema":
                        ProcessSchema(node, dataSource);
                        break;
                }
            }

            report.Dictionary.DataSources.Add(dataSource);

            var dataSource2 = dataSource.Clone() as StiCsvSource;
            dataSource2.Name += "2";
            clonnedDataSourceNames.Add(dataSource2.Name);
            dataSource2.Alias += "2";
            report.Dictionary.DataSources.Add(dataSource2);
        }

        private void ProcessSqlDataSource(XmlNode baseNode, StiReport report, string reference)
        {
            var name = "";
            StiSqlDatabase dataBase = new StiSqlDatabase();
            dataBase.Name = "db";
            dataBase.Alias = "db";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        dataBase.Name = node.Value;
                        dataBase.Alias = node.Value;
                        currentDataSourceName = node.Value;
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            if (!string.IsNullOrEmpty(reference) && !string.IsNullOrEmpty(name))
                dataSourceReferenceCollection.Add(reference, name);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Query":
                        ProcessQuery(node, queryCollection);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Connection":
                        dataBase.ConnectionString = ProcessConnection(node);
                        break;

                    case "ResultSchema":
                        ResultSchema(node, report);
                        break;
                }
            }

            report.Dictionary.Databases.Add(dataBase);
        }

        private void ProcessQuery(XmlNode baseNode, Hashtable queryCollection)
        {
            var type = "SelectQuery";
            var name = "";
            var tableName = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Type":
                        type = node.Value;
                        break;

                    case "Name":
                        name = node.Value;
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
                    case "Tables":
                        foreach (XmlNode nodee in node.ChildNodes)
                        {
                            switch (nodee.Name)
                            {
                                case "Table":
                                    foreach (XmlNode nodeee in nodee.Attributes)
                                    {
                                        switch (nodeee.Name)
                                        {
                                            case "Name":
                                                tableName = nodeee.Value;
                                                break;

                                            default:
                                                ThrowError(baseNode.Name, "#" + nodee.Name);
                                                break;
                                        }
                                    }
                                    break;

                                default:
                                    ThrowError(baseNode.Name, "#" + node.Name);
                                    break;
                            }
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            var comand = "";

            if (type == "SelectQuery")
                comand = $"select * from {tableName}";

            if (type == "SelectQuery")
            {
                queryCollection.Add(name, comand);
                dataSourseCorrectNameCollection.Add(name, tableName);
            }
            else if (type == "StoredProcQuery")
            {

            }
        }

        private void ProcessSqlParameters(XmlNode baseNode, StiReport report, StiSqlSource dataSource)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "SqlDataSourceParameter":
                        ProcessSqlDataSourceParameter(node, report, dataSource);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessSqlDataSourceParameter(XmlNode baseNode, StiReport report, StiSqlSource dataSource)
        {
            StiDataParameter dp = new StiDataParameter();

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        dp.Name = node.Value;
                        break;
                    case "DbType":
                        if (node.Value == "String") dp.Type = (int)System.Data.SqlDbType.Variant;
                        if (node.Value == "Int32") dp.Type = (int)System.Data.SqlDbType.Int;
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
                    case "Value":
                        string value = node.ChildNodes[0].Value;
                        if ((node.ChildNodes[0].Name == "String") && (node.ChildNodes[0].ChildNodes.Count == 1) && (node.ChildNodes[0].ChildNodes[0].Name == "#text"))
                        {
                            value = node.ChildNodes[0].ChildNodes[0].Value;
                        }
                        if (!string.IsNullOrEmpty(value))
                        {
                            dp.Value = value;
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            dataSource.Parameters.Add(dp);
        }

        private string ProcessConnection(XmlNode baseNode)
        {
            var name = "";
            var providerKey = "";
            var connectionString = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        break;
                    case "ProviderKey":
                        providerKey = node.Value;
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
                    case "Parameters":
                        connectionString = ProcessSqlConnectionParameters(node);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return connectionString;
        }

        private string ProcessSqlConnectionParameters(XmlNode baseNode)
        {
            var server = "";
            var database = "";
            var useIntegratedSecurity = "true";
            var readOnly = "1";
            var generateConnectionHelper = "false";
            var userID = "";
            var password = "";

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Parameter":
                        var parameterName = "";
                        var parameterValue = "";

                        foreach (XmlNode nodee in node.Attributes)
                        {
                            switch (nodee.Name)
                            {
                                case "Name":
                                    parameterName = nodee.Value;
                                    break;

                                case "Value":
                                    parameterValue = nodee.Value;
                                    break;

                                default:
                                    ThrowError(baseNode.Name, "#" + nodee.Name);
                                    break;
                            }
                        }

                        switch (parameterName)
                        {
                            case "server":
                                server = parameterValue;
                                break;
                            case "database":
                                database = parameterValue;
                                break;
                            case "useIntegratedSecurity":
                                useIntegratedSecurity = parameterValue;
                                break;
                            case "read only":
                                readOnly = node.Value;
                                break;
                            case "generateConnectionHelper":
                                generateConnectionHelper = parameterValue;
                                break;
                            case "userId":
                                userID = parameterValue;
                                break;
                            case "password":
                                password = parameterValue;
                                break;
                        }

                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            var connectionString = $"Server={server};Initial Catalog={database};Integrated Security={useIntegratedSecurity};" +
                $"User ID={userID};Password={password};";

            return connectionString;
        }

        private void ResultSchema(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSet":
                        ResultDataSet(node, report);
                        break;
                }
            }
        }

        private void ResultDataSet(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "View":
                        ResultView(node, report);
                        break;
                }
            }
        }

        private void ResultView(XmlNode baseNode, StiReport report)
        {
            var viewName = "";
            var name = "";

            StiSqlSource dataSource = new StiSqlSource();
            dataSource.Name = "ds";
            dataSource.Alias = "ds";
            dataSource.NameInSource = "db";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        viewName = node.Value;
                        var query = (string)queryCollection[viewName];

                        if (!string.IsNullOrEmpty(query))
                        {
                            name = query.Split(' ')[query.Split(' ').Length - 1];

                            dataSource.Name = name;
                            dataSource.Alias = name;
                            dataSource.NameInSource = currentDataSourceName;
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            dataSource.SqlCommand = (string)queryCollection[viewName];

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Field":
                        var fieldName = "";
                        var type = "";

                        foreach (XmlNode nodee in node.Attributes)
                        {
                            switch (nodee.Name)
                            {
                                case "Name":
                                    fieldName = nodee.Value;
                                    break;

                                case "Type":
                                    type = $"System.{nodee.Value}";
                                    break;
                            }
                        }

                        var newColumn = new StiDataColumn(fieldName);
                        newColumn.Name = fieldName;

                        Type newType = Type.GetType(type);

                        if (newType == null)
                            newType = Type.GetType("System.Object");

                        newColumn.Type = newType;

                        dataSource.Columns.Add(newColumn);

                        break;
                }
            }

            if (!string.IsNullOrEmpty(name))
                report.Dictionary.DataSources.Add(dataSource);
        }

        private void ProcessSchema(XmlNode baseNode, StiFileDataSource dataSourse)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "FieldInfo":
                        var name = "";
                        var originalName = "";
                        var type = "";
                        var selected = "";

                        foreach (XmlNode nodee in node.Attributes)
                        {
                            switch (nodee.Name)
                            {
                                case "Name":
                                    name = nodee.Value;
                                    break;

                                case "OriginalName":
                                    originalName = nodee.Value;
                                    break;

                                case "Type":
                                    type = nodee.Value;
                                    break;

                                case "Selected":
                                    selected = nodee.Value;
                                    break;
                            }
                        }

                        var newColumn = new StiDataColumn(originalName);
                        newColumn.Name = name;

                        Type newType = Type.GetType(type);
                        newColumn.Type = newType;

                        dataSourse.Columns.Add(newColumn);

                        break;
                }
            }
        }

        private void ProcessControls(XmlNode baseNode, StiContainer container, Color foreColor, StiReport report, StiBorder border = null,
            Font font = null, string textAlignment = null)
        {
            var itemIndex = 1;
            var newFont = new Font("Arial", 8);
            var backColor = Color.Transparent;

            if (font != null)
                newFont = font;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    if (border != null)
                        ProcessItem(node, container, foreColor, backColor, newFont, report, border, textAlignment);

                    else
                        ProcessItem(node, container, foreColor, backColor, newFont, report, null, textAlignment);

                    itemIndex++;
                }
            }
        }

        private void ProcessItem(XmlNode baseNode, StiContainer container, Color foreColor, Color defaultBackColor, Font defaultFont,
            StiReport report, StiBorder border = null, string textAlign = null)
        {
            var itemType = "";
            var itemName = "";
            var multiline = "false";
            var text = "";
            var sizeF = "";
            var locationFloat = "";
            var padding = "";
            var isChecked = "false";
            var borderSide = StiBorderSides.All;
            var borderColor = Color.Black;
            var borderWidth = 1f;
            var borderDashStyle = StiPenStyle.Solid;
            var isBorderNeedChange = false;
            var imageSource = "";
            var sizing = "NormalImage";
            var serializableRtfString = "";
            var cellWidth = "1";
            var cellHeight = "1";
            var lineWidth = "1";
            var lineDirection = "Horizontal";
            var lineStyle = "Solid";
            var angle = "";
            var fillColor = Color.Transparent;
            var backColor = defaultBackColor;
            var isForeColorNeedChange = false;
            var isBackColorNeedChange = false;
            var module = "";
            var autoModule = "";
            var dataSourceRef = "";
            var dataMember = "";
            var font = defaultFont;
            var isFontNeedChange = false;
            var textAlignment = "";
            var styleName = "";
            var cellHorizontalSpacing = "2";
            var cellVerticalSpacing = "2";

            if (textAlign != null && !string.IsNullOrEmpty(textAlign))
                textAlignment = textAlign;

            if (border != null)
            {
                borderSide = border.Side;
                borderColor = border.Color;
                borderWidth = (float)border.Size;
                borderDashStyle = border.Style;
                isBorderNeedChange = true;
            }

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "ControlType":
                        itemType = node.Value;
                        break;

                    case "Name":
                        itemName = node.Value;
                        break;

                    case "Multiline":
                        multiline = node.Value;
                        break;

                    case "Text":
                        text = node.Value;
                        break;

                    case "SizeF":
                        sizeF = node.Value;
                        break;

                    case "LocationFloat":
                        locationFloat = node.Value;
                        break;

                    case "Padding":
                        padding = node.Value;
                        break;

                    case "Checked":
                        isChecked = node.Value;
                        break;

                    case "Borders":
                        borderSide = BorderSide(node.Value);
                        isBorderNeedChange = true;
                        break;

                    case "Sides":
                        borderSide = BorderSide(node.Value);
                        isBorderNeedChange = true;
                        break;

                    case "BorderColor":
                        borderColor = ParseStyleColor(node.Value);
                        isBorderNeedChange = true;
                        break;

                    case "BorderWidth":
                        borderWidth = float.Parse(node.Value);
                        isBorderNeedChange = true;
                        break;

                    case "BorderDashStyle":
                        borderDashStyle = ParseBorderStyle(node.Value);
                        isBorderNeedChange = true;
                        break;

                    case "ImageSource":
                        imageSource = node.Value;
                        break;

                    case "Sizing":
                        sizing = node.Value;
                        break;

                    case "SerializableRtfString":
                        serializableRtfString = node.Value;
                        break;

                    case "CellWidth":
                        cellWidth = node.Value;
                        break;

                    case "CellHeight":
                        cellHeight = node.Value;
                        break;

                    case "LineWidth":
                        lineWidth = node.Value;
                        break;

                    case "LineDirection":
                        lineDirection = node.Value;
                        break;

                    case "LineStyle":
                        lineStyle = node.Value;
                        break;

                    case "ForeColor":
                        foreColor = ParseStyleColor(node.Value);
                        isForeColorNeedChange = true;
                        break;

                    case "BackColor":
                        backColor = ParseStyleColor(node.Value);
                        isBackColorNeedChange = true;
                        break;

                    case "Angle":
                        angle = node.Value;
                        break;

                    case "FillColor":
                        fillColor = ParseStyleColor(node.Value);
                        break;

                    case "Module":
                        module = node.Value;
                        break;

                    case "AutoModule":
                        autoModule = node.Value;
                        break;

                    case "DataSource":
                        dataSourceRef = node.Value.Replace("#Ref-", "");
                        break;

                    case "DataMember":
                        dataMember = node.Value;
                        break;

                    case "Font":
                        font = ParseFont(node.Value);
                        isFontNeedChange = true;
                        break;

                    case "TextAlignment":
                        textAlignment = node.Value;
                        break;

                    case "StyleName":
                        styleName = node.Value;
                        break;

                    case "CellHorizontalSpacing":
                        cellHorizontalSpacing = node.Value;
                        break;

                    case "CellVerticalSpacing":
                        cellVerticalSpacing = node.Value;
                        break;
                }
            }

            if (dataSourceReferenceCollection.Count > 0 && string.IsNullOrEmpty(dataSourceReference)
                && !string.IsNullOrEmpty(dataSourceRef) && string.IsNullOrEmpty(dataMember))
                mainDataSourceName = dataSourceReferenceCollection[dataSourceRef].ToString();

            else if (dataSourceReferenceCollection.Count > 0 && string.IsNullOrEmpty(dataSourceReference)
                && !string.IsNullOrEmpty(dataSourceRef) && !string.IsNullOrEmpty(dataMember))
                mainDataSourceName = dataMember;

            border = new StiBorder(borderSide, borderColor, borderWidth, borderDashStyle);

            if (itemType == "XRLabel")
                ProcessTextBox(baseNode, itemName, text, locationFloat, sizeF, padding, border, container, foreColor, backColor,
                    font, textAlignment, styleName, report, "true", isBorderNeedChange, isFontNeedChange, isForeColorNeedChange,
                    isBackColorNeedChange);

            if (itemType == "XRCheckBox")
                ProcessCheckBox(itemName, text, locationFloat, sizeF, isChecked, border, container, styleName, report);

            if (itemType == "XRPictureBox")
                ProcessPictureBox(itemName, imageSource, sizing, locationFloat, sizeF, container);

            if (itemType == "XRPanel")
                ProcessPanel(baseNode, itemName, locationFloat, sizeF, container, foreColor, report, styleName, backColor);

            if (itemType == "XRRichText")
                ProcessRichText(itemName, serializableRtfString, locationFloat, sizeF, padding, border, container, styleName, report);

            if (itemType == "XRCharacterComb")
                ProcessCharacterComb(itemName, text, cellWidth, cellHeight, locationFloat, sizeF, padding, border, container, styleName,
                    report, cellHorizontalSpacing, cellVerticalSpacing, multiline);

            if (itemType == "XRLine")
                ProcessLine(itemName, lineWidth, lineDirection, lineStyle, foreColor, locationFloat, sizeF, container, styleName,
                    report);

            if (itemType == "XRShape")
                ProcessShape(baseNode, itemName, lineWidth, lineStyle, angle, fillColor, locationFloat, sizeF, foreColor,
                    borderDashStyle, container, styleName, report);

            if (itemType == "XRBarCode")
                ProcessBarCode(baseNode, itemName, module, autoModule, text, locationFloat, sizeF, container, textAlignment, styleName,
                    report, backColor, foreColor);

            if (itemType == "XRChart")
                ProcessChart(baseNode, itemName, border, locationFloat, sizeF, container, styleName, report);

            if (itemType == "XRTable")
                ProcessTable(baseNode, itemName, border, locationFloat, sizeF, container, foreColor, backColor, font, report);

            if (itemType == "XRCrossTab")
                ProcessCrossTab(baseNode, container, dataSourceRef, dataMember);

            if (itemType == "XRTableCell")
                ProcessTextBox(baseNode, itemName, text, locationFloat, sizeF, padding, border, container, foreColor, backColor,
                    font, textAlignment, styleName, report, multiline, isBorderNeedChange, isFontNeedChange, isForeColorNeedChange,
                    isBackColorNeedChange);
        }

        private void ProcessTextBox(XmlNode baseNode, string itemName, string text, string locationFloat, string sizeF, string padding,
            StiBorder border, StiContainer container, Color foreColor, Color backColor, Font font, string textAlignment,
            string styleName, StiReport report, string multiline, bool isBorderNeedChange, bool isFontNeedChange,
            bool isForeColorNeedChange, bool isBackColorNeedChange)
        {
            StiText component = new StiText();

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(component);

            component.Name = itemName;
            component.CanGrow = true;
            component.CanShrink = false;

            if (multiline == "true")
                component.WordWrap = true;

            else
                component.WordWrap = false;

            component.Text = text;

            if (isBorderNeedChange || style == null)
                component.Border = border;

            if (isBackColorNeedChange || style == null)
                component.Brush = new StiSolidBrush(backColor);

            if (isForeColorNeedChange || style == null)
                component.TextBrush = new StiSolidBrush(foreColor);

            if (isFontNeedChange || style == null)
                component.Font = font;

            if (!string.IsNullOrEmpty(textAlignment))
            {
                TextHorizontalAlignment(component, textAlignment);
                VerticalAlignment(component, textAlignment);
            }

            ClientRectangle(locationFloat, sizeF, component);

            if (!string.IsNullOrEmpty(padding))
            {
                var paddingArray = padding.Split(',');

                if (paddingArray.Length > 3)
                {
                    component.Margins = new StiMargins(
                        ParseDouble(paddingArray[0]),
                        ParseDouble(paddingArray[1]),
                        ParseDouble(paddingArray[2]),
                        ParseDouble(paddingArray[3]));
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ExpressionBindings":
                        ProcessExpressionBindings(node, component);
                        break;
                }
            }

            CheckExpression(component);

            container.Components.Add(component);
        }

        private void ProcessCheckBox(string itemName, string text, string locationFloat, string sizeF, string isChecked,
            StiBorder border, StiContainer container, string styleName, StiReport report)
        {
            var panel = new StiPanel();
            panel.Name = $"{itemName}Panel";

            var component = new StiCheckBox();
            component.Name = itemName;

            if (isChecked == "true")
            {
                component.Checked = new StiCheckedExpression();
                component.Checked.Value = "true";
            }

            ClientRectangle(locationFloat, sizeF, component);

            if (!string.IsNullOrEmpty(text))
            {
                var newText = new StiText();
                newText.Text = text;
                newText.Name = $"{itemName}Text";
                newText.Font = defaultFont;

                var style = report.Styles[styleName];

                if (style != null)
                    style.SetStyleToComponent(newText);

                var g = Graphics.FromImage(new Bitmap(1, 1));
                var textSize = g.MeasureString(text, defaultFont);

                component.ClientRectangle = new RectangleD(
                    component.ClientRectangle.X,
                    component.ClientRectangle.Y,
                    15,
                    15);

                newText.ClientRectangle = new RectangleD(
                    component.ClientRectangle.X + component.ClientRectangle.Width,
                    component.ClientRectangle.Y,
                    ToHi(reportUnit, textSize.Width.ToString()) / 2,
                    ToHi(reportUnit, textSize.Height.ToString()) / 2);

                component.Border = border;

                panel.Components.Add(newText);
                panel.Components.Add(component);
                container.Components.Add(panel);
            }
            else
            {
                component.Border = border;

                container.Components.Add(component);
            }
        }

        private void ProcessPictureBox(string itemName, string imageSource, string sizing, string locationFloat, string sizeF,
            StiContainer container)
        {
            StiImage component = new StiImage();
            component.Name = itemName;
            component.CanGrow = true;
            component.CanShrink = false;

            if (imageSource.Length > 2 && imageSource.Remove(3, imageSource.Length - 3) == "img")
            {
                var im = Convert.FromBase64String(imageSource.Remove(0, 4));

                using (MemoryStream ms = new MemoryStream(im))
                {
                    component.Image = Image.FromStream(ms);
                }
            }
            else if (imageSource.Length > 2 && imageSource.Remove(3, imageSource.Length - 3) == "svg")
            {
                var bytes = Convert.FromBase64String(imageSource.Remove(0, 4));
                component.ImageBytes = bytes;
            }

            if (sizing == "NormalImage")
            {
                component.CanGrow = false;
            }
            else if (sizing == "StretchImage")
            {
                component.CanGrow = false;
                component.Stretch = true;
            }
            else if (sizing == "ZoomImage")
            {
                component.CanGrow = false;
                component.Stretch = true;
                component.AspectRatio = true;
                component.VertAlignment = StiVertAlignment.Center;
                component.HorAlignment = StiHorAlignment.Center;
            }

            ClientRectangle(locationFloat, sizeF, component);

            container.Components.Add(component);
        }

        private void ProcessPanel(XmlNode baseNode, string itemName, string locationFloat, string sizeF, StiContainer container,
            Color forecolor, StiReport report, string styleName, Color backColor)
        {
            StiPanel component = new StiPanel();
            component.Name = itemName;
            component.CanGrow = true;
            component.CanShrink = false;
            component.Brush = new StiSolidBrush(backColor);

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(component); ;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Controls":
                        ProcessControls(node, component, forecolor, report);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            ClientRectangle(locationFloat, sizeF, component);

            container.Components.Add(component);
        }

        private void ProcessRichText(string itemName, string serializableRtfString, string locationFloat, string sizeF, string padding,
            StiBorder border, StiContainer container, string styleName, StiReport report)
        {
            StiRichText component = new StiRichText();
            component.Name = itemName;
            component.CanGrow = true;
            component.CanShrink = false;
            component.WordWrap = true;
            component.Border = border;
            component.RtfText = Base64Decode(serializableRtfString);

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(component);

            ClientRectangle(locationFloat, sizeF, component);

            if (!string.IsNullOrEmpty(padding))
            {
                var paddingArray = padding.Split(',');

                if (paddingArray.Length > 3)
                {
                    component.Margins = new StiMargins(
                        ParseDouble(paddingArray[0]),
                        ParseDouble(paddingArray[1]),
                        ParseDouble(paddingArray[2]),
                        ParseDouble(paddingArray[3]));
                }
            }

            container.Components.Add(component);
        }

        private void ProcessCharacterComb(string itemName, string text, string cellWidth, string cellHeight, string locationFloat,
            string sizeF, string padding, StiBorder border, StiContainer container, string styleName, StiReport report,
            string horizontalSpacing, string verticalSpacing, string multiline)
        {
            var component = new StiTextInCells();
            component.Name = itemName;
            component.CellWidth = (float)ToHi(reportUnit, cellWidth);
            component.CellHeight = (float)ToHi(reportUnit, cellHeight);
            component.HorSpacing = (float)ToHi(reportUnit, horizontalSpacing);
            component.VertSpacing = (float)ToHi(reportUnit, verticalSpacing);
            component.CanGrow = false;
            component.CanShrink = false;

            if (multiline == "true")
                component.WordWrap = true;

            else
                component.WordWrap = false;

            component.Text = text;
            component.Border = border;

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(component);

            ClientRectangle(locationFloat, sizeF, component);

            if (!string.IsNullOrEmpty(padding))
            {
                var paddingArray = padding.Split(',');

                if (paddingArray.Length > 3)
                {
                    component.Margins = new StiMargins(
                        ParseDouble(paddingArray[0]),
                        ParseDouble(paddingArray[1]),
                        ParseDouble(paddingArray[2]),
                        ParseDouble(paddingArray[3]));
                }
            }

            container.Components.Add(component);
        }

        private void ProcessLine(string itemName, string lineWidth, string lineDirection, string lineStyle, Color foreColor,
            string locationFloat, string sizeF, StiContainer container, string styleName, StiReport report)
        {
            var component = ShapeLineComponent(lineDirection);
            component.Name = itemName;
            component.Width = ParseDouble(lineWidth);
            component.Style = ParseBorderStyle(lineStyle);
            component.BorderColor = foreColor;

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(component);

            ClientRectangle(locationFloat, sizeF, component);

            container.Components.Add(component);
        }

        private void ProcessShape(XmlNode baseNode, string itemName, string lineWidth, string lineStyle, string angle, Color fillColor,
            string locationFloat, string sizeF, Color foreColor, StiPenStyle borderDashStyle, StiContainer container, string styleName,
            StiReport report)
        {
            StiShape component = new StiShape { ShapeType = new StiOvalShapeType() };

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Shape":
                        component = ProcessShapeType(node, angle);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            component.Name = baseNode.Name;
            component.Brush = new StiSolidBrush(fillColor);

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(component);

            ClientRectangle(locationFloat, sizeF, component);

            container.Components.Add(component);
        }

        private StiShape ProcessShapeType(XmlNode baseNode, string angle)
        {
            var component = new StiShape { ShapeType = new StiOvalShapeType() };
            var shapeName = "Ellips";
            var fillet = 0d;
            var numberOfSides = "3";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "ShapeName":
                        shapeName = node.Value;
                        break;

                    case "Fillet":
                        fillet = ParseDouble(node.Value);
                        break;

                    case "NumberOfSides":
                        numberOfSides = node.Value;
                        break;
                }
            }

            if (shapeName == "Rectangle")
            {
                component = new StiShape { ShapeType = new StiRectangleShapeType() };
            }
            else if (shapeName == "Arrow")
            {
                if (angle == "0")
                    component = new StiShape { ShapeType = new StiArrowShapeType() { Direction = StiShapeDirection.Up } };

                else if (angle == "270")
                    component = new StiShape { ShapeType = new StiArrowShapeType() { Direction = StiShapeDirection.Right } };

                else if (angle == "180")
                    component = new StiShape { ShapeType = new StiArrowShapeType() { Direction = StiShapeDirection.Down } };

                else if (angle == "90")
                    component = new StiShape { ShapeType = new StiArrowShapeType() { Direction = StiShapeDirection.Left } };
            }
            else if (shapeName == "Polygon")
            {
                if (numberOfSides == "3")
                    component = new StiShape { ShapeType = new StiTriangleShapeType() };

                else if (numberOfSides == "4")
                    component = new StiShape { ShapeType = new StiRectangleShapeType() };

                else if (numberOfSides == "5")
                    component = new StiShape { ShapeType = new StiRegularPentagonShapeType() };

                else if (numberOfSides == "6")
                    component = new StiShape { ShapeType = new StiFlowchartPreparationShapeType() };

                else if (numberOfSides == "8")
                    component = new StiShape { ShapeType = new StiOctagonShapeType() };
            }
            else if (shapeName == "Star")
            {
                component = new StiShape { ShapeType = new StiRectangleShapeType() };
            }
            else if (shapeName == "Line")
            {
                if (angle == "0")
                    component = new StiShape { ShapeType = new StiVerticalLineShapeType() };

                else if (angle == "270")
                    component = new StiShape { ShapeType = new StiHorizontalLineShapeType() };

                else if (angle == "135")
                    component = new StiShape { ShapeType = new StiDiagonalUpLineShapeType() };

                else if (angle == "225")
                    component = new StiShape { ShapeType = new StiDiagonalDownLineShapeType() };
            }
            else if (shapeName == "Cross")
            {
                component = new StiShape { ShapeType = new StiPlusShapeType() };
            }
            else if (shapeName == "Bracket")
            {
                component = new StiShape { ShapeType = new StiVerticalLineShapeType() };

            }
            else if (shapeName == "Brace")
            {
                component = new StiShape { ShapeType = new StiVerticalLineShapeType() };
            }

            return component;
        }

        private void ProcessBarCode(XmlNode baseNode, string itemName, string module, string autoModule, string text,
            string locationFloat, string sizeF, StiContainer container, string textAlignment, string styleName, StiReport report
            , Color backColor, Color foreColor)
        {
            var component = new StiBarCode();
            component.Name = itemName;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Symbology":
                        ProcessSymbology(node, component);
                        break;
                }
            }

            component.Code = new StiBarCodeExpression(text);
            component.BackColor = backColor;
            component.ForeColor = foreColor;

            HorizontalAlignment(component, textAlignment);
            VerticalAlignment(component, textAlignment);

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(component);

            ClientRectangle(locationFloat, sizeF, component);

            container.Components.Add(component);
        }

        private void ProcessSymbology(XmlNode baseNode, StiBarCode component)
        {
            var name = "Code93";
            var wideNarrowRatio = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        break;

                    case "WideNarrowRatio":
                        wideNarrowRatio = node.Value;
                        break;
                }
            }

            BarcodeType(component, name);
        }

        private void ProcessChart(XmlNode baseNode, string itemName, StiBorder border, string locationFloat, string sizeF,
            StiContainer container, string styleName, StiReport report)
        {
            var component = new StiChart();
            component.Name = itemName;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Parameters":
                        ProcessChartParameters(node, component);
                        break;

                    case "Chart":
                        ProcessChartOptions(node, component);
                        break;
                }
            }

            ClientRectangle(locationFloat, sizeF, component);

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(component);

            container.Components.Add(component);
        }

        private void ProcessChartParameters(XmlNode baseNode, StiChart container)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    ProcessChartItem(node, container);
                    itemIndex++;
                }
            }
        }

        private void ProcessChartItem(XmlNode baseNode, StiChart chart)
        {
            var name = "";
            var dataSource = "";
            var dataMember = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        break;

                    case "DataSource":
                        dataSource = node.Value;
                        break;

                    case "DataMember":
                        dataMember = node.Value;
                        break;
                }
            }
        }

        private void ProcessChartOptions(XmlNode baseNode, StiChart chart)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataContainer":
                        ProcessDataContainer(node, chart);
                        break;

                    case "Diagram":
                        ProcessDiagram(node, chart);
                        break;

                    case "Legend":
                        //ProcessLegend(node, component);
                        break;

                    case "OptionsPrint":
                        //ProcessOptionsPrint(node, component);
                        break;
                }
            }
        }

        private void ProcessDataContainer(XmlNode baseNode, StiChart chart)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "SeriesSerializable":
                        ProcessSeriesSerializable(node, chart);
                        break;
                }
            }
        }

        private void ProcessSeriesSerializable(XmlNode baseNode, StiChart chart)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    ProcessSeriesSerializableItem(node, chart);
                    itemIndex++;
                }
            }
        }

        private void ProcessSeriesSerializableItem(XmlNode baseNode, StiChart chart)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "View":
                        ProcessView(node, chart);
                        break;
                }
            }
        }

        private void ProcessView(XmlNode baseNode, StiChart chart)
        {
            var name = "Bar";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "TypeNameSerializable":
                        name = node.Value;
                        break;
                }
            }

            if (name == "Bar")
            {
                var area = new StiClusteredColumnArea();
                var series = new StiClusteredColumnSeries();
                chart.Area = area;
                chart.Series.Add(series);
            }
            else if (name == "StackedBarSeriesView")
            {
                var area = new StiStackedColumnArea();
                var series = new StiStackedColumnSeries();
                chart.Area = area;
                chart.Series.Add(series);
            }
        }

        private void ProcessDiagram(XmlNode baseNode, StiChart chart)
        {
            var type = "";
            var axisX = "";
            var axisY = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "TypeNameSerializable":
                        type = node.Value;
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "AxisX":
                        axisX = ProcessAxis(node);
                        break;

                    case "AxisY":
                        axisY = ProcessAxis(node);
                        break;

                    case "DefaultPane":
                        //ProcessPane(node, chart);
                        break;
                }
            }
        }

        private string ProcessAxis(XmlNode baseNode)
        {
            var visibleInPanesSerializable = "0";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "VisibleInPanesSerializable":
                        visibleInPanesSerializable = node.Value;
                        break;
                }
            }

            return visibleInPanesSerializable;
        }

        private void ProcessTable(XmlNode baseNode, string itemName, StiBorder border, string locationFloat, string sizeF,
            StiContainer container, Color forecolor, Color backColor, Font font, StiReport report)
        {
            var rows = new List<List<StiPanel>>();
            var tablePanel = new StiPanel();
            tablePanel.CanGrow = true;
            StiComponent[,] cells = null;

            if (!string.IsNullOrEmpty(itemName))
            {
                tablePanel.Name = itemName;
            }
            else
            {
                tablePanel.Name = $"{container.Name}Table{bandCountForDifferentName}";
                bandCountForDifferentName++;
            }

            ClientRectangle(locationFloat, sizeF, tablePanel);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Rows":
                        cells = ProcessTableRows(node, rows, border, tablePanel.ClientRectangle.Width,
                            tablePanel.ClientRectangle.Height, forecolor, backColor, font, report);
                        break;
                }
            }

            for (int indexRow = 0; indexRow < rows.Count; indexRow++)
            {
                var posX = 0d;

                for (int indexColumn = 0; indexColumn < rows[0].Count; indexColumn++)
                {
                    StiComponent component = cells[indexRow, indexColumn];

                    if (component != null)
                    {
                        if (indexRow > 0)
                            component.ClientRectangle = new RectangleD(
                                posX,
                                cells[indexRow - 1, indexColumn].ClientRectangle.Y + cells[indexRow - 1, indexColumn].ClientRectangle.Height,
                                component.ClientRectangle.Width,
                                component.ClientRectangle.Height);

                        component.Linked = true;

                        if (component.Bottom > tablePanel.Height)
                            tablePanel.Height = component.Bottom;

                        tablePanel.Components.Add(component);
                    }

                    if (component != null)
                        posX += component.ClientRectangle.Width;
                }
            }

            container.Components.Add(tablePanel);
        }

        private StiComponent[,] ProcessTableRows(XmlNode baseNode, List<List<StiPanel>> rows, StiBorder border, double rowWidth,
            double tableHeight, Color forecolor, Color backColor, Font font, StiReport report)
        {
            var itemIndex = 1;
            var tableCellsRowSpanCollection = new List<List<int>>();
            var tableRowsHeightCollection = new List<double>();
            var allTableRowsWeight = 0d;
            StiComponent[,] tableCells = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    foreach (XmlNode nodee in node.Attributes)
                    {
                        switch (nodee.Name)
                        {
                            case "Weight":
                                allTableRowsWeight += ParseDouble(nodee.Value);
                                break;
                        }
                    }

                    itemIndex++;
                }
            }

            itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    var rowSPanCollection = new List<int>();
                    List<StiPanel> row = ProcessTableRowItem(node, border, rowWidth, allTableRowsWeight, tableHeight, forecolor,
                        rowSPanCollection, backColor, font, report, tableRowsHeightCollection);
                    rows.Add(row);
                    tableCellsRowSpanCollection.Add(rowSPanCollection);
                    itemIndex++;
                }
            }

            #region RowSpans manifest
            for (int rowIndex = 0; rowIndex < tableCellsRowSpanCollection.Count; rowIndex++)
            {
                var rowSpans = tableCellsRowSpanCollection[rowIndex];

                for (int spanIndex = 0; spanIndex < rowSpans.Count; spanIndex++)
                {
                    if (rowSpans[spanIndex] > 1)
                    {
                        for (int index = rowIndex + 1; index < rowIndex + rowSpans[spanIndex]; index++)
                        {
                            tableCellsRowSpanCollection[index][spanIndex] = 0;
                        }
                    }
                }
            }
            #endregion

            if (rows.Count > 0)
            {
                if (rows[0].Count > 0)
                {
                    tableCells = new StiComponent[rows.Count, rows[0].Count];

                    TableFilling(rows, tableCells, tableCellsRowSpanCollection, tableRowsHeightCollection);
                }
            }

            return tableCells;
        }

        private List<StiPanel> ProcessTableRowItem(XmlNode baseNode, StiBorder border, double rowWidth, double allTableRowsWeight,
            double tableHeight, Color forecolor, List<int> rowSPanCollection, Color defaultBackColor, Font defaultFont,
            StiReport report, List<double> tableRowsHeightCollection)
        {
            var name = "";
            var controlType = "";
            var weight = 1d;
            var panels = new List<StiPanel>();
            var backColor = defaultBackColor;
            var font = defaultFont;
            var textAlignment = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        break;

                    case "ControlType":
                        controlType = node.Value;
                        break;

                    case "Weight":
                        weight = ParseDouble(node.Value);
                        break;

                    case "BackColor":
                        backColor = ParseStyleColor(node.Value);
                        break;

                    case "TextAlignment":
                        textAlignment = node.Value;
                        break;
                }
            }

            var rowHeight = tableHeight / (allTableRowsWeight / weight);
            tableRowsHeightCollection.Add(rowHeight);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Cells":
                        panels = ProcessCells(node, border, backColor, rowWidth, rowHeight, forecolor, rowSPanCollection, font,
                            report, textAlignment);
                        break;
                }
            }

            var maxItemHeight = 0d;

            foreach (var item in panels)
            {
                var comp = item as StiComponent;

                if (comp.ClientRectangle.Height > maxItemHeight)
                    maxItemHeight = comp.ClientRectangle.Height;
            }

            foreach (var item in panels)
            {
                if (item.Components.Count == 1 && item.Components[0].ClientRectangle.Height == item.ClientRectangle.Height)
                {
                    var comp = item.Components[0];

                    comp.ClientRectangle = new RectangleD(
                        comp.ClientRectangle.X,
                        comp.ClientRectangle.Y,
                        comp.ClientRectangle.Width,
                        maxItemHeight);
                }

                item.ClientRectangle = new RectangleD(
                    item.ClientRectangle.X,
                    item.ClientRectangle.Y,
                    item.ClientRectangle.Width,
                    maxItemHeight);
            }

            return panels;
        }

        private List<StiPanel> ProcessCells(XmlNode baseNode, StiBorder border, Color backColor, double rowWidth, double rowHeight,
            Color foreColor, List<int> rowSPanCollection, Font font, StiReport report, string textAlignment)
        {
            var itemIndex = 1;
            var rowCellsCount = 1;
            var panels = new List<StiPanel>();
            var defaultCellWidth = 0d;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{rowCellsCount}")
                    rowCellsCount++;
            }

            if (rowCellsCount > 0)
                rowCellsCount--;

            if (rowCellsCount <= 1)
                defaultCellWidth = rowWidth;

            else
                defaultCellWidth = rowWidth / rowCellsCount;

            allRowCellsWeight = 0d;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    foreach (XmlNode nodee in node.Attributes)
                    {
                        switch (nodee.Name)
                        {
                            case "Weight":
                                allRowCellsWeight += ParseDouble(nodee.Value);
                                break;
                        }
                    }

                    itemIndex++;
                }
            }

            itemIndex = 1;

            var rowWeight = allRowCellsWeight;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    panels.Add(ProcessCellItem(node, border, backColor, defaultCellWidth, rowHeight, foreColor, rowSPanCollection,
                        rowCellsCount, rowWidth, rowWeight, font, report, textAlignment));
                    itemIndex++;
                }
            }

            allRowCellsWeight = 0d;

            return panels;
        }

        private StiPanel ProcessCellItem(XmlNode baseNode, StiBorder border, Color backColor, double defaultCellWidth,
            double rowHeight, Color foreColor, List<int> rowSPanCollection, int rowCellsCount, double rowWidth,
            double rowWeight, Font defFont, StiReport report, string textAlignment = null)
        {
            var name = "";
            var controlType = "";
            var weight = 1d;
            var text = "";
            var textFromDataBinding = "";
            var backcolor = backColor;
            var isBorderColorNeedChange = false;
            var isFontNeedChange = false;
            var borderColor = border.Color;
            var rowSpan = 1;
            var font = defFont;
            var styleName = "";
            var measureCoeff = 1.4f;
            var canGrow = "true";
            var multiline = "false";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        break;

                    case "ControlType":
                        controlType = node.Value;
                        break;

                    case "Weight":
                        weight = ParseDouble(node.Value);
                        break;

                    case "Text":
                        text = node.Value;
                        break;

                    case "BorderColor":
                        borderColor = ParseStyleColor(node.Value);
                        isBorderColorNeedChange = true;
                        break;

                    case "BackColor":
                        backcolor = ParseStyleColor(node.Value);
                        break;

                    case "RowSpan":
                        rowSpan = int.Parse(node.Value);
                        break;

                    case "Font":
                        font = ParseFont(node.Value);
                        isFontNeedChange = true;
                        break;

                    case "StyleName":
                        styleName = node.Value;
                        break;

                    case "CanGrow":
                        canGrow = node.Value;
                        break;

                    case "Multiline":
                        multiline = node.Value;
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataBindings":
                        var itemIndex = 1;
                        var propertyName = "";
                        var propertyValue = "";

                        foreach (XmlNode nodee in node.ChildNodes)
                        {
                            if (nodee.Name == $"Item{itemIndex}")
                            {
                                foreach (XmlNode nodeeee in nodee.Attributes)
                                {
                                    switch (nodeeee.Name)
                                    {
                                        case "PropertyName":
                                            propertyName = nodeeee.Value;
                                            break;

                                        case "DataMember":
                                            propertyValue = nodeeee.Value;
                                            break;
                                    }
                                }

                                itemIndex++;
                            }
                        }

                        if (!string.IsNullOrEmpty(propertyName) && propertyName == "Text")
                            textFromDataBinding = $"{{{propertyValue}}}";

                        break;
                }
            }

            rowSPanCollection.Add(rowSpan);

            var isDefaultCellText = false;
            var panel = new StiPanel();

            if (canGrow == "true")
                panel.CanGrow = true;
            else
                panel.CanGrow = false;

            panel.CanShrink = false;

            var style = report.Styles[styleName];

            if (style != null)
                style.SetStyleToComponent(panel);

            if (isBorderColorNeedChange)
                panel.Border = new StiBorder(border.Side, borderColor, border.Size, border.Style);

            else
                panel.Border = border;

            panel.Brush = new StiSolidBrush(backcolor);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Controls":
                        ProcessControls(node, panel, foreColor, report, border, font, textAlignment);
                        break;
                }
            }

            if (panel.Components.Count == 0)
            {
                ProcessItem(baseNode, panel, foreColor, backColor, font, report, border, textAlignment);

                if (panel.Components.Count > 0)
                {
                    panel.Components[0].ClientRectangle = new RectangleD(
                        panel.ClientRectangle.X,
                        panel.ClientRectangle.Y,
                        panel.ClientRectangle.Width,
                        panel.ClientRectangle.Height);

                    panel.Components[0].Name = $"{panel.Components[0].Name}Label";

                    if (panel.Components[0] is StiText)
                    {
                        var newText = panel.Components[0] as StiText;

                        if (isBorderColorNeedChange)
                            newText.Border = new StiBorder(border.Side, borderColor, border.Size, border.Style);

                        if (isFontNeedChange)
                            newText.Font = font;

                        if (!string.IsNullOrEmpty(textFromDataBinding))
                            newText.Text = textFromDataBinding;
                    }
                    else
                    {
                        if (style != null)
                            style.SetStyleToComponent(panel.Components[0]);
                    }
                }

                isDefaultCellText = true;
            }

            if (panel.Components.Count == 0)
            {
                var newText = new StiText();
                newText.Name = $"{name}Label";
                newText.Text = text;
                newText.Font = font;
                newText.ClientRectangle = new RectangleD(
                    panel.ClientRectangle.X,
                    panel.ClientRectangle.Y,
                    panel.ClientRectangle.Width,
                    panel.ClientRectangle.Height);

                if (style != null)
                    style.SetStyleToComponent(newText);

                if (isBorderColorNeedChange)
                    newText.Border = new StiBorder(border.Side, borderColor, border.Size, border.Style);

                else
                    newText.Border = border;

                if (isFontNeedChange)
                    newText.Font = font;

                if (!string.IsNullOrEmpty(textFromDataBinding))
                    newText.Text = textFromDataBinding;

                panel.Components.Add(newText);

                isDefaultCellText = true;
            }

            if (string.IsNullOrEmpty(panel.Name))
                panel.Name = $"{name}Panel";

            var maxItemHeight = rowHeight;

            foreach (var item in panel.Components)
            {
                var comp = item as StiComponent;

                if (comp is StiText t)
                {
                    var g = Graphics.FromImage(new Bitmap(1, 1));

                    var correctString = StringWordWrap(t.Text, t.Font, (float)(rowWidth / (rowWeight / weight) -
                        t.Margins.Left - t.Margins.Right), 1000, false, measureCoeff);

                    var size = MeasureSize(correctString, t.Font);

                    if (size.Height * measureCoeff + t.Margins.Top + t.Margins.Bottom > maxItemHeight
                        && !t.Text.ToString().StartsWith("{") && !t.Text.ToString().StartsWith("[")
                        && !t.Text.ToString().Contains("[") && !t.Text.ToString().Contains("]"))
                        maxItemHeight = size.Height * measureCoeff + t.Margins.Top + t.Margins.Bottom;
                }

                if (comp.ClientRectangle.Height > maxItemHeight)
                    maxItemHeight = comp.ClientRectangle.Height;
            }

            if (rowCellsCount == 1 || (rowCellsCount > 1 && rowWidth / (rowWeight / weight) > rowWidth))
                panel.ClientRectangle = new RectangleD(0, 0, rowWidth, maxItemHeight);

            else
                panel.ClientRectangle = new RectangleD(0, 0, rowWidth / (rowWeight / weight), maxItemHeight);

            if (isDefaultCellText && panel.Components.Count == 1)
            {
                panel.Components[0].ClientRectangle = panel.ClientRectangle;
            }

            return panel;
        }

        private void ProcessExpressionBindings(XmlNode baseNode, StiText component)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    ProcessExpressionItem(node, component);
                    itemIndex++;
                }
            }
        }

        private void ProcessExpressionItem(XmlNode baseNode, StiText component)
        {
            var eventName = "";
            var propertyName = "";
            var expression = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "EventName":
                        eventName = node.Value;
                        break;

                    case "PropertyName":
                        propertyName = node.Value;
                        break;

                    case "Expression":
                        expression = node.Value;
                        break;
                }
            }

            if (!string.IsNullOrEmpty(propertyName) && !string.IsNullOrEmpty(expression))
            {
                if (propertyName == "Text")
                    component.Text = $"[{expression}]";
            }
        }

        private void ProcessCrossTab(XmlNode baseNode, StiContainer container, string dataSourceRef, string dataMember)
        {
            var crossTable = new StiCrossTab();
            crossTable.Name = $"{container.Name}CrossTab";
            crossTable.ClientRectangle = new RectangleD(0, 0, crossTable.Width, crossTable.Height);
            crossTable.CanGrow = true;

            if (dataSourceReferenceCollection.Count > 0)
                currentDataSourceName = dataSourceReferenceCollection[dataSourceRef].ToString();

            if (!string.IsNullOrEmpty(dataMember))
                currentDataSourceName = (string)dataSourseCorrectNameCollection[dataMember];

            crossTable.DataSourceName = currentDataSourceName;

            var storeDataSource = currentDataSourceName;
            var rowFields = new List<StiText>();
            var columnFields = new List<StiText>();
            var dataFields = new List<StiText>();
            var columnWidths = new List<double>();
            var rowHeights = new List<double>();
            StiText[,] cells = null;

            var table = new StiPanel();
            table.CanGrow = true;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ColumnDefinitions":
                        ProcessCrossTabColumnWidth(node, columnWidths);
                        break;

                    case "RowDefinitions":
                        ProcessCrossTabRowHeight(node, rowHeights);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "RowFields":
                        ProcessCrossTabFields(node, rowFields);
                        break;

                    case "ColumnFields":
                        ProcessCrossTabFields(node, columnFields);
                        break;

                    case "DataFields":
                        ProcessCrossTabFields(node, dataFields);
                        break;

                    case "Cells":
                        if (rowHeights.Count > 0 && columnWidths.Count > 0)
                            cells = new StiText[rowHeights.Count, columnWidths.Count];

                        ProcessCrossTabCells(node, cells);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            currentDataSourceName = storeDataSource;

            #region CrossTab Builder
            var rowTotalMaxYCoordinate = 0d;
            var crossRowX = 0d;
            var defaultRowHeight = 16d;
            var defaultCellMargin = 2d;
            var rowWidthCollection = new List<double>();
            var allRowsWidth = 0d;
            var measureCoeff = 1.8d;

            for (int index = 0; index < rowFields.Count; index++)
            {
                var value = rowFields[index].Alias.Replace("[", "").Replace("]", "").Replace("\"", "");
                var width = MeasureSize(value, rowFields[index].Font).Width * measureCoeff;

                if (width < 30)
                    width = 30;

                allRowsWidth += width;
                rowWidthCollection.Add(width);
            }

            var leftTitleText = "";

            if (cells[0, 0] != null)
                leftTitleText = cells[0, 0].Text.ToString().Replace("{", "").Replace("}", "");

            else
                leftTitleText = currentDataSourceName.Replace("{", "").Replace("}", "");

            var leftTitleWidth = MeasureSize(leftTitleText, rowFields[0].Font);

            if (allRowsWidth < leftTitleWidth.Width * measureCoeff)
            {
                var difference = (leftTitleWidth.Width * measureCoeff - allRowsWidth) / rowWidthCollection.Count;

                allRowsWidth = leftTitleWidth.Width * measureCoeff;

                for (int index = 0; index < rowWidthCollection.Count; index++)
                {
                    rowWidthCollection[index] += difference;
                }
            }

            #region Rows
            for (int index = 0; index < rowFields.Count; index++)
            {
                if (rowFields[index] is StiText)
                {
                    var rowWidth = ParseDouble(rowWidthCollection[index].ToString());
                    var rowHeight = defaultRowHeight;

                    if (dataFields.Count > 0)
                        rowHeight = defaultRowHeight * dataFields.Count;

                    var rowTotal = new StiCrossRowTotal();
                    rowTotal.Conditions = rowFields[index].Conditions;
                    rowTotal.Font = rowFields[index].Font;
                    rowTotal.Margins = rowFields[index].Margins;
                    rowTotal.Name = $"{crossTable.Name}_RowTotal{index}";
                    rowTotal.Page = rowFields[index].Page;
                    rowTotal.Parent = rowFields[index].Parent;
                    rowTotal.Restrictions = rowFields[index].Restrictions;
                    rowTotal.TextBrush = rowFields[index].TextBrush;

                    var row = new StiCrossRow();

                    row.Alias = rowFields[index].Alias.Replace("[", "").Replace("]", "").Replace("\"", "");
                    row.ComponentStyle = rowFields[index].ComponentStyle;
                    row.ClientRectangle = new RectangleD
                        (
                        crossRowX,
                        10 + defaultCellMargin + defaultCellMargin + defaultRowHeight * columnFields.Count,
                        rowWidth,
                        rowHeight * (rowFields.Count - index)
                        );
                    row.Conditions = rowFields[index].Conditions;
                    row.DisplayValue = new StiDisplayCrossValueExpression(rowFields[index].Text.ToString().Replace(" ", ""));
                    row.Font = rowFields[index].Font;
                    row.Margins = rowFields[index].Margins;
                    row.Name = $"{crossTable.Name}_Row{index}";
                    row.Page = rowFields[index].Page;
                    row.Parent = rowFields[index].Parent;
                    row.Restrictions = rowFields[index].Restrictions;
                    row.TextBrush = rowFields[index].TextBrush;
                    row.Value = new StiCrossValueExpression(rowFields[index].Text.ToString().Replace(" ", ""));
                    row.Guid = Guid.NewGuid().ToString();
                    row.TotalGuid = rowTotal.Guid;

                    var rowTotalWidth = 0d;

                    if (index > 0 && index != rowWidthCollection.Count - 1)
                        rowTotalWidth = allRowsWidth - ParseDouble(rowWidthCollection[index - 1].ToString());

                    else if (index > 0 && index == rowWidthCollection.Count - 1)
                        rowTotalWidth = row.ClientRectangle.Width;

                    else
                        rowTotalWidth = allRowsWidth;

                    rowTotal.ClientRectangle = new RectangleD
                        (
                        crossRowX,
                        row.ClientRectangle.Y + row.ClientRectangle.Height,
                        rowTotalWidth,
                        rowHeight
                        );

                    if (rowTotal.ClientRectangle.Y + rowTotal.ClientRectangle.Height > rowTotalMaxYCoordinate)
                        rowTotalMaxYCoordinate = rowTotal.ClientRectangle.Y + rowTotal.ClientRectangle.Height;

                    crossTable.Components.Add(rowTotal);
                    crossTable.Components.Add(row);

                    var crossTitle = new StiCrossTitle();
                    crossTitle.ComponentStyle = rowFields[index].ComponentStyle;
                    crossTitle.Name = $"{crossTable.Name}_Row{index}_Title";
                    crossTitle.ClientRectangle = new RectangleD
                        (
                        crossRowX,
                        10 + defaultCellMargin,
                        rowWidth,
                        defaultRowHeight * columnFields.Count
                        );
                    crossTitle.Conditions = rowFields[index].Conditions;
                    crossTitle.TypeOfComponent = $"Row:{crossTable.Name}_Row{index}";
                    crossTitle.Font = rowFields[index].Font;
                    crossTitle.Margins = rowFields[index].Margins;
                    crossTitle.Page = rowFields[index].Page;
                    crossTitle.Parent = rowFields[index].Parent;
                    crossTitle.Restrictions = rowFields[index].Restrictions;
                    crossTitle.TextBrush = rowFields[index].TextBrush;

                    var rowValue = row.Value.ToString().Split('.');

                    if (rowValue.Length > 0)
                        crossTitle.Text = rowValue[rowValue.Length - 1].Replace("{", "").Replace("}", "");

                    crossTitle.Guid = Guid.NewGuid().ToString();

                    crossRowX = crossTitle.ClientRectangle.X + crossTitle.ClientRectangle.Width;

                    crossTable.Components.Add(crossTitle);
                }
            }
            #endregion

            var rowComponent = rowFields[0] as StiText;

            var leftTitle = new StiCrossTitle();
            leftTitle.ComponentStyle = rowComponent.ComponentStyle;
            leftTitle.Name = $"{crossTable.Name}_LeftTitle";
            leftTitle.TypeOfComponent = "LeftTitle";
            leftTitle.ClientRectangle = new RectangleD
                (
                0,
                0,
                crossRowX,
                10
                );
            leftTitle.Conditions = rowComponent.Conditions;
            leftTitle.Font = rowComponent.Font;
            leftTitle.Margins = rowComponent.Margins;
            leftTitle.Page = rowComponent.Page;
            leftTitle.Parent = rowComponent.Parent;
            leftTitle.Restrictions = rowComponent.Restrictions;
            leftTitle.TextBrush = rowComponent.TextBrush;

            if (cells[0, 0] != null)
                leftTitle.Text = cells[0, 0].Text.ToString().Replace("{", "").Replace("}", "");

            else
                leftTitle.Text = currentDataSourceName.Replace("{", "").Replace("}", "");

            leftTitle.Guid = Guid.NewGuid().ToString();

            var rightTitle = new StiCrossTitle();
            rightTitle.ComponentStyle = rowComponent.ComponentStyle;
            rightTitle.Name = $"{crossTable.Name}_RightTitle";
            rightTitle.TypeOfComponent = "RightTitle";
            rightTitle.Conditions = rowComponent.Conditions;
            rightTitle.Font = rowComponent.Font;
            rightTitle.Margins = rowComponent.Margins;
            rightTitle.Page = rowComponent.Page;
            rightTitle.Parent = rowComponent.Parent;
            rightTitle.Restrictions = rowComponent.Restrictions;
            rightTitle.TextBrush = rowComponent.TextBrush;
            rightTitle.Guid = Guid.NewGuid().ToString();

            var columnTotalMaxXCoordinate = 0d;
            var lastColumnWidthForData = 20d;
            var columnWidthCollection = new List<double>();

            for (int index = 0; index < columnFields.Count; index++)
            {
                var value = columnFields[index].Alias.Replace("[", "").Replace("]", "").Replace("\"", "");
                var width = MeasureSize(value, columnFields[index].Font).Width * measureCoeff;
                columnWidthCollection.Add(width);
            }

            var columnTotalWidthCollection = new double[columnWidthCollection.Count];

            for (int ind = columnWidthCollection.Count - 1; ind >= 0; ind--)
            {
                var width = columnWidthCollection[ind];
                var totalWidth = 0d;

                if (ind > 0)
                {
                    var nextWidth = columnWidthCollection[ind - 1];

                    if (nextWidth > width)
                    {
                        totalWidth = nextWidth - width;

                        if (totalWidth > 30)
                        {
                            columnWidthCollection[ind] = width + (totalWidth - 30);
                            totalWidth = 30;
                        }
                    }
                    else
                    {
                        totalWidth = 30;
                        columnWidthCollection[ind - 1] = width + totalWidth;
                    }
                }
                else
                {
                    totalWidth = 30;
                }

                columnTotalWidthCollection[ind] = totalWidth;
            }

            var allColumnTextsForRightTitle = "";

            #region Columns
            for (int index = 0; index < columnFields.Count; index++)
            {
                var columnWidth = ParseDouble(columnWidthCollection[index].ToString());
                var totalWidth = ParseDouble(columnTotalWidthCollection[index].ToString());

                var columnTotal = new StiCrossColumnTotal();
                columnTotal.Guid = Guid.NewGuid().ToString();

                var column = new StiCrossColumn();
                column.Alias = columnFields[index].Alias.Replace("[", "").Replace("]", "").Replace("\"", "");

                if (index < columnFields.Count - 1)
                    allColumnTextsForRightTitle += $"{column.Alias}, ";

                else
                    allColumnTextsForRightTitle += $"{column.Alias}";

                column.ComponentStyle = columnFields[index].ComponentStyle;

                column.ClientRectangle = new RectangleD
                    (
                    leftTitle.ClientRectangle.X + leftTitle.ClientRectangle.Width + defaultCellMargin,
                    10 + defaultCellMargin + defaultRowHeight * index,
                    columnWidth,
                    defaultRowHeight
                    );
                column.Conditions = columnFields[index].Conditions;
                column.DisplayValue = new StiDisplayCrossValueExpression(columnFields[index].Text.ToString().Replace(" ", ""));
                column.Font = columnFields[index].Font;
                column.Margins = columnFields[index].Margins;
                column.Name = $"{crossTable.Name}_Column{index}";
                column.Page = columnFields[index].Page;
                column.Parent = columnFields[index].Parent;
                column.Restrictions = columnFields[index].Restrictions;
                column.TextBrush = columnFields[index].TextBrush;
                column.Value = new StiCrossValueExpression(columnFields[index].Text.ToString().Replace(" ", ""));
                column.Guid = Guid.NewGuid().ToString();
                column.TotalGuid = columnTotal.Guid;

                var columnValue = column.Value.ToString().Split('.');

                crossTable.Components.Add(column);

                columnTotal.ClientRectangle = new RectangleD
                    (
                    leftTitle.ClientRectangle.X + leftTitle.ClientRectangle.Width + column.ClientRectangle.Width + defaultCellMargin,
                    column.ClientRectangle.Y,
                    totalWidth,
                    defaultRowHeight * (columnFields.Count - index)
                    );
                columnTotal.Conditions = columnFields[index].Conditions;
                columnTotal.Font = columnFields[index].Font;
                columnTotal.Margins = columnFields[index].Margins;
                columnTotal.Name = $"{crossTable.Name}_ColTotal{index}";
                columnTotal.Page = columnFields[index].Page;
                columnTotal.Parent = columnFields[index].Parent;
                columnTotal.Restrictions = columnFields[index].Restrictions;
                columnTotal.TextBrush = columnFields[index].TextBrush;

                if (columnTotal.ClientRectangle.X + columnTotal.ClientRectangle.Width > columnTotalMaxXCoordinate)
                    columnTotalMaxXCoordinate = columnTotal.ClientRectangle.X + columnTotal.ClientRectangle.Width;

                crossTable.Components.Add(columnTotal);

                if (index == columnFields.Count - 1)
                    lastColumnWidthForData = column.ClientRectangle.Width;
            }
            #endregion

            rightTitle.Text = allColumnTextsForRightTitle;

            rightTitle.ClientRectangle = new RectangleD
                (
                leftTitle.ClientRectangle.Width + defaultCellMargin,
                0,
                columnWidthCollection[0] + 30,
                10
                );

            #region Data
            for (int index = 0; index < dataFields.Count; index++)
            {
                var summary = new StiCrossSummary();

                if (dataFields[index] is StiText)
                {
                    summary.Alias = dataFields[index].Alias.Replace("[", "").Replace("]", "").Replace("\"", "");
                    summary.ComponentStyle = dataFields[index].ComponentStyle;
                    summary.ClientRectangle = new RectangleD
                        (
                        leftTitle.ClientRectangle.Width + defaultCellMargin,
                        10 + defaultCellMargin + defaultCellMargin + defaultRowHeight * columnFields.Count + defaultRowHeight * index,
                        lastColumnWidthForData,
                        defaultRowHeight
                        );
                    summary.Conditions = dataFields[index].Conditions;
                    summary.Font = dataFields[index].Font;
                    summary.Margins = dataFields[index].Margins;
                    summary.Name = $"{crossTable.Name}_Sum{index}";
                    summary.Page = dataFields[index].Page;
                    summary.Parent = dataFields[index].Parent;
                    summary.Restrictions = dataFields[index].Restrictions;
                    summary.TextBrush = dataFields[index].TextBrush;
                    summary.Value = new StiCrossValueExpression(dataFields[index].Text.ToString().Replace(" ", ""));
                    summary.Guid = Guid.NewGuid().ToString();

                    if (summaryTypeCollection.Count == 1)
                    {
                        if (summaryTypeCollection[0] == "Count")
                            summary.Summary = CrossTab.Core.StiSummaryType.Count;

                        else if (summaryTypeCollection[0] == "Avg")
                            summary.Summary = CrossTab.Core.StiSummaryType.Average;

                        else if (summaryTypeCollection[0] == "Sum")
                            summary.Summary = CrossTab.Core.StiSummaryType.Sum;

                        else if (summaryTypeCollection[0] == "Min")
                            summary.Summary = CrossTab.Core.StiSummaryType.Min;

                        else if (summaryTypeCollection[0] == "Max")
                            summary.Summary = CrossTab.Core.StiSummaryType.Max;
                    }
                    else if (summaryTypeCollection.Count == 0)
                    {
                        summary.Summary = CrossTab.Core.StiSummaryType.Sum;
                    }
                }
                else if (dataFields.Count == 0)
                {
                    summary.Alias = $"{crossTable.Name}_Sum1";
                    summary.ClientRectangle = new RectangleD
                        (
                        32,
                        34,
                        lastColumnWidthForData,
                        defaultRowHeight);
                    summary.Name = $"{crossTable.Name}_Sum1";
                    summary.Guid = Guid.NewGuid().ToString();
                }

                crossTable.Components.Add(summary);
            }
            #endregion

            crossTable.Components.Add(leftTitle);
            crossTable.Components.Add(rightTitle);

            crossTable.ClientRectangle = table.ClientRectangle;

            if (crossTable.Height < rowTotalMaxYCoordinate)
                crossTable.Height = rowTotalMaxYCoordinate;

            if (crossTable.Width < columnTotalMaxXCoordinate)
                crossTable.Width = columnTotalMaxXCoordinate;
            #endregion

            container.Components.Add(crossTable);
        }

        private void ProcessCrossTabFields(XmlNode baseNode, List<StiText> fields)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    var text = ProcessCrossTabFieldItem(node);
                    fields.Add(text);
                    itemIndex++;
                }
            }
        }

        private StiText ProcessCrossTabFieldItem(XmlNode baseNode)
        {
            var newText = new StiText();
            var fieldName = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "FieldName":
                        fieldName = $"[{node.Value}]";
                        break;
                }
            }

            newText.Text = fieldName;
            newText.Alias = fieldName;

            CheckExpression(newText);

            return newText;
        }

        private void ProcessCrossTabColumnWidth(XmlNode baseNode, List<double> widthCollection)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    var width = ProcessCrossTabFieldWidth(node);
                    widthCollection.Add(width);
                    itemIndex++;
                }
            }
        }

        private double ProcessCrossTabFieldWidth(XmlNode baseNode)
        {
            var width = 0d;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Width":
                        width = ParseDouble(node.Value);
                        break;
                }
            }

            return width;
        }

        private void ProcessCrossTabRowHeight(XmlNode baseNode, List<double> heightCollection)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    var height = ProcessCrossTabFieldHeight(node);
                    heightCollection.Add(height);
                    itemIndex++;
                }
            }
        }

        private double ProcessCrossTabFieldHeight(XmlNode baseNode)
        {
            var height = 0d;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Height":
                        height = ParseDouble(node.Value);
                        break;
                }
            }

            return height;
        }

        private void ProcessCrossTabCells(XmlNode baseNode, StiText[,] cells)
        {
            var itemIndex = 1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == $"Item{itemIndex}")
                {
                    ProcessCrossTabCellsItem(node, cells);
                    itemIndex++;
                }
            }
        }

        private void ProcessCrossTabCellsItem(XmlNode baseNode, StiText[,] cells)
        {
            var newText = new StiText();
            var name = "";
            var text = "";
            var columnIndex = "";
            var rowIndex = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        name = node.Value;
                        break;

                    case "Text":
                        text = node.Value;
                        break;

                    case "ColumnIndex":
                        columnIndex = node.Value;
                        break;

                    case "RowIndex":
                        rowIndex = node.Value;
                        break;
                }
            }

            newText.Text = text;
            newText.Name = name;

            CheckExpression(newText);

            cells[int.Parse(rowIndex), int.Parse(columnIndex)] = newText;
        }
        #endregion

        #region Utils
        private string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            var resultRtf = System.Text.Encoding.Unicode.GetString(base64EncodedBytes)
                .Replace("\\}", "}")
                .Replace("\\{", "{");

            if (resultRtf.StartsWith("{{") && resultRtf.EndsWith("}}"))
            {
                resultRtf = resultRtf.Remove(resultRtf.Length - 1, 1).Remove(0, 1);
            }

            return resultRtf;
        }

        private void CheckExpression(StiText label)
        {
            if (label.Text.ToString().StartsWith("[") && !label.Text.ToString().StartsWith("[sumCount"))
            {
                var text = label.Text.ToString().Replace("[", "").Replace("]", "");
                label.Text = new StiExpression($"{{{mainDataSourceName}.{text}}}");
            }
        }

        private void TextHorizontalAlignment(IStiTextHorAlignment component, string horizontalAlignment)
        {
            var horAlign = StiTextHorAlignment.Left;

            if (horizontalAlignment == "TopCenter")
                horAlign = StiTextHorAlignment.Center;

            else if (horizontalAlignment == "TopRight")
                horAlign = StiTextHorAlignment.Right;

            else if (horizontalAlignment == "MiddleLeft")
                horAlign = StiTextHorAlignment.Left;

            else if (horizontalAlignment == "MiddleCenter")
                horAlign = StiTextHorAlignment.Center;

            else if (horizontalAlignment == "MiddleRight")
                horAlign = StiTextHorAlignment.Right;

            else if (horizontalAlignment == "BottomLeft")
                horAlign = StiTextHorAlignment.Left;

            else if (horizontalAlignment == "BottomCenter")
                horAlign = StiTextHorAlignment.Center;

            else if (horizontalAlignment == "BottomRight")
                horAlign = StiTextHorAlignment.Right;

            else if (horizontalAlignment == "TopJustify")
                horAlign = StiTextHorAlignment.Width;

            else if (horizontalAlignment == "MiddleJustify")
                horAlign = StiTextHorAlignment.Width;

            else if (horizontalAlignment == "BottomJustify")
                horAlign = StiTextHorAlignment.Width;

            component.HorAlignment = horAlign;
        }

        private void VerticalAlignment(IStiVertAlignment component, string horizontalAlignment)
        {
            var vertAlign = StiVertAlignment.Top;

            if (horizontalAlignment == "TopCenter")
                vertAlign = StiVertAlignment.Top;

            else if (horizontalAlignment == "TopRight")
                vertAlign = StiVertAlignment.Top;

            else if (horizontalAlignment == "MiddleLeft")
                vertAlign = StiVertAlignment.Center;

            else if (horizontalAlignment == "MiddleCenter")
                vertAlign = StiVertAlignment.Center;

            else if (horizontalAlignment == "MiddleRight")
                vertAlign = StiVertAlignment.Center;

            else if (horizontalAlignment == "BottomLeft")
                vertAlign = StiVertAlignment.Bottom;

            else if (horizontalAlignment == "BottomCenter")
                vertAlign = StiVertAlignment.Bottom;

            else if (horizontalAlignment == "BottomRight")
                vertAlign = StiVertAlignment.Bottom;

            else if (horizontalAlignment == "TopJustify")
                vertAlign = StiVertAlignment.Top;

            else if (horizontalAlignment == "MiddleJustify")
                vertAlign = StiVertAlignment.Center;

            else if (horizontalAlignment == "BottomJustify")

                vertAlign = StiVertAlignment.Bottom;

            component.VertAlignment = vertAlign;
        }

        private void HorizontalAlignment(IStiHorAlignment component, string horizontalAlignment)
        {
            var horAlign = StiHorAlignment.Left;

            if (horizontalAlignment == "TopCenter")
                horAlign = StiHorAlignment.Center;

            else if (horizontalAlignment == "TopRight")
                horAlign = StiHorAlignment.Right;

            else if (horizontalAlignment == "MiddleLeft")
                horAlign = StiHorAlignment.Left;

            else if (horizontalAlignment == "MiddleCenter")
                horAlign = StiHorAlignment.Center;

            else if (horizontalAlignment == "MiddleRight")
                horAlign = StiHorAlignment.Right;

            else if (horizontalAlignment == "BottomLeft")
                horAlign = StiHorAlignment.Left;

            else if (horizontalAlignment == "BottomCenter")
                horAlign = StiHorAlignment.Center;

            else if (horizontalAlignment == "BottomRight")
                horAlign = StiHorAlignment.Right;

            else if (horizontalAlignment == "TopJustify")
                horAlign = StiHorAlignment.Center;

            else if (horizontalAlignment == "MiddleJustify")
                horAlign = StiHorAlignment.Center;

            else if (horizontalAlignment == "BottomJustify")
                horAlign = StiHorAlignment.Center;

            component.HorAlignment = horAlign;
        }

        private void TableFilling(List<List<StiPanel>> rows, StiComponent[,] tableCells, List<List<int>> tableCellsRowSpanCollection,
            List<double> tableRowsHeightCollection)
        {
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                var rowSpanCollection = tableCellsRowSpanCollection[rowIndex];

                var mainPanel = new StiPanel();
                mainPanel.CanGrow = true;
                mainPanel.CanShrink = false;
                posX = 0d;

                for (int cellIndex = 0; cellIndex < rows[rowIndex].Count; cellIndex++)
                {
                    var rowSpan = rowSpanCollection[cellIndex];
                    var newCellHeight = 0d;

                    for (int index = rowIndex; index <= rowIndex + rowSpan - 1; index++)
                    {
                        newCellHeight += tableRowsHeightCollection[index];
                    }

                    if (rows[rowIndex][cellIndex].ClientRectangle.Height > newCellHeight)
                        newCellHeight = rows[rowIndex][cellIndex].ClientRectangle.Height;

                    if (rows[rowIndex][cellIndex].Components.Count == 1 && rows[rowIndex][cellIndex].ClientRectangle.Height ==
                        rows[rowIndex][cellIndex].Components[0].ClientRectangle.Height)
                    {
                        rows[rowIndex][cellIndex].ClientRectangle = new RectangleD
                        (
                        rows[rowIndex][cellIndex].ClientRectangle.X,
                        rows[rowIndex][cellIndex].ClientRectangle.Y,
                        rows[rowIndex][cellIndex].ClientRectangle.Width,
                        newCellHeight
                        );

                        rows[rowIndex][cellIndex].Components[0].ClientRectangle = new RectangleD
                        (
                        rows[rowIndex][cellIndex].ClientRectangle.X,
                        rows[rowIndex][cellIndex].ClientRectangle.Y,
                        rows[rowIndex][cellIndex].ClientRectangle.Width,
                        rows[rowIndex][cellIndex].ClientRectangle.Height
                        );
                    }
                    else
                    {
                        rows[rowIndex][cellIndex].ClientRectangle = new RectangleD
                        (
                        rows[rowIndex][cellIndex].ClientRectangle.X,
                        rows[rowIndex][cellIndex].ClientRectangle.Y,
                        rows[rowIndex][cellIndex].ClientRectangle.Width,
                        newCellHeight
                        );
                    }

                    if (rowSpan == 0)
                    {
                        foreach (StiComponent item in rows[rowIndex][cellIndex].Components)
                        {
                            item.ClientRectangle = new RectangleD
                                                    (
                                                    item.ClientRectangle.X,
                                                    item.ClientRectangle.Y,
                                                    item.ClientRectangle.Width,
                                                    item.ClientRectangle.Height * rowSpan
                                                    );
                        }

                        var cellPanel = rows[rowIndex][cellIndex];

                        cellPanel.ClientRectangle = new RectangleD
                                            (
                                            cellPanel.ClientRectangle.X,
                                            cellPanel.ClientRectangle.Y,
                                            cellPanel.ClientRectangle.Width,
                                            cellPanel.ClientRectangle.Height * rowSpan
                                            );
                    }

                    if (rows[rowIndex][cellIndex].Components.Count == 1)
                    {
                        if (rows[rowIndex][cellIndex].Components[0] is StiPanel)
                        {
                            if (rows[rowIndex].Count >= tableCells.Rank)
                            {
                                var component = rows[rowIndex][cellIndex];

                                AddComponent(component, mainPanel);
                            }
                            else
                            {
                                tableCells[rowIndex, cellIndex] = rows[rowIndex][cellIndex].Components[0];
                            }
                        }
                        else
                        {
                            if (rows[rowIndex].Count >= tableCells.Rank)
                            {
                                var component = rows[rowIndex][cellIndex];

                                AddComponent(component, mainPanel);
                            }
                            else
                            {
                                var panel = new StiPanel();
                                panel.CanGrow = true;
                                panel.CanShrink = false;
                                panel.Name = $"Panel{rows[rowIndex][cellIndex].Components[0].Name}";
                                panel.Border = rows[rowIndex][cellIndex].Border;
                                panel.Components.Add(rows[rowIndex][cellIndex].Components[0]);
                                panel.Brush = rows[rowIndex][cellIndex].Brush;
                                panel.ClientRectangle = rows[rowIndex][cellIndex].ClientRectangle;

                                tableCells[rowIndex, cellIndex] = panel;
                            }
                        }
                    }
                    else
                    {
                        if (rows[rowIndex].Count >= tableCells.Rank)
                        {
                            var component = rows[rowIndex][cellIndex];

                            AddComponent(component, mainPanel);
                        }
                        else
                        {
                            tableCells[rowIndex, cellIndex] = rows[rowIndex][cellIndex];
                        }
                    }
                }

                if (rows[rowIndex].Count >= tableCells.Rank)
                {
                    var mainPanelWidth = 0d;

                    foreach (StiComponent cell in mainPanel.Components)
                        mainPanelWidth += cell.ClientRectangle.Width;

                    mainPanel.Name = $"Row{rowIndex}";
                    mainPanel.Border = rows[rowIndex][0].Border;
                    mainPanel.Brush = rows[rowIndex][0].Brush;
                    mainPanel.ClientRectangle = new RectangleD(
                        rows[rowIndex][0].ClientRectangle.X,
                        rows[rowIndex][0].ClientRectangle.Y,
                        mainPanelWidth,
                        rows[rowIndex][0].ClientRectangle.Height
                        );

                    if (tableCells.GetLength(1) >= mainPanel.Components.Count)
                    {
                        for (int index = 0; index < mainPanel.Components.Count; index++)
                        {
                            tableCells[rowIndex, index] = mainPanel.Components[index];
                        }
                    }
                    else
                    {
                        tableCells[rowIndex, 0] = mainPanel;
                    }
                }
            }
        }

        private void AddComponent(StiComponent component, StiPanel mainPanel)
        {
            component.ClientRectangle = new RectangleD(
                                    posX,
                                    component.ClientRectangle.Y,
                                    component.ClientRectangle.Width,
                                    component.ClientRectangle.Height
                                    );

            mainPanel.Components.Add(component);

            posX += component.ClientRectangle.Width;
        }

        private static string ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

        private double ToHi(string reportUnit, string strValue)
        {
            string stSize = strValue;
            double factor = 1;

            if (reportUnit == "HundredthsOfAnInch")
                factor = 96f / 100;

            else if (reportUnit == "TenthsOfAMillimeter")
                factor = 1 / 2.54 * 96f / 100;

            else if (reportUnit == "Pixels")
                //factor = 0.905f;

                if (stSize.Length == 0)
                {
                    ThrowError(null, null, "Expression in SizeType: " + strValue);
                    return 0;
                }

            double size = Convert.ToDouble(stSize.Replace(",", ".").Replace(".", ApplicationDecimalSeparator)) * factor;
            return Math.Round(size, 2);
        }

        private Font ParseFont(string strValue)
        {
            var font = new Font("Arial", 8);
            var fontParts = strValue.Replace("style=", "=").Split('=');

            if (fontParts.Length > 0)
            {
                var fontArray = fontParts[0].Split(',');

                var fontFamily = "";
                var fontSize = "";

                if (!string.IsNullOrEmpty(fontArray[0]))
                    fontFamily = fontArray[0];

                if (!string.IsNullOrEmpty(fontArray[1]))
                    fontSize = fontArray[1].Substring(0, fontArray[1].Length - 2).Replace(" ", "");

                float size = Convert.ToSingle(fontSize.Replace(",", ".").Replace(".", ApplicationDecimalSeparator));

                if (size < 1) size = 1;

                font = new Font(fontFamily, size);
            }

            if (fontParts.Length > 1)
            {
                var fontStyleArray = fontParts[1].Replace(" ", "").Split(',');

                if (fontStyleArray.Length > 0)
                {
                    foreach (var style in fontStyleArray)
                    {
                        if (style == "Bold")
                            font = new Font(font.FontFamily, font.Size, font.Style | FontStyle.Bold);

                        if (style == "Italic")
                            font = new Font(font.FontFamily, font.Size, font.Style | FontStyle.Italic);

                        if (style == "Underline")
                            font = new Font(font.FontFamily, font.Size, font.Style | FontStyle.Underline);

                        if (style == "Strikeout")
                            font = new Font(font.FontFamily, font.Size, font.Style | FontStyle.Strikeout);
                    }
                }
            }

            return font;
        }

        private double ParseDouble(string strValue)
        {
            return Convert.ToDouble(strValue.Trim().Replace(",", ".").Replace(".", ApplicationDecimalSeparator));
        }

        private void ClientRectangle(string locationFloat, string sizeF, StiComponent component)
        {
            if (!string.IsNullOrEmpty(locationFloat) && !string.IsNullOrEmpty(sizeF))
            {
                var locationArray = locationFloat.Split(',');
                var sizeArray = sizeF.Split(',');
                var x = "";
                var y = "";
                var width = "";
                var height = "";

                if (sizeArray.Length > 1)
                {
                    width = sizeArray[0];
                    height = sizeArray[1];
                }

                if (locationArray.Length > 1)
                {
                    x = locationArray[0];
                    y = locationArray[1];
                }

                if (reportUnit != "Pixels")
                {
                    component.ClientRectangle = new RectangleD(
                    ToHi(reportUnit, x),
                    ToHi(reportUnit, y),
                    ToHi(reportUnit, width),
                    ToHi(reportUnit, height));
                }
                else
                {
                    component.ClientRectangle = new RectangleD(
                    ParseDouble(x),
                    ParseDouble(y),
                    ParseDouble(width),
                    ParseDouble(height));
                }
            }
        }

        private Hashtable HtmlNameToColor = null;

        private Color ParseStyleColor(string colorAttribute)
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
                            ch == '8' || ch == '9' || ch == 'a' || ch == 'b' || ch == 'c' || ch == 'd' || ch == 'e'
                            || ch == 'f') sbc.Append(ch);
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
                        color = Color.FromArgb(0xFF, colorInt >> 16 & 0xFF, colorInt >> 8 & 0xFF, colorInt & 0xFF);
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
                else if (colorAttribute.IndexOf(',') != -1)
                {
                    #region Parse RGB values
                    string[] colors = colorAttribute.Split(new char[] { ',' });
                    if (colors.Length == 3)
                    {
                        color = Color.FromArgb(0xFF, int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]));
                    }
                    else if (colors.Length == 4)
                    {
                        color = Color.FromArgb(int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]), int.Parse(colors[3]));
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
                            Color value = Color.FromArgb(0xFF, colorInt >> 16 & 0xFF, colorInt >> 8 & 0xFF, colorInt & 0xFF);
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

        private StiPenStyle ParseBorderStyle(string styleName)
        {
            switch (styleName)
            {
                case "None":
                    return StiPenStyle.None;
                case "Dash":
                    return StiPenStyle.Dash;
                case "Dot":
                    return StiPenStyle.Dot;
                case "DashDot":
                    return StiPenStyle.DashDot;
                case "DashDotDot":
                    return StiPenStyle.DashDotDot;
                case "Double":
                    return StiPenStyle.Double;
            }
            return StiPenStyle.Solid;
        }

        private StiBorderSides BorderSide(string bandItemBorders)
        {
            var borders = bandItemBorders.Split(',');
            var borderSides = StiBorderSides.All;

            if (borders.Length == 1)
            {
                if (borders[0] == "None")
                    borderSides = StiBorderSides.None;

                else if (borders[0] == "All")
                    borderSides = StiBorderSides.All;

                else if (borders[0] == "Left")
                    borderSides = StiBorderSides.Left;

                else if (borders[0] == "Top")
                    borderSides = StiBorderSides.Top;

                else if (borders[0] == "Right")
                    borderSides = StiBorderSides.Right;

                else if (borders[0] == "Bottom")
                    borderSides = StiBorderSides.Bottom;
            }

            if (borders.Length > 1)
            {
                var isFirstBorder = true;

                foreach (var border in borders)
                {
                    var borderSide = border.Replace(" ", "");

                    if (borderSide == "Left")
                    {
                        if (isFirstBorder)
                            borderSides = StiBorderSides.Left;

                        else
                            borderSides |= StiBorderSides.Left;
                    }

                    if (borderSide == "Top")
                    {
                        if (isFirstBorder)
                            borderSides = StiBorderSides.Top;

                        else
                            borderSides |= StiBorderSides.Top;
                    }

                    if (borderSide == "Right")
                    {
                        if (isFirstBorder)
                            borderSides = StiBorderSides.Right;
                        else
                            borderSides |= StiBorderSides.Right;
                    }

                    if (borderSide == "Bottom")
                    {
                        if (isFirstBorder)
                            borderSides = StiBorderSides.Bottom;
                        else
                            borderSides |= StiBorderSides.Bottom;
                    }

                    if (isFirstBorder)
                        isFirstBorder = false;
                }
            }

            return borderSides;
        }

        private StiShape ShapeLineComponent(string lineDirection)
        {
            var component = new StiShape { ShapeType = new StiHorizontalLineShapeType() };

            if (lineDirection == "Horizontal")
                component = new StiShape { ShapeType = new StiHorizontalLineShapeType() };

            else if (lineDirection == "Slant")
                component = new StiShape { ShapeType = new StiDiagonalUpLineShapeType() };

            else if (lineDirection == "BackSlant")
                component = new StiShape { ShapeType = new StiDiagonalDownLineShapeType() };

            else if (lineDirection == "Vertical")
                component = new StiShape { ShapeType = new StiVerticalLineShapeType() };

            return component;
        }

        private void BarcodeType(StiBarCode component, string typeName)
        {
            if (typeName == "Code93") component.BarCodeType = new StiCode93BarCodeType();
            else if (typeName == "Code93Extended") component.BarCodeType = new StiCode93ExtBarCodeType();
            else if (typeName == "CodeMSI") component.BarCodeType = new StiMsiBarCodeType();
            else if (typeName == "EAN8") component.BarCodeType = new StiEAN8BarCodeType();
            else if (typeName == "EAN13") component.BarCodeType = new StiEAN13BarCodeType();
            else if (typeName == "EAN128") component.BarCodeType = new StiEAN128AutoBarCodeType();
            else if (typeName == "Postnet") component.BarCodeType = new StiPostnetBarCodeType();
            else if (typeName == "UPCA") component.BarCodeType = new StiUpcABarCodeType();
            else if (typeName == "UPCSupplement5") component.BarCodeType = new StiUpcSup5BarCodeType();
            else if (typeName == "UPCSupplement2") component.BarCodeType = new StiUpcSup2BarCodeType();
            else if (typeName == "Code39") component.BarCodeType = new StiCode39BarCodeType();
            else if (typeName == "Code39Extended") component.BarCodeType = new StiCode39ExtBarCodeType();
            else if (typeName == "Codabar") component.BarCodeType = new StiCodabarBarCodeType();
            else if (typeName == "Code11") component.BarCodeType = new StiCode11BarCodeType();
            else if (typeName == "Interleaved2of5") component.BarCodeType = new StiInterleaved2of5BarCodeType();
            else if (typeName == "PDF417") component.BarCodeType = new StiPdf417BarCodeType();
            else if (typeName == "DataMatrix") component.BarCodeType = new StiDataMatrixBarCodeType();
            else if (typeName == "QRCode") component.BarCodeType = new StiQRCodeBarCodeType();
            else if (typeName == "IntelligentMail") component.BarCodeType = new StiIntelligentMail4StateBarCodeType();
            else if (typeName == "DataMatrixGS1") component.BarCodeType = new StiGS1DataMatrixBarCodeType();
            else if (typeName == "ITF14") component.BarCodeType = new StiITF14BarCodeType();
            else if (typeName == "Pharmacode") component.BarCodeType = new StiPharmacodeBarCodeType();
            else if (typeName == "Code128") component.BarCodeType = new StiCode128AutoBarCodeType();
            else if (typeName == "UPCE0") component.BarCodeType = new StiUpcEBarCodeType();
            else if (typeName == "UPCE1") component.BarCodeType = new StiUpcEBarCodeType();
            else if (typeName == "Matrix2of5") component.BarCodeType = new StiStandard2of5BarCodeType();
            else component.BarCodeType = new StiCode93BarCodeType();
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
                message = string.Format("Node not supported: {0}.{1}", baseNodeName, nodeName);
            }
            else
            {
                message = string.Format("{0}", message1);
            }
            errorList.Add(message);
        }

        private string StringWordWrap(string textString, Font font, float availableWidth, float availableHeight, bool needHeightCutting,
            float measureCoeff)
        {
            var splittedString = textString.Split(' ');
            var correctStringCollection = new List<string>();
            var correctString = "";

            for (int ind = 0; ind < splittedString.Length; ind++)
            {
                string newStr;
                if (ind != splittedString.Length - 1)
                    newStr = $"{splittedString[ind]} ";

                else
                    newStr = splittedString[ind];

                if (MeasureSize(newStr, font).Width * measureCoeff < availableWidth)
                {
                    if (MeasureSize(correctString, font).Width * measureCoeff +
                        MeasureSize(newStr, font).Width * measureCoeff < availableWidth)
                    {
                        correctString += newStr;
                    }
                    else
                    {
                        var str = correctString;
                        correctStringCollection.Add(str);
                        correctString = newStr;
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(correctString))
                    {
                        var str = correctString;
                        correctStringCollection.Add(str);
                        correctString = "";
                    }

                    for (int i = 0; i < newStr.Length; i++)
                    {
                        if (MeasureSize(correctString, font).Width * measureCoeff + MeasureSize(newStr[i].ToString(), font).Width * measureCoeff
                            < availableWidth)
                        {
                            correctString += $"{newStr[i]}";
                        }
                        else
                        {
                            var str = correctString;
                            correctStringCollection.Add(str);
                            correctString = $"{newStr[i]}";
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(correctString))
            {
                var str = correctString;
                correctStringCollection.Add(str);
            }

            var finalString = "";
            var textHeight = 0d;

            for (int lineIndex = 0; lineIndex < correctStringCollection.Count; lineIndex++)
            {
                textHeight += MeasureSize(correctStringCollection[lineIndex], font).Height * measureCoeff;

                if (textHeight < availableHeight || !needHeightCutting)
                {
                    if (lineIndex != correctStringCollection.Count - 1)
                        finalString += $"{correctStringCollection[lineIndex]}\n";

                    else
                        finalString += correctStringCollection[lineIndex];
                }
            }

            return finalString;
        }

        private SizeF MeasureSize(string text, Font font)
        {
            Bitmap img = new Bitmap(1, 1);
            Graphics g = Graphics.FromImage(img);
            StringFormat sf = new StringFormat(StringFormat.GenericTypographic);
            g.PageUnit = GraphicsUnit.Point;
            var result = g.MeasureString(text, font, 10000000, sf);

            return result;
        }

        private static bool IsXml(byte[] data)
        {
            return (data != null) && (data.Length > 11) && (IsXml2(data, 0) || (data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF
                && IsXml2(data, 3)));
        }

        private static bool IsXml2(byte[] data, int offset)
        {
            return data[offset] == 0x3c && data[offset + 1] == 0x3f && data[offset + 2] == 0x78 && data[offset + 3] == 0x6d
                && data[offset + 4] == 0x6c;
        }
        #endregion

        #region Methods.Import
        public static StiImportResult Import(byte[] bytes)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = StiCultureInfo.GetEN(false);

                ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

                var report = new StiReport();
                var errors = new List<string>();

                if (!IsXml(bytes))
                {
                    errors.Add("The file is in an old format.To import this file please use an external utility:" +
                        "\r\nhttps://github.com/stimulsoft/Importing.Tools/tree/master/Import.XtraReports");
                    report = null;
                }
                else
                {
                    var helper = new StiDevExpressHelper();
                    helper.ProcessFile(bytes, report, errors);
                }
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
