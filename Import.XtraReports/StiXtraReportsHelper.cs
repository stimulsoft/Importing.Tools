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

using System.Globalization;
using System.Collections;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using DevExpress.XtraPrinting.Shape;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.BarCodes;
using Stimulsoft.Report.Components.ShapeTypes;
using Stimulsoft.Report.Components.TextFormats;
using Stimulsoft.Report.Dictionary;
using DevExpress.XtraReports.UI;

namespace Import.XtraReports
{
    public class StiXtraReportsHelper
    {
        #region Const
        private const string datasetName = "DataSetName";
        #endregion

        #region Fields
        private double totalHeight = 0;
        private StiReport report = null;
        private ReportUnit reportUnit = ReportUnit.HundredthsOfAnInch;
        private int detailLevel = 0;
        private Hashtable fields = new Hashtable();
        private string currentDataSourceName = string.Empty;
        #endregion

        #region Methods
        public StiReport Convert(string fileXtraReports)
        {
            CultureInfo currentCulture = Application.CurrentCulture;
            try
            {
                Application.CurrentCulture = new CultureInfo("en-US", false);

                report = new StiReport();
                report.Pages.Clear();

                XtraReport xtraReport = new XtraReport();
                xtraReport.LoadLayout(fileXtraReports);

                detailLevel = 0;
                currentDataSourceName = xtraReport.DataMember;
                reportUnit = xtraReport.ReportUnit;

                if (reportUnit == ReportUnit.TenthsOfAMillimeter)
                    report.ReportUnit = StiReportUnitType.Millimeters;
                else
                    report.ReportUnit = StiReportUnitType.HundredthsOfInch;
                
                ReadPage(xtraReport, report);

                foreach (StiPage page in report.Pages)
                {
                    StiComponentsCollection comps = page.GetComponents();
                    foreach (StiComponent comp in comps)
                    {
                        comp.Page = page;
                    }

                    page.LargeHeightFactor = 2;
                    page.LargeHeight = true;
                }


                //create datasources and relations, variables
                foreach (DictionaryEntry de in fields)
                {
                    string[] parts = ((string)de.Key).Split(new char[] { '.' });
                    if (parts.Length >= 2)
                    {
                        StiDataSource ds = report.Dictionary.DataSources[parts[0]];
                        if (ds == null)
                        {
                            ds = new StiDataTableSource();
                            ds.Name = parts[0];
                            ds.Alias = parts[0];
                            (ds as StiDataTableSource).NameInSource = datasetName;
                            ds.Columns.Add(new StiDataColumn("id"));
                            report.Dictionary.DataSources.Add(ds);
                        }

                        int pos = 1;
                        while (pos < parts.Length - 1)
                        {
                            string dsName = parts[pos];
                            if (dsName.StartsWith(ds.Name))
                            {
                                dsName = dsName.Substring(ds.Name.Length);
                            }

                            StiDataSource childSource = report.Dictionary.DataSources[dsName];
                            if (childSource == null)
                            {
                                childSource = new StiDataTableSource();
                                childSource.Name = dsName;
                                childSource.Alias = dsName;
                                (childSource as StiDataTableSource).NameInSource = datasetName;
                                childSource.Columns.Add(new StiDataColumn("id"));
                                report.Dictionary.DataSources.Add(childSource);
                            }
                            StiDataRelation relation = ds.GetChildRelations()[parts[pos]];
                            if (relation == null)
                            {
                                relation = new StiDataRelation(parts[pos], ds, childSource, new string[1] { "id" }, new string[1] { "id" });
                                report.Dictionary.Relations.Add(relation);
                            }
                            ds = childSource;
                            pos++;
                        }

                        if (ds.Columns[parts[pos]] == null)
                        {
                            StiDataColumn column = new StiDataColumn();
                            column.Name = parts[pos];
                            ds.Columns.Add(column);
                        }
                    }
                    else if (parts.Length == 1)
                    {
                        StiVariable varr = report.Dictionary.Variables[parts[0]];
                        if (varr == null)
                        {
                            varr = new StiVariable();
                            varr.Name = parts[0];
                            varr.Alias = parts[0];
                            report.Dictionary.Variables.Add(varr);
                        }
                    }
                }

                return report;
            }
            finally
            {
                Application.CurrentCulture = currentCulture;
            }
        }
        #endregion

        #region Methods.Convert
        private double ReadValueFrom(XtraReportBase xtraReport, double value)
        {
            if (reportUnit == ReportUnit.TenthsOfAMillimeter)
                return value / 10;
            else
                return value;
        }

        private StiPenStyle ConvertBorderDashStyle(DevExpress.XtraPrinting.BorderDashStyle style)
        {
            switch (style)
            {
                case DevExpress.XtraPrinting.BorderDashStyle.Dash:
                    return StiPenStyle.Dash;

                case DevExpress.XtraPrinting.BorderDashStyle.DashDot:
                    return StiPenStyle.DashDot;

                case DevExpress.XtraPrinting.BorderDashStyle.DashDotDot:
                    return StiPenStyle.DashDotDot;

                case DevExpress.XtraPrinting.BorderDashStyle.Dot:
                    return StiPenStyle.Dot;

                case DevExpress.XtraPrinting.BorderDashStyle.Double:
                    return StiPenStyle.Double;

                default:
                    return StiPenStyle.Solid;
            }
        }

        private StiPenStyle ConvertBorderDashStyle(DashStyle style)
        {
            switch (style)
            {
                case DashStyle.Dash:
                    return StiPenStyle.Dash;

                case DashStyle.DashDot:
                    return StiPenStyle.DashDot;

                case DashStyle.DashDotDot:
                    return StiPenStyle.DashDotDot;

                case DashStyle.Dot:
                    return StiPenStyle.Dot;

                default:
                    return StiPenStyle.Solid;
            }
        }

        private StiBorderSides ConvertBorderSide(DevExpress.XtraPrinting.BorderSide side)
        {
            if (side == DevExpress.XtraPrinting.BorderSide.All)
                return StiBorderSides.All;
            if (side == DevExpress.XtraPrinting.BorderSide.None)
                return StiBorderSides.None;

            StiBorderSides sides = 0;
            if ((side & DevExpress.XtraPrinting.BorderSide.Left) > 0)
                sides |= StiBorderSides.Left;
            if ((side & DevExpress.XtraPrinting.BorderSide.Right) > 0)
                sides |= StiBorderSides.Right;
            if ((side & DevExpress.XtraPrinting.BorderSide.Top) > 0)
                sides |= StiBorderSides.Top;
            if ((side & DevExpress.XtraPrinting.BorderSide.Bottom) > 0)
                sides |= StiBorderSides.Bottom;

            return sides;
        }
        #endregion

        #region Methods.Read.Properties
        private void ReadFont(XRControl xtraControl, StiComponent comp)
        {
            if (comp is IStiFont)
            {
                IStiFont fontComp = comp as IStiFont;
                fontComp.Font = xtraControl.GetEffectiveFont();
            }
        }

        private void ReadBorder(XRControl xtraControl, StiComponent comp)
        {
            if (comp is IStiBorder)
            {
                IStiBorder borderComp = comp as IStiBorder;
                borderComp.Border.Color = xtraControl.GetEffectiveBorderColor();
                borderComp.Border.Style = ConvertBorderDashStyle(xtraControl.GetEffectiveBorderDashStyle());
                borderComp.Border.Side = ConvertBorderSide(xtraControl.GetEffectiveBorders());
                borderComp.Border.Size = xtraControl.GetEffectiveBorderWidth();
            }
        }

        private void ReadBrush(XRControl xtraControl, StiComponent comp)
        {
            if (comp is IStiBrush)
            {
                IStiBrush brushComp = comp as IStiBrush;
                brushComp.Brush = new StiSolidBrush(xtraControl.GetEffectiveBackColor());
            }
        }

        private void ReadTextBrush(XRControl xtraControl, StiComponent comp)
        {
            if (comp is IStiTextBrush)
            {
                IStiTextBrush brushComp = comp as IStiTextBrush;
                brushComp.TextBrush = new StiSolidBrush(xtraControl.GetEffectiveForeColor());
            }
        }

        private void ReadTextAlignment(XRControl xtraControl, StiText text)
        {
            DevExpress.XtraPrinting.TextAlignment textAlignment = DevExpress.XtraPrinting.TextAlignment.TopCenter;
            if (xtraControl is XRLabel)
                textAlignment = ((XRLabel)xtraControl).GetEffectiveTextAlignment();
            if (xtraControl is XRPageInfo)
                textAlignment = ((XRPageInfo)xtraControl).GetEffectiveTextAlignment();
            if (xtraControl is XRCheckBox)
                textAlignment = ((XRCheckBox)xtraControl).GetEffectiveTextAlignment();

            switch (textAlignment)
            {
                case DevExpress.XtraPrinting.TextAlignment.BottomCenter:
                    text.HorAlignment = StiTextHorAlignment.Center;
                    text.VertAlignment = StiVertAlignment.Bottom;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.BottomJustify:
                    text.HorAlignment = StiTextHorAlignment.Width;
                    text.VertAlignment = StiVertAlignment.Bottom;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.BottomLeft:
                    text.HorAlignment = StiTextHorAlignment.Left;
                    text.VertAlignment = StiVertAlignment.Bottom;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.BottomRight:
                    text.HorAlignment = StiTextHorAlignment.Right;
                    text.VertAlignment = StiVertAlignment.Bottom;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.MiddleCenter:
                    text.HorAlignment = StiTextHorAlignment.Center;
                    text.VertAlignment = StiVertAlignment.Center;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.MiddleJustify:
                    text.HorAlignment = StiTextHorAlignment.Width;
                    text.VertAlignment = StiVertAlignment.Center;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.MiddleLeft:
                    text.HorAlignment = StiTextHorAlignment.Left;
                    text.VertAlignment = StiVertAlignment.Center;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.MiddleRight:
                    text.HorAlignment = StiTextHorAlignment.Right;
                    text.VertAlignment = StiVertAlignment.Center;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.TopCenter:
                    text.HorAlignment = StiTextHorAlignment.Center;
                    text.VertAlignment = StiVertAlignment.Top;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.TopJustify:
                    text.HorAlignment = StiTextHorAlignment.Width;
                    text.VertAlignment = StiVertAlignment.Top;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.TopLeft:
                    text.HorAlignment = StiTextHorAlignment.Left;
                    text.VertAlignment = StiVertAlignment.Top;
                    break;

                case DevExpress.XtraPrinting.TextAlignment.TopRight:
                    text.HorAlignment = StiTextHorAlignment.Right;
                    text.VertAlignment = StiVertAlignment.Top;
                    break;
            }
        }        
        #endregion

        #region Methods.Read.Bands
        private void ReadBand(Band xtraBand, StiBand band)
        {
            ReadComp(xtraBand, band);

            #region IStiPageBreak
            IStiPageBreak pageBreak = band as IStiPageBreak;
            if (pageBreak != null)
            {
                pageBreak.NewPageBefore = xtraBand.PageBreak == PageBreak.BeforeBand;
                pageBreak.NewPageAfter = xtraBand.PageBreak == PageBreak.AfterBand;
            }
            #endregion

            ProcessControls(xtraBand.Controls, band);

            totalHeight += band.Height;
        }

        private void ProcessControls(XRControlCollection controls, StiContainer parent)
        {
            foreach (XRControl control in controls)
            {
                if (control is XRLabel)
                    ReadLabel(control as XRLabel, parent);
                else if (control is XRPageInfo)
                    ReadPageInfo(control as XRPageInfo, parent);
                else if (control is XRCheckBox)
                    ReadCheckBox(control as XRCheckBox, parent);
                else if (control is XRPictureBox)
                    ReadPictureBox(control as XRPictureBox, parent);
                else if (control is XRRichText)
                    ReadRichText(control as XRRichText, parent);
                else if (control is XRPanel)
                    ReadPanel(control as XRPanel, parent);
                else if (control is XRLine)
                    ReadLine(control as XRLine, parent);
                else if (control is XRShape)
                    ReadShape(control as XRShape, parent);
                else if (control is XRBarCode)
                    ReadBarCode(control as XRBarCode, parent);
                else if (control is XRTable)
                    ReadTable(control as XRTable, parent);
                else if (control is XRTableRow)
                    ReadTableRow(control as XRTableRow, parent);
            }
        }

        private void ReadReportHeaderBand(XtraReportBase xtraReport, ReportHeaderBand xtraBand, StiPage page)
        {
            StiBand band = null;
            if (detailLevel == 0)
            {
                band = new StiReportTitleBand();
            }
            else
            {
                band = new StiHeaderBand();
                (band as StiHeaderBand).PrintOnAllPages = false;
            }
            page.Components.Add(band);

            ReadBand(xtraBand, band);

        }

        private void ReadReportFooterBand(XtraReportBase xtraReport, ReportFooterBand xtraBand, StiPage page)
        {
            StiBand band = null;
            if (detailLevel == 0)
            {
                band = new StiReportSummaryBand();
            }
            else
            {
                band = new StiFooterBand();
                (band as StiFooterBand).PrintOnAllPages = false;
            }
            page.Components.Add(band);

            ReadBand(xtraBand, band);
        }

        private void ReadDetailBand(XtraReportBase xtraReport, DetailBand xtraBand, StiPage page)
        {
            StiDataBand band = new StiDataBand();
            page.Components.Add(band);

            ReadBand(xtraBand, band);

            band.EvenStyle = xtraBand.EvenStyleName;
            band.OddStyle = xtraBand.OddStyleName;

            string dsName = ParseDatasourceName(xtraReport.DataMember, true);
            band.DataSourceName = dsName;
            band.DataRelationName = ParseRelationName(xtraReport.DataMember);
        }

        private void ReadPageHeaderBand(XtraReportBase xtraReport, PageHeaderBand xtraBand, StiPage page)
        {
            StiHeaderBand band = new StiHeaderBand();
            band.PrintOnAllPages = true;
            page.Components.Add(band);

            ReadBand(xtraBand, band);
        }

        private void ReadPageFooterBand(XtraReportBase xtraReport, PageFooterBand xtraBand, StiPage page)
        {
            StiFooterBand band = new StiFooterBand();
            band.PrintOnAllPages = true;
            page.Components.Add(band);

            ReadBand(xtraBand, band);
        }

        private void ReadTopMarginBand(XtraReportBase xtraReport, TopMarginBand xtraBand, StiPage page)
        {
            StiPageHeaderBand band = new StiPageHeaderBand();
            page.Components.Add(band);

            ReadBand(xtraBand, band);

            page.Margins.Top = page.Margins.Top - report.Unit.ConvertToHInches(band.Height);
        }

        private void ReadBottomMarginBand(XtraReportBase xtraReport, BottomMarginBand xtraBand, StiPage page)
        {
            StiPageFooterBand band = new StiPageFooterBand();
            page.Components.Add(band);

            ReadBand(xtraBand, band);

            page.Margins.Bottom = page.Margins.Bottom - report.Unit.ConvertToHInches(band.Height);
        }

        private void ReadGroupHeaderBand(XtraReportBase xtraReport, GroupHeaderBand xtraBand, StiPage page)
        {
            StiGroupHeaderBand band = new StiGroupHeaderBand();
            page.Components.Add(band);

            ReadBand(xtraBand, band);
        }

        private void ReadGroupFooterBand(XtraReportBase xtraReport, GroupFooterBand xtraBand, StiPage page)
        {
            StiGroupFooterBand band = new StiGroupFooterBand();
            page.Components.Add(band);

            ReadBand(xtraBand, band);
        }

        private void ReadDetailReportBand(XtraReportBase xtraReport, DetailReportBand band, StiPage page)
        {
            detailLevel++;
            StiPage tempPage = new StiPage(report);

            string storeDataSourceName = currentDataSourceName;
            currentDataSourceName = band.DataMember;

            ProcessBands(band, band.Controls, tempPage);

            //StiDataBand masterBand = page.Components[page.Components.Count - 1] as StiDataBand;
            //if (masterBand != null)
            //{
            //    foreach (StiComponent comp in tempPage.Components)
            //    {
            //        StiDataBand dataBand = comp as StiDataBand;
            //        if (dataBand != null)
            //        {
            //            dataBand.MasterComponent = masterBand;
            //        }
            //    }
            //}

            //page.Components.AddRange(tempPage.Components);
            SortBands(tempPage);
            page.Components.Add(tempPage);

            currentDataSourceName = storeDataSourceName;

            detailLevel--;
        }

        private void SortBands(StiPage input)
        {
            int counterPageHeader = 1000;
            int counterPageFooter = 2000;
            int counterHeader = 3000;
            int counterGroupHeader = 4000;
            int counterData = 5000;
            int counterDetail = 6000;
            int counterGroupFooter = 7000;
            int counterFooter = 8000;
            int counterOther = 10000;

            StiDataBand masterDataBand = null;

            Dictionary<int, StiComponent> dict = new Dictionary<int, StiComponent>();
            List<int> keys = new List<int>();

            foreach (StiComponent comp in input.Components)
            {
                int counter = 0;
                if (comp is StiPageHeaderBand)
                {
                    counter = counterPageHeader++;
                }
                if (comp is StiPageFooterBand)
                {
                    counter = counterPageFooter++;
                }
                if (comp is StiGroupHeaderBand)
                {
                    counter = counterGroupHeader++;
                }
                else if (comp is StiGroupFooterBand)
                {
                    counter = counterGroupFooter++;
                }
                else if (comp is StiHeaderBand)
                {
                    counter = counterHeader++;
                }
                else if (comp is StiFooterBand)
                {
                    counter = counterFooter++;
                }
                else if (comp is StiDataBand)
                {
                    counter = counterData++;
                    masterDataBand = comp as StiDataBand;
                }
                else if (comp is StiPage)
                {
                    counter = counterDetail++;
                }
                else if (comp is StiChildBand)
                {
                    counter = counterData++;
                }
                else
                {
                    counter = counterOther++;
                }

                dict.Add(counter, comp);
                keys.Add(counter);
            }

            keys.Sort();

            StiComponentsCollection comps = new StiComponentsCollection(input);
            foreach (int key in keys)
            {
                StiComponent comp = dict[key];
                if (comp is StiPage)
                {
                    StiPage cont = comp as StiPage;
                    foreach (StiComponent comp2 in cont.Components)
                    {
                        StiDataBand band = comp2 as StiDataBand;
                        if ((band != null) && (band.MasterComponent == null))
                        {
                            band.MasterComponent = masterDataBand;
                        }
                        comps.Add(comp2);
                    }
                }
                else
                {
                    comps.Add(comp);
                }
            }
            input.Components = comps;
        }
        #endregion

        #region Methods.Read.Components
        private void ReadPage(XtraReport xtraReport, StiReport report)
        {
            StiPage page = new StiPage(report);
            report.Pages.Add(page);

            page.PageWidth = ReadValueFrom(xtraReport, xtraReport.PageWidth);
            page.PageHeight = ReadValueFrom(xtraReport, xtraReport.PageHeight);
            page.Brush = new StiSolidBrush(xtraReport.PageColor);

            page.Orientation = xtraReport.Landscape ? StiPageOrientation.Landscape : StiPageOrientation.Portrait;

            page.PaperSize = xtraReport.PaperKind;

            page.Margins.Left = ReadValueFrom(xtraReport, xtraReport.Margins.Left);
            page.Margins.Right = ReadValueFrom(xtraReport, xtraReport.Margins.Right);
            page.Margins.Top = ReadValueFrom(xtraReport, xtraReport.Margins.Top);
            page.Margins.Bottom = ReadValueFrom(xtraReport, xtraReport.Margins.Bottom);

            report.ScriptLanguage = xtraReport.ScriptLanguage == DevExpress.XtraReports.ScriptLanguage.CSharp ? StiReportLanguageType.CSharp : StiReportLanguageType.VB;
            //report.ReferencedAssemblies = xtraReport.ScriptReferences;

            ReadBrush(xtraReport, page);
            ReadBorder(xtraReport, page);

            ProcessBands(xtraReport, xtraReport.Bands, page);

            SortBands(page);

            if (totalHeight * 1.2 > page.Height)
            {
                page.LargeHeight = true;
                page.LargeHeightFactor = 2;

                while (totalHeight * 1.2 > page.Height * page.LargeHeightFactor)
                {
                    page.LargeHeightFactor++;
                }
            }
        }

        private void ProcessBands(XtraReportBase xtraReport, XRControlCollection controls, StiPage page)
        {
            foreach (XRControl band in controls)
            {
                if (band is GroupHeaderBand)
                    ReadGroupHeaderBand(xtraReport, band as GroupHeaderBand, page);
                else if (band is GroupFooterBand)
                    ReadGroupFooterBand(xtraReport, band as GroupFooterBand, page);
                else if (band is DetailBand)
                    ReadDetailBand(xtraReport, band as DetailBand, page);
                else if (band is ReportHeaderBand)
                    ReadReportHeaderBand(xtraReport, band as ReportHeaderBand, page);
                else if (band is ReportFooterBand)
                    ReadReportFooterBand(xtraReport, band as ReportFooterBand, page);
                else if (band is PageHeaderBand)
                    ReadPageHeaderBand(xtraReport, band as PageHeaderBand, page);
                else if (band is PageFooterBand)
                    ReadPageFooterBand(xtraReport, band as PageFooterBand, page);
                else if (band is TopMarginBand)
                    ReadTopMarginBand(xtraReport, band as TopMarginBand, page);
                else if (band is BottomMarginBand)
                    ReadBottomMarginBand(xtraReport, band as BottomMarginBand, page);
                else if (band is DetailReportBand)
                    ReadDetailReportBand(xtraReport, band as DetailReportBand, page);
            }
        }

        private void ReadComp(XRControl xtraControl, StiComponent comp)
        {
            comp.Name = xtraControl.Name;

            comp.Left = ReadValueFrom(xtraControl.Report as XtraReportBase, xtraControl.LeftF);
            comp.Top = ReadValueFrom(xtraControl.Report as XtraReportBase, xtraControl.TopF);
            comp.Width = ReadValueFrom(xtraControl.Report as XtraReportBase, xtraControl.WidthF);
            comp.Height = ReadValueFrom(xtraControl.Report as XtraReportBase, xtraControl.HeightF);

            comp.CanGrow = xtraControl.CanGrow;
            comp.CanShrink = xtraControl.CanShrink;
            comp.Enabled = xtraControl.Visible;

            comp.Bookmark.Value = xtraControl.Bookmark;
            comp.Hyperlink.Value = xtraControl.NavigateUrl;
            comp.Tag.Value = xtraControl.Tag != null ? xtraControl.Tag.ToString() : "";

            comp.AfterPrintEvent.Script = xtraControl.Scripts.OnAfterPrint;
            comp.BeforePrintEvent.Script = xtraControl.Scripts.OnBeforePrint;

            comp.Enabled = xtraControl.Visible;
            comp.ComponentStyle = xtraControl.StyleName;
        }

        private void ReadPageInfo(XRPageInfo xtraControl, StiContainer parent)
        {
            StiText text = new StiText();
            parent.Components.Add(text);

            ReadComp(xtraControl, text);

            ReadBrush(xtraControl, text);
            ReadTextBrush(xtraControl, text);
            ReadBorder(xtraControl, text);
            ReadFont(xtraControl, text);   
            ReadTextAlignment(xtraControl, text);
           
            #region ProcessDuplicates
            switch (xtraControl.ProcessDuplicates)
            {
                case ValueSuppressType.Leave:
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.None;
                    break;

                case ValueSuppressType.Suppress:
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.Hide;
                    break;

                case ValueSuppressType.SuppressAndShrink:
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.Merge;
                    break;
            }           
            #endregion

            #region ProcessNullValues
            text.HideZeros = xtraControl.ProcessNullValues == ValueSuppressType.Suppress || xtraControl.ProcessNullValues == ValueSuppressType.SuppressAndShrink;
            #endregion            

            text.Text.Value = xtraControl.Text;

            #region PageInfo
            switch (xtraControl.PageInfo)
            {
                case DevExpress.XtraPrinting.PageInfo.DateTime:
                    text.Text.Value = "{Today}";
                    text.TextFormat = new StiDateFormatService();
                    break;

                case DevExpress.XtraPrinting.PageInfo.Number:
                    text.Text.Value = "{PageNumber}";
                    break;

                case DevExpress.XtraPrinting.PageInfo.NumberOfTotal:
                    text.Text.Value = "{PageNumber}/{TotalPageCount}";
                    break;

                case DevExpress.XtraPrinting.PageInfo.RomLowNumber:
                    text.Text.Value = "{ToLowerCase(Roman(PageNumber))}";
                    break;

                case DevExpress.XtraPrinting.PageInfo.RomHiNumber:
                    text.Text.Value = "{ToLowerCase(Roman(PageNumber))}";
                    break;

                case DevExpress.XtraPrinting.PageInfo.UserName:
                    text.Text.Value = "{ReportAuthor}";
                    break;
            }
            #endregion
        }

        private void ReadLabel(XRLabel xtraControl, StiContainer parent)
        {
            StiText text = new StiText();
            parent.Components.Add(text);

            ReadComp(xtraControl, text);

            text.Angle = xtraControl.Angle;
            text.AutoWidth = xtraControl.AutoWidth;

            ReadBrush(xtraControl, text);
            ReadTextBrush(xtraControl, text);
            ReadBorder(xtraControl, text);
            ReadFont(xtraControl, text);
            ReadTextAlignment(xtraControl, text);

            text.WordWrap = xtraControl.WordWrap;

            #region ProcessDuplicates
            switch (xtraControl.ProcessDuplicates)
            {
                case ValueSuppressType.Leave:
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.None;
                    break;

                case ValueSuppressType.Suppress:
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.Hide;
                    break;

                case ValueSuppressType.SuppressAndShrink:
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.Merge;
                    break;
            }
            #endregion

            #region ProcessNullValues
            text.HideZeros = xtraControl.ProcessNullValues == ValueSuppressType.Suppress || xtraControl.ProcessNullValues == ValueSuppressType.SuppressAndShrink;
            #endregion

            text.Text.Value = xtraControl.Text;

            foreach (XRBinding bind in xtraControl.DataBindings)
            {
                switch (bind.PropertyName)
                {
                    case "Text":
                        text.Text.Value = "{" + ParseExpression(bind.DataMember) + "}";
                        break;

                    case "Bookmark":
                        text.Bookmark.Value = "{" + ParseExpression(bind.DataMember) + "}";
                        break;

                    case "Hyperlink":
                        text.Hyperlink.Value = "{" + ParseExpression(bind.DataMember) + "}";
                        break;
                }
            }
        }

        private void ReadCheckBox(XRCheckBox xtraControl, StiContainer parent)
        {
            #region Read CheckBox
            StiCheckBox check = new StiCheckBox();
            parent.Components.Add(check);

            ReadComp(xtraControl, check);

            ReadBrush(xtraControl, check);
            ReadTextBrush(xtraControl, check);
            ReadBorder(xtraControl, check);
            ReadFont(xtraControl, check);

            check.Width = check.Height;
            #endregion

            #region Read Text
            StiText text = new StiText();
            parent.Components.Add(text);

            ReadComp(xtraControl, text);
            text.Name += "_Text";
            text.Left += check.Width;
            text.Width -= check.Width;

            ReadBrush(xtraControl, text);
            ReadTextBrush(xtraControl, text);
            ReadBorder(xtraControl, text);
            ReadFont(xtraControl, text);
            ReadTextAlignment(xtraControl, text);

            text.Text.Value = xtraControl.Text;
            #endregion

        }

        private void ReadPictureBox(XRPictureBox xtraControl, StiContainer parent)
        {
            StiImage image = new StiImage();
            parent.Components.Add(image);

            ReadComp(xtraControl, image);

            ReadBrush(xtraControl, image);
            ReadTextBrush(xtraControl, image);
            ReadBorder(xtraControl, image);
            ReadFont(xtraControl, image);

            image.ImageURL.Value = xtraControl.ImageUrl;
            image.Image = xtraControl.Image;

            #region Sizing
            switch (xtraControl.Sizing)
            {
                case DevExpress.XtraPrinting.ImageSizeMode.AutoSize:
                case DevExpress.XtraPrinting.ImageSizeMode.StretchImage:
                    image.Stretch = true;
                    break;

                case DevExpress.XtraPrinting.ImageSizeMode.CenterImage:
                    image.HorAlignment = StiHorAlignment.Center;
                    image.VertAlignment = StiVertAlignment.Center;
                    break;

                case DevExpress.XtraPrinting.ImageSizeMode.ZoomImage:
                    image.AspectRatio = true;
                    image.Stretch = true;
                    break;
            }
            #endregion

        }

        private void ReadRichText(XRRichText xtraControl, StiContainer parent)
        {
            StiRichText richBox = new StiRichText();
            parent.Components.Add(richBox);

            ReadComp(xtraControl, richBox);

            ReadBrush(xtraControl, richBox);
            ReadTextBrush(xtraControl, richBox);
            ReadBorder(xtraControl, richBox);
            ReadFont(xtraControl, richBox);

        }

        private void ReadLine(XRLine xtraControl, StiContainer parent)
        {
            StiShape shape = new StiShape();
            parent.Components.Add(shape);

            ReadComp(xtraControl, shape);

            ReadBrush(xtraControl, shape);
            ReadTextBrush(xtraControl, shape);
            ReadBorder(xtraControl, shape);

            #region LineDirection
            switch (xtraControl.LineDirection)
            {
                case LineDirection.BackSlant:
                    shape.ShapeType = new StiDiagonalDownLineShapeType();
                    break;

                case LineDirection.Slant:
                    shape.ShapeType = new StiDiagonalUpLineShapeType();
                    break;

                case LineDirection.Horizontal:
                    shape.ShapeType = new StiHorizontalLineShapeType();
                    break;

                case LineDirection.Vertical:
                    shape.ShapeType = new StiVerticalLineShapeType();
                    break;
            }
            #endregion

            shape.Style = ConvertBorderDashStyle(xtraControl.LineStyle);
            shape.Size = xtraControl.LineWidth;
            shape.BorderColor = xtraControl.ForeColor;


        }

        private void ReadShape(XRShape xtraControl, StiContainer parent)
        {
            StiShape shape = new StiShape();
            parent.Components.Add(shape);

            ReadComp(xtraControl, shape);
            ReadBorder(xtraControl, shape);

            #region Shape
            if (xtraControl.Shape is ShapeRectangle)
                shape.ShapeType = new StiRectangleShapeType();
            else if (xtraControl.Shape is ShapeArrow)
                shape.ShapeType = new StiArrowShapeType();
            else if (xtraControl.Shape is ShapeCross)
                shape.ShapeType = new StiPlusShapeType();
            else if (xtraControl.Shape is ShapeEllipse)
                shape.ShapeType = new StiOvalShapeType();
            else if (xtraControl.Shape is ShapeLine)
                shape.ShapeType = new StiHorizontalLineShapeType();
            #endregion

            shape.Size = xtraControl.LineWidth;
            shape.BorderColor = xtraControl.ForeColor;
            shape.Brush = new StiSolidBrush(xtraControl.FillColor);


        }

        private void ReadBarCode(XRBarCode xtraControl, StiContainer parent)
        {
            StiBarCode barCode = new StiBarCode();
            parent.Components.Add(barCode);

            ReadComp(xtraControl, barCode);
            ReadBorder(xtraControl, barCode);
            ReadFont(xtraControl, barCode);
            ReadTextBrush(xtraControl, barCode);
        }

        private void ReadPanel(XRPanel xtraControl, StiContainer parent)
        {
            StiPanel panel = new StiPanel();
            parent.Components.Add(panel);

            ReadComp(xtraControl, panel);

            ProcessControls(xtraControl.Controls, panel);
        }

        private void ReadTable(XRTable xtraControl, StiContainer parent)
        {
            StiPanel table = new StiPanel();
            parent.Components.Add(table);

            ReadComp(xtraControl, table);

            ProcessControls(xtraControl.Controls, table);
        }

        private void ReadTableRow(XRTableRow xtraControl, StiContainer parent)
        {
            StiPanel tableRow = new StiPanel();
            parent.Components.Add(tableRow);

            ReadComp(xtraControl, tableRow);

            ProcessControls(xtraControl.Controls, tableRow);
        }
        #endregion

        #region Methods.ParseExpression
        private string ParseExpression(string input)
        {
            string output = input;

            fields[input] = input;

            string[] parts = input.Split(new char[] { '.' });
            if (parts.Length > 2)
            {
                if (input.StartsWith(currentDataSourceName))
                {
                    string dsName = ParseDatasourceName(input, false);
                    output = dsName + "." + parts[parts.Length - 1];
                }
            }

            return output;
        }

        private string ParseDatasourceName(string name, bool full)
        {
            string[] parts = name.Split(new char[] { '.' });
            if (parts.Length > (1 + (full ? 0 : 1)))
            {
                for (int index = 1; index < (parts.Length - (full ? 0 : 1)); index++)
                {
                    string dsName = parts[index];
                    if (dsName.StartsWith(parts[index - 1]))
                    {
                        dsName = dsName.Substring(parts[index - 1].Length);
                    }
                    parts[index] = dsName;
                }
            }
            return (full ? parts[parts.Length - 1] : parts[parts.Length - 2]);
        }

        private string ParseRelationName(string name)
        {
            string[] parts = name.Split(new char[] { '.' });
            if (parts.Length > 1)
            {
                return parts[parts.Length - 1];
            }
            return string.Empty;
        }
        #endregion
    }
}
