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
using System.Drawing.Printing;
using System.Globalization;
using System.Text;
using System.Threading;
using System.Xml;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.BarCodes;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components.TextFormats;
using Stimulsoft.Report.Components.Table;
using Stimulsoft.Base;

namespace Stimulsoft.Report.Import
{
    public class StiComponentOneReportsHelper
    {
        #region Fields
        List<string> errorList = null;

        ArrayList fields = new ArrayList();
        Hashtable fieldsNames = new Hashtable();

        string currentDataSourceName = "ds";
        //string mainDataSourceName = "ds";

        //bool pageSettingsUsed = false;
        //bool dataSourceUsed;

        Font mainFont = new Font("Arial", 8);

        internal bool SetLinked = true;
        #endregion

        #region ReservedWords
        private Hashtable _reservedWords = null;

        public Hashtable reservedWords
        {
            get
            {
                if (_reservedWords == null)
                {
                    _reservedWords = new Hashtable();
                    _reservedWords["vbcrlf"] = "vbcrlf";
                    _reservedWords["Page"] = "Page";
                    _reservedWords["Pages"] = "Pages";
                    _reservedWords["Now"] = "Now";
                    //_reservedWords["vbcrlf"] = "vbcrlf";
                    //_reservedWords["vbcrlf"] = "vbcrlf";
                }
                return _reservedWords;
            }
        }
        #endregion

        #region Root node
        public void ProcessRootNode(XmlNode rootNode, StiReport report, List<string> errorList, int pageIndex)
        {
            this.errorList = errorList;
            this.fields.Clear();
            this.fieldsNames = new Hashtable();
            report.ReportUnit = StiReportUnitType.HundredthsOfInch;
            StiPage page = report.Pages[pageIndex];
            currentDataSourceName = string.Format("ds{0}", pageIndex > 0 ? pageIndex.ToString() : "");

            page.Margins = new StiMargins(100, 100, 100, 100);

            //firts pass
            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Name":
                        page.Name = node.FirstChild.Value;
                        if (pageIndex == 0) report.ReportName = node.FirstChild.Value;
                        break;

                    case "DataSource":
                        //case "DataSources":
                        //dataSourceUsed = true;
                        ProcessDataSource(node, report);
                        //mainDataSourceName = currentDataSourceName;
                        break;

                    case "Sections":
                        ProcessSections(node, report, page);
                        break;

                    case "Fields":
                        ProcessFields(node, report, page);
                        break;

                    case "Layout":
                        ProcessLayout(node, report, page);
                        break;

                    case "Font":
                        StiText tempText = new StiText();
                        ProcessFont(node, tempText);
                        mainFont = tempText.Font;
                        break;

                    case "Groups":
                        //ignored on first pass
                        break;

                    //case "wwwwWidth":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(rootNode.Name, node.Name);
                        break;
                }
            }

            //second pass
            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Groups":
                        ProcessGroups(node, report, page);
                        break;
                }
            }

            #region Check DataSource
            //if (!dataSourceUsed)
            //{
            //    StiDataTableSource dataSource = new StiDataTableSource();
            //    dataSource.Name = "ds";
            //    dataSource.Alias = "ds";
            //    dataSource.NameInSource = "";
            //    report.Dictionary.DataSources.Add(dataSource);
            //}

            foreach (StiDataSource ds in report.Dictionary.DataSources)
            {
                string dsName = ds.Name + ".";
                foreach (string st in fields)
                {
                    if (st.Length > 0 && st.StartsWith(dsName))
                    {
                        ds.Columns.Add(new StiDataColumn(st.Replace(dsName, "")));
                    }
                }
            }
            #endregion

            #region Normalize bands order
            StiComponentsCollection oldComps = page.Components;
            StiComponentsCollection newComps = new StiComponentsCollection();
            if (oldComps.Count > 3) newComps.Add(oldComps[3]); //PageHeader
            if (oldComps.Count > 1) newComps.Add(oldComps[1]); //ReportTitle
            int indexBand = 5;
            while (indexBand < oldComps.Count)
            {
                newComps.Add(oldComps[indexBand]); //GroupHeader
                indexBand += 2;
            }
            if (oldComps.Count > 0) newComps.Add(oldComps[0]); //DataBand
            indexBand = oldComps.Count - 1;
            while (indexBand >= 6)
            {
                newComps.Add(oldComps[indexBand]); //GroupFooter
                indexBand -= 2;
            }
            if (oldComps.Count > 2) newComps.Add(oldComps[2]); //ReportSummary
            if (oldComps.Count > 4) newComps.Add(oldComps[4]); //PageFooter

            page.Components.Clear();
            page.Components.AddRange(newComps);
            #endregion

            foreach (StiPage page2 in report.Pages)
            {
                foreach (StiComponent comp in page2.GetComponents())
                {
                    comp.Page = page2;
                    if (SetLinked)
                    {
                        comp.Linked = true;
                    }
                }
            }

            page.DockToContainer();
            page.Correct();

            page.TitleBeforeHeader = true;
        }
        #endregion

        #region Layout
        private void ProcessLayout(XmlNode baseNode, StiReport report, StiPage page)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Orientation":
                        if (node.FirstChild.Value == "2") page.Orientation = StiPageOrientation.Landscape;
                        break;

                    case "CustomWidth":
                        page.PaperSize = PaperKind.Custom;
                        page.PageWidth = ToHi(node.FirstChild.Value);
                        break;

                    case "CustomHeight":
                        page.PaperSize = PaperKind.Custom;
                        page.PageHeight = ToHi(node.FirstChild.Value);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region DataSource
        private void ProcessDataSource(XmlNode baseNode, StiReport report)
        {
            StiOleDbDatabase dataBase = new StiOleDbDatabase();
            dataBase.Name = "db";
            dataBase.Alias = "db";
            StiOleDbSource dataSource = new StiOleDbSource();
            dataSource.Name = "ds";
            dataSource.Alias = "ds";
            dataSource.NameInSource = "db";

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ConnectionString":
                        string connectionString = node.FirstChild.Value;
                        connectionString = connectionString.Replace("|DataDirectory|", "");
                        dataBase.ConnectionString = connectionString;
                        break;
                    case "RecordSource":
                        if (node.FirstChild.Value.IndexOf(' ') == -1)
                        {
                            dataSource.SqlCommand = "select * from " + node.FirstChild.Value;
                        }
                        else
                        {
                            dataSource.SqlCommand = node.FirstChild.Value;
                        }
                        break;

                    //case "Type":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            report.Dictionary.Databases.Add(dataBase);
            report.Dictionary.DataSources.Add(dataSource);
        }
        #endregion

        #region Sections
        private void ProcessSections(XmlNode baseNode, StiReport report, StiPage page)
        {
            int index = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Section":
                        ProcessSection(node, report, page, index);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
                index++;
            }
        }

        private void ProcessSection(XmlNode baseNode, StiReport report, StiPage page, int index)
        {
            StiBand band = new StiDataBand();
            if (index == 1) band = new StiReportTitleBand();
            if (index == 2) band = new StiReportSummaryBand();
            if (index == 3) band = new StiPageHeaderBand();
            if (index == 4) band = new StiPageFooterBand();
            if (index > 4)
            {
                if (index % 2 == 1) band = new StiGroupHeaderBand();
                if (index % 2 == 0) band = new StiGroupFooterBand();
            }

            if (band is StiDataBand) (band as StiDataBand).DataSourceName = currentDataSourceName;

            band.CanGrow = true;
            band.CanShrink = false;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Name":
                        band.Name = node.FirstChild.Value;
                        break;

                    case "Height":
                        band.Height = ToHi(node.FirstChild.Value);
                        break;

                    case "Visible":
                        if (node.FirstChild.Value == "0") band.Enabled = false;
                        break;

                    case "Repeat":
                        if (node.FirstChild.Value == "-1")
                        {
                            IStiPrintOnAllPages printOnAll = band as IStiPrintOnAllPages;
                            if (printOnAll != null) printOnAll.PrintOnAllPages = true;
                        }
                        break;

                    case "BackColor":
                        band.Brush = new StiSolidBrush(ParseColor(node.FirstChild));
                        break;

                    case "OnPrint":
                        band.BeforePrintEvent.Script = "/* " + node.FirstChild.Value + " */";
                        break;

                    case "Type":
                        //ignore
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            page.Components.Add(band);
        }
        #endregion

        #region ReportParameters
        //private void ProcessReportParameters(XmlNode baseNode, StiReport report)
        //{
        //    foreach (XmlNode node in baseNode.ChildNodes)
        //    {
        //        switch (node.Name)
        //        {
        //            case "ReportParameter":
        //                ProcessReportParameter(node, report);
        //                break;

        //            default:
        //                ThrowError(baseNode.Name, node.Name);
        //                break;
        //        }
        //    }
        //}

        //private void ProcessReportParameter(XmlNode baseNode, StiReport report)
        //{
        //    StiVariable var = new StiVariable("Parameters", "Name", typeof(string));
        //    var.RequestFromUser = true;

        //    foreach (XmlNode node in baseNode.Attributes)
        //    {
        //        switch (node.Name)
        //        {
        //            case "Name":
        //                var.Name = node.Value;
        //                var.Alias = node.Value;
        //                break;
        //            case "Text":
        //                var.Alias = node.Value;
        //                break;
        //            case "Type":
        //                if (node.Value == "Integer") var.Type = typeof(Int32);
        //                break;

        //            default:
        //                ThrowError(baseNode.Name, "#" + node.Name);
        //                break;
        //        }
        //    }

        //    report.Dictionary.Variables[var.Name] = var;

        //    foreach (XmlNode node in baseNode.ChildNodes)
        //    {
        //        switch (node.Name)
        //        {
        //            case "Value":
        //                var.Value = node.ChildNodes[0].Value;
        //                break;

        //            case "AvailableValues":
        //                ProcessAvailableValues(node, report, var);
        //                break;                        

        //            default:
        //                ThrowError(baseNode.Name, node.Name);
        //                break;
        //        }
        //    }
        //}

        //private void ProcessAvailableValues(XmlNode baseNode, StiReport report, StiVariable var)
        //{
        //    foreach (XmlNode node in baseNode.Attributes)
        //    {
        //        switch (node.Name)
        //        {
        //            //case "Name":
        //            //    var.Name = node.Value;
        //            //    break;

        //            default:
        //                ThrowError(baseNode.Name, "#" + node.Name);
        //                break;
        //        }
        //    }

        //    foreach (XmlNode node in baseNode.ChildNodes)
        //    {
        //        switch (node.Name)
        //        {
        //            case "DataSource":
        //            case "DataSources":
        //                string store = currentDataSourceName;
        //                ProcessDataSources(node, report);
        //                currentDataSourceName = store;
        //                break;

        //            default:
        //                ThrowError(baseNode.Name, node.Name);
        //                break;
        //        }
        //    }
        //}
        #endregion

        #region Fields
        private void ProcessFields(XmlNode baseNode, StiReport report, StiPage page)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Field":
                        ProcessField(node, report, page);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessField(XmlNode baseNode, StiReport report, StiPage page)
        {
            bool isPicture = false;
            bool isBarcode = false;
            bool isCheckbox = false;
            bool isRtf = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == "Picture") isPicture = true;
                if (node.Name == "Barcode") isBarcode = true;
                if (node.Name == "CheckBox") isCheckbox = true;
                if (node.Name == "RTF") isRtf = true;
            }

            if (isPicture) ProcessPicture(baseNode, report, page);
            else if (isBarcode) ProcessBarcode(baseNode, report, page);
            else if (isCheckbox) ProcessCheckBox(baseNode, report, page);
            else if (isRtf) ProcessRtf(baseNode, report, page);
            else ProcessText(baseNode, report, page);
        }
        #endregion

        #region ProcessCommonComponentProperties
        private bool ProcessCommonComponentProperties(XmlNode node, StiComponent component)
        {
            switch (node.Name)
            {
                case "Name":
                    component.Name = node.FirstChild.Value;
                    return true;

                case "Left":
                    component.Left = ToHi(node.FirstChild.Value);
                    return true;
                case "Top":
                    component.Top = ToHi(node.FirstChild.Value);
                    return true;
                case "Width":
                    component.Width = ToHi(node.FirstChild.Value);
                    return true;
                case "Height":
                    component.Height = ToHi(node.FirstChild.Value);
                    return true;

                case "BorderStyle":
                    ProcessBorderStyle(node, component);
                    return true;

                case "BorderColor":
                    IStiBorder border = component as IStiBorder;
                    if (border != null)
                    {
                        border.Border.Color = ParseColor(node.FirstChild);
                    }
                    return true;

                case "CanGrow":
                    if (node.FirstChild.Value == "-1")
                    {
                        if (component is IStiCanGrow) (component as IStiCanGrow).CanGrow = true;
                    }
                    return true;

                case "CanShrink":
                    if (node.FirstChild.Value == "-1")
                    {
                        if (component is IStiCanShrink) (component as IStiCanShrink).CanShrink = true;
                    }
                    return true;


                //case "Type":
                //    //ignored
                //    return true;
            }
            return false;
        }
        #endregion

        #region Text
        private void ProcessText(XmlNode baseNode, StiReport report, StiPage page)
        {
            StiText component = new StiText();
            component.CanGrow = false;
            component.CanShrink = false;
            component.WordWrap = true;
            component.Font = mainFont;

            int sectionIndex = -1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Text":
                        component.Text = node.FirstChild.Value;
                        break;

                    case "Calculated":
                        if (node.FirstChild.Value == "-1") component.Text = "{" + ConvertExpression(component.Text, report) + "}";
                        break;

                    case "WordWrap":
                        if (node.FirstChild.Value == "0") component.WordWrap = false;
                        break;

                    case "Section":
                        sectionIndex = int.Parse(node.FirstChild.Value);
                        break;

                    case "ForeColor":
                        component.TextBrush = new StiSolidBrush(ParseColor(node.FirstChild));
                        break;

                    case "BackColor":
                        component.Brush = new StiSolidBrush(ParseColor(node.FirstChild));
                        break;

                    case "Font":
                        ProcessFont(node, component);
                        break;

                    case "Align":
                        ProcessTextAlign(node, component);
                        break;

                    case "MarginLeft":
                        if (component.Margins.IsEmpty) component.Margins = new StiMargins();
                        component.Margins.Left = (int) ToHi(node.FirstChild.Value);
                        break;
                    case "MarginRight":
                        if (component.Margins.IsEmpty) component.Margins = new StiMargins();
                        component.Margins.Right = (int) ToHi(node.FirstChild.Value);
                        break;
                    case "MarginTop":
                        if (component.Margins.IsEmpty) component.Margins = new StiMargins();
                        component.Margins.Top = (int) ToHi(node.FirstChild.Value);
                        break;
                    case "MarginBottom":
                        if (component.Margins.IsEmpty) component.Margins = new StiMargins();
                        component.Margins.Bottom = (int) ToHi(node.FirstChild.Value);
                        break;

                    case "Format":
                        string format = node.FirstChild.Value;
                        if (format == "Currency")
                        {
                            component.TextFormat = new StiCurrencyFormatService();
                        }
                        else
                        {
                            component.TextFormat = new StiCustomFormatService(format);
                        }
                        break;

                    //case "RowIndex":
                    //case "ColumnIndex":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        if (!ProcessCommonComponentProperties(node, component)) ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            if (sectionIndex == -1)
            {
                page.Components.Add(component);
            }
            else
            {
                if (sectionIndex >= page.Components.Count) sectionIndex = page.Components.Count - 1;
                (page.Components[sectionIndex] as StiBand).Components.Add(component);
            }
        }
        #endregion

        #region Picture
        private void ProcessPicture(XmlNode baseNode, StiReport report, StiPage page)
        {
            StiImage component = new StiImage();
            component.CanGrow = false;
            component.CanShrink = false;

            int sectionIndex = -1;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    //case "CanGrow":
                    //    if (node.Value == "False") component.CanGrow = false;
                    //    break;

                    case "PictureScale":
                        if (node.FirstChild.Value == "0" || node.FirstChild.Value == "3")
                        {
                            //component.CanGrow = false;
                        }
                        if (node.FirstChild.Value == "1")
                        {
                            //component.CanGrow = false;
                            component.Stretch = true;
                        }
                        if (node.FirstChild.Value == "2" || node.FirstChild.Value == "4")
                        {
                            //component.CanGrow = false;
                            component.Stretch = true;
                            component.AspectRatio = true;
                        }
                        break;

                    case "Picture":
                        if (node.Attributes["encoding"] != null)
                        {
                            component.ImageBytes = Convert.FromBase64String(node.FirstChild.Value);
                        }
                        else
                        {
                            component.DataColumn = ConvertExpression(node.FirstChild.Value, report);
                        }
                        break;

                    //case "Url":
                    //    if (node.ChildNodes[0].Value.StartsWith("="))
                    //    {
                    //        component.DataColumn = ConvertExpression(node.ChildNodes[0].Value.Substring(1), report);
                    //    }
                    //    else
                    //    {
                    //        component.ImageURL = new StiImageURLExpression(node.ChildNodes[0].Value);
                    //    }
                    //    break;

                    //case "Style":
                    //    ProcessStyle(node, component);
                    //    break;

                    //case "CanShrink":
                    //    if (node.Value == "1") component.CanShrink = true;
                    //    break;

                    case "Section":
                        sectionIndex = int.Parse(node.FirstChild.Value);
                        break;

                    //case "Type":
                    //ignored or not implemented yet
                    //break;

                    default:
                        if (!ProcessCommonComponentProperties(node, component)) ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (sectionIndex == -1)
            {
                page.Components.Add(component);
            }
            else
            {
                if (sectionIndex >= page.Components.Count) sectionIndex = page.Components.Count - 1;
                (page.Components[sectionIndex] as StiBand).Components.Add(component);
            }
        }
        #endregion

        #region Barcode
        private void ProcessBarcode(XmlNode baseNode, StiReport report, StiPage page)
        {
            //    StiBarCode component = new StiBarCode();
            //    component.BarCodeType = new StiCode128AutoBarCodeType();

            //    foreach (XmlNode node in baseNode.Attributes)
            //    {
            //        switch (node.Name)
            //        {
            //            case "Value":
            //                if (node.Value.StartsWith("="))
            //                {
            //                    component.Code = new StiBarCodeExpression("{" + ConvertExpression(node.Value.Substring(1), report) + "}");
            //                }
            //                else
            //                {
            //                    component.Code = new StiBarCodeExpression(node.Value);
            //                }
            //                break;

            //            case "Symbology":
            //                if (node.Value == "Code93") component.BarCodeType = new StiCode93BarCodeType();
            //                else if (node.Value == "Code93Extended") component.BarCodeType = new StiCode93ExtBarCodeType();
            //                else if (node.Value == "Code128A") component.BarCodeType = new StiCode128aBarCodeType();
            //                else if (node.Value == "Code128B") component.BarCodeType = new StiCode128bBarCodeType();
            //                else if (node.Value == "Code128C") component.BarCodeType = new StiCode128cBarCodeType();
            //                else if (node.Value == "CodeMSI") component.BarCodeType = new StiMsiBarCodeType();
            //                else if (node.Value == "EAN8") component.BarCodeType = new StiEAN8BarCodeType();
            //                else if (node.Value == "EAN13") component.BarCodeType = new StiEAN13BarCodeType();
            //                else if (node.Value == "EAN128") component.BarCodeType = new StiEAN128AutoBarCodeType();
            //                else if (node.Value == "Postnet") component.BarCodeType = new StiPostnetBarCodeType();
            //                else if (node.Value == "UPCA") component.BarCodeType = new StiUpcABarCodeType();
            //                else if (node.Value == "UPCSupplement5") component.BarCodeType = new StiUpcSup5BarCodeType();
            //                else if (node.Value == "UPCSupplement2") component.BarCodeType = new StiUpcSup2BarCodeType();
            //                else if (node.Value == "UPCE") component.BarCodeType = new StiUpcEBarCodeType();
            //                else if (node.Value == "Code25Interleaved") component.BarCodeType = new StiInterleaved2of5BarCodeType();
            //                else if (node.Value == "Code39") component.BarCodeType = new StiCode39BarCodeType();
            //                else if (node.Value == "Code39Extended") component.BarCodeType = new StiCode39ExtBarCodeType();
            //                else if (node.Value == "Codabar") component.BarCodeType = new StiCodabarBarCodeType();
            //                else if (node.Value == "Code11") component.BarCodeType = new StiCode11BarCodeType();
            //                else if (node.Value == "Code25Standard") component.BarCodeType = new StiStandard2of5BarCodeType();
            //                break;

            //            case "Angle":
            //                double angle = Convert.ToDouble(node.Value);
            //                if (angle > 45 && angle < 135) component.Angle = StiAngle.Angle270;
            //                if (angle >= 135 && angle < 225) component.Angle = StiAngle.Angle180;
            //                if (angle >= 225 && angle < 315) component.Angle = StiAngle.Angle90;
            //                break;

            //            case "Stretch":
            //                if (node.Value == "True") component.AutoScale = true;
            //                break;

            //            case "RowIndex":
            //            case "ColumnIndex":
            //                //ignored or not implemented yet
            //                break;

            //            default:
            //                if (!ProcessCommonComponentProperties(node, component)) ThrowError(baseNode.Name, "#" + node.Name);
            //                break;
            //        }
            //    }

            //    foreach (XmlNode node in baseNode.ChildNodes)
            //    {
            //        switch (node.Name)
            //        {
            //            case "Style":
            //                ProcessStyle(node, component);
            //                break;

            //            default:
            //                ThrowError(baseNode.Name, node.Name);
            //                break;
            //        }
            //    }

            //    component.VertAlignment = StiVertAlignment.Center;
            //    component.HorAlignment = StiHorAlignment.Center;

            //    container.Components.Add(component);
            //    return component;
        }
        #endregion

        #region CheckBox
        private void ProcessCheckBox(XmlNode baseNode, StiReport report, StiPage page)
        {
        }
        #endregion

        #region Rtf
        private void ProcessRtf(XmlNode baseNode, StiReport report, StiPage page)
        {
        }
        #endregion

        #region Groups
        private void ProcessGroups(XmlNode baseNode, StiReport report, StiPage page)
        {
            int index = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                ProcessGroup(node, report, page, index);
                index++;
            }
        }

        private void ProcessGroup(XmlNode baseNode, StiReport report, StiPage page, int index)
        {
            int bandIndex = 5 + index * 2;
            if (bandIndex > page.Components.Count - 1) return;
            StiGroupHeaderBand group = page.Components[bandIndex] as StiGroupHeaderBand;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupBy":
                        group.Condition = new StiGroupConditionExpression("{" + ConvertExpression(node.FirstChild.Value, report) + "}");
                        break;

                    case "Sort":
                        if (node.FirstChild.Value == "0") group.SortDirection = StiGroupSortDirection.None;
                        if (node.FirstChild.Value == "1") group.SortDirection = StiGroupSortDirection.Ascending;
                        if (node.FirstChild.Value == "2") group.SortDirection = StiGroupSortDirection.Descending;
                        break;

                    case "KeepTogether":
                        if (node.FirstChild.Value == "0")
                        {
                            group.KeepGroupHeaderTogether = false;
                            group.KeepGroupTogether = false;
                        }
                        if (node.FirstChild.Value == "1")
                        {
                            group.KeepGroupHeaderTogether = true;
                            group.KeepGroupTogether = true;
                        }
                        if (node.FirstChild.Value == "2")
                        {
                            group.KeepGroupHeaderTogether = true;
                            group.KeepGroupTogether = false;
                        }
                        break;

                    case "Name":
                        //ignore
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region Style
        private void ProcessFont(XmlNode baseNode, StiComponent component)
        {
            IStiFont font = component as IStiFont;
            if (font == null) return;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Name":
                        font.Font = StiFontUtils.ChangeFontName(font.Font, node.FirstChild.Value);
                        break;

                    case "Bold":
                        if (node.FirstChild.Value == "-1") font.Font = StiFontUtils.ChangeFontStyleBold(font.Font, true);
                        break;

                    case "Italic":
                        if (node.FirstChild.Value == "-1") font.Font = StiFontUtils.ChangeFontStyleItalic(font.Font, true);
                        break;

                    case "Underline":
                        if (node.FirstChild.Value == "-1") font.Font = StiFontUtils.ChangeFontStyleUnderline(font.Font, true);
                        //font.Font = StiFontUtils.ChangeFontStyleStrikeout(font.Font, true);
                        break;

                    case "Size":
                        font.Font = StiFontUtils.ChangeFontSize(font.Font, (float) ParseDouble(node.FirstChild.Value));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessTextAlign(XmlNode baseNode, StiComponent component)
        {
            IStiTextHorAlignment horAlign = component as IStiTextHorAlignment;
            IStiVertAlignment vertAlign = component as IStiVertAlignment;

            switch (baseNode.FirstChild.Value)
            {
                case "0":
                    horAlign.HorAlignment = StiTextHorAlignment.Left;
                    vertAlign.VertAlignment = StiVertAlignment.Top;
                    break;
                case "1":
                    horAlign.HorAlignment = StiTextHorAlignment.Center;
                    vertAlign.VertAlignment = StiVertAlignment.Top;
                    break;
                case "2":
                    horAlign.HorAlignment = StiTextHorAlignment.Right;
                    vertAlign.VertAlignment = StiVertAlignment.Top;
                    break;

                case "6":
                    horAlign.HorAlignment = StiTextHorAlignment.Left;
                    vertAlign.VertAlignment = StiVertAlignment.Center;
                    break;
                case "7":
                    horAlign.HorAlignment = StiTextHorAlignment.Center;
                    vertAlign.VertAlignment = StiVertAlignment.Center;
                    break;
                case "8":
                    horAlign.HorAlignment = StiTextHorAlignment.Right;
                    vertAlign.VertAlignment = StiVertAlignment.Center;
                    break;

                case "3":
                    horAlign.HorAlignment = StiTextHorAlignment.Left;
                    vertAlign.VertAlignment = StiVertAlignment.Bottom;
                    break;
                case "4":
                    horAlign.HorAlignment = StiTextHorAlignment.Center;
                    vertAlign.VertAlignment = StiVertAlignment.Bottom;
                    break;
                case "5":
                    horAlign.HorAlignment = StiTextHorAlignment.Right;
                    vertAlign.VertAlignment = StiVertAlignment.Bottom;
                    break;

                case "9":
                    horAlign.HorAlignment = StiTextHorAlignment.Width;
                    vertAlign.VertAlignment = StiVertAlignment.Top;
                    break;
                case "11":
                    horAlign.HorAlignment = StiTextHorAlignment.Width;
                    vertAlign.VertAlignment = StiVertAlignment.Center;
                    break;
                case "10":
                    horAlign.HorAlignment = StiTextHorAlignment.Width;
                    vertAlign.VertAlignment = StiVertAlignment.Bottom;
                    break;

                default:
                    ThrowError(baseNode.Name, "=" + baseNode.FirstChild.Value);
                    break;
            }
        }

        private void ProcessBorderStyle(XmlNode baseNode, StiComponent component)
        {
            IStiBorder border = component as IStiBorder;
            if (border == null) return;

            switch (baseNode.FirstChild.Value)
            {
                case "0":
                    border.Border.Style = StiPenStyle.None;
                    break;

                case "1":
                    border.Border.Style = StiPenStyle.Solid;
                    border.Border.Side = StiBorderSides.All;
                    break;

                case "2":
                    border.Border.Style = StiPenStyle.Dash;
                    border.Border.Side = StiBorderSides.All;
                    border.Border.Size = 2;
                    break;

                case "3":
                    border.Border.Style = StiPenStyle.Dot;
                    border.Border.Side = StiBorderSides.All;
                    border.Border.Size = 2;
                    break;

                case "4":
                    border.Border.Style = StiPenStyle.DashDot;
                    border.Border.Side = StiBorderSides.All;
                    border.Border.Size = 2;
                    break;

                case "5":
                    border.Border.Style = StiPenStyle.DashDotDot;
                    border.Border.Side = StiBorderSides.All;
                    border.Border.Size = 2;
                    break;

                default:
                    ThrowError(baseNode.Name, "=" + baseNode.Name);
                    break;
            }
        }

        //private static Hashtable HtmlNameToColor = null;

        //private static Color ParseStyleColor(string colorAttribute)
        //{
        //    Color color = Color.Transparent;
        //    if (colorAttribute != null && colorAttribute.Length > 1)
        //    {
        //        if (colorAttribute[0] == '#')
        //        {
        //            #region Parse RGB value in hexadecimal notation
        //            string colorSt = colorAttribute.Substring(1).ToLowerInvariant();
        //            StringBuilder sbc = new StringBuilder();
        //            foreach (char ch in colorSt)
        //            {
        //                if (ch == '0' || ch == '1' || ch == '2' || ch == '3' || ch == '4' || ch == '5' || ch == '6' || ch == '7' ||
        //                    ch == '8' || ch == '9' || ch == 'a' || ch == 'b' || ch == 'c' || ch == 'd' || ch == 'e' || ch == 'f') sbc.Append(ch);
        //            }
        //            if (sbc.Length == 3)
        //            {
        //                colorSt = string.Format("{0}{0}{1}{1}{2}{2}", sbc[0], sbc[1], sbc[2]);
        //            }
        //            else
        //            {
        //                colorSt = sbc.ToString();
        //            }
        //            if (colorSt.Length == 6)
        //            {
        //                int colorInt = Convert.ToInt32(colorSt, 16);
        //                color = Color.FromArgb(0xFF, (colorInt >> 16) & 0xFF, (colorInt >> 8) & 0xFF, colorInt & 0xFF);
        //            }
        //            #endregion
        //        }
        //        else if (colorAttribute.StartsWith("rgb"))
        //        {
        //            #region Parse RGB function
        //            string[] colors = colorAttribute.Substring(4, colorAttribute.Length - 5).Split(new char[] { ',' });
        //            color = Color.FromArgb(0xFF, int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]));
        //            #endregion
        //        }
        //        else if (colorAttribute.IndexOf(',') != -1)
        //        {
        //            #region Parse RGB values
        //            string[] colors = colorAttribute.Split(new char[] { ',' });
        //            if (colors.Length == 3)
        //            {
        //                color = Color.FromArgb(0xFF, int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]));
        //            }
        //            else if (colors.Length == 4)
        //            {
        //                color = Color.FromArgb(int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]), int.Parse(colors[3]));
        //            }
        //            #endregion
        //        }
        //        else
        //        {
        //            #region Parse color keywords
        //            if (HtmlNameToColor == null)
        //            {
        //                #region Init hashtable
        //                string[,] initData = {
        //                {"AliceBlue",	    "#F0F8FF"},
        //                {"AntiqueWhite",	"#FAEBD7"},
        //                {"Aqua",	    "#00FFFF"},
        //                {"Aquamarine",	"#7FFFD4"},
        //                {"Azure",	    "#F0FFFF"},
        //                {"Beige",	    "#F5F5DC"},
        //                {"Bisque",	    "#FFE4C4"},
        //                {"Black",	    "#000000"},
        //                {"BlanchedAlmond",	"#FFEBCD"},
        //                {"Blue",	    "#0000FF"},
        //                {"BlueViolet",	"#8A2BE2"},
        //                {"Brown",	    "#A52A2A"},
        //                {"BurlyWood",	"#DEB887"},
        //                {"CadetBlue",	"#5F9EA0"},
        //                {"Chartreuse",	"#7FFF00"},
        //                {"Chocolate",	"#D2691E"},
        //                {"Coral",	    "#FF7F50"},
        //                {"CornflowerBlue",	"#6495ED"},
        //                {"Cornsilk",	"#FFF8DC"},
        //                {"Crimson",	    "#DC143C"},
        //                {"Cyan",	    "#00FFFF"},
        //                {"DarkBlue",	"#00008B"},
        //                {"DarkCyan",	"#008B8B"},
        //                {"DarkGoldenRod",	"#B8860B"},
        //                {"DarkGray",	"#A9A9A9"},
        //                {"DarkGrey",	"#A9A9A9"},
        //                {"DarkGreen",	"#006400"},
        //                {"DarkKhaki",	"#BDB76B"},
        //                {"DarkMagenta",	"#8B008B"},
        //                {"DarkOliveGreen",	"#556B2F"},
        //                {"Darkorange",	"#FF8C00"},
        //                {"DarkOrchid",	"#9932CC"},
        //                {"DarkRed",	    "#8B0000"},
        //                {"DarkSalmon",	"#E9967A"},
        //                {"DarkSeaGreen",	"#8FBC8F"},
        //                {"DarkSlateBlue",	"#483D8B"},
        //                {"DarkSlateGray",	"#2F4F4F"},
        //                {"DarkSlateGrey",	"#2F4F4F"},
        //                {"DarkTurquoise",	"#00CED1"},
        //                {"DarkViolet",	"#9400D3"},
        //                {"DeepPink",	"#FF1493"},
        //                {"DeepSkyBlue",	"#00BFFF"},
        //                {"DimGray",	    "#696969"},
        //                {"DimGrey",	    "#696969"},
        //                {"DodgerBlue",	"#1E90FF"},
        //                {"FireBrick",	"#B22222"},
        //                {"FloralWhite",	"#FFFAF0"},
        //                {"ForestGreen",	"#228B22"},
        //                {"Fuchsia",	    "#FF00FF"},
        //                {"Gainsboro",	"#DCDCDC"},
        //                {"GhostWhite",	"#F8F8FF"},
        //                {"Gold",	    "#FFD700"},
        //                {"GoldenRod",	"#DAA520"},
        //                {"Gray",	    "#808080"},
        //                {"Grey",	    "#808080"},
        //                {"Green",	    "#008000"},
        //                {"GreenYellow",	"#ADFF2F"},
        //                {"HoneyDew",	"#F0FFF0"},
        //                {"HotPink",	    "#FF69B4"},
        //                {"IndianRed",	"#CD5C5C"},
        //                {"Indigo",	    "#4B0082"},
        //                {"Ivory",	    "#FFFFF0"},
        //                {"Khaki",	    "#F0E68C"},
        //                {"Lavender",	"#E6E6FA"},
        //                {"LavenderBlush",	"#FFF0F5"},
        //                {"LawnGreen",	"#7CFC00"},
        //                {"LemonChiffon",	"#FFFACD"},
        //                {"LightBlue",	"#ADD8E6"},
        //                {"LightCoral",	"#F08080"},
        //                {"LightCyan",	"#E0FFFF"},
        //                {"LightGoldenRodYellow",	"#FAFAD2"},
        //                {"LightGray",	"#D3D3D3"},
        //                {"LightGrey",	"#D3D3D3"},
        //                {"LightGreen",	"#90EE90"},
        //                {"LightPink",	"#FFB6C1"},
        //                {"LightSalmon",	"#FFA07A"},
        //                {"LightSeaGreen",	"#20B2AA"},
        //                {"LightSkyBlue",	"#87CEFA"},
        //                {"LightSlateGray",	"#778899"},
        //                {"LightSlateGrey",	"#778899"},
        //                {"LightSteelBlue",	"#B0C4DE"},
        //                {"LightYellow",	"#FFFFE0"},
        //                {"Lime",	    "#00FF00"},
        //                {"LimeGreen",	"#32CD32"},
        //                {"Linen",	    "#FAF0E6"},
        //                {"Magenta",	    "#FF00FF"},
        //                {"Maroon",	    "#800000"},
        //                {"MediumAquaMarine",	"#66CDAA"},
        //                {"MediumBlue",	"#0000CD"},
        //                {"MediumOrchid",	"#BA55D3"},
        //                {"MediumPurple",	"#9370D8"},
        //                {"MediumSeaGreen",	"#3CB371"},
        //                {"MediumSlateBlue",	"#7B68EE"},
        //                {"MediumSpringGreen",	"#00FA9A"},
        //                {"MediumTurquoise",	"#48D1CC"},
        //                {"MediumVioletRed",	"#C71585"},
        //                {"MidnightBlue",	"#191970"},
        //                {"MintCream",	"#F5FFFA"},
        //                {"MistyRose",	"#FFE4E1"},
        //                {"Moccasin",	"#FFE4B5"},
        //                {"NavajoWhite",	"#FFDEAD"},
        //                {"Navy",	    "#000080"},
        //                {"OldLace",	    "#FDF5E6"},
        //                {"Olive",	    "#808000"},
        //                {"OliveDrab",	"#6B8E23"},
        //                {"Orange",	    "#FFA500"},
        //                {"OrangeRed",	"#FF4500"},
        //                {"Orchid",	    "#DA70D6"},
        //                {"PaleGoldenRod",	"#EEE8AA"},
        //                {"PaleGreen",	"#98FB98"},
        //                {"PaleTurquoise",	"#AFEEEE"},
        //                {"PaleVioletRed",	"#D87093"},
        //                {"PapayaWhip",	"#FFEFD5"},
        //                {"PeachPuff",	"#FFDAB9"},
        //                {"Peru",	    "#CD853F"},
        //                {"Pink",	    "#FFC0CB"},
        //                {"Plum",	    "#DDA0DD"},
        //                {"PowderBlue",	"#B0E0E6"},
        //                {"Purple",	    "#800080"},
        //                {"Red",	        "#FF0000"},
        //                {"RosyBrown",	"#BC8F8F"},
        //                {"RoyalBlue",	"#4169E1"},
        //                {"SaddleBrown",	"#8B4513"},
        //                {"Salmon",	    "#FA8072"},
        //                {"SandyBrown",	"#F4A460"},
        //                {"SeaGreen",	"#2E8B57"},
        //                {"SeaShell",	"#FFF5EE"},
        //                {"Sienna",	    "#A0522D"},
        //                {"Silver",	    "#C0C0C0"},
        //                {"SkyBlue",	    "#87CEEB"},
        //                {"SlateBlue",	"#6A5ACD"},
        //                {"SlateGray",	"#708090"},
        //                {"SlateGrey",	"#708090"},
        //                {"Snow",	    "#FFFAFA"},
        //                {"SpringGreen",	"#00FF7F"},
        //                {"SteelBlue",	"#4682B4"},
        //                {"Tan",	        "#D2B48C"},
        //                {"Teal",	    "#008080"},
        //                {"Thistle",	    "#D8BFD8"},
        //                {"Tomato",	    "#FF6347"},
        //                {"Turquoise",	"#40E0D0"},
        //                {"Violet",	    "#EE82EE"},
        //                {"Wheat",	    "#F5DEB3"},
        //                {"White",	    "#FFFFFF"},
        //                {"WhiteSmoke",	"#F5F5F5"},
        //                {"Yellow",	    "#FFFF00"},
        //                {"YellowGreen",	"#9ACD32"}};

        //                HtmlNameToColor = new Hashtable();
        //                for (int index = 0; index < initData.GetLength(0); index++)
        //                {
        //                    string key = initData[index, 0].ToLowerInvariant();
        //                    int colorInt = Convert.ToInt32(initData[index, 1].Substring(1), 16);
        //                    Color value = Color.FromArgb(0xFF, (colorInt >> 16) & 0xFF, (colorInt >> 8) & 0xFF, colorInt & 0xFF);
        //                    HtmlNameToColor[key] = value;
        //                }
        //                #endregion
        //            }
        //            string colorSt = colorAttribute.ToLowerInvariant();
        //            if (HtmlNameToColor.ContainsKey(colorSt))
        //            {
        //                color = (Color)HtmlNameToColor[colorSt];
        //            }
        //            #endregion
        //        }
        //    }
        //    return color;
        //}

        //private static StiPenStyle ParseBorderStyle(string styleName)
        //{
        //    switch (styleName)
        //    {
        //        case "None":
        //            return StiPenStyle.None;
        //        case "Dashed":
        //            return StiPenStyle.Dash;
        //        case "Dotted":
        //            return StiPenStyle.Dot;
        //        case "Double":
        //            return StiPenStyle.Double;
        //    }
        //    return StiPenStyle.Solid;
        //}
        #endregion

        #region Utils
        private static string ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

        private double ToHi(string strValue)
        {
            string stSize = strValue.Trim();
            double factor = 1 / 14.4;
            //double factor = 1;
            if (strValue.EndsWith("cm"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 100 / 2.54;
            }
            if (strValue.EndsWith("mm"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 10 / 2.54;
            }
            if (strValue.EndsWith("in"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 100;
            }
            if (strValue.EndsWith("pt"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                //factor = 100 / 72f;
                factor = 1;
            }
            if (strValue.EndsWith("px"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 100 / 96f;
            }
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
            uint colorValue = (uint) int.Parse(node.Value);
            return Color.FromArgb(0xFF, (int) (colorValue & 0xFF), (int) ((colorValue >> 8) & 0xFF), (int) ((colorValue >> 16) & 0xFF));
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
        #endregion

        #region Methods.Helper
        public string ConvertExpression(string baseExpression, StiReport report)
        {
            string st = baseExpression.Trim();
            bool flag = true;
            foreach (char ch in st)
            {
                if (!(char.IsLetterOrDigit(ch) || ch == '_')) flag = false;
            }
            if (flag) //only one field
            {
                string fieldName = currentDataSourceName + "." + st;
                if (!fieldsNames.ContainsKey(fieldName))
                {
                    fields.Add(fieldName);
                    fieldsNames[fieldName] = null;
                }
                return fieldName;
            }

            //complex expression
            bool inString = false;
            StringBuilder expr = new StringBuilder();
            for (int index = 0; index < baseExpression.Length; index++)
            {
                char ch = baseExpression[index];
                if (ch == '"') inString = !inString;
                if (!inString && (ch == '&'))
                {
                    ch = '+';
                }
                if (!inString && char.IsLetter(ch))
                {
                    int index2 = index;
                    while (index2 < baseExpression.Length && char.IsLetterOrDigit(baseExpression[index2]))
                    {
                        index2++;
                    }
                    string word = baseExpression.Substring(index, index2 - index);
                    bool isField = true;
                    //if (index > 0 && !char.IsLetterOrDigit(baseExpression[index - 1])) isField = false;
                    if (index > 0 && baseExpression[index - 1] == '[') isField = false;
                    if (index2 < baseExpression.Length && baseExpression[index2] == ']') isField = false;
                    if (index2 < baseExpression.Length && baseExpression[index2] == '[') isField = false;
                    if (index2 < baseExpression.Length && baseExpression[index2] == '(') isField = false;
                    if (reservedWords.Contains(word)) isField = false;
                    if (isField)
                    {
                        string fieldName = currentDataSourceName + "." + word;
                        if (!fieldsNames.ContainsKey(fieldName))
                        {
                            fields.Add(fieldName);
                            fieldsNames[fieldName] = null;
                        }
                        expr.Append(currentDataSourceName + "." + word);
                    }
                    else
                    {
                        expr.Append(word);
                    }
                    index += word.Length - 1;
                }
                else
                {
                    expr.Append(ch);
                }
            }

            expr.Replace("Now()", "Today");
            expr.Replace("[Page]", "PageNumber");
            expr.Replace("[Pages]", "TotalPageCount");
            expr.Replace("vbcrlf", "\"\\r\\n\"");

            return expr.ToString();
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
                var helper = new StiComponentOneReportsHelper();
                var errors = new List<string>();
                var doc = new XmlDocument();

                using (var stream = new MemoryStream(bytes))
                {
                    doc.Load(stream);
                }

                report.Pages.Clear();
                int pageIndex = 0;
                foreach (XmlNode node in doc.DocumentElement.ChildNodes)
                {
                    if (node.NodeType != XmlNodeType.Comment)
                    {
                        report.Pages.Add(new StiPage(report));
                        helper.ProcessRootNode(node, report, errors, pageIndex++);
                    }
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