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
using Stimulsoft.Report.CrossTab;

namespace Stimulsoft.Report.Import
{
    public class StiTelerikReportsHelper
    {
        #region Fields
        ArrayList errorList = null;

        ArrayList fields = new ArrayList();
        Hashtable fieldsNames = new Hashtable();

        string currentDataSourceName = "ds";
        string mainDataSourceName = "ds";

        StiPanel groupHeadersList = new StiPanel();
        StiPanel groupFootersList = new StiPanel();
        #endregion

        #region Root node
        public void ProcessRootNode(XmlNode rootNode, StiReport report, ArrayList errorList)
        {
            this.errorList = errorList;
            this.fields.Clear();
            this.fieldsNames = new Hashtable();
            report.ReportUnit = StiReportUnitType.HundredthsOfInch;
            StiPage page = report.Pages[0];

            if (rootNode.Attributes["Name"] != null)
            {
                report.ReportName = rootNode.Attributes["Name"].Value;
            }

            bool pageSettingsUsed = false;
            bool dataSourceUsed = false;

            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {

                    case "PageSettings":
                        pageSettingsUsed = ProcessPageSettings(node, report, page);
                        break;

                    case "DataSource":
                    case "DataSources":
                        dataSourceUsed = ProcessDataSources(node, report);
                        mainDataSourceName = currentDataSourceName;
                        break;

                    case "ReportParameters":
                        ProcessReportParameters(node, report);
                        break;

                    case "Items":
                        ProcessItems(node, report, page);
                        break;

                    case "Groups":
                        ProcessGroups(node, report, page);
                        break;

                    //case "wwwwWidth":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(rootNode.Name, node.Name);
                        break;
                }
            }

            #region Check PageSetting
            double printWidth = 0;
            if (!pageSettingsUsed && rootNode.Attributes["Width"] != null)
            {
                printWidth = ToHi(rootNode.Attributes["Width"].Value);
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
                report.Dictionary.DataSources.Add(dataSource);
            }

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

            foreach (StiComponent comp in report.GetComponents())
            {
                comp.Page = page;
            }
        }
        #endregion

        #region PageSettings
        private bool ProcessPageSettings(XmlNode baseNode, StiReport report, StiPage page)
        {
            bool pageSettingsUsed = false;
            if (baseNode.Attributes["Landscape"] != null)
            {
                if (baseNode.Attributes["Landscape"].Value == "True") page.Orientation = StiPageOrientation.Landscape;
            }
            if (baseNode.Attributes["PaperKind"] != null)
            {
                page.PaperSize = (PaperKind)Enum.Parse(typeof(PaperKind), baseNode.Attributes["PaperKind"].Value);
                pageSettingsUsed = true;
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Margins":
                        if (node.Attributes["Left"] != null)
                        {
                            page.Margins.Left = ToHi(node.Attributes["Left"].Value);
                        }
                        if (node.Attributes["Right"] != null)
                        {
                            page.Margins.Right = ToHi(node.Attributes["Right"].Value);
                        }
                        if (node.Attributes["Top"] != null)
                        {
                            page.Margins.Top = ToHi(node.Attributes["Top"].Value);
                        }
                        if (node.Attributes["Bottom"] != null)
                        {
                            page.Margins.Bottom = ToHi(node.Attributes["Bottom"].Value);
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return pageSettingsUsed;
        }
        #endregion

        #region DataSource
        private bool ProcessDataSources(XmlNode baseNode, StiReport report)
        {
            bool dataSourceUsed = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "SqlDataSource":
                        ProcessSqlDataSource(node, report);
                        dataSourceUsed = true;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return dataSourceUsed;
        }

        private void ProcessSqlDataSource(XmlNode baseNode, StiReport report)
        {
            StiSqlDatabase dataBase = new StiSqlDatabase();
            dataBase.Name = "db";
            dataBase.Alias = "db";
            StiSqlSource dataSource = new StiSqlSource();
            dataSource.Name = "ds";
            dataSource.Alias = "ds";
            dataSource.NameInSource = "db";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "ConnectionString":
                        dataBase.ConnectionString = node.Value;
                        break;
                    case "SelectCommand":
                        dataSource.SqlCommand = node.Value;
                        break;
                    case "Name":
                        dataSource.Name = node.Value;
                        dataSource.Alias = node.Value;
                        dataSource.NameInSource = "db" + node.Value;
                        dataBase.Name = "db" + node.Value;
                        dataBase.Alias = "db" + node.Value;
                        currentDataSourceName = node.Value;
                        break;

                    //case "Type":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            report.Dictionary.Databases.Add(dataBase);
            report.Dictionary.DataSources.Add(dataSource);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Parameters":
                        ProcessParameters(node, report, dataSource);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessParameters(XmlNode baseNode, StiReport report, StiSqlSource dataSource)
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
                            if (value.StartsWith("="))
                            {
                                dp.Value = ConvertExpression(value.Substring(1), report);
                            }
                            else
                            {
                                dp.Value = value;
                            }
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            dataSource.Parameters.Add(dp);
        }
        #endregion

        #region ReportParameters
        private void ProcessReportParameters(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ReportParameter":
                        ProcessReportParameter(node, report);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessReportParameter(XmlNode baseNode, StiReport report)
        {
            StiVariable var = new StiVariable("Parameters", "Name", typeof(string));
            var.RequestFromUser = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        var.Name = node.Value;
                        var.Alias = node.Value;
                        break;
                    case "Text":
                        var.Alias = node.Value;
                        break;
                    case "Type":
                        if (node.Value == "Integer") var.Type = typeof(Int32);
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            report.Dictionary.Variables[var.Name] = var;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        var.Value = node.ChildNodes[0].Value;
                        break;

                    case "AvailableValues":
                        ProcessAvailableValues(node, report, var);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessAvailableValues(XmlNode baseNode, StiReport report, StiVariable var)
        {
            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    //case "Name":
                    //    var.Name = node.Value;
                    //    break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSource":
                    case "DataSources":
                        string store = currentDataSourceName;
                        ProcessDataSources(node, report);
                        currentDataSourceName = store;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region Items
        private void ProcessItems(XmlNode baseNode, StiReport report, StiContainer container)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DetailSection":
                    case "GroupHeaderSection":
                    case "GroupFooterSection":
                    case "ReportHeaderSection":
                    case "ReportFooterSection":
                    case "PageHeaderSection":
                    case "PageFooterSection":
                        ProcessSection(node, report, container);
                        break;

                    case "Panel":
                        ProcessPanel(node, report, container);
                        break;

                    case "TextBox":
                        ProcessTextBox(node, report, container);
                        break;

                    case "PictureBox":
                        ProcessPictureBox(node, report, container);
                        break;

                    case "Barcode":
                        ProcessBarcode(node, report, container);
                        break;

                    case "Table":
                        ProcessTable(node, report, container);
                        break;

                    case "Crosstab":
                        ProcessCrosstab(node, report, container);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        #region ProcessCommonComponentProperties
        private bool ProcessCommonComponentProperties(XmlNode node, StiComponent component)
        {
            //IStiBorder border = component as IStiBorder;

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


                //case "ClassName":
                //    component.ComponentStyle = node.Value;
                //    return true;

                //case "Type":
                //    //ignored
                //    return true;

            }
            return false;
        }
        #endregion

        #region Section
        private void ProcessSection(XmlNode baseNode, StiReport report, StiContainer container)
        {
            string sectionType = baseNode.Name;
            StiBand band = new StiDataBand();
            if (sectionType == "ReportHeaderSection") band = new StiReportTitleBand();
            if (sectionType == "PageHeaderSection") band = new StiPageHeaderBand();
            if (sectionType == "GroupHeaderSection") band = new StiGroupHeaderBand();
            if (sectionType == "DetailSection") band = new StiDataBand();
            if (sectionType == "GroupFooterSection") band = new StiGroupFooterBand();
            if (sectionType == "PageFooterSection") band = new StiPageFooterBand();
            if (sectionType == "ReportFooterSection") band = new StiReportSummaryBand();

            if (band is StiDataBand) (band as StiDataBand).DataSourceName = currentDataSourceName;

            band.CanShrink = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    ////case "Name":
                    ////    band.Name = node.Value;
                    ////    break;
                    ////case "Height":
                    ////    band.Height = ToHi(node.Value);
                    ////    break;

                    case "ColumnCount":
                        if (band is StiDataBand) (band as StiDataBand).Columns = int.Parse(node.Value);
                        break;

                    //case "CanGrow":
                    //    if (node.Value == "0") band.CanGrow = false;
                    //    break;
                    //case "CanShrink":
                    //    if (node.Value == "1") band.CanShrink = true;
                    //    break;

                    //case "Type":
                    //    //ignored or not implemented yet
                    //    break;

                    //case "DataField":
                    //    if (band is StiGroupHeaderBand) (band as StiGroupHeaderBand).Condition = new StiGroupConditionExpression("{" + ConvertExpression(node.Value, report) + "}");
                    //    break;

                    default:
                        //ThrowError(baseNode.Name, "#" + node.Name);
                        if (!ProcessCommonComponentProperties(node, band)) ThrowError(node.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Items":
                        ProcessItems(node, report, band);
                        break;

                    case "Style":
                        ProcessStyle(node, band);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (sectionType == "DetailSection")
            {
                for (int index = 0; index < groupHeadersList.Components.Count; index++)
                {
                    container.Components.Add(groupHeadersList.Components[index]);
                }
            }
            container.Components.Add(band);
            if (sectionType == "DetailSection")
            {
                for (int index = groupFootersList.Components.Count - 1; index >= 0; index--)
                {
                    container.Components.Add(groupFootersList.Components[index]);
                }
            }
        }
        #endregion

        #region Panel
        private StiComponent ProcessPanel(XmlNode baseNode, StiReport report, StiContainer container)
        {
            StiPanel panel = new StiPanel();
            panel.CanGrow = true;
            panel.CanShrink = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {

                    default:
                        //ThrowError(baseNode.Name, "#" + node.Name);
                        if (!ProcessCommonComponentProperties(node, panel)) ThrowError(node.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Items":
                        ProcessItems(node, report, panel);
                        break;

                    case "Style":
                        ProcessStyle(node, panel);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            container.Components.Add(panel);
            return panel;
        }
        #endregion

        #region TextBox
        private StiComponent ProcessTextBox(XmlNode baseNode, StiReport report, StiContainer container)
        {
            StiText component = new StiText();
            component.CanGrow = true;
            component.CanShrink = false;
            component.WordWrap = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Value":
                        if (node.Value.StartsWith("="))
                        {
                            component.Text = "{" + ConvertExpression(node.Value.Substring(1), report) + "}";
                        }
                        else
                        {
                            component.Text = node.Value;
                        }
                        break;

                    case "CanGrow":
                        if (node.Value == "False") component.CanGrow = false;
                        break;

                    //case "CanShrink":
                    //    if (node.Value == "1") component.CanShrink = true;
                    //    break;

                    case "TextWrap":
                        if (node.Value == "False") component.WordWrap = false;
                        break;

                    case "Format":
                        string format = node.Value;
                        format = format.Substring(1, format.Length - 2);
                        int pos = format.IndexOf(':');
                        if (pos != -1)
                        {
                            format = format.Substring(pos + 1);
                            component.TextFormat = new StiCustomFormatService(format);
                        }
                        break;

                    case "RowIndex":
                    case "ColumnIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        if (!ProcessCommonComponentProperties(node, component)) ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Style":
                        ProcessStyle(node, component);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            container.Components.Add(component);
            return component;
        }
        #endregion

        #region PictureBox
        private StiComponent ProcessPictureBox(XmlNode baseNode, StiReport report, StiContainer container)
        {
            StiImage component = new StiImage();
            component.CanGrow = true;
            component.CanShrink = false;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "CanGrow":
                        if (node.Value == "False") component.CanGrow = false;
                        break;

                    case "Sizing":
                        if (node.Value == "Center")
                        {
                            component.CanGrow = false;
                            component.VertAlignment = StiVertAlignment.Center;
                            component.HorAlignment = StiHorAlignment.Center;
                        }
                        if (node.Value == "Normal")
                        {
                            component.CanGrow = false;
                        }
                        if (node.Value == "Stretch")
                        {
                            component.CanGrow = false;
                            component.Stretch = true;
                        }
                        if (node.Value == "ScaleProportional")
                        {
                            component.CanGrow = false;
                            component.Stretch = true;
                            component.AspectRatio = true;
                            component.VertAlignment = StiVertAlignment.Center;
                            component.HorAlignment = StiHorAlignment.Center;
                        }
                        break;

                    //case "CanShrink":
                    //    if (node.Value == "1") component.CanShrink = true;
                    //    break;

                    //case "Type":
                    //ignored or not implemented yet
                    //break;

                    default:
                        if (!ProcessCommonComponentProperties(node, component)) ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        component.Image = Image.FromStream(new MemoryStream(Convert.FromBase64String(node.ChildNodes[0].Value)));
                        break;

                    case "Url":
                        if (node.ChildNodes[0].Value.StartsWith("="))
                        {
                            component.DataColumn = ConvertExpression(node.ChildNodes[0].Value.Substring(1), report);
                        }
                        else
                        {
                            component.ImageURL = new StiImageURLExpression(node.ChildNodes[0].Value);
                        }
                        break;

                    case "Style":
                        ProcessStyle(node, component);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            container.Components.Add(component);
            return component;
        }
        #endregion

        #region Barcode
        private StiComponent ProcessBarcode(XmlNode baseNode, StiReport report, StiContainer container)
        {
            StiBarCode component = new StiBarCode();
            component.BarCodeType = new StiCode128AutoBarCodeType();

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Value":
                        if (node.Value.StartsWith("="))
                        {
                            component.Code = new StiBarCodeExpression("{" + ConvertExpression(node.Value.Substring(1), report) + "}");
                        }
                        else
                        {
                            component.Code = new StiBarCodeExpression(node.Value);
                        }
                        break;

                    case "Symbology":
                        if (node.Value == "Code93") component.BarCodeType = new StiCode93BarCodeType();
                        else if (node.Value == "Code93Extended") component.BarCodeType = new StiCode93ExtBarCodeType();
                        else if (node.Value == "Code128A") component.BarCodeType = new StiCode128aBarCodeType();
                        else if (node.Value == "Code128B") component.BarCodeType = new StiCode128bBarCodeType();
                        else if (node.Value == "Code128C") component.BarCodeType = new StiCode128cBarCodeType();
                        else if (node.Value == "CodeMSI") component.BarCodeType = new StiMsiBarCodeType();
                        else if (node.Value == "EAN8") component.BarCodeType = new StiEAN8BarCodeType();
                        else if (node.Value == "EAN13") component.BarCodeType = new StiEAN13BarCodeType();
                        else if (node.Value == "EAN128") component.BarCodeType = new StiEAN128AutoBarCodeType();
                        else if (node.Value == "Postnet") component.BarCodeType = new StiPostnetBarCodeType();
                        else if (node.Value == "UPCA") component.BarCodeType = new StiUpcABarCodeType();
                        else if (node.Value == "UPCSupplement5") component.BarCodeType = new StiUpcSup5BarCodeType();
                        else if (node.Value == "UPCSupplement2") component.BarCodeType = new StiUpcSup2BarCodeType();
                        else if (node.Value == "UPCE") component.BarCodeType = new StiUpcEBarCodeType();
                        else if (node.Value == "Code25Interleaved") component.BarCodeType = new StiInterleaved2of5BarCodeType();
                        else if (node.Value == "Code39") component.BarCodeType = new StiCode39BarCodeType();
                        else if (node.Value == "Code39Extended") component.BarCodeType = new StiCode39ExtBarCodeType();
                        else if (node.Value == "Codabar") component.BarCodeType = new StiCodabarBarCodeType();
                        else if (node.Value == "Code11") component.BarCodeType = new StiCode11BarCodeType();
                        else if (node.Value == "Code25Standard") component.BarCodeType = new StiStandard2of5BarCodeType();
                        break;

                    case "Angle":
                        double angle = Convert.ToDouble(node.Value);
                        if (angle > 45 && angle < 135) component.Angle = StiAngle.Angle270;
                        if (angle >= 135 && angle < 225) component.Angle = StiAngle.Angle180;
                        if (angle >= 225 && angle < 315) component.Angle = StiAngle.Angle90;
                        break;

                    case "Stretch":
                        if (node.Value == "True") component.AutoScale = true;
                        break;

                    case "RowIndex":
                    case "ColumnIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        if (!ProcessCommonComponentProperties(node, component)) ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Style":
                        ProcessStyle(node, component);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            component.VertAlignment = StiVertAlignment.Center;
            component.HorAlignment = StiHorAlignment.Center;

            container.Components.Add(component);
            return component;
        }
        #endregion

        #region Table
        private void ProcessTable(XmlNode baseNode, StiReport report, StiContainer container)
        {
            string storeDataSource = currentDataSourceName;

            StiPanel table = new StiPanel();
            table.CanGrow = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    ////case "Name":
                    ////    band.Name = node.Value;
                    ////    break;

                    default:
                        //ThrowError(baseNode.Name, "#" + node.Name);
                        if (!ProcessCommonComponentProperties(node, table)) ThrowError(node.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSource":
                    case "DataSources":
                        ProcessDataSources(node, report);
                        break;

                    case "Body":
                        ProcessTableBody(node, report, table);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            container.Components.Add(table);

            currentDataSourceName = storeDataSource;
        }

        private void ProcessTableBody(XmlNode baseNode, StiReport report, StiPanel container)
        {
            List<double> columnWidths = new List<double>();
            List<double> rowHeights = new List<double>();
            StiComponent[,] cells = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Columns":
                        columnWidths = ProcessTableBodyColumns(node, report);
                        break;

                    case "Rows":
                        rowHeights = ProcessTableBodyRows(node, report);
                        break;

                    case "Cells":
                        cells = new StiComponent[rowHeights.Count, columnWidths.Count];
                        ProcessTableBodyCells(node, report, cells);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            StiDataBand band = new StiDataBand();
            band.Name = container.Name + "DataBand";
            band.DataSourceName = currentDataSourceName;

            double posY = 0;
            for (int indexRow = 0; indexRow < rowHeights.Count; indexRow++)
            {
                double posX = 0;
                for (int indexColumn = 0; indexColumn < columnWidths.Count; indexColumn++)
                {
                    StiComponent component = cells[indexRow, indexColumn];
                    if (component != null)
                    {
                        component.Left = posX;
                        component.Top = posY;
                        band.Components.Add(component);
                    }
                    posX += columnWidths[indexColumn];
                }
                posY += rowHeights[indexRow];
            }

            container.Components.Add(band);
        }

        private List<double> ProcessTableBodyColumns(XmlNode baseNode, StiReport report)
        {
            List<double> columnWidths = new List<double>();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Column":
                        columnWidths.Add(ToHi(node.Attributes["Width"].Value));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return columnWidths;
        }

        private List<double> ProcessTableBodyRows(XmlNode baseNode, StiReport report)
        {
            List<double> rowHeights = new List<double>();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Row":
                        rowHeights.Add(ToHi(node.Attributes["Height"].Value));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return rowHeights;
        }

        private void ProcessTableBodyCells(XmlNode baseNode, StiReport report, StiComponent[,] cells)
        {
            StiPanel tempPanel = new StiPanel();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                StiComponent component = null;
                switch (node.Name)
                {
                    case "TextBox":
                        component = ProcessTextBox(node, report, tempPanel);
                        break;

                    case "PictureBox":
                        component = ProcessPictureBox(node, report, tempPanel);
                        break;

                    case "Barcode":
                        component = ProcessBarcode(node, report, tempPanel);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
                if (component != null)
                {
                    int row = int.Parse(node.Attributes["RowIndex"].Value);
                    int column = int.Parse(node.Attributes["ColumnIndex"].Value);
                    cells[row, column] = component;
                }
            }
        }
        #endregion

        #region Crosstab
        private StiComponent ProcessCrosstab(XmlNode baseNode, StiReport report, StiContainer container)
        {
            StiCrossTab crosstab = new StiCrossTab();

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {

                    default:
                        //ThrowError(baseNode.Name, "#" + node.Name);
                        if (!ProcessCommonComponentProperties(node, crosstab)) ThrowError(node.Name, "#" + node.Name);
                        break;
                }
            }

            container.Components.Add(crosstab);
            return crosstab;
        }
        #endregion

        #endregion

        #region Groups
        private void ProcessGroups(XmlNode baseNode, StiReport report, StiContainer container)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                ProcessGroup(node, report, container);
            }
        }

        private void ProcessGroup(XmlNode baseNode, StiReport report, StiContainer container)
        {

            //foreach (XmlNode node in baseNode.Attributes)
            //{
            //    switch (node.Name)
            //    {
            //        case "Name":
            //            band.Name = node.Value;
            //            break;

            //        default:
            //            ThrowError(baseNode.Name, "#" + node.Name);
            //            break;
            //    }
            //}

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupHeader":
                        ProcessItems(node, report, groupHeadersList);
                        break;

                    case "GroupFooter":
                        ProcessItems(node, report, groupFootersList);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

        }
        #endregion

        #region Style
        private void ProcessStyle(XmlNode baseNode, StiComponent component)
        {
            IStiBrush brush = component as IStiBrush;
            IStiTextBrush textBrush = component as IStiTextBrush;
            IStiTextHorAlignment horAlign = component as IStiTextHorAlignment;
            IStiVertAlignment vertAlign = component as IStiVertAlignment;
            IStiFont font = component as IStiFont;
            IStiBorder border = component as IStiBorder;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Visible":
                        if (node.Value == "True") component.Enabled = true;
                        if (node.Value == "False") component.Enabled = false;
                        break;

                    case "BackgroundColor":
                        if (brush != null)
                        {
                            brush.Brush = new StiSolidBrush(ParseStyleColor(node.Value));
                        }
                        break;

                    case "Color":
                        if (textBrush != null)
                        {
                            textBrush.TextBrush = new StiSolidBrush(ParseStyleColor(node.Value));
                        }
                        break;

                    case "TextAlign":
                        if (horAlign != null)
                        {
                            if (node.Value == "Left") horAlign.HorAlignment = StiTextHorAlignment.Left;
                            else if (node.Value == "Center") horAlign.HorAlignment = StiTextHorAlignment.Center;
                            else if (node.Value == "Right") horAlign.HorAlignment = StiTextHorAlignment.Right;
                        }
                        break;

                    case "VerticalAlign":
                        if (vertAlign != null)
                        {
                            if (node.Value == "Top") vertAlign.VertAlignment = StiVertAlignment.Top;
                            else if (node.Value == "Middle") vertAlign.VertAlignment = StiVertAlignment.Center;
                            else if (node.Value == "Bottom") vertAlign.VertAlignment = StiVertAlignment.Bottom;
                        }
                        break;

                    //case "Type":
                    //ignored or not implemented yet
                    //break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Font":
                        if (font != null) ProcessFont(node, component);
                        break;

                    case "BorderStyle":
                        if (border != null) ProcessBorderStyle(node, component);
                        break;

                    case "BorderColor":
                        if (border != null) ProcessBorderColor(node, component);
                        break;

                    case "BorderWidth":
                        if (border != null) ProcessBorderWidth(node, component);
                        break;

                    case "Padding":
                        if (component is StiBarCode) break;
                        ThrowError(baseNode.Name, node.Name);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

        }

        private void ProcessFont(XmlNode baseNode, StiComponent component)
        {
            IStiFont font = component as IStiFont;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        font.Font = StiFontUtils.ChangeFontName(font.Font, node.Value);
                        break;

                    case "Style":
                        string[] fontStyles = node.Value.Split(new char[] { ',' });
                        foreach (string fontStyle in fontStyles)
                        {
                            if (fontStyle.Trim() == "Bold") font.Font = StiFontUtils.ChangeFontStyleBold(font.Font, true);
                            if (fontStyle.Trim() == "Italic") font.Font = StiFontUtils.ChangeFontStyleItalic(font.Font, true);
                            if (fontStyle.Trim() == "Underline") font.Font = StiFontUtils.ChangeFontStyleUnderline(font.Font, true);
                            if (fontStyle.Trim() == "Strikeout") font.Font = StiFontUtils.ChangeFontStyleStrikeout(font.Font, true);
                        }
                        break;

                    case "Size":
                        font.Font = StiFontUtils.ChangeFontSize(font.Font, (float)ToHi(node.Value));
                        break;

                    case "Bold":
                        font.Font = StiFontUtils.ChangeFontStyleBold(font.Font, true);
                        break;
                    case "Italic":
                        font.Font = StiFontUtils.ChangeFontStyleItalic(font.Font, true);
                        break;
                    case "Underline":
                        font.Font = StiFontUtils.ChangeFontStyleUnderline(font.Font, true);
                        break;
                    case "Strikeout":
                        font.Font = StiFontUtils.ChangeFontStyleStrikeout(font.Font, true);
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }
        }


        private void CheckForAdvancedBorder(IStiBorder border)
        {
            if (!(border.Border is StiAdvancedBorder))
            {
                StiBorder oldBorder = border.Border;
                border.Border = new StiAdvancedBorder();
                border.Border.Color = oldBorder.Color;
                border.Border.Style = oldBorder.Style;
                border.Border.Size = oldBorder.Size;
                border.Border.Side = oldBorder.Side;
            }
        }

        private void ProcessBorderStyle(XmlNode baseNode, StiComponent component)
        {
            IStiBorder border = component as IStiBorder;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Default":
                        if (node.Value != "None") border.Border.Side = StiBorderSides.All;
                        border.Border.Style = ParseBorderStyle(node.Value);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Left":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).LeftSide.Style = ParseBorderStyle(node.Value);
                        break;

                    case "Right":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).RightSide.Style = ParseBorderStyle(node.Value);
                        break;

                    case "Top":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).TopSide.Style = ParseBorderStyle(node.Value);
                        break;

                    case "Bottom":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).BottomSide.Style = ParseBorderStyle(node.Value);
                        break;

                    case "Default":
                        //ignore
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

        }

        private void ProcessBorderColor(XmlNode baseNode, StiComponent component)
        {
            IStiBorder border = component as IStiBorder;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Default":
                        border.Border.Color = ParseStyleColor(node.Value);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Left":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).LeftSide.Color = ParseStyleColor(node.Value);
                        break;

                    case "Right":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).RightSide.Color = ParseStyleColor(node.Value);
                        break;

                    case "Top":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).TopSide.Color = ParseStyleColor(node.Value);
                        break;

                    case "Bottom":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).BottomSide.Color = ParseStyleColor(node.Value);
                        break;

                    case "Default":
                        //ignore
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }
        }

        private void ProcessBorderWidth(XmlNode baseNode, StiComponent component)
        {
            IStiBorder border = component as IStiBorder;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Default":
                        border.Border.Size = ToHi(node.Value);
                        break;
                }
            }

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Left":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).LeftSide.Size = ToHi(node.Value);
                        break;

                    case "Right":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).RightSide.Size = ToHi(node.Value);
                        break;

                    case "Top":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).TopSide.Size = ToHi(node.Value);
                        break;

                    case "Bottom":
                        CheckForAdvancedBorder(border);
                        (border.Border as StiAdvancedBorder).BottomSide.Size = ToHi(node.Value);
                        break;

                    case "Default":
                        //ignore
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }
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

        private static StiPenStyle ParseBorderStyle(string styleName)
        {
            switch (styleName)
            {
                case "None":
                    return StiPenStyle.None;
                case "Dashed":
                    return StiPenStyle.Dash;
                case "Dotted":
                    return StiPenStyle.Dot;
                case "Double":
                    return StiPenStyle.Double;
            }
            return StiPenStyle.Solid;
        }
        #endregion

        #region Utils
        private static string ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

        private double ToHi(string strValue)
        {
            string stSize = strValue.Trim();
            //double factor = 1 / 14.4;
            double factor = 1;
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
            uint colorValue = (uint)int.Parse(node.Value);
            return Color.FromArgb(0xFF, (int)(colorValue & 0xFF), (int)((colorValue >> 8) & 0xFF), (int)((colorValue >> 16) & 0xFF));
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
        #endregion

        #region Methods.Helper
        public string ConvertExpression(string baseExpression, StiReport report)
        {
            int pos = -1;
            while ((pos = baseExpression.IndexOf("Parameters.")) != -1)
            {
                int pos2 = pos + 11;
                int len = 0;
                while (pos2 + len < baseExpression.Length && (char.IsLetterOrDigit(baseExpression[pos2 + len]) || baseExpression[pos2 + len] == '_')) len++;
                if (len > 0)
                {
                    string name = baseExpression.Substring(pos2, len);
                    if (baseExpression.Substring(pos2 + len).StartsWith(".Value")) baseExpression = baseExpression.Remove(pos2 + len, 6);
                    baseExpression = baseExpression.Remove(pos, 11);
                    if (!report.Dictionary.Variables.Contains(name))
                    {
                        StiVariable var = new StiVariable("Parameters", name, typeof(string));
                        var.RequestFromUser = true;
                        report.Dictionary.Variables.Add(var);
                    }
                }
            }
            while ((pos = baseExpression.IndexOf("Fields.")) != -1)
            {
                int pos2 = pos + 7;
                int len = 0;
                while (pos2 + len < baseExpression.Length && (char.IsLetterOrDigit(baseExpression[pos2 + len]) || baseExpression[pos2 + len] == '_')) len++;
                if (len > 0)
                {
                    string fieldName = currentDataSourceName + "." + baseExpression.Substring(pos2, len);
                    baseExpression = baseExpression.Remove(pos, 7);
                    baseExpression = baseExpression.Insert(pos, currentDataSourceName + ".");
                    if (!fieldsNames.ContainsKey(fieldName))
                    {
                        fields.Add(fieldName);
                        fieldsNames[fieldName] = null;
                    }
                }
            }
            baseExpression = baseExpression.Replace("ReportItem.DataObject", mainDataSourceName);
            baseExpression = baseExpression.Replace("Now()", "Today");
            return baseExpression;
        }
        #endregion

        #region Methods.Import
        public static StiImportResult Import(byte[] bytes)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US", false);

                ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

                var report = new StiReport();
                var errors = new ArrayList();
                var doc = new XmlDocument();

                using (var stream = new MemoryStream(bytes))
                {
                    doc.Load(stream);
                }

                var helper = new StiTelerikReportsHelper();
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
