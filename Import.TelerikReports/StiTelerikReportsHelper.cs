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
using Stimulsoft.Report.BarCodes;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components.TextFormats;
using Stimulsoft.Report.CrossTab;
using Stimulsoft.Base.Zip;
using Stimulsoft.Base;

namespace Stimulsoft.Report.Import
{
    public class StiTelerikReportsHelper
    {
        #region Fields
        private List<string> errorList = null;
        private ArrayList fields = new ArrayList();
        private Hashtable fieldsNames = new Hashtable();
        private string currentDataSourceName = "ds";
        private string mainDataSourceName = "ds";
        private StiPanel groupHeadersList = new StiPanel();
        private StiPanel groupFootersList = new StiPanel();
        private Hashtable styles = new Hashtable();
        private List<string> summaryTypeCollection = new List<string>();
        private List<string> clonnedDataSourceNames = new List<string>();
        private Dictionary<string, byte[]> files = new Dictionary<string, byte[]>();
        #endregion

        #region Root node
        public void ProcessFile(byte[] bytes, StiReport report, List<string> errorList)
        {
            #region Check for packed format
            if (bytes[0] == 'P' && bytes[1] == 'K' && bytes[2] == 3 && bytes[3] == 4)
            {
                //packed format
                using (var zipReader = new StiZipReader20(new MemoryStream(bytes)))
                {
                    foreach (var entry in zipReader.Entries)
                    {
                        if (entry.FullName == "definition.xml")
                        {
                            Stream defStream = zipReader.GetEntryStream(entry);
                            byte[] buf = new byte[entry.Size];
                            defStream.Read(buf, 0, buf.Length);
                            defStream.Close();

                            bytes = buf;
                        }
                        if (entry.FullName.StartsWith(@"Images/"))
                        {
                            Stream bufStream = zipReader.GetEntryStream(entry);
                            byte[] buf = new byte[entry.Size];
                            bufStream.Read(buf, 0, buf.Length);
                            bufStream.Close();

                            files.Add(entry.FullName, buf);
                        }
                    }
                }
            }
            #endregion

            var doc = new XmlDocument();

            using (var stream = new MemoryStream(bytes))
            {
                doc.Load(stream);
            }
            var rootNode = doc.DocumentElement;

            this.errorList = errorList;
            this.fields.Clear();
            this.fieldsNames = new Hashtable();
            report.ReportUnit = StiReportUnitType.HundredthsOfInch;
            StiPage page = report.Pages[0];
            page.UnlimitedBreakable = false;

            if (rootNode.Attributes["Name"] != null)
                report.ReportName = rootNode.Attributes["Name"].Value;

            bool pageSettingsUsed = false;
            bool dataSourceUsed = false;

            //first pass
            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "PageSettings":
                        pageSettingsUsed = ProcessPageSettings(node, report, page);
                        break;

                    case "Style":
                        ProcessStyle(node, page);
                        break;

                    case "DataSource":
                    case "DataSources":
                        dataSourceUsed = ProcessDataSources(node, report);
                        mainDataSourceName = currentDataSourceName;
                        break;

                    case "ReportParameters":
                        ProcessReportParameters(node, report);
                        break;

                    case "Groups":
                        ProcessGroups(node, report, page);
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

            //second pass
            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Items":
                        ProcessItems(node, report, page);
                        break;

                    case "PageSettings":
                    case "DataSource":
                    case "DataSources":
                    case "ReportParameters":
                    case "Style":
                    case "StyleSheet":
                    case "Groups":
                        //ignore, processed on first pass
                        break;

                    default:
                        ThrowError(rootNode.Name, node.Name);
                        break;
                }
            }

            #region Check PageSetting
            if (rootNode.Attributes["Width"] != null)
            {
                if (rootNode.Attributes["Width"].Value.LastIndexOf("cm") == rootNode.Attributes["Width"].Value.Length - 2)
                    pageSettingsUsed = false;
            }

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

            #region Add images to resources
            foreach (var pair in files)
            {
                var resource = new StiResource(pair.Key.Substring(7), StiResourceType.Image, pair.Value);
                report.Dictionary.Resources.Add(resource);
            }
            #endregion

            foreach (StiComponent comp in report.GetComponents())
            {
                comp.Page = page;
            }

            page.DockToContainer();
            page.SortByPriority();
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

                    case "CsvDataSource":
                        ProcessCsvDataSource(node, report);
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

        private void ProcessCsvDataSource(XmlNode baseNode, StiReport report)
        {
            StiCsvDatabase dataBase = new StiCsvDatabase();
            dataBase.Name = "db";
            dataBase.Alias = "db";
            StiCsvSource dataSource = new StiCsvSource();
            dataSource.Name = "ds";
            dataSource.Alias = "ds";
            dataSource.NameInSource = "db";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Name":
                        dataSource.Name = node.Value;
                        dataSource.Alias = node.Value;
                        dataSource.NameInSource = "db" + node.Value;
                        dataBase.Name = "db" + node.Value;
                        dataBase.Alias = "db" + node.Value;
                        currentDataSourceName = node.Value;
                        break;

                    case "RecordSeparators":

                        break;

                    case "FieldSeparators":

                        break;

                    case "HasHeaders":

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
                    case "Source":
                        foreach (XmlNode nodee in node.ChildNodes)
                        {
                            switch (nodee.Name)
                            {
                                case "Uri":
                                    foreach (XmlNode nodeee in nodee.Attributes)
                                    {
                                        switch (nodeee.Name)
                                        {
                                            case "Path":
                                                dataBase.PathData = nodeee.Value;

                                                break;
                                        }
                                    }
                                    break;
                            }
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            dataSource.NameInSource = dataBase.PathData;
            if (dataSource.NameInSource.StartsWith("file:///"))
                dataSource.NameInSource = dataSource.NameInSource.Substring(8);
            if (dataSource.NameInSource.StartsWith("file://"))
                dataSource.NameInSource = dataSource.NameInSource.Substring(7);

            report.Dictionary.DataSources.Add(dataSource);

            var dataSource2 = dataSource.Clone() as StiCsvSource;
            dataSource2.Name += "2";
            clonnedDataSourceNames.Add(dataSource2.Name);
            dataSource2.Alias += "2";
            report.Dictionary.DataSources.Add(dataSource2);
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
        private void ProcessStyleSheet(XmlNode baseNode, StiReport report, StiContainer container)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "StyleRule":
                        ProcessStyleRule(node, report, container);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessStyleRule(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var component = new StiText();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Style":
                        ProcessStyle(node, component);
                        break;

                    case "Selectors":
                        ProcessSelectors(node, component);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            var style = new StiStyle
            {
                Name = component.Name,
                Brush = component.Brush,
                Font = component.Font,
                TextBrush = component.TextBrush,
                HorAlignment = component.HorAlignment,
                VertAlignment = component.VertAlignment,
                Border = component.Border
            };

            if (!string.IsNullOrEmpty(style.Name))
                styles.Add(style.Name, style);
        }

        private void ProcessSelectors(XmlNode baseNode, StiComponent component)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "StyleSelector":
                        foreach (XmlNode nodee in node.Attributes)
                        {
                            switch (nodee.Name)
                            {
                                case "StyleName":
                                    component.Name = nodee.Value;
                                    break;

                                case "Type":
                                    break;

                                default:
                                    ThrowError(baseNode.Name, "#" + nodee.Name);
                                    break;
                            }
                        }
                        break;


                    case "DescendantSelector":
                        foreach (XmlNode nodee in node.ChildNodes)
                        {
                            switch (nodee.Name)
                            {
                                case "Selectors":
                                    ProcessSelectors(nodee, component);
                                    break;
                            }
                        }
                        break;

                    case "TypeSelector":
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

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

            if (sectionType == "DetailSection")
                band.CanBreak = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "ColumnCount":
                        if (band is StiDataBand) (band as StiDataBand).Columns = int.Parse(node.Value);
                        break;

                    case "PageBreak":
                        if (band is IStiPageBreak pageBreak)
                        {
                            if (node.Value == "After") pageBreak.NewPageAfter = true;
                            else if (node.Value == "Before") pageBreak.NewPageBefore = true;
                            else ThrowError(node.Name, "#" + node.Name + "." + node.Value);
                        }
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
                    case "StyleName":
                        var styleName = node.Value;

                        if (!string.IsNullOrWhiteSpace(styleName) && styles.ContainsKey(styleName))
                            panel.ComponentStyle = styleName;

                        break;

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
#if DEBUG
                        //only in debug, because the expression can contain all sorts of crap that has nothing to do with the name
                        var alias = node.Value;
                        var valueParts = node.Value.Split('.');
                        if (valueParts.Length > 0)
                            alias = valueParts[valueParts.Length - 1].Replace("(", "").Replace(")", "");
                        component.Alias = alias;
#endif

                        if (node.Value.StartsWith("="))
                        {
                            component.Text = "{" + ConvertExpression(node.Value.Substring(1), report) + "}";
                        }
                        else
                        {
                            component.Text = node.Value;
                        }

                        var newTextValue = component.Text.ToString().Replace("'", "\"");
                        component.Text = newTextValue;

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

                    case "StyleName":
                        var styleName = node.Value;

                        if (!string.IsNullOrWhiteSpace(styleName) && styles.ContainsKey(styleName))
                            component.ComponentStyle = styleName;

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

                    case "Value":
                        if (node.Value.StartsWith("="))
                        {
                            component.ImageURL = new StiImageURLExpression("{" + ConvertExpression(node.Value.Substring(1), report) + "}");
                        }
                        else
                        {
                            component.ImageURL = new StiImageURLExpression(node.Value);
                        }

                        var newPictureValue = component.ImageURL.ToString().Replace("'", "\"");
                        component.ImageURL = new StiImageURLExpression(newPictureValue);

                        break;

                    case "Image":
                        if (node.Value.StartsWith(@"/Images/"))
                        {
                            component.ImageURL = new StiImageURLExpression("resource://" + node.Value.Substring(8));
                        }
                        break;

                    //case "CanShrink":
                    //    if (node.Value == "1") component.CanShrink = true;
                    //    break;

                    //case "Type":
                    //ignored or not implemented yet
                    //break;

                    case "StyleName":
                        var styleName = node.Value;

                        if (!string.IsNullOrWhiteSpace(styleName) && styles.ContainsKey(styleName))
                            component.ComponentStyle = styleName;

                        break;

                    case "MimeType":
                        //component.MimeType = node.Value;
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
                    case "Value":
                        component.ImageBytes = Convert.FromBase64String(node.ChildNodes[0].Value);
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

                    case "StyleName":
                        var styleName = node.Value;

                        if (!string.IsNullOrWhiteSpace(styleName) && styles.ContainsKey(styleName))
                            component.ComponentStyle = styleName;

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
            var mainTableSorting = "";

            StiPanel table = new StiPanel();
            table.CanGrow = true;
            table.CanBreak = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "DataSourceName":

                        break;

                    ////case "Name":
                    ////    band.Name = node.Value;
                    ////    break;

                    default:
                        //ThrowError(baseNode.Name, "#" + node.Name);
                        if (!ProcessCommonComponentProperties(node, table)) ThrowError(node.Name, "#" + node.Name);
                        break;
                }
            }

            var rowGroupCells = new List<StiComponent>();
            var columnGroupCells = new List<StiComponent>();
            StiComponent[,] cornerCells = null;

            foreach (XmlNode nodee in baseNode.ChildNodes)
            {
                switch (nodee.Name)
                {
                    case "RowGroups":
                        ProcessRowGroups(nodee, report, rowGroupCells);
                        break;

                    case "ColumnGroups":
                        var storedCurrentDataSourceName = currentDataSourceName;
                        currentDataSourceName += "2";

                        ProcessColumnGroups(nodee, report, columnGroupCells);

                        currentDataSourceName = storedCurrentDataSourceName;
                        break;

                    case "Corner":
                        cornerCells = ProcessCorner(nodee, report, table);
                        break;

                    case "Sortings":
                        mainTableSorting = ProcessSortings(nodee, report, table);
                        break;

                    default:
                        ThrowError(baseNode.Name, nodee.Name);
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
                        ProcessTableBody(node, report, table, rowGroupCells, columnGroupCells, cornerCells, mainTableSorting);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            var tableHeight = 0d;

            foreach (var comp in table.Components)
            {
                var tableComponent = comp as StiComponent;
                tableComponent.Linked = true;
                tableHeight += tableComponent.Height;
            }

            table.Height = tableHeight;

            if (columnGroupCells.Count == 0 && report.Dictionary.DataSources.Count > 0 && clonnedDataSourceNames.Count > 0)
            {
                var correctDataSourseCollection = new List<StiDataSource>();

                foreach (var dataSource in report.Dictionary.DataSources)
                {
                    var data = dataSource as StiDataSource;

                    if (!clonnedDataSourceNames.Contains(data.Name))
                        correctDataSourseCollection.Add(data);
                }

                report.Dictionary.DataSources.Clear();

                foreach (var data in correctDataSourseCollection)
                    report.Dictionary.DataSources.Add(data);
            }

            container.Components.Add(table);

            currentDataSourceName = storeDataSource;
        }

        private StiComponent[,] ProcessCorner(XmlNode baseNode, StiReport report, StiContainer container)
        {
            StiComponent[,] cornerCells = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Columns":
                    case "Rows":
                        break;

                    case "Cells":
                        cornerCells = ProcessTableCornerCells(node, report);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return cornerCells;
        }

        private void ProcessTableBody(XmlNode baseNode, StiReport report, StiPanel container, List<StiComponent> rowGroupCells
            , List<StiComponent> columnGroupCells, StiComponent[,] cornerCells, string mainTableSorting)
        {
            List<double> columnWidths = new List<double>();
            List<double> rowHeights = new List<double>();
            List<int> crossGroupHeaderBandIndexCollection = new List<int>();
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
                }
            }

            var finalColumnWidth = new List<double>();

            foreach (var rowgroupComponent in rowGroupCells)
            {
                finalColumnWidth.Add(rowgroupComponent.Width);
            }

            foreach (var columnWidth in columnWidths)
            {
                finalColumnWidth.Add(columnWidth);
            }

            columnWidths = finalColumnWidth;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Columns":
                    case "Rows":
                        break;

                    case "Cells":
                        cells = new StiComponent[rowHeights.Count, columnWidths.Count];

                        for (int index = 0; index < rowGroupCells.Count; index++)
                        {
                            cells[0, index] = rowGroupCells[index];
                        }

                        ProcessTableBodyCells(node, report, cells, rowGroupCells.Count);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            StiHeaderBand headerBand = new StiHeaderBand();
            headerBand.Name = container.Name + "HeaderBand";

            var componentCollection = new List<StiComponent>();
            var crossDataBandInsideIndex = 0;

            if (cornerCells != null && cornerCells.Length > 0)
            {
                var crosDataBand = new StiCrossDataBand();
                crosDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                crossDataBandInsideIndex++;

                var x = 0d;

                foreach (var cell in cornerCells)
                {
                    if (cell != null)
                    {
                        cell.ClientRectangle = new RectangleD(x, 0, cell.Width, cell.Height);
                        cell.Linked = true;
                        crosDataBand.Components.Add(cell);
                        crosDataBand.Width += cell.Width;

                        if (crosDataBand.Height < cell.Height)
                            crosDataBand.Height = cell.Height;

                        x += cell.Width;
                    }
                }

                headerBand.Components.Add(crosDataBand);

                crosDataBand = new StiCrossDataBand();
                crosDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                crossDataBandInsideIndex++;
            }

            if (columnGroupCells.Count > 0 && columnGroupCells.Count == columnWidths.Count)
            {
                double positionX = 0;
                double headerHeight = 0;

                for (int indexColumn = 0; indexColumn < columnWidths.Count; indexColumn++)
                {
                    StiComponent component = columnGroupCells[indexColumn];

                    if (component != null)
                    {
                        if (component is StiCrossGroupHeaderBand)
                        {
                            if (componentCollection.Count > 0)
                            {
                                var crosDataBand = new StiCrossDataBand();
                                crosDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                                crossDataBandInsideIndex++;

                                var x = 0d;

                                foreach (var comp in componentCollection)
                                {
                                    comp.ClientRectangle = new RectangleD(x, 0, comp.Width, comp.Height);
                                    crosDataBand.Components.Add(comp);
                                    crosDataBand.Width += comp.Width;

                                    if (crosDataBand.Height < comp.Height)
                                        crosDataBand.Height = comp.Height;

                                    x += comp.Width;
                                }

                                headerBand.Components.Add(crosDataBand);

                                crosDataBand = new StiCrossDataBand();
                                crosDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                                crossDataBandInsideIndex++;

                                componentCollection.Clear();
                            }

                            headerHeight = component.Height;
                            component.Left = positionX;
                            component.Top = 0;
                            component.Linked = true;
                            headerBand.Components.Add(component);

                            crossGroupHeaderBandIndexCollection.Add(indexColumn);

                            StiCrossDataBand crossDataBand = new StiCrossDataBand();
                            crossDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                            crossDataBandInsideIndex++;
                            crossDataBand.Width = 0;
                            crossDataBand.Height = 0;
                            crossDataBand.DataSourceName = $"{currentDataSourceName}2";
                            crossDataBand.Linked = true;
                            headerBand.Components.Add(crossDataBand);

                        }
                        else if (component is StiText)
                        {
                            componentCollection.Add(component);
                        }
                    }

                    positionX += columnWidths[indexColumn];
                }

                headerBand.Height = headerHeight;
            }

            if (componentCollection.Count > 0)
            {
                var crosDataBand = new StiCrossDataBand();
                crosDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                crossDataBandInsideIndex++;

                var x = 0d;

                foreach (var comp in componentCollection)
                {
                    comp.ClientRectangle = new RectangleD(x, 0, comp.Width, comp.Height);
                    crosDataBand.Components.Add(comp);
                    crosDataBand.Width += comp.Width;

                    if (crosDataBand.Height < comp.Height)
                        crosDataBand.Height = comp.Height;

                    x += comp.Width;
                }

                headerBand.Components.Add(crosDataBand);

                crosDataBand = new StiCrossDataBand();
                crosDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                crossDataBandInsideIndex++;

                componentCollection.Clear();
            }

            if (headerBand.Components.Count > 0)
                container.Components.Add(headerBand);

            StiDataBand band = new StiDataBand();
            band.Name = container.Name + "DataBand";
            band.DataSourceName = currentDataSourceName;

            var sorting = "";
            var direction = "";

            if (!string.IsNullOrEmpty(mainTableSorting))
            {
                var splittedSorting = mainTableSorting.Split('/');

                if (splittedSorting.Length > 1)
                {
                    sorting = splittedSorting[0];

                    if (splittedSorting[1] == "Asc")
                        direction = "ASC";

                    else if (splittedSorting[1] == "Desc")
                        direction = "DESC";
                }
            }

            if (!string.IsNullOrEmpty(sorting))
                band.Sort = new string[] { direction, sorting };

            double posY = 0;

            componentCollection = new List<StiComponent>();

            for (int indexRow = 0; indexRow < rowHeights.Count; indexRow++)
            {
                double posX = 0;
                for (int indexColumn = 0; indexColumn < columnWidths.Count; indexColumn++)
                {
                    StiComponent component = cells[indexRow, indexColumn];

                    if (component != null)
                    {
                        component.ClientRectangle = new RectangleD(0, 0, component.Width, component.Height);
                        component.Linked = true;

                        if (crossGroupHeaderBandIndexCollection.Contains(indexColumn))
                        {
                            if (componentCollection.Count > 0)
                            {
                                var crossDataBand = new StiCrossDataBand();
                                crossDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                                crossDataBandInsideIndex++;

                                var x = 0d;

                                foreach (var comp in componentCollection)
                                {
                                    comp.ClientRectangle = new RectangleD(x, 0, comp.Width, comp.Height);
                                    crossDataBand.Components.Add(comp);
                                    crossDataBand.Width += comp.Width;

                                    if (crossDataBand.Height < comp.Height)
                                        crossDataBand.Height = comp.Height;

                                    x += comp.Width;
                                }

                                if (crossDataBand.Height > band.Height)
                                    band.Height = crossDataBand.Height;

                                band.Components.Add(crossDataBand);

                                crossDataBand = new StiCrossDataBand();
                                crossDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                                crossDataBandInsideIndex++;

                                componentCollection.Clear();
                            }

                            var crossGroupHeader = new StiCrossGroupHeaderBand();
                            crossGroupHeader.ClientRectangle = component.ClientRectangle;

                            var columnGroupHeaderBand = columnGroupCells[indexColumn] as StiCrossGroupHeaderBand;

                            crossGroupHeader.Condition = columnGroupHeaderBand.Condition;
                            crossGroupHeader.Name = $"crossHeaderBand{component.Name}";
                            crossGroupHeader.Height = component.Height;
                            crossGroupHeader.Left = posX;
                            crossGroupHeader.Top = posY;
                            crossGroupHeader.Width = component.Width;
                            crossGroupHeader.Components.Add(component);
                            crossGroupHeader.Linked = true;
                            band.Height = crossGroupHeader.Height;

                            if (crossGroupHeader.Height > band.Height)
                                band.Height = crossGroupHeader.Height;

                            band.Components.Add(crossGroupHeader);

                            StiCrossDataBand crossDataBand2 = new StiCrossDataBand();
                            crossDataBand2.Name = crossGroupHeader.Name + "CrossDataBand";
                            crossDataBand2.Width = 0;
                            crossDataBand2.Height = 0;
                            crossDataBand2.DataSourceName = $"{currentDataSourceName}2";
                            crossDataBand2.Linked = true;
                            band.Components.Add(crossDataBand2);
                        }
                        else
                        {
                            if (component is StiText)
                            {
                                component.Linked = true;

                                componentCollection.Add(component);
                            }
                        }
                    }
                    posX += columnWidths[indexColumn];
                }

                posY += rowHeights[indexRow];
            }

            if (componentCollection.Count > 0)
            {
                var crossDataBand = new StiCrossDataBand();
                crossDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                crossDataBandInsideIndex++;

                var x = 0d;

                foreach (var comp in componentCollection)
                {
                    comp.ClientRectangle = new RectangleD(x, 0, comp.Width, comp.Height);
                    crossDataBand.Components.Add(comp);
                    crossDataBand.Width += comp.Width;

                    if (crossDataBand.Height < comp.Height)
                        crossDataBand.Height = comp.Height;

                    x += comp.Width;
                }

                if (crossDataBand.Height > band.Height)
                    band.Height = crossDataBand.Height;

                band.Components.Add(crossDataBand);

                crossDataBand = new StiCrossDataBand();
                crossDataBand.Name = $"{container.Name}CrossDataBand{crossDataBandInsideIndex}";
                crossDataBandInsideIndex++;

                componentCollection.Clear();
            }

            container.Components.Add(band);
        }

        private void ProcessRowGroups(XmlNode baseNode, StiReport report, List<StiComponent> cells)
        {
            StiPanel tempPanel = new StiPanel();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TableGroup":
                        foreach (XmlNode nodee in node.ChildNodes)
                        {
                            StiComponent component = null;
                            switch (nodee.Name)
                            {
                                case "ReportItem":
                                    component = ProcessReportItem(nodee, report, tempPanel);
                                    break;

                                case "ChildGroups":
                                    ProcessRowGroups(nodee, report, cells);
                                    break;
                            }
                            if (component != null)
                            {
                                cells.Add(component);
                            }
                        }

                        break;
                }
            }
        }

        private void ProcessColumnGroups(XmlNode baseNode, StiReport report, List<StiComponent> cells)
        {
            StiPanel tempPanel = new StiPanel();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TableGroup":
                        StiComponent component = null;
                        StiCrossGroupHeaderBand crossGroupHeader = null;

                        foreach (XmlNode nodee in node.ChildNodes)
                        {
                            switch (nodee.Name)
                            {
                                case "ReportItem":
                                    component = ProcessReportItem(nodee, report, tempPanel);
                                    break;

                                case "Groupings":
                                    crossGroupHeader = ProcessGroupings(nodee, report, cells);
                                    break;

                                case "ChildGroups":
                                    ProcessColumnGroups(nodee, report, cells);
                                    break;
                            }

                        }

                        if (crossGroupHeader != null && component != null)
                        {
                            if (component is StiPanel panel)
                            {
                                component.ClientRectangle = new RectangleD(0, 0, component.Width, component.Height);
                                crossGroupHeader.Name = $"groupHeaderBand{component.Name}";

                                foreach (var comp in panel.Components)
                                {
                                    var com = comp as StiComponent;

                                    crossGroupHeader.Components.Add(com);
                                }

                                crossGroupHeader.Height = component.Height;
                                crossGroupHeader.Width = component.Width;

                                cells.Add(crossGroupHeader);
                            }
                            else
                            {
                                component.ClientRectangle = new RectangleD(0, 0, component.Width, component.Height);
                                crossGroupHeader.Name = $"groupHeaderBand{component.Name}";
                                crossGroupHeader.Height = component.Height;
                                crossGroupHeader.Width = component.Width;
                                crossGroupHeader.Components.Add(component);
                                cells.Add(crossGroupHeader);
                            }
                        }
                        else if (crossGroupHeader == null && component != null)
                        {
                            if (component is StiPanel panel)
                            {
                                var crossHeader = new StiCrossHeaderBand();
                                crossHeader.Name = $"crossHeaderBand{component.Name}";

                                foreach (var comp in panel.Components)
                                {
                                    var com = comp as StiComponent;

                                    crossHeader.Components.Add(com);
                                }

                                crossHeader.Height = component.Height;
                                crossHeader.Width = component.Width;

                                cells.Add(crossHeader);
                            }
                            else
                            {
                                cells.Add(component);
                            }
                        }

                        break;
                }
            }
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

        private void ProcessTableBodyCells(XmlNode baseNode, StiReport report, StiComponent[,] cells, int rowGroupCellsCount)
        {
            StiPanel tempPanel = new StiPanel();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                StiComponent component = null;
                switch (node.Name)
                {
                    case "TableCell":
                        component = ProcessTableCell(node, report, tempPanel);
                        break;

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
                    int column = int.Parse(node.Attributes["ColumnIndex"].Value) + rowGroupCellsCount;
                    cells[row, column] = component;
                }
            }
        }

        private StiComponent[,] ProcessTableCornerCells(XmlNode baseNode, StiReport report)
        {
            StiPanel tempPanel = new StiPanel();
            var rowCount = 0;
            var columnCount = 0;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                StiComponent component = null;
                switch (node.Name)
                {
                    case "TableCell":
                        component = ProcessTableCell(node, report, tempPanel);
                        break;

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

                    if (row > rowCount)
                        rowCount = row;

                    int column = int.Parse(node.Attributes["ColumnIndex"].Value);

                    if (column > columnCount)
                        columnCount = column;
                }
            }

            var cells = new StiComponent[rowCount + 1, columnCount + 1];

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                StiComponent component = null;
                switch (node.Name)
                {
                    case "TableCell":
                        component = ProcessTableCell(node, report, tempPanel);
                        break;

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

            return cells;
        }

        private StiComponent ProcessTableCell(XmlNode baseNode, StiReport report, StiPanel tempPanel)
        {
            StiComponent component = null;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ReportItem":
                        component = ProcessReportItem(node, report, tempPanel);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return component;
        }

        private StiComponent ProcessReportItem(XmlNode baseNode, StiReport report, StiPanel tempPanel)
        {
            StiComponent component = null;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
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

                    case "Panel":
                        component = ProcessPanel(node, report, tempPanel);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return component;
        }
        #endregion

        #region Crosstab
        private void ProcessCrosstab(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var crossTable = new StiCrossTab();
            crossTable.Name = $"{container.Name}CrossTab";
            crossTable.ClientRectangle = new RectangleD(0, 0, crossTable.Width, crossTable.Height);
            crossTable.CanGrow = true;
            crossTable.DataSourceName = currentDataSourceName;


            var storeDataSource = currentDataSourceName;
            var mainTableSorting = "";
            var rowGroupCells = new List<StiComponent>();
            var columnGroupCells = new List<StiComponent>();
            var columnWidths = new List<double>();
            var rowHeights = new List<double>();
            StiComponent[,] bodyCells = null;
            StiComponent[,] cornerCells = null;

            var table = new StiPanel();
            table.CanGrow = true;

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "DataSourceName":

                        break;

                    ////case "Name":
                    ////    band.Name = node.Value;
                    ////    break;

                    default:
                        //ThrowError(baseNode.Name, "#" + node.Name);
                        if (!ProcessCommonComponentProperties(node, table)) ThrowError(node.Name, "#" + node.Name);
                        break;
                }
            }

            foreach (XmlNode nodee in baseNode.ChildNodes)
            {
                switch (nodee.Name)
                {
                    case "RowGroups":
                        ProcessRowGroups(nodee, report, rowGroupCells);
                        break;

                    case "ColumnGroups":
                        ProcessColumnGroups(nodee, report, columnGroupCells);
                        break;

                    case "Corner":
                        cornerCells = ProcessCorner(nodee, report, table);
                        break;

                    case "Sortings":
                        mainTableSorting = ProcessSortings(nodee, report, table);
                        break;

                    default:
                        ThrowError(baseNode.Name, nodee.Name);
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
                        foreach (XmlNode nodee in node.ChildNodes)
                        {
                            switch (nodee.Name)
                            {
                                case "Columns":
                                    columnWidths = ProcessTableBodyColumns(nodee, report);
                                    break;

                                case "Rows":
                                    rowHeights = ProcessTableBodyRows(nodee, report);
                                    break;

                                default:
                                    ThrowError(baseNode.Name, nodee.Name);
                                    break;
                            }
                        }

                        foreach (XmlNode nodee in node.ChildNodes)
                        {
                            switch (nodee.Name)
                            {
                                case "Cells":
                                    bodyCells = new StiComponent[rowHeights.Count, columnWidths.Count];

                                    for (int index = 0; index < rowGroupCells.Count; index++)
                                    {
                                        if (index < columnWidths.Count)
                                            bodyCells[0, index] = rowGroupCells[index];
                                    }

                                    ProcessTableBodyCells(nodee, report, bodyCells, 0);
                                    break;

                                default:
                                    ThrowError(baseNode.Name, nodee.Name);
                                    break;
                            }
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            currentDataSourceName = storeDataSource;

            #region CrossTab Build
            var rowFields = new List<StiText>();
            var columnFields = new List<StiText>();
            var dataFields = new List<StiText>();

            foreach (var roww in rowGroupCells)
            {
                if (roww is StiText textRow)
                    rowFields.Add(textRow);
            }

            foreach (var columnn in columnGroupCells)
            {
                if (columnn is StiCrossGroupHeaderBand crossGroupHeader)
                {
                    if (crossGroupHeader.Components.Count > 0)
                    {
                        var columnText = (StiText)crossGroupHeader.Components[0];

                        columnFields.Add(columnText);
                    }
                }
            }

            foreach (var data in bodyCells)
            {
                if (data is StiText dataCell)
                    dataFields.Add(dataCell);
            }

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

            var leftTitleWidth = MeasureSize(currentDataSourceName.Replace("{", "").Replace("}", ""), rowFields[0].Font);

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

                    case "Groupings":
                        ProcessGroupings(node, report, groupHeadersList);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessGroupings(XmlNode baseNode, StiReport report, StiContainer container)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Grouping":
                        ProcessGrouping(node, report, groupHeadersList);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private StiCrossGroupHeaderBand ProcessGroupings(XmlNode baseNode, StiReport report, List<StiComponent> container)
        {
            StiCrossGroupHeaderBand crossGroupHeader = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Grouping":
                        crossGroupHeader = ProcessGrouping(node, report, container);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return crossGroupHeader;
        }

        private void ProcessGrouping(XmlNode baseNode, StiReport report, StiContainer container)
        {
            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Expression":
                        var exp = ConvertExpression(node.Value.Substring(1), report);
                        if (container != null && container.Components.Count > 0)
                        {
                            (container.Components[0] as StiGroupHeaderBand).Condition = new StiGroupConditionExpression(exp);
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }
        }

        private StiCrossGroupHeaderBand ProcessGrouping(XmlNode baseNode, StiReport report, List<StiComponent> container)
        {
            var crossGroupHeader = new StiCrossGroupHeaderBand();

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Expression":
                        var exp = "";

                        if (node.Value.StartsWith("="))
                        {
                            exp = "{" + ConvertExpression(node.Value.Substring(1), report) + "}";
                        }
                        else
                        {
                            exp = node.Value;
                        }

                        if (container != null)
                        {
                            crossGroupHeader.Condition = new StiGroupConditionExpression(exp);
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            return crossGroupHeader;
        }

        private string ProcessSortings(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var mainTableSorting = "";

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Sorting":
                        mainTableSorting = ProcessSorting(node, report, container);
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            return mainTableSorting;
        }

        private string ProcessSorting(XmlNode baseNode, StiReport report, StiContainer container)
        {
            var exp = "";
            var direction = "";

            foreach (XmlNode node in baseNode.Attributes)
            {
                switch (node.Name)
                {
                    case "Expression":
                        exp = ConvertExpression(node.Value.Substring(1), report);
                        break;

                    case "Direction":
                        direction = node.Value;
                        break;

                    default:
                        ThrowError(baseNode.Name, "#" + node.Name);
                        break;
                }
            }

            if (exp.IndexOf(" ") == 0)
                exp = exp.Remove(0, 1);

            var result = $"{exp}/{direction}";

            return result;
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
                        font.Font = ParseFont(node.Value, font.Font);
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

        private Font ParseFont(string strValue, Font font)
        {
            string stSize = strValue.Trim();

            GraphicsUnit gru = GraphicsUnit.Point;
            if (strValue.EndsWith("pt") || strValue.EndsWith("px"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                if (strValue.EndsWith("px")) gru = GraphicsUnit.Pixel;
            }
            float size = Convert.ToSingle(stSize.Replace(",", ".").Replace(".", ApplicationDecimalSeparator));

            if (size < 1) size = 1;

            return new Font(
                font.FontFamily,
                size,
                font.Style,
                gru,
                font.GdiCharSet,
                font.GdiVerticalFont);
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

        private SizeF MeasureSize(string text, Font font)
        {
            Bitmap img = new Bitmap(1, 1);
            Graphics g = Graphics.FromImage(img);
            StringFormat sf = new StringFormat(StringFormat.GenericTypographic);
            g.PageUnit = GraphicsUnit.Point;
            var result = g.MeasureString(text, font, 10000000, sf);

            return result;
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
            baseExpression = baseExpression.TrimStart();

            int pos = 0;
            while ((pos = baseExpression.IndexOf("Parameters.", pos)) != -1)
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
                    pos += len;
                }
                else
                {
                    pos = pos2;
                }
            }
            pos = 0;
            while ((pos = baseExpression.IndexOf("Fields.", pos)) != -1)
            {
                int pos2 = pos + 7;
                if (baseExpression[pos2] == '[')
                {
                    pos2++;
                    int pos3 = baseExpression.IndexOf("]", pos2);
                    if (pos3 != -1)
                    {
                        string fieldName = baseExpression.Substring(pos2, pos3 - pos2).Trim();
                        if (fieldName.StartsWith("\"") && fieldName.EndsWith("\""))
                            fieldName = StiNameValidator.CorrectName(fieldName.Substring(1, fieldName.Length - 2));
                        fieldName = currentDataSourceName + "." + fieldName;
                        baseExpression = baseExpression.Remove(pos, pos3 + 1 - pos);
                        baseExpression = baseExpression.Insert(pos, fieldName);
                        if (!fieldsNames.ContainsKey(fieldName))
                        {
                            fields.Add(fieldName);
                            fieldsNames[fieldName] = null;
                        }
                        pos += fieldName.Length;
                    }
                    else
                    {
                        pos = pos2;
                    }
                }
                else
                {
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
                        pos += fieldName.Length;
                    }
                    else
                    {
                        pos = pos2;
                    }
                }
            }

            foreach (string agr in aggregates)
            {
                if (baseExpression.StartsWith(agr))
                {
                    baseExpression = baseExpression.Substring(agr.Length).Trim();
                    if (baseExpression.StartsWith("(")) baseExpression = baseExpression.Substring(1);
                    if (baseExpression.EndsWith(")")) baseExpression = baseExpression.Remove(baseExpression.Length - 1, 1);
                    summaryTypeCollection.Add(agr);
                }
            }

            baseExpression = baseExpression.Replace("ReportItem.DataObject", mainDataSourceName);
            baseExpression = baseExpression.Replace("Now()", "Today");
            return baseExpression;
        }

        private static string[] aggregates = new string[] { "Count", "Sum", "Max", "Min", "Avg" };
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

                var helper = new StiTelerikReportsHelper();
                helper.ProcessFile(bytes, report, errors);

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
