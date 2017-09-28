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
using System.Windows.Forms;

namespace Stimulsoft.Report.Import
{
    public class StiReportSharpShooterHelper
    {
        #region Const
        private const string containerName = "-=**|container|**=-";
        private const string datasetName = "DataSetName";
        #endregion

        #region Fields
        private string currentStyleName = string.Empty;
        private string currentDataSourceName = string.Empty;
        private string currentDatabandName = string.Empty;
        private bool currentIsGroup = false;
        private Hashtable fields = new Hashtable();
        private Hashtable dataBands = new Hashtable();
        private bool importStyleNames = false;
        #endregion

        #region Methods
        public StiReport Convert(byte[] bytes, bool importStyleNames)
        {
            CultureInfo currentCulture = Application.CurrentCulture;
            try
            {
                Application.CurrentCulture = new CultureInfo("en-US", false);

                this.importStyleNames = importStyleNames;

                StiReport report = new StiReport();
                report.ReportUnit = StiReportUnitType.Centimeters;
                report.Pages.Clear();

                XmlDocument doc = new XmlDocument();
                using (var stream = new MemoryStream(bytes))
                {
                    doc.Load(stream);
                }

                XmlNodeList documentList = doc.GetElementsByTagName("root");
                foreach (XmlNode documentNode in documentList)
                {
                    foreach (XmlNode elementNode in documentNode.ChildNodes)
                    {
                        switch (elementNode.Name)
                        {
                            case "Pages":
                                ReadPages(elementNode, report);
                                break;
                            /*
                            case "Parameters":
                                ReadParameters(elementNode, report);
                                break;*/

                            case "StyleSheet":
                                ReadStyleSheet(elementNode, report);
                                break;
                                /*
                                case "GraphicsSettings":
                                    ReadGraphicsSettings(elementNode, report);
                                    break;*/
                        }
                    }

                    #region Add references list
                    XmlAttribute attr = documentNode.Attributes["ImportsString"];
                    if ((attr != null) && !string.IsNullOrEmpty(attr.Value))
                    {
                        List<string> added = new List<string>();
                        string[] strList = attr.Value.Split(new char[]
                        {
                        '\r', '\n'
                        });
                        foreach (string st in strList)
                        {
                            if (!string.IsNullOrEmpty(st.Trim()))
                            {
                                string dll = (st.Trim() + ".dll").ToLowerInvariant();
                                bool founded = false;
                                foreach (string refst in report.ReferencedAssemblies)
                                {
                                    if (refst.Trim().ToLowerInvariant() == dll)
                                    {
                                        founded = true;
                                    }
                                }
                                if (!founded)
                                {
                                    added.Add(st.Trim());
                                }
                            }
                        }
                        if (added.Count > 0)
                        {
                            //string[] newRefs = new string[report.ReferencedAssemblies.Length + added.Count];
                            //Array.Copy(report.ReferencedAssemblies, newRefs, report.ReferencedAssemblies.Length);
                            for (int index = 0; index < added.Count; index++)
                            {
                                //newRefs[report.ReferencedAssemblies.Length + index] = added[index] + ".dll";
                                report.Script = report.Script.Insert(0, string.Format("using {0};\r\n", added[index]));
                            }

                            //report.ReferencedAssemblies = newRefs;
                        }
                    }
                    #endregion
                }

                foreach (StiPage page in report.Pages)
                {
                    StiComponentsCollection comps = page.GetComponents();
                    foreach (StiComponent comp in comps)
                    {
                        comp.Page = page;
                        comp.Left -= page.Margins.Left;

                        //comp.Top -= page.Margins.Top;
                    }
                }

                report.ApplyStyles();

                if (documentList.Count > 0)
                {
                    report.ReportName = ReadString(documentList[0], "Name", "Report name");
                    report.ReportAlias = ReadString(documentList[0], "Title", "Report title");
                    report.ReportDescription = ReadString(documentList[0], "Description", "Report description");
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

        private void ReadPages(XmlNode dictNode, StiReport report)
        {
            foreach (XmlNode elementNode in dictNode.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "Item":
                    case "item":
                        ReadPage(elementNode, report);
                        break;
                }
            }
        }
        #endregion

        #region Methods.Read.Styles
        private void ReadStyleSheet(XmlNode dictNode, StiReport report)
        {
            foreach (XmlNode elementNode in dictNode.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "Styles":
                        ReadStyles(elementNode, report);
                        break;
                }
            }
        }

        private void ReadStyles(XmlNode dictNode, StiReport report)
        {
            foreach (XmlNode elementNode in dictNode.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "Item":
                        ReadStyle(elementNode, report);
                        break;
                }
            }
        }

        private void ReadStyle(XmlNode node, StiReport report)
        {
            StiStyle style = new StiStyle();
            report.Styles.Add(style);

            style.Name = ReadString(node, "Name", style.Name);

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "Border":
                        ReadBorder(elementNode, style);
                        break;

                    case "Fill":
                        ReadBrush(elementNode, style);
                        break;

                    case "TextFill":
                        ReadTextBrush(elementNode, style);
                        break;

                    case "Font":
                        ReadFont(elementNode, style);
                        break;
                }
            }
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

        private static bool ReadBool(XmlNode node, string name, bool defaultValue)
        {
            XmlAttribute attr = node.Attributes[name];
            if (attr != null)
                return attr.Value.ToLowerInvariant() == "true";
            else
                return defaultValue;//default value
        }

        private static double ReadDouble(XmlNode node, string name, double defaultValue)
        {
            XmlAttribute attr = node.Attributes[name];
            if (attr != null)
                return double.Parse(attr.Value);
            else
                return defaultValue;
        }

        private static int ReadInt(XmlNode node, string name, int defaultValue)
        {
            XmlAttribute attr = node.Attributes[name];
            if (attr != null)
                return int.Parse(attr.Value);
            else
                return defaultValue;
        }
        #endregion

        #region Methods.Convert
        private static double ConvertFromMillimeters(double value, StiComponent comp)
        {
            StiMillimetersUnit mmUnit = new StiMillimetersUnit();
            value = mmUnit.ConvertToHInches(value);
            return comp.Report.Unit.ConvertFromHInches(value);
        }

        private static double ConvertFromRSSUnit(double value)
        {
            return (value / 118.110237121582);// /0.96/124
        }
        #endregion

        #region Methods.Read.Properties
        private void ReadPageMargins(XmlNode node, StiPage page)
        {
            XmlAttribute attr = node.Attributes["Margins"];
            if (attr == null) return;

            string[] strs = attr.Value.Split(new char[] { ';' });

            try
            {
                page.Margins = new StiMargins(
                    Math.Round(ConvertFromRSSUnit(double.Parse(strs[2].Trim())), 2),
                    Math.Round(ConvertFromRSSUnit(double.Parse(strs[3].Trim())), 2),
                    Math.Round(ConvertFromRSSUnit(double.Parse(strs[0].Trim())), 2),
                    Math.Round(ConvertFromRSSUnit(double.Parse(strs[1].Trim())), 2));
            }
            catch
            {
            }
        }

        private void ReadPageSize(XmlNode node, StiPage page)
        {
            XmlAttribute attr = node.Attributes["Size"];
            if (attr == null) return;

            string[] strs = attr.Value.Split(new char[] { ';' });

            try
            {
                page.PageWidth = Math.Round(ConvertFromRSSUnit(double.Parse(strs[0].Trim())), 2);
                page.PageHeight = Math.Round(ConvertFromRSSUnit(double.Parse(strs[1].Trim())), 2);
            }
            catch
            {
            }
        }

        private void ReadComponentLocation(XmlNode node, StiComponent comp)
        {
            XmlAttribute attr = node.Attributes["Location"];
            if (attr == null) return;

            string[] strs = attr.Value.Split(new char[] { ';' });

            try
            {
                comp.Left = Math.Round(ConvertFromRSSUnit(double.Parse(strs[0].Trim())), 2);
                comp.Top = Math.Round(ConvertFromRSSUnit(double.Parse(strs[1].Trim())), 2);
            }
            catch
            {
            }
        }

        private void ReadComponentSize(XmlNode node, StiComponent comp)
        {
            XmlAttribute attr = node.Attributes["Size"];
            if (attr == null) return;

            string[] strs = attr.Value.Split(new char[] { ';' });

            try
            {
                comp.Width = Math.Round(ConvertFromRSSUnit(double.Parse(strs[0].Trim())), 2);
                comp.Height = Math.Round(ConvertFromRSSUnit(double.Parse(strs[1].Trim())), 2);
            }
            catch
            {
            }
        }

        private Color ReadColor(XmlNode node, string key, Color defaultColor)
        {
            ColorConverter colorConverter = new ColorConverter();
            try
            {
                if (node.Attributes[key] == null)
                    return defaultColor;

                return (Color)colorConverter.ConvertFromString(node.Attributes[key].Value);
            }
            catch
            {
            }
            return Color.Black;
        }

        private Color ReadColor(string name)
        {
            ColorConverter colorConverter = new ColorConverter();
            try
            {
                return (Color)colorConverter.ConvertFromString(name);
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
                    if (node.Attributes["TopLine"] != null ||
                        node.Attributes["BottomLine"] != null ||
                        node.Attributes["LeftLine"] != null ||
                        node.Attributes["RightLine"] != null)
                    {
                        StiAdvancedBorder border = new StiAdvancedBorder();
                        if (node.Attributes["TopLine"] != null)
                        {
                            string[] strs = node.Attributes["TopLine"].Value.Split(new char[] { ' ' });
                            border.TopSide.Size = double.Parse(strs[0]);
                            border.TopSide.Style = ReadPenStyle(strs[1]);

                            if (strs.Length == 3)
                                border.TopSide.Color = ReadColor(strs[2]);
                            else
                                border.TopSide.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                        }
                        if (node.Attributes["BottomLine"] != null)
                        {
                            string[] strs = node.Attributes["BottomLine"].Value.Split(new char[] { ' ' });
                            border.BottomSide.Size = double.Parse(strs[0]);
                            border.BottomSide.Style = ReadPenStyle(strs[1]);

                            if (strs.Length == 3)
                                border.BottomSide.Color = ReadColor(strs[2]);
                            else
                                border.BottomSide.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                        }
                        if (node.Attributes["LeftLine"] != null)
                        {
                            string[] strs = node.Attributes["LeftLine"].Value.Split(new char[] { ' ' });
                            border.LeftSide.Size = double.Parse(strs[0]);
                            border.LeftSide.Style = ReadPenStyle(strs[1]);

                            if (strs.Length == 3)
                                border.LeftSide.Color = ReadColor(strs[2]);
                            else
                                border.LeftSide.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                        }
                        if (node.Attributes["RightLine"] != null)
                        {
                            string[] strs = node.Attributes["RightLine"].Value.Split(new char[] { ' ' });
                            border.RightSide.Size = double.Parse(strs[0]);
                            border.RightSide.Style = ReadPenStyle(strs[1]);

                            if (strs.Length == 3)
                                border.RightSide.Color = ReadColor(strs[2]);
                            else
                                border.RightSide.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                        }

                        ((IStiBorder)comp).Border = border;
                    }
                    else
                    {
                        StiBorder border = ((IStiBorder)comp).Border;

                        #region Border.Lines
                        XmlAttribute attr = node.Attributes["All"];
                        if (attr != null)
                        {
                            string[] strs = attr.Value.Split(new char[] { ' ' });
                            border.Size = double.Parse(strs[0]);
                            border.Style = ReadPenStyle(strs[1]);
                            border.Side = StiBorderSides.All;

                            if (strs.Length == 3)
                                border.Color = ReadColor(strs[2]);
                            else
                                border.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                        }
                        #endregion
                    }
                }
            }
            catch
            {
            }
        }

        private void ReadBorder(XmlNode node, StiStyle style)
        {
            try
            {
                if (node.Attributes["TopLine"] != null ||
                    node.Attributes["BottomLine"] != null ||
                    node.Attributes["LeftLine"] != null ||
                    node.Attributes["RightLine"] != null)
                {
                    StiAdvancedBorder border = new StiAdvancedBorder();
                    if (node.Attributes["TopLine"] != null)
                    {
                        string[] strs = node.Attributes["TopLine"].Value.Split(new char[] { ' ' });
                        border.TopSide.Size = double.Parse(strs[0]);
                        border.TopSide.Style = ReadPenStyle(strs[1]);

                        if (strs.Length == 3)
                            border.TopSide.Color = ReadColor(strs[2]);
                        else
                            border.TopSide.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                    }
                    if (node.Attributes["BottomLine"] != null)
                    {
                        string[] strs = node.Attributes["BottomLine"].Value.Split(new char[] { ' ' });
                        border.BottomSide.Size = double.Parse(strs[0]);
                        border.BottomSide.Style = ReadPenStyle(strs[1]);

                        if (strs.Length == 3)
                            border.BottomSide.Color = ReadColor(strs[2]);
                        else
                            border.BottomSide.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                    }
                    if (node.Attributes["LeftLine"] != null)
                    {
                        string[] strs = node.Attributes["LeftLine"].Value.Split(new char[] { ' ' });
                        border.LeftSide.Size = double.Parse(strs[0]);
                        border.LeftSide.Style = ReadPenStyle(strs[1]);

                        if (strs.Length == 3)
                            border.LeftSide.Color = ReadColor(strs[2]);
                        else
                            border.LeftSide.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                    }
                    if (node.Attributes["RightLine"] != null)
                    {
                        string[] strs = node.Attributes["RightLine"].Value.Split(new char[] { ' ' });
                        border.RightSide.Size = double.Parse(strs[0]);
                        border.RightSide.Style = ReadPenStyle(strs[1]);

                        if (strs.Length == 3)
                            border.RightSide.Color = ReadColor(strs[2]);
                        else
                            border.RightSide.Color = ReadColor(strs[2] + strs[3] + strs[4]);
                    }

                    style.Border = border;

                }
            }
            catch
            {
            }
        }

        private StiPenStyle ReadPenStyle(string str)
        {
            switch (str)
            {
                case "Dash":
                    return StiPenStyle.Dash;

                case "DashDot":
                    return StiPenStyle.DashDot;

                case "DashDotDot":
                    return StiPenStyle.DashDotDot;

                case "Dot":
                    return StiPenStyle.Dot;

                case "Double":
                    return StiPenStyle.Double;

                default:
                    return StiPenStyle.Solid;
            }
        }

        private void ReadBrush(XmlNode node, StiComponent comp)
        {
            if (comp is IStiBrush)
            {
                IStiBrush brushComp = comp as IStiBrush;
                brushComp.Brush = GetBrush(node, Color.Transparent);
            }
        }

        private void ReadBrush(XmlNode node, StiStyle style)
        {
            style.Brush = GetBrush(node, Color.Transparent);
        }


        private void ReadTextBrush(XmlNode node, StiComponent comp)
        {
            if (comp is IStiTextBrush)
            {
                IStiTextBrush brushComp = comp as IStiTextBrush;
                brushComp.TextBrush = GetBrush(node, Color.Black);
            }
        }

        private void ReadTextBrush(XmlNode node, StiStyle style)
        {
            style.TextBrush = GetBrush(node, Color.Black);
        }

        private void ReadFont(XmlNode node, StiStyle style)
        {
            style.Font = GetFont(node);
        }

        private void ReadFont(XmlNode node, StiComponent comp)
        {
            IStiFont font = comp as IStiFont;
            if (font != null)
            {
                font.Font = GetFont(node);
            }
        }

        private StiBrush GetBrush(XmlNode node, Color defaultColor)
        {
            try
            {
                XmlAttribute attr = node.Attributes["type"];
                if (attr != null)
                {
                    switch (attr.Value)
                    {
                        case "PerpetuumSoft.Framework.Drawing.SolidFill":
                        case "NineRays.Basics.Drawing.SolidFill":
                            StiSolidBrush solidBrush = new StiSolidBrush();
                            solidBrush.Color = ReadColor(node, "Color", defaultColor);
                            return solidBrush;

                        case "PerpetuumSoft.Framework.Drawing.LinearGradientFill":
                        case "NineRays.Basics.Drawing.LinearGradientFill":
                            StiGradientBrush gradientBrush = new StiGradientBrush();
                            gradientBrush.StartColor = ReadColor(node, "StartColor", defaultColor);
                            gradientBrush.EndColor = ReadColor(node, "EndColor", defaultColor);
                            gradientBrush.Angle = (float)ReadDouble(node, "Angle", 0d);
                            return gradientBrush;



                        case "PerpetuumSoft.Framework.Drawing.HatchFill":
                        case "NineRays.Basics.Drawing.HatchFill":
                            StiHatchBrush hatchBrush = new StiHatchBrush();
                            hatchBrush.ForeColor = ReadColor(node, "ForeColor", defaultColor);
                            hatchBrush.BackColor = ReadColor(node, "BackColor", defaultColor);
                            string str = ReadString(node, "Style", "BackwardDiagonal");
                            hatchBrush.Style = (HatchStyle)Enum.Parse(typeof(HatchStyle), str, false);
                            return hatchBrush;

                        case "PerpetuumSoft.Framework.Drawing.EmptyFill":
                        case "PerpetuumSoft.Framework.Drawing.MultiGradientFill":
                        case "PerpetuumSoft.Framework.Drawing.SphericalFill":
                        case "PerpetuumSoft.Framework.Drawing.ConicalFill":
                        case "NineRays.Basics.Drawing.EmptyFill":
                        case "NineRays.Basics.Drawing.MultiGradientFill":
                        case "NineRays.Basics.Drawing.SphericalFill":
                        case "NineRays.Basics.Drawing.ConicalFill":
                            return new StiEmptyBrush();
                    }
                }
            }
            catch
            {
            }
            return new StiSolidBrush(Color.Black);
        }

        private Font GetFont(XmlNode node)
        {
            string fontName = "Arial";
            float fontSize = 9;
            FontStyle fontStyle = FontStyle.Regular;

            try
            {
                XmlAttribute attr = node.Attributes["FamilyName"];
                if (attr != null)
                    fontName = attr.Value;

                attr = node.Attributes["Size"];
                if (attr != null)
                    fontSize = float.Parse(attr.Value);

                attr = node.Attributes["Bold"];
                if (attr != null && attr.Value == "On")
                    fontStyle |= FontStyle.Bold;

                attr = node.Attributes["Italic"];
                if (attr != null && attr.Value == "On")
                    fontStyle |= FontStyle.Italic;

                attr = node.Attributes["Underline"];
                if (attr != null && attr.Value == "On")
                    fontStyle |= FontStyle.Underline;

                attr = node.Attributes["Strikeout"];
                if (attr != null && attr.Value == "On")
                    fontStyle |= FontStyle.Strikeout;

            }
            catch
            {
            }
            return new Font(fontName, fontSize, fontStyle);
        }
        #endregion

        #region Methods.Read.Bands
        private void ReadBand(XmlNode node, StiBand band)
        {
            ReadComp(node, band);
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

            //band.KeepReportSummaryTogether = ReadBool(node, "KeepWithData", band, false);

        }

        private void ReadChildBand(XmlNode node, StiContainer cont)
        {
            StiChildBand band = new StiChildBand();
            cont.Components.Add(band);

            ReadBand(node, band);

        }

        private void ReadDataBand(XmlNode node, StiContainer cont, StiBand masterBand)
        {
            string oldStyle = currentStyleName;

            if (importStyleNames)
            {
                string componentStyle = ReadString(node, "StyleName", "");
                if (!string.IsNullOrEmpty(componentStyle))
                    currentStyleName = componentStyle;
            }

            bool storeIsGroup = currentIsGroup;
            currentIsGroup = false;

            string storeDatabandName = currentDatabandName;
            currentDatabandName = ReadString(node, "Name", currentDatabandName);

            string storeDataSourceName = currentDataSourceName;
            currentDataSourceName = ReadString(node, "DataSource", "");

            int countData = ReadInt(node, "InstanceCount", 0);
            int columnsCount = ReadInt(node, "ColumnsCount", 0);

            dataBands[currentDatabandName] = currentDataSourceName;

            StiContainer container = new StiContainer();
            container.Name = containerName;
            ReadProperties(node, container);

            string[] parts = currentDataSourceName.Split(new char[] { '.' });
            string dsName = parts[parts.Length - 1];
            string relName = parts[parts.Length - 1];
            if (parts.Length > 1)
            {
                if (dsName.StartsWith(parts[parts.Length - 2]))
                {
                    dsName = dsName.Substring(parts[parts.Length - 2].Length);
                }
            }
            else
            {
                relName = string.Empty;
            }

            bool flag = false;
            foreach (StiComponent comp in container.Components)
            {
                StiDataBand band = comp as StiDataBand;
                if (band != null)
                {
                    flag = true;
                    band.DataSourceName = dsName;
                    band.DataRelationName = relName;
                    band.CountData = countData;
                    band.Columns = columnsCount;
                }
            }
            if (!flag)
            {
                StiDataBand band = new StiDataBand();
                band.CanGrow = true;
                band.CanShrink = true;
                band.DataSourceName = currentDataSourceName;
                band.DataRelationName = relName;
                band.CountData = countData;
                band.Columns = columnsCount;
                band.Name = ReadString(node, "Name", band.Name);

                container.Components.Add(band);
            }

            SortBands(container);
            if (cont is StiPage)
            {
                cont.Components.AddRange(container.Components);
            }
            else
            {
                cont.Components.Add(container);
            }

            currentDataSourceName = storeDataSourceName;
            currentDatabandName = storeDatabandName;
            currentIsGroup = storeIsGroup;
            currentStyleName = oldStyle;
        }

        private void ReadDetailBand(XmlNode node, StiContainer cont, StiBand masterBand)
        {
            StiDataBand band = new StiDataBand();
            cont.Components.Add(band);

            ReadBand(node, band);

            //band.MasterComponent = masterBand;

            string dataSource = ReadString(node, "DataSource", band.Name);
            band.DataSourceName = dataSource;
        }

        private void ReadHeaderBand(XmlNode node, StiContainer cont, StiBand masterBand)
        {
            StiHeaderBand band = new StiHeaderBand();
            if (masterBand != null)
                cont.Components.Insert(cont.Components.IndexOf(masterBand), band);
            else
                cont.Components.Add(band);

            ReadBand(node, band);

            //band.KeepHeaderTogether = ReadBool(node, "KeepWithData", band, false);
        }

        private void ReadFooterBand(XmlNode node, StiContainer cont, StiBand masterBand)
        {
            StiFooterBand band = new StiFooterBand();
            cont.Components.Add(band);

            ReadBand(node, band);

            //band.KeepFooterTogether = ReadBool(node, "KeepWithData", band, false);
        }

        private void ReadPageHeaderBand(XmlNode node, StiContainer cont)
        {
            StiPageHeaderBand band = new StiPageHeaderBand();
            cont.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadPageFooterBand(XmlNode node, StiContainer cont)
        {
            StiPageFooterBand band = new StiPageFooterBand();
            cont.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadGroupBand(XmlNode node, StiContainer cont)
        {
            string oldStyle = currentStyleName;

            if (importStyleNames)
            {
                string componentStyle = ReadString(node, "StyleName", "");
                if (!string.IsNullOrEmpty(componentStyle))
                    currentStyleName = componentStyle;
            }

            bool storeIsGroup = currentIsGroup;
            currentIsGroup = true;

            StiContainer container = new StiContainer();
            container.Name = containerName;
            ReadProperties(node, container);

            string groupCondition = ReadString(node, "GroupExpression", "");
            if (!string.IsNullOrEmpty(groupCondition))
            {
                groupCondition = ParseExpression(groupCondition);
                foreach (StiComponent comp in container.Components)
                {
                    StiGroupHeaderBand band = comp as StiGroupHeaderBand;
                    if (band != null)
                    {
                        (band as StiGroupHeaderBand).Condition.Value = groupCondition;
                    }
                }
            }

            SortBands(container);
            //if (cont is StiPage)
            //{
            cont.Components.AddRange(container.Components);
            //}
            //else
            //{
            //    cont.Components.Add(container);
            //}

            currentIsGroup = storeIsGroup;
            currentStyleName = oldStyle;
        }

        private void ReadGroupHeaderBand(XmlNode node, StiContainer cont, StiBand masterBand)
        {
            StiGroupHeaderBand band = new StiGroupHeaderBand();
            if (masterBand != null)
                cont.Components.Insert(cont.Components.IndexOf(masterBand), band);
            else
                cont.Components.Add(band);

            ReadBand(node, band);
        }

        private void ReadGroupFooterBand(XmlNode node, StiContainer cont, StiBand masterBand)
        {
            StiGroupFooterBand band = new StiGroupFooterBand();
            if (masterBand != null)
                cont.Components.Insert(cont.Components.IndexOf(masterBand), band);
            else
                cont.Components.Add(band);

            ReadBand(node, band);
        }

        private void SortBands(StiContainer input)
        {
            int counterHeader = 1000;
            int counterGroupHeader = 2000;
            int counterData = 3000;
            int counterGroupFooter = 4000;
            int counterFooter = 5000;
            int counterOther = 10000;

            StiDataBand masterDataBand = null;

            Dictionary<int, StiComponent> dict = new Dictionary<int, StiComponent>();
            List<int> keys = new List<int>();

            foreach (StiComponent comp in input.Components)
            {
                int counter = 0;
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
                if ((comp is StiContainer) && (comp.Name == containerName))
                {
                    StiContainer cont = comp as StiContainer;
                    foreach (StiComponent comp2 in cont.Components)
                    {
                        StiDataBand band = comp2 as StiDataBand;
                        if (band != null)
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
        private void ReadPage(XmlNode node, StiReport report)
        {
            string oldStyle = currentStyleName;

            StiPage page = new StiPage(report);
            report.Pages.Add(page);

            page.Name = ReadString(node, "Name", page.Name);
            page.Orientation = ReadString(node, "Orientation", "Portrait") == "Landscape" ? StiPageOrientation.Landscape : StiPageOrientation.Portrait;
            page.Enabled = ReadBool(node, "Visible", true);

            if (importStyleNames)
            {
                page.ComponentStyle = ReadString(node, "StyleName", "");
                if (!string.IsNullOrEmpty(page.ComponentStyle))
                    currentStyleName = page.ComponentStyle;
            }

            ReadPageMargins(node, page);
            ReadPageSize(node, page);
            ReadProperties(node, page);

            currentStyleName = oldStyle;
        }

        private void ReadDataBinding(XmlNode node, StiComponent comp)
        {
            foreach (XmlNode elementNode in node.ChildNodes)
            {
                XmlAttribute attr = elementNode.Attributes["type"];

                switch (attr.Value)
                {
                    case "PerpetuumSoft.Reporting.DOM.ReportDataBinding":
                    case "NineRays.Reporting.DOM.ReportDataBinding":
                        XmlAttribute attr2 = elementNode.Attributes["Expression"];
                        XmlAttribute attr3 = elementNode.Attributes["PropertyName"];
                        if ((attr2 != null) && (attr3 != null))
                        {
                            if ((comp is StiText) && (attr3.Value == "Value"))
                            {
                                ((StiText)comp).Text.Value = "{" + ParseExpression(attr2.Value) + "}";
                            }
                            if ((comp is StiGroupHeaderBand) && (attr3.Value == "Group"))
                            {
                                ((StiGroupHeaderBand)comp).Condition.Value = "{" + ParseExpression(attr2.Value) + "}";
                            }
                        }
                        break;
                }
            }
        }

        private void ReadProperties(XmlNode node, StiComponent comp)
        {
            StiContainer cont = comp as StiContainer;

            ////fix for versions < v2.0
            //foreach (XmlNode elementNode in node.ChildNodes)
            //{
            //    if (elementNode.Name == "Controls" && elementNode.ChildNodes.Count > 0)
            //    {
            //        node = elementNode;
            //        break;
            //    }
            //}

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                XmlAttribute attr = elementNode.Attributes["type"];

                string attrValue = null;
                if (attr != null)
                {
                    attrValue = attr.Value;
                }
                else
                {
                    if (elementNode.Name == "DataBindings") attrValue = "DataBindings";
                }

                switch (attrValue)
                {
                    case "PerpetuumSoft.Reporting.DOM.ReportDataBindingCollection":
                    case "DataBindings":
                        ReadDataBinding(elementNode, comp);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.ReportControlCollection":
                    case "NineRays.Reporting.DOM.ReportControlCollection":
                        ReadControls(elementNode, cont);
                        break;

                    case "Border":
                        ReadBorder(elementNode, comp);
                        break;

                    case "Fill":
                        ReadBrush(elementNode, comp);
                        break;

                    case "TextFill":
                        ReadTextBrush(elementNode, comp);
                        break;

                    case "PerpetuumSoft.Framework.Drawing.FontDescriptor":
                        ReadFont(elementNode, comp);
                        break;

                }
            }
        }

        private void ReadControls(XmlNode node, StiComponent comp)
        {
            StiContainer cont = comp as StiContainer;

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                XmlAttribute attr = elementNode.Attributes["type"];

                string attrValue = null;
                if (attr != null)
                {
                    attrValue = attr.Value;
                }

                switch (attrValue)
                {
                    case "PerpetuumSoft.Reporting.DOM.PageHeader":
                    case "NineRays.Reporting.DOM.PageHeader":
                        ReadPageHeaderBand(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.PageFooter":
                    case "NineRays.Reporting.DOM.PageFooter":
                        ReadPageFooterBand(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.Detail":
                    case "NineRays.Reporting.DOM.Detail":
                        ReadDetailBand(elementNode, cont, null);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.DataBand":
                    case "NineRays.Reporting.DOM.DataBand":
                        ReadDataBand(elementNode, cont, null);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.GroupBand":
                    case "NineRays.Reporting.DOM.GroupBand":
                        ReadGroupBand(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.GroupHeader":
                    case "NineRays.Reporting.DOM.GroupHeader":
                        ReadGroupHeaderBand(elementNode, cont, null);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.GroupFooter":
                    case "NineRays.Reporting.DOM.GroupFooter":
                        ReadGroupFooterBand(elementNode, cont, null);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.Header":
                    case "NineRays.Reporting.DOM.Header":
                        if (currentIsGroup)
                        {
                            ReadGroupHeaderBand(elementNode, cont, null);
                        }
                        else
                        {
                            ReadHeaderBand(elementNode, cont, null);
                        }
                        break;

                    case "PerpetuumSoft.Reporting.DOM.Footer":
                    case "NineRays.Reporting.DOM.Footer":
                        if (currentIsGroup)
                        {
                            ReadFooterBand(elementNode, cont, null);
                        }
                        else
                        {
                            ReadGroupFooterBand(elementNode, cont, null);
                        }
                        break;

                    case "PerpetuumSoft.Reporting.DOM.TextBox":
                    case "NineRays.Reporting.DOM.TextBox":
                        ReadTextBox(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.Picture":
                    case "NineRays.Reporting.DOM.Picture":
                        ReadPictureObject(elementNode, cont);
                        break;
                }
            }
        }

        private void ReadComp(XmlNode node, StiComponent comp)
        {
            string oldStyle = currentStyleName;

            comp.Name = ReadString(node, "Name", comp.Name);

            ReadComponentLocation(node, comp);
            ReadComponentSize(node, comp);

            comp.Enabled = ReadBool(node, "Enabled", true);
            comp.GrowToHeight = ReadBool(node, "GrowToBottom", false);
            comp.CanShrink = ReadBool(node, "CanShrink", false);
            comp.CanGrow = ReadBool(node, "CanGrow", false);

            if (importStyleNames)
            {
                comp.ComponentStyle = ReadString(node, "StyleName", "");

                if (!string.IsNullOrEmpty(comp.ComponentStyle))
                    currentStyleName = comp.ComponentStyle;
                else
                    comp.ComponentStyle = currentStyleName;
            }

            comp.Bookmark.Value = ReadString(node, "Bookmark", "");
            comp.Hyperlink.Value = ReadString(node, "Hyperlink", "");
            comp.ToolTip.Value = ReadString(node, "ToolTip", "");
            comp.Tag.Value = ReadString(node, "Tag", "");

            ReadProperties(node, comp);

            currentStyleName = oldStyle;
        }

        private void ReadTextBox(XmlNode node, StiContainer cont)
        {
            StiText text = new StiText();
            cont.Components.Add(text);

            ReadComp(node, text);

            text.Text.Value = ReadString(node, "Text", "");
            string expr = ReadString(node, "Text", "");

            text.Angle = -(float)ReadDouble(node, "Angle", 0d);

            #region TextAlign
            text.HorAlignment = StiTextHorAlignment.Center;
            text.VertAlignment = StiVertAlignment.Center;

            switch (ReadString(node, "TextAlign", ""))
            {
                case "TopLeft":
                    text.HorAlignment = StiTextHorAlignment.Left;
                    text.VertAlignment = StiVertAlignment.Top;
                    break;

                case "TopCenter":
                    text.HorAlignment = StiTextHorAlignment.Center;
                    text.VertAlignment = StiVertAlignment.Top;
                    break;

                case "TopRight":
                    text.HorAlignment = StiTextHorAlignment.Right;
                    text.VertAlignment = StiVertAlignment.Top;
                    break;

                case "MiddleLeft":
                    text.HorAlignment = StiTextHorAlignment.Left;
                    text.VertAlignment = StiVertAlignment.Center;
                    break;

                case "MiddleCenter":
                    text.HorAlignment = StiTextHorAlignment.Center;
                    text.VertAlignment = StiVertAlignment.Center;
                    break;

                case "MiddleRight":
                    text.HorAlignment = StiTextHorAlignment.Right;
                    text.VertAlignment = StiVertAlignment.Center;
                    break;


                case "BottomLeft":
                    text.HorAlignment = StiTextHorAlignment.Left;
                    text.VertAlignment = StiVertAlignment.Bottom;
                    break;

                case "BottomCenter":
                    text.HorAlignment = StiTextHorAlignment.Center;
                    text.VertAlignment = StiVertAlignment.Bottom;
                    break;

                case "BottomRight":
                    text.HorAlignment = StiTextHorAlignment.Right;
                    text.VertAlignment = StiVertAlignment.Bottom;
                    break;


            }
            #endregion

            ReadProperties(node, text);
        }

        private void ReadPictureObject(XmlNode node, StiContainer cont)
        {
            StiImage image = new StiImage();
            cont.Components.Add(image);

            ReadComp(node, image);
        }
        #endregion

        #region Methods.ParseExpression
        private string ParseExpression(string input)
        {
            string st = input;

            st = st.Replace("ColumnNumber", "Column");
            st = st.Replace("Document.Title", "ReportAlias");
            st = st.Replace("Document.Description", "ReportDescription");
            st = st.Replace("Now", "Today");
            st = st.Replace("PageCount", "TotalPageCount");
            st = st.Replace("LineNumber", "Line");
            //st = st.Replace("PageNumber", "PageNumber");

            int pos = -1;
            string stFind = "GetData(\"";
            while ((pos = st.IndexOf(stFind, StringComparison.InvariantCulture)) != -1)
            {
                int pos2 = st.IndexOf("\")", pos, StringComparison.InvariantCulture);
                if (pos2 != -1)
                {
                    string expr = st.Substring(pos + stFind.Length, pos2 - pos - stFind.Length);
                    fields[expr] = expr;

                    string[] parts = expr.Split(new char[] { '.' });
                    if (parts.Length > 2)
                    {
                        if (expr.StartsWith(currentDataSourceName))
                        {
                            string dsName = parts[parts.Length - 2];
                            if (dsName.StartsWith(parts[parts.Length - 3]))
                            {
                                dsName = dsName.Substring(parts[parts.Length - 3].Length);
                            }
                            expr = dsName + "." + parts[parts.Length - 1];
                        }
                    }

                    st = st.Substring(0, pos) + expr + st.Substring(pos2 + 2);
                }
            }

            stFind = "GetParameter(\"";
            while ((pos = st.IndexOf(stFind, StringComparison.InvariantCulture)) != -1)
            {
                int pos2 = st.IndexOf("\")", pos, StringComparison.InvariantCulture);
                if (pos2 != -1)
                {
                    string expr = st.Substring(pos + stFind.Length, pos2 - pos - stFind.Length);
                    st = st.Substring(0, pos) + expr + st.Substring(pos2 + 2);
                    fields[expr] = expr;
                }
            }

            foreach (DictionaryEntry de in dataBands)
            {
                if (!string.IsNullOrEmpty((string)de.Key) && !string.IsNullOrEmpty((string)de.Value))
                {
                    string st2 = (string)de.Key + "[\"";
                    while ((pos = st.IndexOf(st2, StringComparison.InvariantCulture)) != -1)
                    {
                        int pos2 = st.IndexOf("\"]", pos, StringComparison.InvariantCulture);
                        if (pos2 != -1)
                        {
                            string expr = (string)de.Value + "." + st.Substring(pos + st2.Length, pos2 - pos - st2.Length);
                            st = st.Substring(0, pos) + expr + st.Substring(pos2 + 2);
                            fields[expr] = expr;
                        }
                    }
                }
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
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US", false);

                var helper = new StiReportSharpShooterHelper();
                var report = helper.Convert(bytes, true);

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
