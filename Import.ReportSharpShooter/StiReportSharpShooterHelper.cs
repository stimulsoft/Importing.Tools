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
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Units;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components.TextFormats;
using Stimulsoft.Report.Components.ShapeTypes;
using Stimulsoft.Report.BarCodes;
using Stimulsoft.Base;
using Stimulsoft.Report.Chart;

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
        private StiPage currentPage = null;
        private bool currentIsGroup = false;
        private Hashtable dataBands = new Hashtable();
        private bool importStyleNames = false;
        private List<string> errorList = null;
        private DataSet dataSet = null;
        private StiReport report = null;
        private int version = 1;
        #endregion

        #region Methods
        public StiReport Convert(byte[] bytes, bool importStyleNames, DataSet dataSet, List<string> errorList)
        {
            this.importStyleNames = importStyleNames;
            this.dataSet = dataSet;
            this.errorList = errorList ?? new List<string>();

            report = new StiReport();
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
                version = ReadInt(documentList[0], "version", 1);
                foreach (XmlNode elementNode in documentNode.ChildNodes)
                {
                    switch (elementNode.Name)
                    {
                        case "Pages":
                            ReadPages(elementNode, report);
                            break;

                        case "StyleSheet":
                            ReadStyleSheet(elementNode, report);
                            break;

                        /*
                        case "GraphicsSettings":
                            ReadGraphicsSettings(elementNode, report);
                            break;
                        case "Parameters":
                            ReadParameters(elementNode, report);
                            break; */

                        default:
                            ThrowError(documentNode.Name, elementNode.Name);
                            break;
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
                        var st2 = st.Trim();
                        if (st2.StartsWith("Mikromarc")) continue;
                        if (!string.IsNullOrEmpty(st2))
                        {
                            string dll = (st2 + ".dll").ToLowerInvariant();
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
                                added.Add(st2);
                            }
                        }
                    }
                    if (added.Count > 0)
                    {
                        string[] newRefs = new string[report.ReferencedAssemblies.Length + added.Count];
                        Array.Copy(report.ReferencedAssemblies, newRefs, report.ReferencedAssemblies.Length);
                        for (int index = 0; index < added.Count; index++)
                        {
                            newRefs[report.ReferencedAssemblies.Length + index] = added[index] + ".dll";
                            report.Script = report.Script.Insert(0, string.Format("using {0};\r\n", added[index]));
                        }

                        report.ReferencedAssemblies = newRefs;
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

                    var parent = comp.Parent;
                    bool ignoreLeftMargin = parent is StiCrossDataBand || parent is StiCrossHeaderBand || parent is StiCrossFooterBand;

                    if (!ignoreLeftMargin)
                        comp.Left -= page.Margins.Left;

                    if (!(parent is StiPage) && (comp.Left < 0 || comp.Top < 0))
                    {
                        comp.Linked = true;
                    }
                    //comp.Top -= page.Margins.Top;
                }
            }

            report.ApplyStyles();

            if (documentList.Count > 0)
            {
                report.ReportName = ReadString(documentList[0], "Name", "Report name");
                report.ReportAlias = ReadString(documentList[0], "Title", "Report title");
                report.ReportDescription = ReadString(documentList[0], "Description", "Report description");
                report.NumberOfPass = ReadBool(documentList[0], "DoublePass", false) ? StiNumberOfPass.DoublePass : StiNumberOfPass.SinglePass;
            }

            report.ReportName = StiNameCreation.CreateName(report, report.ReportName, false, true, true);

            #region Scripts
            string commonScript = ReadString(documentList[0], "CommonScript", null);
            if (!string.IsNullOrEmpty(commonScript))
            {
                int posScript = report.Script.IndexOf("#region StiReport Designer generated code - do not modify");
                if (posScript != -1)
                {
                    posScript -= 8;
                    report.Script = report.Script.Insert(posScript, commonScript + "\r\n");
                }
            }

            string generateScript = ReadString(documentList[0], "GenerateScript", null);
            if (!string.IsNullOrEmpty(generateScript))
            {
                report.BeginRenderEvent.Script = generateScript;
            }
            #endregion

            foreach (StiPage page in report.Pages)
            {
                page.DockToContainer();
                page.Correct();
            }

            return report;
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

                    default:
                        ThrowError(dictNode.Name, elementNode.Name);
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

                    default:
                        ThrowError(dictNode.Name, elementNode.Name);
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

                    default:
                        ThrowError(dictNode.Name, elementNode.Name);
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

                    default:
                        ThrowError(node.Name, elementNode.Name);
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

                    default:
                        ThrowError(node.Name, elementNode.Name);
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

                    default:
                        ThrowError(node.Name, elementNode.Name);
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

                    default:
                        ThrowError(node.Name, elementNode.Name);
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

                    default:
                        ThrowError(node.Name, elementNode.Name);
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

                    default:
                        ThrowError(node.Name, elementNode.Name);
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
                return double.Parse(attr.Value.Replace(".", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator));
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

        private static double ParseDouble(string value)
        {
            value = value.Trim().Replace(",", ".").Replace(".", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
            if (double.TryParse(value, out double d)) return d;
            return 0;
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
                    Math.Round(ConvertFromRSSUnit(ParseDouble(strs[2])), 2),
                    Math.Round(ConvertFromRSSUnit(ParseDouble(strs[3])), 2),
                    Math.Round(ConvertFromRSSUnit(ParseDouble(strs[0])), 2),
                    Math.Round(ConvertFromRSSUnit(ParseDouble(strs[1])), 2));
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
                page.PageWidth = Math.Round(ConvertFromRSSUnit(ParseDouble(strs[0])), 2);
                page.PageHeight = Math.Round(ConvertFromRSSUnit(ParseDouble(strs[1])), 2);
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
                comp.Left = Math.Round(ConvertFromRSSUnit(ParseDouble(strs[0])), 2);
                comp.Top = Math.Round(ConvertFromRSSUnit(ParseDouble(strs[1])), 2);
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
                comp.Width = Math.Round(ConvertFromRSSUnit(ParseDouble(strs[0])), 2);
                comp.Height = Math.Round(ConvertFromRSSUnit(ParseDouble(strs[1])), 2);
            }
            catch
            {
            }
        }

        private Font ReadAttributeFont(XmlNode node)
        {
            XmlAttribute attr = node.Attributes["Font"];
            if (attr == null) return null;

            string[] strs = attr.Value.Split(new char[] { ',' });

            try
            {
                string name = strs[0].Trim();

                GraphicsUnit gu = GraphicsUnit.Point;
                string sizeSt = strs[1].Trim();
                if (sizeSt.EndsWith("pt"))
                {
                    gu = GraphicsUnit.Point;
                    sizeSt = sizeSt.Substring(0, sizeSt.Length - 2);
                }
                if (sizeSt.EndsWith("px"))
                {
                    gu = GraphicsUnit.Pixel;
                    sizeSt = sizeSt.Substring(0, sizeSt.Length - 2);
                }

                var size = (float)ParseDouble(sizeSt);

                FontStyle fs = FontStyle.Regular;
                if (strs.Length > 2)
                {
                    string styleSt = strs[2].Trim();
                    if (styleSt.StartsWith("style="))
                    {
                        if (styleSt.Contains("Bold")) fs |= FontStyle.Bold;
                        if (styleSt.Contains("Italic")) fs |= FontStyle.Italic;
                        if (styleSt.Contains("Underline")) fs |= FontStyle.Underline;
                    }
                }

                return new Font(name, size, fs, gu);
            }
            catch { }

            return null;
        }

        private Color ReadColor(XmlNode node, string key, Color defaultColor)
        {
            try
            {
                if (node.Attributes[key] == null)
                    return defaultColor;

                return ReadColor(node.Attributes[key].Value);
            }
            catch
            {
            }
            return Color.Black;
        }

        private Color ReadColor(string name)
        {
            try
            {
                if (name.Contains(","))
                {
                    var arr = name.Split(',');
                    if (arr.Length == 3)
                        return ReadColor(arr[0], arr[1], arr[2]);
                    if (arr.Length == 4)
                        return ReadColor(arr[1], arr[2], arr[3], arr[0]);
                }
                ColorConverter colorConverter = new ColorConverter();
                return (Color)colorConverter.ConvertFromString(name);
            }
            catch
            {
            }
            return Color.Black;
        }

        private Color ReadColor(string s1, string s2, string s3, string alpha = null)
        {
            try
            {
                int i1 = int.Parse(s1.Replace(",", "").Replace(".", ""));
                int i2 = int.Parse(s2.Replace(",", "").Replace(".", ""));
                int i3 = int.Parse(s3.Replace(",", "").Replace(".", ""));

                int i0 = 255; 
                if (alpha != null) i0 = int.Parse(alpha.Replace(",", "").Replace(".", ""));

                return Color.FromArgb(i0, i1, i2, i3);
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
                            border.TopSide.Size = ParseDouble(strs[0]);
                            border.TopSide.Style = ReadPenStyle(strs[1]);

                            if (strs.Length == 3)
                                border.TopSide.Color = ReadColor(strs[2]);
                            else
                                border.TopSide.Color = ReadColor(strs[2], strs[3], strs[4]);
                        }
                        if (node.Attributes["BottomLine"] != null)
                        {
                            string[] strs = node.Attributes["BottomLine"].Value.Split(new char[] { ' ' });
                            border.BottomSide.Size = ParseDouble(strs[0]);
                            border.BottomSide.Style = ReadPenStyle(strs[1]);

                            if (strs.Length == 3)
                                border.BottomSide.Color = ReadColor(strs[2]);
                            else
                                border.BottomSide.Color = ReadColor(strs[2], strs[3], strs[4]);
                        }
                        if (node.Attributes["LeftLine"] != null)
                        {
                            string[] strs = node.Attributes["LeftLine"].Value.Split(new char[] { ' ' });
                            border.LeftSide.Size = ParseDouble(strs[0]);
                            border.LeftSide.Style = ReadPenStyle(strs[1]);

                            if (strs.Length == 3)
                                border.LeftSide.Color = ReadColor(strs[2]);
                            else
                                border.LeftSide.Color = ReadColor(strs[2], strs[3], strs[4]);
                        }
                        if (node.Attributes["RightLine"] != null)
                        {
                            string[] strs = node.Attributes["RightLine"].Value.Split(new char[] { ' ' });
                            border.RightSide.Size = ParseDouble(strs[0]);
                            border.RightSide.Style = ReadPenStyle(strs[1]);

                            if (strs.Length == 3)
                                border.RightSide.Color = ReadColor(strs[2]);
                            else
                                border.RightSide.Color = ReadColor(strs[2], strs[3], strs[4]);
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
                            border.Size = ParseDouble(strs[0]);
                            border.Style = ReadPenStyle(strs[1]);
                            border.Side = StiBorderSides.All;

                            if (strs.Length == 3)
                                border.Color = ReadColor(strs[2]);
                            else
                                border.Color = ReadColor(strs[2], strs[3], strs[4]);
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
                        border.TopSide.Size = ParseDouble(strs[0]);
                        border.TopSide.Style = ReadPenStyle(strs[1]);

                        if (strs.Length == 3)
                            border.TopSide.Color = ReadColor(strs[2]);
                        else
                            border.TopSide.Color = ReadColor(strs[2], strs[3], strs[4]);
                    }
                    if (node.Attributes["BottomLine"] != null)
                    {
                        string[] strs = node.Attributes["BottomLine"].Value.Split(new char[] { ' ' });
                        border.BottomSide.Size = ParseDouble(strs[0]);
                        border.BottomSide.Style = ReadPenStyle(strs[1]);

                        if (strs.Length == 3)
                            border.BottomSide.Color = ReadColor(strs[2]);
                        else
                            border.BottomSide.Color = ReadColor(strs[2], strs[3], strs[4]);
                    }
                    if (node.Attributes["LeftLine"] != null)
                    {
                        string[] strs = node.Attributes["LeftLine"].Value.Split(new char[] { ' ' });
                        border.LeftSide.Size = ParseDouble(strs[0]);
                        border.LeftSide.Style = ReadPenStyle(strs[1]);

                        if (strs.Length == 3)
                            border.LeftSide.Color = ReadColor(strs[2]);
                        else
                            border.LeftSide.Color = ReadColor(strs[2], strs[3], strs[4]);
                    }
                    if (node.Attributes["RightLine"] != null)
                    {
                        string[] strs = node.Attributes["RightLine"].Value.Split(new char[] { ' ' });
                        border.RightSide.Size = ParseDouble(strs[0]);
                        border.RightSide.Style = ReadPenStyle(strs[1]);

                        if (strs.Length == 3)
                            border.RightSide.Color = ReadColor(strs[2]);
                        else
                            border.RightSide.Color = ReadColor(strs[2], strs[3], strs[4]);
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

        private void ReadTextFormat(XmlNode node, StiComponent comp)
        {
            IStiTextFormat format = comp as IStiTextFormat;
            if (format != null)
            {
                format.TextFormat = GetTextFormat(node);
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

                        default:
                            ThrowError(node.Name, attr.Value);
                            break;
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
            var fontSize = 9f;
            FontStyle fontStyle = FontStyle.Regular;

            try
            {
                XmlAttribute attr = node.Attributes["FamilyName"];
                if (attr != null)
                    fontName = attr.Value;

                attr = node.Attributes["Size"];
                if (attr != null)
                    fontSize = (float)ParseDouble(attr.Value);

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

        private StiFormatService GetTextFormat(XmlNode node)
        {
            string formatStyle = ReadString(node, "FormatStyle", "");
            StiFormatService format = new StiGeneralFormatService();

            switch(formatStyle)
            {
                case "Number":
                    var numberFormat = new StiNumberFormatService();
                    numberFormat.UseLocalSetting = ReadBool(node, "UseCultureSettings", true);
                    numberFormat.UseGroupSeparator = ReadBool(node, "UseGroupSeparator", false);
                    numberFormat.DecimalSeparator = ReadString(node, "DecimalSeparator", numberFormat.DecimalSeparator);
                    format = numberFormat;
                    break;

                case "Date":
                    var dateFormat = new StiDateFormatService();
                    dateFormat.StringFormat = ReadString(node, "FormatMask", dateFormat.StringFormat);
                    format = dateFormat;
                    break;

                case "Currency":
                    var currencyFormat = new StiCurrencyFormatService();
                    currencyFormat.UseLocalSetting = ReadBool(node, "UseCultureSettings", true);
                    currencyFormat.UseGroupSeparator = ReadBool(node, "UseGroupSeparator", false);
                    format = currencyFormat;
                    break;

                case "Custom":
                    var customFormat = new StiCustomFormatService();
                    customFormat.StringFormat = ReadString(node, "FormatMask", customFormat.StringFormat);
                    format = customFormat;
                    break;

                default:
                    ThrowError(node.Name, formatStyle);
                    break;
            }

            return format;
        }

        private void ReadBitmap(XmlNode node, StiComponent comp)
        {
            string st = node.ChildNodes[0].Value;
            var bytes = System.Convert.FromBase64String(st);

            var image = comp as StiImage;
            if (image != null)
            {
                image.ImageBytes = bytes;
            }
        }
        #endregion

        #region Methods.Read.Bands
        private void ReadBand(XmlNode node, StiBand band)
        {
            ReadComp(node, band, true);

            var dynamicBand = band as StiDynamicBand;
            if (dynamicBand != null)
            {
                dynamicBand.NewPageBefore = ReadBool(node, "NewPageBefore", false);
                dynamicBand.NewPageAfter = ReadBool(node, "NewPageAfter", false);
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

            string relName = string.Empty;
            StiDataSource ds = null;
            ReadDataBandDataSource(node, storeDatabandName, ref ds, ref relName);

            int countData = ReadInt(node, "InstanceCount", 0);
            int columnsCount = ReadInt(node, "ColumnsCount", 0);

            StiContainer container = new StiContainer();
            container.Name = containerName;
            ReadProperties(node, container, true);

            StiDataBand masterDataBand = new StiDataBand();
            masterDataBand.CanGrow = true;
            masterDataBand.CanShrink = true;
            masterDataBand.CountData = countData;
            masterDataBand.Columns = columnsCount;
            masterDataBand.Name = ReadString(node, "Name", masterDataBand.Name);
            if (ds != null)
                masterDataBand.DataSourceName = ds.Name;
            masterDataBand.DataRelationName = relName;

            if (ReadBool(node, "NewPageBefore", false))
            {
                masterDataBand.NewPageBefore = true;
                masterDataBand.SkipFirst = false;
            }
            if (ReadBool(node, "NewPageAfter", false))
            {
                //emulate
                masterDataBand.NewPageBefore = true;
                masterDataBand.SkipFirst = true;
            }
            if (!string.IsNullOrWhiteSpace(relName) && !string.IsNullOrWhiteSpace(storeDatabandName))
            {
                //masterComponents[band] = storeDatabandName;
            }

            container.Components.Add(masterDataBand);

            SortBands(container, masterDataBand);
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

        private void ReadDataBandDataSource(XmlNode node, string storeDatabandName, ref StiDataSource ds, ref string relName)
        {
            currentDataSourceName = ReadString(node, "DataSource", "");
            if (string.IsNullOrWhiteSpace(currentDataSourceName)) return;

            if ((dataSet != null) && !string.IsNullOrWhiteSpace(dataSet.DataSetName) && currentDataSourceName.StartsWith(dataSet.DataSetName))
            {
                currentDataSourceName = currentDataSourceName.Substring(dataSet.DataSetName.Length + 1);
            }

            string[] parts = currentDataSourceName.Split(new char[] { '.' });
            ds = Get_DataSource(parts[0]);
            if (parts.Length > 1)
            {
                StiDataSource prevDs = null;
                StiDataRelation dr = null;
                for (int index = 1; index < parts.Length; index++)
                {
                    prevDs = ds;
                    dr = Get_ChildDataRelation(parts[index], ds);
                    ds = dr.ChildSource;
                }
                relName = dr.Name;

                dataBands[currentDatabandName + "^r"] = relName;
                dataBands[currentDatabandName + "^m"] = storeDatabandName;
            }

            dataBands[currentDatabandName] = ds.Name;
        }

        private void ReadDetailBand(XmlNode node, StiContainer cont, StiBand masterBand)
        {
            StiDataBand band = new StiDataBand();
            cont.Components.Add(band);

            ReadBand(node, band);

            //band.MasterComponent = masterBand;

            //string dataSource = ReadString(node, "DataSource", band.Name);
            //band.DataSourceName = dataSource;
            band.CountData = 1;
        }

        private void ReadCrossBand(XmlNode node, StiContainer cont, StiBand masterBand)
        {
            string storeDatabandName = currentDatabandName;
            currentDatabandName = ReadString(node, "Name", currentDatabandName);
            string storeDataSourceName = currentDataSourceName;

            StiCrossDataBand band = new StiCrossDataBand();
            cont.Components.Add(band);

            ReadBand(node, band);

            string relName = string.Empty;
            StiDataSource ds = null;
            ReadDataBandDataSource(node, storeDatabandName, ref ds, ref relName);

            int countData = ReadInt(node, "InstanceCount", 0);
            band.CountData = countData;

            currentDataSourceName = storeDataSourceName;
            currentDatabandName = storeDatabandName;
        }

        private void ReadBandContainer(XmlNode node, StiContainer cont)
        {
            string oldStyle = currentStyleName;
            if (importStyleNames)
            {
                string componentStyle = ReadString(node, "StyleName", "");
                if (!string.IsNullOrEmpty(componentStyle))
                    currentStyleName = componentStyle;
            }

            StiContainer container = new StiContainer();
            container.Name = containerName;
            ReadProperties(node, container, true);

            if (container.Conditions.Count > 0)
            {
                foreach (StiComponent comp in container.Components)
                {
                    comp.Conditions.AddRange(container.Conditions);
                }
            }

            if (cont is StiPage)
            {
                cont.Components.AddRange(container.Components);
            }
            else
            {
                cont.Components.Add(container);
            }

            currentStyleName = oldStyle;
        }

        private void ReadPageOverlay(XmlNode node, StiContainer cont)
        {
            StiOverlayBand band = new StiOverlayBand();
            cont.Components.Add(band);

            ReadBand(node, band);
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
            currentIsGroup = true;      // ???

            StiContainer container = new StiContainer();
            container.Name = containerName;
            ReadProperties(node, container, true);

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

            SortBands(container, null);
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

        private List<XmlNode> SortNodeChildsVertical(XmlNode baseNode, bool needSort)
        {
            if (!needSort)
            {
                List<XmlNode> nodes2 = new List<XmlNode>();
                foreach (XmlNode node in baseNode)
                {
                    nodes2.Add(node);
                }
                return nodes2;
            }

            Dictionary<double, List<XmlNode>> dict = new Dictionary<double, List<XmlNode>>();
            List<double> keys = new List<double>();
            StiText tempText = new StiText();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                ReadComponentLocation(node, tempText);
                double y = tempText.Top;
                if (!dict.ContainsKey(y))
                {
                    dict[y] = new List<XmlNode>();
                }
                List<XmlNode> list = dict[y];
                list.Add(node);
                if (!keys.Contains(y))
                {
                    keys.Add(y);
                }
                tempText.Top += 0.01;
            }

            keys.Sort();

            List<XmlNode> nodes = new List<XmlNode>();
            foreach (double key in keys)
            {
                List<XmlNode> list = dict[key];
                nodes.AddRange(list);
            }

            return nodes;
        }

        private void SortBands(StiContainer input, StiDataBand masterDataBand)
        {
            int counterHeader = 1000;
            int counterGroupHeader = 2000;
            int counterData = 3000;
            int counterGroupFooter = 4000;
            int counterFooter = 5000;
            int counterDetail = 6000;
            int counterOther = 10000;

            //StiDataBand masterDataBand = null;

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
                    if (masterDataBand != null)
                    {
                        if (comp == masterDataBand)
                        {
                            counter = counterData;
                        }
                        else
                        {
                            counter = counterDetail++;
                            (comp as StiDataBand).MasterComponent = masterDataBand;
                        }
                    }
                    else
                    {
                        counter = counterData++;
                        masterDataBand = comp as StiDataBand;
                    }
                }
                else if (comp is StiContainer && comp.Name == containerName)
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
                if ((comp is StiContainer) && (comp.Name == containerName))
                {
                    StiContainer cont = comp as StiContainer;
                    foreach (StiComponent comp2 in cont.Components)
                    {
                        StiDataBand band = comp2 as StiDataBand;
                        if (band != null && band.MasterComponent == null)
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
            currentPage = page;

            page.Name = ReadString(node, "Name", page.Name);
            page.Orientation = ReadString(node, "Orientation", "Portrait") == "Landscape" ? StiPageOrientation.Landscape : StiPageOrientation.Portrait;
            page.Enabled = ReadBool(node, "Visible", true);

            if (importStyleNames)
            {
                page.ComponentStyle = ReadString(node, "StyleName", "");
                if (!string.IsNullOrEmpty(page.ComponentStyle))
                    currentStyleName = page.ComponentStyle;
            }

            if (version == 1)
            {
                ReadPageSize(node, page);

                foreach (XmlNode elementNode in node.ChildNodes)
                {
                    switch (elementNode.Name)
                    {
                        case "Margins":
                            page.Margins = new StiMargins(
                                Math.Round(ConvertFromRSSUnit(ReadDouble(elementNode, "Left", 0)), 2),
                                Math.Round(ConvertFromRSSUnit(ReadDouble(elementNode, "Right", 0)), 2),
                                Math.Round(ConvertFromRSSUnit(ReadDouble(elementNode, "Top", 0)), 2),
                                Math.Round(ConvertFromRSSUnit(ReadDouble(elementNode, "Bottom", 0)), 2));
                            break;

                        case "Controls":
                            ReadControls(elementNode, page, true);
                            break;

                        default:
                            ThrowError(node.Name, elementNode.Name);
                            break;
                    }
                }
            }
            else
            {
                ReadPageMargins(node, page);
                ReadPageSize(node, page);
                ReadProperties(node, page, true);
            }

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
                            if ((comp is StiRichText) && (attr3.Value == "RTFText"))
                            {
                                ((StiRichText)comp).DataColumn = ParseExpression(attr2.Value);
                            }
                            if ((comp is StiGroupHeaderBand) && (attr3.Value == "Group"))
                            {
                                ((StiGroupHeaderBand)comp).Condition.Value = "{" + ParseExpression(attr2.Value) + "}";
                            }
                            if (attr3.Value == "Visible")
                            {
                                var cond = new StiCondition();
                                cond.Enabled = false;
                                cond.Permissions = StiConditionPermissions.None;
                                cond.Item = StiFilterItem.Expression;
                                cond.Expression = ParseExpression(attr2.Value);
                                comp.Conditions.Add(cond);
                            }
                        }
                        break;

                    default:
                        ThrowError(node.Name, elementNode.Name);
                        break;
                }
            }
        }

        private void ReadProperties(XmlNode node, StiComponent comp, bool needSortControlsNodes = false)
        {
            StiContainer cont = comp as StiContainer;

            foreach (XmlNode elementNode in node.ChildNodes)
            {
                switch (elementNode.Name)
                {
                    case "DataBindings":
                        ReadDataBinding(elementNode, comp);
                        break;

                    case "Controls":
                        ReadControls(elementNode, cont, needSortControlsNodes);
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

                    case "Font":
                        ReadFont(elementNode, comp);
                        break;

                    case "TextFormat":
                        ReadTextFormat(elementNode, comp);
                        break;

                    case "ShapeStyle":
                        ReadShapeStyle(elementNode, comp);
                        break;

                    case "Image":
                        ReadBitmap(elementNode, comp);
                        break;

                    default:
                        ThrowError(node.Name, elementNode.Name);
                        break;
                }
            }
        }

        private void ReadControls(XmlNode node, StiComponent comp, bool needSortNodes = false)
        {
            StiContainer cont = comp as StiContainer;

            var nodes = SortNodeChildsVertical(node, needSortNodes);

            foreach (XmlNode elementNode in nodes)
            {
                XmlAttribute attr = elementNode.Attributes["type"];

                string attrValue = elementNode.Name;
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
                            ReadGroupFooterBand(elementNode, cont, null);
                        }
                        else
                        {
                            ReadFooterBand(elementNode, cont, null);
                        }
                        break;

                    case "PerpetuumSoft.Reporting.DOM.CrossBand":
                    case "NineRays.Reporting.DOM.CrossBand":
                        ReadCrossBand(elementNode, cont, null);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.BandContainer":
                    case "NineRays.Reporting.DOM.BandContainer":
                        ReadBandContainer(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.PageOverlay":
                    case "NineRays.Reporting.DOM.PageOverlay":
                        ReadPageOverlay(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.TextBox":
                    case "NineRays.Reporting.DOM.TextBox":
                        ReadTextBox(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.AdvancedText":
                    case "NineRays.Reporting.DOM.AdvancedText":
                        ReadTextBox(elementNode, cont, true);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.RichText":
                    case "NineRays.Reporting.DOM.RichText":
                        ReadRichText(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.Picture":
                    case "NineRays.Reporting.DOM.Picture":
                        ReadPictureObject(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.Shape":
                    case "NineRays.Reporting.DOM.Shape":
                        ReadShapeObject(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.BarCode":
                    case "NineRays.Reporting.DOM.BarCode":
                        ReadBarCode(elementNode, cont);
                        break;

                    case "PerpetuumSoft.Reporting.DOM.ChartControl":
                    case "NineRays.Reporting.DOM.ChartControl":
                        ReadChart(elementNode, cont);
                        break;

                    default:
                        ThrowError(node.Name, $"type=\"{attrValue}\"");
                        break;
                }
            }
        }

        private void ReadComp(XmlNode node, StiComponent comp, bool isBand = false)
        {
            string oldStyle = currentStyleName;

            comp.Name = ReadString(node, "Name", comp.Name);

            ReadComponentLocation(node, comp);
            ReadComponentSize(node, comp);

            comp.Enabled = ReadBool(node, "Enabled", true);
            comp.GrowToHeight = ReadBool(node, "GrowToBottom", false);
            comp.CanShrink = ReadBool(node, "CanShrink", false);
            comp.CanGrow = ReadBool(node, "CanGrow", false);
            comp.Enabled = ReadBool(node, "Visible", true);

            if (importStyleNames)
            {
                var style = ReadString(node, "StyleName", "");
                if (!string.IsNullOrWhiteSpace(style))
                    currentStyleName = style;
                if (!isBand)
                    comp.ComponentStyle = currentStyleName;
            }

            comp.Bookmark.Value = ReadString(node, "Bookmark", "");
            comp.Hyperlink.Value = ReadString(node, "Hyperlink", "");
            comp.ToolTip.Value = ReadString(node, "ToolTip", "");
            comp.Tag.Value = ReadString(node, "Tag", "");

            string generateScript = ReadString(node, "GenerateScript", null);
            if (!string.IsNullOrEmpty(generateScript))
            {
                comp.BeforePrintEvent.Script = "/*" + generateScript + "*/";
            }

            ReadProperties(node, comp);

            currentStyleName = oldStyle;
        }

        private void ReadTextBox(XmlNode node, StiContainer cont, bool advanced = false)
        {
            StiText text = new StiText();
            cont.Components.Add(text);

            text.Text.Value = ReadString(node, "Text", "");
            //string expr = ReadString(node, "Text", "");

            ReadComp(node, text);

            text.Angle = -(float)ReadDouble(node, "Angle", 0d);
            text.AllowHtmlTags = advanced;

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

            //ReadProperties(node, text);

            var font = ReadAttributeFont(node);
            if (font != null) text.Font = font;

            if (text.Text.ToString().Contains("#PAGE"))
            {
                text.Text = "#PAGE {PageNumber} #OF {TotalPageCount}"; // Åsmund
            }
        }

        private void ReadRichText(XmlNode node, StiContainer cont)
        {
            StiRichText text = new StiRichText();
            cont.Components.Add(text);

            text.Text.Value = ReadString(node, "Text", "");

            ReadComp(node, text);
        }

        private void ReadPictureObject(XmlNode node, StiContainer cont)
        {
            StiImage image = new StiImage();
            image.Stretch = true;
            cont.Components.Add(image);

            ReadComp(node, image);
        }

        private void ReadChart(XmlNode node, StiContainer cont)
        {
            StiChart chart = new StiChart();
            cont.Components.Add(chart);

            ReadComp(node, chart);
        }

        private void ReadShapeObject(XmlNode node, StiContainer cont)
        {
            StiShape shape = new StiShape();
            cont.Components.Add(shape);

            ReadComp(node, shape);

            if (shape.Width == 0) shape.Width = report.Unit.ConvertFromHInches(1d);
            if (shape.Height == 0) shape.Height = report.Unit.ConvertFromHInches(1d);

            var lineStyle = ReadString(node, "Line", null);
            if (!string.IsNullOrWhiteSpace(lineStyle))
            {
                string[] attrs = lineStyle.Split(new char[] { ' ' });
                shape.Size = (float)ParseDouble(attrs[0]);
                shape.Style = ReadPenStyle(attrs[1]);
                shape.BorderColor = ReadColor(attrs[2]);
            }

            if (shape.ShapeType is StiHorizontalLineShapeType)
            {
                var horLine = new StiHorizontalLinePrimitive(shape.ClientRectangle);
                horLine.Color = shape.BorderColor;
                horLine.Style = shape.Style;
                horLine.Size = shape.Size;
                horLine.Name = shape.Name;

                cont.Components.Remove(shape);
                cont.Components.Add(horLine);
            }

            if ((shape.ShapeType is StiVerticalLineShapeType) || (shape.ShapeType is StiRectangleShapeType) || (shape.ShapeType is StiRoundedRectangleShapeType))
            {
                StiCrossLinePrimitive primitive = null;
                if (shape.ShapeType is StiVerticalLineShapeType) primitive = new StiVerticalLinePrimitive(shape.ClientRectangle);
                if (shape.ShapeType is StiRectangleShapeType) primitive = new StiRectanglePrimitive(shape.ClientRectangle);
                if (shape.ShapeType is StiRoundedRectangleShapeType)
                {
                    primitive = new StiRoundedRectanglePrimitive(shape.ClientRectangle);
                    (primitive as StiRoundedRectanglePrimitive).Round = (shape.ShapeType as StiRoundedRectangleShapeType).Round;
                }

                primitive.Color = shape.BorderColor;
                primitive.Style = shape.Style;
                primitive.Size = shape.Size;
                primitive.Name = shape.Name;

                StiStartPointPrimitive start = new StiStartPointPrimitive();
                start.Left = primitive.Left;
                start.Top = primitive.Top;
                start.ReferenceToGuid = primitive.Guid;
                start.Parent = cont;
                cont.Components.Add(start);
                start.Linked = true;
                start.Name = "sp_" + shape.Name;

                StiEndPointPrimitive end = new StiEndPointPrimitive();
                end.Left = primitive.Right;
                end.Top = primitive.Bottom;
                end.ReferenceToGuid = primitive.Guid;
                end.Parent = cont;
                cont.Components.Add(end);
                end.Linked = true;
                end.Name = "ep_" + shape.Name;

                primitive.Top += cont.Top;
                primitive.Parent = cont.Page;
                cont.Components.Remove(shape);
                currentPage.Components.Add(primitive);
            }

        }

        private void ReadShapeStyle(XmlNode node, StiComponent comp)
        {
            var shape = comp as StiShape;
            if (shape == null) return;

            XmlAttribute attr = node.Attributes["type"];
            if (attr == null || string.IsNullOrWhiteSpace(attr.Value)) return;

            switch (attr.Value)
            {
                case "PerpetuumSoft.Framework.Drawing.LineShape":
                case "NineRays.Basics.Drawing.LineShape":
                    shape.ShapeType = new StiHorizontalLineShapeType();
                    XmlAttribute attrLineKind = node.Attributes["LineKind"];
                    if (attrLineKind.Value == "Slash") shape.ShapeType = new StiDiagonalUpLineShapeType();
                    if (attrLineKind.Value == "BackSlash") shape.ShapeType = new StiDiagonalDownLineShapeType();
                    if (shape.Height == 0) shape.ShapeType = new StiHorizontalLineShapeType();
                    if (shape.Width == 0) shape.ShapeType = new StiVerticalLineShapeType();
                    break;

                case "PerpetuumSoft.Framework.Drawing.EllipseShape":
                case "NineRays.Basics.Drawing.EllipseShape":
                    shape.ShapeType = new StiOvalShapeType();
                    break;

                case "PerpetuumSoft.Framework.Drawing.ParallelogramShape":
                case "NineRays.Basics.Drawing.ParallelogramShape":
                    shape.ShapeType = new StiParallelogramShapeType();
                    //XmlAttribute attr2 = node.Attributes["Angle"];    //not supported
                    break;

                case "PerpetuumSoft.Framework.Drawing.RectangleShape":
                case "NineRays.Basics.Drawing.RectangleShape":
                    shape.ShapeType = new StiRectangleShapeType();
                    break;

                case "PerpetuumSoft.Framework.Drawing.TriangleShape":
                case "NineRays.Basics.Drawing.TriangleShape":
                    shape.ShapeType = new StiTriangleShapeType();
                    XmlAttribute attrDirection = node.Attributes["Direction"];
                    if (attrDirection.Value == "Up") (shape.ShapeType as StiTriangleShapeType).Direction = StiShapeDirection.Up;
                    if (attrDirection.Value == "Down") (shape.ShapeType as StiTriangleShapeType).Direction = StiShapeDirection.Down;
                    if (attrDirection.Value == "Left") (shape.ShapeType as StiTriangleShapeType).Direction = StiShapeDirection.Left;
                    if (attrDirection.Value == "Right") (shape.ShapeType as StiTriangleShapeType).Direction = StiShapeDirection.Right;
                    break;

                case "PerpetuumSoft.Framework.Drawing.DiamondShape":
                case "NineRays.Basics.Drawing.DiamondShape":
                    shape.ShapeType = new StiFlowchartDecisionShapeType();
                    break;

                case "PerpetuumSoft.Framework.Drawing.RoundRectangleShape":
                case "NineRays.Basics.Drawing.RoundRectangleShape":
                    shape.ShapeType = new StiRoundedRectangleShapeType();
                    (shape.ShapeType as StiRoundedRectangleShapeType).Round = (float)ReadDouble(node, "Round", 0.2);
                    break;

                case "PerpetuumSoft.Framework.Drawing.ArrowShape":
                case "NineRays.Basics.Drawing.ArrowShape":
                    shape.ShapeType = new StiArrowShapeType();
                    XmlAttribute attrDirection2 = node.Attributes["Direction"];
                    if (attrDirection2.Value == "Up") (shape.ShapeType as StiArrowShapeType).Direction = StiShapeDirection.Up;
                    if (attrDirection2.Value == "Down") (shape.ShapeType as StiArrowShapeType).Direction = StiShapeDirection.Down;
                    if (attrDirection2.Value == "Left") (shape.ShapeType as StiArrowShapeType).Direction = StiShapeDirection.Left;
                    if (attrDirection2.Value == "Right") (shape.ShapeType as StiArrowShapeType).Direction = StiShapeDirection.Right;
                    break;

                case "PerpetuumSoft.Framework.Drawing.RectTriangleShape":
                case "NineRays.Basics.Drawing.RectTriangleShape":
                    shape.ShapeType = new StiFlowchartOffPageConnectorShapeType();
                    XmlAttribute attrDirection3 = node.Attributes["Direction"];
                    if (attrDirection3.Value == "Up") (shape.ShapeType as StiFlowchartOffPageConnectorShapeType).Direction = StiShapeDirection.Up;
                    if (attrDirection3.Value == "Down") (shape.ShapeType as StiFlowchartOffPageConnectorShapeType).Direction = StiShapeDirection.Down;
                    if (attrDirection3.Value == "Left") (shape.ShapeType as StiFlowchartOffPageConnectorShapeType).Direction = StiShapeDirection.Left;
                    if (attrDirection3.Value == "Right") (shape.ShapeType as StiFlowchartOffPageConnectorShapeType).Direction = StiShapeDirection.Right;
                    break;

                case "PerpetuumSoft.Framework.Drawing.CrossShape":
                case "NineRays.Basics.Drawing.CrossShape":
                    shape.ShapeType = new StiPlusShapeType();
                    break;

                case "PerpetuumSoft.Framework.Drawing.StarShape":
                case "NineRays.Basics.Drawing.StarShape":
                    shape.ShapeType = new StiMultiplyShapeType();   //'Star' not supported, using Multiply
                    break;

                default:
                    ThrowError("ShapeStyle", attr.Value);
                    break;
            }
        }

        private void ReadBarCode(XmlNode node, StiContainer cont)
        {
            StiBarCode barcode = new StiBarCode();
            cont.Components.Add(barcode);

            ReadComp(node, barcode);

            barcode.Code = new StiBarCodeExpression(ReadString(node, "Code", ""));
            barcode.AutoScale = true;
            barcode.ShowQuietZones = false;

            string codeType = ReadString(node, "CodeType", "Code39");
            switch (codeType)
            {
                case "Code39":
                    barcode.BarCodeType = new StiCode39BarCodeType();
                    break;
                case "Code39Extended":
                    barcode.BarCodeType = new StiCode39ExtBarCodeType();
                    break;

                case "Code93":
                    barcode.BarCodeType = new StiCode93BarCodeType();
                    break;
                case "Code93Extended":
                    barcode.BarCodeType = new StiCode93ExtBarCodeType();
                    break;

                case "Code128":
                    barcode.BarCodeType = new StiCode128AutoBarCodeType();
                    break;
                case "Code128A":
                    barcode.BarCodeType = new StiCode128aBarCodeType();
                    break;
                case "Code128B":
                    barcode.BarCodeType = new StiCode128bBarCodeType();
                    break;
                case "Code128C":
                    barcode.BarCodeType = new StiCode128cBarCodeType();
                    break;

                case "CodeEAN128A":
                    barcode.BarCodeType = new StiEAN128aBarCodeType();
                    break;
                case "CodeEAN128B":
                    barcode.BarCodeType = new StiEAN128bBarCodeType();
                    break;
                case "CodeEAN128C":
                    barcode.BarCodeType = new StiEAN128cBarCodeType();
                    break;

                case "Code_2_5_interleaved":
                    barcode.BarCodeType = new StiInterleaved2of5BarCodeType();
                    break;
                case "Code_2_5_industrial": //???
                case "Code_2_5_matrix":     //???
                    barcode.BarCodeType = new StiStandard2of5BarCodeType();
                    break;

                case "CodeMSI":
                    barcode.BarCodeType = new StiMsiBarCodeType();
                    break;

                case "CodePostNet":
                    barcode.BarCodeType = new StiPostnetBarCodeType();
                    break;

                case "CodeCodabar":
                    barcode.BarCodeType = new StiCodabarBarCodeType();
                    break;

                case "CodeEAN8":
                    barcode.BarCodeType = new StiEAN8BarCodeType();
                    break;
                case "CodeEAN13":
                    barcode.BarCodeType = new StiEAN13BarCodeType();
                    break;

                case "CodeUPC_A":
                    barcode.BarCodeType = new StiUpcABarCodeType();
                    break;
                case "CodeUPC_E0":
                case "CodeUPC_E1":
                    barcode.BarCodeType = new StiUpcEBarCodeType();
                    break;

                case "CodeUPC_Supp2":
                    barcode.BarCodeType = new StiUpcSup2BarCodeType();
                    break;
                case "CodeUPC_Supp5":
                    barcode.BarCodeType = new StiUpcSup5BarCodeType();
                    break;

                case "CodeJAN8":
                    barcode.BarCodeType = new StiJan8BarCodeType();
                    break;
                case "CodeJAN13":
                    barcode.BarCodeType = new StiJan13BarCodeType();
                    break;

                case "PDF417":
                case "PDF417Compact":
                    barcode.BarCodeType = new StiPdf417BarCodeType();
                    break;

                case "QRCode":
                    barcode.BarCodeType = new StiQRCodeBarCodeType();
                    break;

                default:
                    ThrowError("BarCodeType", codeType);
                    break;
            }
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
            st = st.Replace("Engine.IsDoublePass", "IsSecondPass");
            
            st = Check_shDbUtils(st);

            int pos = -1;
            string stFind = "(DataObjects[\"";
            string stFind2 = "\"] as DataSet).Tables[\"";
            while ((pos = st.IndexOf(stFind, StringComparison.InvariantCulture)) != -1)
            {
                int pos2 = st.IndexOf("\"]", pos, StringComparison.InvariantCulture);
                if ((pos2 != -1) && st.Substring(pos2).StartsWith(stFind2))
                {
                    int pos3 = st.IndexOf("\"]", pos2 + stFind2.Length, StringComparison.InvariantCulture);
                    if (pos3 != -1)
                    {
                        string tableName = st.Substring(pos2 + stFind2.Length, pos3 - (pos2 + stFind2.Length));

                        string stSum = st.Substring(pos3 + 2);
                        string stSumConst = ".Compute(\"Sum(";
                        if (stSum.StartsWith(stSumConst))
                        {
                            //Convert to SumD
                            int posS1 = st.IndexOf(")", pos3 + 2);
                            int posS2 = st.IndexOf(")", posS1 + 1);
                            string fieldName = st.Substring(pos3 + 2 + stSumConst.Length, posS1 - (pos3 + 2 + stSumConst.Length));

                            string expr = $"SumD({tableName}.{fieldName})";

                            st = st.Substring(0, pos) + expr + st.Substring(posS2 + 1);
                        }
                        else
                        {
                            //Convert to DataTable.Compute
                            string expr = tableName + ".DataTable";
                            st = st.Substring(0, pos) + expr + st.Substring(pos3 + 2);
                        }

                        Get_DataSource(tableName);
                    }
                }
            }

            stFind = "GetData(\"";
            while ((pos = st.IndexOf(stFind, StringComparison.InvariantCulture)) != -1)
            {
                int pos2 = st.IndexOf("\")", pos, StringComparison.InvariantCulture);
                if (pos2 != -1)
                {
                    string expr = st.Substring(pos + stFind.Length, pos2 - pos - stFind.Length);
                    if ((dataSet != null) && !string.IsNullOrWhiteSpace(dataSet.DataSetName) && expr.StartsWith(dataSet.DataSetName))
                    {
                        expr = expr.Substring(dataSet.DataSetName.Length + 1);
                    }

                    string[] parts = expr.Split(new char[] { '.' });
                    if (parts.Length == 1)
                    {
                        Get_Variable(expr);
                    }
                    else if (parts.Length > 2)
                    {
                        var ds = Get_DataSource(parts[0]);
                        for (int index = 1; index < parts.Length - 1; index++)
                        {
                            var dr = Get_ChildDataRelation(parts[index], ds);
                            ds = dr.ChildSource;
                        }
                        Get_DataColumn(parts[parts.Length - 1], ds);

                        expr = ds.Name + "." + parts[parts.Length - 1];
                    }
                    else
                    {
                        if (parts[0] == "Parameters")
                        {
                            var vr = Get_Variable(parts[1], "Parameters");
                            expr = vr.Name;
                        }
                        else
                        {
                            var ds = Get_DataSource(parts[0]);
                            var dc = Get_DataColumn(parts[1], ds);
                            expr = $"{ds.Name}.{dc.Name}";
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
                    Get_Variable(expr);
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
                            string dsName = (string)de.Value;
                            string fieldName = st.Substring(pos + st2.Length, pos2 - pos - st2.Length);
                            string expr = null;

                            if (currentDatabandName == (string)de.Key)  //is current databand
                            {
                                if (fieldName.Contains("."))    //backward relation?
                                {
                                    string[] parts = fieldName.Split(new char[] { '.' });
                                    fieldName = parts[parts.Length - 1];

                                    var ds = Get_DataSource(currentDataSourceName);
                                    for (int index = 0; index < parts.Length - 1; index++)
                                    {
                                        var dr = Get_ChildDataRelation(parts[index], ds);
                                        ds = dr.ChildSource;
                                    }
                                    dsName = ds.Name;
                                }
                                expr = dsName + "." + fieldName;
                            }
                            else
                            {
                                string path = (string)dataBands[currentDatabandName];
                                string oldName = currentDatabandName;
                                string dbName = (string)dataBands[oldName + "^m"];

                                while (!string.IsNullOrWhiteSpace(dbName))
                                {
                                    path += "." + (string)dataBands[oldName + "^r"];
                                    if (dbName == (string)de.Key) break;
                                    oldName = dbName;
                                    dbName = (string)dataBands[dbName + "^n"];
                                }
                                if (string.IsNullOrWhiteSpace(dbName))
                                {
                                    path = (string)de.Value;
                                }

                                if (fieldName.Contains("."))    //backward relation?
                                {
                                    string[] parts = fieldName.Split(new char[] { '.' });
                                    fieldName = parts[parts.Length - 1];

                                    for (int index = 0; index < parts.Length - 1; index++)
                                    {
                                        if (path.EndsWith(parts[index]))
                                        {
                                            path = path.Substring(0, path.Length - parts[index].Length - 1);
                                        }
                                        else break;
                                    }
                                }

                                expr = path + "." + fieldName;
                            }

                            //check elements
                            string[] parts2 = expr.Split(new char[] { '.' });
                            var ds2 = Get_DataSource(parts2[0]);
                            for (int index = 1; index < parts2.Length - 1; index++)
                            {
                                var dr2 = Get_ParentDataRelation(parts2[index], ds2);
                                ds2 = dr2.ParentSource;
                            }
                            Get_DataColumn(parts2[parts2.Length - 1], ds2);

                            st = st.Substring(0, pos) + expr + st.Substring(pos2 + 2);
                        }
                    }
                }
            }

            return st;
        }

        //custom function in referenced assembly
        private string Check_shDbUtils(string st)
        {
            //st = st.Replace("shDbUtils.DBToString", "ToString");
            //st = st.Replace("shDbUtils.DBToFloat", "System.Convert.ToSingle");

            st = st.Replace("shDbUtils.DBToString", "");
            st = st.Replace("shDbUtils.DBToFloat", "");

            return st;
        }


        private StiDataSource Get_DataSource(string name)
        {
            var dataSource = report.Dictionary.DataSources[name];
            if (dataSource != null) return dataSource;

            var ds = new StiDataTableSource();
            ds.Name = name;
            ds.Alias = name;
            ds.NameInSource = datasetName;
            report.Dictionary.DataSources.Add(ds);

            if (dataSet != null)
            {
                ds.NameInSource = dataSet.DataSetName + "." + name;

                DataTable dt = dataSet.Tables[name];
                if (dt != null)
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        Get_DataColumn(column.ColumnName, ds, column);
                    }
                }
            }

            return ds;
        }

        private StiDataColumn Get_DataColumn(string name, StiDataSource ds, DataColumn dc = null)
        {
            var dataColumn = ds.Columns[name];
            if (dataColumn != null) return dataColumn;

            var column = new StiDataColumn();
            column.Name = name;
            ds.Columns.Add(column);

            if (dc == null && dataSet != null)
            {
                var dt = dataSet.Tables[ds.Name];
                if (dt != null)
                {
                    dc = dt.Columns[name];
                }
            }
            if (dc != null)
            {
                column.Type = dc.DataType;
            }

            return column;
        }

        private StiDataRelation Get_ChildDataRelation(string name, StiDataSource ds)
        {
            var dataRelation = ds.GetChildRelations()[name];
            if (dataRelation != null) return dataRelation;

            if (dataSet != null)
            {
                DataTable dt = dataSet.Tables[ds.Name];
                if (dt != null)
                {
                    var dr = dt.ChildRelations[name];
                    if (dr != null)
                    {
                        dataRelation = new StiDataRelation(name, ds, Get_DataSource(dr.ChildTable.TableName), GetColumnsArray(dr.ParentColumns), GetColumnsArray(dr.ChildColumns));
                    }
                }
            }

            if (dataRelation == null)
            {
                string dsName = null;
                if (name.StartsWith(ds.Name))
                {
                    dsName = name.Substring(ds.Name.Length);
                }
                else
                {
                    dsName = "ds_" + name;
                }

                var ds2 = new StiDataTableSource();
                ds2.Name = dsName;
                ds2.Alias = dsName;
                ds2.NameInSource = datasetName;
                report.Dictionary.DataSources.Add(ds2);

                dataRelation = new StiDataRelation(name, ds, ds2, new string[0], new string[0]);
            }

            report.Dictionary.Relations.Add(dataRelation);

            return dataRelation;
        }

        private StiDataRelation Get_ParentDataRelation(string name, StiDataSource ds)
        {
            var dataRelation = ds.GetParentRelations()[name];
            if (dataRelation != null) return dataRelation;

            if (dataSet != null)
            {
                DataTable dt = dataSet.Tables[ds.Name];
                if (dt != null)
                {
                    var dr = dt.ParentRelations[name];
                    if (dr != null)
                    {
                        dataRelation = new StiDataRelation(name, Get_DataSource(dr.ParentTable.TableName), ds, GetColumnsArray(dr.ParentColumns), GetColumnsArray(dr.ChildColumns));
                    }
                }
            }

            if (dataRelation == null)
            {
                var ds2 = new StiDataTableSource();
                ds2.Name = "ds_" + name;
                ds2.Alias = "ds_" + name;
                ds2.NameInSource = datasetName;
                report.Dictionary.DataSources.Add(ds2);

                dataRelation = new StiDataRelation(name, ds2, ds, new string[0], new string[0]);
            }

            report.Dictionary.Relations.Add(dataRelation);

            return dataRelation;
        }

        private string[] GetColumnsArray(DataColumn[] dataColumns)
        {
            var output = new string[dataColumns.Length];
            for (int index = 0; index < dataColumns.Length; index++)
            {
                output[index] = dataColumns[index].ColumnName;
            }
            return output;
        }

        private StiVariable Get_Variable(string name, string category = null)
        {
            StiVariable varr = report.Dictionary.Variables[name];
            if (varr != null) return varr;

            varr = new StiVariable();
            varr.Name = name;
            varr.Alias = name;
            varr.Category = category;
            report.Dictionary.Variables.Add(varr);

            return varr;
        }
        #endregion

        #region ThrowError
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

        #region Methods.Import
        public static StiImportResult Import(byte[] bytes, bool importStyleNames = true, DataSet dataSet = null)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = StiCultureInfo.GetEN(false);

                var errorList = new List<string>();
                var helper = new StiReportSharpShooterHelper();
                var report = helper.Convert(bytes, importStyleNames, dataSet, errorList);

                return new StiImportResult(report, helper.errorList);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }
        #endregion
    }
}
