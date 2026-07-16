#region Copyright (C) 2003-2026 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports  											}
{																	}
{	Copyright (C) 2003-2026 Stimulsoft     							}
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
#endregion Copyright (C) 2003-2026 Stimulsoft

using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Units;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
        #region Fields
        private List<string> errorList = null;
        private StiUnit unitType = new StiHundredthsOfInchUnit();
        private StiText lastTextComponent = null;
        private List<string> methods = new List<string>();
        internal bool ConvertSyntaxToCSharp = false;
        internal bool SetLinked = true;
        private Hashtable backgroundImages = new Hashtable();
        private StiReport report = null;
        private int flagCounter = 0;
        private Hashtable embeddedImages = new Hashtable();
        private Hashtable fieldsNames = new Hashtable();
        #endregion

        #region Process Root node
        public void ProcessRootNode(XmlNode rootNode, StiReport report, List<string> errorList)
        {
            this.report = report;
            this.errorList = errorList;
            report.Unit = unitType;
            StiPage page = report.Pages[0];
            page.PaperSize = System.Drawing.Printing.PaperKind.Letter;
            page.Margins = new StiMargins(0, 0, 0, 0);
            double reportWidth = 0;
            string code = null;

            if (flagCounter == 0)
            {
                if (ConvertSyntaxToCSharp)
                {
                    report.ScriptLanguage = StiReportLanguageType.CSharp;
                    report.Script = report.Script.Insert(0, "using System.Math\r\n");
                }
                else
                {
                    report.ScriptLanguage = StiReportLanguageType.VB;
                    report.Script = report.Script.Insert(0, "Imports System.Math\r\n");
                    report.Script = report.Script.Insert(0, "Imports Microsoft.VisualBasic\r\n");
                }
            }

            //first pass
            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "PageWidth":
                        page.PageWidth = ToHi(GetNodeTextValue(node));
                        break;
                    case "PageHeight":
                        page.PageHeight = ToHi(GetNodeTextValue(node));
                        break;

                    case "LeftMargin":
                        page.Margins.Left = ToHi(GetNodeTextValue(node), false);
                        break;
                    case "RightMargin":
                        page.Margins.Right = ToHi(GetNodeTextValue(node), false);
                        break;
                    case "TopMargin":
                        page.Margins.Top = ToHi(GetNodeTextValue(node), false);
                        break;
                    case "BottomMargin":
                        page.Margins.Bottom = ToHi(GetNodeTextValue(node), false);
                        break;

                    case "ColumnSpacing":
                        page.ColumnGaps = ToHi(GetNodeTextValue(node));
                        break;

                    case "EmbeddedImages":
                        ProcessEmbeddedImagesType(node);
                        break;

                    case "Variables":
                        ProcessVariablesType(node);
                        break;

                    case "ReportParameters":
                        ProcessReportParametersType(node, report);
                        break;

                    case "Code":
                        code = GetNodeTextValue(node);
                        break;
                }
            }

            //second pass
            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSources":
                        ProcessDataSourcesType(node, report);
                        break;

                    case "DataSets":
                        ProcessDataSetsType(node, report);
                        break;
                }
            }

            //third pass
            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "PageHeader":
                        StiPageHeaderBand pageHeader = new StiPageHeaderBand();
                        ProcessPageHeaderFooterType(node, page, pageHeader, "PageHeaderBand1");
                        break;
                    case "PageFooter":
                        StiPageFooterBand pageFooter = new StiPageFooterBand();
                        ProcessPageHeaderFooterType(node, page, pageFooter, "PageFooterBand1");
                        break;

                    case "Page":
                        ProcessPageType(node, page);
                        break;
                }
            }

            //fourth pass
            foreach (XmlNode node in rootNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "rd:ReportID":
                        report.ReportGuid = GetNodeTextValue(node);
                        break;
                    case "rd:DrawGrid":
                        report.Info.ShowGrid = Convert.ToBoolean(GetNodeTextValue(node));
                        break;

                    case "Body":
                        ProcessBody(node, page);
                        break;

                    case "Width":
                        reportWidth = ToHi(GetNodeTextValue(node));
                        break;

                    case "ReportSections":
                        ProcessReportSections(node, report, errorList);
                        break;

                    case "rd:ReportUnitType":
                        unitType = ParseReportUnitType(GetNodeTextValue(node));
                        break;

                    case "Description":
                        report.ReportDescription = GetNodeTextValue(node);
                        break;

                    case "InteractiveHeight":
                    case "InteractiveWidth":
                    case "rd:GridSpacing":
                    case "rd:SnapToGrid":
                    case "Language":    //???
                        //ignored or not implemented yet
                        break;

                    case "PageWidth":
                    case "PageHeight":
                    case "LeftMargin":
                    case "RightMargin":
                    case "TopMargin":
                    case "BottomMargin":
                    case "ColumnSpacing":
                    case "DataSources":
                    case "DataSets":
                    case "EmbeddedImages":
                    case "Variables":
                    case "ReportParameters":
                    case "Page":
                    case "PageHeader":
                    case "PageFooter":
                        //processed in first pass
                        break;

                    default:
                        ThrowError(rootNode.Name, node.Name);
                        break;
                }
            }

            var size = new SizeD();
            GetMaxSizeRecursive(ref size, page);
            while ((decimal)page.Width < (decimal)size.Width) page.SegmentPerWidth++;
            //while (page.Width + page.Margins.Left + page.Margins.Right < reportWidth) page.SegmentPerWidth++;

            report.Unit = unitType;

            if (flagCounter == 0)
            {
                #region Postprocessing
                if (SetLinked)
                {
                    foreach (StiComponent component in report.GetComponents())
                    {
                        component.Linked = true;
                    }
                }
                foreach (DictionaryEntry de in backgroundImages)
                {
                    StiImage image = de.Key as StiImage;
                    StiComponent comp = de.Value as StiComponent;
                    if (comp is StiPage)
                    {
                        page.Watermark.ImageBytes = image.ImageBytes;
                    }
                    else
                    {
                        image.ClientRectangle = comp.ClientRectangle;
                        image.Page = comp.Page;
                        comp.Parent.Components.Insert(comp.Parent.Components.IndexOf(comp), image);
                    }
                    image.Name = CheckComponentName(report, image.Name);
                }
                foreach (StiDataSource ds in report.Dictionary.DataSources)
                {
                    foreach (StiDataColumn column in ds.Columns)
                    {
                        if (column is StiCalcDataColumn cdc)
                        {
                            var st = ConvertExpression(cdc.Value, null);
                            cdc.Value = st.Substring(1, st.Length - 2);
                        }
                    }
                }
                #endregion

                #region Add methods
                int methodsPos = -1;
                if (ConvertSyntaxToCSharp)
                {
                    methods.AddRange(new string[] {
                        "private string format(DateTime dt, string format) { return dt.ToString(format); }",
                        "private string format(object obj) { return obj.ToString(); }"
                    });

                    methodsPos = report.Script.IndexOf("#region StiReport");

                    if (code != null)
                    {
                        methods.Add(string.Empty);
                        methods.Add("#region Code");
                        methods.AddRange(Stimulsoft.Report.Export.StiExportUtils.SplitString(code, true));
                        methods.Add("#endregion");
                    }
                }
                else
                {
                    methods.AddRange(new string[] {
                        "Public Function DateAdd(Interval As Microsoft.VisualBasic.DateInterval, Number As Double, DateValue As String) As DateTime",
                        "    Return DateAdd(Interval, Number, System.Convert.ToString(DateValue))",
                        "End Function",
                        "Protected Overloads Function Format(obj As Object) As string",
		                "   Return obj.ToString()",
		                "End Function"
                    });

                    methodsPos = report.Script.IndexOf("#Region \"StiReport");

                    if (code != null)
                    {
                        methods.Add(string.Empty);
                        methods.Add("#Region \"Code\"");
                        methods.AddRange(Stimulsoft.Report.Export.StiExportUtils.SplitString(ReplaceFields(code, null), true));
                        methods.Add("#End Region 'Code");
                    }
                }

                if (methodsPos > 0)
                {
                    methodsPos -= 9;

                    foreach (string st in methods)
                    {
                        string st2 = "\t\t" + st + "\r\n";
                        report.Script = report.Script.Insert(methodsPos, st2);
                        methodsPos += st2.Length;
                    }
                }
                #endregion
            }
        }

        private void GetMaxSizeRecursive(ref SizeD size, StiComponent component)
        {
            if (!(component is StiBand || component is StiPage))
            {
                var rect = component.ComponentToPage(component.ClientRectangle);
                if (rect.Right > size.Width)
                    size.Width = rect.Right;
            }
            if (component is StiContainer cont)
            {
                foreach (StiComponent comp in cont.Components)
                {
                    GetMaxSizeRecursive(ref size, comp);
                }
            }
        }

        private string CheckComponentName(StiReport report, string baseName)
        {
            string newName = StiNameCreation.CreateName(report, baseName, false, true, true);
            if (newName == baseName) return baseName;
            return newName;
        }
        #endregion

        #region ReportSections
        private void ProcessReportSections(XmlNode baseNode, StiReport report, List<string> errorList)
        {
            StiPagesCollection pages = new StiPagesCollection(report);

            flagCounter++;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ReportSection":
                        report.Pages.Clear();
                        report.Pages.Add(new StiPage(report));
                        ProcessRootNode(node, report, errorList);
                        pages.Add(report.Pages[0]);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (pages.Count > 0)
            {
                report.Pages.Clear();
                report.Pages.AddRange(pages);
            }

            flagCounter--;
        }
        #endregion

        #region Body
        private void ProcessBody(XmlNode baseNode, StiPage page)
        {
            StiDataBand band = new StiDataBand();
            band.Name = "MainBand";
            band.ClientRectangle = new RectangleD(0, 0, page.Width, page.Height);
            band.CanBreak = true;
            band.CanGrow = true;
            band.CanShrink = true;
            band.CountData = 1;
            band.Page = page;
            page.Components.Add(band);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ReportItems":
                        ProcessReportItemsType(node, band, page, null);
                        break;

                    case "Height":
                        band.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, band);
                        break;

                    //case "ColumnSpacing":
                    //ignored or not implemented yet
                    //break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            double maxHeight = 0;
            foreach (StiComponent comp in band.Components)
            {
                if (comp.Bottom > maxHeight) maxHeight = comp.Bottom;
            }
            maxHeight += 20;
            if (maxHeight < band.Height) band.Height = maxHeight;

        }
        #endregion

        #region EmbeddedImages
        private void ProcessEmbeddedImagesType(XmlNode baseNode)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "EmbeddedImage":
                        ProcessEmbeddedImageType(node);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessEmbeddedImageType(XmlNode baseNode)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ImageData":
                        string imageString = Convert.ToString(GetNodeTextValue(node));
                        imageString = imageString.Replace("\r", "").Replace("\n", "");
                        var image = StiImageConverter.StringToByteArray(imageString);
                        string imageName = baseNode.Attributes["Name"].Value;
                        embeddedImages[imageName] = image;
                        break;

                    case "MIMEType":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region ReportParameters
        private void ProcessReportParametersType(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ReportParameter":
                        ProcessReportParameterType(node, report);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessReportParameterType(XmlNode baseNode, StiReport report)
        {
            StiVariable var = new StiVariable();
            var.Name = baseNode.Attributes["Name"].Value;
            var.Alias = var.Name;
            var.Category = "ReportParameters";
            var.RequestFromUser = true;

            //first pass
            bool nullable = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Nullable":
                        nullable = Convert.ToBoolean(GetNodeTextValue(node));
                        break;
                }
            }

            //second pass
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataType":
                        string dataType = Convert.ToString(GetNodeTextValue(node));
                        var.Type = GetTypeFromDataType(dataType, nullable);
                        break;

                    case "Prompt":
                        var.Description = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "Hidden":
                        var.RequestFromUser = Convert.ToBoolean(GetNodeTextValue(node));
                        break;

                    case "DefaultValue":
                    case "ValidValues":
                        //third pass
                        break;

                    case "Nullable":
                        //first pass
                        break;

                    case "AllowBlank":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            //third pass
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DefaultValue":
                        ProcessReportParameterDefaultValue(node, var);
                        break;

                    case "ValidValues":
                        ProcessReportParameterValidValues(node, var);
                        break;


                }
            }

            report.Dictionary.Variables.Add(var);

            string baseField = $"Parameters!{var.Name}.Value";
            string newField = var.Name;
            fieldsNames[baseField] = newField;

            //Get Label is not supported yet, so used Value
            baseField = $"Parameters!{var.Name}.Label";
            fieldsNames[baseField] = newField;
        }

        private void ProcessReportParameterDefaultValue(XmlNode baseNode, StiVariable var)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Values":
                        ProcessReportParameterDefaultValueValues(node, var);
                        break;

                    case "DataSetReference":
                        ProcessReportParameterDefaultValueDataSetReference(node, var);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessReportParameterDefaultValueValues(XmlNode baseNode, StiVariable var)
        {
            var values = new List<string>();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        values.Add(Convert.ToString(GetNodeTextValue(node)));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            //in current revision only first value is used
            var.InitBy = StiVariableInitBy.Value;
            if (values.Count > 0)
            {
                string defaultValue = values[0];
                if (IsExpression(defaultValue))
                {
                    var.InitBy = StiVariableInitBy.Expression;
                    string newExpr = ConvertExpression(defaultValue, null);
                    var.Value = newExpr.Substring(1, newExpr.Length - 2);
                }
                else
                {
                    var.Value = defaultValue;
                }
            }
        }

        private void ProcessReportParameterValidValues(XmlNode baseNode, StiVariable var)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSetReference":
                        ProcessReportParameterValidValuesDataSetReference(node, var);
                        break;

                    case "ParameterValues":
                        ProcessReportParameterValidValuesParameterValues(node, var);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessReportParameterValidValuesDataSetReference(XmlNode baseNode, StiVariable var)
        {
            var.DialogInfo.ItemsInitializationType = StiItemsInitializationType.Columns;

            string dataSetName = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSetName":
                        dataSetName = Convert.ToString(GetNodeTextValue(node));
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ValueField":
                        var.DialogInfo.KeysColumn = dataSetName + "." + Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "LabelField":
                        var.DialogInfo.ValuesColumn = dataSetName + "." + Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "DataSetName":
                        //first pass
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

        }

        private void ProcessReportParameterValidValuesParameterValues(XmlNode baseNode, StiVariable var)
        {
            var values = new List<string>();
            var labels = new List<string>();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ParameterValue":
                        ProcessReportParameterValidValuesParameterValue(node, var, values, labels);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (values.Count > 0)
            {
                var.DialogInfo.Keys = values.ToArray();
                var.DialogInfo.Values = labels.ToArray();
            }
        }

        private void ProcessReportParameterValidValuesParameterValue(XmlNode baseNode, StiVariable var, List<string> values, List<string> labels)
        {
            string value = null;
            string label = null;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        value = GetNodeTextValue(node);
                        break;

                    case "Label":
                        label = GetNodeTextValue(node);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            values.Add(value);
            labels.Add(label);
        }

        private void ProcessReportParameterDefaultValueDataSetReference(XmlNode baseNode, StiVariable var)
        {
            string dataSetName = null;
            string fieldName = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSetName":
                        dataSetName = Convert.ToString(GetNodeTextValue(node));
                        break;
                }
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ValueField":
                        fieldName = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "DataSetName":
                        //first pass
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            var.InitBy = StiVariableInitBy.Expression;
            var.Value = $"{dataSetName}.{fieldName}";
        }
        #endregion

        #region Page
        private void ProcessPageType(XmlNode baseNode, StiPage page)
        {
            double pw = 0;
            double ph = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "PageWidth":
                        pw = ToHi(GetNodeTextValue(node));
                        break;
                    case "PageHeight":
                        ph = ToHi(GetNodeTextValue(node));
                        break;
                }
            }
            if (pw > ph) page.Orientation = StiPageOrientation.Landscape;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "PageWidth":
                        page.PageWidth = ToHi(GetNodeTextValue(node));
                        break;
                    case "PageHeight":
                        page.PageHeight = ToHi(GetNodeTextValue(node));
                        break;

                    case "LeftMargin":
                        page.Margins.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "RightMargin":
                        page.Margins.Right = ToHi(GetNodeTextValue(node));
                        break;
                    case "TopMargin":
                        page.Margins.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "BottomMargin":
                        page.Margins.Bottom = ToHi(GetNodeTextValue(node));
                        break;

                    case "ColumnSpacing":
                        page.ColumnGaps = ToHi(GetNodeTextValue(node));
                        break;

                    case "PageHeader":
                        StiPageHeaderBand pageHeader = new StiPageHeaderBand();
                        ProcessPageHeaderFooterType(node, page, pageHeader, "PageHeaderBand1");
                        break;
                    case "PageFooter":
                        StiPageFooterBand pageFooter = new StiPageFooterBand();
                        ProcessPageHeaderFooterType(node, page, pageFooter, "PageFooterBand1");
                        break;

                    case "Style":
                        ProcessStyleType(node, page);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessPageHeaderFooterType(XmlNode baseNode, StiPage page, StiBand band, string name)
        {
            band.Name = name;
            band.ClientRectangle = new RectangleD(0, 0, page.Width, page.Height);
            band.Page = page;
            page.Components.Add(band);

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ReportItems":
                        ProcessReportItemsType(node, band, page, null);
                        break;

                    case "Height":
                        band.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, band);
                        break;

                    case "PrintOnFirstPage":
                        if (GetNodeTextValue(node).ToLowerInvariant() == "false")
                        {
                            band.PrintOn |= StiPrintOnType.ExceptFirstPage;
                        }
                        break;

                    case "PrintOnLastPage":
                        if (GetNodeTextValue(node).ToLowerInvariant() == "false")
                        {
                            band.PrintOn |= StiPrintOnType.ExceptLastPage;
                        }
                        break;

                    //case "ColumnSpacing":
                    //ignored or not implemented yet
                    //break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region Variables
        private void ProcessVariablesType(XmlNode baseNode)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Variable":
                        var name = node.Attributes["Name"]?.Value;
                        var value = GetNodeTextValue(node.FirstChild);
                        if (!string.IsNullOrWhiteSpace(name))
                            report.Dictionary.Variables.Add("ReportVariables", name, value);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region Utils
        private static void ApplyBreakLocation(StiBand baseBand, string value)
        {
            if (!string.IsNullOrWhiteSpace(value) && baseBand is StiDynamicBand band)
            {
                var st = value.Trim();
                if (st.Equals("Start", StringComparison.OrdinalIgnoreCase))
                {
                    band.NewPageBefore = true;
                    band.SkipFirst = false;
                }
                if (st.Equals("StartAndEnd", StringComparison.OrdinalIgnoreCase))
                {
                    band.NewPageBefore = true;
                    band.SkipFirst = false;
                    band.NewPageAfter = true;
                }
                if (st.Equals("Between", StringComparison.OrdinalIgnoreCase))
                {
                    band.NewPageBefore = true;
                    band.SkipFirst = true;
                }
                if (st.Equals("End", StringComparison.OrdinalIgnoreCase))
                    band.NewPageAfter = true;
            }
        }

        private static bool IsTrue(string value)
        {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsBoldFontWeight(string fontWeight)
        {
            if (string.IsNullOrWhiteSpace(fontWeight))
                return false;
            if (string.Equals(fontWeight, "Bold", StringComparison.OrdinalIgnoreCase))
                return true;
            if (string.Equals(fontWeight, "Normal", StringComparison.OrdinalIgnoreCase))
                return false;
            try
            {
                return Convert.ToDouble(fontWeight) > 550;
            }
            catch
            {
                return false;
            }
        }

        private XmlNode ProcessVisibility(XmlNode baseNode, StiComponent component, string dataSet)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Hidden":
                        var value = GetNodeTextValue(node);
                        if (IsExpression(value))
                        {
                            var expr = ConvertExpression(value, dataSet).Trim();
                            expr = expr.Substring(1, expr.Length - 2);
                            component.Expressions.Add(new StiAppExpression("Enabled", expr));
                        }
                        else
                        {
                            if (IsTrue(value))
                                component.Enabled = false;
                        }
                        return baseNode;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return null;
        }

        private string GetNodeTextValue(XmlNode node)
        {
            if (node.FirstChild != null) return node.FirstChild.Value;
            return string.Empty;
        }

        private static string ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

        private static StiUnit ParseReportUnitType(string reportUnitName)
        {
            if (reportUnitName == "Mm") return new StiMillimetersUnit();
            if (reportUnitName == "Cm") return new StiCentimetersUnit();
            if (reportUnitName == "Inch") return new StiInchesUnit();
            return new StiHundredthsOfInchUnit();
        }

        private double ToHi(string strValue, bool useRound = true)
        {
            string stSize = string.Empty;
            double factor = 1;
            if (strValue.EndsWith("cm"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 100 / 2.54;
            }
            if (strValue.EndsWith("mm"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 100 / 25.4;
            }
            if (strValue.EndsWith("in"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 100;
            }
            if (strValue.EndsWith("pt"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 100 / 72f;
            }
            if (stSize.Length == 0)
            {
                ThrowError(null, null, "Expression in SizeType: " + strValue);
                return 0;
            }
            double size = Convert.ToDouble(stSize.Replace(",", ".").Replace(".", ApplicationDecimalSeparator)) * factor;
            if (factor != 1)
            {
                size = unitType.ConvertFromHInches(size);
            }
            if (useRound)
                return (double)Math.Round((decimal)size, 2);
            return (double)((int)((decimal)size * 100) / 100d);
        }

        private double ToPt(string strValue, bool isFont = false)
        {
            string stSize = string.Empty;
            double factor = 1;
            if (strValue.EndsWith("cm"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 72 / 2.54;
            }
            if (strValue.EndsWith("mm"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 72 / 25.4;
            }
            if (strValue.EndsWith("in"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
                factor = 72;
            }
            if (strValue.EndsWith("pt"))
            {
                stSize = strValue.Substring(0, strValue.Length - 2);
            }
            if (stSize.Length == 0)
            {
                ThrowError(null, null, "Expression in SizeType: " + strValue);
                return 0;
            }
            double size = Convert.ToDouble(stSize.Replace(",", ".").Replace(".", ApplicationDecimalSeparator)) * factor;
            return (double)Math.Round((decimal)size, 2);
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
                message = $"Node not supported: {baseNodeName}.{nodeName}";
            }
            else
            {
                message = message1;
            }
            errorList.Add(message);
        }

        private static bool IsExpression(string value)
        {
            return !string.IsNullOrWhiteSpace(value) && value.TrimStart().StartsWith("=");
        }

        public string ConvertExpression(string baseExpression, string dataset)
        {
            if (baseExpression == null)
                return string.Empty;

            string newExpression = baseExpression;
            if (baseExpression.TrimStart().StartsWith("="))
            {
                newExpression = "{" + baseExpression.Trim().Substring(1) + "}";

                newExpression = ReplaceFields(newExpression, dataset);
                
                if (newExpression == "{Now}") newExpression = "{Today}";
                if (newExpression.StartsWith("{RowNumber(\"") && newExpression.EndsWith("\")}")) newExpression = "{Line}";
                newExpression = newExpression.Replace("RowNumber(Nothing)", "Line");
                newExpression = newExpression.Replace("IIf(", "IIF(");
                newExpression = newExpression.Replace("Today()", "Today");

                if (ConvertSyntaxToCSharp)
                {
                    newExpression = StringReplaceIgnoreCase(newExpression, "vbcrlf", "\"\\r\\n\"");
                    newExpression = newExpression.Replace(" & ", " + ");
                    newExpression = newExpression.Replace(" &\r\n", " +\r\n");
                }
                else
                {
                    newExpression = newExpression.Replace(" &\r\n\r\n", " &\r\n").Replace(" & \r\n\r\n", " & \r\n").Replace(" &\r\n \r\n", " &\r\n").Replace(" & \r\n \r\n", " & \r\n");
                }

                int pos = -1;
                while((pos = newExpression.IndexOf("Code.", pos + 1, StringComparison.OrdinalIgnoreCase)) != -1)
                {
                    char ch = newExpression[pos - 1];
                    if (!(char.IsLetterOrDigit(ch) || ch == '.'))
                    {
                        newExpression = newExpression.Remove(pos, 5);
                    }
                }

                newExpression = TryParseFunctions(newExpression);
                newExpression = RemoveLineFeedFromExpression(newExpression);
            }
            return newExpression;
        }

        private string ReplaceFields(string newExpression, string dataset)
        {
            foreach (DictionaryEntry de in fieldsNames)
            {
                if (dataset == null || ((string)de.Key).StartsWith("Parameters!"))
                {
                    newExpression = StringReplaceIgnoreCase(newExpression, (string)de.Key, (string)de.Value);
                }
                else
                {
                    if (((string)de.Key).StartsWith(dataset + ":", StringComparison.OrdinalIgnoreCase))
                    {
                        string key = ((string)de.Key).Substring(dataset.Length + 1);
                        newExpression = StringReplaceIgnoreCase(newExpression, key, (string)de.Value);
                    }
                }
            }

            //check for function with second parameter "dataset"
            foreach (DictionaryEntry de in fieldsNames)
            {
                string key = (string)de.Key;
                int posKey = newExpression.IndexOf(key);
                if (posKey != -1)
                {
                    int posVal = ((string)de.Value).IndexOf(".");
                    if (posVal != -1)
                    {
                        string possibleDataset = ((string)de.Value).Substring(0, posVal);
                        int posDataset = newExpression.IndexOf("\"" + possibleDataset + "\"", posKey + key.Length);
                        if ((posDataset != -1) && (posDataset - (posKey + key.Length) < 5))
                        {
                            newExpression = StringReplaceIgnoreCase(newExpression, key, (string)de.Value);
                        }
                    }
                }
            }

            foreach (StiVariable variable in report.Dictionary.Variables)
            {
                var stVar = $"Variables!{variable.Name}.Value";
                newExpression = StringReplaceIgnoreCase(newExpression, stVar, variable.Name);
            }

            newExpression = newExpression.Replace("Globals!ExecutionTime", "Time");
            newExpression = newExpression.Replace("Globals!PageNumber", "PageNumber");
            newExpression = newExpression.Replace("Globals!TotalPages", "TotalPageCount");
            newExpression = newExpression.Replace("Globals!OverallPageNumber", "PageNumberThrough");
            newExpression = newExpression.Replace("Globals!OverallTotalPages", "TotalPageCountThrough");
            newExpression = newExpression.Replace("Globals!ReportName", "ReportName");

            return newExpression;
        }

        private string StringReplaceIgnoreCase(string input, string from, string to)
        {
            return Regex.Replace(input, from, to, RegexOptions.IgnoreCase);
        }

        private string TryParseFunctions(string input)
        {
            StringComparison sc = ConvertSyntaxToCSharp ? StringComparison.InvariantCulture : StringComparison.InvariantCultureIgnoreCase;

            foreach (string func in aggregateFunctions)
            {
                var regex = new Regex(func + @"[\s]*\(", RegexOptions.IgnoreCase);
                int startIndex = 0;
                while (startIndex < input.Length)
                {
                    var matches = regex.Matches(input, startIndex);
                    if (matches.Count == 0) break;
                    if (!matches[0].Success) break;
                    string name = matches[0].Value;
                    int index = matches[0].Index;
                    input = input.Substring(0, index) + name + input.Substring(index + name.Length);
                    int indexStart = index + name.Length;
                    int indexEnd = input.IndexOf(")", indexStart);
                    if (indexEnd == -1) break;
                    string stArgs = input.Substring(indexStart, indexEnd - indexStart).Trim();
                    string[] args = stArgs.Split(new char[] { ',' });
                    if (args.Length == 2)
                    {
                        var dsName = args[1].Trim();
                        dsName = dsName.Substring(1, dsName.Length - 2);
                        var field = args[0].Trim();
                        string oldSt = input.Substring(index, indexEnd - index);
                        string newSt = $"Totals.{func}({dsName},{field}";
                        input = input.Remove(index, oldSt.Length);
                        input = input.Insert(index, newSt);
                        indexEnd += newSt.Length - oldSt.Length;
                    }
                    if (args.Length == 1)
                    {
                        if (func.ToLowerInvariant() == "count")
                        {
                            string arg1 = args[0].Trim();
                            int posDot = arg1.LastIndexOf(".");
                            if (posDot != -1)
                            {
                                string oldSt = input.Substring(index, indexEnd - index);
                                string newSt = $"Totals.{func}({arg1.Substring(0, posDot)}";
                                input = input.Remove(index, oldSt.Length);
                                input = input.Insert(index, newSt);
                                indexEnd += newSt.Length - oldSt.Length;
                            }
                        }
                        else
                        {
                            string newSt = $"Totals.{func}(";
                            input = input.Remove(index, name.Length);
                            input = input.Insert(index, newSt);
                            indexEnd += newSt.Length - name.Length;
                        }
                    }
                    startIndex = indexEnd + 1;
                }
            }
            return input;
        }

        private static string[] aggregateFunctions = new string[] { "Avg", "Count", "Max", "Min", "Sum", "First", "Last"};

        private static string RemoveLineFeedFromExpression(string input)
        {
            StringBuilder sb = new StringBuilder();
            int expLevel = 0;
            for (int index = 0; index < input.Length; index++)
            {
                char ch = input[index];
                if (ch == '{')
                {
                    expLevel++;
                }
                if (ch == '}' && expLevel > 0)
                {
                    expLevel--;
                }
                if ((ch == '\r' || ch == '\n') && expLevel > 0) continue;
                sb.Append(ch);
            }
            return sb.ToString();
        }

        private Type GetTypeFromDataType(string dataType, bool nullable)
        {
            switch (dataType)
            {
                case "String":
                    return typeof(string);

                case "Integer":
                    return (nullable ? typeof(int?) : typeof(int));

                case "DateTime":
                    return (nullable ? typeof(DateTime?) : typeof(DateTime));

                default:
                    ThrowError(null, null, $"DataType \"{dataType}\" not found.");
                    return typeof(object);
            }
        }

        private string MakeFullDataSetName(string baseName, string name)
        {
            //By definition, nested data regions are based on the same report dataset. You cannot nest data regions that are based on different datasets.
            //Therefore, the DataSetName for data regions that are not top level must is ignored

            if (baseName == null) return name;
            if (string.IsNullOrWhiteSpace(baseName) && !string.IsNullOrWhiteSpace(name)) return name;
            return baseName;

            /* string[] parts = baseName.Split(new char[] { ':' });
            if (parts[parts.Length - 1] == name)
            {
                return baseName;
            }
            return baseName + ":" + name; */
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
                var doc = new XmlDocument();

                using (var stream = new MemoryStream(bytes))
                {
                    doc.Load(stream);
                }

                var helper = new StiReportingServicesHelper { ConvertSyntaxToCSharp = false };
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
