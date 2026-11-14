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
using Stimulsoft.Report.Import.ReportingServices;
using Stimulsoft.Report.Units;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

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
        private Dictionary<string, string> fieldsNames =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private XmlNode reportNode;
        private HashSet<StiText> generalAlignedTexts = new HashSet<StiText>();
        private HashSet<string> inlinedReportItems = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private bool usesReportCode;

        //the name of the text box being imported, Me.Value in its expressions is its own value
        private string currentTextboxName;
        #endregion

        #region Properties
        private string defaultFontFamily;
        /// <summary>
        /// The font of a text that sets no font. Reporting Services takes Arial, a report of the
        /// new designer names its own font in DefaultFontFamily, Segoe UI as a rule.
        /// </summary>
        private string DefaultFontFamily
        {
            get
            {
                return string.IsNullOrWhiteSpace(defaultFontFamily) ? "Arial" : defaultFontFamily;
            }
        }
        #endregion

        #region Process Root node
        public void ProcessRootNode(XmlNode rootNode, StiReport report, List<string> errorList)
        {
            this.report = report;
            this.errorList = errorList;
            reportNode = rootNode;
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

                    case "df:DefaultFontFamily":
                        defaultFontFamily = GetNodeTextValue(node);
                        break;

                    case "Language":
                        if (flagCounter == 0)
                            ProcessReportLanguage(node, report);
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

                    case "Author":
                        report.ReportAuthor = GetNodeTextValue(node);
                        break;

                    //the growing items of a container take its whitespace first, as the items of a container of
                    //Stimulsoft do; the whitespace is never given up when they do not grow, so nothing to set
                    case "ConsumeContainerWhitespace":
                        break;

                    case "InteractiveHeight":
                    case "InteractiveWidth":
                    case "rd:GridSpacing":
                    case "rd:SnapToGrid":
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
                    case "Code":
                    case "Page":
                    case "PageHeader":
                    case "PageFooter":
                    case "df:DefaultFontFamily":
                    case "Language":
                        //processed in first pass
                        break;

                    default:
                        ThrowError(rootNode.Name, node.Name);
                        break;
                }
            }

            ReportUnresolvedUserSorts();
            ApplyTextLanguages();
            ApplyGaugePanelLayouts();
            MergeListDetails();
            ApplyBodyLayout();
            ApplyShiftModes();
            ApplyMergeTags();

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
                        //a calculated field reads the fields of its own data source, a filtered data set
                        //keeps it in the data source of the query
                        if (column is StiCalcDataColumn cdc)
                        {
                            var st = ConvertExpression(cdc.Value, ds.Name);
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
                    //DateAdd is called as DateAndTime.DateAdd, known both to the compiler and to the parser
                    methods.AddRange(new string[] {
                        "Protected Overloads Function Format(obj As Object) As string",
		                "   Return obj.ToString()",
		                "End Function"
                    });

                    methodsPos = report.Script.IndexOf("#Region \"StiReport");

                    if (code != null)
                    {
                        methods.Add(string.Empty);
                        methods.Add("#Region \"Code\"");
                        var vbCode = ConvertCode(code);
                        methods.AddRange(Stimulsoft.Report.Export.StiExportUtils.SplitString(vbCode, true));
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

            //the data regions of the body are the details of the band, Reporting Services prints the body
            //when they have no rows
            band.PrintIfDetailEmpty = true;
            band.Page = page;
            page.Components.Add(band);
            bodyBands.Add(band);

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

            SplitBodyByPageBreaks(band, page);
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
            var name = baseNode.Attributes["Name"].Value;
            StiVariable var = new StiVariable();
            var.Name = GetParameterVariableName(report, name);
            var.Alias = name;
            var.Category = ParametersCategory;
            var.RequestFromUser = true;

            //first pass
            bool nullable = false;
            var multiValue = false;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Nullable":
                        nullable = Convert.ToBoolean(GetNodeTextValue(node));
                        break;

                    case "MultiValue":
                        multiValue = IsTrue(GetNodeTextValue(node));
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
                        var.Type = GetTypeFromDataType(dataType, nullable, multiValue);
                        break;

                    case "Prompt":
                        var.Description = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "Hidden":
                        //a hidden parameter is not shown to the user
                        var.RequestFromUser = !IsTrue(GetNodeTextValue(node));
                        break;

                    case "DefaultValue":
                    case "ValidValues":
                        //third pass
                        break;

                    case "Nullable":
                    case "MultiValue":
                        //first pass
                        break;

                    case "AllowBlank":
                        //ignored or not implemented yet
                        break;

                    //the settings of Report Builder, as the notifications of the updates of a shared parameter
                    case "cl:ComponentMetadata":
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

            if (multiValue)
                ApplyListDefaultValue(baseNode, var);

            report.Dictionary.Variables.Add(var);

            string baseField = $"Parameters!{name}.Value";
            string newField = var.Name;
            fieldsNames[baseField] = newField;
        }

        /// <summary>
        /// A parameter with many values is a list variable, its items are checked in the dialog.
        /// The default values of Reporting Services become the checked items: all the items when the
        /// default values come from the same field as the valid values, which is the rule for a list
        /// filled from a data set, and the listed items when both are given as values. Any other
        /// default is not known before the report runs, it is named in the list of errors.
        /// </summary>
        private void ApplyListDefaultValue(XmlNode parameterNode, StiVariable var)
        {
            var defaultNode = GetChildNode(parameterNode, "DefaultValue");
            if (defaultNode == null)
                return;

            var defaultReference = GetDataSetReferenceColumn(GetChildNode(defaultNode, "DataSetReference"));
            var validReference = GetDataSetReferenceColumn(
                GetChildNode(GetChildNode(parameterNode, "ValidValues"), "DataSetReference"));

            if (defaultReference != null &&
                string.Equals(defaultReference, validReference, StringComparison.OrdinalIgnoreCase))
                return;

            var valuesNode = GetChildNode(defaultNode, "Values");
            var keys = var.DialogInfo.Keys;

            if (defaultReference == null && valuesNode != null)
            {
                var values = valuesNode.ChildNodes.Cast<XmlNode>()
                    .Where(n => n.Name == "Value")
                    .Select(GetNodeTextValue)
                    .ToList();

                if (!values.Any(IsExpression) && keys != null && keys.Length > 0)
                {
                    var.DialogInfo.CheckedStates = keys.Select(k => values.Contains(k)).ToArray();
                    return;
                }

                //without the valid values the value of a list variable has no items to check, so the default
                //values become the items, all checked
                if (!values.Any(IsExpression) && validReference == null &&
                    GetChildNode(parameterNode, "ValidValues") == null)
                {
                    var.DialogInfo.Keys = values.ToArray();
                    var.DialogInfo.Values = values.ToArray();
                    var.DialogInfo.CheckedStates = values.Select(v => true).ToArray();
                    return;
                }
            }

            ThrowError(null, null, $"Default values of a parameter with many values not imported: {var.Name}");
        }

        /// <summary>
        /// Returns the column of a DataSetReference node as DataSet.Field, or null.
        /// </summary>
        private string GetDataSetReferenceColumn(XmlNode referenceNode)
        {
            var dataSetNode = GetChildNode(referenceNode, "DataSetName");
            var fieldNode = GetChildNode(referenceNode, "ValueField");

            if (dataSetNode == null || fieldNode == null)
                return null;

            return $"{GetNodeTextValue(dataSetNode)}.{GetNodeTextValue(fieldNode)}";
        }

        private static XmlNode GetChildNode(XmlNode node, string name)
        {
            return node?.ChildNodes.Cast<XmlNode>().FirstOrDefault(n => n.Name == name);
        }

        private void ProcessReportParameterDefaultValue(XmlNode baseNode, StiVariable var)
        {
            //the default values of a list are the checked items, see ApplyListDefaultValue
            if (IsListType(var.Type))
                return;

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
            var column = GetColumnReference(dataSetName, fieldName);
            var.Value = ConvertToVariableType(column, var.Type);
        }

        /// <summary>
        /// Returns the reference to the column in the code of the report. The code generator renames
        /// a name that is a keyword of the language of the report, Date turns into Date_ in Visual
        /// Basic, so the reference is corrected the same way.
        /// </summary>
        private string GetColumnReference(string dataSetName, string fieldName)
        {
            var dataSource = StiNameValidator.CorrectName(dataSetName, report);
            var column = StiNameValidator.CorrectName(fieldName, report);

            return $"{dataSource}.{column}";
        }

        /// <summary>
        /// The type of a field is free of the type of the parameter it fills, a String parameter
        /// often takes a Guid field. The value is converted, else the assignment does not compile.
        /// </summary>
        private static string ConvertToVariableType(string expression, Type type)
        {
            if (type == typeof(string))
                return $"System.Convert.ToString({expression})";

            if (type == typeof(int) || type == typeof(int?))
                return $"System.Convert.ToInt32({expression})";

            if (type == typeof(DateTime) || type == typeof(DateTime?))
                return $"System.Convert.ToDateTime({expression})";

            return expression;
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

                    //the body is printed in the columns of the page, the width of a column is the width of the body
                    case "Columns":
                        if (int.TryParse(GetNodeTextValue(node), out var columns) && columns > 1)
                            page.Columns = columns;
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
            XmlNode result = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ToggleItem":
                        //a printed report of Reporting Services shows the initial visibility, which is kept
                        ThrowError(null, null,
                            $"Toggle not supported, the initial visibility is used: {GetNodeTextValue(node)}");
                        break;

                    case "Hidden":
                        var value = GetNodeTextValue(node);
                        if (IsExpression(value))
                        {
                            var expr = ConvertPrintExpression(value, dataSet, out var usesCode);
                            if (expr != null)
                                AddPrintExpression(component, "Enabled", expr, usesCode);
                        }
                        else
                        {
                            if (IsTrue(value))
                                component.Enabled = false;
                        }
                        result = baseNode;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return result;
        }

        /// <summary>
        /// Converts an expression of a property counted when the item prints, the visibility or a color, and
        /// tells whether it calls the functions of the code of the report. An empty expression gives null, the
        /// property keeps its value then.
        /// </summary>
        private string ConvertPrintExpression(string value, string dataSet, out bool usesCode)
        {
            //ConvertIdentifier sets the flag when the expression calls Code.X, an inlined text box included
            usesReportCode = false;
            var expression = ConvertExpression(value, dataSet).Trim();
            usesCode = usesReportCode;

            expression = expression.Substring(1, expression.Length - 2).Trim();
            return expression.Length > 0 ? expression : null;
        }

        /// <summary>
        /// Sets Enabled by the expression of Hidden, or Brush and TextBrush by the expression of a color, when
        /// the component prints. The parser of the report counts the expressions of a component, but it does
        /// not know the functions of the code of the report, so an expression that calls them goes to the
        /// BeforePrint event: it is compiled with the code and runs right after the expressions.
        /// </summary>
        private void AddPrintExpression(StiComponent component, string property, string expression, bool usesCode)
        {
            var isEnabled = property == "Enabled";
            var not = ConvertSyntaxToCSharp ? "!" : "Not ";

            //Hidden is the opposite of Enabled; Nothing means a visible item, as in Reporting Services
            var value = isEnabled
                ? $"{not}System.Convert.ToBoolean({expression})"
                : $"SolidBrushValue({expression})";

            if (!usesCode || ConvertSyntaxToCSharp)
            {
                component.Expressions.Add(new StiAppExpression(property, value));
                return;
            }

            var target = isEnabled ? "StiComponent" : property == "Brush" ? "IStiBrush" : "IStiTextBrush";
            if (!isEnabled)
                value = "Stimulsoft.Report.Dictionary.StiFunctionsDrawing." + value;

            var statement = $"CType(sender, Stimulsoft.Report.Components.{target}).{property} = {value}";
            var script = component.BeforePrintEvent.Script;
            if (!string.IsNullOrEmpty(script))
                statement = script + Environment.NewLine + statement;

            component.BeforePrintEvent.Script = statement;
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

        /// <summary>
        /// Converts a size of Reporting Services to the units of the report, or to hundredths of an inch
        /// when inHInches is set: the margins of a text or an image are kept in them in any unit of the report.
        /// </summary>
        private double ToHi(string strValue, bool useRound = true, bool inHInches = false)
        {
            //an empty size is not set
            if (string.IsNullOrWhiteSpace(strValue))
                return 0;

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
            if (factor != 1 && !inHInches)
            {
                size = unitType.ConvertFromHInches(size);
            }
            if (useRound)
                return (double)Math.Round((decimal)size, 2);
            return (double)((int)((decimal)size * 100) / 100d);
        }

        private double ToPt(string strValue, bool isFont = false)
        {
            //an empty size is not set
            if (string.IsNullOrWhiteSpace(strValue))
                return 0;

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
            if (IsIgnoredNode(baseNodeName, nodeName))
                return;

            ThrowError(baseNodeName, nodeName, null);
        }

        /// <summary>
        /// Checks whether the node holds nothing a rendered report shows: the state of the designer,
        /// the authoring history, the settings of the report server and of the interactive viewer.
        /// Such nodes are skipped silently, so the list of errors keeps only what the import loses.
        /// </summary>
        private static bool IsIgnoredNode(string baseNodeName, string nodeName)
        {
            if (string.IsNullOrEmpty(nodeName))
                return false;

            //the namespaces of the report designer and of the authoring metadata, the server ignores them too
            if (nodeName.StartsWith("rd:") || nodeName.StartsWith("am:"))
                return true;

            return ignoredNodes.Contains(nodeName) || ignoredNodes.Contains($"{baseNodeName}.{nodeName}");
        }

        private static HashSet<string> ignoredNodes = new HashSet<string>
        {
            //the drawing order of the overlapped items, the banded layout has no such order
            "ZIndex",

            //the names and the output of the items in the XML and CSV renderings
            "DataElementName",
            "DataElementOutput",
            "DataElementStyle",

            //the size of the page in the interactive viewer
            "InteractiveHeight",
            "InteractiveWidth",

            //the refresh interval of the report in the browser
            "Report.AutoRefresh",

            //the grid of the parameters panel of the report server
            "Report.ReportParametersLayout",

            //the properties of the report for the custom extensions of the report server
            "Report.CustomProperties",

            //the frozen headers of the interactive viewer
            "TablixMember.FixedData"
        };

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

        private Type GetTypeFromDataType(string dataType, bool nullable, bool multiValue)
        {
            if (multiValue)
                return GetListTypeFromDataType(dataType);

            switch (dataType)
            {
                case "String":
                    return typeof(string);

                case "Integer":
                    return (nullable ? typeof(int?) : typeof(int));

                case "DateTime":
                    return (nullable ? typeof(DateTime?) : typeof(DateTime));

                case "Boolean":
                    return nullable ? typeof(bool?) : typeof(bool);

                case "Float":
                    return nullable ? typeof(double?) : typeof(double);

                default:
                    ThrowError(null, null, $"DataType \"{dataType}\" not found.");
                    return typeof(object);
            }
        }

        /// <summary>
        /// Returns the list type of Stimulsoft for a parameter of Reporting Services with many values.
        /// </summary>
        private Type GetListTypeFromDataType(string dataType)
        {
            switch (dataType)
            {
                case "String":
                    return typeof(StringList);

                case "Integer":
                    return typeof(IntList);

                case "DateTime":
                    return typeof(DateTimeList);

                case "Boolean":
                    return typeof(BoolList);

                case "Float":
                    return typeof(DoubleList);

                default:
                    ThrowError(null, null, $"DataType \"{dataType}\" not found.");
                    return typeof(StringList);
            }
        }

        private static bool IsListType(Type type)
        {
            return type != null && typeof(IStiList).IsAssignableFrom(type);
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
            return Import(bytes, null);
        }

        /// <summary>
        /// Imports a report of Reporting Services. The shared data sources and data sets the report refers to
        /// are read from the .rds and .rsd files of the given folder, the folder of the report as a rule; when
        /// it is null, they are imported without the connection and named in the list of errors.
        /// </summary>
        public static StiImportResult Import(byte[] bytes, string sharedFolder)
        {
            return Import(bytes, sharedFolder, DefaultBodyLayout);
        }

        /// <summary>
        /// Imports a report of Reporting Services and puts the items of its body on the page the given way.
        /// </summary>
        public static StiImportResult Import(byte[] bytes, string sharedFolder, StiReportingServicesBodyLayout bodyLayout)
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

                var helper = new StiReportingServicesHelper
                {
                    ConvertSyntaxToCSharp = false,
                    sharedFolder = sharedFolder,
                    BodyLayout = bodyLayout
                };
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
