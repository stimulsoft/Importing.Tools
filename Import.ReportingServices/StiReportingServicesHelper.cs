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
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Units;
using Stimulsoft.Report.Components.TextFormats;
using System.Data;
using Stimulsoft.Base;

namespace Stimulsoft.Report.Import
{
    public class StiReportingServicesHelper
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
                }
                else
                {
                    report.ScriptLanguage = StiReportLanguageType.VB;
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

                    case "EmbeddedImages":
                        ProcessEmbeddedImagesType(node);
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
                        unitType = ParseReportUnitType(node.Value);
                        break;

                    case "Description":
                        report.ReportDescription = GetNodeTextValue(node);
                        break;


                    //-----<xsd:element name="Report">
                    //<xsd:element name="PageHeight" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="PageWidth" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="LeftMargin" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="RightMargin" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="TopMargin" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="BottomMargin" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Body" type="BodyType"/>
                    //<xsd:element name="DataSources" type="DataSourcesType" minOccurs="0"/>
                    //<xsd:element name="EmbeddedImages" type="EmbeddedImagesType" minOccurs="0"/>
                    //<xsd:element name="DataSets" type="DataSetsType" minOccurs="0"/>
                    //<xsd:element name="ReportParameters" type="ReportParametersType" minOccurs="0"/>

                    //<xsd:element name="Width" type="SizeType"/>
                    //<xsd:element name="InteractiveHeight" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="InteractiveWidth" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Language" type="xsd:string" minOccurs="0"/>

                    //<xsd:element name="Description" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Author" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="AutoRefresh" type="xsd:unsignedInt" minOccurs="0"/>
                    //<xsd:element name="Code" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="PageHeader" type="PageHeaderFooterType" minOccurs="0"/>
                    //<xsd:element name="PageFooter" type="PageHeaderFooterType" minOccurs="0"/>
                    //<xsd:element name="CodeModules" type="CodeModulesType" minOccurs="0"/>
                    //<xsd:element name="Classes" type="ClassesType" minOccurs="0"/>
                    //<xsd:element name="CustomProperties" type="CustomPropertiesType" minOccurs="0"/>
                    //<xsd:element name="DataTransform" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataSchema" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataElementName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataElementStyle" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="AttributeNormal"/>
                    //            <xsd:enumeration value="ElementNormal"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>

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
                    case "DataSources":
                    case "DataSets":
                    case "EmbeddedImages":
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

            while (page.Width < reportWidth) page.SegmentPerWidth++;

            //double height = page.Height * 0.2;
            //foreach (StiComponent comp in page.Components)
            //{
            //    if (comp is StiBand)
            //    {
            //        height += comp.Height;
            //    }
            //}
            //while (page.Height < height) page.SegmentPerHeight++;
            //if (page.SegmentPerHeight > 1)
            //{
            //    page.LargeHeight = true;
            //    page.LargeHeightFactor = page.SegmentPerHeight;
            //    page.SegmentPerHeight = 1;
            //}

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
                    image.ClientRectangle = comp.ClientRectangle;
                    image.Page = comp.Page;
                    comp.Parent.Components.Insert(comp.Parent.Components.IndexOf(comp), image);
                    image.Name = CheckComponentName(report, image.Name);
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
                        methods.AddRange(Stimulsoft.Report.Export.StiExportUtils.SplitString(code, true));
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

                    //-----<xsd:element name="Body" type="BodyType"/>
                    //<xsd:element name="ReportItems" type="ReportItemsType" minOccurs="0"/>
                    //<xsd:element name="Height" type="SizeType"/>

                    //<xsd:element name="Columns" type="xsd:unsignedInt" minOccurs="0"/>
                    //<xsd:element name="ColumnSpacing" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Style" type="StyleType" minOccurs="0"/>

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

        #region Report items
        private void ProcessReportItemsType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            int colspan = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ColSpan":
                        colspan = Convert.ToInt32(GetNodeTextValue(node));
                        break;
                }
            }
            if (colspan > 0)
            {
                container.TagValue = colspan;
            }
            else
            {
                container.TagValue = null;
            }

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Textbox":
                        ProcessTextboxType(node, container, page, dataset);
                        break;
                    case "Rectangle":
                        ProcessRectangleType(node, container, page, dataset);
                        break;
                    case "Table":
                        ProcessTableType(node, container, page, dataset);
                        break;
                    case "Tablix":
                        ProcessTablixType(node, container, page, dataset);
                        break;
                    case "Image":
                        ProcessImageType(node, container, page, dataset);
                        break;
                    case "List":
                        ProcessListType(node, container, page, dataset);
                        break;
                    case "Line":
                        ProcessLineType(node, container, page, dataset);
                        break;
                    case "Subreport":
                        ProcessSubreportType(node, container, page, dataset);
                        break;

                    //case "PageHeight":
                    //    page.PageHeight = ToHi(GetNodeTextValue(node));
                    //    break;

                    //-----<xsd:complexType name="ReportItemsType">
                    //<xsd:element name="Textbox" type="TextboxType"/>
                    //<xsd:element name="Rectangle" type="RectangleType"/>
                    //<xsd:element name="Table" type="TableType"/>
                    //<xsd:element name="Image" type="ImageType"/>
                    //<xsd:element name="List" type="ListType"/>
                    //<xsd:element name="Line" type="LineType"/>

                    //<xsd:element name="Subreport" type="SubreportType"/>
                    //<xsd:element name="Matrix" type="MatrixType"/>
                    //<xsd:element name="Chart" type="ChartType"/>
                    //<xsd:element name="CustomReportItem" type="CustomReportItemType"/>

                    //case "InteractiveHeight":
                    //case "InteractiveWidth":
                    //ignored or not implemented yet
                    //break;

                    case "ColSpan":
                        //already processed
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }

                if ((container.TagValue != null) && (container.Components.Count > 0))
                {
                    container.Components[container.Components.Count - 1].TagValue = container.TagValue;
                }
            }

            container.TagValue = null;
        }

        #region Image
        private StiImage ProcessImageType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiImage component = new StiImage();
            string name = "Image";
            if (baseNode.Attributes["Name"] != null)
            {
                name = baseNode.Attributes["Name"].Value;
            }
            component.Name = CheckComponentName(report, name);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            string source = null;
            string value = null;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
                        component.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
                        component.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
                        component.Width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        component.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, component);
                        break;

                    case "Sizing":
                        string sizing = Convert.ToString(GetNodeTextValue(node));
                        if (sizing == "Fit")
                        {
                            component.Stretch = true;
                        }
                        else if (sizing == "FitProportional")
                        {
                            component.Stretch = true;
                            component.AspectRatio = true;
                        }
                        else if (sizing == "AutoSize")
                        {
                            component.CanGrow = true;
                            component.CanShrink = true;
                        }
                        break;

                    case "Source":
                        source = Convert.ToString(GetNodeTextValue(node));
                        break;
                    case "Value":
                        value = Convert.ToString(GetNodeTextValue(node));
                        break;


                    //-----<xsd:complexType name="ImageType">
                    //<xsd:element name="Left" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Top" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Width" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Height" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Style" type="StyleType" minOccurs="0"/>
                    //<xsd:element name="Sizing" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="AutoSize"/>
                    //            <xsd:enumeration value="Fit"/>
                    //            <xsd:enumeration value="FitProportional"/>
                    //            <xsd:enumeration value="Clip"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>
                    //<xsd:element name="Source">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="External"/>
                    //            <xsd:enumeration value="Embedded"/>
                    //            <xsd:enumeration value="Database"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>
                    //<xsd:element name="Value" type="xsd:string"/>

                    //<xsd:element name="MIMEType" type="xsd:string" minOccurs="0"/>

                    //<xsd:element name="Action" type="ActionType" minOccurs="0"/>
                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>
                    //<xsd:element name="ToolTip" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Label" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="LinkToChild" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Bookmark" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="RepeatWith" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="CustomProperties" type="CustomPropertiesType" minOccurs="0"/>
                    //<xsd:element name="DataElementName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataElementOutput" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Output"/>
                    //            <xsd:enumeration value="NoOutput"/>
                    //            <xsd:enumeration value="ContentsOnly"/>
                    //            <xsd:enumeration value="Auto"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>

                    case "MIMEType":
                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            if (source != null)
            {
                if (source == "Embedded")
                {
                    var image = (byte[])embeddedImages[value];
                    if (image != null) component.ImageBytes = image;
                }
                else if (source == "External")
                {
                    component.ImageURL = new StiImageURLExpression(ConvertExpression(value, dataset));
                }
                else if (source == "Database")
                {
                    string datacolumn = ConvertExpression(value, dataset);
                    if (datacolumn.StartsWith("{")) datacolumn = datacolumn.Substring(1, datacolumn.Length - 2);
                    component.DataColumn = datacolumn;
                }
            }
            if (container != null)
            {
                container.Components.Add(component);
            }

            return component;
        }
        #endregion

        #region Textbox
        private void ProcessTextboxType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiText component = new StiText();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Font = new Font("Arial", 10);
            component.Page = page;

            lastTextComponent = component;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
                        component.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
                        component.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
                        component.Width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        component.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "CanGrow":
                        component.CanGrow = Convert.ToBoolean(GetNodeTextValue(node));
                        break;
                    case "CanShrink":
                        component.CanShrink = Convert.ToBoolean(GetNodeTextValue(node));
                        break;

                    case "Value":
                        component.Text = GetNodeTextValue(node);
                        break;

                    case "Paragraphs":
                        component.Text = ProcessParagraphsType(node, dataset);
                        break;

                    case "Style":
                        ProcessStyleType(node, component);
                        break;


                    case "KeepTogether":
                        string keepTogether = GetNodeTextValue(node).ToLowerInvariant();
                        if (keepTogether == "true") component.CanBreak = false;
                        if (keepTogether == "false") component.CanBreak = true;
                        break;


                    //-----<xsd:complexType name="TextboxType">
                    //-----<xsd:attribute name="Name" type="xsd:normalizedString" use="required"/>
                    //<xsd:element name="Left" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Top" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Width" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Height" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="CanGrow" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="Value" type="xsd:string"/>
                    //<xsd:element name="Style" type="StyleType" minOccurs="0"/>
                    //<xsd:element name="CanShrink" type="xsd:boolean" minOccurs="0"/>

                    //<xsd:element name="ZIndex" type="xsd:unsignedInt" minOccurs="0"/>

                    //<xsd:element name="Action" type="ActionType" minOccurs="0"/>
                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>
                    //<xsd:element name="ToolTip" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Label" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="LinkToChild" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Bookmark" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="RepeatWith" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="CustomProperties" type="CustomPropertiesType" minOccurs="0"/>
                    //<xsd:element name="HideDuplicates" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="ToggleImage" type="ToggleImageType" minOccurs="0"/>
                    //<xsd:element name="UserSort" type="UserSortType" minOccurs="0"/>
                    //<xsd:element name="DataElementName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataElementOutput" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Output"/>
                    //            <xsd:enumeration value="NoOutput"/>
                    //            <xsd:enumeration value="ContentsOnly"/>
                    //            <xsd:enumeration value="Auto"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>
                    //<xsd:element name="DataElementStyle" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Auto"/>
                    //            <xsd:enumeration value="AttributeNormal"/>
                    //            <xsd:enumeration value="ElementNormal"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>		

                    //case "InteractiveHeight":
                    //case "InteractiveWidth":
                    case "ZIndex":
                    case "rd:DefaultName":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            component.Text = ConvertExpression(component.Text, dataset);
            if (container.TagValue != null)
            {
                component.TagValue = container.TagValue;
            }

            container.Components.Add(component);

            lastTextComponent = null;
        }

        private string ProcessParagraphsType(XmlNode baseNode, string dataset)
        {
            StringBuilder text = new StringBuilder();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Paragraph":
                        if (text.Length > 0) text.Append("\r\n");
                        ProcessParagraphType(node, text, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return text.ToString();
        }

        private void ProcessParagraphType(XmlNode baseNode, StringBuilder text, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TextRuns":
                        ProcessTextRunsType(node, text, dataset);
                        break;

                    case "Style":
                        if (lastTextComponent != null)
                        {
                            ProcessStyleType(node, lastTextComponent);
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessTextRunsType(XmlNode baseNode, StringBuilder text, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TextRun":
                        ProcessTextRunType(node, text, dataset);
                        break;

                    //case "Style":
                    //    ProcessStyleType(node, component);
                    //    break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessTextRunType(XmlNode baseNode, StringBuilder text, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        string expr = GetNodeTextValue(node);
                        text.Append(ConvertExpression(expr, dataset));
                        break;

                    case "Style":
                        if (lastTextComponent != null)
                        {
                            ProcessStyleType(node, lastTextComponent);
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region Rectangle
        private void ProcessRectangleType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiPanel component = new StiPanel();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
                        component.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
                        component.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
                        component.Width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        component.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, component);
                        break;

                    case "ReportItems":
                        ProcessReportItemsType(node, component, page, dataset);
                        break;


                    case "KeepTogether":
                        string keepTogether = GetNodeTextValue(node).ToLowerInvariant();
                        if (keepTogether == "true") component.CanBreak = false;
                        if (keepTogether == "false") component.CanBreak = true;
                        break;


                    //-----<xsd:complexType name="RectangleType">
                    //-----<xsd:attribute name="Name" type="xsd:normalizedString" use="required"/>
                    //<xsd:element name="Left" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Top" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Width" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Height" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Style" type="StyleType" minOccurs="0"/>
                    //<xsd:element name="ReportItems" type="ReportItemsType" minOccurs="0"/>

                    //<xsd:element name="Action" type="ActionType" minOccurs="0"/>
                    //<xsd:element name="ZIndex" type="xsd:unsignedInt" minOccurs="0"/>
                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>
                    //<xsd:element name="ToolTip" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Label" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="LinkToChild" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Bookmark" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="RepeatWith" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="CustomProperties" type="CustomPropertiesType" minOccurs="0"/>
                    //<xsd:element name="PageBreakAtStart" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="PageBreakAtEnd" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="DataElementName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataElementOutput" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Output"/>
                    //            <xsd:enumeration value="NoOutput"/>
                    //            <xsd:enumeration value="ContentsOnly"/>
                    //            <xsd:enumeration value="Auto"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>


                    //ignored or not implemented yet
                    //break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            container.Components.Add(component);
        }
        #endregion

        #region Line
        private void ProcessLineType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiHorizontalLinePrimitive component = new StiHorizontalLinePrimitive();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Page = page;
            StiText text = new StiText();

            //double left = 0;
            //double top = 0;
            double width = 0;
            double height = 0;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
                        component.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
                        component.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
                        width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, text);
                        break;


                    //-----<xsd:complexType name="LineType">
                    //<xsd:element name="Top" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Left" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Height" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Width" type="SizeType" minOccurs="0"/>

                    //<xsd:element name="ZIndex" type="xsd:unsignedInt" minOccurs="0"/>

                    //<xsd:element name="Style" type="StyleType" minOccurs="0"/>
                    //<xsd:element name="Action" type="ActionType" minOccurs="0"/>
                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>
                    //<xsd:element name="ToolTip" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Label" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="LinkToChild" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Bookmark" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="RepeatWith" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="CustomProperties" type="CustomPropertiesType" minOccurs="0"/>
                    //<xsd:element name="DataElementName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataElementOutput" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Output"/>
                    //            <xsd:enumeration value="NoOutput"/>
                    //            <xsd:enumeration value="ContentsOnly"/>
                    //            <xsd:enumeration value="Auto"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>

                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (height < 0)
            {
                component.Top += height;
                height = Math.Abs(height);
            }
            if (width < 0)
            {
                component.Left += width;
                width = Math.Abs(width);
            }

            if (width > height)
            {
                component.Width = width;
                component.Height = height;
                component.Style = text.Border.Style;
                component.Color = text.Border.Color;
                component.Size = (float)text.Border.Size;
                container.Components.Add(component);
            }
            else
            {
                StiVerticalLinePrimitive line = new StiVerticalLinePrimitive();
                line.Name = component.Name;
                line.Page = component.Page;
                line.Left = component.Left;
                line.Top = component.Top;
                line.Width = width;
                line.Height = height;
                line.Style = text.Border.Style;
                line.Color = text.Border.Color;
                line.Size = (float)text.Border.Size;
                page.Components.Add(line);

                StiStartPointPrimitive start = new StiStartPointPrimitive();
                start.Left = line.Left;
                start.Top = line.Top;
                start.ReferenceToGuid = line.Guid;
                start.Page = page;
                container.Components.Add(start);

                StiEndPointPrimitive end = new StiEndPointPrimitive();
                end.Left = line.Right;
                end.Top = line.Bottom;
                end.ReferenceToGuid = line.Guid;
                end.Page = page;
                container.Components.Add(end);
            }
        }
        #endregion

        #region Table
        private void ProcessTableType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiPanel component = new StiPanel();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;

            string datasetName = null;
            string fullDatasetName = null;

            //first pass - get dataset name
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSetName":
                        datasetName = Convert.ToString(GetNodeTextValue(node));
                        break;
                }
            }
            fullDatasetName = MakeFullDataSetName(dataset, datasetName);

            ArrayList tableColumns = new ArrayList();
            ArrayList tableHeaders = new ArrayList();
            ArrayList tableDetails = new ArrayList();
            ArrayList tableFooters = new ArrayList();

            //second pass
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
                        component.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
                        component.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
                        component.Width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        component.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, component);
                        break;

                    case "TableColumns":
                        tableColumns = ProcessTableColumnsType(node);
                        break;

                    case "Header":
                        tableHeaders = ProcessHeaderType(node, component, page, fullDatasetName);
                        break;
                    case "Details":
                        tableDetails = ProcessDetailsType(node, component, page, fullDatasetName);
                        break;
                    case "Footer":
                        tableFooters = ProcessFooterType(node, component, page, fullDatasetName);
                        break;


                    //-----<xsd:complexType name="TableType">
                    //-----<xsd:attribute name="Name" type="xsd:normalizedString" use="required"/>
                    //<xsd:element name="Left" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Top" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Width" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Height" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Style" type="StyleType" minOccurs="0"/>
                    //<xsd:element name="TableColumns" type="TableColumnsType"/>
                    //<xsd:element name="Header" type="HeaderType" minOccurs="0"/>
                    //<xsd:element name="Details" type="DetailsType" minOccurs="0"/>
                    //<xsd:element name="Footer" type="FooterType" minOccurs="0"/>
                    //<xsd:element name="DataSetName" type="xsd:string" minOccurs="0"/>

                    //<xsd:element name="ZIndex" type="xsd:unsignedInt" minOccurs="0"/>

                    //<xsd:element name="Action" type="ActionType" minOccurs="0"/>
                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>
                    //<xsd:element name="ToolTip" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Label" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="LinkToChild" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Bookmark" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="RepeatWith" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="CustomProperties" type="CustomPropertiesType" minOccurs="0"/>
                    //<xsd:element name="KeepTogether" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="NoRows" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="PageBreakAtStart" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="PageBreakAtEnd" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="Filters" type="FiltersType" minOccurs="0"/>
                    //<xsd:element name="TableGroups" type="TableGroupsType" minOccurs="0"/>
                    //<xsd:element name="FillPage" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="DataElementName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataElementOutput" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Output"/>
                    //            <xsd:enumeration value="NoOutput"/>
                    //            <xsd:enumeration value="ContentsOnly"/>
                    //            <xsd:enumeration value="Auto"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>
                    //<xsd:element name="DetailDataElementName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DetailDataCollectionName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DetailDataElementOutput" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Output"/>
                    //            <xsd:enumeration value="NoOutput"/>
                    //            <xsd:enumeration value="ContentsOnly"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>
                    //<xsd:any namespace="##other" processContents="skip"/>			


                    case "DataSetName":
                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            foreach (StiBand band in tableHeaders) component.Components.Add(band);
            foreach (StiBand band in tableDetails)
            {
                (band as StiDataBand).DataSourceName = datasetName;
                component.Components.Add(band);
            }
            foreach (StiBand band in tableFooters) component.Components.Add(band);

            foreach (StiBand band in component.Components)
            {
                double offset = 0;
                for (int index = 0; index < tableColumns.Count; index++)
                {
                    if (index < band.Components.Count)
                    {
                        band.Components[index].Left = offset;
                        band.Components[index].Width = (double)tableColumns[index];
                        band.Components[index].Height = band.Height;
                        offset += (double)tableColumns[index];
                    }
                }
            }

            container.Components.Add(component);
        }
        #endregion

        #region Table columns
        private ArrayList ProcessTableColumnsType(XmlNode baseNode)
        {
            ArrayList list = new ArrayList();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {

                    case "TableColumn":
                        list.Add(ProcessTableColumnType(node));
                        break;

                    //-----<xsd:complexType name="TableColumnsType">
                    //<xsd:sequence>
                    //    <xsd:element name="TableColumn" type="TableColumnType" maxOccurs="unbounded"/>
                    //</xsd:sequence>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return list;
        }

        private double ProcessTableColumnType(XmlNode baseNode)
        {
            double width = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {

                    case "Width":
                        width = ToHi(GetNodeTextValue(node));
                        break;

                    //-----<xsd:complexType name="TableColumnType">
                    //<xsd:element name="Width" type="SizeType"/>

                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>
                    //<xsd:element name="FixedHeader" type="xsd:boolean" minOccurs="0"/>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return width;
        }
        #endregion

        #region Table rows
        private ArrayList ProcessHeaderType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            ArrayList tableRows = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TableRows":
                        tableRows = ProcessTableRowsType(node, container, page, new StiHeaderBand(), dataset);
                        break;

                    //-----<xsd:complexType name="HeaderType">
                    //<xsd:element name="TableRows" type="TableRowsType"/>

                    //<xsd:element name="FixedHeader" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="RepeatOnNewPage" type="xsd:boolean" minOccurs="0"/>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return tableRows;
        }

        private ArrayList ProcessDetailsType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            ArrayList tableRows = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TableRows":
                        tableRows = ProcessTableRowsType(node, container, page, new StiDataBand(), dataset);
                        break;

                    //-----<xsd:complexType name="DetailsType">
                    //<xsd:element name="TableRows" type="TableRowsType"/>

                    //<xsd:element name="Grouping" type="GroupingType" minOccurs="0"/>
                    //<xsd:element name="Sorting" type="SortingType" minOccurs="0"/>
                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return tableRows;
        }

        private ArrayList ProcessFooterType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            ArrayList tableRows = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TableRows":
                        tableRows = ProcessTableRowsType(node, container, page, new StiFooterBand(), dataset);
                        break;

                    //-----<xsd:complexType name="FooterType">
                    //<xsd:element name="TableRows" type="TableRowsType"/>

                    //<xsd:element name="RepeatOnNewPage" type="xsd:boolean" minOccurs="0"/>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return tableRows;
        }


        private ArrayList ProcessTableRowsType(XmlNode baseNode, StiContainer container, StiPage page, StiBand bandType, string dataset)
        {
            ArrayList list = new ArrayList();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {

                    case "TableRow":
                        list.Add(ProcessTableRowType(node, container, page, bandType, dataset));
                        break;

                    //-----<xsd:complexType name="TableRowsType">
                    //<xsd:sequence>
                    //    <xsd:element name="TableRow" type="TableRowType" maxOccurs="unbounded"/>
                    //</xsd:sequence>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return list;
        }

        private StiBand ProcessTableRowType(XmlNode baseNode, StiContainer container, StiPage page, StiBand bandType, string dataset)
        {
            StiBand component = null;
            if (bandType is StiHeaderBand) component = new StiHeaderBand();
            if (bandType is StiDataBand) component = new StiDataBand();
            if (bandType is StiFooterBand) component = new StiFooterBand();

            component.Page = page;
            component.CanGrow = true;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TableCells":
                        ProcessTableCellsType(node, component, page, dataset);
                        break;

                    case "Height":
                        component.Height = ToHi(GetNodeTextValue(node));
                        break;

                    //-----<xsd:complexType name="TableRowType">
                    //<xsd:element name="TableCells" type="TableCellsType"/>
                    //<xsd:element name="Height" type="SizeType"/>

                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return component;
        }
        #endregion

        #region Table cells
        private void ProcessTableCellsType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {

                    case "TableCell":
                        ProcessTableCellType(node, container, page, dataset);
                        break;

                    //-----<xsd:complexType name="TableCellsType">
                    //<xsd:sequence>
                    //    <xsd:element name="TableCell" type="TableCellType" maxOccurs="unbounded"/>
                    //</xsd:sequence>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessTableCellType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {

                    case "ReportItems":
                        ProcessReportItemsType(node, container, page, dataset);
                        break;

                    //-----<xsd:complexType name="TableCellType">
                    //<xsd:element name="ReportItems" type="ReportItemsType"/>

                    //<xsd:element name="ColSpan" type="xsd:unsignedInt" minOccurs="0"/>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #region List
        private void ProcessListType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiPanel component = new StiPanel();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value + "Panel");
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;

            StiDataBand band = new StiDataBand();
            band.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            band.Page = page;
            band.CanGrow = true;

            //first pass - get dataset name
            string datasetName = null;
            string fullDatasetName = null;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSetName":
                        datasetName = Convert.ToString(GetNodeTextValue(node));
                        break;
                }
            }
            fullDatasetName = MakeFullDataSetName(dataset, datasetName);

            //second pass
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
                        component.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
                        component.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
                        component.Width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        component.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, component);
                        break;

                    case "ReportItems":
                        ProcessReportItemsType(node, band, page, fullDatasetName);
                        break;

                    //-----<xsd:complexType name="ListType">
                    //<xsd:element name="Left" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Top" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Width" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Height" type="SizeType" minOccurs="0"/>
                    //<xsd:element name="Style" type="StyleType" minOccurs="0"/>
                    //<xsd:element name="DataSetName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="ReportItems" type="ReportItemsType" minOccurs="0"/>

                    //<xsd:element name="ZIndex" type="xsd:unsignedInt" minOccurs="0"/>

                    //<xsd:element name="Action" type="ActionType" minOccurs="0"/>
                    //<xsd:element name="Visibility" type="VisibilityType" minOccurs="0"/>
                    //<xsd:element name="ToolTip" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Label" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="LinkToChild" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Bookmark" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="RepeatWith" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="CustomProperties" type="CustomPropertiesType" minOccurs="0"/>
                    //<xsd:element name="KeepTogether" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="NoRows" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="PageBreakAtStart" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="PageBreakAtEnd" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="Filters" type="FiltersType" minOccurs="0"/>
                    //<xsd:element name="Grouping" type="GroupingType" minOccurs="0"/>
                    //<xsd:element name="Sorting" type="SortingType" minOccurs="0"/>
                    //<xsd:element name="FillPage" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="DataInstanceName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataInstanceElementOutput" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Output"/>
                    //            <xsd:enumeration value="NoOutput"/>
                    //            <xsd:enumeration value="ContentsOnly"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>
                    //<xsd:element name="DataElementName" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="DataElementOutput" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Output"/>
                    //            <xsd:enumeration value="NoOutput"/>
                    //            <xsd:enumeration value="ContentsOnly"/>
                    //            <xsd:enumeration value="Auto"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>
                    //<xsd:attribute name="Name" type="xsd:normalizedString" use="required"/>

                    case "DataSetName":
                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            band.Height = component.Height;
            band.DataSourceName = datasetName;
            component.Components.Add(band);
            container.Components.Add(component);
        }
        #endregion

        #region Tablix
        private void ProcessTablixType(XmlNode baseNode, StiContainer baseContainer, StiPage page, string dataset)
        {
            StiPanel container = new StiPanel();
			container.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
			container.Border = new StiAdvancedBorder();
			container.Page = page;
			container.CanGrow = true;
            container.CanBreak = true;

            string datasetName = null;
            string fullDatasetName = null;
            StiFiltersCollection filters = null;
            string[] sortExpr = null;

            GroupsInfo rowsHierarchy = new GroupsInfo() { Page = page };

			//first pass - get dataset name
			foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSetName":
                        datasetName = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "TablixRowHierarchy":
                        ProcessTablixRowHierarchyType(node, dataset, rowsHierarchy);
                        break;

					case "SortExpressions":
						sortExpr = ProcessTablixSortExpressionsType(node, datasetName);
						break;
				}
            }

            fullDatasetName = MakeFullDataSetName(dataset, datasetName);

            List<double> tableColumnsWidths = new List<double>();
            ArrayList tableDetails = new ArrayList();

            //second pass
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
						container.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
						container.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
						container.Width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
						container.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, container);
                        break;

                    case "TablixBody":
                        ProcessTablixBodyType(node, page, fullDatasetName, tableColumnsWidths, tableDetails);
                        break;

                    case "Filters":
                        filters = ProcessFiltersType(node, fullDatasetName);
                        break;


                    case "TablixRowHierarchy":
                    case "DataSetName":
                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            //add bands
            if (rowsHierarchy.HeaderLines.Count > 0)
				AddTablixCellsToContainer(rowsHierarchy, tableColumnsWidths, tableDetails, new StiHeaderBand() { PrintIfEmpty = true }, rowsHierarchy.HeaderLines, container);

            for (int groupIndex = 0; groupIndex < rowsHierarchy.Levels.Count; groupIndex++)
            {
                var group = rowsHierarchy.Levels[groupIndex];
                var groupHeader = new StiGroupHeaderBand();
                groupHeader.Condition = new StiGroupConditionExpression("{" + group.GroupExpression + "}");
                AddTablixCellsToContainer(rowsHierarchy, tableColumnsWidths, tableDetails, groupHeader, group.HeaderLines, container);
            }

            var dataBand = new StiDataBand();
            dataBand.DataSourceName = datasetName;
            if (sortExpr != null && sortExpr.Length > 0)
                dataBand.Sort = sortExpr;
			if (filters != null) dataBand.Filters = filters;
			AddTablixCellsToContainer(rowsHierarchy, tableColumnsWidths, tableDetails, dataBand, rowsHierarchy.DataLines, container);

            bool needGroupFooters = false;
			for (int groupIndex = 0; groupIndex < rowsHierarchy.Levels.Count; groupIndex++)
			{
                if (rowsHierarchy.Levels[groupIndex].FooterLines.Count > 0) needGroupFooters = true;
			}
            if (needGroupFooters)
            {
                for (int groupIndex = 0; groupIndex < rowsHierarchy.Levels.Count; groupIndex++)
                {
                    var group = rowsHierarchy.Levels[rowsHierarchy.Levels.Count - 1 - groupIndex];
                    AddTablixCellsToContainer(rowsHierarchy, tableColumnsWidths, tableDetails, new StiGroupFooterBand(), group.FooterLines, container);
                }
            }

			if (rowsHierarchy.FooterLines.Count > 0)
				AddTablixCellsToContainer(rowsHierarchy, tableColumnsWidths, tableDetails, new StiFooterBand() { PrintIfEmpty = true }, rowsHierarchy.FooterLines, container);

			baseContainer.Components.Add(container);
        }

        private void AddTablixCellsToContainer(GroupsInfo rowsHierarchy, List<double> tableColumnsWidths, ArrayList tableDetails, StiBand baseBand, List<int> lines, StiContainer baseContainer)
		{
			baseContainer.Components.Add(baseBand);

			baseBand.Locked = true;
            baseBand.Height = 0;
            double offsetY = 0;

			for (int indexLine = 0; indexLine < lines.Count; indexLine++)
            {
                int lineIndex = lines[indexLine];
				var container = tableDetails[lineIndex] as StiPanel;

				double offset = 0;
				int indexComp = 0;
				for (int index = 0; index < tableColumnsWidths.Count; index++)
				{
					StiComponent comp = container.Components[indexComp++];

					int colspan = 0;
					if (comp.TagValue != null)
					{
						colspan = Convert.ToInt32(comp.TagValue) - 1;
					}
					decimal cellWidth = (decimal)tableColumnsWidths[index];
					while (colspan > 0)
					{
						index++;
						cellWidth += (decimal)tableColumnsWidths[index];
						colspan--;
					}

					comp.Left = offset;
					comp.Width = (double)cellWidth;
					comp.Height = container.Height;
					offset += (double)cellWidth;
				}

				if (rowsHierarchy.cells[lineIndex].Count > 0)
				{
					offset = 0;
					for (int index2 = 0; index2 < rowsHierarchy.cells[lineIndex].Count; index2++)
					{
						var cont = rowsHierarchy.cells[lineIndex][index2];
						if (cont != null)
						{
							var comp = cont.Components[0];
                            comp.Left = offset;
							comp.Width = cont.Width;
							comp.Height = container.Height;
							container.Components.Insert(index2, comp);
							offset += cont.Width;
						}
					}
                    for (int index3 = rowsHierarchy.cells[lineIndex].Count; index3 < container.Components.Count; index3++)
                    {
                        container.Components[index3].Left += offset;
                    }
				}

                for (int index = 0; index < container.Components.Count; index++)
                {
                    container.Components[index].GrowToHeight = true;
                }

				if (lines.Count > 1)
                {
                    container.Name = baseBand.Name + "_c" + indexLine.ToString();
                    container.Width = baseContainer.Width;
                    container.Top = offsetY;
                    offsetY += container.Height;
					baseBand.Components.Add(container);
				}
                else
                {
					baseBand.Components.AddRange(container.Components);
				}
				baseBand.Height += container.Height;
			}
		}

		private class GroupLevelInfo
        {
            public List<int> HeaderLines = new List<int>();
			public List<int> FooterLines = new List<int>();
            public bool HasNestedGroup = false;
			public string GroupName = string.Empty;
            public string GroupExpression = string.Empty;
            public string[] SortExpression = null;  //not used in current revision
            public List<int> Lines = new List<int>();
		}

        private class GroupsInfo
        {
            public List<GroupLevelInfo> Levels = new List<GroupLevelInfo>();
            public int Line = 0;
            public int Column = 0;
            public List<List<StiContainer>> cells = new List<List<StiContainer>>();
            public List<string> LineTypes = new List<string>();
            public StiPage Page = null;
			public List<int> HeaderLines = new List<int>();
			public List<int> FooterLines = new List<int>();
			public List<int> DataLines = new List<int>();
		}

		private void ProcessTablixBodyType(XmlNode baseNode, StiPage page, string dataset, List<double> tableColumns, ArrayList tableDetails)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixColumns":
                        var tableColumns2 = ProcessTablixColumnsType(node);
                        tableColumns.AddRange(tableColumns2);
                        break;

                    case "TablixRows":
                        ArrayList tableDetails2 = ProcessTablixRowsType(node, page, dataset);
                        tableDetails.AddRange(tableDetails2);
                        break;

                    //case "ZIndex":
                    //    //ignored or not implemented yet
                    //    break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private ArrayList ProcessTablixRowsType(XmlNode baseNode, StiPage page, string dataset)
        {
            ArrayList list = new ArrayList();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixRow":
                        list.Add(ProcessTablixRowType(node, page, dataset));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return list;
        }

        private StiComponent ProcessTablixRowType(XmlNode baseNode, StiPage page, string dataset)
        {
            StiPanel container = new StiPanel();

			container.Page = page;
			container.CanGrow = true;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixCells":
                        ProcessTablixCellsType(node, container, page, dataset);
                        break;

                    case "Height":
						container.Height = ToHi(GetNodeTextValue(node));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return container;
        }

        private void ProcessTablixCellsType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixCell":
                        ProcessTablixCellType(node, container, page, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessTablixCellType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ReportItems":
                    case "CellContents":
                        ProcessReportItemsType(node, container, page, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private List<double> ProcessTablixColumnsType(XmlNode baseNode)
        {
            var list = new List<double>();
			foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixColumn":
                        list.Add(ProcessTablixColumnType(node));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return list;
        }

        private double ProcessTablixColumnType(XmlNode baseNode)
        {
            double width = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Width":
                        width = ToHi(GetNodeTextValue(node));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return width;
        }

        private void ProcessTablixRowHierarchyType(XmlNode baseNode, string dataSet, GroupsInfo groupsLevels)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixMembers":
                        ProcessTablixMembersType(node, dataSet, groupsLevels, null, false);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (groupsLevels.Levels.Count > 0)
			{
				var fg = groupsLevels.Levels[0];
				for (int index = 0; index < fg.Lines[0]; index++)
                {
					groupsLevels.HeaderLines.Add(index);
					groupsLevels.LineTypes[index] = "h";
				}
				for (int index = fg.Lines[fg.Lines.Count - 1] + 1; index < groupsLevels.LineTypes.Count; index++)
				{
					groupsLevels.FooterLines.Add(index);
					groupsLevels.LineTypes[index] = "f";
				}
			}
            if (groupsLevels.Levels.Count > 0)
            {
                GroupLevelInfo group = groupsLevels.Levels[groupsLevels.Levels.Count - 1];
                if (group.GroupExpression.Length == 0)
                {
                    //change last group to data
                    foreach (int line in group.Lines)
                    {
                        groupsLevels.LineTypes[line] = "d";
                    }
                    groupsLevels.DataLines.AddRange(group.Lines);
                    groupsLevels.Levels.Remove(group);
                }
                else
                {
                    //no data lines, only groupheader
                    group.HeaderLines.AddRange(group.Lines);
                }
            }			

		}

        private void ProcessTablixMembersType(XmlNode baseNode, string dataSet, GroupsInfo groupsInfo, GroupLevelInfo group, bool incrementColumn)
        {
            int column = groupsInfo.Column + (incrementColumn ? 1 : 0);
            bool hasGroup = false;
			foreach (XmlNode node in baseNode.ChildNodes)
            {
                groupsInfo.Column = column;
				int currentLine = groupsInfo.Line;
				bool isGroup = false;

				switch (node.Name)
                {
                    case "TablixMember":
                        isGroup = ProcessTablixMemberType(node, dataSet, groupsInfo, group);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }

				if (group != null && isGroup) group.HasNestedGroup = true;
                hasGroup |= isGroup;

				if (!isGroup)
				{
					if (group != null)
					{
						if (!group.HasNestedGroup)
						{
							group.HeaderLines.Add(currentLine);
							groupsInfo.LineTypes[currentLine] = "gh";
						}
						else
						{
							group.FooterLines.Add(currentLine);
							groupsInfo.LineTypes[currentLine] = "gf";
						}
					}				
				}

			}
		}

        private bool ProcessTablixMemberType(XmlNode baseNode, string dataSet, GroupsInfo groupsInfo, GroupLevelInfo baseGroup)
        {
            CheckCells(groupsInfo, groupsInfo.Line, -1);
            /*if (groupsInfo.Line + 1 > groupsInfo.cells.Count)
            {
                groupsInfo.cells.Add(new List<StiComponent>());
            }*/

            bool hasHeader = false;
            bool hasMembers = false;
            GroupLevelInfo group = null;
            int currentLine = groupsInfo.Line;
            int currentColumn = groupsInfo.Column;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Group":
                        group = ProcessTablixGroupType(node, dataSet, groupsInfo);
                        //hasGroup = true;
                        //if (baseGroup != null) baseGroup.HasNestedGroup = true;
                        break;

                    case "SortExpressions":
						group.SortExpression = ProcessTablixSortExpressionsType(node, dataSet);
						break;

                    case "TablixHeader":
                        ProcessTablixHeaderType(node, dataSet, groupsInfo);
                        hasHeader = true;
                        break;

                    case "TablixMembers":
                        ProcessTablixMembersType(node, dataSet, groupsInfo, group, hasHeader);
                        hasMembers = true;
                        break;

                    case "KeepWithGroup":
                    case "KeepTogether":
                    case "DataElementName":
                    case "DataElementOutput":
                        //ignored in this revision
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            //duplicate textbox for all groupheader
            if (hasHeader && (groupsInfo.Line > currentLine) && (groupsInfo.cells[currentLine][currentColumn] != null))
            {
                for (int index = currentLine + 1; index < groupsInfo.Line; index++)
                {
                    CheckCells(groupsInfo, index, currentColumn);

                    var cont = groupsInfo.cells[currentLine][currentColumn];
					var newCont = new StiPanel() { Width = cont.Width };

                    var newComp = cont.Components[0].Clone() as StiComponent;
                    newComp.Name += "_" + index.ToString();
                    (newComp as StiText).ProcessingDuplicates = StiProcessingDuplicatesType.GlobalMerge;

                    newCont.Components.Add(newComp);
					groupsInfo.cells[index][currentColumn] = newCont;
				}
                (groupsInfo.cells[currentLine][currentColumn].Components[0] as StiText).ProcessingDuplicates = StiProcessingDuplicatesType.GlobalMerge;
			}		

			if (/* (group == null) && !hasHeader || */ !hasMembers)
            {
                groupsInfo.Line++;
            }

			if (group != null)
			{
				for (int index = currentLine; index < groupsInfo.Line; index++)
				{
					group.Lines.Add(index);
				}
			}

			return group != null;
        }


		private GroupLevelInfo ProcessTablixGroupType(XmlNode baseNode, string dataSet, GroupsInfo groupsInfo)
        {
			GroupLevelInfo group = new GroupLevelInfo();
			groupsInfo.Levels.Add(group);
			//group.IsGroup = true;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupExpressions":
                        ProcessGroupExpressionsType(node, dataSet, group);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return group;
        }

        private void ProcessGroupExpressionsType(XmlNode baseNode, string dataSet, GroupLevelInfo group)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupExpression":
                        string st = ConvertExpression(GetNodeTextValue(node), dataSet);
                        if (group.GroupExpression.Length > 0)
                        {
							group.GroupExpression += ".ToString() + ";      //!!! todo - check what comes here
                        }
						group.GroupExpression += st.Substring(1, st.Length - 2);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

		private string[] ProcessTablixSortExpressionsType(XmlNode baseNode, string dataSet)
		{
            List<string> sort = new List<string>();
			foreach (XmlNode node in baseNode.ChildNodes)
			{
				switch (node.Name)
				{
					case "SortExpression":
						ProcessTablixSortExpressionType(node, dataSet, sort);
						break;

					default:
						ThrowError(baseNode.Name, node.Name);
						break;
				}
			}
            return sort.ToArray();
		}

		private void ProcessTablixSortExpressionType(XmlNode baseNode, string dataSet, List<string> sort)
		{
            string expr = string.Empty;
            string direction = "ASC";
			foreach (XmlNode node in baseNode.ChildNodes)
			{
				switch (node.Name)
				{
					case "Value":
						expr = ConvertExpression(GetNodeTextValue(node), dataSet);
						break;

					case "Direction":
						string st = GetNodeTextValue(node);
                        if (st == "Descending") direction = "DESC";
						break;

					default:
						ThrowError(baseNode.Name, node.Name);
						break;
				}
			}
            if (expr.Length > 0)
            {
                sort.Add(direction);
                sort.Add(expr);
            }
		}

		private void ProcessTablixHeaderType(XmlNode baseNode, string dataSet, GroupsInfo groupsInfo)
		{
            double size = 0;
			foreach (XmlNode node in baseNode.ChildNodes)
			{
				switch (node.Name)
				{
					case "Size":
						size = ToHi(GetNodeTextValue(node));
						break;

					case "CellContents":
                        var cont = new StiContainer() { Name = $"Cont{groupsInfo.Line}*{groupsInfo.Column}" };
                        cont.Width = size;
						ProcessReportItemsType(node, cont, groupsInfo.Page, dataSet);
						CheckCells(groupsInfo, groupsInfo.Line, groupsInfo.Column);
                        groupsInfo.cells[groupsInfo.Line][groupsInfo.Column] = cont;
						break;

					default:
						ThrowError(baseNode.Name, node.Name);
						break;
				}
			}
		}

        private void CheckCells(GroupsInfo groupsInfo, int line, int column)
        {
			for (int index = 0; index <= line; index++)
			{
                if (groupsInfo.cells.Count <= line) groupsInfo.cells.Add(new List<StiContainer>());
				while (groupsInfo.cells[index].Count <= column) groupsInfo.cells[index].Add(null);

                if (groupsInfo.LineTypes.Count <= line) groupsInfo.LineTypes.Add(string.Empty);
			}
		}
		#endregion

		#region Subreport
		private void ProcessSubreportType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            StiSubReport component = new StiSubReport();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Border = new StiAdvancedBorder();
            component.Page = page;
            component.CanGrow = true;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Left":
                        component.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "Top":
                        component.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "Width":
                        component.Width = ToHi(GetNodeTextValue(node));
                        break;
                    case "Height":
                        component.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "Style":
                        ProcessStyleType(node, component);
                        break;

                    //case "ReportName":
                    //    component.SubReportPage = GetNodeTextValue(node);
                    //    break;


                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            container.Components.Add(component);
        }
        #endregion

        #region Filters
        private StiFiltersCollection ProcessFiltersType(XmlNode baseNode, string dataSet)
        {
            StiFiltersCollection filters = new StiFiltersCollection();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Filter":
                        ProcessFilterType(node, filters, dataSet);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return filters;
        }

        private void ProcessFilterType(XmlNode baseNode, StiFiltersCollection filters, string dataSet)
        {
            string expression = string.Empty;
            string op = string.Empty;
            string val = string.Empty;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "FilterExpression":
                        expression = ConvertExpression(GetNodeTextValue(node), dataSet);
                        break;

                    case "Operator":
                        op = GetNodeTextValue(node);
                        break;

                    case "FilterValues":
                        val = ProcessFilterValueType(node);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            expression = expression.Substring(1, expression.Length - 2);

            if (op == "Equal") op = "==";
            if (op == "NotEqual") op = "!=";
            if (op == "GreaterThan") op = ">";
            if (op == "GreaterThanOrEqual") op = ">=";
            if (op == "LessThan") op = "<";
            if (op == "LessThanOrEqual") op = "<=";

            StiFilter filter = new StiFilter(expression + " " + op + " " + val);
            filters.Add(filter);
        }

        private string ProcessFilterValueType(XmlNode baseNode)
        {
            string val = string.Empty;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "FilterValue":
                        val = GetNodeTextValue(node);
                        return val;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return val;
        }
        #endregion

        #endregion

        #region DataSources
        private void ProcessDataSourcesType(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSource":
                        ProcessDataSourceType(node, report);
                        break;

                    //-----<xsd:complexType name="DataSourcesType">
                    //<xsd:sequence>
                    //    <xsd:element name="DataSource" type="DataSourceType" maxOccurs="unbounded"/>
                    //</xsd:sequence>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessDataSourceType(XmlNode baseNode, StiReport report)
        {
            StiDatabase dataBase = null;
            string databaseName = baseNode.Attributes["Name"].Value;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ConnectionProperties":
                        dataBase = ProcessConnectionPropertiesType(node);
                        break;

                    //-----<xsd:complexType name="DataSourceType">
                    //<xsd:element name="ConnectionProperties" type="ConnectionPropertiesType" minOccurs="0"/>

                    //<xsd:element name="Transaction" type="xsd:boolean" minOccurs="0"/>
                    //<xsd:element name="DataSourceReference" type="xsd:string" minOccurs="0"/>

                    case "rd:DataSourceID":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            if (dataBase != null)
            {
                report.Dictionary.Databases.Add(dataBase);
            }
        }

        private StiDatabase ProcessConnectionPropertiesType(XmlNode baseNode)
        {
            StiDatabase dataBase = null;
            string databaseName = baseNode.ParentNode.Attributes["Name"].Value;
            string dataProvider = null;
            string connectString = $"Data Source=.;Initial Catalog={databaseName};Integrated Security=True";

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataProvider":
                        dataProvider = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "ConnectString":
                        connectString = Convert.ToString(GetNodeTextValue(node));
                        break;

					case "IntegratedSecurity":
						var st = Convert.ToString(GetNodeTextValue(node)).ToLowerInvariant();
                        if (st == "true") connectString += ";Integrated Security=True";
						break;

					//-----<xsd:complexType name="ConnectionPropertiesType">
					//<xsd:element name="DataProvider" type="xsd:string"/>
					//<xsd:element name="ConnectString" type="xsd:string"/>

					//<xsd:element name="Prompt" type="xsd:string" minOccurs="0"/>

					default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            switch (dataProvider)
            {
                case "OLEDB":
                    dataBase = new StiOleDbDatabase(databaseName, connectString);
                    break;

                default:
                    dataBase = new StiSqlDatabase(databaseName, connectString);
                    break;
            }
            return dataBase;
        }

        #endregion

        #region DataSets
        private void ProcessDataSetsType(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSet":
                        ProcessDataSetType(node, report);
                        break;

                    //-----<xsd:complexType name="DataSetsType">
                    //<xsd:sequence>
                    //    <xsd:element name="DataSet" type="DataSetType" maxOccurs="unbounded"/>
                    //</xsd:sequence>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessDataSetType(XmlNode baseNode, StiReport report)
        {
            StiDataSource dataSource = null;
            ArrayList columns = new ArrayList();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Fields":
                        ProcessFieldsType(node, columns);
                        break;

                    case "Query":
                        dataSource = ProcessQueryType(node, report);
                        break;

                    //-----<xsd:complexType name="DataSetType">
                    //<xsd:element name="Fields" type="FieldsType" minOccurs="0"/>
                    //<xsd:element name="Query" type="QueryType"/>

                    //<xsd:element name="Collation" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="CaseSensitivity" string = "True", "False", "Auto"
                    //<xsd:element name="AccentSensitivity" string = "True", "False", "Auto"
                    //<xsd:element name="KanatypeSensitivity" string = "True", "False", "Auto"
                    //<xsd:element name="WidthSensitivity" string = "True", "False", "Auto"
                    //<xsd:element name="Filters" type="FiltersType" minOccurs="0"/>
                    //<xsd:any namespace="##other" processContents="skip"/>			

                    case "rd:DataSetInfo":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (dataSource != null)
            {
                foreach (StiDataColumn column in columns)
                {
                    column.DataSource = dataSource;
                    dataSource.Columns.Add(column);

                    string baseField = $"Fields!{column.Name}.Value";
                    string newField = $"{dataSource.Name}.{column.Name}";
                    fieldsNames[baseField] = newField;
                    fieldsNames[dataSource.Name + ":" + baseField] = newField;
                }
                report.Dictionary.DataSources.Add(dataSource);
            }
        }

        private void ProcessFieldsType(XmlNode baseNode, ArrayList columns)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Field":
                        ProcessFieldType(node, columns);
                        break;

                    //-----<xsd:complexType name="FieldsType">
                    //<xsd:sequence>
                    //    <xsd:element name="Field" type="FieldType" maxOccurs="unbounded"/>
                    //</xsd:sequence>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessFieldType(XmlNode baseNode, ArrayList columns)
        {
            StiDataColumn column = new StiDataColumn();
            column.Name = baseNode.Attributes["Name"].Value;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataField":
                        column.NameInSource = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "rd:TypeName":
                        string typeName = Convert.ToString(GetNodeTextValue(node));
                        Type type = Type.GetType(typeName);
                        if (type == null) type = typeof(object);
                        column.Type = type;
                        break;

                    //-----<xsd:complexType name="FieldType">
                    //<xsd:element name="DataField" type="xsd:string" minOccurs="0"/>

                    //<xsd:element name="Value" type="xsd:string" minOccurs="0"/>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            columns.Add(column);
        }

        private StiDataSource ProcessQueryType(XmlNode baseNode, StiReport report)
        {
            StiDataSource dataSource = null;
            string databaseName = null;
            string commandText = null;
            string commandType = null;
            StiDataParametersCollection parameters = new StiDataParametersCollection();
            int timeout = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSourceName":
                        databaseName = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "CommandText":
                        commandText = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "CommandType":
                        commandType = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "Timeout":
                        timeout = Convert.ToInt32(GetNodeTextValue(node));
                        break;

                    case "QueryParameters":
                        ProcessQueryParameters(node, parameters);
                        break;

                    //-----<xsd:complexType name="QueryType">
                    //<xsd:element name="DataSourceName" type="xsd:string"/>
                    //<xsd:element name="CommandText" type="xsd:string"/>
                    //<xsd:element name="CommandType" minOccurs="0">
                    //    <xsd:simpleType>
                    //        <xsd:restriction base="xsd:string">
                    //            <xsd:enumeration value="Text"/>
                    //            <xsd:enumeration value="StoredProcedure"/>
                    //            <xsd:enumeration value="TableDirect"/>
                    //        </xsd:restriction>
                    //    </xsd:simpleType>
                    //</xsd:element>
                    //<xsd:element name="Timeout" type="xsd:unsignedInt" minOccurs="0"/>
                    //<xsd:element name="QueryParameters" type="QueryParametersType" minOccurs="0"/>

                    case "rd:UseGenericDesigner":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            StiDatabase database = report.Dictionary.Databases[databaseName];
            if (database != null)
            {
                if (database is StiOleDbDatabase)
                {
                    StiOleDbSource oledbSource = new StiOleDbSource();
                    oledbSource.NameInSource = databaseName;
                    oledbSource.SqlCommand = commandText;
                    if (timeout > 0) oledbSource.CommandTimeout = timeout;
                    dataSource = oledbSource;
                }
            }
            if (dataSource == null)
            {
                if (!string.IsNullOrWhiteSpace(commandText))
                {
                    StiSqlSource sqlSource = new StiSqlSource();
                    sqlSource.NameInSource = databaseName;
                    sqlSource.SqlCommand = commandText;
                    if (commandType == "StoredProcedure") sqlSource.Type = StiSqlSourceType.StoredProcedure;
                    if (timeout > 0) sqlSource.CommandTimeout = timeout;
                    dataSource = sqlSource;
                }
                else
                {
                    StiDataTableSource datatableSource = new StiDataTableSource();
                    datatableSource.NameInSource = databaseName;
                    dataSource = datatableSource;
                }
            }
            dataSource.Name = baseNode.ParentNode.Attributes["Name"].Value;
            dataSource.Alias = dataSource.Name;
            dataSource.Parameters.AddRange(parameters);
            return dataSource;
        }

        private void ProcessQueryParameters(XmlNode baseNode, StiDataParametersCollection parameters)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "QueryParameter":
                        parameters.Add(ProcessQueryParameter(node));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private StiDataParameter ProcessQueryParameter(XmlNode baseNode)
        {
            StiDataParameter parameter = new StiDataParameter();
            parameter.Type = 18;    //Text
            parameter.Name = baseNode.Attributes["Name"].Value.Substring(1);    //remove @
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        string st = ConvertExpression(GetNodeTextValue(node), null);
                        if (st.StartsWith("{") && st.EndsWith("}"))
                        {
                            st = st.Substring(1, st.Length - 2);
                        }
                        parameter.Expression = st;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return parameter;
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

                    //-----<xsd:complexType name="EmbeddedImagesType">
                    //<xsd:sequence>
                    //    <xsd:element name="EmbeddedImage" type="EmbeddedImageType" maxOccurs="unbounded"/>
                    //</xsd:sequence>

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

                    //-----<xsd:complexType name="EmbeddedImageType">
                    //<xsd:element name="ImageData" type="xsd:string"/>
                    //<xsd:element name="MIMEType" type="xsd:string"/>

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

        #region Style
        private void ProcessStyleType(XmlNode baseNode, StiComponent component)
        {
            IStiBorder border = component as IStiBorder;
            IStiBrush brush = component as IStiBrush;
            IStiFont font = component as IStiFont;
            StiText text = component as StiText;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Border":
                        ProcessBorder(node, component);
                        break;

                    case "BorderColor":
                        string[] colors = ProcessBorderColorStyleWidthType(node);
                        if (border != null)
                        {
                            if (border.Border is StiAdvancedBorder)
                            {
                                StiAdvancedBorder advBorder = border.Border as StiAdvancedBorder;
                                advBorder.LeftSide.Color = ParseColor(colors[0]);
                                advBorder.RightSide.Color = ParseColor(colors[1]);
                                advBorder.TopSide.Color = ParseColor(colors[2]);
                                advBorder.BottomSide.Color = ParseColor(colors[3]);
                            }
                            else
                            {
                                border.Border.Color = ParseColor(colors[0]);
                            }
                        }
                        break;

                    case "BorderWidth":
                        string[] widths = ProcessBorderColorStyleWidthType(node);
                        if (border != null)
                        {
                            if (border.Border is StiAdvancedBorder)
                            {
                                StiAdvancedBorder advBorder = border.Border as StiAdvancedBorder;
                                if (widths[0] != null) advBorder.LeftSide.Size = ToHi(widths[0]);
                                if (widths[1] != null) advBorder.RightSide.Size = ToHi(widths[1]);
                                if (widths[2] != null) advBorder.TopSide.Size = ToHi(widths[2]);
                                if (widths[3] != null) advBorder.BottomSide.Size = ToHi(widths[3]);
                            }
                            else
                            {
                                if (widths[0] != null) border.Border.Size = ToHi(widths[0]);
                            }
                        }
                        break;

                    case "BorderStyle":
                        string[] styles = ProcessBorderColorStyleWidthType(node);
                        if (border != null)
                        {
                            if (border.Border is StiAdvancedBorder)
                            {
                                StiAdvancedBorder advBorder = border.Border as StiAdvancedBorder;
                                advBorder.LeftSide.Style = ParseStyle(styles[0]);
                                advBorder.RightSide.Style = ParseStyle(styles[1]);
                                advBorder.TopSide.Style = ParseStyle(styles[2]);
                                advBorder.BottomSide.Style = ParseStyle(styles[3]);
                            }
                            else
                            {
                                border.Border.Style = ParseStyle(styles[0]);
                            }
                        }
                        break;

                    case "PaddingLeft":
                        if (text != null) text.Margins.Left = ToHi(GetNodeTextValue(node));
                        break;
                    case "PaddingRight":
                        if (text != null) text.Margins.Right = ToHi(GetNodeTextValue(node));
                        break;
                    case "PaddingTop":
                        if (text != null) text.Margins.Top = ToHi(GetNodeTextValue(node));
                        break;
                    case "PaddingBottom":
                        if (text != null) text.Margins.Bottom = ToHi(GetNodeTextValue(node));
                        break;

                    case "FontSize":
                        if (font != null)
                        {
                            try
                            {
                                font.Font = new Font(font.Font.Name, (float)ToHi(GetNodeTextValue(node)), font.Font.Style);
                            }
                            catch
                            {
                                font.Font = new Font(font.Font.Name, 8, font.Font.Style);
                            }
                        }
                        break;
                    case "FontFamily":
                        if (font != null) font.Font = new Font(GetNodeTextValue(node), font.Font.Size, font.Font.Style);
                        break;
                    case "FontStyle":
                        if (font != null)
                        {
                            if (GetNodeTextValue(node) == "Italic")
                                font.Font = new Font(font.Font.Name, font.Font.Size, font.Font.Style | FontStyle.Italic);
                        }
                        break;
                    case "FontWeight":
                        string fontWeight = GetNodeTextValue(node);
                        if ((font != null) && !string.IsNullOrEmpty(fontWeight))
                        {
                            bool needBold = false;
                            if (fontWeight.ToLowerInvariant() == "bold")
                            {
                                needBold = true;
                            }
                            else if (fontWeight.ToLowerInvariant() == "normal")
                            {
                                needBold = false;
                            }
                            else
                            {
                                try
                                {
                                    if (Convert.ToDouble(fontWeight) > 550) needBold = true;
                                }
                                catch
                                {
                                }
                            }
                            if (needBold)
                            {
                                font.Font = new Font(font.Font.Name, font.Font.Size, font.Font.Style | FontStyle.Bold);
                            }
                        }
                        break;

                    case "BackgroundColor":
                        if (brush != null) (brush.Brush as StiSolidBrush).Color = ParseColor(GetNodeTextValue(node));
                        break;

                    case "Color":
                        if (text != null) (text.TextBrush as StiSolidBrush).Color = ParseColor(GetNodeTextValue(node));
                        break;

                    case "TextAlign":
                        if (text != null) text.HorAlignment = ParseTextAlign(GetNodeTextValue(node));
                        break;

                    case "VerticalAlign":
                        if (text != null) text.VertAlignment = ParseVerticalAlign(GetNodeTextValue(node));
                        break;

                    case "TextDecoration":
                        if (font != null)
                        {
                            if (GetNodeTextValue(node) == "Underline")
                                font.Font = new Font(font.Font.Name, font.Font.Size, font.Font.Style | FontStyle.Underline);
                            if (GetNodeTextValue(node) == "LineThrough")
                                font.Font = new Font(font.Font.Name, font.Font.Size, font.Font.Style | FontStyle.Strikeout);
                        }
                        break;

                    case "Direction":
                        if (text != null)
                        {
                            if (GetNodeTextValue(node) == "RTL") text.TextOptions.RightToLeft = true;
                        }
                        break;

                    case "Format":
                        string format = GetNodeTextValue(node);
                        if (!string.IsNullOrWhiteSpace(format) && text != null)
                        {
                            if (format == "P") text.TextFormat = new StiPercentageFormatService();
                            else if (format == "C") text.TextFormat = new StiCurrencyFormatService();
                            else text.TextFormat = new StiCustomFormatService(format);
                        }
                        break;

                    case "LeftBorder":
                        ProcessBorderSide(node, component, StiBorderSides.Left);
                        break;
                    case "RightBorder":
                        ProcessBorderSide(node, component, StiBorderSides.Right);
                        break;
                    case "TopBorder":
                        ProcessBorderSide(node, component, StiBorderSides.Top);
                        break;
                    case "BottomBorder":
                        ProcessBorderSide(node, component, StiBorderSides.Bottom);
                        break;

                    case "BackgroundImage":
                        ProcessBackgroundImage(node, component);
                        break;


                    //-----<xsd:complexType name="StyleType">
                    //<xsd:element name="BorderColor" type="BorderColorStyleWidthType" minOccurs="0"/>
                    //<xsd:element name="BorderStyle" type="BorderColorStyleWidthType" minOccurs="0"/>
                    //<xsd:element name="BorderWidth" type="BorderColorStyleWidthType" minOccurs="0"/>
                    //<xsd:element name="PaddingLeft" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="PaddingRight" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="PaddingTop" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="PaddingBottom" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="FontSize" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="FontFamily" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="FontStyle" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="FontWeight" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="BackgroundColor" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Color" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="TextAlign" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="VerticalAlign" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="TextDecoration" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Direction" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Format" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="BackgroundImage" type="BackgroundImageType" minOccurs="0"/>

                    //<xsd:element name="BackgroundGradientType" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="BackgroundGradientEndColor" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="LineHeight" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="WritingMode" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Language" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="UnicodeBiDi" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Calendar" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="NumeralLanguage" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="NumeralVariant" type="xsd:string" minOccurs="0"/>

                    //case "InteractiveHeight":
                    //case "InteractiveWidth":
                    //ignored or not implemented yet
                    //break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessBorderSide(XmlNode baseNode, StiComponent component, StiBorderSides borderSide)
        {
            IStiBorder border = component as IStiBorder;
            if (border == null) return;

            if (!(border.Border is StiAdvancedBorder)) border.Border = new StiAdvancedBorder();
            var adv = border.Border as StiAdvancedBorder;
            StiBorderSide side = new StiBorderSide();
            if (borderSide == StiBorderSides.Left) side = adv.LeftSide;
            if (borderSide == StiBorderSides.Right) side = adv.RightSide;
            if (borderSide == StiBorderSides.Top) side = adv.TopSide;
            if (borderSide == StiBorderSides.Bottom) side = adv.BottomSide;

            var tempText = new StiText();
            ProcessBorder(baseNode, tempText);

            side.Color = tempText.Border.Color;
            side.Size = tempText.Border.Size;
            side.Style = tempText.Border.Style;
        }

        private void ProcessBackgroundImage(XmlNode baseNode, StiComponent component)
        {
            StiImage image = ProcessImageType(baseNode, null, null, null);
            backgroundImages[image] = component;
        }

        private void ProcessBorder(XmlNode baseNode, StiComponent component)
        {
            IStiBorder border = component as IStiBorder;
            if (border == null) return;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Color":
                        border.Border.Color = ParseColor(GetNodeTextValue(node));
                        break;

                    case "Width":
                        string borderWidth = GetNodeTextValue(node);
                        if (!string.IsNullOrEmpty(borderWidth)) border.Border.Size = ToHi(borderWidth);
                        break;

                    case "Style":
                        border.Border.Style = ParseStyle(GetNodeTextValue(node));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private string[] ProcessBorderColorStyleWidthType(XmlNode baseNode)
        {
            string[] arrayResult = new string[4];
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Default":
                        arrayResult[0] = GetNodeTextValue(node);
                        arrayResult[1] = arrayResult[0];
                        arrayResult[2] = arrayResult[0];
                        arrayResult[3] = arrayResult[0];
                        break;
                    case "Left":
                        arrayResult[0] = GetNodeTextValue(node);
                        break;
                    case "Right":
                        arrayResult[1] = GetNodeTextValue(node);
                        break;
                    case "Top":
                        arrayResult[2] = GetNodeTextValue(node);
                        break;
                    case "Bottom":
                        arrayResult[3] = GetNodeTextValue(node);
                        break;

                    //-----<xsd:complexType name="BorderColorStyleWidthType">
                    //<xsd:element name="Default" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Left" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Right" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Top" type="xsd:string" minOccurs="0"/>
                    //<xsd:element name="Bottom" type="xsd:string" minOccurs="0"/>

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return arrayResult;
        }

        private static StiPenStyle ParseStyle(string styleAttribute)
        {
            switch (styleAttribute)
            {
                case null:
                case "None":
                    return StiPenStyle.None;

                case "Dotted":
                    return StiPenStyle.Dot;

                case "Dashed":
                    return StiPenStyle.Dash;

                case "Double":
                    return StiPenStyle.Double;
            }
            return StiPenStyle.Solid;
        }

        private static StiTextHorAlignment ParseTextAlign(string alignAttribute)
        {
            switch (alignAttribute)
            {
                case "Center":
                    return StiTextHorAlignment.Center;

                case "Right":
                    return StiTextHorAlignment.Right;
            }
            return StiTextHorAlignment.Left;
        }

        private static StiVertAlignment ParseVerticalAlign(string alignAttribute)
        {
            switch (alignAttribute)
            {
                case "Middle":
                    return StiVertAlignment.Center;

                case "Bottom":
                    return StiVertAlignment.Bottom;
            }
            return StiVertAlignment.Top;
        }

        private static Hashtable HtmlNameToColor = null;

        private static Color ParseColor(string colorAttribute)
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
            List<string> values = new List<string>();
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
                if (defaultValue.StartsWith("="))
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
            List<string> values = new List<string>();
            List<string> labels = new List<string>();
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


                    //-----<xsd:element name="Body" type="BodyType"/>
                    //<xsd:element name="ReportItems" type="ReportItemsType" minOccurs="0"/>
                    //<xsd:element name="Height" type="SizeType"/>
                    //<xsd:element name="Style" type="StyleType" minOccurs="0"/>

                    //<xsd:element name="Columns" type="xsd:unsignedInt" minOccurs="0"/>
                    //<xsd:element name="ColumnSpacing" type="SizeType" minOccurs="0"/>

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

        #region Utils
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

        private double ToHi(string strValue)
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
                //factor = 100 / 72f;
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

        Hashtable embeddedImages = new Hashtable();
        Hashtable fieldsNames = new Hashtable();

        public string ConvertExpression(string baseExpression, string dataset)
        {
            string newExpression = baseExpression;
            if (baseExpression.StartsWith("="))
            {
                newExpression = "{" + baseExpression.Substring(1) + "}";

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

                newExpression = newExpression.Replace("Globals!ExecutionTime", "Time");
                newExpression = newExpression.Replace("Globals!PageNumber", "PageNumber");
                newExpression = newExpression.Replace("Globals!TotalPages", "TotalPageCount");
                newExpression = newExpression.Replace("Globals!OverallPageNumber", "PageNumberThrough");
                newExpression = newExpression.Replace("Globals!OverallTotalPages", "TotalPageCountThrough");
                newExpression = newExpression.Replace("Globals!ReportName", "ReportName");

                int pos = -1;
                while((pos = newExpression.IndexOf("Code.", pos + 1)) != -1)
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
                    if ((args.Length == 1) && (func.ToLowerInvariant() == "count"))
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
            if (baseName == null) return name;
            if (string.IsNullOrWhiteSpace(name)) return baseName;

            string[] parts = baseName.Split(new char[] { ':' });
            if (parts[parts.Length - 1] == name)
            {
                return baseName;
            }

            return baseName + ":" + name;
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