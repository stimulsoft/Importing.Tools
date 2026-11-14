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
using Stimulsoft.Report.BarCodes;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
        #region CustomReportItem
        private void ProcessCustomReportItemType(XmlNode baseNode, StiContainer container, StiPage page, string dataset)
        {
            var component = new StiPanel();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"].Value);
            component.Page = page;
            component.CanGrow = true;

            string type = null;
            XmlNode customData = null;
            XmlNode bandedListConfig = null;
            XmlNode altReportItem = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Type":
                        type = GetNodeTextValue(node);
                        break;

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
                        ProcessStyleType(node, component, dataset);
                        break;

                    case "Visibility":
                        ProcessVisibility(node, component, dataset);
                        break;

                    case "CustomData":
                        //processed together with the layout config below
                        customData = node;
                        break;

                    case "BandedListConfig":
                        bandedListConfig = node;
                        break;

                    case "AltReportItem":
                        altReportItem = node;
                        break;

                    case "CustomProperties":
                    case "ZIndex":
                    case "DataElementName":
                    case "DataElementOutput":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            //a barcode of Bold Reports is a barcode of Stimulsoft in the place of the item
            var barCode = CreateBoldBarCode(baseNode, type, component, dataset);
            if (barCode != null)
            {
                container.Components.Add(barCode);
                return;
            }

            if (string.Equals(type, "PDFSignature", StringComparison.OrdinalIgnoreCase))
            {
                container.Components.Add(CreateBoldPdfSignature(baseNode, component));
                return;
            }

            if (string.Equals(type, "BandedList", StringComparison.OrdinalIgnoreCase) && bandedListConfig != null)
            {
                ProcessBandedList(customData, bandedListConfig, component, page, dataset);
            }
            else if (altReportItem != null)
            {
                //The custom item type is unknown, so use the alternate representation that the
                //report provides for renderers without support for this type.
                ProcessReportItemsType(altReportItem, component, page, dataset);
            }
            else
            {
                ThrowError(baseNode.Name, "Type", $"CustomReportItem type not supported: {type}");
            }

            container.Components.Add(component);
        }

        #region Barcode
        /// <summary>
        /// Makes the barcode of a custom report item of Bold Reports, barcode or matrixbarcode, with the value,
        /// the type and the text of its custom properties. Returns null for another item or an unknown type of the
        /// barcode; such a barcode keeps its place as an empty panel and is reported.
        /// </summary>
        private StiBarCode CreateBoldBarCode(XmlNode baseNode, string type, StiPanel place, string dataset)
        {
            if (!string.Equals(type, "barcode", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(type, "matrixbarcode", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var properties = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var propertiesNode = GetChildNode(baseNode, "CustomProperties");
            if (propertiesNode != null)
            {
                foreach (XmlNode node in propertiesNode.ChildNodes)
                {
                    var name = GetChildText(node, "Name");
                    if (!string.IsNullOrEmpty(name))
                        properties[name] = GetChildNode(node, "Value")?.InnerText ?? string.Empty;
                }
            }

            properties.TryGetValue("BarcodeType", out var barCodeTypeName);
            var barCodeType = GetBoldBarCodeType(barCodeTypeName);
            if (barCodeType == null)
            {
                ThrowError(null, null, $"Barcode type not supported: {barCodeTypeName}");
                return null;
            }

            properties.TryGetValue("BarcodeValue", out var value);
            properties.TryGetValue("DisplayBarcodeText", out var displayText);

            var barCode = new StiBarCode
            {
                Name = place.Name,
                Page = place.Page,
                ClientRectangle = place.ClientRectangle,
                BarCodeType = barCodeType,
                ShowLabelText = !string.Equals(displayText, "false", StringComparison.OrdinalIgnoreCase),
                Code = new StiBarCodeExpression(IsExpression(value) ? ConvertExpression(value, dataset) : value)
            };

            barCode.Enabled = place.Enabled;
            foreach (StiAppExpression expression in place.Expressions)
            {
                barCode.Expressions.Add(expression);
            }

            return barCode;
        }

        /// <summary>
        /// Makes the digital signature of a PDF of Bold Reports, PDFSignature, in the place of the item. Bold
        /// Reports signs with the certificate of its server, named in the custom properties; the certificate of a
        /// PDF of Stimulsoft is set in the export, so the name of the certificate is named in the list of errors.
        /// </summary>
        private StiPdfDigitalSignature CreateBoldPdfSignature(XmlNode baseNode, StiPanel place)
        {
            var signature = new StiPdfDigitalSignature(place.ClientRectangle)
            {
                Name = place.Name,
                Page = place.Page,
                Enabled = place.Enabled
            };

            foreach (StiAppExpression expression in place.Expressions)
            {
                signature.Expressions.Add(expression);
            }

            var certificate = GetChildNode(baseNode, "CustomProperties")?.ChildNodes.Cast<XmlNode>()
                .Where(n => GetChildText(n, "Name") == "CertificateFileName")
                .Select(n => GetChildText(n, "Value"))
                .FirstOrDefault();
            ThrowError(null, null,
                $"Certificate of a PDF signature not imported, set it in the PDF export: {certificate}");

            return signature;
        }

        /// <summary>
        /// Returns the barcode type of Stimulsoft for the barcode type of Bold Reports, or null.
        /// </summary>
        private static StiBarCodeTypeService GetBoldBarCodeType(string name)
        {
            switch ((name ?? string.Empty).Replace(" ", string.Empty).Replace("-", string.Empty).ToLowerInvariant())
            {
                case "code39":
                    return new StiCode39BarCodeType();

                case "code39extended":
                    return new StiCode39ExtBarCodeType();

                case "code93":
                    return new StiCode93BarCodeType();

                case "code93extended":
                    return new StiCode93ExtBarCodeType();

                case "code128":
                    return new StiCode128AutoBarCodeType();

                case "code128a":
                    return new StiCode128aBarCodeType();

                case "code128b":
                    return new StiCode128bBarCodeType();

                case "code128c":
                    return new StiCode128cBarCodeType();

                case "code11":
                    return new StiCode11BarCodeType();

                case "codabar":
                    return new StiCodabarBarCodeType();

                case "ean13":
                    return new StiEAN13BarCodeType();

                case "ean8":
                    return new StiEAN8BarCodeType();

                case "interleaved2of5":
                    return new StiInterleaved2of5BarCodeType();

                case "upcbarcode":
                case "upca":
                    return new StiUpcABarCodeType();

                case "msi":
                    return new StiMsiBarCodeType();

                case "qrbarcode":
                case "qrcode":
                    return new StiQRCodeBarCodeType();

                case "datamatrix":
                case "datamatrixbarcode":
                    return new StiDataMatrixBarCodeType();

                case "pdf417":
                case "pdf417barcode":
                    return new StiPdf417BarCodeType();

                default:
                    return null;
            }
        }
        #endregion

        #region BandedList
        //A BandedList declares its report items inside CustomData and the bands of
        //BandedListConfig reference them by name, so the items are collected first.
        private void ProcessBandedList(XmlNode customData, XmlNode bandedListConfig, StiContainer component, StiPage page, string dataset)
        {
            var items = new Dictionary<string, StiComponent>();
            var groupExpressions = new Dictionary<string, string>();

            if (customData != null)
                CollectBandedListData(customData, new StiContainer(), page, dataset, items, groupExpressions);

            ProcessBandedListConfigType(bandedListConfig, component, page, items, groupExpressions, dataset);
        }

        private void CollectBandedListData(XmlNode baseNode, StiContainer temp, StiPage page, string dataset,
            Dictionary<string, StiComponent> items, Dictionary<string, string> groupExpressions)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.LocalName)
                {
                    case "Grouping":
                        var groupName = node.Attributes?["Name"]?.Value;
                        var expression = GetBandedListGroupExpression(node, dataset);
                        if (!string.IsNullOrEmpty(groupName) && !string.IsNullOrWhiteSpace(expression))
                            groupExpressions[groupName] = expression;
                        break;

                    case "ReportItems":
                        CollectBandedListItems(node, temp, page, dataset, items);
                        break;

                    default:
                        CollectBandedListData(node, temp, page, dataset, items, groupExpressions);
                        break;
                }
            }
        }

        private string GetBandedListGroupExpression(XmlNode baseNode, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.LocalName != "GroupExpressions")
                    continue;

                foreach (XmlNode expressionNode in node.ChildNodes)
                {
                    if (expressionNode.LocalName != "GroupExpression")
                        continue;

                    var expression = ConvertExpression(GetNodeTextValue(expressionNode), dataset);
                    if (!string.IsNullOrWhiteSpace(expression))
                        return expression;      //todo use all expressions
                }
            }

            return null;
        }

        private void CollectBandedListItems(XmlNode baseNode, StiContainer temp, StiPage page, string dataset,
            Dictionary<string, StiComponent> items)
        {
            var startIndex = temp.Components.Count;
            ProcessReportItemsType(baseNode, temp, page, dataset);

            //The bands reference the items by the name from the report definition, which can differ
            //from the component name after the uniqueness check, so map them by the document order.
            var index = startIndex;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                var name = node.Attributes?["Name"]?.Value;
                if (string.IsNullOrEmpty(name))
                    continue;

                if (index >= temp.Components.Count)
                    break;

                items[name] = temp.Components[index++];
            }
        }

        private void ProcessBandedListConfigType(XmlNode baseNode, StiContainer component, StiPage page,
            Dictionary<string, StiComponent> items, Dictionary<string, string> groupExpressions, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.LocalName)
                {
                    case "Groups":
                        ProcessBandedListGroupsType(node, component, page, items, groupExpressions, dataset);
                        break;

                    case "Details":
                        var dataBand = new StiDataBand();
                        dataBand.Name = CheckComponentName(page.Report, component.Name + "DataBand");
                        dataBand.Page = page;
                        dataBand.CanGrow = true;
                        ProcessBandedListBandType(node, dataBand, items, dataset);
                        component.Components.Add(dataBand);
                        break;

                    case "FixedWidth":
                    case "FixedHeight":
                        //the size is taken from the custom report item itself
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessBandedListGroupsType(XmlNode baseNode, StiContainer component, StiPage page,
            Dictionary<string, StiComponent> items, Dictionary<string, string> groupExpressions, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.LocalName)
                {
                    case "Group":
                        ProcessBandedListGroupType(node, component, page, items, groupExpressions, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessBandedListGroupType(XmlNode baseNode, StiContainer component, StiPage page,
            Dictionary<string, StiComponent> items, Dictionary<string, string> groupExpressions, string dataset)
        {
            var groupName = baseNode.Attributes?["Name"]?.Value;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.LocalName)
                {
                    case "HeaderBand":
                        var groupHeader = new StiGroupHeaderBand();
                        groupHeader.Name = CheckComponentName(page.Report, component.Name + "GroupHeader");
                        groupHeader.Page = page;
                        groupHeader.CanGrow = true;

                        string expression;
                        if (!string.IsNullOrEmpty(groupName) && groupExpressions.TryGetValue(groupName, out expression))
                            groupHeader.Condition = new StiGroupConditionExpression(expression);

                        ProcessBandedListBandType(node, groupHeader, items, dataset);
                        component.Components.Add(groupHeader);
                        break;

                    case "FooterBand":
                        var groupFooter = new StiGroupFooterBand();
                        groupFooter.Name = CheckComponentName(page.Report, component.Name + "GroupFooter");
                        groupFooter.Page = page;
                        groupFooter.CanGrow = true;
                        ProcessBandedListBandType(node, groupFooter, items, dataset);
                        component.Components.Add(groupFooter);
                        break;

                    case "Visibility":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessBandedListBandType(XmlNode baseNode, StiBand band, Dictionary<string, StiComponent> items, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.LocalName)
                {
                    case "Height":
                        band.Height = ToHi(GetNodeTextValue(node));
                        break;

                    case "ReportItems":
                        ProcessBandedListReportItemsType(node, band, items);
                        break;

                    case "Visibility":
                        ProcessVisibility(node, band, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessBandedListReportItemsType(XmlNode baseNode, StiBand band, Dictionary<string, StiComponent> items)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.LocalName)
                {
                    case "ReportItemReference":
                        var name = node.Attributes?["Name"]?.Value;
                        StiComponent comp;
                        if (!string.IsNullOrEmpty(name) && items.TryGetValue(name, out comp))
                            band.Components.Add(comp);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }
        #endregion

        #endregion
    }
}
