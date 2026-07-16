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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
        #region FixedPage
        private void ProcessFixedPageType(XmlNode baseNode, StiContainer container, StiPage page)
        {
            //make main container
            var component = new StiPanel();
            component.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"]?.Value);
            component.Page = page;
            component.CanGrow = true;
            component.CanBreak = true;

            var size = GetFixedSize(baseNode);
            if (size.Width > 0)
                component.Width = size.Width;
            else
                component.Width = page.Width;
            if (size.Height > 0)
                component.Height = size.Height;
            else
                component.Height = container.Height;

            //first pass - resolve the dataset name, it is needed to convert the grouping expression.
            string datasetName = null;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == "DataSetName")
                    datasetName = Convert.ToString(GetNodeTextValue(node));
            }
            var fullDatasetName = MakeFullDataSetName(null, datasetName);

            var dataSetNameCopy = MakeDataSourceCopy(datasetName, "Fixed");
            //var fullDatasetNameCopy = MakeFullDataSetName(null, dataSetNameCopy);

            //make bands
            var groupHeader = new StiGroupHeaderBand();
            groupHeader.Name = CheckComponentName(page.Report, component.Name + "GroupHeader");
            groupHeader.Page = page;

            var dataBand = new StiDataBand();
            dataBand.Name = CheckComponentName(page.Report, component.Name + "DataBand");
            dataBand.Page = page;
            dataBand.CanGrow = true;
            dataBand.CanShrink = true;
            dataBand.DataSourceName = dataSetNameCopy;
            dataBand.NewPageBefore = true;

            //second pass - the page content goes onto the data band, the grouping onto the
            //group header that is placed before it.
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Pages":
                        ProcessFixedPagePagesType(node, dataBand, page, fullDatasetName);
                        break;

                    case "Grouping":
                        ProcessFixedPageGroupingType(node, fullDatasetName, groupHeader);
                        break;

                    case "DataSetName":
                        //already processed
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            var maxHeight = 0d;
            foreach (StiComponent comp in dataBand.Components)
            {
                if (comp.Bottom > maxHeight)
                    maxHeight = comp.Bottom;
            }
            dataBand.Height = maxHeight + dataBand.HeaderSize;

            component.Components.Add(groupHeader);
            component.Components.Add(dataBand);
            container.Components.Add(component);
        }

        private void ProcessFixedPagePagesType(XmlNode baseNode, StiContainer band, StiPage page, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Page":
                        ProcessFixedPagePageType(node, band, page, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessFixedPagePageType(XmlNode baseNode, StiContainer band, StiPage page, string dataset)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ReportItems":
                        ProcessReportItemsType(node, band, page, dataset);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessFixedPageGroupingType(XmlNode baseNode, string dataset, StiGroupHeaderBand groupHeader)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupExpressions":
                        ProcessFixedPageGroupExpressionsType(node, dataset, groupHeader);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessFixedPageGroupExpressionsType(XmlNode baseNode, string dataset, StiGroupHeaderBand groupHeader)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupExpression":
                        var expression = ConvertExpression(GetNodeTextValue(node), dataset);
                        if (!string.IsNullOrWhiteSpace(expression))
                            groupHeader.Condition = new StiGroupConditionExpression(expression);      //todo use all expressions
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        //Recursively walks the nested nodes and collects the dd:FixedWidth/dd:FixedHeight
        //extension values (Data Dynamics). Returns the largest fixed size found, in report units.
        private SizeD GetFixedSize(XmlNode baseNode)
        {
            var size = new SizeD();
            GetFixedSizeRecursive(baseNode, ref size);
            return size;
        }

        private void GetFixedSizeRecursive(XmlNode baseNode, ref SizeD size)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.LocalName)
                {
                    case "FixedWidth":
                        var width = ToHi(GetNodeTextValue(node));
                        if (width > size.Width)
                            size.Width = width;
                        break;

                    case "FixedHeight":
                        var height = ToHi(GetNodeTextValue(node));
                        if (height > size.Height)
                            size.Height = height;
                        break;

                    default:
                        GetFixedSizeRecursive(node, ref size);
                        break;
                }
            }
        }
        #endregion
    }
}
