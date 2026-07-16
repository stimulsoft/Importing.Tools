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
using Stimulsoft.Report.Chart;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.TextFormats;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Units;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
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
                        ProcessStyleType(node, component, dataset);
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

                    case "Visibility":
                        ProcessVisibility(node, component, dataset);
                        break;

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
    }
}
