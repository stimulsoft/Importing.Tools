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
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
        #region Tablix

        #region Classes
        private class TablixMember
        {
            public string GroupName = null;
            public List<string> GroupExpressions = null;
            public List<string> SortExpressions = null;
            public StiComponent Header = null;
            public List<TablixMember> Childs = null;
            public string KeepWithGroup;
            public XmlNode VisibilityNode;
            public bool ResetPageNumber = false;
            public string BreakLocation = null;
        }

        private class GroupInfo2
        {
            public List<int> HeaderLines = new List<int>();
            public List<int> FooterLines = new List<int>();
            public List<int> GroupHeaderLines = new List<int>();
            public List<int> GroupFooterLines = new List<int>();
            public List<int> DataLines = new List<int>();

            public TablixMember Member;
            public GroupInfo2 Parent;
        }

        private class Info2
        {
            public int RowIndex;
            public int ColumnIndex;
            public List<List<StiComponent>> Cells = new List<List<StiComponent>>();
            public StiPage Page;
            public List<List<GroupInfo2>> Groups;

            public List<List<StiComponent>> ColumnsCells = new List<List<StiComponent>>();
            public List<List<GroupInfo2>> ColumnsGroups;

            public List<double> TableColumnsWidths;
            public List<double> TableHeaderRowsHeights;
            public List<StiComponent> TableDetails;
            public List<StiComponent> TablixCorner;
            public string DataSetName;
            public string DataSetName2;
        }
        #endregion

        private void ProcessTablixType(XmlNode baseNode, StiContainer baseContainer, StiPage page, string dataset)
        {
            var container = new StiPanel();
            container.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"]?.Value);
            container.Border = new StiAdvancedBorder();
            container.Page = page;
            container.CanGrow = true;
            container.CanBreak = true;

            string datasetName = null;
            string fullDatasetName = null;
            StiFiltersCollection filters = null;
            string[] sortExpr = null;

            var rowsHierarchy2 = new List<TablixMember>();
            var columnsHierarchy2 = new List<TablixMember>();

            //preliminary pass - resolve the dataset name first. It is needed to convert
            //field references in the hierarchies and sort expressions below, which in RDL
            //come before the DataSetName node.
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                if (node.Name == "DataSetName")
                    datasetName = Convert.ToString(GetNodeTextValue(node));
            }

            fullDatasetName = MakeFullDataSetName(dataset, datasetName);

            //first pass - hierarchies and sort, resolved against the tablix own dataset
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixRowHierarchy":
                        ProcessTablixRowHierarchyType2(node, fullDatasetName, page, rowsHierarchy2);
                        break;

                    case "TablixColumnHierarchy":
                        ProcessTablixRowHierarchyType2(node, fullDatasetName, page, columnsHierarchy2); //todo columns
                        break;

                    case "SortExpressions":
                        sortExpr = ProcessTablixSortExpressionsType(node, fullDatasetName).ToArray();
                        break;
                }
            }

            var tableColumnsWidths = new List<double>();
            var tableDetails = new List<StiComponent>();
            List<StiComponent> tablixCorner = null;

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
                        ProcessStyleType(node, container, dataset);
                        break;

                    case "TablixBody":
                        ProcessTablixBodyType(node, page, fullDatasetName, tableColumnsWidths, tableDetails);
                        break;

                    case "Filters":
                        filters = ProcessFiltersType(node, fullDatasetName);
                        break;

                    case "TablixCorner":
                        tablixCorner = ProcessTablixCornerType(node, fullDatasetName, page);
                        break;

                    case "TablixRowHierarchy":
                    case "TablixColumnHierarchy":
                    case "SortExpressions":
                    case "RepeatRowHeaders":
                    case "RepeatColumnHeaders":
                    case "KeepTogether":
                    case "CustomProperties":
                    case "DataSetName":
                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            var info = AnalyzeHierarchy(rowsHierarchy2);
            info.TablixCorner = tablixCorner;
            info.TableColumnsWidths = tableColumnsWidths;
            info.TableDetails = tableDetails;
            info.Page = page;

            var info2 = AnalyzeHierarchy(columnsHierarchy2);
            info.ColumnsGroups = info2.Groups;
            info.ColumnsCells = info2.Cells;
            info.DataSetName = datasetName;
            info.DataSetName2 = datasetName;

            var hasRowColumns = !((info.Groups.Count == 1) && (info.Groups[0].Count == 1) && (info.Groups[0][0].HeaderLines.Count == tableDetails.Count));
            var hasColumnGroups = !((info.ColumnsGroups.Count == 1) && (info.ColumnsGroups[0].Count == 1) && (info.ColumnsGroups[0][0].HeaderLines.Count == info.ColumnsCells.Count));
            if (hasColumnGroups)
            {
                //calculate row height for header lines
                info.TableHeaderRowsHeights = new List<double>();
                for (int indexRow = 0; indexRow < info.ColumnsCells[0].Count; indexRow++)
                {
                    info.TableHeaderRowsHeights.Add(0);
                    for (int indexCol = 0; indexCol < info.ColumnsCells.Count; indexCol++)
                    {
                        var comp = info.ColumnsCells[indexCol][indexRow];
                        if (comp != null && comp.Width > 0)
                        {
                            if (info.TableHeaderRowsHeights[indexRow] < comp.Width)
                                info.TableHeaderRowsHeights[indexRow] = comp.Width;
                        }
                    }
                }

                if (hasRowColumns)
                {
                    string newDataSetName = MakeDataSourceCopy(datasetName);
                    info.DataSetName2 = newDataSetName;
                }
                else
                {
                    info.DataSetName2 = datasetName;
                }

                AddColumnsHeaders(info, container);
            }

            if (!hasRowColumns)
            {
                //only static rows
                var cont = new StiContainer() { Page = page };
                var dataBand = new StiDataBand();
                dataBand.Name = $"{container.Name}_Db";
                dataBand.CountData = 1;
                AddTablixCellsToContainer2(info, dataBand, info.Groups[0][0].HeaderLines, cont, null);
                if (hasColumnGroups)
                    container.Components.Add(dataBand);
                else
                    container.Components.AddRange(dataBand.Components);
            }
            else
            {
                //add bands
                //first level
                for (int index = 0; index < info.Groups.Count; index++)
                {
                    var groups = info.Groups[index];
                    string levelId = info.Groups.Count > 1 ? $"{index + 1}_" : "";

                    //second level
                    var firstGroup = groups[0];
                    if (firstGroup.HeaderLines.Count > 0)
                    {
                        var header = new StiHeaderBand() { PrintIfEmpty = true };
                        header.Name = $"{container.Name}_{levelId}Hd";
                        AddTablixCellsToContainer2(info, header, firstGroup.HeaderLines, container, firstGroup);
                    }

                    for (int groupIndex = 0; groupIndex < groups.Count - 1; groupIndex++)
                    {
                        var group = groups[groupIndex];
                        var groupHeader = new StiGroupHeaderBand();
                        groupHeader.Name = $"{container.Name}_{levelId}Gh{groupIndex + 1}";
                        groupHeader.Condition = new StiGroupConditionExpression("{" + group.Member.GroupExpressions[0] + "}");      //todo use all expressions
                        AddTablixCellsToContainer2(info, groupHeader, group.GroupHeaderLines, container, group);
                    }

                    var dataBand = new StiDataBand();
                    dataBand.Name = $"{container.Name}_{levelId}Db";
                    dataBand.DataSourceName = datasetName;
                    if (sortExpr != null && sortExpr.Length > 0)
                        dataBand.Sort = sortExpr;
                    if (filters != null) dataBand.Filters = filters;
                    AddTablixCellsToContainer2(info, dataBand, groups[groups.Count - 1].DataLines, container, groups[groups.Count - 1]);

                    bool needGroupFooters = false;
                    for (int groupIndex = 0; groupIndex < groups.Count - 1; groupIndex++)
                    {
                        if (groups[groupIndex].GroupFooterLines.Count > 0) needGroupFooters = true;
                    }
                    if (needGroupFooters)
                    {
                        for (int groupIndex = groups.Count - 2; groupIndex >= 0; groupIndex--)
                        {
                            var group = groups[groupIndex];
                            var groupFooter = new StiGroupFooterBand();
                            groupFooter.Name = $"{container.Name}_{levelId}Gf{groupIndex + 1}";
                            AddTablixCellsToContainer2(info, groupFooter, group.GroupFooterLines, container, group);
                        }
                    }

                    if (firstGroup.FooterLines.Count > 0)
                    {
                        var footer = new StiFooterBand() { PrintIfEmpty = true };
                        footer.Name = $"{container.Name}_{levelId}Ft";
                        AddTablixCellsToContainer2(info, footer, firstGroup.FooterLines, container, firstGroup);
                    }
                }
            }

            baseContainer.Components.Add(container);
        }

        private void AddTablixCellsToContainer2(Info2 info, StiBand baseBand, List<int> lines, StiContainer baseContainer, GroupInfo2 group)
        {
            baseBand.Height = 0;
            baseBand.Page = info.Page;
            baseContainer.Components.Add(baseBand);

            if (group?.Member != null)
            {
                if (group.Member.VisibilityNode != null)
                    ProcessVisibility(group.Member.VisibilityNode, baseBand, info.DataSetName);

                if (group.Member.ResetPageNumber)
                    baseBand.ResetPageNumber = true;

                if (!string.IsNullOrWhiteSpace(group.Member.BreakLocation))
                    ApplyBreakLocation(baseBand, group.Member.BreakLocation);
            }

            if (lines.Count == 0) return;

            double offsetY = 0;
            var hasColumnGroups = !((info.ColumnsGroups.Count == 1) && (info.ColumnsGroups[0].Count == 1) && (info.ColumnsGroups[0][0].HeaderLines.Count == info.ColumnsCells.Count));

            for (int indexLine = 0; indexLine < lines.Count; indexLine++)
            {
                int lineIndex = lines[indexLine];
                double offset = 0;

                var newCont = new StiContainer();
                if (lines.Count > 1)
                    newCont.Name = baseBand.Name + "_c" + indexLine.ToString();
                else
                    newCont.Name = baseBand.Name;

                for (int index = 0; index < info.Cells[lineIndex].Count; index++)
                {
                    var comp = info.Cells[lineIndex][index];
                    comp.Left = offset;
                    offset += comp.Width;
                    newCont.Components.Add(comp);
                }

                var container = info.TableDetails[lineIndex] as StiPanel;
                var rowHeight = container.Height;

                int indexComp = 0;
                for (int index = 0; index < info.TableColumnsWidths.Count && indexComp < container.Components.Count; index++)
                {
                    var comp = container.Components[indexComp++];

                    int colspan = 0;
                    if (comp.TagValue != null && comp.TagValue is int tagInt)
                    {
                        colspan = Convert.ToInt32(tagInt) - 1;
                    }
                    var cellWidth = (decimal)info.TableColumnsWidths[index];
                    while (colspan > 0 && index + 1 < info.TableColumnsWidths.Count)
                    {
                        index++;
                        cellWidth += (decimal)info.TableColumnsWidths[index];
                        colspan--;
                    }

                    comp.Left = offset;
                    comp.Width = (double)cellWidth;
                    offset += (double)cellWidth;

                    newCont.Components.Add(comp);
                }                               

                for (int index = 0; index < newCont.Components.Count; index++)
                {
                    var comp = newCont.Components[index];
                    comp.GrowToHeight = true;
                    comp.Height = rowHeight;
                }

                if (hasColumnGroups)
                {
                    info.RowIndex = lineIndex;
                    ApplyColumnsGrouping(info, newCont);
                }

                if (lines.Count > 1)
                {
                    newCont.Name = baseBand.Name + "_c" + indexLine.ToString();
                    newCont.Width = baseContainer.Width;
                    newCont.Height = rowHeight;
                    newCont.Top = offsetY;
                    offsetY += rowHeight;
                    baseBand.Components.Add(newCont);
                }
                else
                {
                    baseBand.Components.AddRange(newCont.Components);
                }
                baseBand.Height += container.Height;
            }
        }

        #region Process columns
        private void ApplyColumnsGrouping(Info2 info, StiContainer baseCont)
        {
            var newCont = new StiContainer();

            var headerCellCount = info.Cells[info.RowIndex].Count;
            if (headerCellCount > 0)
            {
                var list = new List<int>();
                for (int index = 0; index < headerCellCount; index++)
                {
                    list.Add(index - headerCellCount);
                }
                var header = new StiCrossHeaderBand() { PrintIfEmpty = true };
                header.Name = $"{baseCont.Name}_CHdc";
                AddColumnCellsToContainer2(info, header, list, baseCont, newCont);
            }

            for (int index = 0; index < info.ColumnsGroups.Count; index++)
            {
                var groups = info.ColumnsGroups[index];
                string levelId = info.ColumnsGroups.Count > 1 ? $"{index + 1}_" : "";

                //second level
                var firstGroup = groups[0];
                if (firstGroup.HeaderLines.Count > 0)
                {
                    var header = new StiCrossHeaderBand() { PrintIfEmpty = true };
                    header.Name = $"{baseCont.Name}_{levelId}CHd";
                    AddColumnCellsToContainer2(info, header, firstGroup.HeaderLines, baseCont, newCont);
                }

                for (int groupIndex = 0; groupIndex < groups.Count - 1; groupIndex++)
                {
                    var group = groups[groupIndex];
                    var groupHeader = new StiCrossGroupHeaderBand();
                    groupHeader.Name = $"{baseCont.Name}_{levelId}CGh{groupIndex + 1}";
                    var groupExpr = group.Member.GroupExpressions[0];     //todo use all expressions
                    groupExpr = groupExpr.Replace(info.DataSetName, info.DataSetName2);
                    groupHeader.Condition = new StiGroupConditionExpression("{" + groupExpr + "}");
                    AddColumnCellsToContainer2(info, groupHeader, group.GroupHeaderLines, baseCont, newCont);
                }

                var dataBand = new StiCrossDataBand();
                dataBand.DataSourceName = info.DataSetName2;
                dataBand.Name = $"{baseCont.Name}_{levelId}CDb";
                /*if (sortExpr != null && sortExpr.Length > 0)
                    dataBand.Sort = sortExpr;
                if (filters != null) dataBand.Filters = filters; */     //todo ???
                AddColumnCellsToContainer2(info, dataBand, groups[groups.Count - 1].DataLines, baseCont, newCont);

                bool needGroupFooters = false;
                for (int groupIndex = 0; groupIndex < groups.Count - 1; groupIndex++)
                {
                    if (groups[groupIndex].GroupFooterLines.Count > 0) needGroupFooters = true;
                }
                if (needGroupFooters)
                {
                    for (int groupIndex = groups.Count - 2; groupIndex >= 0; groupIndex--)
                    {
                        var group = groups[groupIndex];
                        var groupFooter = new StiCrossGroupFooterBand();
                        groupFooter.Name = $"{baseCont.Name}_{levelId}CGf{groupIndex + 1}";
                        AddColumnCellsToContainer2(info, groupFooter, group.GroupFooterLines, baseCont, newCont);
                    }
                }

                if (firstGroup.FooterLines.Count > 0)
                {
                    var footer = new StiCrossFooterBand() { PrintIfEmpty = true };
                    footer.Name = $"{baseCont.Name}_{levelId}CFt";
                    AddColumnCellsToContainer2(info, footer, firstGroup.FooterLines, baseCont, newCont);
                }
            }

            baseCont.Components.Clear();
            baseCont.Components.AddRange(newCont.Components);
        }

        private void AddColumnCellsToContainer2(Info2 info, StiBand baseBand, List<int> lines, StiContainer baseContainer, StiContainer newContainer)
        {
            baseBand.Page = info.Page;

            var headerCellCount = info.Cells[info.RowIndex].Count;

            double rowHeight = 0;
            double sumWidth = 0;
            for (int index = 0; index < lines.Count; index++)
            {
                var indexLine = headerCellCount + lines[index];

                var comp = baseContainer.Components[indexLine];
                comp.Left = sumWidth;
                sumWidth += comp.Width;
                rowHeight = comp.Height;
                baseBand.Components.Add(comp);
            }
            baseBand.Width = sumWidth;
            //baseBand.Height = rowHeight;  ?? todo?

            newContainer.Components.Add(baseBand);
        }

        private void AddColumnsHeaders(Info2 info, StiContainer baseContainer)
        {
            var baseBand = new StiHeaderBand();
            baseBand.Page = info.Page;
            baseBand.PrintIfEmpty = true;
            baseBand.Name = $"{baseContainer.Name}_Hdc";
            baseContainer.Components.Add(baseBand);

            double bandHeight = 0;
            foreach (var value in info.TableHeaderRowsHeights)
                bandHeight += value;
            baseBand.Height += bandHeight;

            if (info.TablixCorner != null)
            {
                var headerCellCount = info.TablixCorner.Count;
                if (headerCellCount > 0)
                {
                    var header = new StiCrossHeaderBand() { PrintIfEmpty = true };
                    header.Name = $"{baseBand.Name}_CHdc";
                    AddColumnHeaderToContainer2(info, header, info.TablixCorner, baseBand);
                }
            }

            for (int index = 0; index < info.ColumnsGroups.Count; index++)
            {
                var groups = info.ColumnsGroups[index];
                string levelId = info.ColumnsGroups.Count > 1 ? $"{index + 1}_" : "";

                //second level
                var firstGroup = groups[0];
                if (firstGroup.HeaderLines.Count > 0)
                {
                    var header = new StiCrossHeaderBand() { PrintIfEmpty = true };
                    header.Name = $"{baseBand.Name}_{levelId}CHd";
                    AddColumnHeaderToContainer3(info, header, firstGroup.HeaderLines, baseBand);
                }

                for (int groupIndex = 0; groupIndex < groups.Count - 1; groupIndex++)
                {
                    var group = groups[groupIndex];
                    var groupHeader = new StiCrossGroupHeaderBand();
                    groupHeader.Name = $"{baseBand.Name}_{levelId}CGh{groupIndex + 1}";
                    var groupExpr = group.Member.GroupExpressions[0];     //todo use all expressions
                    groupExpr = groupExpr.Replace(info.DataSetName, info.DataSetName2);
                    groupHeader.Condition = new StiGroupConditionExpression("{" + groupExpr + "}");
                    AddColumnHeaderToContainer3(info, groupHeader, group.GroupHeaderLines, baseBand);
                }

                var dataBand = new StiCrossDataBand();
                dataBand.DataSourceName = info.DataSetName2;
                dataBand.Name = $"{baseBand.Name}_{levelId}CDb";
                /*if (sortExpr != null && sortExpr.Length > 0)
                    dataBand.Sort = sortExpr;
                if (filters != null) dataBand.Filters = filters; */     //todo ???
                AddColumnHeaderToContainer3(info, dataBand, groups[groups.Count - 1].DataLines, baseBand);

                bool needGroupFooters = false;
                for (int groupIndex = 0; groupIndex < groups.Count - 1; groupIndex++)
                {
                    if (groups[groupIndex].GroupFooterLines.Count > 0) needGroupFooters = true;
                }
                if (needGroupFooters)
                {
                    for (int groupIndex = groups.Count - 2; groupIndex >= 0; groupIndex--)
                    {
                        var group = groups[groupIndex];
                        var groupFooter = new StiCrossGroupFooterBand();
                        groupFooter.Name = $"{baseBand.Name}_{levelId}CGf{groupIndex + 1}";
                        AddColumnHeaderToContainer3(info, groupFooter, group.GroupFooterLines, baseBand);
                    }
                }

                if (firstGroup.FooterLines.Count > 0)
                {
                    var footer = new StiCrossFooterBand() { PrintIfEmpty = true };
                    footer.Name = $"{baseBand.Name}_{levelId}CFt";
                    AddColumnHeaderToContainer3(info, footer, firstGroup.FooterLines, baseBand);
                }
            }
        }

        private void AddColumnHeaderToContainer2(Info2 info, StiBand baseBand, List<StiComponent> cells, StiBand baseContainer)
        {
            baseBand.Page = info.Page;
            baseContainer.Components.Add(baseBand);

            var headerCellCount = info.Cells[info.RowIndex].Count;
            var widths = new double[headerCellCount];
            var sumWidths = 0d;
            for (int index = 0; index < headerCellCount; index++)
            {
                var comp = info.Cells[0][index];
                if (comp != null)
                    widths[index] = comp.Width;
                sumWidths += widths[index];
            }

            double offset = 0;
            for (int index = 0; index < cells.Count; index++)
            {
                var comp = cells[index];

                if (comp is StiContainer cont)
                {
                    double offsetX = 0;
                    for (int indexColumn = 0; indexColumn < cont.Components.Count; indexColumn++)
                    {
                        var comp2 = cont.Components[indexColumn];
                        comp2.Left = offsetX;
                        comp2.Width = widths[indexColumn];
                        comp2.Height = info.TableHeaderRowsHeights[index];
                        comp2.GrowToHeight = true;
                        offsetX += comp2.Width;
                    }
                }

                comp.Width = sumWidths;
                comp.Height = info.TableHeaderRowsHeights[index];
                comp.Top = offset;
                comp.GrowToHeight = true;
                offset += comp.Height;

                if (comp is StiContainer cont2 && cont2.Components.Count == 1)
                {
                    var comp3 = cont2.Components[0];
                    comp3.Top = cont2.Top;
                    baseBand.Components.Add(comp3);
                }
                else
                {
                    if (comp.Height > 0)
                        baseBand.Components.Add(comp);
                }
            }
            baseBand.Width = sumWidths;
            baseBand.Height = offset;
        }

        private void AddColumnHeaderToContainer3(Info2 info, StiBand baseBand, List<int> lines, StiBand baseContainer)
        {
            baseBand.Page = info.Page;
            baseContainer.Components.Add(baseBand);

            if (lines.Count == 0) return;

            double offsetY = 0;
            for (int indexRow = 0; indexRow < info.ColumnsCells[0].Count; indexRow++)
            {
                var cont = new StiContainer();
                cont.Top = offsetY;
                cont.Name = $"HeadCont{indexRow}";

                double offset = 0;
                for (int index = 0; index < lines.Count; index++)
                {
                    var indexLine = lines[index];

                    var compWidth = info.TableColumnsWidths[indexLine];

                    var comp = info.ColumnsCells[indexLine][indexRow];
                    if (comp != null)
                    {
                        comp.Height = info.TableHeaderRowsHeights[indexRow];
                        comp.Left = offset;
                        comp.Width = compWidth;
                        cont.Components.Add(comp);
                    }
                    offset += compWidth;
                }
                cont.Width = offset;
                cont.Height = info.TableHeaderRowsHeights[indexRow];

                if (cont.Components.Count == 1)
                {
                    var comp3 = cont.Components[0];
                    comp3.Top = cont.Top;
                    baseBand.Components.Add(comp3);
                }
                else
                    baseBand.Components.Add(cont);

                offsetY += cont.Height;
            }

            if (baseBand.Components.Count > 0)
                baseBand.Width = baseBand.Components[0].Width;
            baseBand.Height = offsetY;
        }
        #endregion

        #region Process body
        private void ProcessTablixBodyType(XmlNode baseNode, StiPage page, string dataset, List<double> tableColumns, List<StiComponent> tableDetails)
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
                        var tableDetails2 = ProcessTablixRowsType(node, page, dataset);
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

        private List<StiComponent> ProcessTablixRowsType(XmlNode baseNode, StiPage page, string dataset)
        {
            var list = new List<StiComponent>();
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
            var container = new StiPanel();

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

                    case "DataElementName":
                    case "DataElementOutput":
                        //ignored or not implemented yet
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
        #endregion

        #region Analyze hierarchy
        private Info2 AnalyzeHierarchy(List<TablixMember> hierarchy)
        {
            var info = new Info2();
            //info.Hierarchy = hierarchy;
            info.RowIndex = 0;
            info.Groups = new List<List<GroupInfo2>>();
            info.Groups.Add(new List<GroupInfo2>());

            var group = new GroupInfo2();
            info.Groups[0].Add(group);

            var tempMember = new TablixMember();
            tempMember.Childs = hierarchy;

            AnalyzeGroupRecursive(tempMember, info, group, false, true);

            //remove empty cells
            int maxLen = 0;
            foreach (var list in info.Cells)
            {
                var index = list.Count - 1;
                while (index >= 0 && list[index] == null) index--;
                index++;
                if (maxLen < index) maxLen = index;
            }
            foreach (var list in info.Cells)
            {
                if (list.Count > maxLen)
                    list.RemoveRange(maxLen, list.Count - maxLen);
            }

            return info;
        }

        private void AnalyzeGroupRecursive(TablixMember baseMember, Info2 info, GroupInfo2 baseGroup, bool baseHasGroup, bool isRoot = false)
        {
            if (baseMember.Childs == null || baseMember.Childs.Count == 0) return;

            var storedColumnIndex = info.ColumnIndex;

            bool wasGroup = false;
            for (int index = 0; index < baseMember.Childs.Count; index++)
            {
                var member = baseMember.Childs[index];

                bool hasGroup = member.GroupExpressions != null && member.GroupExpressions.Count > 0;

                if (isRoot && wasGroup && (hasGroup || member.KeepWithGroup == "after"))
                {
                    //start new group
                    info.Groups.Add(new List<GroupInfo2>());
                    baseGroup = new GroupInfo2();
                    info.Groups[info.Groups.Count - 1].Add(baseGroup);
                    wasGroup = false;
                }

                if (isRoot)
                    info.ColumnIndex = 0;
                else
                    info.ColumnIndex = storedColumnIndex;

                //add header to cells
                var rowHeight = ScanMembersHeightRecursive(member);
                for (int index2 = 0; index2 < rowHeight; index2++)
                {
                    var newComp = member.Header?.Clone() as StiComponent;
                    if (newComp != null && index2 > 0)
                        newComp.Name += "_" + index2.ToString();

                    if (newComp is StiText stiText)
                        stiText.ProcessingDuplicates = StiProcessingDuplicatesType.GlobalMerge;
                    if (newComp is StiImage stiImage)
                        stiImage.ProcessingDuplicates = StiImageProcessingDuplicatesType.GlobalMerge;

                    var rowIndex2 = info.RowIndex + index2; 
                    while (info.Cells.Count < rowIndex2 + 1)
                        info.Cells.Add(new List<StiComponent>());
                    while (info.Cells[rowIndex2].Count < info.ColumnIndex + 1)
                        info.Cells[rowIndex2].Add(null);
                    info.Cells[rowIndex2][info.ColumnIndex] = newComp;
                }

                GroupInfo2 group2 = null;
              
                //check row type
                if (baseHasGroup || isRoot)
                {
                    if (member.GroupName == null)
                    {
                        bool hasNestedGroup = ScanMembersGroupRecursive(member);
                        if (!hasNestedGroup)
                        {
                            if (isRoot)
                            {
                                if (!wasGroup)
                                    baseGroup.HeaderLines.Add(info.RowIndex);
                                else
                                    baseGroup.FooterLines.Add(info.RowIndex);
                            }
                            else
                            {
                                if (baseGroup.Parent != null)
                                {
                                    if (!wasGroup)
                                        baseGroup.Parent.GroupHeaderLines.Add(info.RowIndex);
                                    else
                                        baseGroup.Parent.GroupFooterLines.Add(info.RowIndex);
                                }
                                else
                                {
                                    if (!wasGroup)
                                        baseGroup.GroupHeaderLines.Add(info.RowIndex);
                                    else
                                        baseGroup.GroupFooterLines.Add(info.RowIndex);
                                }
                            }
                        }
                        else
                        {
                            hasGroup = baseHasGroup;
                            wasGroup = true;
                        }
                    }
                    else
                    {
                        wasGroup = true;
                        if (hasGroup)
                        {
                            //nested group
                            baseGroup.Member = member;

                            group2 = new GroupInfo2();
                            group2.Parent = baseGroup;
                            info.Groups[info.Groups.Count - 1].Add(group2);

                            //check for group in last member - need to add empty member to easy processing
                            if (member.Childs == null ||  member.Childs.Count == 0)
                            {
                                var newMember = new TablixMember();
                                member.Childs = new List<TablixMember>();
                                member.Childs.Add(newMember);
                            }
                        }
                        else
                        {
                            baseGroup.Member = member;
                            for (int i = 0; i < rowHeight; i++)
                            {
                                baseGroup.DataLines.Add(info.RowIndex + i);
                            }
                        }
                    }
                }

                info.ColumnIndex++;
                AnalyzeGroupRecursive(member, info, group2 ?? baseGroup, hasGroup);

                info.RowIndex++;
            }
            info.RowIndex--;
        }

        private int ScanMembersHeightRecursive(TablixMember member)
        {
            if (member.Childs == null || member.Childs.Count == 0) return 1;
            int len = 0;
            for (int index = 0; index < member.Childs.Count; index++)
            {
                var row = member.Childs[index];
                len += ScanMembersHeightRecursive(row);
            }
            return len;
        }

        private bool ScanMembersGroupRecursive(TablixMember member)
        {
            if (member.GroupName != null) return true;
            if (member.Childs == null || member.Childs.Count == 0) return false;
            for (int index = 0; index < member.Childs.Count; index++)
            {
                var row = member.Childs[index];
                if (ScanMembersGroupRecursive(row)) return true;
            }
            return false;
        }
        #endregion

        #region Process row hierarchy
        private void ProcessTablixRowHierarchyType2(XmlNode baseNode, string dataSet, StiPage page, List<TablixMember> rowsHierarchy)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixMembers":
                        ProcessTablixMembersType2(node, dataSet, page, rowsHierarchy);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessTablixMembersType2(XmlNode baseNode, string dataSet, StiPage page, List<TablixMember> childs)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixMember":
                        var item = ProcessTablixMemberType2(node, dataSet, page);
                        childs.Add(item);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private TablixMember ProcessTablixMemberType2(XmlNode baseNode, string dataSet, StiPage page)
        {
            var tablixMember = new TablixMember();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Group":
                        ProcessTablixGroupType2(node, dataSet, tablixMember);
                        break;

                    case "SortExpressions":
                        tablixMember.SortExpressions = ProcessTablixSortExpressionsType(node, dataSet);
                        break;

                    case "TablixHeader":
                        tablixMember.Header = ProcessTablixHeaderType2(node, dataSet, page);
                        break;

                    case "TablixMembers":
                        tablixMember.Childs = new List<TablixMember>();
                        ProcessTablixMembersType2(node, dataSet, page, tablixMember.Childs);
                        break;

                    case "KeepWithGroup":
                        tablixMember.KeepWithGroup = GetNodeTextValue(node)?.ToLowerInvariant();
                        break;

                    case "Visibility":
                        var tempText = new StiText();
                        tablixMember.VisibilityNode = ProcessVisibility(node, tempText, dataSet);
                        break;

                    //case "KeepTogether":
                    //case "RepeatOnNewPage":
                    //case "FixedData":
                    case "DataElementName":
                    case "DataElementOutput":
                    case "HideIfNoRows":    //not applicable
                        //ignored in this revision
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return tablixMember;
        }

        private void ProcessTablixGroupType2(XmlNode baseNode, string dataSet, TablixMember tablixMember)
        {
            tablixMember.GroupName = baseNode.Attributes["Name"]?.Value;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupExpressions":
                        ProcessGroupExpressionsType2(node, dataSet, tablixMember);
                        break;

                    case "PageBreak":
                        ProcessPageBreakType2(node, dataSet, tablixMember);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessPageBreakType2(XmlNode baseNode, string dataSet, TablixMember tablixMember)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "BreakLocation":
                        tablixMember.BreakLocation = GetNodeTextValue(node);
                        break;

                    case "ResetPageNumber":
                        if (IsTrue(GetNodeTextValue(node)))
                            tablixMember.ResetPageNumber = true;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessGroupExpressionsType2(XmlNode baseNode, string dataSet, TablixMember tablixMember)
        {
            tablixMember.GroupExpressions = new List<string>();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "GroupExpression":
                        var st = ConvertExpression(GetNodeTextValue(node), dataSet);
                        if (st.StartsWith("{") && st.EndsWith("}"))
                        {
                            st = st.Substring(1, st.Length - 2);
                        }
                        if (!string.IsNullOrWhiteSpace(st))
                            tablixMember.GroupExpressions.Add(st);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private StiComponent ProcessTablixHeaderType2(XmlNode baseNode, string dataSet, StiPage page)
        {
            var cont = new StiContainer();
            double size = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Size":
                        size = ToHi(GetNodeTextValue(node));
                        break;

                    case "CellContents":
                        cont.Width = size;
                        ProcessReportItemsType(node, cont, page, dataSet);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            if (cont.Components.Count > 1)
                return cont;
            if (cont.Components.Count == 1)
            {
                var comp = cont.Components[0];
                comp.Width = cont.Width;
                return comp;
            }
            return null;
        }
        #endregion

        #region Process corner
        private List<StiComponent> ProcessTablixCornerType(XmlNode baseNode, string dataSet, StiPage page)
        {
            var list = new List<StiComponent>();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixCornerRows":
                        ProcessTablixCornerRowsType(node, dataSet, page, list);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return list;
        }

        private void ProcessTablixCornerRowsType(XmlNode baseNode, string dataSet, StiPage page, List<StiComponent> list)
        {
            int index = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixCornerRow":
                        var cont = ProcessTablixCornerRowType(node, dataSet, page, index++);
                        list.Add(cont);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private StiContainer ProcessTablixCornerRowType(XmlNode baseNode, string dataSet, StiPage page, int rowIndex)
        {
            var container = new StiContainer();
            container.Name = $"CornerRow{rowIndex}";

            int index = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "TablixCornerCell":
                        ProcessTablixCornerCellType(node, dataSet, page, container, index++);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return container;
        }

        private void ProcessTablixCornerCellType(XmlNode baseNode, string dataSet, StiPage page, StiContainer baseCont, int cellIndex)
        {
            var cont = new StiContainer();
            cont.Name = $"{baseCont.Name}_{cellIndex}";

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "CellContents":
                        ProcessReportItemsType(node, cont, page, dataSet);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            if (cont.Components.Count > 1)
                baseCont.Components.Add(cont);
            if (cont.Components.Count == 1)
                baseCont.Components.Add(cont.Components[0]);
        }
        #endregion

		private List<string> ProcessTablixSortExpressionsType(XmlNode baseNode, string dataSet)
		{
            var sort = new List<string>();
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
            return sort;
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
		#endregion
    }
}
