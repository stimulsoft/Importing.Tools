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
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
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
            public bool RepeatOnNewPage;
            public bool KeepTogether;
            public string DocumentMapLabel;
            //the filters of the details, a group of the rows without group expressions
            public XmlNode DetailsFilters;
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

            //RepeatOnNewPage of the static members of the lines
            public bool RepeatHeader;
            public bool RepeatFooter;
            public bool RepeatGroupHeader;
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

            //the visibility of each column of the body, null when no column can be hidden
            public List<XmlNode> ColumnVisibilities;

            //the visibility of each static row of the body, null when no row can be hidden
            public List<XmlNode> RowVisibilities;

            //the group that collapses the rows of the band being made when they are toggled
            public StiGroupHeaderBand ToggleGroup;
        }
        #endregion

        private void ProcessTablixType(XmlNode baseNode, StiContainer baseContainer, StiPage page, string dataset)
        {
            var wasInDataRegion = isInDataRegion;
            isInDataRegion = true;

            var container = new StiPanel();
            container.Name = CheckComponentName(page.Report, baseNode.Attributes["Name"]?.Value);
            container.Border = new StiAdvancedBorder();
            container.Page = page;
            container.CanGrow = true;
            container.CanBreak = true;

            string datasetName = null;
            string fullDatasetName = null;
            var repeatColumnHeaders = false;
            string noRowsMessage = null;
            Tuple<string, string, int> rowLimit = null;
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
                        rowLimit = GetTablixRowLimit(node, fullDatasetName);
                        filters = ProcessFiltersType(node, fullDatasetName, rowLimitTaken: rowLimit != null);
                        break;

                    case "TablixCorner":
                        tablixCorner = ProcessTablixCornerType(node, fullDatasetName, page);
                        break;

                    case "RepeatColumnHeaders":
                        repeatColumnHeaders = IsTrue(GetNodeTextValue(node));
                        break;

                    case "PageBreak":
                        ProcessItemPageBreak(node, container);
                        break;

                    case "NoRowsMessage":
                        noRowsMessage = GetNodeTextValue(node);
                        break;

                    case "Visibility":
                        ProcessVisibility(node, container, dataset);
                        break;

                    //the headers stay in view when the viewer scrolls the region, the viewer has no such headers
                    case "FixedRowHeaders":
                    case "FixedColumnHeaders":
                        break;

                    case "TablixRowHierarchy":
                    case "TablixColumnHierarchy":
                    case "SortExpressions":
                    case "RepeatRowHeaders":
                    case "KeepTogether":
                    case "CustomProperties":
                    case "DataSetName":
                    case "ZIndex":
                        //ignored or not implemented yet
                        break;

                    default:
                        if (!ProcessInteractivityNode(node, container, dataset))
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

            info.RowVisibilities = GetRowVisibilities(rowsHierarchy2, tableDetails.Count);
            info.ColumnVisibilities = GetColumnVisibilities(columnsHierarchy2, tableColumnsWidths.Count);
            if (info.ColumnVisibilities != null && hasColumnGroups)
            {
                ThrowError(null, null, $"Visibility of columns of a matrix not applied: {container.Name}");
                info.ColumnVisibilities = null;
            }
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
                StiDataBand firstDataBand = null;
                var countScopes = new Dictionary<StiBand, StiBand>();
                var rowScopes = new Dictionary<StiBand, List<string>>();
                var groupBands = new Dictionary<string, StiBand>(StringComparer.OrdinalIgnoreCase);

                //add bands
                //first level
                for (int index = 0; index < info.Groups.Count; index++)
                {
                    var groups = info.Groups[index];
                    string levelId = info.Groups.Count > 1 ? $"{index + 1}_" : "";

                    //second level
                    var firstGroup = groups[0];
                    StiHeaderBand header = null;
                    if (firstGroup.HeaderLines.Count > 0)
                    {
                        //Reporting Services repeats a header only by RepeatOnNewPage, a header band of
                        //Stimulsoft is printed on all pages by default
                        //without rows Reporting Services shows the message instead of the whole region
                        header = new StiHeaderBand() { PrintIfEmpty = string.IsNullOrEmpty(noRowsMessage) };
                        header.Name = $"{container.Name}_{levelId}Hd";
                        header.PrintOnAllPages = firstGroup.RepeatHeader || repeatColumnHeaders;
                        AddTablixCellsToContainer2(info, header, firstGroup.HeaderLines, container, firstGroup);
                    }

                    var sort = new List<string>();
                    var groupHeaders = new List<StiGroupHeaderBand>();
                    var groupConditions = new List<string>();
                    for (int groupIndex = 0; groupIndex < groups.Count - 1; groupIndex++)
                    {
                        var group = groups[groupIndex];
                        var groupHeader = new StiGroupHeaderBand();
                        groupHeader.Name = $"{container.Name}_{levelId}Gh{groupIndex + 1}";
                        groupConditions.Add(GetGroupCondition(group.Member.GroupExpressions));
                        groupHeader.Condition = new StiGroupConditionExpression("{" + groupConditions.Last() + "}");
                        AddGroupSort(sort, groupHeader, group.Member);
                        groupHeader.PrintOnAllPages = group.RepeatGroupHeader;
                        groupHeader.KeepGroupTogether = group.Member.KeepTogether;

                        if (group.Member.DocumentMapLabel != null)
                            groupHeader.Bookmark = new StiBookmarkExpression(group.Member.DocumentMapLabel);

                        info.ToggleGroup = groupHeaders.LastOrDefault();
                        AddTablixCellsToContainer2(info, groupHeader, group.GroupHeaderLines, container, group);
                        info.ToggleGroup = null;
                        groupHeaders.Add(groupHeader);
                        countScopes[groupHeader] = groupHeader;
                        rowScopes[groupHeader] = groupConditions.ToList();

                        if (group.Member.GroupName != null)
                            groupBands[group.Member.GroupName] = groupHeader;
                    }

                    //the rows are sorted by the groups from the outer one, by the details and by the region
                    var detailsSort = groups[groups.Count - 1].Member?.SortExpressions;
                    if (detailsSort != null)
                        sort.AddRange(GetRowSort(detailsSort));

                    if (sortExpr != null)
                        sort.AddRange(GetRowSort(sortExpr));

                    //the first or the last rows by a value: the rows sorted by it, the band prints the first ones
                    if (rowLimit != null && groups.Count == 1)
                        sort.InsertRange(0, new[] { rowLimit.Item1, rowLimit.Item2 });

                    var dataBand = new StiDataBand();
                    dataBand.Name = $"{container.Name}_{levelId}Db";
                    dataBand.DataSourceName = datasetName;
                    if (sort.Count > 0)
                        dataBand.Sort = sort.ToArray();

                    var detailsFilters = groups[groups.Count - 1].Member?.DetailsFilters;
                    var bandFilters = detailsFilters != null
                        ? MergeFilters(filters, ProcessFiltersType(detailsFilters, fullDatasetName))
                        : filters;
                    if (bandFilters != null) dataBand.Filters = bandFilters;
                    info.ToggleGroup = groupHeaders.LastOrDefault();
                    AddTablixCellsToContainer2(info, dataBand, groups[groups.Count - 1].DataLines, container, groups[groups.Count - 1]);
                    info.ToggleGroup = null;

                    if (rowLimit != null && groups.Count == 1)
                    {
                        dataBand.Expressions.Add(new StiAppExpression("Enabled", $"Line <= {rowLimit.Item3}"));
                    }
                    else if (rowLimit != null)
                    {
                        ThrowError(null, null,
                            $"Filter of the first rows of a grouped table not imported: {container.Name}");
                    }

                    var detailsLabel = groups[groups.Count - 1].Member?.DocumentMapLabel;
                    if (detailsLabel != null)
                        dataBand.Bookmark = new StiBookmarkExpression(detailsLabel);

                    if (firstDataBand == null)
                        firstDataBand = dataBand;

                    countScopes[dataBand] = dataBand;
                    rowScopes[dataBand] = null;
                    if (header != null)
                    {
                        countScopes[header] = dataBand;
                        rowScopes[header] = new List<string>();
                    }

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
                            countScopes[groupFooter] = groupHeaders[groupIndex];
                            rowScopes[groupFooter] = groupConditions.Take(groupIndex + 1).ToList();
                        }
                    }

                    if (firstGroup.FooterLines.Count > 0)
                    {
                        var footer = new StiFooterBand() { PrintIfEmpty = string.IsNullOrEmpty(noRowsMessage) };
                        footer.Name = $"{container.Name}_{levelId}Ft";
                        footer.PrintOnAllPages = firstGroup.RepeatFooter;
                        AddTablixCellsToContainer2(info, footer, firstGroup.FooterLines, container, firstGroup);
                        countScopes[footer] = dataBand;
                        rowScopes[footer] = new List<string>();
                    }
                }

                //the scope of the region itself is all its rows, the rows of its data band
                var regionName = baseNode.Attributes["Name"]?.Value;
                if (regionName != null && firstDataBand != null && !groupBands.ContainsKey(regionName))
                    groupBands[regionName] = firstDataBand;

                ResolveGroupScopes(container, groupBands, info.DataSetName, info.DataSetName2);

                if (hasColumnGroups)
                    ResolveMatrixCells(container, rowScopes, countScopes, info.DataSetName, info.DataSetName2);

                ResolveChartScopes(rowScopes);

                ResolveUserSorts(container, firstDataBand);
                ResolveCountScopes(countScopes, container);
                ResolveRunningScopes(container, groupBands, firstDataBand);

                if (!string.IsNullOrEmpty(noRowsMessage))
                    AddTablixNoRowsMessage(baseNode, container, page, datasetName, noRowsMessage);
            }

            ResolveGroupScopes(container, new Dictionary<string, StiBand>(), null, null);
            ResolveCountScopes(new Dictionary<StiBand, StiBand>(), container);
            ResolveRunningScopes(container, new Dictionary<string, StiBand>(), null);
            isInDataRegion = wasInDataRegion;

            baseContainer.Components.Add(container);
        }

        /// <summary>
        /// Returns the first filter of the first or the last rows of a tablix, TopN or BottomN, as the direction of
        /// the sort, the expression to sort by and the number of the rows; null when there is none or its number
        /// is not a constant. A data band of Stimulsoft has no such filter, the rows are sorted by the expression
        /// and the band prints the first ones. Reporting Services takes the rows tied with the last one too, the
        /// band does not; the order of the rows is the order of the expression, not of the data.
        /// </summary>
        private Tuple<string, string, int> GetTablixRowLimit(XmlNode filtersNode, string dataSet)
        {
            var limits = filtersNode.ChildNodes.Cast<XmlNode>().Where(IsRowLimitFilter).ToList();
            if (limits.Count == 0)
                return null;

            var filter = limits[0];
            var op = filter.SelectSingleNode("*[local-name()='Operator']").InnerText.Trim();
            var expression = filter.SelectSingleNode("*[local-name()='FilterExpression']")?.InnerText;
            var value = filter.SelectSingleNode(".//*[local-name()='FilterValue']")?.InnerText ?? string.Empty;

            var match = Regex.Match(value, @"^\s*=?\s*(C?Int\s*\(\s*)?(?<count>\d+)\s*\)?\s*$",
                RegexOptions.IgnoreCase);
            if (!match.Success || string.IsNullOrWhiteSpace(expression))
                return null;

            if (limits.Count > 1)
                ThrowError(null, null, $"Only the first filter of the first rows is imported: {op}");

            return Tuple.Create(op == "TopN" ? "DESC" : "ASC", ConvertExpression(expression, dataSet),
                int.Parse(match.Groups["count"].Value));
        }

        /// <summary>
        /// Returns the filters of the tablix and of its details together, a row passes both.
        /// </summary>
        private static StiFiltersCollection MergeFilters(StiFiltersCollection tablixFilters,
            StiFiltersCollection detailsFilters)
        {
            var filters = new StiFiltersCollection();
            foreach (var collection in new[] { tablixFilters, detailsFilters })
            {
                if (collection == null)
                    continue;

                foreach (StiFilter filter in collection)
                {
                    filters.Add(filter);
                }
            }

            return filters.Count > 0 ? filters : null;
        }

        /// <summary>
        /// Adds the message of a tablix without rows: a band at the end of the region that prints once when the
        /// data set has no rows, with the text in the font of the tablix. The headers and the footers do not
        /// print then, see ProcessTablixType. The rows are counted in the data set, the filters of the tablix
        /// are not taken into account.
        /// </summary>
        private void AddTablixNoRowsMessage(XmlNode baseNode, StiContainer container, StiPage page,
            string datasetName, string message)
        {
            if (string.IsNullOrEmpty(datasetName))
                return;

            var height = ToHi("0.25in");
            var band = new StiDataBand
            {
                Name = CheckComponentName(page.Report, container.Name + "NoRows"),
                Page = page,
                CountData = 1,
                CanGrow = true,
                ClientRectangle = new RectangleD(0, 0, container.Width, height)
            };

            var equal = ConvertSyntaxToCSharp ? "==" : "=";
            band.Expressions.Add(new StiAppExpression("Enabled", $"{datasetName}.Count {equal} 0"));

            var text = new StiText
            {
                Name = CheckComponentName(page.Report, container.Name + "NoRowsMessage"),
                Page = page,
                CanGrow = true,
                WordWrap = true,
                ClientRectangle = new RectangleD(0, 0, container.Width, height)
            };

            var style = GetChildNode(baseNode, "Style");
            if (style != null)
                ProcessStyleType(style, text, datasetName);

            //the frame and the fill of the tablix are on its panel
            text.Border = new StiBorder();
            text.Brush = new StiSolidBrush(Color.Transparent);
            text.Text = IsExpression(message) ? ConvertExpression(message, datasetName) : message;

            band.Components.Add(text);
            container.Components.Add(band);
        }

        /// <summary>
        /// Returns the visibility of each column of the body, the visibility of a static member of the
        /// columns or of its parent. Null when no column can be hidden or the members do not match the
        /// columns.
        /// </summary>
        private static List<XmlNode> GetColumnVisibilities(List<TablixMember> hierarchy, int columnCount)
        {
            var visibilities = new List<XmlNode>();
            AddColumnVisibilities(hierarchy, null, visibilities);

            if (visibilities.Count != columnCount || visibilities.All(v => v == null))
                return null;

            return visibilities;
        }

        /// <summary>
        /// Returns the visibility of each static row of the body. The visibility of a group is applied
        /// to its bands, so the row of a group gets null. Null when no row can be hidden or the members
        /// do not match the rows.
        /// </summary>
        private static List<XmlNode> GetRowVisibilities(List<TablixMember> hierarchy, int rowCount)
        {
            var visibilities = new List<XmlNode>();
            AddRowVisibilities(hierarchy, visibilities);

            if (visibilities.Count != rowCount || visibilities.All(v => v == null))
                return null;

            return visibilities;
        }

        private static void AddRowVisibilities(List<TablixMember> members, List<XmlNode> visibilities)
        {
            foreach (var member in members)
            {
                if (member.Childs != null && member.Childs.Count > 0)
                    AddRowVisibilities(member.Childs, visibilities);

                else
                    visibilities.Add(member.GroupName == null ? member.VisibilityNode : null);
            }
        }

        /// <summary>
        /// Hides a static row: the band when it has the row only, the line of the row otherwise. A row
        /// toggled by a text box is collapsed by the group around it, so it can be shown in the viewer;
        /// a row toggled out of any group keeps its initial visibility. Returns true when the row itself is
        /// hidden, false when the group collapses it.
        /// </summary>
        private bool ApplyRowVisibility(Info2 info, XmlNode visibility, StiBand band, StiContainer line,
            bool isSeveralLines)
        {
            var interaction = info.ToggleGroup?.Interaction as StiBandInteraction;
            if (visibility["ToggleItem"] != null && interaction != null)
            {
                interaction.CollapsingEnabled = true;
                info.ToggleGroup.Collapsed = new StiCollapsedExpression(GetCollapsedValue(visibility, info.DataSetName));
                return false;
            }

            if (isSeveralLines)
            {
                ProcessVisibility(visibility, line, info.DataSetName);
                band.CanShrink = true;
            }
            else
            {
                ProcessVisibility(visibility, band, info.DataSetName);
            }

            return true;
        }

        private string GetCollapsedValue(XmlNode visibility, string dataSet)
        {
            var hidden = GetNodeTextValue(visibility["Hidden"]);
            if (!IsExpression(hidden))
                return IsTrue(hidden) ? "{true}" : "{false}";

            var expression = ConvertPrintExpression(hidden, dataSet, out _);

            return expression != null ? "{" + expression + "}" : "{false}";
        }

        private static void AddColumnVisibilities(List<TablixMember> members, XmlNode parentVisibility,
            List<XmlNode> visibilities)
        {
            foreach (var member in members)
            {
                var visibility = member.VisibilityNode ?? parentVisibility;

                if (member.Childs != null && member.Childs.Count > 0)
                    AddColumnVisibilities(member.Childs, visibility, visibilities);

                else
                    visibilities.Add(visibility);
            }
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
            var hasHiddenRow = false;
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

                    var visibility = info.ColumnVisibilities?[index];
                    if (visibility != null)
                        ProcessVisibility(visibility, comp, info.DataSetName);

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

                    //a hidden column gives its place to the columns on the right, as in Reporting Services
                    if (info.ColumnVisibilities != null)
                        comp.DockStyle = StiDockStyle.Left;
                }

                if (hasColumnGroups)
                {
                    info.RowIndex = lineIndex;
                    ApplyColumnsGrouping(info, newCont);
                }

                var rowVisibility = info.RowVisibilities?[lineIndex];
                if (rowVisibility != null &&
                    ApplyRowVisibility(info, rowVisibility, baseBand, newCont, lines.Count > 1))
                    hasHiddenRow = true;

                if (lines.Count > 1 && !hasColumnGroups && TakeRowVisibility(newCont))
                {
                    //one cell for the whole row: the cell goes to the band without the row around it
                    var cell = newCont.Components[0];
                    newCont.Components.Remove(cell);
                    cell.Top = offsetY;
                    cell.GrowToHeight = false;
                    offsetY += rowHeight;
                    baseBand.Components.Add(cell);
                }

                else if (lines.Count > 1)
                {
                    newCont.Name = baseBand.Name + "_c" + indexLine.ToString();
                    newCont.Width = baseContainer.Width;
                    newCont.Height = rowHeight;
                    newCont.Top = offsetY;
                    newCont.CanGrow = true;
                    offsetY += rowHeight;
                    baseBand.Components.Add(newCont);
                }
                else
                {
                    baseBand.Components.AddRange(newCont.Components);
                }
                baseBand.Height += container.Height;
            }

            //the engine does not print a hidden row at all, so the DecreasingSize shift mode of the rows below
            //has nothing to move up from; the rows docked to the top take the place of a hidden row, as in
            //Reporting Services
            if (hasHiddenRow && lines.Count > 1)
            {
                foreach (StiComponent row in baseBand.Components)
                {
                    row.DockStyle = StiDockStyle.Top;
                }
            }
        }

        /// <summary>
        /// Returns true when the row has one cell that can stand for it. The visibility of the row goes to the cell:
        /// two expressions of the visibility are joined, the cell prints when both allow it. A script of the
        /// visibility, which calls the code of the report, is not joined with anything, the row stays then.
        /// </summary>
        private bool TakeRowVisibility(StiContainer line)
        {
            if (line.Components.Count != 1)
                return false;

            var cell = line.Components[0];
            var lineScript = line.BeforePrintEvent.Script;
            var cellScript = cell.BeforePrintEvent.Script;
            var lineEnabled = GetEnabledExpression(line);
            var cellEnabled = GetEnabledExpression(cell);
            if (!string.IsNullOrEmpty(lineScript) && (!string.IsNullOrEmpty(cellScript) || cellEnabled != null) ||
                !string.IsNullOrEmpty(cellScript) && lineEnabled != null)
                return false;

            if (lineEnabled != null && cellEnabled != null)
            {
                var and = ConvertSyntaxToCSharp ? "&&" : "AndAlso";
                cellEnabled.Expression = $"({lineEnabled.Expression}) {and} ({cellEnabled.Expression})";
            }

            else if (lineEnabled != null)
            {
                cell.Expressions.Add(lineEnabled);
            }

            if (!line.Enabled)
                cell.Enabled = false;

            if (!string.IsNullOrEmpty(lineScript))
                cell.BeforePrintEvent.Script = lineScript;

            line.Expressions.Clear();
            return true;
        }

        private static StiAppExpression GetEnabledExpression(StiComponent component)
        {
            return component.Expressions.Cast<StiAppExpression>().FirstOrDefault(e => e.Name == "Enabled");
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
                    columnGroupNames[groupHeader.Name] = group.Member.GroupName;
                    ApplyColumnGroup(groupHeader, group.Member, info.DataSetName, info.DataSetName2);
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
                    columnGroupNames[groupHeader.Name] = group.Member.GroupName;
                    ApplyColumnGroup(groupHeader, group.Member, info.DataSetName, info.DataSetName2);
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

                //add header to cells. A member without a header takes no column of the headers, as in
                //Reporting Services: a group may leave its header to a static member inside it
                var rowHeight = ScanMembersHeightRecursive(member);
                for (int index2 = 0; index2 < rowHeight; index2++)
                {
                    var rowIndex2 = info.RowIndex + index2;
                    while (info.Cells.Count < rowIndex2 + 1)
                        info.Cells.Add(new List<StiComponent>());

                    var newComp = member.Header?.Clone() as StiComponent;
                    if (newComp == null)
                        continue;

                    if (index2 > 0)
                        newComp.Name += "_" + index2.ToString();

                    //the copies of the header are merged by the value and the tag, the tag is the expression of
                    //the text given at the end of the import, see ApplyMergeTags
                    if (newComp is StiText stiText)
                        stiText.ProcessingDuplicates = StiProcessingDuplicatesType.GlobalBasedOnValueAndTagMerge;

                    if (newComp is StiImage stiImage)
                        stiImage.ProcessingDuplicates = StiImageProcessingDuplicatesType.GlobalMerge;

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
                            //a static member of static members takes all their rows, its members are not
                            //looked at again, so only the first row was taken and the others were lost
                            var memberLines = Enumerable.Range(info.RowIndex, rowHeight).ToList();

                            if (isRoot)
                            {
                                if (!wasGroup)
                                {
                                    baseGroup.HeaderLines.AddRange(memberLines);
                                    baseGroup.RepeatHeader |= member.RepeatOnNewPage;
                                }
                                else
                                {
                                    baseGroup.FooterLines.AddRange(memberLines);
                                    baseGroup.RepeatFooter |= member.RepeatOnNewPage;
                                }
                            }
                            else
                            {
                                var lineGroup = baseGroup.Parent ?? baseGroup;

                                if (!wasGroup)
                                {
                                    lineGroup.GroupHeaderLines.AddRange(memberLines);
                                    lineGroup.RepeatGroupHeader |= member.RepeatOnNewPage;
                                }
                                else
                                {
                                    lineGroup.GroupFooterLines.AddRange(memberLines);
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

                if (member.Header != null)
                    info.ColumnIndex++;

                AnalyzeGroupRecursive(member, info, group2 ?? baseGroup, hasGroup);

                info.RowIndex++;
            }
            info.RowIndex--;
        }

        /// <summary>
        /// Gives a text merged by the value and the tag the expression of its text as the tag: a text that repeats
        /// the value and the tag of the text above prints no text when it renders, and the texts merge by their
        /// tags, not by the texts they print. The tag is given at the end, when the expression of the text is
        /// final: the scopes of the aggregates are put into the text after the cells are copied.
        /// </summary>
        private void ApplyMergeTags()
        {
            foreach (var text in report.GetComponents().Cast<StiComponent>().OfType<StiText>())
            {
                var isMerged = text.ProcessingDuplicates == StiProcessingDuplicatesType.BasedOnValueAndTagMerge ||
                    text.ProcessingDuplicates == StiProcessingDuplicatesType.GlobalBasedOnValueAndTagMerge;

                if (isMerged && string.IsNullOrEmpty(text.Tag.Value))
                    text.Tag = new StiTagExpression(text.Text.Value);
            }
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
                        //applied to the bands, the cells or the groups when they are made
                        if (node["Hidden"] != null)
                            tablixMember.VisibilityNode = node;
                        break;

                    case "RepeatOnNewPage":
                        tablixMember.RepeatOnNewPage = IsTrue(GetNodeTextValue(node));
                        break;

                    case "KeepTogether":
                        tablixMember.KeepTogether = IsTrue(GetNodeTextValue(node));
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

                    case "DocumentMapLabel":
                        tablixMember.DocumentMapLabel = ConvertInteractiveValue(node, dataSet);
                        break;

                    case "Filters":
                        tablixMember.DetailsFilters = node;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            //the filters of the details take the rows; a filter of a group takes the whole instances of the
            //group by the values of the group, the data band has no such filter
            if (tablixMember.DetailsFilters != null && tablixMember.GroupExpressions?.Count > 0)
            {
                ThrowError(baseNode.Name, "Filters");
                tablixMember.DetailsFilters = null;
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
