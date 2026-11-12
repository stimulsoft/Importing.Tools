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

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.CrossTab;
using Stimulsoft.Report.CrossTab.Core;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The Tablix data region. A Tablix whose columns are fixed is a table with nested row groups, it becomes
    /// bands: a row of a group goes to a group header band, a row of details goes to the data band, a row
    /// that stands before or after the rows of a group ("Total") goes to the header or the footer of the outer
    /// group. A Tablix with column groups is a matrix, it becomes a cross-tab.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Methods.Tablix
        private void BuildTablix(StiRdlxNode tablix, StiContainer container, double left, StiRdlxScope scope)
        {
            var dataSet = GetRegionDataSet(tablix, scope);
            var rowHierarchy = tablix.Child("RowHierarchy", "TablixRowHierarchy") ?? new StiRdlxNode("RowHierarchy");
            var columnHierarchy = tablix.Child("ColumnHierarchy", "TablixColumnHierarchy") ??
                new StiRdlxNode("ColumnHierarchy");
            var body = tablix.Child("Body", "TablixBody") ?? new StiRdlxNode("Body");
            var isStatic = dataSet == null || IsCurrentRowRegion(dataSet, scope);

            if (!isStatic && HasGroup(GetMembers(columnHierarchy)))
            {
                var rect = new RectangleD(left, 0, ToHi(tablix.Get("Width")), ToHi(tablix.Get("Height")));
                var band = CreateStaticBand(rect.Height);
                container.Components.Add(band);
                band.Components.Add(CreateCrossTab(tablix, rect, dataSet, rowHierarchy, columnHierarchy, body));
                return;
            }

            var context = new StiRdlxTablixContext();
            context.Container = container;
            context.Scope = scope;
            context.LevelEdges = ToEdges(GetLevelSizes(rowHierarchy), left);
            context.ColumnEdges = GetEdges(body.Items("Columns", "TablixColumns"),
                context.LevelEdges[context.LevelEdges.Length - 1]);
            context.BodyRows = body.Items("Rows", "TablixRows");

            if (isStatic)
            {
                context.HeaderBand = CreateStaticBand(0);
                context.FooterBand = context.HeaderBand;
            }
            else
            {
                context.DataBand = CreateDataBand(tablix, dataSet, scope);
                context.Scope = CreateRegionScope(dataSet, context.DataBand);
                context.HeaderBand = new StiHeaderBand { Name = CreateName("HeaderBand") };
                context.FooterBand = new StiFooterBand { Name = CreateName("FooterBand") };
            }

            container.Components.Add(context.HeaderBand);
            context.Bottoms[context.HeaderBand] = AddTablixColumnHeaders(tablix, columnHierarchy, context);

            AddTablixRows(GetMembers(rowHierarchy), 0, new List<StiRdlxTablixHeader>(), context.HeaderBand,
                context.FooterBand, context);

            if (!isStatic)
            {
                container.Components.Add(context.DataBand);

                for (var index = context.GroupFooters.Count - 1; index >= 0; index--)
                {
                    container.Components.Add(context.GroupFooters[index]);
                }

                if (context.Bottoms.ContainsKey(context.FooterBand))
                    container.Components.Add(context.FooterBand);
            }

            foreach (var pair in context.Bottoms)
            {
                pair.Key.Height = pair.Value;
            }
        }

        /// <summary>
        /// The members of a hierarchy or of a member: Members and Children in JSON, TablixMembers in XML.
        /// </summary>
        private static List<StiRdlxNode> GetMembers(StiRdlxNode node)
        {
            foreach (var name in new[] { "Members", "Children", "TablixMembers" })
            {
                var members = node.Items(name);
                if (members.Count > 0)
                    return members;
            }

            return new List<StiRdlxNode>();
        }

        private static StiRdlxNode GetMemberHeader(StiRdlxNode member)
        {
            return member.Child("Header", "TablixHeader");
        }

        private static bool HasGroup(List<StiRdlxNode> members)
        {
            foreach (var member in members)
            {
                if (GetGroupExpression(member.Child("Group")) != null || HasGroup(GetMembers(member)))
                    return true;
            }

            return false;
        }

        /// <summary>
        /// The sizes of the levels of a hierarchy: LevelSizes in JSON, the sizes of the headers in XML.
        /// </summary>
        private List<double> GetLevelSizes(StiRdlxNode hierarchy)
        {
            var sizes = new List<double>();

            foreach (var size in hierarchy.Items("LevelSizes"))
            {
                sizes.Add(ToHi(size.Value));
            }

            if (sizes.Count > 0)
                return sizes;

            var members = GetMembers(hierarchy);

            while (members.Count > 0)
            {
                var header = GetMemberHeader(members[0]);
                if (header != null)
                    sizes.Add(ToHi(header.Get("Size")));

                members = GetMembers(members[0]);
            }

            return sizes;
        }

        private static double[] ToEdges(List<double> sizes, double start)
        {
            var edges = new double[sizes.Count + 1];
            edges[0] = start;

            for (var index = 0; index < sizes.Count; index++)
            {
                edges[index + 1] = edges[index] + sizes[index];
            }

            return edges;
        }

        private static double GetEdge(double[] edges, int index)
        {
            return edges[Math.Max(0, Math.Min(index, edges.Length - 1))];
        }

        /// <summary>
        /// Places the corner and the headers of the columns on the header band and returns their height.
        /// </summary>
        private double AddTablixColumnHeaders(StiRdlxNode tablix, StiRdlxNode columnHierarchy,
            StiRdlxTablixContext context)
        {
            var levelEdges = ToEdges(GetLevelSizes(columnHierarchy), 0);
            var rowIndex = 0;

            foreach (var row in tablix.Items("Corner"))
            {
                var column = 0;

                foreach (var cell in row.Nodes)
                {
                    var colSpan = Math.Max(1, GetInt(cell.Get("ColSpan"), 1));
                    var rowSpan = Math.Max(1, GetInt(cell.Get("RowSpan"), 1));
                    var item = GetCellItem(cell);

                    if (item != null)
                    {
                        var x = GetEdge(context.LevelEdges, column);
                        var y = GetEdge(levelEdges, rowIndex);
                        var rect = new RectangleD(x, y, GetEdge(context.LevelEdges, column + colSpan) - x,
                            GetEdge(levelEdges, rowIndex + rowSpan) - y);

                        AddComponent(context.HeaderBand, CreateItem(item, rect, context.Scope));
                    }

                    column += colSpan;
                }

                rowIndex++;
            }

            var leaf = 0;
            AddTablixColumnMembers(GetMembers(columnHierarchy), 0, levelEdges, context, ref leaf);

            return levelEdges[levelEdges.Length - 1];
        }

        private void AddTablixColumnMembers(List<StiRdlxNode> members, int level, double[] levelEdges,
            StiRdlxTablixContext context, ref int leaf)
        {
            foreach (var member in members)
            {
                var header = GetMemberHeader(member);
                var levelCount = header == null ? 0 : Math.Max(1, GetInt(header.Get("LevelCount"), 1));
                var first = leaf;
                var children = GetMembers(member);

                if (children.Count > 0)
                    AddTablixColumnMembers(children, level + levelCount, levelEdges, context, ref leaf);

                else
                    leaf++;

                var item = header == null ? null : GetCellItem(header);
                if (item == null)
                    continue;

                var x = GetEdge(context.ColumnEdges, first);
                var y = GetEdge(levelEdges, level);
                var rect = new RectangleD(x, y, GetEdge(context.ColumnEdges, leaf) - x,
                    GetEdge(levelEdges, level + levelCount) - y);

                AddComponent(context.HeaderBand, CreateItem(item, rect, context.Scope));
            }
        }

        private static void AddComponent(StiContainer container, StiComponent component)
        {
            if (component != null)
                container.Components.Add(component);
        }

        /// <summary>
        /// Places the rows of the row members. A member before the first group of its level puts its rows on
        /// the header band of the outer group, a member after it puts them on the footer band.
        /// </summary>
        private void AddTablixRows(List<StiRdlxNode> members, int level, List<StiRdlxTablixHeader> path,
            StiBand headerBand, StiBand footerBand, StiRdlxTablixContext context)
        {
            var groupIndex = -1;

            for (var index = 0; index < members.Count && groupIndex < 0; index++)
            {
                if (members[index].Has("Group"))
                    groupIndex = index;
            }

            for (var index = 0; index < members.Count; index++)
            {
                var member = members[index];
                var header = GetMemberHeader(member);
                var levelCount = header == null ? 0 : Math.Max(1, GetInt(header.Get("LevelCount"), 1));
                var memberPath = new List<StiRdlxTablixHeader>(path);
                var item = header == null ? null : GetCellItem(header);

                if (item != null)
                    memberPath.Add(new StiRdlxTablixHeader { Item = item, Level = level, LevelCount = levelCount });

                var group = member.Child("Group");
                var isGroup = context.DataBand != null && group != null;
                var isCondition = GetGroupExpression(group) != null;
                var band = groupIndex < 0 || index < groupIndex ? headerBand : footerBand;
                var children = GetMembers(member);

                if (isGroup)
                    ApplySorting(member, context.DataBand, context.Scope.DataSet);

                if (children.Count > 0 && isGroup && isCondition)
                {
                    var groupHeader = CreateGroupHeader(group, false, context.Scope);
                    var groupFooter = CreateGroupFooter();

                    context.Container.Components.Add(groupHeader);
                    context.GroupFooters.Add(groupFooter);
                    context.Bottoms[groupHeader] = 0;
                    context.Bottoms[groupFooter] = 0;

                    AddTablixRows(children, level + levelCount, memberPath, groupHeader, groupFooter, context);
                }
                else if (children.Count > 0)
                {
                    var isBetween = groupIndex >= 0 && !isGroup;

                    AddTablixRows(children, level + levelCount, memberPath, isBetween ? band : headerBand,
                        isBetween ? band : footerBand, context);
                }
                else
                {
                    var rowIndex = GetInt(member.Get("BodyIndex"), context.RowIndex);
                    context.RowIndex = rowIndex + 1;

                    if (rowIndex >= context.BodyRows.Count)
                        continue;

                    if (isGroup && isCondition)
                    {
                        band = CreateGroupHeader(group, false, context.Scope);
                        context.Container.Components.Add(band);
                    }
                    else if (isGroup)
                    {
                        band = context.DataBand;
                    }

                    AddTablixRow(band, memberPath, context.BodyRows[rowIndex], context, isGroup);
                }
            }
        }

        /// <summary>
        /// Places the header cells of a row and the cells of its body on a band.
        /// </summary>
        private void AddTablixRow(StiBand band, List<StiRdlxTablixHeader> path, StiRdlxNode row,
            StiRdlxTablixContext context, bool isRepeated)
        {
            double top;
            context.Bottoms.TryGetValue(band, out top);

            var height = ToHi(row.Get("Height"));

            for (var index = 0; index < path.Count; index++)
            {
                var header = path[index];
                if (header.IsPlaced)
                    continue;

                header.IsPlaced = true;

                var x = GetEdge(context.LevelEdges, header.Level);
                var rect = new RectangleD(x, top, GetEdge(context.LevelEdges, header.Level + header.LevelCount) - x,
                    height);

                var component = CreateItem(header.Item, rect, context.Scope);
                var text = component as StiText;

                // The header of an outer group stands on every row of the inner group, so it is merged.
                if (text != null && isRepeated && index < path.Count - 1)
                    text.ProcessingDuplicates = StiProcessingDuplicatesType.Merge;

                AddComponent(band, component);
            }

            var column = 0;

            foreach (var cell in row.Items("Cells", "TablixCells"))
            {
                var colSpan = Math.Max(1, GetInt(cell.Get("ColSpan"), 1));
                var item = GetCellItem(cell);

                if (item != null)
                {
                    var x = GetEdge(context.ColumnEdges, column);
                    var rect = new RectangleD(x, top, GetEdge(context.ColumnEdges, column + colSpan) - x, height);
                    AddComponent(band, CreateItem(item, rect, context.Scope));
                }

                column += colSpan;
            }

            context.Bottoms[band] = top + height;
        }
        #endregion

        #region Methods.CrossTab
        /// <summary>
        /// Makes a cross-tab of a Tablix with column groups: the groups of the rows and of the columns become
        /// the rows and the columns of the cross-tab, the cells of the body become its summaries.
        /// </summary>
        private StiComponent CreateCrossTab(StiRdlxNode tablix, RectangleD rect, StiRdlxDataSet dataSet,
            StiRdlxNode rowHierarchy, StiRdlxNode columnHierarchy, StiRdlxNode body)
        {
            var crossTab = new StiCrossTab(rect);
            crossTab.Name = CreateName(tablix.Get("Name", "CrossTab"));
            crossTab.DataSourceName = dataSet.SourceName;

            var scope = new StiRdlxScope { DataSet = dataSet, IsRegion = true };

            foreach (var filter in tablix.Items("Filters"))
            {
                var expression = ConvertFilter(filter, scope);
                if (expression == null)
                    continue;

                crossTab.Filters.Add(new StiFilter("{" + expression + "}"));
                crossTab.FilterOn = true;
            }

            var rowGroups = new List<StiRdlxNode>();
            var columnGroups = new List<StiRdlxNode>();
            CollectGroups(GetMembers(rowHierarchy), rowGroups);
            CollectGroups(GetMembers(columnHierarchy), columnGroups);

            var rowSizes = GetLevelSizes(rowHierarchy);
            var columnSizes = GetLevelSizes(columnHierarchy);
            var rows = body.Items("Rows", "TablixRows");
            var cellHeight = rows.Count > 0 ? ToHi(rows[0].Get("Height")) : 20;
            var columns = body.Items("Columns", "TablixColumns");
            var cellWidth = columns.Count > 0 ? ToHi(columns[0].Value ?? columns[0].Get("Width")) : 80;

            for (var index = 0; index < rowGroups.Count; index++)
            {
                var row = new StiCrossRow();
                var total = new StiCrossRowTotal();
                var width = index < rowSizes.Count ? rowSizes[index] : cellWidth;

                InitCrossHeader(row, total, rowGroups[index], crossTab.Name + "_Row" + (index + 1), scope);
                row.ClientRectangle = new RectangleD(0, 0, width, cellHeight);
                total.ClientRectangle = new RectangleD(0, 0, width, cellHeight);

                crossTab.Components.Add(row);
                crossTab.Components.Add(total);
            }

            for (var index = 0; index < columnGroups.Count; index++)
            {
                var column = new StiCrossColumn();
                var total = new StiCrossColumnTotal();
                var height = index < columnSizes.Count ? columnSizes[index] : cellHeight;

                InitCrossHeader(column, total, columnGroups[index], crossTab.Name + "_Column" + (index + 1), scope);
                column.ClientRectangle = new RectangleD(0, 0, cellWidth, height);
                total.ClientRectangle = new RectangleD(0, 0, cellWidth, height);

                crossTab.Components.Add(column);
                crossTab.Components.Add(total);
            }

            var summaryIndex = 1;
            var used = new HashSet<string>(StringComparer.Ordinal);

            foreach (var row in rows)
            {
                foreach (var cell in row.Items("Cells", "TablixCells"))
                {
                    var item = GetCellItem(cell);
                    var value = item == null ? null : item.Get("Value");

                    if (string.IsNullOrEmpty(value) || !used.Add(value))
                        continue;

                    var summary = new StiCrossSummary();
                    summary.Name = CreateName(crossTab.Name + "_Sum" + summaryIndex++);
                    InitCrossSummary(summary, ConvertToExpression(value, scope));
                    ApplyTextStyle(item, summary);
                    summary.ClientRectangle = new RectangleD(0, 0, cellWidth, cellHeight);

                    crossTab.Components.Add(summary);
                }
            }

            return crossTab;
        }

        private void CollectGroups(List<StiRdlxNode> members, List<StiRdlxNode> groups)
        {
            foreach (var member in members)
            {
                if (GetGroupExpression(member.Child("Group")) != null)
                    groups.Add(member);

                CollectGroups(GetMembers(member), groups);
            }
        }

        private void InitCrossHeader(StiCrossHeader header, StiCrossTotal total, StiRdlxNode member, string name,
            StiRdlxScope scope)
        {
            var value = "{" + ConvertToExpression(GetGroupExpression(member.Child("Group")), scope) + "}";
            var memberHeader = GetMemberHeader(member);
            var item = memberHeader == null ? null : GetCellItem(memberHeader);

            header.Name = CreateName(name);
            header.Value.Value = value;
            header.DisplayValue.Value = item != null && !string.IsNullOrEmpty(item.Get("Value"))
                ? ConvertValue(item.Get("Value"), scope)
                : value;

            if (item != null)
            {
                ApplyTextStyle(item, header);
                ApplyTextStyle(item, total);
            }

            total.Name = CreateName(name + "_Total");
            total.Text.Value = "Total";
            total.Guid = Guid.NewGuid().ToString("N");
            header.TotalGuid = total.Guid;
        }

        /// <summary>
        /// Splits the expression of a cell into the function of the summary and its value.
        /// </summary>
        private static void InitCrossSummary(StiCrossSummary summary, string expression)
        {
            var match = Regex.Match(expression, @"\b(Sum|Avg|Count|Min|Max|CountDistinct|Median)\(([^()]*)\)");

            summary.Summary = StiSummaryType.None;
            summary.Value.Value = "{" + expression + "}";

            if (!match.Success)
                return;

            switch (match.Groups[1].Value)
            {
                case "Sum":
                    summary.Summary = StiSummaryType.Sum;
                    break;

                case "Avg":
                    summary.Summary = StiSummaryType.Average;
                    break;

                case "Count":
                    summary.Summary = StiSummaryType.Count;
                    break;

                case "Min":
                    summary.Summary = StiSummaryType.Min;
                    break;

                case "Max":
                    summary.Summary = StiSummaryType.Max;
                    break;

                case "CountDistinct":
                    summary.Summary = StiSummaryType.CountDistinct;
                    break;

                case "Median":
                    summary.Summary = StiSummaryType.Median;
                    break;
            }

            var inner = match.Groups[2].Value.Trim();
            summary.Value.Value = "{" + (inner.Length == 0 ? "1" : inner) + "}";
        }
        #endregion
    }
}
