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

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The data regions: Table, List and BandedList. A region becomes a set of bands: a header band, group
    /// header bands, a data band, group footer bands and a footer band. A region without data, or one that
    /// stands inside a region of the same data set and so shows its current row, is a band printed once.
    /// A grouped list, and a grouped fixed page, that holds one table of the same data set is joined with
    /// it: the group of the list becomes the outer group of the table, the items above the table go to the
    /// group header, the items below it go to the group footer.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Fields.Regions
        /// <summary>
        /// The bands of a table are made inside a group that starts a new page for each of its values, so
        /// the page breaks of the groups of the table are not imported.
        /// </summary>
        private bool isPageBreakSuppressed;
        #endregion

        #region Methods.Regions
        private void BuildDataRegion(StiRdlxNode node, StiContainer container, double left, StiRdlxScope scope)
        {
            switch ((node.ItemType ?? "").ToLowerInvariant())
            {
                case "table":
                    BuildTable(node, container, left, scope);
                    break;

                case "list":
                    BuildList(node, container, left, scope);
                    break;

                case "bandedlist":
                    BuildBandedList(node, container, left, scope);
                    break;

                case "tablix":
                    BuildTablix(node, container, left, scope);
                    break;
            }
        }

        /// <summary>
        /// The region shows the current row of the outer region: it has the same data set and is printed
        /// on a band of that region.
        /// </summary>
        private static bool IsCurrentRowRegion(StiRdlxDataSet dataSet, StiRdlxScope scope)
        {
            return scope.IsRegion && scope.DataSet == dataSet;
        }

        private StiRdlxScope CreateRegionScope(StiRdlxDataSet dataSet, StiDataBand band)
        {
            return new StiRdlxScope { DataSet = dataSet, Band = band, IsRegion = true };
        }

        private StiDataBand CreateDataBand(StiRdlxNode region, StiRdlxDataSet dataSet, StiRdlxScope scope)
        {
            var band = new StiDataBand();
            band.Name = CreateName("DataBand");
            band.DataSourceName = dataSet.SourceName;

            if (dataSet.Band == null)
                dataSet.Band = band;

            // A nested data set is printed inside a band of its parent: only the rows of the current row of
            // the parent are left. The tables of a JSON database are filled on demand, so a relation of the
            // dictionary can not be made of them, a filter by the keys does the same.
            if (dataSet.Parent != null && dataSet.Rows != null && scope.IsRegion && scope.DataSet == dataSet.Parent)
            {
                band.Filters.Add(new StiFilter("{" + dataSet.SourceName + "." + ParentKeyColumn + " == " +
                    dataSet.Parent.SourceName + "." + KeyColumn + "}"));
                band.FilterOn = true;
            }

            var regionScope = CreateRegionScope(dataSet, band);

            if (!dataSet.IsFiltered)
                ApplyFilters(dataSet.Node.Items("Filters"), band, regionScope);

            ApplyFilters(region.Items("Filters"), band, regionScope);

            return band;
        }

        private StiGroupHeaderBand CreateGroupHeader(StiRdlxNode group, bool isSorted, StiRdlxScope scope)
        {
            var band = new StiGroupHeaderBand();
            band.Name = CreateName("GroupHeaderBand");
            band.Condition.Value = GetGroupCondition(group, scope);
            band.SortDirection = isSorted ? StiGroupSortDirection.None : StiGroupSortDirection.Ascending;
            band.CanBreak = true;

            // A page break of a group, before it, after it or between the groups, is a new page before
            // the header of the next group: a break after the header band itself would tear the group.
            // Inside a group that already starts a page for each of its values the break is not repeated.
            var pageBreak = (group == null ? "" : group.Get("PageBreak", "")).ToLowerInvariant();

            if (!isPageBreakSuppressed && (pageBreak == "between" || pageBreak == "start" || pageBreak == "end" ||
                pageBreak == "startandend"))
                band.NewPageBefore = true;

            return band;
        }

        private StiGroupFooterBand CreateGroupFooter()
        {
            var band = new StiGroupFooterBand();
            band.Name = CreateName("GroupFooterBand");
            return band;
        }

        /// <summary>
        /// Returns the first expression of a group, null when the group has none.
        /// </summary>
        private static string GetGroupExpression(StiRdlxNode group)
        {
            if (group == null)
                return null;

            foreach (var expression in group.Items("GroupExpressions"))
            {
                if (!string.IsNullOrEmpty(expression.Value))
                    return expression.Value;
            }

            return null;
        }

        /// <summary>
        /// The condition of a group band. Several expressions of a group are joined into one text.
        /// </summary>
        private string GetGroupCondition(StiRdlxNode group, StiRdlxScope scope)
        {
            var parts = new List<string>();

            if (group != null)
            {
                foreach (var expression in group.Items("GroupExpressions"))
                {
                    if (!string.IsNullOrEmpty(expression.Value))
                        parts.Add(ConvertToExpression(expression.Value, scope));
                }
            }

            if (parts.Count == 0)
                return "{\"\"}";

            if (parts.Count == 1)
                return "{" + parts[0] + "}";

            for (var index = 0; index < parts.Count; index++)
            {
                parts[index] = "ToString(" + parts[index] + ")";
            }

            return "{" + string.Join(" + \"|\" + ", parts) + "}";
        }

        private void ApplyFilters(List<StiRdlxNode> filters, StiDataBand band, StiRdlxScope scope)
        {
            foreach (var filter in filters)
            {
                var expression = ConvertFilter(filter, scope);
                if (expression == null)
                    continue;

                band.Filters.Add(new StiFilter("{" + expression + "}"));
                band.FilterOn = true;
            }
        }

        private string ConvertFilter(StiRdlxNode filter, StiRdlxScope scope)
        {
            var left = ConvertToExpression(filter.Get("FilterExpression"), scope);
            var values = new List<string>();

            foreach (var value in filter.Items("FilterValues"))
            {
                values.Add(ConvertToExpression(value.Value, scope));
            }

            if (values.Count == 0)
                return null;

            var operation = filter.Get("Operator", "Equal").ToLowerInvariant();

            switch (operation)
            {
                case "equal":
                    return left + " == " + values[0];

                case "notequal":
                    return left + " != " + values[0];

                case "greaterthan":
                    return left + " > " + values[0];

                case "greaterthanorequal":
                    return left + " >= " + values[0];

                case "lessthan":
                    return left + " < " + values[0];

                case "lessthanorequal":
                    return left + " <= " + values[0];

                case "between":
                    if (values.Count < 2)
                        return null;

                    return "(" + left + " >= " + values[0] + " && " + left + " <= " + values[1] + ")";

                case "in":
                    var parts = new List<string>();

                    foreach (var value in values)
                    {
                        parts.Add(left + " == " + value);
                    }

                    return "(" + string.Join(" || ", parts) + ")";

                default:
                    Warn($"The filter operator '{filter.Get("Operator")}' is not imported.");
                    return null;
            }
        }

        /// <summary>
        /// Applies the sorting by fields. A sorting by an expression is not imported.
        /// </summary>
        private bool ApplySorting(StiRdlxNode node, StiDataBand band, StiRdlxDataSet dataSet)
        {
            if (node == null)
                return false;

            var sort = new List<string>(band.Sort ?? new string[0]);
            var count = sort.Count;

            foreach (var item in node.Items("SortExpressions", "Sorting"))
            {
                var value = item.Get("Value") ?? item.Get("SortExpression") ?? "";
                var match = Regex.Match(value, @"^=\s*Fields!(\w+)\.Value\s*$|^\{(\w+)\}$|^=\s*Fields\.(\w+)\s*$");
                if (!match.Success)
                    continue;

                var field = match.Groups[1].Value + match.Groups[2].Value + match.Groups[3].Value;
                if (!dataSet.Columns.ContainsKey(field))
                    continue;

                var isDescending = item.Get("Direction", "").StartsWith("Desc", StringComparison.OrdinalIgnoreCase);
                sort.Add(isDescending ? "DESC" : "ASC");
                sort.Add(dataSet.Columns[field]);
            }

            band.Sort = sort.ToArray();
            return sort.Count > count;
        }

        private static List<StiRdlxNode> GetRows(StiRdlxNode part)
        {
            return part == null ? new List<StiRdlxNode>() : part.Items("TableRows");
        }

        private double[] GetEdges(List<StiRdlxNode> columns, double left)
        {
            var edges = new double[columns.Count + 1];
            edges[0] = left;

            for (var index = 0; index < columns.Count; index++)
            {
                var width = columns[index].Value ?? columns[index].Get("Width");
                edges[index + 1] = edges[index] + ToHi(width);
            }

            return edges;
        }

        /// <summary>
        /// Places the cells of table rows on a band from the given top and returns the bottom of the rows.
        /// </summary>
        private double AddRows(List<StiRdlxNode> rows, double[] edges, StiContainer band, double top,
            StiRdlxScope scope)
        {
            var heights = new List<double>();
            var visibleRows = new List<StiRdlxNode>();

            foreach (var row in rows)
            {
                var visibility = row.Child("Visibility");
                if (visibility != null && !visibility.Has("ToggleItem") && visibility.GetBool("Hidden"))
                    continue;

                visibleRows.Add(row);
                heights.Add(ToHi(row.Get("Height")));
            }

            var busy = new int[edges.Length];

            for (var rowIndex = 0; rowIndex < visibleRows.Count; rowIndex++)
            {
                var column = 0;

                foreach (var cell in visibleRows[rowIndex].Items("TableCells"))
                {
                    while (column < busy.Length - 1 && busy[column] > 0)
                    {
                        column++;
                    }

                    if (column >= edges.Length - 1)
                        break;

                    var colSpan = Math.Max(1, GetInt(cell.Get("ColSpan"), 1));
                    var rowSpan = Math.Max(1, GetInt(cell.Get("RowSpan"), 1));
                    var last = Math.Min(column + colSpan, edges.Length - 1);
                    var height = 0d;

                    for (var index = rowIndex; index < Math.Min(rowIndex + rowSpan, heights.Count); index++)
                    {
                        height += heights[index];
                    }

                    for (var index = column; index < last && rowSpan > 1; index++)
                    {
                        busy[index] = rowSpan;
                    }

                    var item = GetCellItem(cell);
                    if (item != null)
                    {
                        var rect = new RectangleD(edges[column], top, edges[last] - edges[column], height);
                        var component = CreateItem(item, rect, scope);

                        if (component != null)
                            band.Components.Add(component);
                    }

                    column = last;
                }

                for (var index = 0; index < busy.Length; index++)
                {
                    if (busy[index] > 0)
                        busy[index]--;
                }

                top += heights[rowIndex];
            }

            return top;
        }

        /// <summary>
        /// The item of a cell: Item in JSON, the first item of ReportItems or CellContents in XML.
        /// </summary>
        private static StiRdlxNode GetCellItem(StiRdlxNode cell)
        {
            var item = cell.Child("Item");
            if (item != null)
                return item;

            var contents = cell.Child("CellContents");
            if (contents != null && contents.Nodes.Count > 0)
                return contents.Nodes[0];

            var items = cell.Items("ReportItems");
            return items.Count > 0 ? items[0] : null;
        }

        private static int GetInt(string value, int defaultValue)
        {
            int number;
            return int.TryParse(value, out number) ? number : defaultValue;
        }

        private void AddBand(StiBand band, StiContainer container, double height)
        {
            band.Height = Math.Max(0, height);
            container.Components.Add(band);
        }
        #endregion

        #region Methods.Table
        private void BuildTable(StiRdlxNode table, StiContainer container, double left, StiRdlxScope scope)
        {
            var dataSet = GetRegionDataSet(table, scope);
            var edges = GetEdges(table.Items("TableColumns"), left);
            var headerRows = GetRows(table.Child("Header"));
            var footerRows = GetRows(table.Child("Footer"));
            var hasRows = GetRows(table.Child("Details")).Count > 0 || table.Items("TableGroups").Count > 0;

            if (dataSet == null || !hasRows || IsCurrentRowRegion(dataSet, scope))
            {
                AddStaticTable(table, edges, container, scope);
                return;
            }

            var dataBand = CreateDataBand(table, dataSet, scope);
            var regionScope = CreateRegionScope(dataSet, dataBand);

            if (headerRows.Count > 0)
            {
                var header = new StiHeaderBand();
                header.Name = CreateName("HeaderBand");
                header.PrintOnAllPages = table.Child("Header").GetBool("RepeatOnNewPage");
                AddBand(header, container, AddRows(headerRows, edges, header, 0, regionScope));
            }

            AddTableDataBands(table, edges, container, dataBand, dataSet, regionScope);

            if (footerRows.Count > 0)
            {
                var footer = new StiFooterBand();
                footer.Name = CreateName("FooterBand");
                AddBand(footer, container, AddRows(footerRows, edges, footer, 0, regionScope));
            }
        }

        /// <summary>
        /// Adds the bands of the groups and of the details of a table: group headers, the data band, group
        /// footers. A grouping of the details is the innermost group: its rows go to the group header.
        /// </summary>
        private void AddTableDataBands(StiRdlxNode table, double[] edges, StiContainer container,
            StiDataBand dataBand, StiRdlxDataSet dataSet, StiRdlxScope regionScope)
        {
            var details = table.Child("Details");
            var isSorted = ApplySorting(details, dataBand, dataSet);
            var footers = new List<StiGroupFooterBand>();
            var footerRows = new List<List<StiRdlxNode>>();

            foreach (var group in table.Items("TableGroups"))
            {
                isSorted |= ApplySorting(group, dataBand, dataSet);

                var header = CreateGroupHeader(group.Child("Group", "Grouping"), isSorted, regionScope);
                AddBand(header, container, AddRows(GetRows(group.Child("Header")), edges, header, 0, regionScope));

                footers.Add(CreateGroupFooter());
                footerRows.Add(GetRows(group.Child("Footer")));
            }

            var detailRows = GetRows(details);
            var detailGroup = details == null ? null : details.Child("Group", "Grouping");

            if (GetGroupExpression(detailGroup) != null)
            {
                var header = CreateGroupHeader(detailGroup, isSorted, regionScope);
                int nestedRow;
                double nestedLeft;
                var nested = FindNestedTable(detailRows, edges, dataSet, out nestedRow, out nestedLeft);

                if (nested == null)
                {
                    AddBand(header, container, AddRows(detailRows, edges, header, 0, regionScope));
                    AddBand(dataBand, container, 0);
                }
                else
                {
                    AddNestedTableBands(nested, nestedRow, nestedLeft, detailRows, edges, container, header, dataBand,
                        dataSet, regionScope);
                }
            }
            else
            {
                AddBand(dataBand, container, AddRows(detailRows, edges, dataBand, 0, regionScope));
            }

            for (var index = footers.Count - 1; index >= 0; index--)
            {
                AddBand(footers[index], container, AddRows(footerRows[index], edges, footers[index], 0, regionScope));
            }
        }

        /// <summary>
        /// Finds a table that stands alone in a row of a group and reads the data set of the group: it
        /// shows the rows of the current group. Returns null when there is no such table or several.
        /// </summary>
        private StiRdlxNode FindNestedTable(List<StiRdlxNode> rows, double[] edges, StiRdlxDataSet dataSet,
            out int rowIndex, out double left)
        {
            StiRdlxNode found = null;
            rowIndex = -1;
            left = 0;

            for (var index = 0; index < rows.Count; index++)
            {
                var column = 0;
                var count = 0;
                StiRdlxNode table = null;
                var tableLeft = 0d;

                foreach (var cell in rows[index].Items("TableCells"))
                {
                    var item = GetCellItem(cell);

                    if (item != null)
                        count++;

                    if (item != null && item.IsType("Table") && GetRows(item.Child("Details")).Count > 0)
                    {
                        var name = item.Get("DataSetName");

                        if (string.IsNullOrEmpty(name) || FindDataSet(name) == dataSet)
                        {
                            table = item;
                            tableLeft = GetEdge(edges, column);
                        }
                    }

                    column += Math.Max(1, GetInt(cell.Get("ColSpan"), 1));
                }

                if (table == null || count != 1)
                    continue;

                if (found != null)
                    return null;

                found = table;
                rowIndex = index;
                left = tableLeft;
            }

            return found;
        }

        /// <summary>
        /// A group whose rows hold a table of the rows of the group: the rows above the table and the header
        /// of the table go to the group header, the details of the table go to the data band, the footer
        /// of the table and the rows below it go to the group footer.
        /// </summary>
        private void AddNestedTableBands(StiRdlxNode nested, int nestedRow, double nestedLeft,
            List<StiRdlxNode> rows, double[] edges, StiContainer container, StiGroupHeaderBand header,
            StiDataBand dataBand, StiRdlxDataSet dataSet, StiRdlxScope regionScope)
        {
            var nestedEdges = GetEdges(nested.Items("TableColumns"), nestedLeft);

            var top = AddRows(rows.GetRange(0, nestedRow), edges, header, 0, regionScope);
            top = AddRows(GetRows(nested.Child("Header")), nestedEdges, header, top, regionScope);
            AddBand(header, container, top);

            // The nested table may have groups and a nested table of its own, it gives its bands the same way.
            ApplyFilters(nested.Items("Filters"), dataBand, regionScope);
            AddTableDataBands(nested, nestedEdges, container, dataBand, dataSet, regionScope);

            var footer = CreateGroupFooter();
            top = AddRows(GetRows(nested.Child("Footer")), nestedEdges, footer, 0, regionScope);
            top = AddRows(rows.GetRange(nestedRow + 1, rows.Count - nestedRow - 1), edges, footer, top, regionScope);
            AddBand(footer, container, top);
        }

        /// <summary>
        /// A table without data: all its rows are placed on one band printed once.
        /// </summary>
        private void AddStaticTable(StiRdlxNode table, double[] edges, StiContainer container, StiRdlxScope scope)
        {
            var band = CreateStaticBand(0);
            var top = AddRows(GetRows(table.Child("Header")), edges, band, 0, scope);

            foreach (var group in table.Items("TableGroups"))
            {
                top = AddRows(GetRows(group.Child("Header")), edges, band, top, scope);
            }

            top = AddRows(GetRows(table.Child("Details")), edges, band, top, scope);

            var groups = table.Items("TableGroups");

            for (var index = groups.Count - 1; index >= 0; index--)
            {
                top = AddRows(GetRows(groups[index].Child("Footer")), edges, band, top, scope);
            }

            top = AddRows(GetRows(table.Child("Footer")), edges, band, top, scope);
            AddBand(band, container, top);
        }
        #endregion

        #region Methods.List
        private void BuildList(StiRdlxNode list, StiContainer container, double left, StiRdlxScope scope)
        {
            var dataSet = GetRegionDataSet(list, scope);
            var items = list.Items("ReportItems");
            var group = list.Child("Group", "Grouping");
            var height = ToHi(list.Get("Height"));

            if (dataSet == null || IsCurrentRowRegion(dataSet, scope))
            {
                var band = CreateStaticBand(height);
                container.Components.Add(band);
                AddItems(items, band, -left, 0, scope);
                return;
            }

            if (GetGroupExpression(group) != null)
            {
                BuildGroupedItems(items, group, dataSet, container, scope, left, list, isNewPage: false);
                return;
            }

            var dataBand = CreateDataBand(list, dataSet, scope);
            ApplySorting(list, dataBand, dataSet);
            AddBand(dataBand, container, height);
            AddItems(items, dataBand, -left, 0, CreateRegionScope(dataSet, dataBand));
        }

        private void BuildGroupedItems(List<StiRdlxNode> items, StiRdlxNode group, StiRdlxDataSet dataSet,
            StiContainer container, StiRdlxScope scope, bool isNewPage)
        {
            BuildGroupedItems(items, group, dataSet, container, scope, 0, null, isNewPage);
        }

        /// <summary>
        /// Places the items of a grouped list or of a grouped fixed page. When one of the items is a table of
        /// the same data set, the group becomes the outer group of the table. Otherwise all the items go to
        /// the group header.
        /// </summary>
        private void BuildGroupedItems(List<StiRdlxNode> items, StiRdlxNode group, StiRdlxDataSet dataSet,
            StiContainer container, StiRdlxScope scope, double left, StiRdlxNode region, bool isNewPage)
        {
            var dataBand = CreateDataBand(region ?? new StiRdlxNode("Region"), dataSet, scope);
            var regionScope = CreateRegionScope(dataSet, dataBand);
            var isSorted = region != null && ApplySorting(region, dataBand, dataSet);

            var header = CreateGroupHeader(group, isSorted, regionScope);
            header.CanBreak = true;

            if (isNewPage)
                header.NewPageBefore = true;

            var table = FindJoinedTable(items, dataSet);
            if (table == null)
            {
                var bottom = 0d;

                foreach (var item in items)
                {
                    bottom = Math.Max(bottom, GetRect(item).Bottom);
                }

                AddBand(header, container, bottom);
                AddItems(items, header, -left, 0, regionScope);
                AddBand(dataBand, container, 0);
                return;
            }

            var tableRect = GetRect(table);
            var edges = GetEdges(table.Items("TableColumns"), left + tableRect.Left);
            var above = new List<StiRdlxNode>();
            var below = new List<StiRdlxNode>();

            foreach (var item in items)
            {
                if (item == table)
                    continue;

                if (GetRect(item).Top < tableRect.Top + Tolerance)
                    above.Add(item);

                else
                    below.Add(item);
            }

            ApplyFilters(table.Items("Filters"), dataBand, regionScope);

            AddItems(above, header, -left, 0, regionScope);
            var headerBottom = AddRows(GetRows(table.Child("Header")), edges, header, tableRect.Top, regionScope);
            AddBand(header, container, headerBottom);

            isPageBreakSuppressed = isNewPage;
            AddTableDataBands(table, edges, container, dataBand, dataSet, regionScope);
            isPageBreakSuppressed = false;

            var footer = CreateGroupFooter();
            var footerBottom = AddRows(GetRows(table.Child("Footer")), edges, footer, 0, regionScope);

            // A table of a fixed page keeps the place it may grow into (FixedHeight). The items below it
            // stand after this place and not after the rows drawn in the designer, so the place is not
            // copied into the footer as an empty gap.
            var tableBottom = tableRect.Top + Math.Max(tableRect.Height, ToHi(table.Get("FixedHeight")));
            var offset = tableBottom - footerBottom;

            // The items below the table of a fixed page are pinned to the bottom of the page, so the place
            // between the table and these items is not kept at all.
            if (isNewPage)
            {
                footer.PrintAtBottom = true;
                offset = below.Count > 0 ? double.MaxValue : 0;
            }

            foreach (var item in below)
            {
                offset = Math.Min(offset, GetRect(item).Top - footerBottom);
            }

            AddItems(below, footer, -left, offset, regionScope);

            foreach (var item in below)
            {
                footerBottom = Math.Max(footerBottom, GetRect(item).Bottom - offset);
            }

            AddBand(footer, container, footerBottom);
        }

        /// <summary>
        /// Finds the only table among the items that reads the same data set and has detail rows.
        /// </summary>
        private StiRdlxNode FindJoinedTable(List<StiRdlxNode> items, StiRdlxDataSet dataSet)
        {
            StiRdlxNode found = null;

            foreach (var item in items)
            {
                if (!item.IsType("Table") || GetRows(item.Child("Details")).Count == 0)
                    continue;

                var name = item.Get("DataSetName");
                if (!string.IsNullOrEmpty(name) && FindDataSet(name) != dataSet)
                    continue;

                if (found != null)
                    return null;

                found = item;
            }

            return found;
        }
        #endregion

        #region Methods.BandedList
        private void BuildBandedList(StiRdlxNode list, StiContainer container, double left, StiRdlxScope scope)
        {
            var dataSet = GetRegionDataSet(list, scope);
            var isStatic = dataSet == null || IsCurrentRowRegion(dataSet, scope);
            var dataBand = isStatic ? null : CreateDataBand(list, dataSet, scope);
            var regionScope = isStatic ? scope : CreateRegionScope(dataSet, dataBand);
            var footers = new List<StiRdlxNode>();

            AddListBand(list.Child("Header"), isStatic ? (StiBand)CreateStaticBand(0) : CreateHeaderBand(list),
                container, left, regionScope);

            foreach (var group in list.Items("Groups"))
            {
                var header = isStatic
                    ? (StiBand)CreateStaticBand(0)
                    : CreateGroupHeader(group.Child("Group", "Grouping"), false, regionScope);

                if (!AddListBand(group.Child("Header"), header, container, left, regionScope) && !isStatic)
                    AddBand(header, container, 0);

                footers.Add(group.Child("Footer"));
            }

            if (!AddListBand(list.Child("Details"), isStatic ? CreateStaticBand(0) : dataBand, container, left,
                regionScope) && !isStatic)
                AddBand(dataBand, container, 0);

            for (var index = footers.Count - 1; index >= 0; index--)
            {
                var footer = isStatic ? (StiBand)CreateStaticBand(0) : CreateGroupFooter();

                if (!AddListBand(footers[index], footer, container, left, regionScope) && !isStatic)
                    AddBand(footer, container, 0);
            }

            StiBand listFooter = new StiFooterBand();
            listFooter.Name = CreateName("FooterBand");
            AddListBand(list.Child("Footer"), isStatic ? CreateStaticBand(0) : listFooter, container, left,
                regionScope);
        }

        private StiHeaderBand CreateHeaderBand(StiRdlxNode list)
        {
            var header = new StiHeaderBand();
            header.Name = CreateName("HeaderBand");

            var node = list.Child("Header");
            header.PrintOnAllPages = node != null && node.GetBool("RepeatOnNewPage");

            return header;
        }

        /// <summary>
        /// Adds a band of a banded list with its items. Returns false when the list has no such band.
        /// </summary>
        private bool AddListBand(StiRdlxNode node, StiBand band, StiContainer container, double left,
            StiRdlxScope scope)
        {
            if (node == null)
                return false;

            AddBand(band, container, ToHi(node.Get("Height")));
            ApplyContainerStyle(node, band);
            AddItems(node.Items("ReportItems"), band, -left, 0, scope);

            return true;
        }
        #endregion
    }
}
