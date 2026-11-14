#region Copyright (C) 2003-2026 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports												}
{	                         										}
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
using Stimulsoft.Base.Helpers;
using Stimulsoft.Report.Components;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts the groups of a tablix of Reporting Services: the condition of a group with several
    /// expressions, the order of the groups and of the rows, and Count without the scope of a data set,
    /// which counts the rows of the group it is printed in.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts
        private const string CountScopeMarker = "RdlCountScope_";

        private static readonly Regex CountScopeRegex = new Regex(CountScopeMarker + @"(\d+)");

        private const string GroupScopeMarker = "RdlGroupScope_";

        private const string GroupScopeBand = "RdlGroupScopeBand";

        private static readonly Regex GroupScopeRegex = new Regex(GroupScopeMarker + @"(\d+)");

        private static readonly Regex SummarySortRegex = new Regex(
            @"^\{(?<wrap>System\.Convert\.ToDouble\()?" +
            @"Totals\.(Sum|Avg|MinDate|MaxDate|Min|Max|Count|CountDistinct)\((.*)\)(?(wrap)\))\}$");

        //an aggregate of one value over one row, with the data source as the scope or without it
        private static readonly Regex RowAggregateRegex = new Regex(
            @"^\{(System\.Convert\.ToDouble\()?Totals\.(Sum|Avg|MinDate|MaxDate|Min|Max|First|Last)\(" +
            @"(\w+, )?(?<value>[^(),]*(\([^()]*\))?[^(),]*)\)\)?\}$");
        #endregion

        #region Fields.Groups
        /// <summary>
        /// The data sources of Count written in a data region without the scope of a data set, by the
        /// number in the marker. The band to count is known only when the bands of the region are made.
        /// </summary>
        private List<string> countScopeSources = new List<string>();

        /// <summary>
        /// The aggregates with the scope of a group by the number in the marker: the name of the group, the
        /// function over its header with a placeholder of the band, and the function without the scope for
        /// a region without the group.
        /// </summary>
        private List<Tuple<string, string, string>> groupScopes = new List<Tuple<string, string, string>>();

        /// <summary>
        /// The names of the groups of the columns by the names of their cross headers.
        /// </summary>
        private Dictionary<string, string> columnGroupNames = new Dictionary<string, string>();
        private bool isInDataRegion;
        #endregion

        #region Groups
        /// <summary>
        /// Joins the expressions of a group into one condition. A group of Reporting Services starts a
        /// new instance when any of its expressions changes, a group of Stimulsoft has one condition.
        /// </summary>
        private string GetGroupCondition(List<string> expressions)
        {
            if (expressions.Count == 1)
                return expressions[0];

            var separator = ConvertSyntaxToCSharp ? " + (char)31 + " : " & ChrW(31) & ";

            return string.Join(separator, expressions.Select(e => $"System.Convert.ToString({e})"));
        }

        /// <summary>
        /// Adds the order of a group to the sorting of its data band, the group itself does not sort.
        /// Stimulsoft sorts by the conditions of the sorted groups first and by the band next, so the
        /// groups of one band are all sorted in the band, from the outer one. The expressions of a group
        /// always go to the sorting: Reporting Services gathers all rows of a group, Stimulsoft joins
        /// only the rows next to each other. A group sorted by its own first expression takes the
        /// direction; a group sorted by an aggregate is sorted by the summary of the same function; a group
        /// sorted by another number or date is sorted by the smallest or the largest value in the group,
        /// as its rows may differ in it.
        /// </summary>
        private void AddGroupSort(List<string> sort, StiGroupHeaderBand groupHeader, TablixMember member)
        {
            groupHeader.SortDirection = StiGroupSortDirection.None;

            var direction = "ASC";
            var sortExpressions = member.SortExpressions;
            if (sortExpressions != null && sortExpressions.Count >= 2)
            {
                if (sortExpressions[1] == "{" + member.GroupExpressions[0] + "}")
                    direction = sortExpressions[0];

                else
                    SortGroupBySummary(groupHeader, member, sortExpressions[0], sortExpressions[1]);

                if (sortExpressions.Count > 2)
                    ThrowError(null, null, $"Only the first sort expression of a group is used: {member.GroupName}");
            }
            else
            {
                ReportUnsortedGroup(member);
            }

            for (var index = 0; index < member.GroupExpressions.Count; index++)
            {
                sort.Add(index == 0 ? direction : "ASC");
                sort.Add("{" + member.GroupExpressions[index] + "}");
            }
        }

        private void SortGroupBySummary(StiGroupHeaderBand groupHeader, TablixMember member, string direction,
            string expression)
        {
            var isAscending = direction == "ASC";
            if (SortGroupByAggregate(groupHeader, expression, isAscending))
                return;

            var type = GetExpressionType(expression.Substring(1, expression.Length - 2));

            if (type != null && type.IsNumericType())
                groupHeader.SummaryType = isAscending ? StiGroupSummaryType.Min : StiGroupSummaryType.Max;

            else if (type == typeof(DateTime) || type == typeof(DateTime?))
                groupHeader.SummaryType = isAscending ? StiGroupSummaryType.MinDate : StiGroupSummaryType.MaxDate;

            else
            {
                ThrowError(null, null, $"Group sorted by another text keeps its own order: {member.GroupName}");
                return;
            }

            groupHeader.SummaryExpression = new StiGroupSummaryExpression(expression);
            groupHeader.SummarySortDirection = isAscending
                ? StiGroupSortDirection.Ascending
                : StiGroupSortDirection.Descending;
        }

        /// <summary>
        /// Sorts a group by an aggregate of its rows, as Sum(Fields!Amount.Value), with the summary of the
        /// same function. Count needs no expression, its argument is the band to count.
        /// </summary>
        private static bool SortGroupByAggregate(StiGroupHeaderBand groupHeader, string expression, bool isAscending)
        {
            var match = SummarySortRegex.Match(expression);
            if (!match.Success)
                return false;

            var function = match.Groups[1].Value;
            groupHeader.SummaryType = (StiGroupSummaryType)Enum.Parse(typeof(StiGroupSummaryType), function);

            if (groupHeader.SummaryType != StiGroupSummaryType.Count)
                groupHeader.SummaryExpression = new StiGroupSummaryExpression("{" + match.Groups[2].Value + "}");

            groupHeader.SummarySortDirection = isAscending
                ? StiGroupSortDirection.Ascending
                : StiGroupSortDirection.Descending;

            return true;
        }

        /// <summary>
        /// Sets the condition and the direction of a group of the columns. The columns are not sorted
        /// by their band, so the group sorts by its condition, descending when its own first expression
        /// is sorted so.
        /// </summary>
        private void ApplyColumnGroup(StiGroupHeaderBand groupHeader, TablixMember member, string dataSetName,
            string copyName)
        {
            var condition = GetGroupCondition(member.GroupExpressions);
            if (dataSetName != copyName)
                condition = condition.Replace(dataSetName, copyName);

            groupHeader.Condition = new StiGroupConditionExpression("{" + condition + "}");

            var sortExpressions = member.SortExpressions;
            var isDescending = sortExpressions != null && sortExpressions.Count >= 2 &&
                sortExpressions[0] == "DESC" && sortExpressions[1] == "{" + member.GroupExpressions[0] + "}";

            if (isDescending)
                groupHeader.SortDirection = StiGroupSortDirection.Descending;

            if (sortExpressions == null || sortExpressions.Count == 0)
                ReportUnsortedGroup(member);
        }

        /// <summary>
        /// Returns the sorting of the rows of a data band, the pairs of the direction and the expression. An
        /// aggregate over one row is the value of the row itself, so First(Fields!Name.Value) of the details
        /// sorts by the name; the function of Totals in the sorting of a band does not compile. Count of one
        /// row is always 1 and another aggregate is not known, such a key is skipped and named in the errors.
        /// </summary>
        private IEnumerable<string> GetRowSort(IList<string> sort)
        {
            for (var index = 0; index + 1 < sort.Count; index += 2)
            {
                var expression = sort[index + 1];
                if (!expression.Contains("Totals."))
                {
                    yield return sort[index];
                    yield return expression;
                    continue;
                }

                var match = RowAggregateRegex.Match(expression);
                if (match.Success)
                {
                    yield return sort[index];
                    yield return "{" + match.Groups["value"].Value + "}";
                    continue;
                }

                ThrowError(null, null, $"Sorting of the rows by an aggregate not imported: {expression}");
            }
        }

        /// <summary>
        /// Reporting Services prints the groups without the sorting in the order they first appear in
        /// the data, Stimulsoft joins only the rows next to each other, so such a group is sorted
        /// ascending by its expressions.
        /// </summary>
        private void ReportUnsortedGroup(TablixMember member)
        {
            ThrowError(null, null, $"Group without sorting is ordered ascending: {member.GroupName}");
        }
        #endregion

        #region Groups.Count
        /// <summary>
        /// Writes Count without the scope of a data set in a data region with a marker in place of the
        /// band. Reporting Services counts the rows of the current group, Totals.Count of the data source
        /// counts all its rows.
        /// </summary>
        private string MakeCountInScope(string dataSourceReference)
        {
            countScopeSources.Add(dataSourceReference);

            return $"Totals.Count({CountScopeMarker}{countScopeSources.Count - 1})";
        }

        /// <summary>
        /// Puts the band to count into Count of the components of a data region: the group header for
        /// its header and footer, the data band for the details, the header and the footer of the
        /// region. The other components of the region count the data source.
        /// </summary>
        private void ResolveCountScopes(Dictionary<StiBand, StiBand> scopes, StiContainer region)
        {
            foreach (var pair in scopes)
            {
                foreach (StiComponent component in pair.Key.GetComponents())
                {
                    ReplaceScopes(component, CountScopeRegex, m => pair.Value.Name);
                }
            }

            foreach (StiComponent component in region.GetComponents())
            {
                ReplaceScopes(component, CountScopeRegex, m => countScopeSources[GetScopeIndex(m)]);
            }
        }

        /// <summary>
        /// Writes an aggregate with the scope of a group with a marker in place of the whole function. Reporting
        /// Services gathers the rows of the named group, the band of the component may be an inner group.
        /// </summary>
        private string MakeGroupScope(string groupName, string totalOfGroup, string total)
        {
            groupScopes.Add(Tuple.Create(groupName, totalOfGroup, total));

            return GroupScopeMarker + (groupScopes.Count - 1);
        }

        /// <summary>
        /// Puts the aggregates with the scope of a group or of the region into the components of a data region:
        /// over the header of a group of the rows, over the data band for the region, over the header of a group
        /// of the columns in the band of the row of the component, which runs over the copy of the data source;
        /// without the scope when the region has no such group. Goes before ResolveMatrixCells, which converts
        /// only the functions without a band, and before ResolveCountScopes, Count without the group has a
        /// marker of its own.
        /// </summary>
        private void ResolveGroupScopes(StiContainer region, Dictionary<string, StiBand> groupBands,
            string dataSetName, string copyName)
        {
            foreach (StiComponent component in region.GetComponents())
            {
                ReplaceScopes(component, GroupScopeRegex, m => GetGroupScope(groupScopes[GetScopeIndex(m)],
                    component, region, groupBands, dataSetName, copyName));
            }
        }

        private string GetGroupScope(Tuple<string, string, string> scope, StiComponent component, StiContainer region,
            Dictionary<string, StiBand> groupBands, string dataSetName, string copyName)
        {
            if (groupBands.TryGetValue(scope.Item1, out var band))
                return scope.Item2.Replace(GroupScopeBand, band.Name);

            var columnBand = copyName != null ? FindColumnGroupHeader(component, region, scope.Item1) : null;
            if (columnBand == null)
                return scope.Item3;

            var total = scope.Item2.Replace(GroupScopeBand, columnBand.Name);

            return ToCopyReference(total, GetDataSourceReference(dataSetName), GetDataSourceReference(copyName));
        }

        /// <summary>
        /// Finds the header of the named group of the columns in the band of the row of the component, every
        /// band of the rows has the cross bands of the columns of its own.
        /// </summary>
        private StiBand FindColumnGroupHeader(StiComponent component, StiContainer region, string groupName)
        {
            var rowBand = component;
            while (rowBand != null && rowBand.Parent != region)
            {
                rowBand = rowBand.Parent;
            }

            return (rowBand as StiContainer)?.GetComponents().Cast<StiComponent>()
                .OfType<StiCrossGroupHeaderBand>()
                .FirstOrDefault(b => columnGroupNames.TryGetValue(b.Name, out var name) &&
                    string.Equals(name, groupName, StringComparison.OrdinalIgnoreCase));
        }

        private static void ReplaceScopes(StiComponent component, Regex regex, MatchEvaluator getScope)
        {
            if (component is StiText text)
                text.Text.Value = ReplaceScopes(text.Text.Value, regex, getScope);

            if (component is StiImage image)
                image.ImageURL.Value = ReplaceScopes(image.ImageURL.Value, regex, getScope);

            foreach (StiAppExpression expression in component.Expressions)
            {
                expression.Expression = ReplaceScopes(expression.Expression, regex, getScope);
            }

            component.BeforePrintEvent.Script = ReplaceScopes(component.BeforePrintEvent.Script, regex, getScope);
            component.Bookmark.Value = ReplaceScopes(component.Bookmark.Value, regex, getScope);
            component.ToolTip.Value = ReplaceScopes(component.ToolTip.Value, regex, getScope);
            component.Hyperlink.Value = ReplaceScopes(component.Hyperlink.Value, regex, getScope);
        }

        private static string ReplaceScopes(string text, Regex regex, MatchEvaluator getScope)
        {
            if (string.IsNullOrEmpty(text) || !regex.IsMatch(text))
                return text;

            return regex.Replace(text, getScope);
        }

        private static int GetScopeIndex(Match match)
        {
            return int.Parse(match.Groups[1].Value);
        }
        #endregion
    }
}
