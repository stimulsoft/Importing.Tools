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

using Stimulsoft.Report.Components;
using Stimulsoft.Report.Import.ReportingServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts the functions of Reporting Services that depend on the rows printed before: RunningValue,
    /// RowNumber and Previous.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts
        private const string RunningScopeMarker = "RdlRunningScope_";

        private static readonly Regex RunningScopeRegex = new Regex(RunningScopeMarker + @"(\d+)");

        private static readonly string[] RunningFunctions = { "Sum", "Avg", "Count", "CountDistinct", "Min", "Max" };
        #endregion

        #region Fields.Running
        /// <summary>
        /// The scopes of RunningValue by the number in the marker: the name of a group, or null for the
        /// data region. The band is known only when the bands of the region are made.
        /// </summary>
        private List<string> runningScopes = new List<string>();

        private HashSet<string> groupNames;

        private HashSet<string> regionNames;
        #endregion

        #region Running
        /// <summary>
        /// Converts RunningValue to a running function of Stimulsoft, Sum to SumRunning and so on. The
        /// running function goes over the band given first: the header of a group restarts it on each
        /// group, the data band goes over all the rows of the region.
        /// </summary>
        private int ConvertRunningValue(List<StiVbToken> tokens, int index, int open, int end, string dataset,
            StringBuilder result)
        {
            var close = FindClosingParenthesis(tokens, open, end);
            if (close == -1)
            {
                result.Append(tokens[index].Text);
                return index + 1;
            }

            var arguments = SplitArguments(tokens, open + 1, close);
            var function = arguments.Count >= 2
                ? FindName(RunningFunctions, GetArgumentName(tokens, arguments[1]))
                : null;

            if (function == null || !isInDataRegion)
            {
                ThrowError(null, null, "Function not supported: RunningValue");
                result.Append(NothingKeyword);
                return close + 1;
            }

            var argument = ConvertTokens(tokens, arguments[0].Item1, arguments[0].Item2, dataset, isCode: false).Trim();

            //Min and Max of Stimulsoft give a number, the dates have functions of their own
            if ((function == "Min" || function == "Max") && IsDateExpression(argument))
                function += "Date";

            runningScopes.Add(GetScope(tokens, arguments, 2));
            var band = RunningScopeMarker + (runningScopes.Count - 1);

            var call = function == "Count" ? $"CountRunning({band})" : $"{function}Running({band}, {argument})";

            //Avg of Reporting Services is Double for any type but Decimal
            if (function == "Avg" && !IsDecimalExpression(argument))
                call = $"System.Convert.ToDouble({call})";

            result.Append(call);

            return close + 1;
        }

        /// <summary>
        /// Converts RowNumber: the number of the row in a group is Line, which starts again in each group,
        /// the number in the data region or the data set is LineThrough.
        /// </summary>
        private int ConvertRowNumber(List<StiVbToken> tokens, int index, int open, int end, StringBuilder result)
        {
            var close = FindClosingParenthesis(tokens, open, end);
            if (close == -1)
            {
                result.Append(tokens[index].Text);
                return index + 1;
            }

            var scope = GetScope(tokens, SplitArguments(tokens, open + 1, close), 0);
            result.Append(scope != null && IsGroupName(scope) ? "Line" : "LineThrough");

            return close + 1;
        }

        /// <summary>
        /// Converts Previous of a field to the function Previous of Stimulsoft, which reads the column in
        /// the row before the current one of the data source. Previous of another expression has no
        /// counterpart and is named in the list of errors.
        /// </summary>
        private int ConvertPrevious(List<StiVbToken> tokens, int index, int open, int end, string dataset,
            StringBuilder result)
        {
            var close = FindClosingParenthesis(tokens, open, end);
            if (close == -1)
            {
                result.Append(tokens[index].Text);
                return index + 1;
            }

            var arguments = SplitArguments(tokens, open + 1, close);
            var column = arguments.Count > 0 ? GetSingleFieldName(tokens, arguments[0]) : null;
            var dataSource = column != null ? FindDataSetOfFields(tokens, arguments, dataset) : null;

            if (dataSource == null)
            {
                ThrowError(null, null, "Function not supported: Previous");
                result.Append(NothingKeyword);
                return close + 1;
            }

            result.Append($"Previous({GetDataSourceReference(dataSource)}, \"{column}\")");

            return close + 1;
        }

        /// <summary>
        /// Returns the name of the field when the argument is Fields!Name.Value and nothing else.
        /// </summary>
        private static string GetSingleFieldName(List<StiVbToken> tokens, Tuple<int, int> argument)
        {
            var start = GetNextSignificant(tokens, argument.Item1, argument.Item2);
            if (start >= argument.Item2 || !IsCollection(tokens, start, argument.Item2) ||
                !tokens[start].IsName("Fields"))
                return null;

            //Fields, !, the name, the dot and Value
            var after = GetNextSignificant(tokens, start + 5, argument.Item2);
            if (after < argument.Item2 || start + 4 >= argument.Item2 || !tokens[start + 4].IsName("Value"))
                return null;

            return tokens[start + 2].Name;
        }

        private static string GetArgumentName(List<StiVbToken> tokens, Tuple<int, int> argument)
        {
            var start = GetNextSignificant(tokens, argument.Item1, argument.Item2);

            return start < argument.Item2 ? tokens[start].Name : null;
        }

        private bool IsGroupName(string name)
        {
            if (groupNames == null)
            {
                groupNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (XmlNode node in reportNode.SelectNodes(".//*[local-name()='Group']"))
                {
                    var groupName = node.Attributes?["Name"]?.Value;
                    if (groupName != null)
                        groupNames.Add(groupName);
                }
            }

            return groupNames.Contains(name);
        }

        /// <summary>
        /// Returns true for the name of a group or of a tablix, the scopes an aggregate can name in a data region
        /// besides a data set.
        /// </summary>
        private bool IsGroupOrRegionName(string name)
        {
            if (IsGroupName(name))
                return true;

            if (regionNames == null)
            {
                regionNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (XmlNode node in reportNode.SelectNodes(".//*[local-name()='Tablix']"))
                {
                    var regionName = node.Attributes?["Name"]?.Value;
                    if (regionName != null)
                        regionNames.Add(regionName);
                }
            }

            return regionNames.Contains(name);
        }

        /// <summary>
        /// Puts the bands into the running functions of a data region: the header of the named group, or
        /// the data band for the data region, a data set or a group the region does not have.
        /// </summary>
        private void ResolveRunningScopes(StiContainer region, Dictionary<string, StiBand> groupBands,
            StiBand dataBand)
        {
            if (dataBand == null)
                dataBand = region.GetComponents().Cast<StiComponent>().OfType<StiDataBand>().FirstOrDefault();

            foreach (StiComponent component in region.GetComponents())
            {
                ReplaceScopes(component, RunningScopeRegex, m =>
                {
                    var scope = runningScopes[GetScopeIndex(m)];
                    var band = scope != null && groupBands.TryGetValue(scope, out var groupBand)
                        ? groupBand
                        : dataBand;

                    return band?.Name ?? NothingKeyword;
                });
            }
        }
        #endregion
    }
}
