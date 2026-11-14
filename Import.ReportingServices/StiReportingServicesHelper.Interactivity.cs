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

using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts the interactivity of a report item of Reporting Services: the label in the document
    /// map, the bookmark, the tooltip, the action and the sorting by the user.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts
        private const int MaxDrillDownParameters = 10;
        #endregion

        #region Fields.Interactivity
        /// <summary>
        /// The text boxes with UserSort and the field they sort by. The data band that is sorted is
        /// made after the text box, so the column is set when the data region is built.
        /// </summary>
        private List<Tuple<StiComponent, string>> pendingUserSorts = new List<Tuple<StiComponent, string>>();
        #endregion

        #region Interactivity
        /// <summary>
        /// Converts a node of interactivity of a report item. Returns false for any other node, so
        /// the caller reports it as not supported.
        /// </summary>
        private bool ProcessInteractivityNode(XmlNode node, StiComponent component, string dataset)
        {
            switch (node.Name)
            {
                case "DocumentMapLabel":
                    component.Bookmark = new StiBookmarkExpression(ConvertInteractiveValue(node, dataset));
                    return true;

                case "Bookmark":
                    //the label of the document map is the bookmark of Stimulsoft as well, the label wins
                    if (string.IsNullOrEmpty(component.Bookmark?.Value))
                        component.Bookmark = new StiBookmarkExpression(ConvertInteractiveValue(node, dataset));

                    return true;

                case "ToolTip":
                    component.ToolTip = new StiToolTipExpression(ConvertInteractiveValue(node, dataset));
                    return true;

                case "ActionInfo":
                    ProcessActionInfo(node, component, dataset);
                    return true;

                case "UserSort":
                    ProcessUserSort(node, component);
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>
        /// A value of interactivity is a text with expressions in braces in Stimulsoft.
        /// </summary>
        private string ConvertInteractiveValue(XmlNode node, string dataset)
        {
            var value = GetNodeTextValue(node);

            return IsExpression(value) ? ConvertExpression(value, dataset) : value;
        }

        /// <summary>
        /// Converts the action of a report item. A hyperlink goes to the hyperlink of the component,
        /// a link to a bookmark is a hyperlink that starts with the number sign, a drillthrough opens
        /// the other report with its parameters. Only the first action is taken, a component of
        /// Stimulsoft has one action.
        /// </summary>
        private void ProcessActionInfo(XmlNode baseNode, StiComponent component, string dataset)
        {
            var actions = baseNode.SelectNodes("*[local-name()='Actions']/*[local-name()='Action']")
                .Cast<XmlNode>()
                .ToList();

            if (actions.Count == 0)
                actions = baseNode.SelectNodes("*[local-name()='Action']").Cast<XmlNode>().ToList();

            if (actions.Count > 1)
                ThrowError(null, null, $"Only the first action is imported: {component.Name}");

            if (actions.Count == 0)
                return;

            foreach (XmlNode node in actions[0].ChildNodes)
            {
                switch (node.Name)
                {
                    case "Hyperlink":
                        component.Hyperlink = new StiHyperlinkExpression(ConvertInteractiveValue(node, dataset));
                        break;

                    case "BookmarkLink":
                        component.Hyperlink = new StiHyperlinkExpression("#" + ConvertInteractiveValue(node, dataset));
                        break;

                    case "Drillthrough":
                        ProcessDrillthrough(node, component, dataset);
                        break;

                    default:
                        ThrowError(actions[0].Name, node.Name);
                        break;
                }
            }
        }

        /// <summary>
        /// A drillthrough opens the other report of the server with the given parameters. The report
        /// is taken from the file with the same name next to this one, it has to be imported too. A
        /// report given by an expression is not known before the report runs.
        /// </summary>
        private void ProcessDrillthrough(XmlNode baseNode, StiComponent component, string dataset)
        {
            var interaction = component.Interaction;
            if (interaction == null)
            {
                ThrowError(null, null, $"Drillthrough not supported by the item: {component.Name}");
                return;
            }

            var reportNode = baseNode.SelectSingleNode("*[local-name()='ReportName']");
            var reportName = GetConstantText(reportNode != null ? GetNodeTextValue(reportNode) : null);

            if (string.IsNullOrWhiteSpace(reportName) || IsExpression(reportName))
            {
                ThrowError(null, null,
                    $"Drillthrough to a report given by an expression not imported: {component.Name}");
                return;
            }

            interaction.DrillDownEnabled = true;
            interaction.DrillDownReport = reportName.Split('/', '\\').Last() + ".mrt";
            ThrowError(null, null, $"Drillthrough report has to be imported too: {reportName}");

            var parameters = baseNode.SelectNodes("*[local-name()='Parameters']/*[local-name()='Parameter']")
                .Cast<XmlNode>()
                .ToList();

            if (parameters.Count > MaxDrillDownParameters)
                ThrowError(null, null,
                    $"Only {MaxDrillDownParameters} drillthrough parameters are imported: {component.Name}");

            for (var index = 0; index < parameters.Count && index < MaxDrillDownParameters; index++)
            {
                var parameter = parameters[index];
                var valueNode = parameter.SelectSingleNode("*[local-name()='Value']");
                var text = valueNode != null ? GetNodeTextValue(valueNode) ?? string.Empty : string.Empty;
                var value = valueNode != null ? ConvertInteractiveValue(valueNode, dataset) : string.Empty;

                //the value of a parameter is code, a text is a string of the language of the report
                if (text.Length == 0)
                    value = "\"\"";

                else if (!IsExpression(text))
                    value = ConvertSyntaxToCSharp ? ToCSharpString(text) : ToVbString(text);

                else if (IsSingleExpression(value))
                    value = value.Substring(1, value.Length - 2);

                var drillDownParameter = new StiDrillDownParameter(parameter.Attributes?["Name"]?.Value)
                {
                    Expression = new StiExpression(value)
                };

                typeof(StiInteraction).GetProperty($"DrillDownParameter{index + 1}")
                    .SetValue(interaction, drillDownParameter);
            }
        }

        /// <summary>
        /// Returns the text of an expression that is only a string, ="Report" gives Report; any other value as is.
        /// </summary>
        private static string GetConstantText(string value)
        {
            var match = Regex.Match(value ?? string.Empty, @"^\s*=\s*""(?<text>[^""]*)""\s*$");

            return match.Success ? match.Groups["text"].Value : value;
        }

        /// <summary>
        /// A text box sorts the data by a click when its interaction has the column to sort by. Only
        /// a field can be the column, an expression has no counterpart.
        /// </summary>
        private void ProcessUserSort(XmlNode baseNode, StiComponent component)
        {
            var expressionNode = baseNode.SelectSingleNode("*[local-name()='SortExpression']");
            var expression = expressionNode != null ? GetNodeTextValue(expressionNode) : null;
            var match = Regex.Match(expression ?? string.Empty, @"^\s*=\s*Fields!(\w+)\.Value\s*$",
                RegexOptions.IgnoreCase);

            if (!match.Success || component.Interaction == null)
            {
                ThrowError(null, null, $"UserSort by an expression not imported: {component.Name}");
                return;
            }

            pendingUserSorts.Add(Tuple.Create(component, match.Groups[1].Value));
        }

        /// <summary>
        /// Sets the column to sort by for the text boxes with UserSort inside the data region, now
        /// that its data band has a name.
        /// </summary>
        private void ResolveUserSorts(StiContainer region, StiDataBand dataBand)
        {
            if (dataBand == null)
                return;

            var components = region.GetComponents().Cast<StiComponent>().ToList();

            foreach (var sort in pendingUserSorts.Where(s => components.Contains(s.Item1)).ToList())
            {
                sort.Item1.Interaction.SortingEnabled = true;
                sort.Item1.Interaction.SortingColumn = $"{dataBand.Name}.{sort.Item2}";
                pendingUserSorts.Remove(sort);
            }
        }

        /// <summary>
        /// The text boxes with UserSort out of a data region that could be sorted are named in the
        /// list of errors.
        /// </summary>
        private void ReportUnresolvedUserSorts()
        {
            foreach (var sort in pendingUserSorts)
            {
                ThrowError(null, null, $"UserSort out of a table not imported: {sort.Item1.Name}");
            }

            pendingUserSorts.Clear();
        }
        #endregion
    }
}
