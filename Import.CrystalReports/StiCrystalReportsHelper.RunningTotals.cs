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
using System.Linq;
using System.Text.RegularExpressions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Import.CrystalReports.Formulas;
using Stimulsoft.Report.Components;

namespace Import.CrystalReports
{
    /// <summary>
    /// The running total fields of Crystal become the running aggregate functions of Stimulsoft: SumRunning,
    /// CountIfRunning and others. The function goes over the data band, or over the band of the group which resets
    /// the total, and a formula of the evaluation becomes the condition of the If variant of the function.
    /// </summary>
    public sealed partial class StiCrystalReportsHelper
    {
        #region Fields.RunningTotals
        private Dictionary<string, RunningTotalFieldDefinition> runningTotals;
        private Dictionary<string, StiCrystalRunningTotal> runningTotalTranslations;
        private Dictionary<string, string> runningTotalVariables;
        #endregion

        #region Methods.RunningTotals
        /// <summary>
        /// Keeps the definition of a running total until the bands are known; the name is the content of the braces,
        /// for example #RTotal0, the variable is the name which a text object uses for the field.
        /// </summary>
        private void AddRunningTotal(string name, RunningTotalFieldDefinition definition, string variable)
        {
            if (runningTotals == null)
            {
                runningTotals = new Dictionary<string, RunningTotalFieldDefinition>(StringComparer.OrdinalIgnoreCase);
                runningTotalTranslations =
                    new Dictionary<string, StiCrystalRunningTotal>(StringComparer.OrdinalIgnoreCase);
                runningTotalVariables = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            }

            runningTotals[name] = definition;
            runningTotalVariables[name] = variable;
        }

        /// <summary>
        /// A text object shows its fields by the names of the variables; a running total and a formula with summaries
        /// are calculated only in a text, so their expressions replace the variables.
        /// </summary>
        private void AddEmbeddedFieldsText(StiText text)
        {
            var value = text.Text?.Value;
            if (string.IsNullOrEmpty(value) || value.IndexOf('{') < 0)
                return;

            pendingTranslations.Add(() =>
            {
                var result = text.Text.Value;
                if (runningTotalTranslations != null)
                {
                    foreach (var pair in runningTotalTranslations)
                    {
                        result = result.Replace("{" + runningTotalVariables[pair.Key] + "}",
                            "{" + GetRunningTotalCode(pair.Value, text) + "}");
                    }
                }

                foreach (var formula in formulaFields.Values.Where(f => f.InlineExpression != null))
                {
                    result = result.Replace("{" + formula.Variable.Name + "}", "{" + formula.InlineExpression + "}");
                }

                text.Text = result;
            });
        }

        /// <summary>
        /// Translates the running totals before the formulas: a formula which uses a running total gets its expression
        /// with the aggregate function, so it is shown by the text itself, like a formula with summaries.
        /// </summary>
        private void TranslateRunningTotals()
        {
            if (runningTotals == null)
                return;

            foreach (var pair in runningTotals)
            {
                var total = TranslateRunningTotal(pair.Key, pair.Value);
                if (total == null)
                    continue;

                runningTotalTranslations[pair.Key] = total;
                var expression = new StiCrystalExpression(total.Code, total.Type, isAtomic: true);

                var reference = resolver.GetReference(pair.Key);
                if (reference == null)
                {
                    reference = new StiCrystalReference(expression.Code, expression.Type);
                    resolver.AddReference(pair.Key, reference);
                }

                reference.Expression = expression.Code;
                reference.Type = expression.Type;
                reference.HasSummaries = true;
            }
        }

        private StiCrystalRunningTotal TranslateRunningTotal(string name, RunningTotalFieldDefinition definition)
        {
            var function = GetRunningFunction(definition.Operation);
            if (function == null)
            {
                warningList.Add(StiWarning.GetWarningFormula(name,
                    $"The operation {definition.Operation} of the running total is not supported"));
                return null;
            }

            var value = TranslateFormula(definition.SummarizedField?.FormulaName, name);
            if (value == null)
                return null;

            var band = GetRunningTotalBand(name, definition);
            if (band == null)
                return null;

            var type = function == "Count" || function == "CountDistinct" ? typeof(int) : typeof(decimal);
            if (function == "Min" || function == "Max")
            {
                if (value.Type == typeof(DateTime))
                {
                    type = typeof(DateTime);
                    function += "Date";
                }

                else if (value.Type == typeof(string))
                {
                    type = typeof(string);
                    function += "Str";
                }
            }

            string condition;
            if (!TryGetRunningTotalCondition(name, definition, out condition))
                return null;

            return new StiCrystalRunningTotal(function, band, value.Code, condition, type);
        }

        /// <summary>
        /// The expression of the running total in the text. In a group header Crystal shows the value at the first
        /// row of the group; it is calculated when the total is reset outside the group: never, by the same group or
        /// by an outer one.
        /// </summary>
        private static string GetRunningTotalCode(StiCrystalRunningTotal total, StiText text)
        {
            var band = text.Parent as StiBand;
            var child = band as StiChildBand;
            if (child != null)
                band = child.GetMaster();

            var header = band as StiGroupHeaderBand;
            if (header == null || !total.CanBeShownInHeader)
                return total.Code;

            var components = header.Parent.Components;
            var resetBand = components[total.Band];
            var resetOutside = resetBand is StiDataBand ||
                resetBand is StiGroupHeaderBand && components.IndexOf(resetBand) <= components.IndexOf(header);

            return resetOutside ? total.GetHeaderCode(header.Name) : total.Code;
        }

        private static bool IsRunningExpression(string code)
        {
            return Regex.IsMatch(code, @"\b\w+Running\(");
        }

        private static string GetRunningFunction(SummaryOperation operation)
        {
            switch (operation)
            {
                case SummaryOperation.Sum:
                    return "Sum";

                case SummaryOperation.Average:
                    return "Avg";

                case SummaryOperation.Maximum:
                    return "Max";

                case SummaryOperation.Minimum:
                    return "Min";

                case SummaryOperation.Count:
                    return "Count";

                case SummaryOperation.DistinctCount:
                    return "CountDistinct";

                default:
                    return null;
            }
        }

        /// <summary>
        /// The band of the function: the band of the group which resets the total, or the data band for the total
        /// which is never reset. A reset by a field or by a formula has no analog, the total is not reset then.
        /// </summary>
        private string GetRunningTotalBand(string name, RunningTotalFieldDefinition definition)
        {
            if (definition.ResetConditionType == RunningTotalCondition.OnChangeOfGroup)
            {
                var group = definition.ResetCondition as CrystalDecisions.CrystalReports.Engine.Group ??
                    definition.Group;
                var field = group?.ConditionField?.FormulaName?.Trim('{', '}');
                var groupBand = field != null ? resolver.ResolveGroupBand(field, GetGroupPeriod(group)) : null;
                if (groupBand != null)
                    return groupBand;

                warningList.Add(StiWarning.GetWarningFormula(name,
                    "The group which resets the running total is unknown"));
                return null;
            }

            if (definition.ResetConditionType != RunningTotalCondition.NoCondition)
            {
                warningList.Add(StiWarning.GetWarningFormula(name,
                    $"The reset {definition.ResetConditionType} of the running total is not supported, " +
                    "it is never reset"));
            }

            var dataBand = formulaPage?.GetComponents().OfType<StiDataBand>().FirstOrDefault();
            if (dataBand != null)
                return dataBand.Name;

            warningList.Add(StiWarning.GetWarningFormula(name, "The report has no data band for the running total"));
            return null;
        }

        /// <summary>
        /// The condition of the evaluation, null if the total takes every row. A formula which can't be translated
        /// fails the running total: the total without its condition would show wrong values. A change of a field or
        /// of a group has no analog, the total takes every row then.
        /// </summary>
        private bool TryGetRunningTotalCondition(string name, RunningTotalFieldDefinition definition,
            out string condition)
        {
            condition = null;
            switch (definition.EvaluationConditionType)
            {
                case RunningTotalCondition.NoCondition:
                    return true;

                case RunningTotalCondition.OnFormula:
                    var expression = TranslateFormula(definition.EvaluationCondition as string, name);
                    if (expression == null)
                        return false;

                    if (expression.Type != typeof(bool))
                    {
                        warningList.Add(StiWarning.GetWarningFormula(name, "The evaluation formula is not a condition"));
                        return false;
                    }

                    condition = expression.Code;
                    return true;

                default:
                    warningList.Add(StiWarning.GetWarningFormula(name,
                        $"The evaluation {definition.EvaluationConditionType} is not supported, every row is taken"));
                    return true;
            }
        }

        /// <summary>
        /// A field object of a running total shows the aggregate function itself: in a variable it wouldn't be running.
        /// </summary>
        private void AddRunningTotalText(StiText text, string name)
        {
            pendingTranslations.Add(() =>
            {
                StiCrystalRunningTotal total;
                if (runningTotalTranslations != null && runningTotalTranslations.TryGetValue(name, out total))
                    text.Text = "{" + GetRunningTotalCode(total, text) + "}";
            });
        }
        #endregion
    }
}
