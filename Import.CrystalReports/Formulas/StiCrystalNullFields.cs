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

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// Crystal stops a formula on a NULL value of a field of a database and the formula gives NULL, which is
    /// printed as nothing. Stimulsoft converts NULL to 0 or to the empty string and calculates on. This class finds
    /// the fields which a formula reads in any case: the condition of If, the subject of Select, the operands of the
    /// operators and the arguments of the functions; the branches are not taken, they may be skipped. A field which
    /// the formula checks by IsNull is not taken either: the formula expects NULL there.
    /// </summary>
    public static class StiCrystalNullFields
    {
        #region Methods
        /// <summary>
        /// In a condition NULL means false, so both operands of and are taken: when the left one is false the
        /// condition is false anyway.
        /// </summary>
        public static List<StiCrystalReference> Collect(StiCrystalNode node, IStiCrystalResolver resolver,
            bool isCondition = false)
        {
            var checkedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectCheckedFields(node, resolver, checkedFields);

            var fields = new List<StiCrystalReference>();
            CollectFields(node, resolver, fields, isCondition);

            return fields
                .Where(f => !checkedFields.Contains(GetKey(f)))
                .GroupBy(GetKey, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
        }

        private static void CollectFields(StiCrystalNode node, IStiCrystalResolver resolver,
            List<StiCrystalReference> fields, bool isCondition)
        {
            switch (node.Kind)
            {
                case StiCrystalNodeKind.Reference:
                    AddReferenceFields(node, resolver, fields);
                    break;

                case StiCrystalNodeKind.Call:
                    // IIF reads only its condition in any case
                    var isIif = string.Equals(node.Text, "iif", StringComparison.OrdinalIgnoreCase);
                    if (!IsNullSafeFunction(node.Text))
                        CollectChildren(node, resolver, fields, isIif ? 1 : int.MaxValue, isCondition);

                    break;

                case StiCrystalNodeKind.Binary:
                    var isShortCircuit = node.Text == "or" || node.Text == "and" && !isCondition;
                    CollectChildren(node, resolver, fields, isShortCircuit ? 1 : int.MaxValue, isCondition);
                    break;

                case StiCrystalNodeKind.Unary:
                case StiCrystalNodeKind.Index:
                case StiCrystalNodeKind.Array:
                case StiCrystalNodeKind.Range:
                    CollectChildren(node, resolver, fields, int.MaxValue, isCondition);
                    break;

                // The condition of If and the subject of Select are read in any case
                case StiCrystalNodeKind.If:
                case StiCrystalNodeKind.Select:
                    CollectChildren(node, resolver, fields, 1, isCondition);
                    break;
            }
        }

        private static void CollectChildren(StiCrystalNode node, IStiCrystalResolver resolver,
            List<StiCrystalReference> fields, int count, bool isCondition)
        {
            foreach (var child in node.Children.Take(count))
            {
                CollectFields(child, resolver, fields, isCondition);
            }
        }

        /// <summary>
        /// A field of a database gives itself, a formula gives its own fields.
        /// </summary>
        private static void AddReferenceFields(StiCrystalNode node, IStiCrystalResolver resolver,
            List<StiCrystalReference> fields)
        {
            var reference = resolver.ResolveReference(node.Text);
            if (reference == null)
                return;

            if (!string.IsNullOrEmpty(reference.DataSourceName))
            {
                if (reference.Name == null)
                    reference.Name = node.Text;

                fields.Add(reference);
            }

            else if (reference.NullFields != null)
            {
                fields.AddRange(reference.NullFields);
            }
        }

        /// <summary>
        /// The fields are compared by the data source and the column: the same field may be written differently.
        /// </summary>
        private static void CollectCheckedFields(StiCrystalNode node, IStiCrystalResolver resolver,
            HashSet<string> checkedFields)
        {
            if (node.Kind == StiCrystalNodeKind.Call && IsNullSafeFunction(node.Text))
            {
                foreach (var child in node.Children.Where(c => c.Kind == StiCrystalNodeKind.Reference))
                {
                    var reference = resolver.ResolveReference(child.Text);
                    if (!string.IsNullOrEmpty(reference?.DataSourceName))
                        checkedFields.Add(GetKey(reference));
                }
            }

            foreach (var child in node.Children)
            {
                CollectCheckedFields(child, resolver, checkedFields);
            }
        }

        private static string GetKey(StiCrystalReference field)
        {
            return field.DataSourceName + "." + field.ColumnName;
        }

        /// <summary>
        /// The functions which take NULL as a value: the checks of NULL and the summaries, which skip NULL.
        /// </summary>
        private static bool IsNullSafeFunction(string name)
        {
            switch (name.ToLowerInvariant())
            {
                case "isnull":
                case "previousisnull":
                case "nextisnull":
                case "hasvalue":
                case "previous":
                case "next":
                case "sum":
                case "average":
                case "count":
                case "distinctcount":
                case "minimum":
                case "maximum":
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>
        /// The code of a record selection which is true only when Crystal keeps the record. Crystal stops the
        /// formula at the first NULL field which it reads and drops the record; and and or are short-circuit, so the
        /// right operand of or is read only when the left one gave false without NULL. Returns null when the formula
        /// reads no field of a database or is not a single condition.
        /// </summary>
        public static string GetSelectionCode(StiCrystalNode node, IStiCrystalResolver resolver)
        {
            if (node.Kind == StiCrystalNodeKind.Sequence)
            {
                if (node.Children.Count != 1)
                    return null;

                node = node.Children[0];
            }

            var isGuarded = false;
            var code = GetSafeCode(node, resolver, ref isGuarded);
            return isGuarded ? code : null;
        }

        private static string GetSafeCode(StiCrystalNode node, IStiCrystalResolver resolver, ref bool isGuarded)
        {
            if (node.Kind == StiCrystalNodeKind.Binary && node.Text == "and")
            {
                var left = GetSafeCode(node.Children[0], resolver, ref isGuarded);
                var right = GetSafeCode(node.Children[1], resolver, ref isGuarded);
                return $"({left} && {right})";
            }

            if (node.Kind == StiCrystalNodeKind.Binary && node.Text == "or")
            {
                var left = GetSafeCode(node.Children[0], resolver, ref isGuarded);
                var noNulls = GetNoNullsCode(node.Children[0], resolver, ref isGuarded);
                var right = GetSafeCode(node.Children[1], resolver, ref isGuarded);
                return noNulls == null ? $"({left} || {right})" : $"({left} || {noNulls} && {right})";
            }

            // The condition of If is read in any case, then one of the branches
            if (node.Kind == StiCrystalNodeKind.If)
            {
                var condition = node.Children[0];
                var whenTrue = GetSafeCode(condition, resolver, ref isGuarded) + " && " +
                    GetSafeCode(node.Children[1], resolver, ref isGuarded);

                if (node.Children.Count < 3)
                    return $"({whenTrue})";

                var operand = TranslateCondition(condition, resolver).ToOperand();
                var whenFalse = Combine(GetNoNullsCode(condition, resolver, ref isGuarded), "!" + operand) + " && " +
                    GetSafeCode(node.Children[2], resolver, ref isGuarded);

                return $"({whenTrue} || {whenFalse})";
            }

            if (node.Kind == StiCrystalNodeKind.Unary && node.Text == "not")
            {
                var operand = TranslateCondition(node.Children[0], resolver);
                return Combine(GetNoNullsCode(node.Children[0], resolver, ref isGuarded), "!" + operand.ToOperand());
            }

            var expression = TranslateCondition(node, resolver);
            return Combine(GetNoNullsCode(node, resolver, ref isGuarded), expression.ToOperand());
        }

        private static StiCrystalExpression TranslateCondition(StiCrystalNode node, IStiCrystalResolver resolver)
        {
            var expression = new StiCrystalTranslator(resolver).Translate(node);
            if (expression == null || expression.Type != typeof(bool))
                throw new StiCrystalFormulaException("A part of the condition is not boolean", node.Position);

            return expression;
        }

        /// <summary>
        /// The check that the fields which the condition reads in any case are not NULL; null if there are none.
        /// </summary>
        private static string GetNoNullsCode(StiCrystalNode node, IStiCrystalResolver resolver, ref bool isGuarded)
        {
            var fields = Collect(node, resolver, isCondition: true);
            if (fields.Count == 0)
                return null;

            isGuarded = true;
            var nulls = fields.Select(f => f.GetIsNullCode());
            return $"!({string.Join(" || ", nulls)})";
        }

        private static string Combine(string noNulls, string condition)
        {
            return noNulls == null ? condition : $"({noNulls} && {condition})";
        }
        #endregion
    }
}
