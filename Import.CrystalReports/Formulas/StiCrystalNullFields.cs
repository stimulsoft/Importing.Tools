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
        public static List<StiCrystalReference> Collect(StiCrystalNode node, IStiCrystalResolver resolver)
        {
            var checkedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            CollectCheckedFields(node, checkedFields);

            var fields = new List<StiCrystalReference>();
            CollectFields(node, resolver, fields);

            return fields
                .Where(f => f.Name == null || !checkedFields.Contains(f.Name))
                .GroupBy(f => f.DataSourceName + "." + f.ColumnName, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
        }

        private static void CollectFields(StiCrystalNode node, IStiCrystalResolver resolver,
            List<StiCrystalReference> fields)
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
                        CollectChildren(node, resolver, fields, isIif ? 1 : int.MaxValue);

                    break;

                case StiCrystalNodeKind.Binary:
                    var isLogical = node.Text == "and" || node.Text == "or";
                    CollectChildren(node, resolver, fields, isLogical ? 1 : int.MaxValue);
                    break;

                case StiCrystalNodeKind.Unary:
                case StiCrystalNodeKind.Index:
                case StiCrystalNodeKind.Array:
                case StiCrystalNodeKind.Range:
                    CollectChildren(node, resolver, fields, int.MaxValue);
                    break;

                // The condition of If and the subject of Select are read in any case
                case StiCrystalNodeKind.If:
                case StiCrystalNodeKind.Select:
                    CollectChildren(node, resolver, fields, 1);
                    break;
            }
        }

        private static void CollectChildren(StiCrystalNode node, IStiCrystalResolver resolver,
            List<StiCrystalReference> fields, int count)
        {
            foreach (var child in node.Children.Take(count))
            {
                CollectFields(child, resolver, fields);
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

        private static void CollectCheckedFields(StiCrystalNode node, HashSet<string> checkedFields)
        {
            if (node.Kind == StiCrystalNodeKind.Call && IsNullSafeFunction(node.Text))
            {
                foreach (var child in node.Children.Where(c => c.Kind == StiCrystalNodeKind.Reference))
                {
                    checkedFields.Add(child.Text);
                }
            }

            foreach (var child in node.Children)
            {
                CollectCheckedFields(child, checkedFields);
            }
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
        #endregion
    }
}
