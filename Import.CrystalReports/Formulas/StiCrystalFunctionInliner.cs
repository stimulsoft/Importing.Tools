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
    /// Puts the bodies of the custom functions of the report instead of their calls: a parameter in the body becomes
    /// the expression of its argument. A function called by another one is put too, up to a depth that stops
    /// a recursion. The custom functions failed about 140 formulas of the corpus (Lbl of the labels in several
    /// languages, Monthfirstday and Daterangestring of the date ranges, the colors of Disciplinebackgroundcolor).
    /// The tree is changed before the translation, so the constants of a formula of the formatting are found in
    /// the body of a function too.
    /// </summary>
    public static class StiCrystalFunctionInliner
    {
        #region Consts
        private const int MaxDepth = 10;
        #endregion

        #region Methods
        public static StiCrystalNode Inline(StiCrystalNode node, IStiCrystalResolver resolver)
        {
            return Inline(node, resolver, 0);
        }

        private static StiCrystalNode Inline(StiCrystalNode node, IStiCrystalResolver resolver, int depth)
        {
            if (node == null)
                return null;

            var function = depth < MaxDepth ? GetFunction(node, resolver) : null;
            if (function != null)
            {
                var arguments = node.Children.Select(c => Inline(c, resolver, depth)).ToList();
                var body = Substitute(function.Body, GetValues(function, arguments, node.Position));
                return Inline(body, resolver, depth + 1);
            }

            var copy = Copy(node);
            foreach (var child in node.Children)
            {
                copy.Add(Inline(child, resolver, depth));
            }

            return copy;
        }

        /// <summary>
        /// The custom function of a call; a name without parentheses calls a function without required parameters.
        /// </summary>
        private static StiCrystalCustomFunction GetFunction(StiCrystalNode node, IStiCrystalResolver resolver)
        {
            if (node.Kind != StiCrystalNodeKind.Call && node.Kind != StiCrystalNodeKind.Identifier)
                return null;

            var function = resolver.ResolveFunction(node.Text);
            if (function == null)
                return null;

            if (node.Kind == StiCrystalNodeKind.Identifier && function.Parameters.Any(p => p.Children.Count == 0))
                return null;

            return function;
        }

        private static Dictionary<string, StiCrystalNode> GetValues(StiCrystalCustomFunction function,
            List<StiCrystalNode> arguments, int position)
        {
            if (arguments.Count > function.Parameters.Count)
                throw new StiCrystalFormulaException(
                    $"The function {function.Name} has {function.Parameters.Count} parameters", position);

            var values = new Dictionary<string, StiCrystalNode>(StringComparer.OrdinalIgnoreCase);
            for (var index = 0; index < function.Parameters.Count; index++)
            {
                var parameter = function.Parameters[index];
                var value = index < arguments.Count ? arguments[index] : parameter.Children.FirstOrDefault();
                if (value == null)
                    throw new StiCrystalFormulaException(
                        $"The function {function.Name} needs the parameter {parameter.Text}", position);

                values[parameter.Text] = value;
            }

            return values;
        }

        /// <summary>
        /// A copy of the body where the parameters are replaced by the expressions of their arguments.
        /// </summary>
        private static StiCrystalNode Substitute(StiCrystalNode node, Dictionary<string, StiCrystalNode> values)
        {
            if (node == null)
                return null;

            StiCrystalNode value;
            if (node.Kind == StiCrystalNodeKind.Identifier && values.TryGetValue(node.Text, out value))
                return value;

            var copy = Copy(node);
            foreach (var child in node.Children)
            {
                copy.Add(Substitute(child, values));
            }

            return copy;
        }

        private static StiCrystalNode Copy(StiCrystalNode node)
        {
            return new StiCrystalNode(node.Kind, node.Text, node.Position)
            {
                Scope = node.Scope,
                VariableType = node.VariableType,
                IsArrayVariable = node.IsArrayVariable,
                IsRangeVariable = node.IsRangeVariable
            };
        }
        #endregion
    }
}
