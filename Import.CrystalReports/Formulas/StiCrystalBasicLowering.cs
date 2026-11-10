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
    /// Turns the statements of a Basic formula into one expression of the tree of Crystal, which the translator
    /// knows. The value of Formula and of the local variables is followed through the statements as an expression:
    /// an assignment replaces it, the branches of If give a conditional expression of the values of the branches,
    /// a variable which a branch doesn't assign keeps its value. The global and shared variables live in the report,
    /// so their declarations and assignments stay statements before the value of the formula, the assignments in
    /// the branches stay in an If. The result is Sequence(declarations, assignments, value of Formula).
    /// </summary>
    public class StiCrystalBasicLowering
    {
        #region Consts
        private const string FormulaName = "formula";
        #endregion

        #region Fields
        private readonly List<StiCrystalNode> declarations = new List<StiCrystalNode>();
        private readonly HashSet<string> globals = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Methods
        /// <summary>
        /// The expression of the formula; an empty Sequence for a formula without statements, like the one of
        /// comments only.
        /// </summary>
        public StiCrystalNode Lower(StiCrystalNode block)
        {
            var values = new Dictionary<string, StiCrystalNode>(StringComparer.OrdinalIgnoreCase);
            var effects = new List<StiCrystalNode>();
            LowerStatement(block, values, effects);

            StiCrystalNode result;
            if (!values.TryGetValue(FormulaName, out result) || result == null)
            {
                if (declarations.Count == 0 && effects.Count == 0)
                    return new StiCrystalNode(StiCrystalNodeKind.Sequence, null, block.Position);

                throw new StiCrystalFormulaException("The formula has no value", block.Position);
            }

            var statements = declarations.Concat(effects).ToList();
            if (statements.Count == 0)
                return result;

            statements.Add(result);
            return new StiCrystalNode(StiCrystalNodeKind.Sequence, null, block.Position, statements.ToArray());
        }

        private void LowerStatement(StiCrystalNode node, Dictionary<string, StiCrystalNode> values,
            List<StiCrystalNode> effects)
        {
            switch (node.Kind)
            {
                case StiCrystalNodeKind.Sequence:
                    foreach (var child in node.Children)
                    {
                        LowerStatement(child, values, effects);
                    }

                    break;

                case StiCrystalNodeKind.Declare:
                    LowerDeclaration(node, values, effects);
                    break;

                case StiCrystalNodeKind.Assign:
                    LowerAssignment(node.Children[0].Text, node.Children[1], node.Position, values, effects);
                    break;

                case StiCrystalNodeKind.If:
                    LowerIf(node, values, effects);
                    break;

                default:
                    // WhilePrintingRecords, EvaluateAfter and the other statements without a value
                    effects.Add(Substitute(node, values));
                    break;
            }
        }

        private void LowerDeclaration(StiCrystalNode node, Dictionary<string, StiCrystalNode> values,
            List<StiCrystalNode> effects)
        {
            if (node.IsArrayVariable)
                throw new StiCrystalFormulaException($"The array variable {node.Text} is not supported", node.Position);

            if (node.Scope == "local")
            {
                values[node.Text] = node.Children.Count > 0
                    ? Substitute(node.Children[0], values)
                    : GetDefaultValue(node.VariableType, node.Position);

                return;
            }

            if (globals.Add(node.Text))
            {
                declarations.Add(new StiCrystalNode(StiCrystalNodeKind.Declare, node.Text, node.Position)
                {
                    Scope = node.Scope,
                    VariableType = node.VariableType
                });
            }

            values.Remove(node.Text);
            if (node.Children.Count > 0)
                LowerAssignment(node.Text, node.Children[0], node.Position, values, effects);
        }

        private void LowerAssignment(string name, StiCrystalNode value, int position,
            Dictionary<string, StiCrystalNode> values, List<StiCrystalNode> effects)
        {
            var substituted = Substitute(value, values);
            if (!globals.Contains(name))
            {
                values[name] = substituted;
                return;
            }

            var target = new StiCrystalNode(StiCrystalNodeKind.Identifier, name, position);
            effects.Add(new StiCrystalNode(StiCrystalNodeKind.Assign, ":=", position, target, substituted));
        }

        /// <summary>
        /// The branches start from the values before If; a variable which a branch assigns gets the conditional
        /// value, a branch which doesn't assign it gives the value before If. A variable which had no value gets
        /// If without Else, which the translator completes by the default value of the type, as Crystal does.
        /// </summary>
        private void LowerIf(StiCrystalNode node, Dictionary<string, StiCrystalNode> values,
            List<StiCrystalNode> effects)
        {
            var condition = Substitute(node.Children[0], values);

            var thenValues = new Dictionary<string, StiCrystalNode>(values, StringComparer.OrdinalIgnoreCase);
            var thenEffects = new List<StiCrystalNode>();
            LowerStatement(node.Children[1], thenValues, thenEffects);

            var elseValues = new Dictionary<string, StiCrystalNode>(values, StringComparer.OrdinalIgnoreCase);
            var elseEffects = new List<StiCrystalNode>();
            if (node.Children.Count > 2)
                LowerStatement(node.Children[2], elseValues, elseEffects);

            var names = thenValues.Keys.Concat(elseValues.Keys).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            foreach (var name in names)
            {
                StiCrystalNode before;
                StiCrystalNode thenValue;
                StiCrystalNode elseValue;
                values.TryGetValue(name, out before);
                thenValues.TryGetValue(name, out thenValue);
                elseValues.TryGetValue(name, out elseValue);

                if (thenValue == before && elseValue == before)
                    continue;

                values[name] = CreateIf(condition, thenValue, elseValue, node.Position);
            }

            if (thenEffects.Count > 0 || elseEffects.Count > 0)
                effects.Add(CreateIf(condition, ToSequence(thenEffects, node.Position),
                    ToSequence(elseEffects, node.Position), node.Position));
        }

        private static StiCrystalNode CreateIf(StiCrystalNode condition, StiCrystalNode thenValue,
            StiCrystalNode elseValue, int position)
        {
            if (thenValue == null)
            {
                var negation = new StiCrystalNode(StiCrystalNodeKind.Unary, "not", position, condition);
                return new StiCrystalNode(StiCrystalNodeKind.If, null, position, negation, elseValue);
            }

            var node = new StiCrystalNode(StiCrystalNodeKind.If, null, position, condition, thenValue);
            if (elseValue != null)
                node.Add(elseValue);

            return node;
        }

        private static StiCrystalNode ToSequence(List<StiCrystalNode> statements, int position)
        {
            if (statements.Count == 0)
                return null;

            return statements.Count == 1
                ? statements[0]
                : new StiCrystalNode(StiCrystalNodeKind.Sequence, null, position, statements.ToArray());
        }

        /// <summary>
        /// A copy of the expression where the local variables and Formula are replaced by their current values;
        /// IIf becomes If.
        /// </summary>
        private StiCrystalNode Substitute(StiCrystalNode node, Dictionary<string, StiCrystalNode> values)
        {
            if (node == null)
                return null;

            if (node.Kind == StiCrystalNodeKind.Identifier && !globals.Contains(node.Text))
            {
                StiCrystalNode value;
                if (values.TryGetValue(node.Text, out value))
                {
                    if (value == null)
                        throw new StiCrystalFormulaException($"The variable {node.Text} has no value", node.Position);

                    return value;
                }

                if (string.Equals(node.Text, FormulaName, StringComparison.OrdinalIgnoreCase))
                    throw new StiCrystalFormulaException("The formula has no value yet", node.Position);
            }

            if (node.Kind == StiCrystalNodeKind.Call && values.ContainsKey(node.Text))
                throw new StiCrystalFormulaException($"The array variable {node.Text} is not supported", node.Position);

            if (node.Kind == StiCrystalNodeKind.Call && node.Children.Count == 3 &&
                string.Equals(node.Text, "iif", StringComparison.OrdinalIgnoreCase))
            {
                return new StiCrystalNode(StiCrystalNodeKind.If, null, node.Position,
                    node.Children.Select(c => Substitute(c, values)).ToArray());
            }

            var copy = new StiCrystalNode(node.Kind, node.Text, node.Position)
            {
                Scope = node.Scope,
                VariableType = node.VariableType,
                IsArrayVariable = node.IsArrayVariable,
                IsRangeVariable = node.IsRangeVariable
            };

            foreach (var child in node.Children)
            {
                copy.Add(Substitute(child, values));
            }

            return copy;
        }

        /// <summary>
        /// The value of a declared variable before an assignment, as Crystal gives it; a date has no such value
        /// here, its reading before an assignment is not supported.
        /// </summary>
        private static StiCrystalNode GetDefaultValue(string type, int position)
        {
            switch (type)
            {
                case "numbervar":
                case "currencyvar":
                    return new StiCrystalNode(StiCrystalNodeKind.Number, "0", position);

                case "stringvar":
                    return new StiCrystalNode(StiCrystalNodeKind.String, string.Empty, position);

                case "booleanvar":
                    return new StiCrystalNode(StiCrystalNodeKind.Boolean, "false", position);

                default:
                    return null;
            }
        }
        #endregion
    }
}
