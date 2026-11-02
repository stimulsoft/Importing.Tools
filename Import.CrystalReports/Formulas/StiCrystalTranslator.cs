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
using System.Globalization;
using System.Linq;
using System.Text;

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// Translates a formula in Crystal syntax to a C# expression of a Stimulsoft report. The expression uses only
    /// the subset of C# which both the compiler and the parser of the interpretation mode understand: the operators,
    /// the ternary operator, the casts, the members of the base types and the functions of the dictionary.
    /// A formula of several statements is translated if the statements only declare and assign local variables:
    /// the values of the variables are substituted into the last expression. A global or shared variable becomes
    /// a variable of the dictionary; its assignment is the value of the whole formula, Variable = value, because
    /// the parser of the interpretation mode accepts an assignment only as the whole expression.
    /// </summary>
    public partial class StiCrystalTranslator
    {
        #region Consts
        /// <summary>
        /// The words which set the time of evaluation of a formula; Stimulsoft has no such notion, they are skipped.
        /// </summary>
        private static readonly string[] EvaluationTimeWords =
        {
            "whileprintingrecords", "whilereadingrecords", "beforereadingrecords"
        };
        #endregion

        #region Fields
        private readonly IStiCrystalResolver resolver;
        private Dictionary<string, StiCrystalExpression> variables;
        #endregion

        #region Properties
        /// <summary>
        /// True if the last translated formula contains a summary, like Sum or Count. Such an expression works only
        /// in the text of a component, so it is placed there instead of a variable.
        /// </summary>
        public bool HasSummaries { get; private set; }
        #endregion

        #region Methods
        /// <summary>
        /// Translates the formula. Returns null for an empty formula, for example the one of comments only.
        /// Throws StiCrystalFormulaException if the formula can't be translated.
        /// </summary>
        public StiCrystalExpression Translate(string formula)
        {
            return Translate(StiCrystalParser.Parse(formula));
        }

        /// <summary>
        /// Translates the parsed formula. Returns null for an empty formula.
        /// </summary>
        public StiCrystalExpression Translate(StiCrystalNode node)
        {
            HasSummaries = false;
            variables = new Dictionary<string, StiCrystalExpression>(StringComparer.OrdinalIgnoreCase);

            if (node.Kind == StiCrystalNodeKind.Sequence && node.Children.Count == 0)
                return null;

            var result = node.Kind == StiCrystalNodeKind.Sequence ? TranslateSequence(node) : TranslateStatement(node);
            if (result == null)
                throw new StiCrystalFormulaException("The formula has no value");

            if (result.IsArray || result.IsRange)
                throw new StiCrystalFormulaException("The value of the formula is an array or a range");

            return result;
        }

        /// <summary>
        /// Translates the formula and converts a number to the specified type of the result, if it is known.
        /// </summary>
        public StiCrystalExpression Translate(string formula, Type resultType)
        {
            var result = Translate(formula);
            if (result == null || resultType == null || resultType == typeof(object) || result.Type == resultType)
                return result;

            if (IsNumber(resultType) && IsNumber(result.Type))
            {
                var needsCast = result.Type == typeof(double) && IsInteger(resultType) ||
                    result.Type == typeof(decimal) || resultType == typeof(decimal);

                var code = needsCast
                    ? $"({GetTypeName(resultType)}){result.ToOperand()}"
                    : result.Code;

                return new StiCrystalExpression(code, resultType)
                {
                    Reference = result.Reference,
                    IsAssignment = result.IsAssignment
                };
            }

            return result;
        }

        /// <summary>
        /// Translates the statements in order and returns the value of the last one. An assignment of a global
        /// variable may be followed only by the reading of the same variable, which gives the same value.
        /// </summary>
        private StiCrystalExpression TranslateSequence(StiCrystalNode node)
        {
            StiCrystalExpression result = null;
            foreach (var statement in node.Children)
            {
                var value = TranslateStatement(statement);
                if (value == null)
                    continue;

                if (result != null && result.IsAssignment)
                {
                    if (value.Code != result.AssignedVariable)
                        throw new StiCrystalFormulaException("An assignment of a global variable must be the value " +
                            "of the formula", statement.Position);

                    continue;
                }

                result = value;
            }

            return result;
        }

        /// <summary>
        /// Translates a statement; returns null for a statement without a value, like WhilePrintingRecords.
        /// </summary>
        private StiCrystalExpression TranslateStatement(StiCrystalNode node)
        {
            switch (node.Kind)
            {
                case StiCrystalNodeKind.Identifier:
                    if (EvaluationTimeWords.Contains(node.Text.ToLowerInvariant()))
                        return null;

                    break;

                case StiCrystalNodeKind.Call:
                    if (string.Equals(node.Text, "EvaluateAfter", StringComparison.OrdinalIgnoreCase))
                        return null;

                    break;

                case StiCrystalNodeKind.Declare:
                    return TranslateDeclaration(node);

                case StiCrystalNodeKind.Assign:
                    return TranslateAssignment(node);

                case StiCrystalNodeKind.Sequence:
                    return TranslateSequence(node);
            }

            return TranslateNode(node);
        }

        private StiCrystalExpression TranslateDeclaration(StiCrystalNode node)
        {
            if (node.IsArrayVariable || node.IsRangeVariable)
                throw new StiCrystalFormulaException($"The array variable {node.Text} is not supported", node.Position);

            var type = GetVariableType(node.VariableType);
            if (node.Scope != "local")
                return TranslateGlobalDeclaration(node, type);

            var value = node.Children.Count > 0 ? TranslateNode(node.Children[0]) : GetDefaultValue(type);
            RequireNoAssignment(value, node);

            variables[node.Text] = Parenthesize(value);
            return value;
        }

        /// <summary>
        /// A global or shared variable lives in the dictionary of the report and keeps its value between formulas.
        /// The declaration without a value only reads the variable.
        /// </summary>
        private StiCrystalExpression TranslateGlobalDeclaration(StiCrystalNode node, Type type)
        {
            var reference = resolver.ResolveVariable(node.Text, type);
            if (reference == null)
                throw new StiCrystalFormulaException($"The {node.Scope} variable {node.Text} is not supported",
                    node.Position);

            var variable = new StiCrystalExpression(reference.Expression, reference.Type, isAtomic: true)
            {
                AssignedVariable = reference.Expression
            };

            variables[node.Text] = variable;
            if (node.Children.Count == 0)
                return variable;

            return CreateAssignment(variable, TranslateNode(node.Children[0]), node);
        }

        private StiCrystalExpression TranslateAssignment(StiCrystalNode node)
        {
            var target = node.Children[0];
            StiCrystalExpression variable;
            if (target.Kind != StiCrystalNodeKind.Identifier || !variables.TryGetValue(target.Text, out variable))
                throw new StiCrystalFormulaException("Only an assignment of a declared variable is supported",
                    node.Position);

            var value = TranslateNode(node.Children[1]);
            if (variable.AssignedVariable != null)
                return CreateAssignment(variable, value, node);

            RequireNoAssignment(value, node);
            variables[target.Text] = Parenthesize(value);
            return value;
        }

        private static StiCrystalExpression CreateAssignment(StiCrystalExpression variable, StiCrystalExpression value,
            StiCrystalNode node)
        {
            RequireNoAssignment(value, node);

            var converted = Coerce(value, variable.Type);
            if (converted.Type != variable.Type && !(IsNumber(variable.Type) && IsNumber(converted.Type)))
                throw new StiCrystalFormulaException($"Can't assign {value.Type.Name} to {variable.Type.Name}",
                    node.Position);

            if (IsNumber(variable.Type) && converted.Type != variable.Type)
                converted = new StiCrystalExpression($"({GetTypeName(variable.Type)}){converted.ToOperand()}",
                    variable.Type);

            return new StiCrystalExpression($"{variable.AssignedVariable} = {converted.Code}", variable.Type)
            {
                IsAssignment = true,
                AssignedVariable = variable.AssignedVariable
            };
        }

        /// <summary>
        /// An assignment of a global variable can't be a part of an expression: the parser of the interpretation
        /// mode accepts it only as the whole expression.
        /// </summary>
        private static void RequireNoAssignment(StiCrystalExpression value, StiCrystalNode node)
        {
            if (value != null && value.IsAssignment)
                throw new StiCrystalFormulaException("An assignment of a global variable inside an expression " +
                    "is not supported", node.Position);
        }

        /// <summary>
        /// The value of a variable is substituted into the following expressions, so it must be an operand.
        /// </summary>
        private static StiCrystalExpression Parenthesize(StiCrystalExpression value)
        {
            if (value.IsAtomic || value.IsArray || value.IsRange)
                return value;

            return new StiCrystalExpression("(" + value.Code + ")", value.Type, isAtomic: true)
            {
                Reference = value.Reference,
                Literal = value.Literal
            };
        }

        private StiCrystalExpression TranslateNode(StiCrystalNode node)
        {
            switch (node.Kind)
            {
                case StiCrystalNodeKind.Number:
                    return TranslateNumber(node);

                case StiCrystalNodeKind.String:
                    return new StiCrystalExpression(ToCSharpString(node.Text), typeof(string), isAtomic: true)
                    {
                        Literal = node.Text
                    };

                case StiCrystalNodeKind.Boolean:
                    return new StiCrystalExpression(node.Text, typeof(bool), isAtomic: true);

                case StiCrystalNodeKind.Date:
                    return TranslateDate(node);

                case StiCrystalNodeKind.Reference:
                    return TranslateReference(node);

                case StiCrystalNodeKind.Identifier:
                    return TranslateIdentifier(node);

                case StiCrystalNodeKind.Call:
                    return TranslateCall(node.Text, node.Children, node.Position);

                case StiCrystalNodeKind.Unary:
                    return TranslateUnary(node);

                case StiCrystalNodeKind.Binary:
                    return TranslateBinary(node);

                case StiCrystalNodeKind.If:
                    return TranslateIf(node);

                case StiCrystalNodeKind.Select:
                    return TranslateSelect(node);

                case StiCrystalNodeKind.Array:
                    return StiCrystalExpression.CreateArray(node.Children.Select(TranslateNode).ToList());

                case StiCrystalNodeKind.Range:
                    return TranslateRange(node);

                case StiCrystalNodeKind.Index:
                    return TranslateIndex(node);

                case StiCrystalNodeKind.Sequence:
                case StiCrystalNodeKind.Declare:
                case StiCrystalNodeKind.Assign:
                    var value = TranslateStatement(node);
                    if (value == null)
                        throw new StiCrystalFormulaException("The statement has no value", node.Position);

                    RequireNoAssignment(value, node);
                    return value;
            }

            throw new StiCrystalFormulaException($"{node.Kind} is not supported", node.Position);
        }

        private static StiCrystalExpression TranslateNumber(StiCrystalNode node)
        {
            int intValue;
            if (int.TryParse(node.Text, NumberStyles.None, CultureInfo.InvariantCulture, out intValue))
                return new StiCrystalExpression(intValue.ToString(CultureInfo.InvariantCulture), typeof(int),
                    isAtomic: true);

            var value = double.Parse(node.Text, CultureInfo.InvariantCulture);
            var code = value.ToString("R", CultureInfo.InvariantCulture);
            if (code.IndexOf('.') < 0 && code.IndexOf('E') < 0)
                code += ".0";

            return new StiCrystalExpression(code, typeof(double), isAtomic: true);
        }

        /// <summary>
        /// A date literal #1/31/2020# or #2020-01-31 10:30# becomes a call of DateSerial.
        /// </summary>
        private static StiCrystalExpression TranslateDate(StiCrystalNode node)
        {
            DateTime date;
            var styles = DateTimeStyles.AllowWhiteSpaces;
            if (!DateTime.TryParse(node.Text, CultureInfo.InvariantCulture, styles, out date))
                throw new StiCrystalFormulaException($"Unknown date #{node.Text}#", node.Position);

            var code = $"DateSerial({date.Year}, {date.Month}, {date.Day})";
            if (date.TimeOfDay.Ticks != 0)
                code += $".AddSeconds({(long)date.TimeOfDay.TotalSeconds})";

            return new StiCrystalExpression(code, typeof(DateTime), isAtomic: true);
        }

        private StiCrystalExpression TranslateReference(StiCrystalNode node)
        {
            var reference = resolver.ResolveReference(node.Text);
            if (reference == null)
                throw new StiCrystalFormulaException($"Unknown reference {{{node.Text}}}", node.Position);

            if (reference.Name == null)
                reference.Name = node.Text;

            if (reference.HasSummaries)
                HasSummaries = true;

            return CreateFromReference(reference);
        }

        private static StiCrystalExpression CreateFromReference(StiCrystalReference reference)
        {
            if (reference.IsRange)
            {
                var isDate = IsDate(reference.Type);
                var fromCode = reference.Expression + (isDate ? ".FromDate" : ".From");
                var toCode = reference.Expression + (isDate ? ".ToDate" : ".To");
                var from = new StiCrystalExpression(fromCode, reference.Type, isAtomic: true);
                var to = new StiCrystalExpression(toCode, reference.Type, isAtomic: true);
                var range = StiCrystalExpression.CreateRange(from, true, to, true);
                range.Reference = reference;
                return range;
            }

            // Crystal reads the strings of a database without the trailing spaces, Stimulsoft keeps them: a column
            // of the type char, or a value with a space at the end, would compare and concatenate differently
            var code = reference.Expression;
            if (!string.IsNullOrEmpty(reference.DataSourceName) && reference.Type == typeof(string))
                code += ".TrimEnd()";

            return new StiCrystalExpression(code, reference.Type ?? typeof(object), isAtomic: true)
            {
                Reference = reference
            };
        }

        private StiCrystalExpression TranslateIdentifier(StiCrystalNode node)
        {
            StiCrystalExpression variable;
            if (variables.TryGetValue(node.Text, out variable))
                return variable;

            return TranslateCall(node.Text, new List<StiCrystalNode>(), node.Position);
        }

        private StiCrystalExpression TranslateUnary(StiCrystalNode node)
        {
            var operand = TranslateNode(node.Children[0]);
            switch (node.Text)
            {
                case "not":
                    RequireBoolean(operand, node);
                    return new StiCrystalExpression("!" + operand.ToOperand(), typeof(bool));

                case "-":
                    RequireNumber(operand, node);
                    return new StiCrystalExpression("-" + operand.ToOperand(), operand.Type);

                case "+":
                    RequireNumber(operand, node);
                    return operand;

                case "$":
                    RequireNumber(operand, node);
                    return Coerce(operand, typeof(decimal));
            }

            throw new StiCrystalFormulaException($"The operator {node.Text} is not supported", node.Position);
        }

        private StiCrystalExpression TranslateBinary(StiCrystalNode node)
        {
            var op = node.Text;
            var equality = TranslateStrCmpEquality(node);
            if (equality != null)
                return equality;

            var left = TranslateNode(node.Children[0]);
            var right = TranslateNode(node.Children[1]);

            switch (op)
            {
                case "and":
                case "or":
                case "xor":
                case "eqv":
                case "imp":
                    return TranslateLogical(op, left, right, node);

                case "=":
                case "<>":
                case "<":
                case ">":
                case "<=":
                case ">=":
                    return TranslateComparison(op, left, right, node);

                case "in":
                    return TranslateIn(left, right, node);

                case "startswith":
                    RequireString(left, node);
                    RequireString(right, node);
                    return new StiCrystalExpression($"{left.ToOperand()}.StartsWith({right.Code})", typeof(bool));

                case "+":
                case "-":
                    return TranslateAdditive(op, left, right, node);

                case "&":
                    return new StiCrystalExpression($"{ToText(left).ToOperand()} + {ToText(right).ToOperand()}",
                        typeof(string));

                case "*":
                    return TranslateArithmetic("*", left, right, node);

                case "/":
                    return TranslateDivision(left, right, node);

                case "%":
                    var percent = TranslateDivision(left, right, node);
                    return new StiCrystalExpression($"{percent.ToOperand()} * 100", percent.Type);

                case "\\":
                    var quotient = TranslateDivision(left, right, node);
                    return new StiCrystalExpression($"Math.Truncate({quotient.Code})", quotient.Type, isAtomic: true);

                case "mod":
                    return TranslateArithmetic("%", left, right, node);

                case "^":
                    RequireNumber(left, node);
                    RequireNumber(right, node);
                    var code = $"Math.Pow({Coerce(left, typeof(double)).Code}, {Coerce(right, typeof(double)).Code})";
                    return new StiCrystalExpression(code, typeof(double), isAtomic: true);
            }

            throw new StiCrystalFormulaException($"The operator {op} is not supported", node.Position);
        }

        /// <summary>
        /// StrCmp(a, b) = 0 and StrCmp(a, b) <> 0 are the equality of the strings: the interpretation can't
        /// compare the order of the strings, but it can check that they are equal.
        /// </summary>
        private StiCrystalExpression TranslateStrCmpEquality(StiCrystalNode node)
        {
            if (node.Text != "=" && node.Text != "<>")
                return null;

            var call = node.Children[0];
            var zero = node.Children[1];
            var isStrCmp = call.Kind == StiCrystalNodeKind.Call && call.Children.Count == 2 &&
                string.Equals(call.Text, "StrCmp", StringComparison.OrdinalIgnoreCase);
            if (!isStrCmp || zero.Kind != StiCrystalNodeKind.Number || zero.Text != "0")
                return null;

            var first = TranslateNode(call.Children[0]);
            var second = TranslateNode(call.Children[1]);
            if (first.Type != typeof(string) || second.Type != typeof(string))
                return null;

            var op = node.Text == "=" ? "==" : "!=";
            return new StiCrystalExpression($"{first.ToOperand()} {op} {second.ToOperand()}", typeof(bool));
        }

        private static StiCrystalExpression TranslateLogical(string op, StiCrystalExpression left,
            StiCrystalExpression right, StiCrystalNode node)
        {
            RequireBoolean(left, node);
            RequireBoolean(right, node);

            switch (op)
            {
                case "and":
                    return new StiCrystalExpression($"{left.ToOperand()} && {right.ToOperand()}", typeof(bool));

                case "or":
                    return new StiCrystalExpression($"{left.ToOperand()} || {right.ToOperand()}", typeof(bool));

                case "xor":
                    return new StiCrystalExpression($"{left.ToOperand()} != {right.ToOperand()}", typeof(bool));

                case "eqv":
                    return new StiCrystalExpression($"{left.ToOperand()} == {right.ToOperand()}", typeof(bool));

                default:
                    return new StiCrystalExpression($"!{left.ToOperand()} || {right.ToOperand()}", typeof(bool));
            }
        }

        private StiCrystalExpression TranslateComparison(string op, StiCrystalExpression left,
            StiCrystalExpression right, StiCrystalNode node)
        {
            if (op == "=" && (right.IsRange || right.Reference != null && right.Reference.IsMultiValue))
                return TranslateIn(left, right, node);

            if (op == "=" && (left.IsRange || left.Reference != null && left.Reference.IsMultiValue))
                return TranslateIn(right, left, node);

            var csharpOp = op == "=" ? "==" : op == "<>" ? "!=" : op;

            if (IsNumber(left.Type) && IsNumber(right.Type))
            {
                var type = GetCommonNumberType(left.Type, right.Type);
                return new StiCrystalExpression(
                    $"{Coerce(left, type).ToOperand()} {csharpOp} {Coerce(right, type).ToOperand()}", typeof(bool));
            }

            if (left.Type == typeof(string) && right.Type == typeof(string) && csharpOp != "==" && csharpOp != "!=")
            {
                return new StiCrystalExpression(
                    $"string.Compare({left.Code}, {right.Code}, StringComparison.Ordinal) {csharpOp} 0", typeof(bool));
            }

            if (left.Type != right.Type && left.Type != typeof(object) && right.Type != typeof(object))
                throw new StiCrystalFormulaException($"Can't compare {left.Type.Name} and {right.Type.Name}",
                    node.Position);

            return new StiCrystalExpression($"{left.ToOperand()} {csharpOp} {right.ToOperand()}", typeof(bool));
        }

        /// <summary>
        /// The operator In checks an array, a range or a parameter with several values.
        /// </summary>
        private StiCrystalExpression TranslateIn(StiCrystalExpression left, StiCrystalExpression right,
            StiCrystalNode node)
        {
            if (right.Reference != null && right.Reference.IsMultiValue)
                return new StiCrystalExpression($"{right.Code}.Contains({left.Code})", typeof(bool));

            if (right.IsArray)
            {
                if (right.Items.Count == 0)
                    return new StiCrystalExpression("false", typeof(bool), isAtomic: true);

                var parts = right.Items.Select(item => item.IsRange
                    ? TranslateInRange(left, item, node)
                    : TranslateComparison("=", left, item, node));

                return JoinConditions(parts.ToList(), " || ");
            }

            if (right.IsRange)
                return TranslateInRange(left, right, node);

            if (left.Type == typeof(string) && right.Type == typeof(string))
                return new StiCrystalExpression($"{right.ToOperand()}.Contains({left.Code})", typeof(bool));

            throw new StiCrystalFormulaException("The operator In needs an array or a range", node.Position);
        }

        private StiCrystalExpression TranslateInRange(StiCrystalExpression value, StiCrystalExpression range,
            StiCrystalNode node)
        {
            var parts = new List<StiCrystalExpression>();
            if (range.From != null)
                parts.Add(TranslateComparison(range.IncludeFrom ? ">=" : ">", value, range.From, node));

            if (range.To != null)
                parts.Add(TranslateComparison(range.IncludeTo ? "<=" : "<", value, range.To, node));

            return JoinConditions(parts, " && ");
        }

        /// <summary>
        /// Joins the conditions with the logical operator; one condition is returned as is.
        /// </summary>
        private static StiCrystalExpression JoinConditions(List<StiCrystalExpression> parts, string op)
        {
            if (parts.Count == 1)
                return parts[0];

            return new StiCrystalExpression(string.Join(op, parts.Select(p => p.ToOperand())), typeof(bool));
        }

        private StiCrystalExpression TranslateRange(StiCrystalNode node)
        {
            var first = TranslateNode(node.Children[0]);
            switch (node.Text)
            {
                case "upfrom":
                    return StiCrystalExpression.CreateRange(first, true, null, false);

                case "upfrom_":
                    return StiCrystalExpression.CreateRange(first, false, null, false);

                case "upto":
                    return StiCrystalExpression.CreateRange(null, false, first, true);

                case "upto_":
                    return StiCrystalExpression.CreateRange(null, false, first, false);
            }

            var second = TranslateNode(node.Children[1]);
            var includeFrom = !node.Text.StartsWith("_");
            var includeTo = !node.Text.EndsWith("_");
            return StiCrystalExpression.CreateRange(first, includeFrom, second, includeTo);
        }

        /// <summary>
        /// The subscript of a string takes a character or, with a range, a substring; the characters count from 1.
        /// </summary>
        private StiCrystalExpression TranslateIndex(StiCrystalNode node)
        {
            var split = TranslateFirstSplitItem(node);
            if (split != null)
                return split;

            var target = TranslateNode(node.Children[0]);
            var subscript = TranslateNode(node.Children[1]);
            if (target.Type != typeof(string))
                throw new StiCrystalFormulaException("Only the subscript of a string is supported", node.Position);

            if (!subscript.IsRange)
            {
                var start = $"{Coerce(subscript, typeof(int)).ToOperand()} - 1";
                return new StiCrystalExpression($"Mid({target.Code}, {start}, 1)", typeof(string), isAtomic: true);
            }

            if (subscript.From == null || subscript.To == null || !subscript.IncludeFrom || !subscript.IncludeTo)
                throw new StiCrystalFormulaException("Only a closed range is supported in a subscript", node.Position);

            var from = Coerce(subscript.From, typeof(int)).ToOperand();
            var to = Coerce(subscript.To, typeof(int)).ToOperand();
            return new StiCrystalExpression($"Mid({target.Code}, {from} - 1, {to} - {from} + 1)", typeof(string),
                isAtomic: true);
        }

        /// <summary>
        /// Split(text, delimiter)[1], the text before the first delimiter. The interpretation has no arrays of
        /// strings, so only the first item is taken, by the position of the delimiter.
        /// </summary>
        private StiCrystalExpression TranslateFirstSplitItem(StiCrystalNode node)
        {
            var call = node.Children[0];
            var subscript = node.Children[1];
            if (call.Kind != StiCrystalNodeKind.Call || !string.Equals(call.Text, "split",
                StringComparison.OrdinalIgnoreCase) || call.Children.Count != 2)
                return null;

            if (subscript.Kind != StiCrystalNodeKind.Number || subscript.Text != "1")
                return null;

            var text = TranslateNode(call.Children[0]);
            var delimiter = TranslateNode(call.Children[1]);
            if (text.Type != typeof(string) || string.IsNullOrEmpty(delimiter.Literal))
                return null;

            var position = $"{text.ToOperand()}.IndexOf({delimiter.Code})";
            var code = $"({position} < 0 ? {text.Code} : Left({text.Code}, {position}))";
            return new StiCrystalExpression(code, typeof(string), isAtomic: true);
        }

        /// <summary>
        /// Adds and subtracts numbers, concatenates strings, adds days to a date and subtracts dates.
        /// The difference of two dates is the number of days.
        /// </summary>
        private StiCrystalExpression TranslateAdditive(string op, StiCrystalExpression left, StiCrystalExpression right,
            StiCrystalNode node)
        {
            if (op == "+" && left.Type == typeof(string) && right.Type == typeof(string))
                return new StiCrystalExpression($"{left.ToOperand()} + {right.ToOperand()}", typeof(string));

            if (IsDate(left.Type) && IsNumber(right.Type))
            {
                var number = Coerce(right, typeof(double));
                var days = op == "+" ? number.Code : "-" + number.ToOperand();
                return new StiCrystalExpression($"{left.ToOperand()}.AddDays({days})", typeof(DateTime), isAtomic: true);
            }

            if (op == "+" && IsNumber(left.Type) && IsDate(right.Type))
            {
                var code = $"{right.ToOperand()}.AddDays({Coerce(left, typeof(double)).Code})";
                return new StiCrystalExpression(code, typeof(DateTime), isAtomic: true);
            }

            if (op == "-" && IsDate(left.Type) && IsDate(right.Type))
                return new StiCrystalExpression($"({left.ToOperand()} - {right.ToOperand()}).TotalDays", typeof(double),
                    isAtomic: true);

            return TranslateArithmetic(op, left, right, node);
        }

        private static StiCrystalExpression TranslateArithmetic(string op, StiCrystalExpression left,
            StiCrystalExpression right, StiCrystalNode node)
        {
            RequireNumber(left, node);
            RequireNumber(right, node);

            var type = GetCommonNumberType(left.Type, right.Type);
            return new StiCrystalExpression($"{Coerce(left, type).ToOperand()} {op} {Coerce(right, type).ToOperand()}",
                type);
        }

        /// <summary>
        /// The division in Crystal is never an integer one, so integer operands are converted to double.
        /// </summary>
        private static StiCrystalExpression TranslateDivision(StiCrystalExpression left, StiCrystalExpression right,
            StiCrystalNode node)
        {
            RequireNumber(left, node);
            RequireNumber(right, node);

            var type = GetCommonNumberType(left.Type, right.Type);
            if (type == typeof(decimal))
                return new StiCrystalExpression($"{Coerce(left, type).ToOperand()} / {Coerce(right, type).ToOperand()}",
                    type);

            var dividend = left.Type == typeof(double) ? left.ToOperand() : "(double)" + left.ToOperand();
            return new StiCrystalExpression($"{dividend} / {right.ToOperand()}", typeof(double));
        }

        /// <summary>
        /// If without Else gives the default value of the type of the branch Then, as in Crystal.
        /// </summary>
        private StiCrystalExpression TranslateIf(StiCrystalNode node)
        {
            var condition = TranslateNode(node.Children[0]);
            RequireBoolean(condition, node);

            var thenValue = TranslateStatement(node.Children[1]);
            var elseValue = node.Children.Count > 2 ? TranslateStatement(node.Children[2]) : null;
            if (thenValue == null)
                throw new StiCrystalFormulaException("The branch Then has no value", node.Position);

            if (elseValue == null)
                elseValue = GetDefaultValue(thenValue.Type);

            return CreateConditional(condition, thenValue, elseValue, node);
        }

        private StiCrystalExpression CreateConditional(StiCrystalExpression condition, StiCrystalExpression thenValue,
            StiCrystalExpression elseValue, StiCrystalNode node)
        {
            RequireNoAssignment(thenValue, node);
            RequireNoAssignment(elseValue, node);

            var type = GetCommonType(thenValue, elseValue, node);

            // The parser of the interpretation mode gives the ternary operator the type object, which the logical
            // operators don't accept, so a boolean choice is written with the logical operators only
            if (type == typeof(bool))
                return CreateLogicalChoice(condition, thenValue, elseValue);

            var code = $"{condition.ToOperand()} ? {Coerce(thenValue, type).ToOperand()} : " +
                Coerce(elseValue, type).ToOperand();
            return new StiCrystalExpression(code, type);
        }

        /// <summary>
        /// The boolean choice with the logical operators; the constant branches are folded, so the conditions of
        /// formatting, whose branches are true or false, stay readable.
        /// </summary>
        private static StiCrystalExpression CreateLogicalChoice(StiCrystalExpression condition,
            StiCrystalExpression thenValue, StiCrystalExpression elseValue)
        {
            var thenCode = thenValue.Code;
            var elseCode = elseValue.Code;
            if ((thenCode == "true" || thenCode == "false") && thenCode == elseCode)
                return thenValue;

            if (thenCode == "true" && elseCode == "false")
                return condition;

            if (thenCode == "false" && elseCode == "true")
                return new StiCrystalExpression("!" + condition.ToOperand(), typeof(bool));

            if (thenCode == "true")
                return new StiCrystalExpression($"{condition.ToOperand()} || {elseValue.ToOperand()}", typeof(bool));

            if (thenCode == "false")
                return new StiCrystalExpression($"!{condition.ToOperand()} && {elseValue.ToOperand()}", typeof(bool));

            if (elseCode == "true")
                return new StiCrystalExpression($"!{condition.ToOperand()} || {thenValue.ToOperand()}", typeof(bool));

            if (elseCode == "false")
                return new StiCrystalExpression($"{condition.ToOperand()} && {thenValue.ToOperand()}", typeof(bool));

            var logical = $"{condition.ToOperand()} && {thenValue.ToOperand()} || " +
                $"!{condition.ToOperand()} && {elseValue.ToOperand()}";
            return new StiCrystalExpression(logical, typeof(bool));
        }

        /// <summary>
        /// Select becomes a chain of the ternary operators; a case may list several values and ranges.
        /// </summary>
        private StiCrystalExpression TranslateSelect(StiCrystalNode node)
        {
            var subject = TranslateNode(node.Children[0]);
            var cases = node.Children.Skip(1).ToList();

            StiCrystalExpression result = null;
            var defaultCase = cases.FirstOrDefault(c => c.Text == "default");
            if (defaultCase != null)
                result = TranslateStatement(defaultCase.Children[0]);

            var branches = cases.Where(c => c.Text != "default").Reverse().ToList();
            if (result == null)
            {
                var firstValue = branches.Count > 0 ? TranslateStatement(branches.Last().Children.Last()) : null;
                if (firstValue == null)
                    throw new StiCrystalFormulaException("Select without cases", node.Position);

                result = GetDefaultValue(firstValue.Type);
            }

            foreach (var branch in branches)
            {
                var values = branch.Children.Take(branch.Children.Count - 1).Select(TranslateNode).ToList();
                var condition = TranslateIn(subject, StiCrystalExpression.CreateArray(values), node);
                result = CreateConditional(condition, TranslateStatement(branch.Children.Last()), result, node);
            }

            return result;
        }

        private static Type GetVariableType(string variableType)
        {
            switch (variableType)
            {
                case "numbervar":
                    return typeof(double);

                case "currencyvar":
                    return typeof(decimal);

                case "booleanvar":
                    return typeof(bool);

                case "datevar":
                case "datetimevar":
                case "timevar":
                    return typeof(DateTime);

                default:
                    return typeof(string);
            }
        }

        /// <summary>
        /// The value of an empty Else branch or of a declared variable without a value.
        /// </summary>
        private static StiCrystalExpression GetDefaultValue(Type type)
        {
            if (type == typeof(string))
                return new StiCrystalExpression("\"\"", type, isAtomic: true);

            if (type == typeof(bool))
                return new StiCrystalExpression("false", type, isAtomic: true);

            if (IsDate(type))
                return new StiCrystalExpression("DateSerial(1, 1, 1)", typeof(DateTime), isAtomic: true);

            if (IsNumber(type))
                return Coerce(new StiCrystalExpression("0", typeof(int), isAtomic: true), type);

            throw new StiCrystalFormulaException($"No default value of the type {type.Name}");
        }

        private static Type GetCommonType(StiCrystalExpression first, StiCrystalExpression second, StiCrystalNode node)
        {
            if (first.Type == second.Type)
                return first.Type;

            if (IsNumber(first.Type) && IsNumber(second.Type))
                return GetCommonNumberType(first.Type, second.Type);

            if (IsDate(first.Type) && IsDate(second.Type))
                return typeof(DateTime);

            if (first.Type == typeof(object))
                return second.Type;

            if (second.Type == typeof(object))
                return first.Type;

            throw new StiCrystalFormulaException(
                $"The branches have different types {first.Type.Name} and {second.Type.Name}",
                node.Position);
        }

        private static Type GetCommonNumberType(Type first, Type second)
        {
            if (first == typeof(decimal) || second == typeof(decimal))
                return typeof(decimal);

            if (!IsInteger(first) || !IsInteger(second))
                return typeof(double);

            if (first == typeof(long) || second == typeof(long))
                return typeof(long);

            return typeof(int);
        }

        /// <summary>
        /// Converts a number to the specified numeric type with a cast; other values are returned as is.
        /// </summary>
        private static StiCrystalExpression Coerce(StiCrystalExpression value, Type type)
        {
            if (value.Type == type || !IsNumber(type) || !IsNumber(value.Type) && value.Type != typeof(object))
                return value;

            if (type == typeof(double) && value.Type != typeof(decimal) && value.Type != typeof(object))
                return value;

            if (type == typeof(long) && IsInteger(value.Type) && value.Type != typeof(ulong))
                return value;

            return new StiCrystalExpression($"({GetTypeName(type)}){value.ToOperand()}", type);
        }

        private static string GetTypeName(Type type)
        {
            if (type == typeof(int))
                return "int";

            if (type == typeof(long))
                return "long";

            if (type == typeof(double))
                return "double";

            if (type == typeof(decimal))
                return "decimal";

            if (type == typeof(string))
                return "string";

            if (type == typeof(bool))
                return "bool";

            return type.Name;
        }

        private static bool IsNumber(Type type)
        {
            return type == typeof(int) || type == typeof(long) || type == typeof(double) || type == typeof(decimal) ||
                type == typeof(short) || type == typeof(byte) || type == typeof(float) || type == typeof(uint) ||
                type == typeof(ulong) || type == typeof(ushort) || type == typeof(sbyte);
        }

        private static bool IsInteger(Type type)
        {
            return type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte) ||
                type == typeof(uint) || type == typeof(ulong) || type == typeof(ushort) || type == typeof(sbyte);
        }

        private static bool IsDate(Type type)
        {
            return type == typeof(DateTime);
        }

        private static void RequireNumber(StiCrystalExpression value, StiCrystalNode node)
        {
            if (!IsNumber(value.Type) && value.Type != typeof(object))
                throw new StiCrystalFormulaException($"A number expected instead of {value.Type.Name}", node.Position);
        }

        private static void RequireString(StiCrystalExpression value, StiCrystalNode node)
        {
            if (value.Type != typeof(string) && value.Type != typeof(object))
                throw new StiCrystalFormulaException($"A string expected instead of {value.Type.Name}", node.Position);
        }

        private static void RequireBoolean(StiCrystalExpression value, StiCrystalNode node)
        {
            if (value.Type != typeof(bool) && value.Type != typeof(object))
                throw new StiCrystalFormulaException($"A boolean expected instead of {value.Type.Name}", node.Position);
        }

        private static string ToCSharpString(string text)
        {
            var sb = new StringBuilder("\"");
            foreach (var ch in text)
            {
                switch (ch)
                {
                    case '"':
                        sb.Append("\\\"");
                        break;

                    case '\\':
                        sb.Append("\\\\");
                        break;

                    case '\r':
                        sb.Append("\\r");
                        break;

                    case '\n':
                        sb.Append("\\n");
                        break;

                    case '\t':
                        sb.Append("\\t");
                        break;

                    default:
                        if (char.IsControl(ch))
                            sb.Append($"\\u{(int)ch:X4}");

                        else
                            sb.Append(ch);

                        break;
                }
            }

            return sb.Append('"').ToString();
        }
        #endregion

        public StiCrystalTranslator(IStiCrystalResolver resolver)
        {
            this.resolver = resolver;
        }
    }
}
