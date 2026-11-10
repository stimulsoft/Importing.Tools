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
    /// Builds the syntax tree of a formula in Crystal syntax by recursive descent. The precedence of the operators
    /// from the highest to the lowest: negation, ^, * / %, \, mod, + - &amp;, ranges, comparisons and in, like,
    /// startswith, not, and, or, xor, eqv, imp. If, Select and the assignment have the lowest precedence.
    /// </summary>
    public class StiCrystalParser
    {
        #region Consts
        private static readonly string[] Scopes = { "local", "global", "shared" };

        private static readonly string[] VariableTypes =
        {
            "numbervar", "currencyvar", "stringvar", "booleanvar", "datevar", "timevar", "datetimevar"
        };

        private static readonly string[] Comparisons = { "=", "<>", "<", ">", "<=", ">=" };

        private static readonly string[] RangeWords = { "to", "_to", "to_", "_to_" };

        private static readonly string[] OpenRangeWords = { "upfrom", "upfrom_", "upto", "upto_" };
        #endregion

        #region Fields
        private readonly List<StiCrystalToken> tokens;
        private int index;
        #endregion

        #region Properties
        private StiCrystalToken Current
        {
            get
            {
                return tokens[index];
            }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Parses the text of a formula and returns the tree; a formula of several statements gives a Sequence node.
        /// </summary>
        public static StiCrystalNode Parse(string formula)
        {
            var parser = new StiCrystalParser(new StiCrystalLexer(formula).Tokenize());
            var node = parser.ParseSequence();
            if (parser.Current.Type != StiCrystalTokenType.End)
                throw new StiCrystalFormulaException($"Unexpected '{parser.Current.Text}'", parser.Current.Position);

            return node;
        }

        private StiCrystalNode ParseSequence()
        {
            var position = Current.Position;
            var statements = new List<StiCrystalNode>();
            while (Current.Type != StiCrystalTokenType.End && Current.Type != StiCrystalTokenType.RightParenthesis)
            {
                if (Current.Type == StiCrystalTokenType.Semicolon)
                {
                    index++;
                    continue;
                }

                statements.Add(ParseExpression());

                if (Current.Type != StiCrystalTokenType.Semicolon)
                    break;
            }

            if (statements.Count == 1)
                return statements[0];

            return new StiCrystalNode(StiCrystalNodeKind.Sequence, null, position, statements.ToArray());
        }

        private StiCrystalNode ParseExpression()
        {
            if (Current.IsWord("if"))
                return ParseIf();

            if (Current.IsWord("select"))
                return ParseSelect();

            if (IsDeclaration())
                return ParseDeclaration();

            if (Current.IsWord("for"))
                return ParseFor();

            if (Current.IsWord("while"))
                return ParseWhile();

            if (Current.IsWord("do"))
                return ParseDoWhile();

            if (Current.IsWord("exit"))
            {
                var position = Current.Position;
                index++;
                var loop = Current.Text;
                Expect(StiCrystalTokenType.Identifier);
                return new StiCrystalNode(StiCrystalNodeKind.Exit, loop.ToLowerInvariant(), position);
            }

            var left = ParseImp();
            if (Current.IsOperator(":="))
            {
                var position = Current.Position;
                index++;
                return new StiCrystalNode(StiCrystalNodeKind.Assign, ":=", position, left, ParseExpression());
            }

            return left;
        }

        private StiCrystalNode ParseIf()
        {
            var position = Current.Position;
            index++;
            var condition = ParseExpression();
            ExpectWord("then");
            var node = new StiCrystalNode(StiCrystalNodeKind.If, null, position, condition, ParseExpression());

            if (Current.IsWord("else"))
            {
                index++;
                node.Add(ParseExpression());
            }

            return node;
        }

        /// <summary>
        /// Select subject Case values: result Case values: result Default: result. Each case gives a Case node,
        /// its last child is the result; the Default branch gives a Case node with the only child.
        /// </summary>
        private StiCrystalNode ParseSelect()
        {
            var position = Current.Position;
            index++;
            var node = new StiCrystalNode(StiCrystalNodeKind.Select, null, position, ParseImp());

            while (Current.IsWord("case"))
            {
                var casePosition = Current.Position;
                index++;
                var caseNode = new StiCrystalNode(StiCrystalNodeKind.Case, null, casePosition);
                caseNode.Add(ParseImp());

                while (Current.Type == StiCrystalTokenType.Comma)
                {
                    index++;
                    caseNode.Add(ParseImp());
                }

                Expect(StiCrystalTokenType.Colon);
                caseNode.Add(ParseExpression());
                node.Add(caseNode);
            }

            if (Current.IsWord("default"))
            {
                var defaultPosition = Current.Position;
                index++;
                Expect(StiCrystalTokenType.Colon);
                node.Add(new StiCrystalNode(StiCrystalNodeKind.Case, "default", defaultPosition, ParseExpression()));
            }

            return node;
        }

        private bool IsDeclaration()
        {
            if (Current.Type != StiCrystalTokenType.Identifier)
                return false;

            var word = Current.Text.ToLowerInvariant();
            return Scopes.Contains(word) || VariableTypes.Contains(word);
        }

        /// <summary>
        /// [Local | Global | Shared] TypeVar [Array | Range] name [:= value]; the scope is Global by default.
        /// </summary>
        private StiCrystalNode ParseDeclaration()
        {
            var position = Current.Position;
            var scope = "global";
            if (Scopes.Contains(Current.Text.ToLowerInvariant()))
            {
                scope = Current.Text.ToLowerInvariant();
                index++;
            }

            if (Current.Type != StiCrystalTokenType.Identifier ||
                !VariableTypes.Contains(Current.Text.ToLowerInvariant()))
                throw new StiCrystalFormulaException("The type of the variable expected", Current.Position);

            var variableType = Current.Text.ToLowerInvariant();
            index++;

            var isArray = false;
            var isRange = false;
            if (Current.IsWord("array"))
            {
                isArray = true;
                index++;
            }

            else if (Current.IsWord("range"))
            {
                isRange = true;
                index++;
            }

            var nameIndex = index;
            var name = Current.Text;
            Expect(StiCrystalTokenType.Identifier);

            var node = new StiCrystalNode(StiCrystalNodeKind.Declare, name, position)
            {
                Scope = scope,
                VariableType = variableType,
                IsArrayVariable = isArray,
                IsRangeVariable = isRange
            };

            if (Current.IsOperator(":="))
            {
                index++;
                node.Add(ParseExpression());
            }

            // A declaration gives the value of the variable, so an operator may follow it: "numbervar header <> 1",
            // "BooleanVar Suppress and ..." in the conditions of suppression. The rest starts again from the name.
            else if (IsBinaryOperator())
            {
                index = nameIndex;
                return new StiCrystalNode(StiCrystalNodeKind.Sequence, null, position, node, ParseExpression());
            }

            return node;
        }

        private bool IsBinaryOperator()
        {
            if (Current.Type == StiCrystalTokenType.Operator)
                return true;

            return Current.IsWord("and") || Current.IsWord("or") || Current.IsWord("xor") || Current.IsWord("eqv") ||
                Current.IsWord("imp") || Current.IsWord("in") || Current.IsWord("like") ||
                Current.IsWord("startswith") || Current.IsWord("mod");
        }

        /// <summary>
        /// For i := from To to [Step step] Do body; the children are from, to, step (or null) and the body.
        /// </summary>
        private StiCrystalNode ParseFor()
        {
            var position = Current.Position;
            index++;
            var name = Current.Text;
            Expect(StiCrystalTokenType.Identifier);
            ExpectOperator(":=");
            var from = ParseImp();
            ExpectWord("to");
            var to = ParseImp();

            StiCrystalNode step = null;
            if (Current.IsWord("step"))
            {
                index++;
                step = ParseImp();
            }

            ExpectWord("do");
            return new StiCrystalNode(StiCrystalNodeKind.For, name, position, from, to, step, ParseExpression());
        }

        private StiCrystalNode ParseWhile()
        {
            var position = Current.Position;
            index++;
            var condition = ParseImp();
            ExpectWord("do");
            return new StiCrystalNode(StiCrystalNodeKind.While, null, position, condition, ParseExpression());
        }

        private StiCrystalNode ParseDoWhile()
        {
            var position = Current.Position;
            index++;
            var body = ParseExpression();
            ExpectWord("while");
            return new StiCrystalNode(StiCrystalNodeKind.DoWhile, null, position, body, ParseImp());
        }

        private StiCrystalNode ParseImp()
        {
            return ParseBinary(ParseEqv, "imp");
        }

        private StiCrystalNode ParseEqv()
        {
            return ParseBinary(ParseXor, "eqv");
        }

        private StiCrystalNode ParseXor()
        {
            return ParseBinary(ParseOr, "xor");
        }

        private StiCrystalNode ParseOr()
        {
            return ParseBinary(ParseAnd, "or");
        }

        private StiCrystalNode ParseAnd()
        {
            return ParseBinary(ParseNot, "and");
        }

        private StiCrystalNode ParseBinary(Func<StiCrystalNode> next, string word)
        {
            var left = next();
            while (Current.IsWord(word))
            {
                var position = Current.Position;
                index++;
                left = new StiCrystalNode(StiCrystalNodeKind.Binary, word, position, left, next());
            }

            return left;
        }

        private StiCrystalNode ParseNot()
        {
            if (!Current.IsWord("not"))
                return ParseComparison();

            var position = Current.Position;
            index++;
            return new StiCrystalNode(StiCrystalNodeKind.Unary, "not", position, ParseNot());
        }

        private StiCrystalNode ParseComparison()
        {
            var left = ParseRange();
            while (true)
            {
                var position = Current.Position;
                string op = null;
                if (Current.Type == StiCrystalTokenType.Operator && Comparisons.Contains(Current.Text))
                    op = Current.Text;

                else if (Current.IsWord("in") || Current.IsWord("like") || Current.IsWord("startswith"))
                    op = Current.Text.ToLowerInvariant();

                if (op == null)
                    return left;

                index++;
                left = new StiCrystalNode(StiCrystalNodeKind.Binary, op, position, left, ParseRange());
            }
        }

        private StiCrystalNode ParseRange()
        {
            var position = Current.Position;
            if (Current.Type == StiCrystalTokenType.Identifier &&
                OpenRangeWords.Contains(Current.Text.ToLowerInvariant()))
            {
                var word = Current.Text.ToLowerInvariant();
                index++;
                return new StiCrystalNode(StiCrystalNodeKind.Range, word, position, ParseAdditive());
            }

            var left = ParseAdditive();
            if (Current.Type == StiCrystalTokenType.Identifier && RangeWords.Contains(Current.Text.ToLowerInvariant()))
            {
                var word = Current.Text.ToLowerInvariant();
                position = Current.Position;
                index++;
                return new StiCrystalNode(StiCrystalNodeKind.Range, word, position, left, ParseAdditive());
            }

            return left;
        }

        private StiCrystalNode ParseAdditive()
        {
            var left = ParseMod();
            while (Current.IsOperator("+") || Current.IsOperator("-") || Current.IsOperator("&"))
            {
                var token = Current;
                index++;
                left = new StiCrystalNode(StiCrystalNodeKind.Binary, token.Text, token.Position, left, ParseMod());
            }

            return left;
        }

        private StiCrystalNode ParseMod()
        {
            var left = ParseIntegerDivision();
            while (Current.IsWord("mod"))
            {
                var position = Current.Position;
                index++;
                left = new StiCrystalNode(StiCrystalNodeKind.Binary, "mod", position, left, ParseIntegerDivision());
            }

            return left;
        }

        private StiCrystalNode ParseIntegerDivision()
        {
            var left = ParseMultiplicative();
            while (Current.IsOperator("\\"))
            {
                var position = Current.Position;
                index++;
                left = new StiCrystalNode(StiCrystalNodeKind.Binary, "\\", position, left, ParseMultiplicative());
            }

            return left;
        }

        private StiCrystalNode ParseMultiplicative()
        {
            var left = ParsePower();
            while (Current.IsOperator("*") || Current.IsOperator("/") || Current.IsOperator("%"))
            {
                var token = Current;
                index++;
                left = new StiCrystalNode(StiCrystalNodeKind.Binary, token.Text, token.Position, left, ParsePower());
            }

            return left;
        }

        private StiCrystalNode ParsePower()
        {
            var left = ParseUnary();
            while (Current.IsOperator("^"))
            {
                var position = Current.Position;
                index++;
                left = new StiCrystalNode(StiCrystalNodeKind.Binary, "^", position, left, ParseUnary());
            }

            return left;
        }

        private StiCrystalNode ParseUnary()
        {
            if (Current.IsOperator("-") || Current.IsOperator("+") || Current.IsOperator("$"))
            {
                var token = Current;
                index++;
                return new StiCrystalNode(StiCrystalNodeKind.Unary, token.Text, token.Position, ParseUnary());
            }

            return ParsePostfix();
        }

        private StiCrystalNode ParsePostfix()
        {
            var node = ParsePrimary();
            while (Current.Type == StiCrystalTokenType.LeftBracket)
            {
                var position = Current.Position;
                index++;
                var subscript = ParseRange();
                Expect(StiCrystalTokenType.RightBracket);
                node = new StiCrystalNode(StiCrystalNodeKind.Index, null, position, node, subscript);
            }

            return node;
        }

        private StiCrystalNode ParsePrimary()
        {
            var token = Current;
            switch (token.Type)
            {
                case StiCrystalTokenType.Number:
                    index++;
                    return new StiCrystalNode(StiCrystalNodeKind.Number, token.Text, token.Position);

                case StiCrystalTokenType.String:
                    index++;
                    return new StiCrystalNode(StiCrystalNodeKind.String, token.Text, token.Position);

                case StiCrystalTokenType.Date:
                    index++;
                    return new StiCrystalNode(StiCrystalNodeKind.Date, token.Text, token.Position);

                case StiCrystalTokenType.Reference:
                    index++;
                    return new StiCrystalNode(StiCrystalNodeKind.Reference, token.Text.Trim(), token.Position);

                case StiCrystalTokenType.LeftParenthesis:
                    index++;
                    var inner = ParseSequence();
                    Expect(StiCrystalTokenType.RightParenthesis);
                    return inner;

                case StiCrystalTokenType.LeftBracket:
                    return ParseArray();

                case StiCrystalTokenType.Identifier:
                    return ParseIdentifier();
            }

            throw new StiCrystalFormulaException($"Unexpected '{token.Text}'", token.Position);
        }

        private StiCrystalNode ParseArray()
        {
            var node = new StiCrystalNode(StiCrystalNodeKind.Array, null, Current.Position);
            index++;
            if (Current.Type != StiCrystalTokenType.RightBracket)
            {
                node.Add(ParseRange());
                while (Current.Type == StiCrystalTokenType.Comma)
                {
                    index++;
                    node.Add(ParseRange());
                }
            }

            Expect(StiCrystalTokenType.RightBracket);
            return node;
        }

        private StiCrystalNode ParseIdentifier()
        {
            if (Current.IsWord("if"))
                return ParseIf();

            if (Current.IsWord("select"))
                return ParseSelect();

            var token = Current;
            index++;

            if (token.IsWord("true") || token.IsWord("false"))
                return new StiCrystalNode(StiCrystalNodeKind.Boolean, token.Text.ToLowerInvariant(), token.Position);

            if (Current.Type != StiCrystalTokenType.LeftParenthesis)
                return new StiCrystalNode(StiCrystalNodeKind.Identifier, token.Text, token.Position);

            index++;
            var call = new StiCrystalNode(StiCrystalNodeKind.Call, token.Text, token.Position);
            if (Current.Type != StiCrystalTokenType.RightParenthesis)
            {
                call.Add(ParseExpression());
                while (Current.Type == StiCrystalTokenType.Comma)
                {
                    index++;
                    call.Add(ParseExpression());
                }
            }

            Expect(StiCrystalTokenType.RightParenthesis);
            return call;
        }

        private void Expect(StiCrystalTokenType type)
        {
            if (Current.Type != type)
                throw new StiCrystalFormulaException($"{type} expected instead of '{Current.Text}'", Current.Position);

            index++;
        }

        private void ExpectWord(string word)
        {
            if (!Current.IsWord(word))
                throw new StiCrystalFormulaException($"'{word}' expected instead of '{Current.Text}'", Current.Position);

            index++;
        }

        private void ExpectOperator(string op)
        {
            if (!Current.IsOperator(op))
                throw new StiCrystalFormulaException($"'{op}' expected instead of '{Current.Text}'", Current.Position);

            index++;
        }
        #endregion

        private StiCrystalParser(List<StiCrystalToken> tokens)
        {
            this.tokens = tokens;
        }
    }
}
