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

using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// The statements of a formula in Basic syntax: declarations (Dim, Local, Global, Shared ... As Type),
    /// assignments (Formula = ..., x = ...), If ... Then ... ElseIf ... Else ... End If in a block or in one line,
    /// Select Case with Case Is, Case a To b and Case Else. The expressions are the expressions of Crystal syntax;
    /// a string is in double quotes, the comments and the ends of lines are dropped by the lexer. The statements are
    /// turned into one expression by StiCrystalBasicLowering. Loops, arrays and own functions are not supported.
    /// </summary>
    public partial class StiCrystalParser
    {
        #region Consts.Basic
        private static readonly string[] BasicScopes = { "dim", "local", "global", "shared" };

        private static readonly Dictionary<string, string> BasicTypes = new Dictionary<string, string>
        {
            { "number", "numbervar" },
            { "currency", "currencyvar" },
            { "string", "stringvar" },
            { "boolean", "booleanvar" },
            { "date", "datevar" },
            { "time", "timevar" },
            { "datetime", "datetimevar" }
        };

        private static readonly string[] BasicLoops = { "for", "do", "while", "loop", "next", "wend", "exit" };

        private static readonly Regex BasicFormulaAssignment = new Regex(@"^\s*formula\s*=",
            RegexOptions.IgnoreCase | RegexOptions.Multiline | RegexOptions.Compiled);
        #endregion

        #region Methods.Basic
        /// <summary>
        /// Parses a formula in Basic syntax and turns its statements into one expression of the tree of Crystal.
        /// </summary>
        public static StiCrystalNode ParseBasic(string formula)
        {
            var parser = new StiCrystalParser(new StiCrystalLexer(formula, isBasic: true).Tokenize());
            var block = parser.ParseBasicBlock();
            if (parser.Current.Type != StiCrystalTokenType.End)
                throw new StiCrystalFormulaException($"Unexpected '{parser.Current.Text}'", parser.Current.Position);

            return new StiCrystalBasicLowering().Lower(block);
        }

        /// <summary>
        /// True if the text is a formula in Basic syntax: a line assigns Formula. The conditions of the formatting
        /// come without their syntax, and such a line is not valid in Crystal syntax (G0318: the styles and the
        /// suppression of the areas in Basic).
        /// </summary>
        public static bool LooksLikeBasic(string formula)
        {
            return formula != null && BasicFormulaAssignment.IsMatch(formula);
        }

        /// <summary>
        /// Parses a formula in the syntax of Crystal or of Basic.
        /// </summary>
        public static StiCrystalNode Parse(string formula, bool isBasic)
        {
            return isBasic ? ParseBasic(formula) : Parse(formula);
        }

        /// <summary>
        /// The statements up to the end of the formula or up to one of the words which end the block.
        /// </summary>
        private StiCrystalNode ParseBasicBlock(params string[] endWords)
        {
            var block = new StiCrystalNode(StiCrystalNodeKind.Sequence, null, Current.Position);
            while (true)
            {
                SkipBasicSeparators();
                if (Current.Type == StiCrystalTokenType.End || endWords.Any(Current.IsWord))
                    return block;

                block.Add(ParseBasicStatement());
            }
        }

        private void SkipBasicSeparators()
        {
            while (Current.Type == StiCrystalTokenType.NewLine || Current.Type == StiCrystalTokenType.Colon)
            {
                index++;
            }
        }

        private bool IsBasicLineEnd()
        {
            return Current.Type == StiCrystalTokenType.NewLine || Current.Type == StiCrystalTokenType.End;
        }

        private StiCrystalNode ParseBasicStatement()
        {
            var position = Current.Position;
            if (BasicScopes.Any(Current.IsWord))
                return ParseBasicDeclaration();

            if (Current.IsWord("if"))
                return ParseBasicIf();

            if (Current.IsWord("select"))
                return ParseBasicSelect();

            if (Current.IsWord("redim"))
                throw new StiCrystalFormulaException("The arrays of Basic are not supported", position);

            if (BasicLoops.Any(Current.IsWord))
                throw new StiCrystalFormulaException("The loops of Basic are not supported", position);

            if (Current.Type == StiCrystalTokenType.Identifier && tokens[index + 1].IsOperator("="))
            {
                var name = Current.Text;
                index += 2;
                var target = new StiCrystalNode(StiCrystalNodeKind.Identifier, name, position);
                return new StiCrystalNode(StiCrystalNodeKind.Assign, "=", position, target, ParseImp());
            }

            var expression = ParseImp();
            if (Current.IsOperator("="))
                throw new StiCrystalFormulaException("The arrays of Basic are not supported", position);

            return expression;
        }

        /// <summary>
        /// Dim a As Number, b As String = "x"; Dim c, d As Number gives the type to both; Dim a() As String is an
        /// array. Dim is a local variable.
        /// </summary>
        private StiCrystalNode ParseBasicDeclaration()
        {
            var position = Current.Position;
            var scope = Current.Text.ToLowerInvariant();
            if (scope == "dim")
                scope = "local";

            index++;
            var declarations = new List<StiCrystalNode>();
            var untyped = new List<StiCrystalNode>();
            while (true)
            {
                var name = Current.Text;
                var namePosition = Current.Position;
                Expect(StiCrystalTokenType.Identifier);

                var node = new StiCrystalNode(StiCrystalNodeKind.Declare, name, namePosition) { Scope = scope };
                if (Current.Type == StiCrystalTokenType.LeftParenthesis)
                {
                    while (Current.Type != StiCrystalTokenType.RightParenthesis && !IsBasicLineEnd())
                    {
                        index++;
                    }

                    Expect(StiCrystalTokenType.RightParenthesis);
                    node.IsArrayVariable = true;
                }

                untyped.Add(node);
                declarations.Add(node);

                if (Current.IsWord("as"))
                {
                    index++;
                    string type;
                    if (Current.Type != StiCrystalTokenType.Identifier ||
                        !BasicTypes.TryGetValue(Current.Text.ToLowerInvariant(), out type))
                        throw new StiCrystalFormulaException("The type of the variable expected", Current.Position);

                    index++;
                    foreach (var declaration in untyped)
                    {
                        declaration.VariableType = type;
                    }

                    untyped.Clear();
                }

                if (Current.IsOperator("="))
                {
                    index++;
                    node.Add(ParseImp());
                }

                if (Current.Type != StiCrystalTokenType.Comma)
                    break;

                index++;
            }

            if (untyped.Count > 0)
                throw new StiCrystalFormulaException("The type of the variable expected", Current.Position);

            return declarations.Count == 1
                ? declarations[0]
                : new StiCrystalNode(StiCrystalNodeKind.Sequence, null, position, declarations.ToArray());
        }

        /// <summary>
        /// If in a block: If c Then / statements / ElseIf c Then / statements / Else / statements / End If, every
        /// ElseIf becomes an If in the branch Else. In one line: If c Then statements [Else statements].
        /// </summary>
        private StiCrystalNode ParseBasicIf()
        {
            var position = Current.Position;
            index++;
            var condition = ParseImp();
            ExpectWord("then");

            if (!IsBasicLineEnd())
                return ParseBasicLineIf(condition, position);

            var node = new StiCrystalNode(StiCrystalNodeKind.If, null, position, condition,
                ParseBasicBlock("elseif", "else", "end"));
            var last = node;
            while (Current.IsWord("elseif"))
            {
                var elseIfPosition = Current.Position;
                index++;
                var elseIfCondition = ParseImp();
                ExpectWord("then");
                var elseIf = new StiCrystalNode(StiCrystalNodeKind.If, null, elseIfPosition, elseIfCondition,
                    ParseBasicBlock("elseif", "else", "end"));
                last.Add(elseIf);
                last = elseIf;
            }

            if (Current.IsWord("else"))
            {
                index++;
                last.Add(ParseBasicBlock("end"));
            }

            ExpectWord("end");
            ExpectWord("if");
            return node;
        }

        private StiCrystalNode ParseBasicLineIf(StiCrystalNode condition, int position)
        {
            var node = new StiCrystalNode(StiCrystalNodeKind.If, null, position, condition, ParseBasicLine());
            if (Current.IsWord("else"))
            {
                index++;
                node.Add(ParseBasicLine());
            }

            if (Current.IsWord("end") && tokens[index + 1].IsWord("if"))
                index += 2;

            return node;
        }

        /// <summary>
        /// The statements of one line, separated by colons, up to the end of the line or Else.
        /// </summary>
        private StiCrystalNode ParseBasicLine()
        {
            var block = new StiCrystalNode(StiCrystalNodeKind.Sequence, null, Current.Position);
            while (!IsBasicLineEnd() && !Current.IsWord("else") && !Current.IsWord("end"))
            {
                if (Current.Type == StiCrystalTokenType.Colon)
                {
                    index++;
                    continue;
                }

                block.Add(ParseBasicStatement());
            }

            return block;
        }

        /// <summary>
        /// Select Case subject / Case values / statements / Case Else / statements / End Select. A value is an
        /// expression, a range a To b or Is with a comparison; the condition of every Case is built here.
        /// </summary>
        private StiCrystalNode ParseBasicSelect()
        {
            var position = Current.Position;
            index++;
            ExpectWord("case");
            var subject = ParseImp();

            StiCrystalNode first = null;
            StiCrystalNode last = null;
            StiCrystalNode otherwise = null;
            while (true)
            {
                SkipBasicSeparators();
                if (!Current.IsWord("case"))
                    break;

                var casePosition = Current.Position;
                index++;
                if (Current.IsWord("else"))
                {
                    index++;
                    otherwise = ParseBasicBlock("case", "end");
                    continue;
                }

                var condition = ParseBasicCaseCondition(subject);
                var node = new StiCrystalNode(StiCrystalNodeKind.If, null, casePosition, condition,
                    ParseBasicBlock("case", "end"));

                if (first == null)
                    first = node;

                else
                    last.Add(node);

                last = node;
            }

            ExpectWord("end");
            ExpectWord("select");

            if (first == null)
                return otherwise ?? new StiCrystalNode(StiCrystalNodeKind.Sequence, null, position);

            if (otherwise != null)
                last.Add(otherwise);

            return first;
        }

        private StiCrystalNode ParseBasicCaseCondition(StiCrystalNode subject)
        {
            StiCrystalNode condition = null;
            while (true)
            {
                var position = Current.Position;
                StiCrystalNode item;
                if (Current.IsWord("is"))
                {
                    index++;
                    var op = Current.Text;
                    if (Current.Type != StiCrystalTokenType.Operator || !Comparisons.Contains(op))
                        throw new StiCrystalFormulaException($"A comparison expected instead of '{op}'", position);

                    index++;
                    item = new StiCrystalNode(StiCrystalNodeKind.Binary, op, position, subject, ParseRange());
                }
                else
                {
                    var value = ParseRange();
                    var op = value.Kind == StiCrystalNodeKind.Range ? "in" : "=";
                    item = new StiCrystalNode(StiCrystalNodeKind.Binary, op, position, subject, value);
                }

                condition = condition == null
                    ? item
                    : new StiCrystalNode(StiCrystalNodeKind.Binary, "or", position, condition, item);

                if (Current.Type != StiCrystalTokenType.Comma)
                    return condition;

                index++;
            }
        }
        #endregion
    }
}
