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
using System.Text;

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// Splits the text of a formula in Crystal syntax into tokens.
    /// </summary>
    public class StiCrystalLexer
    {
        #region Consts
        /// <summary>
        /// The operators of two characters; they are checked before the operators of one character.
        /// </summary>
        private static readonly string[] LongOperators = { ":=", "<>", "<=", ">=" };

        private const string ShortOperators = "+-*/\\^%&=<>$";
        #endregion

        #region Fields
        private readonly string text;
        private int position;
        #endregion

        #region Methods
        public List<StiCrystalToken> Tokenize()
        {
            var tokens = new List<StiCrystalToken>();
            while (true)
            {
                SkipSpacesAndComments();

                if (position >= text.Length)
                {
                    tokens.Add(new StiCrystalToken(StiCrystalTokenType.End, string.Empty, position));
                    return tokens;
                }

                tokens.Add(ReadToken());
            }
        }

        private void SkipSpacesAndComments()
        {
            while (position < text.Length)
            {
                var ch = text[position];
                if (char.IsWhiteSpace(ch))
                {
                    position++;
                }
                else if (ch == '/' && position + 1 < text.Length && text[position + 1] == '/')
                {
                    while (position < text.Length && text[position] != '\r' && text[position] != '\n')
                    {
                        position++;
                    }
                }
                else
                {
                    return;
                }
            }
        }

        private StiCrystalToken ReadToken()
        {
            var start = position;
            var ch = text[position];

            if (char.IsLetter(ch) || ch == '_')
                return ReadWord();

            if (char.IsDigit(ch) || (ch == '.' && position + 1 < text.Length && char.IsDigit(text[position + 1])))
                return ReadNumber();

            switch (ch)
            {
                case '"':
                case '\'':
                    return ReadString(ch);

                case '{':
                    return ReadEnclosed('}', StiCrystalTokenType.Reference);

                case '#':
                    return ReadEnclosed('#', StiCrystalTokenType.Date);

                case ',':
                    position++;
                    return new StiCrystalToken(StiCrystalTokenType.Comma, ",", start);

                case ';':
                    position++;
                    return new StiCrystalToken(StiCrystalTokenType.Semicolon, ";", start);

                case '(':
                    position++;
                    return new StiCrystalToken(StiCrystalTokenType.LeftParenthesis, "(", start);

                case ')':
                    position++;
                    return new StiCrystalToken(StiCrystalTokenType.RightParenthesis, ")", start);

                case '[':
                    position++;
                    return new StiCrystalToken(StiCrystalTokenType.LeftBracket, "[", start);

                case ']':
                    position++;
                    return new StiCrystalToken(StiCrystalTokenType.RightBracket, "]", start);
            }

            foreach (var op in LongOperators)
            {
                if (string.CompareOrdinal(text, position, op, 0, op.Length) == 0)
                {
                    position += op.Length;
                    return new StiCrystalToken(StiCrystalTokenType.Operator, op, start);
                }
            }

            if (ch == ':')
            {
                position++;
                return new StiCrystalToken(StiCrystalTokenType.Colon, ":", start);
            }

            if (ShortOperators.IndexOf(ch) >= 0)
            {
                position++;
                return new StiCrystalToken(StiCrystalTokenType.Operator, ch.ToString(), start);
            }

            throw new StiCrystalFormulaException($"Unexpected character '{ch}'", start);
        }

        private StiCrystalToken ReadWord()
        {
            var start = position;
            while (position < text.Length && (char.IsLetterOrDigit(text[position]) || text[position] == '_'))
            {
                position++;
            }

            return new StiCrystalToken(StiCrystalTokenType.Identifier, text.Substring(start, position - start), start);
        }

        private StiCrystalToken ReadNumber()
        {
            var start = position;
            while (position < text.Length && char.IsDigit(text[position]))
            {
                position++;
            }

            if (position + 1 < text.Length && text[position] == '.' && char.IsDigit(text[position + 1]))
            {
                position++;
                while (position < text.Length && char.IsDigit(text[position]))
                {
                    position++;
                }
            }

            return new StiCrystalToken(StiCrystalTokenType.Number, text.Substring(start, position - start), start);
        }

        /// <summary>
        /// Reads a string in single or double quotes; a doubled quote inside the string stands for one quote.
        /// </summary>
        private StiCrystalToken ReadString(char quote)
        {
            var start = position;
            var sb = new StringBuilder();
            position++;
            while (true)
            {
                if (position >= text.Length)
                    throw new StiCrystalFormulaException("Unterminated string", start);

                var ch = text[position++];
                if (ch != quote)
                {
                    sb.Append(ch);
                    continue;
                }

                if (position < text.Length && text[position] == quote)
                {
                    sb.Append(quote);
                    position++;
                    continue;
                }

                return new StiCrystalToken(StiCrystalTokenType.String, sb.ToString(), start);
            }
        }

        private StiCrystalToken ReadEnclosed(char closing, StiCrystalTokenType type)
        {
            var start = position;
            var end = text.IndexOf(closing, position + 1);
            if (end < 0)
                throw new StiCrystalFormulaException($"Missing '{closing}'", start);

            position = end + 1;
            return new StiCrystalToken(type, text.Substring(start + 1, end - start - 1), start);
        }
        #endregion

        public StiCrystalLexer(string text)
        {
            this.text = text ?? string.Empty;
        }
    }
}
