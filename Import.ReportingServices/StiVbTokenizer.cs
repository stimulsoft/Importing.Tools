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

namespace Stimulsoft.Report.Import.ReportingServices
{
    /// <summary>
    /// Splits an expression or the code of Visual Basic into tokens. Every token keeps its text as in
    /// the source, so the tokens joined together give the source back, and a converter changes only
    /// the tokens it knows. A string, a comment or a date is one token, so nothing inside them is
    /// taken for a name.
    /// </summary>
    internal static class StiVbTokenizer
    {
        #region Consts
        private static readonly string[] TwoCharOperators = { "<>", "<=", ">=", ":=", "&=", "+=", "-=", "*=", "/=",
            "\\=", "^=", "<<", ">>" };
        #endregion

        #region Methods
        public static List<StiVbToken> Tokenize(string text)
        {
            var tokens = new List<StiVbToken>();
            var index = 0;

            while (index < text.Length)
            {
                var start = index;
                var type = ReadToken(text, ref index);

                tokens.Add(new StiVbToken(type, text.Substring(start, index - start)));
            }

            return tokens;
        }

        private static StiVbTokenType ReadToken(string text, ref int index)
        {
            var ch = text[index];

            if (ch == '\r' || ch == '\n')
            {
                index++;

                if (ch == '\r' && index < text.Length && text[index] == '\n')
                    index++;

                return StiVbTokenType.NewLine;
            }

            if (IsSpace(ch))
            {
                SkipWhile(text, ref index, IsSpace);
                return StiVbTokenType.Whitespace;
            }

            if (ch == '"')
            {
                ReadString(text, ref index);
                return StiVbTokenType.String;
            }

            if (ch == '\'')
            {
                SkipToEndOfLine(text, ref index);
                return StiVbTokenType.Comment;
            }

            if (ch == '[')
            {
                var end = text.IndexOf(']', index);
                index = end == -1 ? text.Length : end + 1;
                return StiVbTokenType.Identifier;
            }

            if (char.IsLetter(ch) || ch == '_')
                return ReadIdentifier(text, ref index);

            if (char.IsDigit(ch) || ch == '.' && index + 1 < text.Length && char.IsDigit(text[index + 1]))
            {
                ReadNumber(text, ref index);
                return StiVbTokenType.Number;
            }

            if (ch == '&' && index + 2 < text.Length && IsRadixPrefix(text[index + 1]) &&
                char.IsLetterOrDigit(text[index + 2]))
            {
                index += 2;
                SkipWhile(text, ref index, char.IsLetterOrDigit);
                return StiVbTokenType.Number;
            }

            if (ch == '#' && TryReadDate(text, ref index))
                return StiVbTokenType.Date;

            index += IsTwoCharOperator(text, index) ? 2 : 1;
            return StiVbTokenType.Operator;
        }

        private static StiVbTokenType ReadIdentifier(string text, ref int index)
        {
            var start = index;
            SkipWhile(text, ref index, IsIdentifierChar);

            //the type character of a string, as in Left$(text, 2)
            if (index < text.Length && text[index] == '$')
                index++;

            if (index - start == 3 && string.Compare(text, start, "REM", 0, 3, true) == 0 &&
                (index == text.Length || IsSpace(text[index]) || text[index] == '\r' || text[index] == '\n'))
            {
                SkipToEndOfLine(text, ref index);
                return StiVbTokenType.Comment;
            }

            return StiVbTokenType.Identifier;
        }

        /// <summary>
        /// Reads a string up to its closing quote, a doubled quote does not close it. The type
        /// character of a char, as in "a"c, belongs to the string.
        /// </summary>
        private static void ReadString(string text, ref int index)
        {
            index++;

            while (index < text.Length)
            {
                if (text[index] == '"')
                {
                    if (index + 1 < text.Length && text[index + 1] == '"')
                    {
                        index += 2;
                        continue;
                    }

                    index++;

                    if (index < text.Length && (text[index] == 'c' || text[index] == 'C') &&
                        (index + 1 == text.Length || !IsIdentifierChar(text[index + 1])))
                        index++;

                    return;
                }

                index++;
            }
        }

        /// <summary>
        /// Reads a number with its fraction, exponent and type character, as 1.5E+3 or 10D.
        /// </summary>
        private static void ReadNumber(string text, ref int index)
        {
            index++;

            while (index < text.Length)
            {
                var ch = text[index];
                var previous = text[index - 1];

                var isExponentSign = (ch == '+' || ch == '-') && (previous == 'E' || previous == 'e') &&
                    index + 1 < text.Length && char.IsDigit(text[index + 1]);

                if (!char.IsLetterOrDigit(ch) && ch != '.' && !isExponentSign)
                    break;

                if (ch == '.' && (index + 1 == text.Length || !char.IsDigit(text[index + 1])))
                    break;

                index++;
            }
        }

        /// <summary>
        /// Reads a date such as #1/31/2024#. A number sign without a closing one on the same line, or
        /// with no digit after it, is an operator.
        /// </summary>
        private static bool TryReadDate(string text, ref int index)
        {
            if (index + 1 >= text.Length || !char.IsDigit(text[index + 1]))
                return false;

            for (var end = index + 1; end < text.Length; end++)
            {
                var ch = text[end];

                if (ch == '\r' || ch == '\n')
                    return false;

                if (ch == '#')
                {
                    index = end + 1;
                    return true;
                }
            }

            return false;
        }

        private static bool IsTwoCharOperator(string text, int index)
        {
            if (index + 1 >= text.Length)
                return false;

            foreach (var op in TwoCharOperators)
            {
                if (text[index] == op[0] && text[index + 1] == op[1])
                    return true;
            }

            return false;
        }

        private static void SkipToEndOfLine(string text, ref int index)
        {
            while (index < text.Length && text[index] != '\r' && text[index] != '\n')
            {
                index++;
            }
        }

        private static void SkipWhile(string text, ref int index, Func<char, bool> condition)
        {
            while (index < text.Length && condition(text[index]))
            {
                index++;
            }
        }

        private static bool IsSpace(char ch)
        {
            return ch != '\r' && ch != '\n' && char.IsWhiteSpace(ch);
        }

        private static bool IsIdentifierChar(char ch)
        {
            return char.IsLetterOrDigit(ch) || ch == '_';
        }

        private static bool IsRadixPrefix(char ch)
        {
            return ch == 'H' || ch == 'h' || ch == 'O' || ch == 'o';
        }
        #endregion
    }
}
