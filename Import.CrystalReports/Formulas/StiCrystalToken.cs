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

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// A token of a Crystal formula.
    /// </summary>
    public class StiCrystalToken
    {
        #region Properties
        public StiCrystalTokenType Type { get; }

        /// <summary>
        /// The text of the token: the word, the operator, the content of a string without quotes,
        /// the content of a reference or a date without the braces and the hashes.
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// The position of the token in the formula.
        /// </summary>
        public int Position { get; }
        #endregion

        #region Methods
        /// <summary>
        /// Checks whether the token is the specified word, ignoring the case as Crystal does.
        /// </summary>
        public bool IsWord(string word)
        {
            return Type == StiCrystalTokenType.Identifier &&
                string.Equals(Text, word, System.StringComparison.OrdinalIgnoreCase);
        }

        public bool IsOperator(string op)
        {
            return Type == StiCrystalTokenType.Operator && Text == op;
        }

        public override string ToString()
        {
            return $"{Type} {Text}";
        }
        #endregion

        public StiCrystalToken(StiCrystalTokenType type, string text, int position)
        {
            Type = type;
            Text = text;
            Position = position;
        }
    }
}
