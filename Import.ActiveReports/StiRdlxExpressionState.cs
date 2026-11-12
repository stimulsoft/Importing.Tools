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

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The state of the conversion of one RDLX expression.
    /// </summary>
    public class StiRdlxExpressionState
    {
        #region Properties
        public List<StiRdlxToken> Tokens { get; set; }

        public int Index { get; set; }

        public StiRdlxScope Scope { get; set; }

        /// <summary>
        /// The expression is written in the syntax of ActiveReportsJS: a name without a prefix is a field,
        /// a name with @ is a parameter.
        /// </summary>
        public bool IsTemplate { get; set; }

        /// <summary>
        /// A total out of a data region may get the marker of its data band.
        /// </summary>
        public bool AllowBandMarker { get; set; }

        public bool IsEnd => Index >= Tokens.Count;

        public StiRdlxToken Current => Index < Tokens.Count ? Tokens[Index] : null;
        #endregion

        #region Methods
        public bool IsOperator(string text)
        {
            return IsOperator(0, text);
        }

        public bool IsOperator(int offset, string text)
        {
            var index = Index + offset;

            return index < Tokens.Count && Tokens[index].Type == StiRdlxTokenType.Operator &&
                Tokens[index].Text == text;
        }

        public bool IsType(int offset, StiRdlxTokenType type)
        {
            var index = Index + offset;
            return index < Tokens.Count && Tokens[index].Type == type;
        }
        #endregion
    }
}
