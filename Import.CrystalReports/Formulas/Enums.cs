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
    #region StiCrystalTokenType
    /// <summary>
    /// The kinds of the tokens of a Crystal formula.
    /// </summary>
    public enum StiCrystalTokenType
    {
        Identifier,
        Number,
        String,
        Date,
        Reference,
        Operator,
        Comma,
        Semicolon,
        Colon,
        LeftParenthesis,
        RightParenthesis,
        LeftBracket,
        RightBracket,

        /// <summary>
        /// The end of a line in Basic syntax, which ends a statement.
        /// </summary>
        NewLine,
        End
    }
    #endregion

    #region StiCrystalNodeKind
    /// <summary>
    /// The kinds of the nodes of the syntax tree of a Crystal formula.
    /// </summary>
    public enum StiCrystalNodeKind
    {
        Number,
        String,
        Date,
        Boolean,
        Reference,
        Identifier,
        Call,
        Unary,
        Binary,
        If,
        Select,
        Case,
        Array,
        Range,
        Index,
        Assign,
        Declare,
        Sequence,
        For,
        While,
        DoWhile,
        Exit
    }
    #endregion
}
