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

namespace Stimulsoft.Report.Import.ReportingServices
{
    #region StiVbTokenType
    /// <summary>
    /// The kind of a token of an expression or of the code of Visual Basic.
    /// </summary>
    internal enum StiVbTokenType
    {
        /// <summary>
        /// Spaces and tabs.
        /// </summary>
        Whitespace,

        /// <summary>
        /// The end of a line.
        /// </summary>
        NewLine,

        /// <summary>
        /// A comment up to the end of the line, started with an apostrophe or REM.
        /// </summary>
        Comment,

        /// <summary>
        /// A name or a keyword, also a name in square brackets.
        /// </summary>
        Identifier,

        /// <summary>
        /// A number, also a hexadecimal or an octal one.
        /// </summary>
        Number,

        /// <summary>
        /// A string in quotes, a quote inside is doubled.
        /// </summary>
        String,

        /// <summary>
        /// A date between two number signs.
        /// </summary>
        Date,

        /// <summary>
        /// An operator or a punctuation mark.
        /// </summary>
        Operator
    }
    #endregion

    #region StiReportingServicesBodyLayout
    /// <summary>
    /// How the items of the body of a report of Reporting Services are put on the page.
    /// </summary>
    public enum StiReportingServicesBodyLayout
    {
        /// <summary>
        /// The bands of a data region that has nothing beside it are put on the page, the items between
        /// the data regions go to the report title, to the report summary or to a band that prints once.
        /// </summary>
        Bands,

        /// <summary>
        /// All the items of the body are on one band that prints once, every data region is a panel
        /// with its own bands, as in the body of Reporting Services.
        /// </summary>
        Panels
    }
    #endregion
}
