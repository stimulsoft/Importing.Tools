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

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The format of an ActiveReports file.
    /// </summary>
    public enum StiActiveReportsFormat
    {
        /// <summary>
        /// A section report of ActiveReports, an XML file with the root ActiveReportsLayout.
        /// </summary>
        Rpx,

        /// <summary>
        /// A page or RDL report of ActiveReports.NET, an XML file with the root Report.
        /// </summary>
        RdlxXml,

        /// <summary>
        /// A report of ActiveReportsJS, the same model as RDLX written as JSON.
        /// </summary>
        RdlxJson
    }

    /// <summary>
    /// The kind of a token of an RDLX expression.
    /// </summary>
    public enum StiRdlxTokenType
    {
        Identifier,
        Number,
        String,
        Date,
        Operator,
        OpenParen,
        CloseParen,
        Comma
    }
}
