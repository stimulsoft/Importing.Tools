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
    /// The header cell of a row member of a Tablix on the way from the root to a row.
    /// </summary>
    public class StiRdlxTablixHeader
    {
        #region Properties
        public StiRdlxNode Item { get; set; }

        /// <summary>
        /// The index of the first level the cell takes.
        /// </summary>
        public int Level { get; set; }

        public int LevelCount { get; set; }

        /// <summary>
        /// The cell is already placed on a row, so the rows below leave its place empty.
        /// </summary>
        public bool IsPlaced { get; set; }
        #endregion
    }
}
