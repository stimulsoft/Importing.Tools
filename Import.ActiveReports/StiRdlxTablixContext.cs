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
using Stimulsoft.Report.Components;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The state of the conversion of one Tablix to bands.
    /// </summary>
    public class StiRdlxTablixContext
    {
        #region Properties
        /// <summary>
        /// The left edges of the levels of the row headers and the right edge of the last level.
        /// </summary>
        public double[] LevelEdges { get; set; }

        /// <summary>
        /// The left edges of the columns of the body and the right edge of the last column.
        /// </summary>
        public double[] ColumnEdges { get; set; }

        public List<StiRdlxNode> BodyRows { get; set; }

        public StiContainer Container { get; set; }

        public StiRdlxScope Scope { get; set; }

        /// <summary>
        /// The data band of the Tablix, null when the Tablix has no data.
        /// </summary>
        public StiDataBand DataBand { get; set; }

        /// <summary>
        /// The band of the rows that stand before the rows of the groups.
        /// </summary>
        public StiBand HeaderBand { get; set; }

        /// <summary>
        /// The band of the rows that stand after the rows of the groups.
        /// </summary>
        public StiBand FooterBand { get; set; }

        /// <summary>
        /// The group footer bands in the order of their group headers.
        /// </summary>
        public List<StiBand> GroupFooters { get; } = new List<StiBand>();

        /// <summary>
        /// The bottom of the rows placed on each band.
        /// </summary>
        public Dictionary<StiBand, double> Bottoms { get; } = new Dictionary<StiBand, double>();

        /// <summary>
        /// The index of the next row of the body.
        /// </summary>
        public int RowIndex { get; set; }
        #endregion
    }
}
