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

using Stimulsoft.Report.Components;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The place of a report item in the data of the report: the data set its expressions read and the data band
    /// it is printed with.
    /// </summary>
    public class StiRdlxScope
    {
        #region Properties
        public StiRdlxDataSet DataSet { get; set; }

        /// <summary>
        /// The data band of the nearest data region, null out of data regions.
        /// </summary>
        public StiDataBand Band { get; set; }

        /// <summary>
        /// The item is printed on a band of a data region, so a total without a band is counted by this region.
        /// </summary>
        public bool IsRegion { get; set; }
        #endregion

        #region Methods
        public StiRdlxScope Clone()
        {
            return new StiRdlxScope
            {
                DataSet = DataSet,
                Band = Band,
                IsRegion = IsRegion
            };
        }
        #endregion
    }
}
