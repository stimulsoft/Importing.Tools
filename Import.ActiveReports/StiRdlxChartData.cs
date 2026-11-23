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

using Stimulsoft.Report.Dictionary;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The data of a chart of an RDLX report: the data transformation that counts the totals of the chart
    /// and the data source the transformation reads.
    /// </summary>
    public class StiRdlxChartData
    {
        #region Properties
        public StiDataTransformation Transformation { get; }

        /// <summary>
        /// The name of the data source the transformation reads the rows of.
        /// </summary>
        public string RowsName { get; }

        public StiRdlxDataSet DataSet { get; }
        #endregion

        public StiRdlxChartData(StiDataTransformation transformation, string rowsName, StiRdlxDataSet dataSet)
        {
            Transformation = transformation;
            RowsName = rowsName;
            DataSet = dataSet;
        }
    }
}
