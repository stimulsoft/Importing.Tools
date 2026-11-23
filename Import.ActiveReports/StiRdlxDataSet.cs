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
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// A data set of an RDLX report and the data source of the Stimulsoft report made of it.
    /// </summary>
    public class StiRdlxDataSet
    {
        #region Properties
        public string Name { get; }

        public StiRdlxNode Node { get; }

        /// <summary>
        /// The name of the data source in the dictionary of the report.
        /// </summary>
        public string SourceName { get; set; }

        public StiDataSource Source { get; set; }

        /// <summary>
        /// The names of the columns of the data source by the names of the fields of the data set.
        /// </summary>
        public Dictionary<string, string> Columns { get; } =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// The data set this one is nested in: its rows are the items of an array in a row of the parent.
        /// </summary>
        public StiRdlxDataSet Parent { get; set; }

        /// <summary>
        /// The field of the parent that holds the rows of this data set.
        /// </summary>
        public string ParentField { get; set; }

        /// <summary>
        /// The column of the parent that holds the key of the rows of this data set, null when the rows
        /// are not linked with the rows of the parent.
        /// </summary>
        public string ParentColumn { get; set; }

        /// <summary>
        /// The column of this data set that holds the key of the row of the parent.
        /// </summary>
        public string KeyColumn { get; set; }

        /// <summary>
        /// The database of the data set, null when its data has no database of the report.
        /// </summary>
        public StiRdlxConnection Connection { get; set; }

        /// <summary>
        /// The names of the path from the root of the data to the rows of the data set.
        /// </summary>
        public List<string> PathNames { get; set; } = new List<string>();

        /// <summary>
        /// The table of the database that holds the rows of the data set.
        /// </summary>
        public string Table { get; set; }

        /// <summary>
        /// The name of the data source that reads the table. It differs from the name of the data source of
        /// the data set when the data set has filters: a virtual source with the name of the data set
        /// filters the rows of the table.
        /// </summary>
        public string BaseName { get; set; }

        /// <summary>
        /// The data set is the root object of the data, it has one row and no table.
        /// </summary>
        public bool IsSingleRow { get; set; }

        /// <summary>
        /// The data source of the data set is already made.
        /// </summary>
        public bool IsBuilt { get; set; }

        /// <summary>
        /// The filters of the data set are applied by its data source.
        /// </summary>
        public bool IsFiltered { get; set; }

        /// <summary>
        /// The data set this one is a copy of. A copy reads the rows of the current group of the original:
        /// it is made for a data region that stands on a band of a group of the same data set.
        /// </summary>
        public StiRdlxDataSet Original { get; set; }

        /// <summary>
        /// The filters every data band of a copy gets: they leave the rows of the current group.
        /// </summary>
        public List<string> RowFilters { get; } = new List<string>();

        /// <summary>
        /// The first data band made for this data set, the totals placed out of data regions refer to it.
        /// </summary>
        public StiDataBand Band { get; set; }
        #endregion

        public StiRdlxDataSet(string name, StiRdlxNode node)
        {
            Name = name;
            Node = node;
        }
    }
}
