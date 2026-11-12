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
using Stimulsoft.Base.Json.Linq;
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
        /// The rows read at the import, null when the data is not available at the import.
        /// </summary>
        public List<JObject> Rows { get; set; }

        /// <summary>
        /// The filters of the data set are already applied to the rows read at the import.
        /// </summary>
        public bool IsFiltered { get; set; }

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
