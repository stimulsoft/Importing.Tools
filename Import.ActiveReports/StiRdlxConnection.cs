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
    /// A database of the Stimulsoft report made of the data of an RDLX data source: a JSON, XML or CSV
    /// database that reads a resource of the report, a file or an address.
    /// </summary>
    public class StiRdlxConnection
    {
        #region Properties
        /// <summary>
        /// The name of the database in the dictionary of the report.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// The database reads JSON: its tables are named by the path to the rows.
        /// </summary>
        public bool IsJson { get; }

        /// <summary>
        /// The names of the tables of the database, null when its data is not available at the import.
        /// </summary>
        public List<string> Tables { get; set; }
        #endregion

        public StiRdlxConnection(string name, bool isJson)
        {
            Name = name;
            IsJson = isJson;
        }
    }
}
