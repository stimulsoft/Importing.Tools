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
using CrystalDecisions.CrystalReports.Engine;
using Stimulsoft.Report.Dictionary;

namespace Import.CrystalReports
{
    /// <summary>
    /// The names of the members of the report class, see StiCrystalNameTable.
    /// </summary>
    public sealed partial class StiCrystalReportsHelper
    {
        #region Fields.Names
        private StiCrystalNameTable memberNames;
        private Dictionary<StiDataSource, string> dataSourceQueries;
        #endregion

        #region Methods.Names
        /// <summary>
        /// Starts the names of a report: the tables of the report and of its subreports are reserved first.
        /// </summary>
        private void InitMemberNames(ReportDocument crystalReport)
        {
            memberNames = new StiCrystalNameTable();
            dataSourceQueries = new Dictionary<StiDataSource, string>();
            ReserveTableNames(crystalReport);

            foreach (ReportDocument subreport in crystalReport.Subreports)
            {
                ReserveTableNames(subreport);
            }
        }

        private void ReserveTableNames(ReportDocument crystalReport)
        {
            foreach (Table table in crystalReport.Database.Tables)
            {
                memberNames.ReserveTable(table.Name);
            }
        }

        private static string GetDataSourceName(string tableName)
        {
            return StiCrystalNameTable.GetDataSourceName(tableName);
        }

        private string CreateMemberName(string name)
        {
            return memberNames.Create(name);
        }

        private string GetParameterVariableName(string parameterName)
        {
            return memberNames.GetParameterName(parameterName);
        }

        private static string CreateColumnName(StiDataSource dataSource, string fieldName)
        {
            return StiCrystalNameTable.CreateColumnName(dataSource, fieldName);
        }

        private static string GetColumnName(StiDataSource dataSource, string fieldName)
        {
            return StiCrystalNameTable.GetColumnName(dataSource, fieldName);
        }

        private string CreateRelationName(StiDataRelation relation, string parentTableName)
        {
            return StiCrystalNameTable.CreateRelationName(stimulReport.Dictionary, relation, parentTableName);
        }
        #endregion
    }
}
