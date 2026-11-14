#region Copyright (C) 2003-2026 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports												}
{	                         										}
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

using Stimulsoft.Base;
using Stimulsoft.Base.Json.Linq;
using Stimulsoft.Report.Dictionary;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Makes the Excel data sources of Bold Reports. The workbook is written into the report as Base64 in the
    /// embedded data, and a data set names its worksheet in the query, {"Name":"Sheet1"}. The workbook goes into
    /// a resource of the report read by an Excel database of Stimulsoft, the first row of a sheet is the header.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Fields.Excel
        /// <summary>
        /// The workbooks of the Excel data sources of Bold Reports by the name of the database.
        /// </summary>
        private Dictionary<string, byte[]> boldExcelData =
            new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Methods.Excel
        /// <summary>
        /// Makes the database of an Excel data source of Bold Reports: the workbook written into the report, else
        /// the file of the URL read when the report runs.
        /// </summary>
        private StiDatabase CreateBoldExcelDatabase(string name, string connectString, string embeddedData)
        {
            var data = DecodeBoldEmbeddedBytes(embeddedData);
            if (data != null)
            {
                var resourceName = StiNameCreation.CreateResourceName(report, baseName: name);
                report.Dictionary.Resources.Add(new StiResource(resourceName, StiResourceType.Excel, data));

                boldExcelData[name] = data;
                return new StiExcelDatabase(name, $"resource://{resourceName}");
            }

            string url = null;
            try
            {
                url = JObject.Parse(connectString ?? string.Empty).Value<string>("URL");
            }
            catch (Exception)
            {
            }

            if (string.IsNullOrWhiteSpace(url))
                ThrowError(null, null, $"Excel data source without data not imported: {name}");

            return new StiExcelDatabase(name, ConvertPathExpression(url ?? string.Empty));
        }

        private static byte[] DecodeBoldEmbeddedBytes(string data)
        {
            if (string.IsNullOrWhiteSpace(data))
                return null;

            try
            {
                return Convert.FromBase64String(data.Trim());
            }
            catch (FormatException)
            {
                return null;
            }
        }

        /// <summary>
        /// Makes the data source of a data set of an Excel data source of Bold Reports. The table is the worksheet
        /// named in the query as the Excel database reads it, else the sheet with all the fields of the data set.
        /// </summary>
        private StiDataSource CreateBoldExcelDataSource(StiExcelDatabase database, byte[] data, string query,
            ArrayList columns)
        {
            string sheet = null;
            try
            {
                sheet = JObject.Parse(query ?? string.Empty).Value<string>("Name");
            }
            catch (Exception)
            {
            }

            List<DataTable> tables;
            try
            {
                var options = new StiExcelOptions(data, database.FirstRowIsHeader);
                tables = StiExcelConnector.Get().GetDataSet(options)?.Tables.Cast<DataTable>().ToList() ??
                    new List<DataTable>();
            }
            catch (Exception)
            {
                tables = new List<DataTable>();
            }

            var fields = columns.OfType<StiDataColumn>()
                .Where(c => !(c is StiCalcDataColumn))
                .Select(GetBoldJsonFieldName)
                .ToList();
            var table = tables.FirstOrDefault(t => string.Equals(t.TableName, sheet,
                    StringComparison.OrdinalIgnoreCase)) ??
                tables.FirstOrDefault(t => fields.All(f => t.Columns.Contains(f))) ??
                tables.FirstOrDefault();

            var tableName = table?.TableName ?? sheet;
            if (tableName == null)
                ThrowError(null, null, $"Excel worksheet not found, set it in the data source: {database.Name}");

            return new StiDataTableSource { NameInSource = $"{database.Name}.{tableName}" };
        }
        #endregion
    }
}
