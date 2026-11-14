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

using Stimulsoft.Base.Json.Linq;
using Stimulsoft.Report.Dictionary;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Makes the databases and the data sources of Reporting Services that are not a plain SQL query:
    /// the data of a local report, the XML written into the query and a web API; the shared data sources
    /// are in StiReportingServicesHelper.Shared.cs. Nothing is read from the web while a report is imported,
    /// and no file but the definitions of the shared sources in the folder the caller names: the data written
    /// into the report goes into its resources, and a table that cannot be found without the data is named
    /// in the list of errors.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Fields.Sources
        /// <summary>
        /// The data sources of a local report, its data is given by the application at run time.
        /// </summary>
        private HashSet<string> localDataSources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// The XML databases that already hold the data written into the query of a data set.
        /// </summary>
        private HashSet<string> filledXmlDatabases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Sources
        /// <summary>
        /// Makes the database for the data provider of a data source. A local report, RDLC, takes
        /// its data from the application, so it has no database: its data sets get the data
        /// registered by the name of the data set, as ReportDataSource of ReportViewer does.
        /// </summary>
        private StiDatabase CreateDatabase(string name, string dataProvider, string connectString,
            string baseConnectionString, string embeddedData = null)
        {
            //a local report of Bold Reports may have no data provider, only the connection of a local report
            if (string.IsNullOrEmpty(dataProvider) && baseConnectionString?.Trim() == "/* Local Connection */")
            {
                localDataSources.Add(name);
                return null;
            }

            switch (dataProvider?.ToUpperInvariant())
            {
                case null:
                case "SQL":
                case "SQLAZURE":
                case "SQLPDW":
                    return new StiSqlDatabase(name, connectString);

                case "OLEDB":
                    return new StiOleDbDatabase(name, connectString);

                case "ODBC":
                    return new StiOdbcDatabase(name, connectString);

                case "ORACLE":
                    return new StiOracleDatabase(name, connectString);

                case "XML":
                    return new StiXmlDatabase(name, ConvertPathExpression(baseConnectionString));

                //the data entered in Report Builder or Power BI Report Builder, written into the query of every
                //data set as XmlData, the same as the data of an XML data source
                case "ENTERDATA":
                    return new StiXmlDatabase(name, string.Empty);

                case "WEBAPI":
                    return CreateWebApiDatabase(name, baseConnectionString);

                case "JSON":
                    return CreateBoldJsonDatabase(name, baseConnectionString, embeddedData);

                case "EXCEL":
                    return CreateBoldExcelDatabase(name, baseConnectionString, embeddedData);

                case "SYSTEM.DATA.DATASET":
                case "DATASET":
                    localDataSources.Add(name);
                    return null;

                default:
                    ThrowError(null, null, $"Data provider not supported, SQL is used: {dataProvider}");
                    return new StiSqlDatabase(name, connectString);
            }
        }

        /// <summary>
        /// The web API data source of Reporting Services is a JSON object with the URL, the method,
        /// the headers and the format of the data. A GET request of JSON data becomes a JSON database,
        /// the URL may be an expression. The value of a header given by an expression is not known
        /// before the report runs, so such a header is left out and named in the list of errors.
        /// </summary>
        private StiDatabase CreateWebApiDatabase(string name, string connectString)
        {
            JObject properties = null;
            try
            {
                properties = JObject.Parse(connectString ?? string.Empty);
            }
            catch (Exception)
            {
            }

            var url = properties?.Value<string>("URL");
            if (string.IsNullOrWhiteSpace(url))
            {
                ThrowError(null, null, $"Web API data source not imported: {name}");
                return new StiJsonDatabase(name, string.Empty);
            }

            var method = properties.Value<string>("MethodType");
            var format = properties.Value<string>("DataFormat");

            if (method != null && !string.Equals(method, "Get", StringComparison.OrdinalIgnoreCase) ||
                format != null && !string.Equals(format, "JSON", StringComparison.OrdinalIgnoreCase))
                ThrowError(null, null, $"Web API data source supports only GET of JSON: {name}");

            var database = new StiJsonDatabase(name, ConvertPathExpression(url));

            if (properties["Headers"] is JArray headers)
            {
                foreach (var header in headers.OfType<JObject>())
                {
                    var key = header.Value<string>("Key");
                    var value = header.Value<string>("Value");

                    if (string.IsNullOrEmpty(key))
                        continue;

                    if (IsExpression(value))
                        ThrowError(null, null, $"Web API header given by an expression not imported: {name}.{key}");

                    else
                        database.Headers.Add(key, value);
                }
            }

            return database;
        }

        /// <summary>
        /// A path or a URL given by an expression becomes a text with the expression in braces,
        /// a file database evaluates it before it loads the data.
        /// </summary>
        private string ConvertPathExpression(string path)
        {
            return IsExpression(path) ? ConvertExpression(path, null) : path;
        }

        /// <summary>
        /// Makes the data source of a data set of a local report. The data is registered by the
        /// application with the name of the data set, report.RegData("DataSet1", table).
        /// </summary>
        private StiDataSource CreateLocalDataSource(string dataSetName)
        {
            return new StiDataTableSource { NameInSource = dataSetName };
        }

        /// <summary>
        /// Makes the data source of a JSON database. A web API returns an array of rows as a rule,
        /// the JSON database names the table of such an array root.
        /// </summary>
        private StiDataSource CreateJsonDataSource(StiDatabase database, string dataSetName)
        {
            ThrowError(null, null, $"Web API data set is taken from the root array, check it: {dataSetName}");

            return new StiDataTableSource { NameInSource = $"{database.Name}.root" };
        }

        /// <summary>
        /// Makes the data source of an XML data set. The data written into the query, the XmlData
        /// element, goes into a resource of the report and the database reads the resource, so the
        /// imported report keeps its data. The table is the element of ElementPath, else the table
        /// that has all the fields of the data set, else the first table. Without the data the table
        /// is known only from ElementPath: the file or the URL of the data is not read on the import.
        /// </summary>
        private StiDataSource CreateXmlDataSource(StiXmlDatabase database, string query, ArrayList columns,
            string dataSetName)
        {
            var queryNode = ParseXmlQuery(query);
            var data = queryNode?.SelectSingleNode("XmlData")?.InnerXml.Trim();
            var tableName = GetElementPathTable(queryNode?.SelectSingleNode("ElementPath")?.InnerText);

            if (!string.IsNullOrEmpty(data))
            {
                var resourceName = StiNameCreation.CreateResourceName(report, baseName: dataSetName);
                report.Dictionary.Resources.Add(new StiResource(resourceName, StiResourceType.Xml,
                    Encoding.UTF8.GetBytes(data)));

                database = GetInlineXmlDatabase(database, dataSetName, resourceName);
                tableName = FindXmlTable(data, tableName, columns);
            }

            if (tableName == null)
                ThrowError(null, null, $"XML table not found, set it in the data source: {dataSetName}");

            var nameInSource = tableName != null ? $"{database.Name}.{tableName}" : database.Name;

            return new StiDataTableSource { NameInSource = nameInSource };
        }

        /// <summary>
        /// The first data set with its own data takes the database of its data source, the next ones
        /// get a database of their own, as one data source may serve many data sets.
        /// </summary>
        private StiXmlDatabase GetInlineXmlDatabase(StiXmlDatabase database, string dataSetName, string resourceName)
        {
            var path = $"resource://{resourceName}";

            if (filledXmlDatabases.Add(database.Name))
            {
                database.PathData = path;
                database.XmlType = StiXmlType.AdoNetXml;
                return database;
            }

            //the data written into the query has no schema, every database of it reads the data as ADO.NET XML
            var newDatabase = new StiXmlDatabase($"{database.Name}_{dataSetName}", path)
            {
                XmlType = StiXmlType.AdoNetXml
            };
            report.Dictionary.Databases.Add(newDatabase);

            return newDatabase;
        }

        /// <summary>
        /// Reads the query of an XML data set, the Query element with ElementPath and XmlData. A query
        /// that is not XML, or is empty, gives null.
        /// </summary>
        private static XmlNode ParseXmlQuery(string query)
        {
            if (string.IsNullOrWhiteSpace(query) || IsExpression(query))
                return null;

            try
            {
                var document = new XmlDocument { XmlResolver = null };
                using (var reader = XmlReader.Create(new StringReader(query), GetSafeXmlSettings()))
                {
                    document.Load(reader);
                }

                return document.DocumentElement?.LocalName == "Query" ? document.DocumentElement : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Returns the name of the last element of ElementPath, Customers {}/Customer{@ID} gives
        /// Customer: the rows of the data set are these elements.
        /// </summary>
        private static string GetElementPathTable(string elementPath)
        {
            if (string.IsNullOrWhiteSpace(elementPath))
                return null;

            var last = elementPath.Split('/').Last();
            var brace = last.IndexOf('{');
            if (brace != -1)
                last = last.Substring(0, brace);

            var colon = last.IndexOf(':');
            if (colon != -1)
                last = last.Substring(colon + 1);

            last = last.Trim();

            return last.Length > 0 ? last : null;
        }

        /// <summary>
        /// Finds the table of the data set in the data the way the XML database reads it, with
        /// DataSet.ReadXml. The data is parsed in memory without a DTD and without external entities.
        /// </summary>
        private static string FindXmlTable(string data, string elementPathTable, ArrayList columns)
        {
            DataSet dataSet;
            try
            {
                dataSet = new DataSet();
                using (var reader = XmlReader.Create(new StringReader(data), GetSafeXmlSettings()))
                {
                    dataSet.ReadXml(reader);
                }
            }
            catch (Exception)
            {
                return elementPathTable;
            }

            var tables = dataSet.Tables.Cast<DataTable>().ToList();

            if (elementPathTable != null && tables.Any(t => t.TableName == elementPathTable))
                return elementPathTable;

            var fields = columns.OfType<StiDataColumn>()
                .Where(c => !(c is StiCalcDataColumn))
                .Select(c => string.IsNullOrEmpty(c.NameInSource) ? c.Name : c.NameInSource.TrimStart('@'))
                .ToList();

            var table = tables.FirstOrDefault(t => fields.All(f => t.Columns.Contains(f))) ?? tables.FirstOrDefault();

            return table?.TableName;
        }

        private static XmlReaderSettings GetSafeXmlSettings()
        {
            return new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null
            };
        }
        #endregion
    }
}
