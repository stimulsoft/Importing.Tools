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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Makes the JSON data sources of Bold Reports. The data is written into the report, in the Data of the
    /// connection string or in the embedded data as Base64, and the data set takes a flat table of it: the rows
    /// are the root array, the array in the root object or the root object itself; a nested object gives its
    /// values to the row, a nested array of objects gives a row per item. A value of a nested object is named
    /// "name (Parent_Child)" when its name is used more than once, else just by its name.
    /// Flat data, an array of objects with plain values, is the same table in the JSON database of Stimulsoft,
    /// so it goes into a JSON resource read by a JSON database and the report keeps its JSON. The JSON database
    /// makes a table of every nested object and array, linked by relations, so nested data is flattened on the
    /// import and goes into a resource as ADO.NET XML with its schema, the same as the data of an XML query.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts.Json
        private const string BoldJsonTable = "Rows";
        #endregion

        #region Fields.Json
        /// <summary>
        /// The data of the JSON data sources of Bold Reports by the name of the database.
        /// </summary>
        private Dictionary<string, JToken> boldJsonData =
            new Dictionary<string, JToken>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// The flat data of the JSON data sources of Bold Reports that are read by a JSON database.
        /// </summary>
        private Dictionary<string, JToken> boldFlatJsonData =
            new Dictionary<string, JToken>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Methods.Json
        /// <summary>
        /// Makes the database of a JSON data source of Bold Reports. The data written into the report goes to an
        /// XML database, see CreateBoldJsonDataSource; a URL is read by a JSON database when the report runs.
        /// </summary>
        private StiDatabase CreateBoldJsonDatabase(string name, string connectString, string embeddedData)
        {
            JObject properties = null;
            try
            {
                properties = JObject.Parse(connectString ?? string.Empty);
            }
            catch (Exception)
            {
            }

            var text = properties?.Value<string>("Data");
            if (string.IsNullOrWhiteSpace(text) && !string.IsNullOrWhiteSpace(embeddedData))
                text = DecodeBoldEmbeddedData(embeddedData);

            if (!string.IsNullOrWhiteSpace(text))
            {
                JToken data;
                try
                {
                    data = JToken.Parse(text);
                }
                catch (Exception)
                {
                    ThrowError(null, null, $"JSON data of a data source not read: {name}");
                    return new StiXmlDatabase(name, string.Empty);
                }

                if (!IsFlatBoldJson(data))
                {
                    boldJsonData[name] = data;
                    return new StiXmlDatabase(name, string.Empty);
                }

                //the JSON written again is strict, Bold Reports allows the quotes of JavaScript
                var resourceName = StiNameCreation.CreateResourceName(report, baseName: name);
                report.Dictionary.Resources.Add(new StiResource(resourceName, StiResourceType.Json,
                    Encoding.UTF8.GetBytes(data.ToString())));

                boldFlatJsonData[name] = data;
                return new StiJsonDatabase(name, $"resource://{resourceName}");
            }

            var url = properties?.Value<string>("URL");
            if (!string.IsNullOrWhiteSpace(url))
                return new StiJsonDatabase(name, ConvertPathExpression(url));

            ThrowError(null, null, $"JSON data source without data not imported: {name}");
            return new StiXmlDatabase(name, string.Empty);
        }

        /// <summary>
        /// Returns true when the rows of the data are objects with plain values only.
        /// </summary>
        private static bool IsFlatBoldJson(JToken data)
        {
            if (!(data is JArray) && !(data is JObject root && root.Count == 1 && root.First.First is JArray))
                return false;

            var items = GetBoldJsonItems(data).ToList();

            return items.Count > 0 && items.All(i => i is JObject item && item.Properties().All(p => p.Value is JValue));
        }

        /// <summary>
        /// Makes the data source of a data set of a flat JSON data source of Bold Reports. The table is the one
        /// the JSON database makes of the data with all the fields of the data set, else the first one.
        /// </summary>
        private StiDataSource CreateBoldFlatJsonDataSource(StiJsonDatabase database, JToken data, ArrayList columns)
        {
            var options = new StiJsonOptions(Encoding.UTF8.GetBytes(data.ToString()), database.RelationDirection);
            var tables = StiJsonConnector.Get().GetDataSet(options)?.Tables.Cast<DataTable>().ToList() ??
                new List<DataTable>();

            var fields = columns.OfType<StiDataColumn>()
                .Where(c => !(c is StiCalcDataColumn))
                .Select(GetBoldJsonFieldName)
                .ToList();
            var table = tables.FirstOrDefault(t => fields.All(f => t.Columns.Contains(f))) ?? tables.FirstOrDefault();

            return new StiDataTableSource { NameInSource = $"{database.Name}.{table?.TableName}" };
        }

        private static string DecodeBoldEmbeddedData(string data)
        {
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(data.Trim())).TrimStart((char)0xFEFF);
            }
            catch (FormatException)
            {
                return null;
            }
        }

        /// <summary>
        /// Makes the data source of a data set of a JSON data source of Bold Reports: the flat table of the data
        /// with a column for every field of the data set, typed as the field.
        /// </summary>
        private StiDataSource CreateBoldJsonDataSource(StiXmlDatabase database, JToken data, ArrayList columns,
            string dataSetName)
        {
            var table = new DataTable(BoldJsonTable);
            var fields = columns.OfType<StiDataColumn>().Where(c => !(c is StiCalcDataColumn)).ToList();
            foreach (var field in fields)
            {
                var type = Nullable.GetUnderlyingType(field.Type) ?? field.Type;
                table.Columns.Add(GetBoldJsonFieldName(field), type == typeof(object) ? typeof(string) : type);
            }

            foreach (var row in GetBoldJsonRows(data))
            {
                var values = table.NewRow();
                foreach (DataColumn column in table.Columns)
                {
                    values[column] = GetBoldJsonValue(row, column);
                }

                table.Rows.Add(values);
            }

            var dataSet = new DataSet("Data");
            dataSet.Tables.Add(table);

            string xml;
            using (var writer = new StringWriter())
            {
                dataSet.WriteXml(writer, XmlWriteMode.WriteSchema);
                xml = writer.ToString();
            }

            var resourceName = StiNameCreation.CreateResourceName(report, baseName: dataSetName);
            report.Dictionary.Resources.Add(new StiResource(resourceName, StiResourceType.Xml,
                Encoding.UTF8.GetBytes(xml)));

            database = GetInlineXmlDatabase(database, dataSetName, resourceName);

            return new StiDataTableSource { NameInSource = $"{database.Name}.{BoldJsonTable}" };
        }

        private static string GetBoldJsonFieldName(StiDataColumn field)
        {
            return string.IsNullOrEmpty(field.NameInSource) ? field.Name : field.NameInSource;
        }

        /// <summary>
        /// Returns the rows of the data: the items of the root array, of the only array of the root object, or
        /// the root object; every row is the list of its values with their names and the path of their object.
        /// </summary>
        private static List<List<Tuple<string, string, JToken>>> GetBoldJsonRows(JToken data)
        {
            var rows = new List<List<Tuple<string, string, JToken>>>();
            foreach (var item in GetBoldJsonItems(data).OfType<JObject>())
            {
                rows.AddRange(FlattenBoldJsonObject(item, null));
            }

            return rows;
        }

        private static IEnumerable<JToken> GetBoldJsonItems(JToken data)
        {
            if (data is JArray array)
                return array.Children();

            var root = data as JObject;
            if (root?.Count == 1 && root.First is JProperty single && single.Value is JArray items)
                return items.Children();

            return new[] { data };
        }

        /// <summary>
        /// Flattens an object into rows: a value is added to every row, an object adds its own values, an array
        /// of objects gives a row for every item.
        /// </summary>
        private static List<List<Tuple<string, string, JToken>>> FlattenBoldJsonObject(JObject item, string path)
        {
            var rows = new List<List<Tuple<string, string, JToken>>> { new List<Tuple<string, string, JToken>>() };
            foreach (var property in item.Properties())
            {
                var childPath = path == null ? property.Name : path + "_" + property.Name;
                if (property.Value is JObject child)
                {
                    rows = CrossBoldJsonRows(rows, FlattenBoldJsonObject(child, childPath));
                }
                else if (property.Value is JArray array && array.Children().OfType<JObject>().Any())
                {
                    var itemRows = array.Children().OfType<JObject>()
                        .SelectMany(o => FlattenBoldJsonObject(o, childPath))
                        .ToList();
                    rows = CrossBoldJsonRows(rows, itemRows);
                }
                else
                {
                    var value = property.Value is JArray values
                        ? new JValue(string.Join(", ", values.Children().Select(v => v.ToString())))
                        : property.Value;

                    foreach (var row in rows)
                    {
                        row.Add(Tuple.Create(property.Name, path, value));
                    }
                }
            }

            return rows;
        }

        private static List<List<Tuple<string, string, JToken>>> CrossBoldJsonRows(
            List<List<Tuple<string, string, JToken>>> rows, List<List<Tuple<string, string, JToken>>> childRows)
        {
            if (childRows.Count == 0)
                return rows;

            return rows.SelectMany(r => childRows.Select(c => r.Concat(c).ToList())).ToList();
        }

        /// <summary>
        /// Returns the value of the column in the row converted to the type of the column, or DBNull. The column
        /// is the name of the value, or the name with the path of its object in brackets.
        /// </summary>
        private static object GetBoldJsonValue(List<Tuple<string, string, JToken>> row, DataColumn column)
        {
            var value = row.FirstOrDefault(v => v.Item1 == column.ColumnName && v.Item2 == null) ??
                row.FirstOrDefault(v => v.Item2 != null && $"{v.Item1} ({v.Item2})" == column.ColumnName) ??
                row.FirstOrDefault(v => v.Item1 == column.ColumnName);

            if (value?.Item3 is JValue jsonValue && jsonValue.Value != null)
            {
                try
                {
                    return column.DataType == typeof(string)
                        ? Convert.ToString(jsonValue.Value, CultureInfo.InvariantCulture)
                        : StiConvert.ChangeType(jsonValue.Value, column.DataType);
                }
                catch (Exception)
                {
                }
            }

            return DBNull.Value;
        }
        #endregion
    }
}
