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
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Stimulsoft.Base.Json;
using Stimulsoft.Base.Json.Linq;
using Stimulsoft.Report.Dictionary;
using JsonFormatting = Stimulsoft.Base.Json.Formatting;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The parameters, the data sources and the data sets of the report.
    /// The data that is available at the import - written in the report itself (jsondata, xmldata) or lying in
    /// a file beside the report (jsondoc, endpoint, CSV) - is read here: the query of every data set is applied
    /// and its rows are written into a JSON resource of the report, a table per data set. So the imported
    /// report keeps its data and does not depend on the query syntax of ActiveReports.
    /// A data set whose data is on a server gets a JSON database with the address of the service. The import
    /// itself never asks a server: the table of such a data set is named by the path of its query.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Consts.Data
        private const string EmbeddedDataName = "ReportData";
        private const string KeyColumn = "_id";
        private const string ParentKeyColumn = "_parentId";
        #endregion

        #region Fields.Data
        private Dictionary<string, StiRdlxDataSet> dataSets =
            new Dictionary<string, StiRdlxDataSet>(StringComparer.OrdinalIgnoreCase);
        private List<StiRdlxDataSet> dataSetList = new List<StiRdlxDataSet>();
        private Dictionary<string, StiRdlxNode> dataSourceNodes =
            new Dictionary<string, StiRdlxNode>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, JToken> loadedData = new Dictionary<string, JToken>(StringComparer.Ordinal);
        private Dictionary<string, List<int>> parentRows = new Dictionary<string, List<int>>();
        private Dictionary<string, string> remoteDatabases = new Dictionary<string, string>(StringComparer.Ordinal);
        private JObject embeddedData = new JObject();
        #endregion

        #region Methods.Parameters
        private void ProcessParameters()
        {
            foreach (var parameter in root.Items("ReportParameters"))
            {
                var name = parameter.Get("Name");
                if (string.IsNullOrEmpty(name))
                    continue;

                string value = null;
                var defaultValue = parameter.Child("DefaultValue");
                if (defaultValue != null)
                {
                    var values = defaultValue.Items("Values");
                    if (values.Count > 0)
                        value = values[0].Value;
                }

                AddVariable(name, GetParameterType(parameter.Get("DataType")), parameter.Get("Prompt", name), value,
                    parameter.GetBool("Hidden"));
            }
        }

        private static Type GetParameterType(string dataType)
        {
            switch ((dataType ?? "").ToLowerInvariant())
            {
                case "integer":
                    return typeof(int);

                case "float":
                    return typeof(decimal);

                case "boolean":
                    return typeof(bool);

                case "date":
                case "datetime":
                    return typeof(DateTime);

                default:
                    return typeof(string);
            }
        }

        /// <summary>
        /// Adds a variable for a parameter and returns its name.
        /// </summary>
        private string AddVariable(string name, Type type, string alias, string value, bool isHidden)
        {
            var variableName = CreateName(name);
            var variable = new StiVariable("", variableName, type);
            variable.Alias = alias;
            variable.RequestFromUser = !isHidden;

            if (!string.IsNullOrEmpty(value) && value.StartsWith("="))
            {
                variable.InitBy = StiVariableInitBy.Expression;
                variable.Value = ConvertExpression(value.Substring(1), null, isTemplate: false, allowBandMarker: false);
            }
            else if (!string.IsNullOrEmpty(value))
            {
                try
                {
                    variable.ValueObject = Convert.ChangeType(value, type, CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {
                    variable.Value = value;
                }
            }

            report.Dictionary.Variables.Add(variable);
            parameters[name] = variableName;

            return variableName;
        }
        #endregion

        #region Methods.Data
        private void ProcessData()
        {
            foreach (var source in root.Items("DataSources"))
            {
                var name = source.Get("Name");
                if (!string.IsNullOrEmpty(name))
                    dataSourceNodes[name] = source;
            }

            foreach (var node in root.Items("DataSets"))
            {
                var name = node.Get("Name");
                if (string.IsNullOrEmpty(name) || dataSets.ContainsKey(name))
                    continue;

                NormalizeFieldNames(node);

                var dataSet = new StiRdlxDataSet(name, node);
                dataSet.SourceName = CreateName(name);
                dataSets[name] = dataSet;
                dataSetList.Add(dataSet);
            }

            foreach (var dataSet in dataSetList)
            {
                ResolveParent(dataSet);
                ReadColumns(dataSet);
            }

            foreach (var dataSet in dataSetList)
            {
                LoadRows(dataSet);
            }

            foreach (var dataSet in dataSetList)
            {
                BuildDataSource(dataSet);
            }

            foreach (var dataSet in dataSetList)
            {
                AddCalculatedColumns(dataSet);
            }
        }

        /// <summary>
        /// Writes the data read at the import into a JSON resource of the report. It is called when all
        /// the report is made: the charts add their own tables to this data.
        /// </summary>
        private void AddEmbeddedData()
        {
            if (embeddedData.Count == 0)
                return;

            var bytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(embeddedData, JsonFormatting.None));

            report.Dictionary.Resources.Add(new StiResource(EmbeddedDataName, StiResourceType.Json, bytes));
            report.Dictionary.Databases.Add(new StiJsonDatabase(EmbeddedDataName, "resource://" + EmbeddedDataName));
        }

        /// <summary>
        /// A field whose name has spaces or signs is written in brackets and quotes, ['Net Income'], both
        /// as the name and as the data field. The brackets are taken off: the name is 'Net Income'.
        /// </summary>
        private static void NormalizeFieldNames(StiRdlxNode dataSet)
        {
            foreach (var field in dataSet.Items("Fields"))
            {
                foreach (var node in field.Nodes)
                {
                    var isName = string.Equals(node.Name, "Name", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(node.Name, "DataField", StringComparison.OrdinalIgnoreCase);

                    if (isName && node.Value != null && node.Value.StartsWith("['") && node.Value.EndsWith("']") &&
                        node.Value.Length > 4)
                        node.Value = node.Value.Substring(2, node.Value.Length - 4);
                }
            }
        }

        private StiRdlxDataSet FindDataSet(string name)
        {
            StiRdlxDataSet dataSet;
            if (!string.IsNullOrEmpty(name) && dataSets.TryGetValue(name, out dataSet))
                return dataSet;

            return null;
        }

        /// <summary>
        /// The data set of a data region: the named one, else the only one of the report, else the one of
        /// the outer region.
        /// </summary>
        private StiRdlxDataSet GetRegionDataSet(StiRdlxNode region, StiRdlxScope scope)
        {
            var name = region.Get("DataSetName");
            if (!string.IsNullOrEmpty(name))
            {
                var dataSet = FindDataSet(name);
                if (dataSet == null)
                    Warn($"The data set '{name}' of '{region.Get("Name")}' is not found.");

                return dataSet;
            }

            if (dataSetList.Count == 1)
                return dataSetList[0];

            return scope.DataSet;
        }

        /// <summary>
        /// A data set is nested when its data source is "$dataset:Parent/Field": its rows are the items of
        /// an array in a row of the parent data set.
        /// </summary>
        private void ResolveParent(StiRdlxDataSet dataSet)
        {
            var sourceName = GetQuery(dataSet).Get("DataSourceName", "");
            if (!sourceName.StartsWith("$dataset:", StringComparison.OrdinalIgnoreCase))
                return;

            var path = sourceName.Substring("$dataset:".Length);
            var slash = path.IndexOf('/');
            var parent = FindDataSet(slash < 0 ? path : path.Substring(0, slash));

            if (parent == null || parent == dataSet)
            {
                Warn($"The parent data set of '{dataSet.Name}' is not found.");
                return;
            }

            dataSet.Parent = parent;
            dataSet.ParentField = slash < 0 ? "" : path.Substring(slash + 1);
        }

        private static StiRdlxNode GetQuery(StiRdlxDataSet dataSet)
        {
            return dataSet.Node.Child("Query") ?? new StiRdlxNode("Query");
        }

        private void ReadColumns(StiRdlxDataSet dataSet)
        {
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var field in dataSet.Node.Items("Fields"))
            {
                var name = field.Get("Name");
                if (string.IsNullOrEmpty(name) || dataSet.Columns.ContainsKey(name))
                    continue;

                dataSet.Columns[name] = CreateColumnName(name, used);
            }
        }

        private static string CreateColumnName(string field)
        {
            var builder = new StringBuilder();

            foreach (var ch in field)
            {
                builder.Append(char.IsLetterOrDigit(ch) || ch == '_' ? ch : '_');
            }

            if (builder.Length == 0 || char.IsDigit(builder[0]))
                builder.Insert(0, '_');

            return builder.ToString();
        }

        private static string CreateColumnName(string field, HashSet<string> used)
        {
            var baseName = CreateColumnName(field);
            var name = baseName;
            var index = 2;

            while (!used.Add(name))
            {
                name = baseName + index++;
            }

            return name;
        }
        #endregion

        #region Methods.Rows
        /// <summary>
        /// Reads the rows of a data set when its data is available at the import.
        /// </summary>
        private void LoadRows(StiRdlxDataSet dataSet)
        {
            if (dataSet.Rows != null)
                return;

            try
            {
                if (dataSet.Parent != null)
                {
                    LoadRows(dataSet.Parent);
                    LoadNestedRows(dataSet);
                    return;
                }

                var query = GetQuery(dataSet);
                StiRdlxNode source;
                if (!dataSourceNodes.TryGetValue(query.Get("DataSourceName", ""), out source))
                    return;

                var connection = source.Child("ConnectionProperties") ?? new StiRdlxNode("ConnectionProperties");
                var command = ParsePairs(query.Get("CommandText", ""), "jpath");
                var data = LoadData(connection, command);
                if (data == null)
                    return;

                dataSet.Rows = SelectRows(data, GetPair(command, "jpath"));

                // The filters that compare fields with constants are applied here, to the rows written into
                // the report, so the bands and the charts get the data already filtered.
                var filters = dataSet.Node.Items("Filters");
                if (filters.Count > 0 && AreStaticFilters(filters, dataSet))
                {
                    dataSet.Rows = FilterRows(dataSet.Rows, filters, dataSet, isTable: false);
                    dataSet.IsFiltered = true;
                }
            }
            catch (Exception e)
            {
                Warn($"The data of the data set '{dataSet.Name}' is not read: {e.Message}");
            }
        }

        private void LoadNestedRows(StiRdlxDataSet dataSet)
        {
            if (dataSet.Parent.Rows == null)
                return;

            var rows = new List<JObject>();
            var parents = new List<int>();

            for (var index = 0; index < dataSet.Parent.Rows.Count; index++)
            {
                var token = dataSet.ParentField.Length == 0
                    ? dataSet.Parent.Rows[index]
                    : SelectField(dataSet.Parent.Rows[index], dataSet.ParentField);

                foreach (var row in ToRows(token))
                {
                    rows.Add(row);
                    parents.Add(index);
                }
            }

            dataSet.Rows = rows;
            parentRows[dataSet.Name] = parents;
        }

        /// <summary>
        /// Loads the data of a data source: JSON, XML or CSV written in the connection string or lying in
        /// a file beside the report. Returns null when the data is on a server or is not found.
        /// </summary>
        private JToken LoadData(StiRdlxNode connection, Dictionary<string, string> command)
        {
            var connectString = connection.Get("ConnectString", "");
            if (connectString.StartsWith("="))
                return null;

            var separator = connectString.IndexOf('=');
            if (separator < 0)
                return null;

            var kind = connectString.Substring(0, separator).Trim().ToLowerInvariant();
            var value = connectString.Substring(separator + 1);

            switch (kind)
            {
                case "jsondata":
                    return LoadCached("json:" + value, value, isFile: false);

                case "xmldata":
                    return LoadXmlData(value, GetPair(command, "jpath"));

                case "jsondoc":
                case "endpoint":
                case "csvdoc":
                    var path = value.Split(';')[0].Trim() + GetPair(command, "uri");
                    if (path.Length == 0 || IsUrl(path) || dataFolder == null)
                        return null;

                    var file = Path.Combine(dataFolder, path.TrimStart('/', '\\'));
                    return File.Exists(file) ? LoadCached("file:" + file, file, isFile: true) : null;

                default:
                    return null;
            }
        }

        private JToken LoadCached(string key, string value, bool isFile)
        {
            JToken data;
            if (loadedData.TryGetValue(key, out data))
                return data;

            if (!isFile)
                data = ParseEmbeddedJson(value);

            else if (value.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                data = ParseCsv(File.ReadAllText(value));

            else
                data = StiRdlxReader.ParseJson(File.ReadAllText(value));

            loadedData[key] = data;
            return data;
        }

        /// <summary>
        /// Parses the JSON written in a connection string. Some designers write it escaped, with "\n" and
        /// "\"" as two characters each.
        /// </summary>
        private static JToken ParseEmbeddedJson(string value)
        {
            try
            {
                return StiRdlxReader.ParseJson(value);
            }
            catch (JsonException)
            {
                var text = value.Replace("\\r", " ").Replace("\\n", " ").Replace("\\t", " ").Replace("\\\"", "\"");
                return StiRdlxReader.ParseJson(text);
            }
        }

        /// <summary>
        /// Reads the rows of XML data: the query is the path to the elements of the rows.
        /// </summary>
        private static JToken LoadXmlData(string xml, string path)
        {
            var doc = new XmlDocument();
            doc.LoadXml(xml);

            var rows = new JArray();
            var nodes = string.IsNullOrEmpty(path)
                ? doc.DocumentElement.ChildNodes
                : doc.SelectNodes("/" + path.TrimStart('/'));

            foreach (XmlNode node in nodes)
            {
                var element = node as XmlElement;
                if (element == null)
                    continue;

                var row = new JObject();

                foreach (XmlAttribute attribute in element.Attributes)
                {
                    row[attribute.LocalName] = attribute.Value;
                }

                foreach (XmlNode child in element.ChildNodes)
                {
                    if (child is XmlElement)
                        row[child.LocalName] = ToJsonValue(child.InnerText);
                }

                rows.Add(row);
            }

            return rows;
        }

        private static JToken ParseCsv(string text)
        {
            var rows = new JArray();
            text = text.TrimStart((char)0xFEFF);
            string[] header = null;

            foreach (var line in SplitCsv(text))
            {
                if (header == null)
                {
                    header = line.ToArray();
                    continue;
                }

                var row = new JObject();

                for (var index = 0; index < header.Length && index < line.Count; index++)
                {
                    row[header[index].Trim()] = ToJsonValue(line[index]);
                }

                rows.Add(row);
            }

            return rows;
        }

        private static List<List<string>> SplitCsv(string text)
        {
            var lines = new List<List<string>>();
            var line = new List<string>();
            var value = new StringBuilder();
            var isQuoted = false;

            for (var index = 0; index < text.Length; index++)
            {
                var ch = text[index];

                if (isQuoted)
                {
                    if (ch != '"')
                        value.Append(ch);

                    else if (index + 1 < text.Length && text[index + 1] == '"')
                        value.Append(text[++index]);

                    else
                        isQuoted = false;
                }
                else if (ch == '"')
                {
                    isQuoted = true;
                }
                else if (ch == ',')
                {
                    line.Add(value.ToString());
                    value.Clear();
                }
                else if (ch == '\n')
                {
                    line.Add(value.ToString());
                    value.Clear();

                    if (line.Count > 1 || line[0].Length > 0)
                        lines.Add(line);

                    line = new List<string>();
                }
                else if (ch != '\r')
                {
                    value.Append(ch);
                }
            }

            line.Add(value.ToString());

            if (line.Count > 1 || line[0].Length > 0)
                lines.Add(line);

            return lines;
        }

        /// <summary>
        /// A value of text data: a number when the text is a number, else the text.
        /// </summary>
        private static JToken ToJsonValue(string text)
        {
            decimal number;
            if (decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number) &&
                (text.Length < 2 || text[0] != '0' || text[1] == '.'))
                return new JValue(number);

            return new JValue(text);
        }

        /// <summary>
        /// Applies the JSONPath of the query. An array that the path gives is opened into its items.
        /// </summary>
        private List<JObject> SelectRows(JToken data, string path)
        {
            path = (path ?? "").Trim();

            if (path.Length == 0 || path == "*")
                path = "$";

            if (!path.StartsWith("$"))
                path = "$." + path;

            path = path.Replace(".[", "[");

            // The path of the root alone selects nothing in the JSON library, the root is taken as it is.
            if (path == "$")
                return ToRows(data);

            // ActiveReports takes ".*" and "[*]" as the same: all the items of an array or of an object.
            // The JSON library takes the first only for an object and the second only for an array.
            var paths = new[] { path, path.Replace(".*", "[*]"), path.Replace("[*]", ".*") };
            var rows = new List<JObject>();

            foreach (var candidate in paths)
            {
                foreach (var token in data.SelectTokens(candidate))
                {
                    rows.AddRange(ToRows(token));
                }

                if (rows.Count > 0)
                    break;
            }

            return rows;
        }

        private static List<JObject> ToRows(JToken token)
        {
            var rows = new List<JObject>();

            var array = token as JArray;
            if (array != null)
            {
                foreach (var item in array)
                {
                    rows.AddRange(ToRows(item));
                }
            }
            else if (token is JObject)
            {
                rows.Add((JObject)token);
            }
            else if (token is JValue && token.Type != JTokenType.Null)
            {
                rows.Add(new JObject { ["Value"] = token.DeepClone() });
            }

            return rows;
        }

        /// <summary>
        /// Reads a field of a row. The name may be a path through nested objects ("address.city") and may
        /// end with the type of the value in brackets ("price[Number]").
        /// </summary>
        private static JToken SelectField(JObject row, string dataField)
        {
            var path = StripFieldType(dataField);

            JToken token;
            if (row.TryGetValue(path, out token))
                return token;

            JToken current = row;

            foreach (var part in path.Split('.'))
            {
                var obj = current as JObject;
                if (obj == null || !obj.TryGetValue(part, out current))
                    return null;
            }

            return current;
        }

        private static string StripFieldType(string dataField)
        {
            var bracket = dataField.LastIndexOf('[');

            return bracket > 0 && dataField.EndsWith("]") ? dataField.Substring(0, bracket) : dataField;
        }

        private static string GetFieldTypeHint(string dataField)
        {
            var bracket = dataField.LastIndexOf('[');

            return bracket > 0 && dataField.EndsWith("]")
                ? dataField.Substring(bracket + 1, dataField.Length - bracket - 2).ToLowerInvariant()
                : "";
        }

        /// <summary>
        /// Returns the field of a filter that compares a plain field with constants, null for other filters.
        /// </summary>
        private static string GetStaticFilterField(StiRdlxNode filter)
        {
            var match = Regex.Match(filter.Get("FilterExpression", ""), @"^=\s*Fields!(\w+)\.Value\s*$|^\{(\w+)\}$");
            var values = filter.Items("FilterValues");

            if (!match.Success || values.Count == 0)
                return null;

            foreach (var value in values)
            {
                if (value.Value == null || value.Value.StartsWith("=") || value.Value.IndexOf('{') >= 0)
                    return null;
            }

            switch (filter.Get("Operator", "Equal").ToLowerInvariant())
            {
                case "equal":
                case "notequal":
                case "greaterthan":
                case "greaterthanorequal":
                case "lessthan":
                case "lessthanorequal":
                case "in":
                    return match.Groups[1].Value + match.Groups[2].Value;

                case "between":
                    return values.Count >= 2 ? match.Groups[1].Value + match.Groups[2].Value : null;

                default:
                    return null;
            }
        }

        /// <summary>
        /// All the filters compare fields of the data set with constants.
        /// </summary>
        private static bool AreStaticFilters(List<StiRdlxNode> filters, StiRdlxDataSet dataSet)
        {
            foreach (var filter in filters)
            {
                var field = GetStaticFilterField(filter);
                if (field == null || !dataSet.Columns.ContainsKey(field))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Leaves the rows that pass the filters. The rows are either the rows of the source data, read by
        /// the data fields of the data set, or the rows of the table of the report data, read by the columns.
        /// </summary>
        private static List<JObject> FilterRows(IEnumerable<JObject> rows, List<StiRdlxNode> filters,
            StiRdlxDataSet dataSet, bool isTable)
        {
            var result = new List<JObject>();

            foreach (var row in rows)
            {
                var isAccepted = true;

                foreach (var filter in filters)
                {
                    var field = GetStaticFilterField(filter);
                    var value = isTable ? row[dataSet.Columns[field]] : SelectField(row, GetDataField(dataSet, field));

                    if (!IsFilterPassed(value, filter))
                    {
                        isAccepted = false;
                        break;
                    }
                }

                if (isAccepted)
                    result.Add(row);
            }

            return result;
        }

        private static string GetDataField(StiRdlxDataSet dataSet, string name)
        {
            foreach (var field in dataSet.Node.Items("Fields"))
            {
                if (string.Equals(field.Get("Name"), name, StringComparison.OrdinalIgnoreCase))
                    return field.Get("DataField", name);
            }

            return name;
        }

        private static bool IsFilterPassed(JToken value, StiRdlxNode filter)
        {
            var values = filter.Items("FilterValues");
            var result = CompareFilterValue(value, values[0].Value);

            switch (filter.Get("Operator", "Equal").ToLowerInvariant())
            {
                case "notequal":
                    return result != 0;

                case "greaterthan":
                    return result > 0;

                case "greaterthanorequal":
                    return result >= 0;

                case "lessthan":
                    return result < 0;

                case "lessthanorequal":
                    return result <= 0;

                case "between":
                    return result >= 0 && CompareFilterValue(value, values[1].Value) <= 0;

                case "in":
                    foreach (var item in values)
                    {
                        if (CompareFilterValue(value, item.Value) == 0)
                            return true;
                    }

                    return false;

                default:
                    return result == 0;
            }
        }

        /// <summary>
        /// Compares a value of a row with a constant: as numbers when both are numbers, else as texts.
        /// </summary>
        private static int CompareFilterValue(JToken value, string constant)
        {
            var text = value == null || value.Type == JTokenType.Null ? "" : value.ToString();

            double left;
            double right;

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out left) &&
                double.TryParse(constant, NumberStyles.Float, CultureInfo.InvariantCulture, out right))
                return left.CompareTo(right);

            return string.Compare(text, constant, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Parses "key=value;key=value". A text without a known key is the value of the default key.
        /// </summary>
        private static Dictionary<string, string> ParsePairs(string text, string defaultKey)
        {
            var pairs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var known = new[] { "uri", "jpath", "method", "body", "path" };

            foreach (var part in Regex.Split(text ?? "", @"(?<!\\);"))
            {
                var separator = part.IndexOf('=');
                var key = separator < 0 ? "" : part.Substring(0, separator).Trim().ToLowerInvariant();

                if (Array.IndexOf(known, key) >= 0)
                    pairs[key] = part.Substring(separator + 1).Trim().Replace("\\;", ";");

                else if (!pairs.ContainsKey(defaultKey) && IsPath(part.Trim()))
                    pairs[defaultKey] = part.Trim();
            }

            return pairs;
        }

        /// <summary>
        /// The text is a path to the rows and not a setting like "HasHeaders=true" of a CSV query.
        /// </summary>
        private static bool IsPath(string text)
        {
            return text.Length > 0 && (text[0] == '$' || text[0] == '*' || text.IndexOf('=') < 0);
        }

        private static string GetPair(Dictionary<string, string> pairs, string key)
        {
            string value;
            return pairs.TryGetValue(key, out value) ? value : "";
        }

        private static bool IsUrl(string path)
        {
            return path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
        }
        #endregion

        #region Methods.Lookup
        /// <summary>
        /// Converts Lookup(source, destination, result, "DataSet"): the value of the result in the row of
        /// another data set whose destination equals the source. It becomes the VLookup function of
        /// Stimulsoft, which returns one column of the found row, so every field of the result is looked up
        /// by its own call. Returns null when the function can not be converted.
        /// </summary>
        private string ConvertLookup(List<string> arguments)
        {
            if (arguments.Count < 4 || !IsStringLiteral(arguments[3]))
                return null;

            var lookupSet = FindDataSet(GetLiteralText(arguments[3]));
            var destination = Regex.Match(arguments[1], @"^\w+\.(\w+)$");

            if (lookupSet == null || !destination.Success)
            {
                Warn("A Lookup function is not imported: its data set is not found or its key is not a plain field.");
                return null;
            }

            var result = arguments[2];
            var columns = new List<string>();

            foreach (Match match in Regex.Matches(result, @"\b" + Regex.Escape(lookupSet.SourceName) + @"\.(\w+)"))
            {
                if (!columns.Contains(match.Groups[1].Value))
                    columns.Add(match.Groups[1].Value);
            }

            // A field that both data sets have is read as a field of the outer data set, though the result
            // of Lookup always reads the data set of the lookup.
            var plain = Regex.Match(result, @"^\w+\.(\w+)$");

            if (columns.Count == 0 && plain.Success && lookupSet.Columns.ContainsValue(plain.Groups[1].Value))
                return GetLookupCall(arguments[0], lookupSet, destination.Groups[1].Value, plain.Groups[1].Value);

            foreach (var column in columns)
            {
                result = Regex.Replace(result, @"\b" + Regex.Escape(lookupSet.SourceName + "." + column) + @"\b",
                    GetLookupCall(arguments[0], lookupSet, destination.Groups[1].Value, column));
            }

            return columns.Count > 1 ? "(" + result + ")" : result;
        }

        /// <summary>
        /// The call of VLookup for one column. The function returns an object, a text column is converted
        /// to a text so that it can be joined with other texts and its methods can be called.
        /// </summary>
        private string GetLookupCall(string value, StiRdlxDataSet lookupSet, string keyColumn, string column)
        {
            var call = "VLookup(" + value + ", " + lookupSet.SourceName + ", " + ToLiteral(keyColumn) + ", " +
                ToLiteral(column) + ")";

            return GetColumnType(lookupSet, column) == typeof(string) ? "ToString(" + call + ")" : call;
        }
        #endregion

        #region Methods.DataSources
        private void BuildDataSource(StiRdlxDataSet dataSet)
        {
            if (dataSet.Rows != null)
            {
                BuildEmbeddedSource(dataSet);
                return;
            }

            var query = GetQuery(dataSet);
            StiRdlxNode sourceNode;
            dataSourceNodes.TryGetValue(query.Get("DataSourceName", ""), out sourceNode);

            if (dataSet.Parent != null)
            {
                // The rows of the parent are not known, so the nested rows are not known either.
                dataSet.Source = new StiDataTableSource(dataSet.SourceName, dataSet.SourceName, dataSet.Name);
                AddDeclaredColumns(dataSet, dataSet.Source);
                report.Dictionary.DataSources.Add(dataSet.Source);
                return;
            }

            var connection = sourceNode == null ? null : sourceNode.Child("ConnectionProperties");
            var provider = connection == null ? "" : connection.Get("DataProvider", "").ToUpperInvariant();
            var connectString = connection == null ? "" : connection.Get("ConnectString", "");
            var commandText = query.Get("CommandText", "");

            StiDataSource source;

            if (provider.StartsWith("JSON") || provider == "CSV")
            {
                source = CreateRemoteSource(dataSet, connectString, commandText);
            }
            else if (provider.Contains("OLEDB"))
            {
                var database = AddDatabase(new StiOleDbDatabase(), sourceNode, connectString);
                source = new StiOleDbSource(database, dataSet.SourceName, dataSet.Name, commandText);
            }
            else if (provider.Contains("ODBC"))
            {
                var database = AddDatabase(new StiOdbcDatabase(), sourceNode, connectString);
                source = new StiOdbcSource(database, dataSet.SourceName, dataSet.Name, commandText);
            }
            else if (provider.Contains("SQL"))
            {
                var database = AddDatabase(new StiSqlDatabase(), sourceNode, connectString);
                source = new StiSqlSource(database, dataSet.SourceName, dataSet.Name, commandText);
            }
            else
            {
                Warn($"The data provider '{provider}' of the data set '{dataSet.Name}' is not supported, " +
                    "the data source is added without a connection.");
                source = new StiDataTableSource(dataSet.SourceName, dataSet.SourceName, dataSet.Name);
            }

            AddDeclaredColumns(dataSet, source);

            dataSet.Source = source;
            report.Dictionary.DataSources.Add(source);
        }

        /// <summary>
        /// Adds the columns of the fields declared in the data set, their data is not known at the import.
        /// </summary>
        private void AddDeclaredColumns(StiRdlxDataSet dataSet, StiDataSource source)
        {
            foreach (var field in dataSet.Node.Items("Fields"))
            {
                var name = field.Get("Name", "");
                var dataField = field.Get("DataField");
                if (dataField == null || !dataSet.Columns.ContainsKey(name))
                    continue;

                source.Columns.Add(new StiDataColumn(StripFieldType(dataField), dataSet.Columns[name], name,
                    GetHintType(GetFieldTypeHint(dataField))));
            }
        }

        private string AddDatabase(StiSqlDatabase database, StiRdlxNode sourceNode, string connectString)
        {
            var name = sourceNode.Get("Name", "Connection");

            foreach (StiDatabase existing in report.Dictionary.Databases)
            {
                if (existing.Name == name)
                    return name;
            }

            database.Name = name;
            database.Alias = name;
            database.ConnectionString = connectString;
            report.Dictionary.Databases.Add(database);

            return name;
        }

        /// <summary>
        /// The data is on a server or in a file that is not found: the data set reads a JSON database
        /// by the address, the table is the one the path of the query points to.
        /// </summary>
        private StiDataSource CreateRemoteSource(StiRdlxDataSet dataSet, string connectString, string commandText)
        {
            var address = "";
            var separator = connectString.IndexOf('=');
            var kind = separator < 0 ? "" : connectString.Substring(0, separator).Trim().ToLowerInvariant();

            if (kind == "jsondoc" || kind == "endpoint" || kind == "csvdoc")
                address = connectString.Substring(separator + 1).Split(';')[0].Trim();

            if (commandText.StartsWith("=") || connectString.StartsWith("="))
            {
                Warn($"The query of the data set '{dataSet.Name}' is an expression, the address of its data " +
                    "is not known.");
            }

            var command = ParsePairs(commandText.StartsWith("=") ? "" : commandText, "jpath");
            address += GetPair(command, "uri");

            if (address.Length == 0)
            {
                Warn($"The data of the data set '{dataSet.Name}' is not found, the data source is added " +
                    "without a connection.");
                return new StiDataTableSource(dataSet.SourceName, dataSet.SourceName, dataSet.Name);
            }

            string databaseName;
            if (!remoteDatabases.TryGetValue(address, out databaseName))
            {
                databaseName = CreateName(dataSet.Name + "Data");
                remoteDatabases[address] = databaseName;
                report.Dictionary.Databases.Add(new StiJsonDatabase(databaseName, address));
            }

            if (!IsUrl(address))
                Warn($"The data file '{address}' of the data set '{dataSet.Name}' is not found beside the report.");

            var table = GetTableName(GetPair(command, "jpath"));
            return new StiDataTableSource(databaseName + "." + table, dataSet.SourceName, dataSet.Name);
        }

        private static List<string> GetPathNames(string path)
        {
            var names = new List<string>();

            foreach (Match match in Regex.Matches(path ?? "", @"[A-Za-z_][A-Za-z0-9_]*"))
            {
                names.Add(match.Value);
            }

            return names;
        }

        /// <summary>
        /// The name of the table that the JSON database of Stimulsoft makes of the array the path points to.
        /// </summary>
        private static string GetTableName(string path)
        {
            var names = GetPathNames(path);
            if (names.Count == 0)
                return "root";

            // An answer of OData is an object with the rows in "value" and with plain properties beside it.
            // For such an object the JSON database makes the table "root" of the properties and names the
            // tables of its arrays with this prefix.
            return names[0] == "value" ? "root_" + string.Join("_", names) : string.Join("_", names);
        }

        private static Type GetHintType(string hint)
        {
            switch (hint)
            {
                case "number":
                    return typeof(decimal);

                case "boolean":
                    return typeof(bool);

                case "date":
                case "datetime":
                    return typeof(DateTime);

                case "string":
                    return typeof(string);

                default:
                    return typeof(object);
            }
        }

        /// <summary>
        /// Writes the rows of a data set into the JSON resource of the report and adds a data source of it.
        /// </summary>
        private void BuildEmbeddedSource(StiRdlxDataSet dataSet)
        {
            var dataFields = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (var field in dataSet.Node.Items("Fields"))
            {
                var name = field.Get("Name", "");
                var dataField = field.Get("DataField");

                if (dataField != null && dataSet.Columns.ContainsKey(name))
                    dataFields[name] = dataField;
            }

            if (dataFields.Count == 0 && dataSet.Rows.Count > 0)
            {
                var used = new HashSet<string>(dataSet.Columns.Values, StringComparer.OrdinalIgnoreCase);

                foreach (var property in dataSet.Rows[0].Properties())
                {
                    if (!(property.Value is JValue) || dataSet.Columns.ContainsKey(property.Name))
                        continue;

                    dataSet.Columns[property.Name] = CreateColumnName(property.Name, used);
                    dataFields[property.Name] = property.Name;
                }
            }

            var hasChildren = false;

            foreach (var other in dataSetList)
            {
                if (other.Parent == dataSet)
                    hasChildren = true;
            }

            List<int> parents;
            parentRows.TryGetValue(dataSet.Name, out parents);

            var table = new JArray();
            var types = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase);

            for (var index = 0; index < dataSet.Rows.Count; index++)
            {
                var row = new JObject();

                foreach (var pair in dataFields)
                {
                    var value = SelectField(dataSet.Rows[index], pair.Value) as JValue;
                    if (value == null || value.Type == JTokenType.Null)
                        continue;

                    row[dataSet.Columns[pair.Key]] = value.DeepClone();
                    MergeType(types, pair.Key, value);
                }

                if (hasChildren)
                    row[KeyColumn] = index;

                if (parents != null)
                    row[ParentKeyColumn] = parents[index];

                table.Add(row);
            }

            NormalizeNumbers(table, dataSet, dataFields, types);
            NormalizeDates(table, dataSet, dataFields, types);
            embeddedData[dataSet.SourceName] = table;

            var source = new StiDataTableSource(EmbeddedDataName + "." + dataSet.SourceName, dataSet.SourceName,
                dataSet.Name);

            foreach (var pair in dataFields)
            {
                Type type;
                if (!types.TryGetValue(pair.Key, out type))
                    type = GetHintType(GetFieldTypeHint(pair.Value));

                if (type == typeof(object))
                    type = typeof(string);

                var column = dataSet.Columns[pair.Key];
                source.Columns.Add(new StiDataColumn(column, column, pair.Key, type));
            }

            if (hasChildren)
                source.Columns.Add(new StiDataColumn(KeyColumn, KeyColumn, KeyColumn, typeof(long)));

            if (parents != null)
                source.Columns.Add(new StiDataColumn(ParentKeyColumn, ParentKeyColumn, ParentKeyColumn, typeof(long)));

            dataSet.Source = source;
            report.Dictionary.DataSources.Add(source);
        }

        /// <summary>
        /// Keeps the type of a column that fits all its values.
        /// </summary>
        private static void MergeType(Dictionary<string, Type> types, string field, JValue value)
        {
            Type type;

            switch (value.Type)
            {
                case JTokenType.Integer:
                    type = typeof(long);
                    break;

                case JTokenType.Float:
                    type = typeof(decimal);
                    break;

                case JTokenType.Boolean:
                    type = typeof(bool);
                    break;

                case JTokenType.Date:
                    type = typeof(DateTime);
                    break;

                default:
                    type = typeof(string);
                    break;
            }

            Type current;
            if (!types.TryGetValue(field, out current) || current == type)
            {
                types[field] = type;
                return;
            }

            var isNumber = (current == typeof(long) || current == typeof(decimal)) &&
                (type == typeof(long) || type == typeof(decimal));

            types[field] = isNumber ? typeof(decimal) : typeof(string);
        }

        /// <summary>
        /// A text column whose values are all numbers becomes a number column when the report counts with
        /// it: multiplies, divides, subtracts or sums it. A column of codes that only look like numbers, and
        /// that the report compares with texts, stays a text column.
        /// </summary>
        private void NormalizeNumbers(JArray table, StiRdlxDataSet dataSet, Dictionary<string, string> dataFields,
            Dictionary<string, Type> types)
        {
            foreach (var field in dataFields.Keys)
            {
                Type type;
                if (!types.TryGetValue(field, out type) || type != typeof(string) || !IsCountedField(field))
                    continue;

                var column = dataSet.Columns[field];
                var numbers = new List<KeyValuePair<JObject, decimal>>();
                var isNumber = true;
                var isInteger = true;

                foreach (JObject row in table)
                {
                    var value = row[column];
                    if (value == null)
                        continue;

                    var text = value.ToString().Trim();
                    decimal number;

                    if (!Regex.IsMatch(text, @"^-?\d+(\.\d+)?$") || !decimal.TryParse(text, NumberStyles.Number,
                        CultureInfo.InvariantCulture, out number))
                    {
                        isNumber = false;
                        break;
                    }

                    if (text.Contains("."))
                        isInteger = false;

                    numbers.Add(new KeyValuePair<JObject, decimal>(row, number));
                }

                if (!isNumber || numbers.Count == 0)
                    continue;

                foreach (var pair in numbers)
                {
                    pair.Key[column] = isInteger ? new JValue((long)pair.Value) : new JValue(pair.Value);
                }

                types[field] = isInteger ? typeof(long) : typeof(decimal);
            }
        }

        /// <summary>
        /// The report counts with the field: an expression multiplies, divides or subtracts it, or takes its
        /// sum or average.
        /// </summary>
        private bool IsCountedField(string field)
        {
            var name = Regex.Escape(field);
            var value = @"Fields!" + name + @"\.Value";
            var options = RegexOptions.IgnoreCase | RegexOptions.CultureInvariant;

            return Regex.IsMatch(expressionText, value + @"\s*[*/-]", options) ||
                Regex.IsMatch(expressionText, @"[*/-]\s*" + value, options) ||
                Regex.IsMatch(expressionText, @"\b(Sum|Avg)\(\s*" + value, options) ||
                Regex.IsMatch(expressionText, @"\{[^{}]*(?<![\w.!])" + name + @"\s*[*/][^{}]*\}", options) ||
                Regex.IsMatch(expressionText, @"\{[^{}]*[*/]\s*" + name + @"(?![\w.])[^{}]*\}", options) ||
                Regex.IsMatch(expressionText, @"\{[^{}]*\b(Sum|Avg)\(\s*" + name + @"\s*[,)][^{}]*\}", options);
        }

        /// <summary>
        /// A text column whose values are all dates becomes a date column. The dates are rewritten in the
        /// full ISO form: the JSON database of Stimulsoft takes only such a text as a date.
        /// </summary>
        private static void NormalizeDates(JArray table, StiRdlxDataSet dataSet, Dictionary<string, string> dataFields,
            Dictionary<string, Type> types)
        {
            foreach (var field in dataFields.Keys)
            {
                Type type;
                if (!types.TryGetValue(field, out type) || type != typeof(string))
                    continue;

                var column = dataSet.Columns[field];
                var dates = new List<KeyValuePair<JObject, DateTime>>();
                var isDate = true;

                foreach (JObject row in table)
                {
                    var value = row[column];
                    if (value == null)
                        continue;

                    DateTime date;
                    if (!TryParseDate(value.ToString(), out date))
                    {
                        isDate = false;
                        break;
                    }

                    dates.Add(new KeyValuePair<JObject, DateTime>(row, date));
                }

                if (!isDate || dates.Count == 0)
                    continue;

                foreach (var pair in dates)
                {
                    pair.Key[column] = pair.Value.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture);
                }

                types[field] = typeof(DateTime);
            }
        }

        /// <summary>
        /// Parses a date written as ISO (2020-01-31, 2020-01-31T10:00:00) or as month/day/year.
        /// </summary>
        private static bool TryParseDate(string text, out DateTime date)
        {
            date = DateTime.MinValue;
            text = text.Trim();

            var isIso = Regex.IsMatch(text, @"^\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}.*)?$");
            var isUs = Regex.IsMatch(text, @"^\d{1,2}/\d{1,2}/\d{4}(\s+\d{1,2}:\d{2}(:\d{2})?(\s*[AaPp][Mm])?)?$");

            if (!isIso && !isUs)
                return false;

            return DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date);
        }

        /// <summary>
        /// A field with an expression becomes a calculated column.
        /// </summary>
        private void AddCalculatedColumns(StiRdlxDataSet dataSet)
        {
            if (dataSet.Source == null)
                return;

            var scope = new StiRdlxScope { DataSet = dataSet, IsRegion = true };

            foreach (var field in dataSet.Node.Items("Fields"))
            {
                var name = field.Get("Name", "");
                var value = field.Get("Value");

                if (string.IsNullOrEmpty(value) || field.Has("DataField") || !dataSet.Columns.ContainsKey(name))
                    continue;

                var column = dataSet.Columns[name];
                dataSet.Source.Columns.Add(new StiCalcDataColumn(column, name, typeof(object),
                    ConvertToExpression(value, scope)));
            }
        }
        #endregion
    }
}
