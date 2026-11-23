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
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Stimulsoft.Base.Json;
using Stimulsoft.Base.Json.Linq;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The parameters, the data sources and the data sets of the report.
    /// The data of the report stays as it is. The data written in the report itself (jsondata, xmldata) or
    /// lying in a file beside the report becomes a resource of the Stimulsoft report with the same content,
    /// and a database of its kind - JSON, XML or CSV - reads the resource. The data on a server gets
    /// a database with the address of the service. A data set is a data source of a table of the database:
    /// its columns are the fields declared in the data set, with their types.
    /// Everything a data set does with its rows is done by the data sources of Stimulsoft: the filters
    /// of a data set are the filters of a virtual source, a field of a nested object is read through
    /// a relation, the rows of a nested array are linked with the row that holds them by the key the database
    /// gives them, the totals of a chart are counted by a data transformation.
    /// The import itself never asks a server.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Consts.Data
        private const string RelationColumn = "relationId";

        /// <summary>
        /// The ending of the name of a custom property of the report that holds the types of the fields of
        /// a data set: "Orders_Data_Set_Fields".
        /// </summary>
        private const string FieldTypesSuffix = "_Data_Set_Fields";
        #endregion

        #region Fields.Data
        private static readonly Regex FieldTypeRegex = new Regex(
            @"(?<=.)\[\s*(number|string|boolean|date|datetime)\s*(\|[^\]]*)?\]$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private Dictionary<string, StiRdlxDataSet> dataSets =
            new Dictionary<string, StiRdlxDataSet>(StringComparer.OrdinalIgnoreCase);
        private List<StiRdlxDataSet> dataSetList = new List<StiRdlxDataSet>();
        private Dictionary<string, StiRdlxNode> dataSourceNodes =
            new Dictionary<string, StiRdlxNode>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, StiRdlxConnection> connections =
            new Dictionary<string, StiRdlxConnection>(StringComparer.Ordinal);
        private Dictionary<string, StiDataSource> objectSources =
            new Dictionary<string, StiDataSource>(StringComparer.Ordinal);
        private Dictionary<string, string> objectRelations = new Dictionary<string, string>(StringComparer.Ordinal);
        private Dictionary<string, Type> fieldTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase);
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

            ReadFieldTypes();

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
                BuildDataSource(dataSet);
            }

            foreach (var dataSet in dataSetList)
            {
                AddCalculatedColumns(dataSet);
            }

            foreach (var dataSet in dataSetList)
            {
                AddFilteredSource(dataSet);
            }
        }

        /// <summary>
        /// The designer of ActiveReportsJS writes the types of the fields of a data set into a custom property
        /// of the report, named by the data set: {"price":{"dataType":"Float"},"name":{"dataType":"String"}}.
        /// These types are read first, a column of a field gets the type written here.
        /// </summary>
        private void ReadFieldTypes()
        {
            foreach (var property in root.Items("CustomProperties"))
            {
                var name = property.Get("Name", "");
                var value = property.Get("Value");

                if (!name.EndsWith(FieldTypesSuffix, StringComparison.Ordinal) || string.IsNullOrEmpty(value))
                    continue;

                var dataSetName = name.Substring(0, name.Length - FieldTypesSuffix.Length);

                try
                {
                    var fields = StiRdlxReader.ParseJson(value) as JObject;
                    if (fields == null)
                        continue;

                    foreach (var field in fields.Properties())
                    {
                        var description = field.Value as JObject;
                        var dataType = description == null ? null : description["dataType"];
                        var type = dataType == null
                            ? typeof(object)
                            : GetHintType(dataType.ToString().ToLowerInvariant());

                        if (type != typeof(object))
                            fieldTypes[GetFieldKey(dataSetName, field.Name)] = type;
                    }
                }
                catch (JsonException)
                {
                    Warn($"The types of the fields of the data set '{dataSetName}' are not read.");
                }
            }
        }

        private static string GetFieldKey(string dataSetName, string field)
        {
            return dataSetName + "\n" + field;
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

        /// <summary>
        /// The name of a new column of a data source, one that the data source does not have yet.
        /// </summary>
        private static string CreateColumnName(string field, StiDataSource source)
        {
            var baseName = CreateColumnName(field);
            var name = baseName;
            var index = 2;

            while (source.Columns[name] != null)
            {
                name = baseName + index++;
            }

            return name;
        }

        /// <summary>
        /// The data field without the type written after its name. A name in brackets, "['Net Income']",
        /// is not a type.
        /// </summary>
        private static string StripFieldType(string dataField)
        {
            var match = FieldTypeRegex.Match(dataField);

            return match.Success ? dataField.Substring(0, match.Index) : dataField;
        }

        /// <summary>
        /// The type written after the name of a data field: "price[Number]", "date[Date|M/D/YYYY]".
        /// </summary>
        private static string GetFieldTypeHint(string dataField)
        {
            var match = FieldTypeRegex.Match(dataField);

            return match.Success ? match.Groups[1].Value.ToLowerInvariant() : "";
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

        #region Methods.Connections
        /// <summary>
        /// Returns the database of the data of a data source, null when the data has no database: the
        /// connection string is an expression or is of an unknown kind.
        /// </summary>
        private StiRdlxConnection GetConnection(StiRdlxDataSet dataSet, StiRdlxNode sourceNode, string connectString,
            Dictionary<string, string> command)
        {
            var separator = connectString.IndexOf('=');
            if (separator <= 0)
                return null;

            var sourceName = sourceNode.Get("Name", "Data");
            var kind = connectString.Substring(0, separator).Trim().ToLowerInvariant();
            var value = connectString.Substring(separator + 1);

            switch (kind)
            {
                case "jsondata":
                    // A description of the data may follow the data itself.
                    var schema = Regex.Match(value, @";\s*schemadata\s*=");
                    if (schema.Success)
                        value = value.Substring(0, schema.Index);

                    var json = GetEmbeddedJson(value, dataSet);
                    if (json == null)
                        return null;

                    return GetResourceConnection("data:" + sourceName, sourceName, StiResourceType.Json,
                        Encoding.UTF8.GetBytes(json));

                case "xmldata":
                    return GetResourceConnection("data:" + sourceName, sourceName, StiResourceType.Xml,
                        Encoding.UTF8.GetBytes(value));

                case "jsondoc":
                case "endpoint":
                case "csvdoc":
                    var address = value.Split(';')[0].Trim() + GetPair(command, "uri");
                    if (address.Length == 0)
                        return null;

                    return GetAddressConnection(dataSet, sourceName, address);

                default:
                    return null;
            }
        }

        /// <summary>
        /// The JSON written in a connection string. Some designers write it escaped, with a backslash
        /// before every quote and with a line break as a backslash and the letter n. Such a text is
        /// unescaped, a JSON database can not read it otherwise.
        /// Returns null when the text is not a JSON.
        /// </summary>
        private string GetEmbeddedJson(string value, StiRdlxDataSet dataSet)
        {
            var text = UnescapeJson(value);

            // An escaped text is unescaped before the first reading: the reading of a wrong JSON ends with
            // an exception.
            if (IsEscapedJson(value) && IsJson(text))
                return text;

            if (IsJson(value))
                return value;

            if (IsJson(text))
                return text;

            Warn($"The data of the data set '{dataSet.Name}' is not a correct JSON.");
            return null;
        }

        /// <summary>
        /// A backslash stands right after the opening bracket of the text: a JSON has a backslash only
        /// inside its strings, so the text is an escaped one.
        /// </summary>
        private static bool IsEscapedJson(string value)
        {
            var text = value.TrimStart();
            if (text.Length == 0 || (text[0] != '{' && text[0] != '['))
                return false;

            text = text.Substring(1).TrimStart();
            return text.Length > 0 && text[0] == '\\';
        }

        private static string UnescapeJson(string value)
        {
            var builder = new StringBuilder(value.Length);

            for (var index = 0; index < value.Length; index++)
            {
                var ch = value[index];

                if (ch != '\\' || index + 1 >= value.Length)
                {
                    builder.Append(ch);
                    continue;
                }

                var next = value[++index];

                switch (next)
                {
                    case 'n':
                    case 'r':
                    case 't':
                        builder.Append(' ');
                        break;

                    case '"':
                    case '\\':
                        builder.Append(next);
                        break;

                    default:
                        builder.Append(ch).Append(next);
                        break;
                }
            }

            return builder.ToString();
        }

        private static bool IsJson(string text)
        {
            try
            {
                StiRdlxReader.ParseJson(text);
                return true;
            }
            catch (JsonException)
            {
                return false;
            }
        }

        /// <summary>
        /// The database of a file or of a service. A file found beside the report is put into the report
        /// as a resource, as it is.
        /// </summary>
        private StiRdlxConnection GetAddressConnection(StiRdlxDataSet dataSet, string sourceName, string address)
        {
            StiRdlxConnection connection;
            if (connections.TryGetValue(address, out connection))
                return connection;

            var isCsv = address.EndsWith(".csv", StringComparison.OrdinalIgnoreCase);

            if (!IsUrl(address) && dataFolder != null)
            {
                var file = Path.Combine(dataFolder, address.TrimStart('/', '\\'));

                if (File.Exists(file))
                {
                    connection = GetResourceConnection(address, Path.GetFileNameWithoutExtension(file),
                        isCsv ? StiResourceType.Csv : StiResourceType.Json, File.ReadAllBytes(file));

                    return connection;
                }
            }

            if (!IsUrl(address))
                Warn($"The data file '{address}' of the data set '{dataSet.Name}' is not found beside the report.");

            var name = CreateName(sourceName);
            connection = new StiRdlxConnection(name, !isCsv);

            if (isCsv)
                report.Dictionary.Databases.Add(new StiCsvDatabase(name, address));

            else
                report.Dictionary.Databases.Add(CreateJsonDatabase(name, address));

            connections[address] = connection;
            return connection;
        }

        /// <summary>
        /// Puts the data into a resource of the report and adds a database that reads the resource.
        /// </summary>
        private StiRdlxConnection GetResourceConnection(string key, string preferredName, StiResourceType type,
            byte[] content)
        {
            StiRdlxConnection connection;
            if (connections.TryGetValue(key, out connection))
                return connection;

            var name = CreateName(preferredName);
            var path = "resource://" + name;
            StiFileDatabase database;

            switch (type)
            {
                case StiResourceType.Xml:
                    database = new StiXmlDatabase(name, path);
                    database.CastToColumnType = StiColumnTypeProcessing.CastToColumnType;
                    break;

                case StiResourceType.Csv:
                    database = new StiCsvDatabase(name, path);
                    database.CastToColumnType = StiColumnTypeProcessing.CastToColumnType;
                    break;

                default:
                    database = CreateJsonDatabase(name, path);
                    break;
            }

            report.Dictionary.Resources.Add(new StiResource(name, type, content));
            report.Dictionary.Databases.Add(database);

            connection = new StiRdlxConnection(name, type == StiResourceType.Json);
            connection.Tables = GetTableNames(database);
            connections[key] = connection;

            return connection;
        }

        /// <summary>
        /// A JSON database that gives a column the type declared in the data source: a number or a date
        /// written as a text becomes a number or a date.
        /// </summary>
        private static StiJsonDatabase CreateJsonDatabase(string name, string path)
        {
            var database = new StiJsonDatabase(name, path);
            database.CastToColumnType = StiColumnTypeProcessing.CastToColumnType;

            return database;
        }

        /// <summary>
        /// The names of the tables the database makes of the data of a resource.
        /// </summary>
        private List<string> GetTableNames(StiFileDatabase database)
        {
            try
            {
                var information = database.GetDatabaseInformation(report);
                if (information == null)
                    return null;

                var tables = new List<string>();

                foreach (DataTable table in information.Tables)
                {
                    tables.Add(table.TableName);
                }

                return tables;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static List<string> GetPathNames(string path)
        {
            var names = new List<string>();

            foreach (Match match in Regex.Matches(path ?? "", @"[\p{L}_][\p{L}\p{N}_]*"))
            {
                names.Add(match.Value);
            }

            return names;
        }

        /// <summary>
        /// The table of the database that holds the rows the path points to. A JSON database names the table
        /// of an array by the names of the path to it; when the root of the data has plain properties beside
        /// the arrays, as an answer of OData has, the names get the prefix "root". Returns null when the
        /// database has no such table.
        /// </summary>
        private static string FindTable(StiRdlxConnection connection, List<string> names)
        {
            var joined = string.Join("_", names);

            if (connection.Tables == null)
            {
                if (!connection.IsJson)
                    return names.Count == 0 ? null : names[names.Count - 1];

                if (names.Count == 0)
                    return "root";

                return names[0] == "value" ? "root_" + joined : joined;
            }

            if (!connection.IsJson)
            {
                if (names.Count > 0 && connection.Tables.Contains(names[names.Count - 1]))
                    return names[names.Count - 1];

                return connection.Tables.Count > 0 ? connection.Tables[0] : null;
            }

            if (names.Count == 0)
                return connection.Tables.Contains("root") ? "root" : null;

            if (connection.Tables.Contains(joined))
                return joined;

            if (connection.Tables.Contains("root_" + joined))
                return "root_" + joined;

            // A path with a wildcard in the middle has no names for a part of the way.
            string found = null;

            foreach (var table in connection.Tables)
            {
                if (!table.EndsWith("_" + joined, StringComparison.Ordinal))
                    continue;

                if (found != null)
                    return null;

                found = table;
            }

            return found;
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
        /// <summary>
        /// Makes the data source of a data set. The data source of the parent of a nested data set is made
        /// first.
        /// </summary>
        private void BuildDataSource(StiRdlxDataSet dataSet)
        {
            if (dataSet.IsBuilt)
                return;

            dataSet.IsBuilt = true;

            // The rows of a data set with filters are filtered by a virtual source that takes the name of
            // the data set, the data source of the table gets another name.
            dataSet.BaseName = dataSet.Node.Items("Filters").Count > 0
                ? CreateName(dataSet.SourceName + "All")
                : dataSet.SourceName;

            if (dataSet.Parent != null)
            {
                BuildDataSource(dataSet.Parent);
                BuildNestedSource(dataSet);
                return;
            }

            var query = GetQuery(dataSet);
            StiRdlxNode sourceNode;
            dataSourceNodes.TryGetValue(query.Get("DataSourceName", ""), out sourceNode);

            var connection = sourceNode == null ? null : sourceNode.Child("ConnectionProperties");
            var provider = connection == null ? "" : connection.Get("DataProvider", "").ToUpperInvariant();
            var connectString = connection == null ? "" : connection.Get("ConnectString", "");
            var commandText = query.Get("CommandText", "");

            StiDataSource source;

            if (provider.StartsWith("JSON") || provider == "CSV" || provider == "XML")
            {
                source = CreateFileSource(dataSet, sourceNode, connectString, commandText);
                if (source == null)
                    return;
            }
            else if (provider.Contains("OLEDB"))
            {
                var database = AddDatabase(new StiOleDbDatabase(), sourceNode, connectString);
                source = new StiOleDbSource(database, dataSet.BaseName, GetBaseAlias(dataSet), commandText);
                AddDeclaredColumns(dataSet, source);
            }
            else if (provider.Contains("ODBC"))
            {
                var database = AddDatabase(new StiOdbcDatabase(), sourceNode, connectString);
                source = new StiOdbcSource(database, dataSet.BaseName, GetBaseAlias(dataSet), commandText);
                AddDeclaredColumns(dataSet, source);
            }
            else if (provider.Contains("SQL"))
            {
                var database = AddDatabase(new StiSqlDatabase(), sourceNode, connectString);
                source = new StiSqlSource(database, dataSet.BaseName, GetBaseAlias(dataSet), commandText);
                AddDeclaredColumns(dataSet, source);
            }
            else
            {
                Warn($"The data provider '{provider}' of the data set '{dataSet.Name}' is not supported, " +
                    "the data source is added without a connection.");
                source = CreateEmptySource(dataSet);
            }

            dataSet.Source = source;
            report.Dictionary.DataSources.Add(source);
        }

        private static string GetBaseAlias(StiRdlxDataSet dataSet)
        {
            return dataSet.BaseName == dataSet.SourceName ? dataSet.Name : dataSet.BaseName;
        }

        /// <summary>
        /// A data source without a connection: the data of the data set is not found.
        /// </summary>
        private StiDataSource CreateEmptySource(StiRdlxDataSet dataSet)
        {
            var source = new StiDataTableSource(dataSet.BaseName, dataSet.BaseName, GetBaseAlias(dataSet));
            AddDeclaredColumns(dataSet, source);

            return source;
        }

        /// <summary>
        /// The data source of a data set that reads JSON, XML or CSV: a table of the database of the data.
        /// Returns null when the data set is the root object of the data and has no table.
        /// </summary>
        private StiDataSource CreateFileSource(StiRdlxDataSet dataSet, StiRdlxNode sourceNode, string connectString,
            string commandText)
        {
            if (commandText.StartsWith("=") || connectString.StartsWith("="))
            {
                Warn($"The query of the data set '{dataSet.Name}' is an expression, the address of its data " +
                    "is not known.");
            }

            var command = ParsePairs(commandText.StartsWith("=") ? "" : commandText, "jpath");

            if (!connectString.StartsWith("="))
                dataSet.Connection = GetConnection(dataSet, sourceNode, connectString, command);

            if (dataSet.Connection == null)
            {
                Warn($"The data of the data set '{dataSet.Name}' is not found, the data source is added " +
                    "without a connection.");
                return CreateEmptySource(dataSet);
            }

            var path = GetPair(command, "jpath");

            if (Regex.IsMatch(path, @"\[\s*\?|\[\s*\d|\.\."))
            {
                Warn($"The path '{path}' of the data set '{dataSet.Name}' has a condition, an index or a search " +
                    "in depth, the data set reads all the rows of the table.");
            }

            dataSet.PathNames = GetPathNames(Regex.Replace(path, @"\[\s*\?.*?\]", ""));
            dataSet.Table = FindTable(dataSet.Connection, dataSet.PathNames);

            if (dataSet.Table == null && dataSet.PathNames.Count == 0 && dataSet.Connection.IsJson)
            {
                // The root object of the data has arrays only, the database makes no table of it.
                dataSet.IsSingleRow = true;
                return CreateRootSource(dataSet);
            }

            if (dataSet.Table == null)
            {
                Warn($"The table of the data set '{dataSet.Name}' is not found in its data by the path '{path}'.");
                dataSet.Table = string.Join("_", dataSet.PathNames);
            }

            return CreateTableSource(dataSet);
        }

        /// <summary>
        /// The data source of a data set that is the root object of the data. Such a data set is one row,
        /// and a field of it may be a property of the first item of an array, "Revenue[*].Sales". The array
        /// is a table of the database; a data transformation takes the first value of the column of this
        /// table and so makes the row of the data set. Returns null when the data set has no such fields.
        /// </summary>
        private StiDataSource CreateRootSource(StiRdlxDataSet dataSet)
        {
            var transformation = new StiDataTransformation(dataSet.BaseName, dataSet.BaseName);
            transformation.Alias = GetBaseAlias(dataSet);

            foreach (var field in dataSet.Node.Items("Fields"))
            {
                var name = field.Get("Name", "");
                var dataField = field.Get("DataField");
                if (dataField == null || !dataSet.Columns.ContainsKey(name))
                    continue;

                var match = Regex.Match(StripFieldType(dataField),
                    @"^([\p{L}_][\p{L}\p{N}_]*)(?:\[\*\])?(?:\.(.+)|\['(.+)'\])$");
                if (!match.Success)
                    continue;

                var table = FindTable(dataSet.Connection, new List<string> { match.Groups[1].Value });
                if (table == null)
                    continue;

                var key = dataSet.Connection.Name + "." + table;

                StiDataSource arraySource;
                if (!objectSources.TryGetValue(key, out arraySource))
                {
                    var sourceName = CreateName(dataSet.SourceName + "_" + CreateColumnName(match.Groups[1].Value));
                    arraySource = new StiDataTableSource(key, sourceName, sourceName);
                    report.Dictionary.DataSources.Add(arraySource);
                    objectSources[key] = arraySource;
                }

                var property = match.Groups[2].Value + match.Groups[3].Value;
                var type = GetFieldType(dataSet, name, GetFieldTypeHint(dataField));
                StiDataColumn column = null;

                foreach (StiDataColumn existing in arraySource.Columns)
                {
                    if (existing.NameInSource == property)
                        column = existing;
                }

                if (column == null)
                {
                    var columnName = CreateColumnName(property, arraySource);
                    column = new StiDataColumn(property, columnName, columnName, type);
                    arraySource.Columns.Add(column);
                }

                transformation.Columns.Add(new StiDataTransformationColumn(dataSet.Columns[name], name, type,
                    "First(" + arraySource.Name + "." + column.Name + ")")
                {
                    Mode = StiDataTransformationMode.Measure
                });
            }

            if (transformation.Columns.Count == 0)
                return null;

            dataSet.IsSingleRow = false;
            return transformation;
        }

        private StiDataSource CreateTableSource(StiRdlxDataSet dataSet)
        {
            var source = new StiDataTableSource(dataSet.Connection.Name + "." + dataSet.Table, dataSet.BaseName,
                GetBaseAlias(dataSet));

            AddDeclaredColumns(dataSet, source);

            return source;
        }

        /// <summary>
        /// The data source of a nested data set: the table the database makes of the array of the parent.
        /// A row of this table has the key of the array it lies in, the row of the parent has the same key
        /// in the column of the array.
        /// </summary>
        private void BuildNestedSource(StiRdlxDataSet dataSet)
        {
            var parent = dataSet.Parent;

            if (parent.Connection == null)
            {
                // The data of the parent has no database, so the nested rows have no table either.
                dataSet.Source = CreateEmptySource(dataSet);
                report.Dictionary.DataSources.Add(dataSet.Source);
                return;
            }

            dataSet.Connection = parent.Connection;
            dataSet.PathNames = new List<string>(parent.PathNames);
            dataSet.PathNames.AddRange(GetPathNames(dataSet.ParentField));
            dataSet.Table = FindTable(dataSet.Connection, dataSet.PathNames);

            if (dataSet.Table == null && parent.Table != null)
                dataSet.Table = parent.Table + "_" + string.Join("_", GetPathNames(dataSet.ParentField));

            if (dataSet.Table == null)
            {
                Warn($"The table of the data set '{dataSet.Name}' is not found in its data.");
                dataSet.Table = string.Join("_", dataSet.PathNames);
            }

            var source = CreateTableSource(dataSet);
            dataSet.Source = source;
            report.Dictionary.DataSources.Add(source);

            var parentSource = parent.Source as StiDataTableSource;
            var isPlainField = dataSet.ParentField.Length > 0 && dataSet.ParentField.IndexOf('.') < 0;

            if (parentSource == null || !isPlainField || !dataSet.Connection.IsJson)
                return;

            var parentColumn = GetKeyColumn(parentSource, dataSet.ParentField);
            var keyColumn = GetKeyColumn(source, RelationColumn);
            var relationName = CreateName(parent.BaseName + "_" + CreateColumnName(dataSet.ParentField));

            report.Dictionary.Relations.Add(new StiDataRelation(relationName, relationName, relationName, parentSource,
                source, new[] { parentColumn.Name }, new[] { keyColumn.Name }));

            dataSet.ParentColumn = parentColumn.Name;
            dataSet.KeyColumn = keyColumn.Name;
        }

        /// <summary>
        /// The column of a data source that holds a key the JSON database gives to a nested array or object.
        /// The key is a text.
        /// </summary>
        private static StiDataColumn GetKeyColumn(StiDataSource source, string nameInSource)
        {
            foreach (StiDataColumn column in source.Columns)
            {
                if (column is StiCalcDataColumn || column.NameInSource != nameInSource)
                    continue;

                column.Type = typeof(string);
                return column;
            }

            var name = CreateColumnName(nameInSource, source);
            var keyColumn = new StiDataColumn(nameInSource, name, name, typeof(string));
            source.Columns.Add(keyColumn);

            return keyColumn;
        }

        /// <summary>
        /// Adds the columns of the fields declared in the data set. The type of a column is the type written
        /// in its data field; the database gives the values this type.
        /// </summary>
        private void AddDeclaredColumns(StiRdlxDataSet dataSet, StiDataSource source)
        {
            foreach (var field in dataSet.Node.Items("Fields"))
            {
                var name = field.Get("Name", "");
                var dataField = field.Get("DataField");
                if (dataField == null || !dataSet.Columns.ContainsKey(name))
                    continue;

                var path = StripFieldType(dataField).TrimStart('@');
                var type = GetFieldType(dataSet, name, GetFieldTypeHint(dataField));
                var column = dataSet.Columns[name];

                if (IsObjectField(dataSet, path))
                    AddObjectColumn(dataSet, source, column, name, path, type);

                else
                    source.Columns.Add(new StiDataColumn(path, column, name, type));
            }
        }

        /// <summary>
        /// The data field is a path to a property of a nested object, "customer.name", and not a property
        /// whose name has a dot. The JSON database makes a table of the nested objects; when the tables of
        /// the database are known, the table of the object has to be among them.
        /// </summary>
        private static bool IsObjectField(StiRdlxDataSet dataSet, string path)
        {
            var dot = path.IndexOf('.');
            if (dot <= 0 || dataSet.Connection == null || !dataSet.Connection.IsJson || dataSet.Table == null)
                return false;

            return dataSet.Connection.Tables == null ||
                dataSet.Connection.Tables.Contains(dataSet.Table + "_" + path.Substring(0, dot));
        }

        /// <summary>
        /// The type of the column of a field. It is the type the report declares: the one written in the
        /// custom property with the types of the fields of the data set, else the one written after the name
        /// of the data field. Only when the report declares no type, or declares a text, the use of the field
        /// is looked at: a field the report counts with - multiplies, divides, subtracts or sums it - is
        /// a number even when its values are written as texts.
        /// </summary>
        private Type GetFieldType(StiRdlxDataSet dataSet, string field, string hint)
        {
            Type type;
            if (!fieldTypes.TryGetValue(GetFieldKey(dataSet.Name, field), out type))
                type = GetHintType(hint);

            if ((type == typeof(object) || type == typeof(string)) && IsCountedField(field))
                return typeof(decimal);

            return type;
        }

        /// <summary>
        /// A field of a nested object, "customer.name". The JSON database makes a table of the objects and
        /// writes the key of the object into the row that holds it. A data source reads this table, a relation
        /// links it with the data source of the rows, and the field is a calculated column that reads the
        /// column of the object through the relation.
        /// </summary>
        private void AddObjectColumn(StiRdlxDataSet dataSet, StiDataSource source, string column, string alias,
            string path, Type type)
        {
            var parts = path.Split('.');
            var current = source;
            var table = dataSet.Table;
            var reference = source.Name;

            for (var index = 0; index < parts.Length - 1; index++)
            {
                table += "_" + parts[index];
                var key = dataSet.Connection.Name + "." + table;

                StiDataSource objectSource;
                if (!objectSources.TryGetValue(key, out objectSource))
                {
                    var name = CreateName(current.Name + "_" + CreateColumnName(parts[index]));
                    objectSource = new StiDataTableSource(key, name, name);
                    report.Dictionary.DataSources.Add(objectSource);
                    objectSources[key] = objectSource;

                    var objectKey = GetKeyColumn(objectSource, RelationColumn);
                    var rowKey = GetKeyColumn(current, parts[index]);
                    var relationName = CreateName(CreateColumnName(parts[index]));

                    report.Dictionary.Relations.Add(new StiDataRelation(relationName, relationName, relationName,
                        objectSource, current, new[] { objectKey.Name }, new[] { rowKey.Name }));
                    objectRelations[key] = relationName;
                }

                reference += "." + objectRelations[key];
                current = objectSource;
            }

            var last = parts[parts.Length - 1];
            StiDataColumn objectColumn = null;

            foreach (StiDataColumn existing in current.Columns)
            {
                if (existing.NameInSource == last)
                    objectColumn = existing;
            }

            if (objectColumn == null)
            {
                var name = CreateColumnName(last, current);
                objectColumn = new StiDataColumn(last, name, name, type);
                current.Columns.Add(objectColumn);
            }

            source.Columns.Add(new StiCalcDataColumn(column, alias, type, reference + "." + objectColumn.Name));
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

        private static Type GetHintType(string hint)
        {
            switch (hint)
            {
                case "number":
                case "float":
                    return typeof(decimal);

                case "integer":
                    return typeof(long);

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
        /// A field with an expression becomes a calculated column of the data source of the table.
        /// </summary>
        private void AddCalculatedColumns(StiRdlxDataSet dataSet)
        {
            if (dataSet.Source == null)
                return;

            var scope = new StiRdlxScope { DataSet = dataSet, IsRegion = true };
            var sourceName = dataSet.SourceName;

            // The expression of the column reads the data source of the table, not the filtered one.
            dataSet.SourceName = dataSet.BaseName;

            try
            {
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
            finally
            {
                dataSet.SourceName = sourceName;
            }
        }

        /// <summary>
        /// Applies the filters of a data set to all its rows, as ActiveReports does. A data source of
        /// Stimulsoft has no filters, so a virtual source with the name of the data set filters the rows of
        /// the data source of the table: every band, chart and total of the report keeps the name of the data
        /// set and sees the filtered rows. A calculated column stays in the data source of the table, the
        /// virtual source takes its value as a plain column.
        /// </summary>
        private void AddFilteredSource(StiRdlxDataSet dataSet)
        {
            var filters = dataSet.Node.Items("Filters");
            if (filters.Count == 0 || dataSet.Source == null || dataSet.BaseName == dataSet.SourceName)
                return;

            var source = new StiVirtualSource(dataSet.BaseName, dataSet.SourceName);
            source.Alias = dataSet.Name;
            var results = new List<string>();

            foreach (StiDataColumn column in dataSet.Source.Columns)
            {
                source.Columns.Add(new StiDataColumn(column.Name, column.Name, column.Alias, column.Type));
                results.Add(column.Name);
                results.Add(string.Empty);
                results.Add(column.Name);
            }

            source.Results = results.ToArray();

            var scope = new StiRdlxScope { DataSet = dataSet, IsRegion = true };
            var sourceName = dataSet.SourceName;

            // The filters read the data source of the table.
            dataSet.SourceName = dataSet.BaseName;

            try
            {
                foreach (var filter in filters)
                {
                    var expression = ConvertFilter(filter, scope);
                    if (expression == null)
                        continue;

                    source.Filters.Add(new StiFilter("{" + expression + "}"));
                    source.FilterOn = true;
                }
            }
            finally
            {
                dataSet.SourceName = sourceName;
            }

            dataSet.Source = source;
            dataSet.IsFiltered = true;
            report.Dictionary.DataSources.Add(source);
        }
        #endregion
    }
}
