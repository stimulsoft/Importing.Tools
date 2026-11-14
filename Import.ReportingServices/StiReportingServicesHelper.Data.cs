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

using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Dictionary.Databases.Azure;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    public partial class StiReportingServicesHelper
    {
        #region DataSources
        private void ProcessDataSourcesType(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSource":
                        ProcessDataSourceType(node, report);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessDataSourceType(XmlNode baseNode, StiReport report)
        {
            StiDatabase dataBase = null;
            string databaseName = baseNode.Attributes["Name"].Value;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "ConnectionProperties":
                        dataBase = ProcessConnectionPropertiesType(node);
                        break;

                    case "DataSourceReference":
                        dataBase = CreateSharedDatabase(databaseName, GetNodeTextValue(node));
                        break;

                    case "rd:DataSourceID":
                    case "rd:SecurityType":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            if (dataBase != null)
            {
                report.Dictionary.Databases.Add(dataBase);
            }
        }

        private StiDatabase ProcessConnectionPropertiesType(XmlNode baseNode)
        {
            string databaseName = baseNode.ParentNode.Attributes["Name"].Value;
            string dataProvider = null;
            string connectString = $"Data Source=.;Initial Catalog={databaseName};Integrated Security=True";
            string baseConnectionString = null;
            string embeddedData = null;

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataProvider":
                        dataProvider = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "ConnectString":
                        baseConnectionString = Convert.ToString(GetNodeTextValue(node));
                        connectString = baseConnectionString;
                        break;

					case "IntegratedSecurity":
						var st = Convert.ToString(GetNodeTextValue(node));
                        if (IsTrue(st))
                            connectString += ";Integrated Security=True";
						break;

                    case "rd:CustomProperties":
                        //ignored or not implemented yet
                        break;

                    //the data of a JSON data source of Bold Reports, see CreateBoldJsonDatabase
                    case "rd:EmbeddedData":
                        embeddedData = node.SelectSingleNode("*[local-name()='Data']")?.InnerText;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return CreateDatabase(databaseName, dataProvider, connectString, baseConnectionString,
                embeddedData: embeddedData);
        }

        #endregion

        #region DataSets
        private void ProcessDataSetsType(XmlNode baseNode, StiReport report)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSet":
                        ProcessDataSetType(node, report);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessDataSetType(XmlNode baseNode, StiReport report)
        {
            StiDataSource dataSource = null;
            ArrayList columns = new ArrayList();
            XmlNode filtersNode = null;

            //first pass, only columns
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Fields":
                        ProcessFieldsType(node, columns);
                        break;
                }
            }

            //second pass
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Query":
                        dataSource = ProcessQueryType(node, report, columns);
                        break;

                    case "rd:DataSetInfo":
                    case "Fields":
                        //ignored or not implemented yet
                        break;

                    case "SharedDataSet":
                        dataSource = ProcessSharedDataSetType(node, report, columns);
                        break;

                    case "Filters":
                        filtersNode = node;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (dataSource != null)
            {
                foreach (StiDataColumn column in columns)
                {
                    column.DataSource = dataSource;
                    dataSource.Columns.Add(column);

                    string baseField = $"Fields!{column.Name}.Value";
                    string newField = GetColumnReference(dataSource.Name, column.Name);
                    fieldsNames[baseField] = newField;
                    fieldsNames[dataSource.Name + ":" + baseField] = newField;
                }
                report.Dictionary.DataSources.Add(dataSource);

                if (filtersNode != null)
                    FilterDataSet(dataSource, filtersNode, report);
            }
        }

        /// <summary>
        /// Applies the filters of a data set to all its rows, as Reporting Services does. A data source of
        /// Stimulsoft has no filters, so the data source of the query takes the name with the suffix All, and
        /// a virtual source with the name of the data set filters its rows; every expression of the report
        /// keeps the name of the data set and sees the filtered rows. The filters read the data source of the
        /// query, the band a virtual source filters with runs over it.
        /// </summary>
        private void FilterDataSet(StiDataSource dataSource, XmlNode filtersNode, StiReport report)
        {
            var dataSetName = dataSource.Name;
            var allName = StiNameCreation.CreateDataSourcesName(report, dataSetName + "All");

            dataSource.Name = allName;
            dataSource.Alias = allName;

            foreach (StiDataColumn column in dataSource.Columns)
            {
                fieldsNames[$"{allName}:Fields!{column.Name}.Value"] = GetColumnReference(allName, column.Name);
            }

            var filters = ProcessFiltersType(filtersNode, allName);
            var filteredSource = new StiVirtualSource(allName, dataSetName);
            var results = new List<string>();

            //a calculated field stays in the data source of the query, a filter may read it, and the virtual
            //source takes its value as a plain field: a copy of it would be calculated twice
            foreach (StiDataColumn column in dataSource.Columns)
            {
                var filteredColumn = column is StiCalcDataColumn
                    ? new StiDataColumn(column.Name, column.Type)
                    : (StiDataColumn)column.Clone();

                //the table of the virtual source names its columns by the names of the fields, a field of the
                //query named otherwise, as "name (Sender)", would read a column that is not there
                filteredColumn.NameInSource = column.Name;
                filteredSource.Columns.Add(filteredColumn);
                results.Add(column.Name);
                results.Add(string.Empty);
                results.Add(column.Name);
            }

            filteredSource.Results = results.ToArray();
            filteredSource.Filters = filters;
            report.Dictionary.DataSources.Add(filteredSource);
        }

        private void ProcessFieldsType(XmlNode baseNode, ArrayList columns)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Field":
                        ProcessFieldType(node, columns);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private void ProcessFieldType(XmlNode baseNode, ArrayList columns)
        {
            StiDataColumn column = new StiDataColumn();
            column.Name = baseNode.Attributes["Name"].Value;
            Type type = null;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataField":
                        column.NameInSource = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "rd:TypeName":
                        string typeName = Convert.ToString(GetNodeTextValue(node));
                        type = Type.GetType(typeName);
                        if (type == null) type = typeof(object);
                        column.Type = type;
                        break;

                    //a calculated field of Reporting Services is an Object unless its type is written
                    case "Value":
                        var calcColumn = new StiCalcDataColumn();
                        calcColumn.Name = column.Name;
                        calcColumn.Type = type ?? typeof(object);
                        calcColumn.Value = Convert.ToString(GetNodeTextValue(node)).Trim();
                        column = calcColumn;
                        break;

                    case "rd:UserDefined":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            columns.Add(column);
        }

        private StiDataSource ProcessQueryType(XmlNode baseNode, StiReport report, ArrayList columns)
        {
            StiDataSource dataSource = null;
            string databaseName = null;
            string commandText = null;
            string rawCommandText = null;
            string commandType = null;
            StiDataParametersCollection parameters = new StiDataParametersCollection();
            int timeout = 0;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataSourceName":
                        databaseName = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "CommandText":
                        rawCommandText = Convert.ToString(GetNodeTextValue(node));
                        commandText = ConvertCommandTextExpression(rawCommandText);
                        break;

                    case "CommandType":
                        commandType = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "Timeout":
                        timeout = Convert.ToInt32(GetNodeTextValue(node));
                        break;

                    case "QueryParameters":
                        ProcessQueryParameters(node, parameters);
                        break;

                    case "rd:UseGenericDesigner":
                    case "QueryDesignerState":
                        //ignored or not implemented yet
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            var dataSetName = baseNode.ParentNode.Attributes["Name"].Value;
            var database = databaseName != null ? report.Dictionary.Databases[databaseName] : null;

            if (databaseName != null && localDataSources.Contains(databaseName))
                dataSource = CreateLocalDataSource(dataSetName);

            else if (database is StiOleDbDatabase)
            {
                StiOleDbSource oledbSource = new StiOleDbSource();
                oledbSource.NameInSource = databaseName;
                oledbSource.SqlCommand = commandText;
                if (timeout > 0) oledbSource.CommandTimeout = timeout;
                dataSource = oledbSource;
            }

            else if (database is StiXmlDatabase jsonDatabase &&
                boldJsonData.TryGetValue(databaseName, out var jsonData))
            {
                dataSource = CreateBoldJsonDataSource(jsonDatabase, jsonData, columns, dataSetName);
            }

            else if (database is StiXmlDatabase xmlDatabase)
                dataSource = CreateXmlDataSource(xmlDatabase, rawCommandText, columns, dataSetName);

            else if (database is StiExcelDatabase excelDatabase &&
                boldExcelData.TryGetValue(databaseName, out var excelData))
            {
                dataSource = CreateBoldExcelDataSource(excelDatabase, excelData, rawCommandText, columns);
            }

            else if (database is StiJsonDatabase flatJsonDatabase &&
                boldFlatJsonData.TryGetValue(databaseName, out var flatJsonData))
            {
                dataSource = CreateBoldFlatJsonDataSource(flatJsonDatabase, flatJsonData, columns);
            }

            else if (database is StiJsonDatabase)
                dataSource = CreateJsonDataSource(database, dataSetName);

            if (dataSource == null)
            {
                if (!string.IsNullOrWhiteSpace(commandText))
                {
                    StiSqlSource sqlSource = new StiSqlSource();
                    sqlSource.NameInSource = databaseName;
                    sqlSource.SqlCommand = commandText;
                    if (commandType == "StoredProcedure") sqlSource.Type = StiSqlSourceType.StoredProcedure;
                    if (timeout > 0) sqlSource.CommandTimeout = timeout;
                    dataSource = sqlSource;
                }
                else
                {
                    StiDataTableSource datatableSource = new StiDataTableSource();
                    datatableSource.NameInSource = databaseName;
                    dataSource = datatableSource;
                }
            }
            dataSource.Name = baseNode.ParentNode.Attributes["Name"].Value;
            dataSource.Alias = dataSource.Name;

            if (dataSource is StiSqlSource && IsMsSqlDatabase(database))
            {
                foreach (StiDataParameter parameter in parameters)
                {
                    SetMsSqlParameterType(parameter);
                }
            }

            if (dataSource is StiSqlSource shadowedSource)
                RenameShadowedParameters(shadowedSource, parameters);

            dataSource.Parameters.AddRange(parameters);
            return dataSource;
        }

        /// <summary>
        /// A query parameter of the name of a variable of the report takes the value of the variable when the
        /// report is interpreted, its own expression is not evaluated: the variables are looked for by the name of
        /// the parameter, without the case. A parameter with another expression, as the list of a variable joined
        /// by commas, gets a name of its own, in the query too; the strings of the query are kept. A stored
        /// procedure names its parameters itself, so its parameter is named in the list of errors.
        /// </summary>
        private void RenameShadowedParameters(StiSqlSource source, StiDataParametersCollection parameters)
        {
            foreach (StiDataParameter parameter in parameters)
            {
                var variable = FindVariableByName(parameter.Name);
                if (variable == null || string.Equals(parameter.Expression?.Trim(), variable.Name,
                    StringComparison.OrdinalIgnoreCase))
                    continue;

                if (source.Type == StiSqlSourceType.StoredProcedure)
                {
                    ThrowError(null, null, "Parameter of a stored procedure takes the value of the variable of its " +
                        $"name in the interpretation: {parameter.Name}");
                    continue;
                }

                var oldName = parameter.Name;
                var newName = oldName + "Query";
                var index = 1;
                while (FindVariableByName(newName) != null || parameters.Cast<StiDataParameter>()
                    .Any(p => string.Equals(p.Name, newName, StringComparison.OrdinalIgnoreCase)))
                {
                    newName = oldName + "Query" + index++;
                }

                parameter.Name = newName;
                source.SqlCommand = Regex.Replace(source.SqlCommand ?? "",
                    @"'(?:[^']|'')*'|@" + Regex.Escape(oldName) + @"\b",
                    m => m.Value.StartsWith("'") ? m.Value : "@" + newName,
                    RegexOptions.IgnoreCase);
            }
        }

        private StiVariable FindVariableByName(string name)
        {
            return report.Dictionary.Variables.Cast<StiVariable>()
                .FirstOrDefault(v => string.Equals(v.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// The other databases derived from the database of MS SQL have their own numbers of the types.
        /// </summary>
        private static bool IsMsSqlDatabase(StiDatabase database)
        {
            return database != null &&
                (database.GetType() == typeof(StiSqlDatabase) || database is StiAzureSqlDatabase);
        }

        /// <summary>
        /// Gives a query parameter of MS SQL the type of the parameter of the report it takes, other values
        /// are strings. The type Text of the parameter, the default, is the old type text of SQL Server: it
        /// can not be compared with a string column, and the query of parameters is prepared before it runs,
        /// which needs the size of a parameter of a variable length; a parameter of the size 0 failed the
        /// query. A string is nvarchar(4000), as Reporting Services passes a parameter.
        /// </summary>
        private void SetMsSqlParameterType(StiDataParameter parameter)
        {
            var type = FindVariable(parameter.Expression?.Trim() ?? "")?.Type;

            if (type == typeof(int) || type == typeof(int?) || type == typeof(short) || type == typeof(short?))
                parameter.Type = (int)SqlDbType.Int;

            else if (type == typeof(long) || type == typeof(long?))
                parameter.Type = (int)SqlDbType.BigInt;

            //a prepared decimal needs its precision and scale, which a parameter of the report does not give
            else if (type == typeof(decimal) || type == typeof(decimal?) || type == typeof(double) ||
                type == typeof(double?) || type == typeof(float) || type == typeof(float?))
                parameter.Type = (int)SqlDbType.Float;

            else if (type == typeof(DateTime) || type == typeof(DateTime?))
                parameter.Type = (int)SqlDbType.DateTime;

            else if (type == typeof(bool) || type == typeof(bool?))
                parameter.Type = (int)SqlDbType.Bit;

            else
            {
                parameter.Type = (int)SqlDbType.NVarChar;
                parameter.Size = 4000;
            }
        }

        private void ProcessQueryParameters(XmlNode baseNode, StiDataParametersCollection parameters)
        {
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "QueryParameter":
                        parameters.Add(ProcessQueryParameter(node));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
        }

        private StiDataParameter ProcessQueryParameter(XmlNode baseNode)
        {
            StiDataParameter parameter = new StiDataParameter();
            parameter.Type = 18;    //Text
            //the name of a query parameter of MS SQL starts with @, a name without it is taken as is
            var name = baseNode.Attributes["Name"].Value;
            parameter.Name = name.StartsWith("@") ? name.Substring(1) : name;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Value":
                        string st = ConvertExpression(GetNodeTextValue(node), null);
                        if (st.StartsWith("{") && st.EndsWith("}"))
                        {
                            st = st.Substring(1, st.Length - 2);
                        }
                        parameter.Expression = JoinListParameter(st);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return parameter;
        }

        /// <summary>
        /// Reporting Services passes the values of a parameter with many values to a query parameter
        /// as one text, the values separated by commas, as a stored procedure expects them. A list
        /// variable given to a query parameter is joined the same way, a list itself cannot be sent.
        /// </summary>
        private string JoinListParameter(string expression)
        {
            var variable = FindVariable(expression.Trim());
            if (variable == null || !IsListType(variable.Type))
                return expression;

            return $"String.Join(\",\", {variable.Name})";
        }

        //Converts a query text that is written as an expression into a text with placeholders.
        //The string literals are unquoted and put into the text as is, everything between them is an expression and becomes a placeholder.
        //A plain text without the leading "=" is returned unchanged.
        private string ConvertCommandTextExpression(string commandText)
        {
            if (!IsExpression(commandText))
                return EscapeBraces(commandText);

            var text = commandText.TrimStart().Substring(1);
            var result = new StringBuilder();
            var expression = new StringBuilder();

            var index = 0;
            while (index < text.Length)
            {
                if (text[index] != '"')
                {
                    expression.Append(text[index++]);
                    continue;
                }

                AppendCommandTextExpression(result, expression);

                index++;
                while (index < text.Length)
                {
                    if (text[index] == '"')
                    {
                        //a doubled quote inside a literal is an escaped quote, not the end of it
                        if (index + 1 >= text.Length || text[index + 1] != '"')
                        {
                            index++;
                            break;
                        }
                        index++;
                    }
                    result.Append(EscapeBraces(text[index++].ToString()));
                }
            }
            AppendCommandTextExpression(result, expression);

            return result.ToString();
        }

        /// <summary>
        /// A brace in the query text opens an expression of the report, so a literal brace, like the
        /// escape sequence {fn NOW()} of ODBC or the JSON query of a web data source, does not compile.
        /// Each literal brace turns into an expression that returns the brace itself.
        /// </summary>
        private static string EscapeBraces(string text)
        {
            if (string.IsNullOrEmpty(text) || text.IndexOfAny(new[] { '{', '}' }) == -1)
                return text;

            var result = new StringBuilder(text.Length + 16);

            foreach (var ch in text)
            {
                if (ch == '{')
                    result.Append("{\"{\"}");

                else if (ch == '}')
                    result.Append("{\"}\"}");

                else
                    result.Append(ch);
            }

            return result.ToString();
        }

        //The text between the literals holds the concatenation operators, they are dropped and the rest, if there is any, becomes a placeholder.
        private void AppendCommandTextExpression(StringBuilder result, StringBuilder expression)
        {
            var text = expression.ToString().Trim().Trim('+', '&').Trim();
            expression.Length = 0;

            if (text.Length > 0)
                result.Append(ConvertExpression("=" + text, null));
        }
        #endregion

        #region Filters
        /// <summary>
        /// Converts the filters. A filter of the first or the last rows, TopN or BottomN, is skipped with a message
        /// unless the caller takes it by itself, see GetTablixRowLimit.
        /// </summary>
        private StiFiltersCollection ProcessFiltersType(XmlNode baseNode, string dataSet, bool rowLimitTaken = false)
        {
            StiFiltersCollection filters = new StiFiltersCollection();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Filter":
                        if (!rowLimitTaken || !IsRowLimitFilter(node))
                            ProcessFilterType(node, filters, dataSet);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return filters;
        }

        private static bool IsRowLimitFilter(XmlNode filter)
        {
            var op = filter.SelectSingleNode("*[local-name()='Operator']")?.InnerText.Trim();

            return op == "TopN" || op == "BottomN";
        }

        private void ProcessFilterType(XmlNode baseNode, StiFiltersCollection filters, string dataSet)
        {
            var expression = string.Empty;
            var op = string.Empty;
            var values = new List<string>();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "FilterExpression":
                        expression = ConvertExpression(GetNodeTextValue(node), dataSet);
                        break;

                    case "Operator":
                        op = GetNodeTextValue(node);
                        break;

                    case "FilterValues":
                        values = ProcessFilterValuesType(node, dataSet);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            expression = UnwrapFilterExpression(expression);
            var filterExpression = CreateFilterExpression(expression, op, values);

            if (!string.IsNullOrWhiteSpace(filterExpression))
                filters.Add(new StiFilter(filterExpression));
        }

        private List<string> ProcessFilterValuesType(XmlNode baseNode, string dataSet)
        {
            var values = new List<string>();
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "FilterValue":
                        values.Add(ProcessFilterValueType(node, dataSet));
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return values;
        }

        private string ProcessFilterValueType(XmlNode baseNode, string dataSet)
        {
            var value = GetNodeTextValue(baseNode);
            var isExpression = IsExpression(value);
            value = UnwrapFilterExpression(ConvertExpression(value, dataSet)).Trim();

            return NormalizeFilterValue(value, baseNode.Attributes["DataType"]?.Value, isExpression);
        }

        private string CreateFilterExpression(string expression, string op, List<string> values)
        {
            if (string.IsNullOrWhiteSpace(expression))
                return string.Empty;

            if (string.Equals(op, "In", StringComparison.OrdinalIgnoreCase))
                return CreateInFilterExpression(expression, values);

            if (string.Equals(op, "Between", StringComparison.OrdinalIgnoreCase))
                return CreateBetweenFilterExpression(expression, values);

            if (values == null || values.Count == 0)
                return string.Empty;

            var filterOperator = ConvertFilterOperator(op);
            if (filterOperator == null)
            {
                ThrowError(null, null, $"Filter operator not supported, the filter is skipped: {op}");
                return string.Empty;
            }

            return expression + " " + filterOperator + " " + values[0];
        }

        /// <summary>
        /// Makes the filter of the operator In: the expression equals one of the values. A value that is a
        /// parameter with many values, a list variable, contains the expression instead.
        /// </summary>
        private string CreateInFilterExpression(string expression, List<string> values)
        {
            if (values == null || values.Count == 0)
                return string.Empty;

            if (values.Count == 1)
                return CreateInFilterPart(expression, values[0]);

            var sb = new StringBuilder("(");
            for (var index = 0; index < values.Count; index++)
            {
                if (index > 0)
                    sb.Append(ConvertSyntaxToCSharp ? " || " : " OrElse ");

                sb.Append(CreateInFilterPart(expression, values[index]));
            }
            sb.Append(")");

            return sb.ToString();
        }

        private string CreateInFilterPart(string expression, string value)
        {
            var type = GetExpressionType(value);
            if (type != null && IsListType(type))
                return $"{value}.Contains({expression})";

            return expression + (ConvertSyntaxToCSharp ? " == " : " = ") + value;
        }

        private string CreateBetweenFilterExpression(string expression, List<string> values)
        {
            if (values == null || values.Count < 2)
                return string.Empty;

            return "(" + expression + " >= " + values[0] + (ConvertSyntaxToCSharp ? " && " : " AndAlso ") + expression + " <= " + values[1] + ")";
        }

        private string ConvertFilterOperator(string op)
        {
            if (string.Equals(op, "Equal", StringComparison.OrdinalIgnoreCase))
                return ConvertSyntaxToCSharp ? "==" : "=";
            if (string.Equals(op, "NotEqual", StringComparison.OrdinalIgnoreCase))
                return ConvertSyntaxToCSharp ? "!=" : "<>";
            if (string.Equals(op, "GreaterThan", StringComparison.OrdinalIgnoreCase))
                return ">";
            if (string.Equals(op, "GreaterThanOrEqual", StringComparison.OrdinalIgnoreCase))
                return ">=";
            if (string.Equals(op, "LessThan", StringComparison.OrdinalIgnoreCase))
                return "<";
            if (string.Equals(op, "LessThanOrEqual", StringComparison.OrdinalIgnoreCase))
                return "<=";
            if (string.Equals(op, "Like", StringComparison.OrdinalIgnoreCase))
                return "Like";

            //TopN, BottomN, TopPercent and BottomPercent pick rows by the order of all of them, a filter of a
            //band of Stimulsoft sees one row at a time
            return null;
        }

        private static string NormalizeFilterValue(string value, string dataType, bool isExpression)
        {
            if (value == null)
                return string.Empty;

            value = value.Trim();

            if (IsSingleQuotedFilterValue(value))
                return "\"" + value.Substring(1, value.Length - 2).Replace("\"", "\\\"") + "\"";

            if (IsBooleanFilterValue(value))
                return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ? "true" : "false";

            if (isExpression || IsDoubleQuotedFilterValue(value) || IsNumericFilterValue(value))
                return value;

            if (string.Equals(dataType, "String", StringComparison.OrdinalIgnoreCase) ||
                string.IsNullOrWhiteSpace(dataType))
                return "\"" + value.Replace("\"", "\\\"") + "\"";

            return value;
        }

        private static bool IsDoubleQuotedFilterValue(string value)
        {
            return value.Length >= 2 && value.StartsWith("\"") && value.EndsWith("\"");
        }

        private static bool IsSingleQuotedFilterValue(string value)
        {
            return value.Length >= 2 && value.StartsWith("'") && value.EndsWith("'");
        }

        private static bool IsNumericFilterValue(string value)
        {
            double number;
            return double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out number) ||
                double.TryParse(value, NumberStyles.Any, CultureInfo.CurrentCulture, out number);
        }

        private static bool IsBooleanFilterValue(string value)
        {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(value, "false", StringComparison.OrdinalIgnoreCase);
        }

        private static string UnwrapFilterExpression(string expression)
        {
            if (expression != null && expression.StartsWith("{") && expression.EndsWith("}"))
                return expression.Substring(1, expression.Length - 2);

            return expression;
        }
        #endregion

        #region Make DataSource Copy
        private string MakeDataSourceCopy(string name, string copySuffix = "Copy")
        {
            if (string.IsNullOrEmpty(name)) 
                return name;

            string newDataSetName = name + copySuffix;
            if (!report.Dictionary.DataSources.ExistsName(newDataSetName))
            {
                var oldSource = report.Dictionary.DataSources[name];
                if (oldSource == null) 
                    return newDataSetName;
                var newSource = new StiVirtualSource(name, newDataSetName);

                var results = new List<string>();
                foreach (StiDataColumn column in oldSource.Columns)
                {
                    var newColumn = column.Clone() as StiDataColumn;

                    newSource.Columns.Add(newColumn);

                    results.Add(column.Name);
                    results.Add(string.Empty);
                    results.Add(column.Name);
                }
                newSource.Results = results.ToArray();

                report.Dictionary.DataSources.Add(newSource);
            }
            return newDataSetName;
        }
        #endregion
    }
}
