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
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text;
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
            StiDatabase dataBase = null;
            string databaseName = baseNode.ParentNode.Attributes["Name"].Value;
            string dataProvider = null;
            string connectString = $"Data Source=.;Initial Catalog={databaseName};Integrated Security=True";
            string baseConnectionString = null;

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

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            switch (dataProvider)
            {
                case "OLEDB":
                    dataBase = new StiOleDbDatabase(databaseName, connectString);
                    break;

                case "XML":
                    dataBase = new StiXmlDatabase(databaseName, baseConnectionString);
                    break;

                default:
                    dataBase = new StiSqlDatabase(databaseName, connectString);
                    break;
            }
            return dataBase;
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
                    string newField = $"{dataSource.Name}.{column.Name}";
                    fieldsNames[baseField] = newField;
                    fieldsNames[dataSource.Name + ":" + baseField] = newField;
                }
                report.Dictionary.DataSources.Add(dataSource);
            }
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
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "DataField":
                        column.NameInSource = Convert.ToString(GetNodeTextValue(node));
                        break;

                    case "rd:TypeName":
                        string typeName = Convert.ToString(GetNodeTextValue(node));
                        Type type = Type.GetType(typeName);
                        if (type == null) type = typeof(object);
                        column.Type = type;
                        break;

                    case "Value":
                        var calcColumn = new StiCalcDataColumn();
                        calcColumn.Name = column.Name;
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
                        string command = Convert.ToString(GetNodeTextValue(node));
                        commandText = ConvertCommandTextExpression(command);
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
            StiDatabase database = report.Dictionary.Databases[databaseName];
            if (database != null)
            {
                if (database is StiOleDbDatabase)
                {
                    StiOleDbSource oledbSource = new StiOleDbSource();
                    oledbSource.NameInSource = databaseName;
                    oledbSource.SqlCommand = commandText;
                    if (timeout > 0) oledbSource.CommandTimeout = timeout;
                    dataSource = oledbSource;
                }
                if (database is StiXmlDatabase xmlDatabase)
                {
                    #region try to get table name from file
                    string tableName = null;
                    try
                    {
                        DataSet ds = new DataSet();
                        ds.ReadXml(xmlDatabase.PathData);

                        foreach (DataTable table in ds.Tables)
                        {
                            var flag = true;
                            foreach (StiDataColumn column in columns)
                            {
                                if (!table.Columns.Contains(column.Name))
                                {
                                    flag = false;
                                    break;
                                }
                            }
                            if (flag)
                                tableName = ds.DataSetName + "_" + table.TableName;
                        }

                        if ((tableName == null) && ds.Tables.Count > 0)
                            tableName = ds.DataSetName + "_" + ds.Tables[0].TableName;
                    }
                    catch
                    {
                    }
                    #endregion

                    var xmlDataSource = new StiDataTableSource();
                    xmlDataSource.NameInSource = databaseName + (tableName != null ? "." + tableName : string.Empty);
                    dataSource = xmlDataSource;
                }
            }
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
            dataSource.Parameters.AddRange(parameters);
            return dataSource;
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
            parameter.Name = baseNode.Attributes["Name"].Value.Substring(1);    //remove @
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
                        parameter.Expression = st;
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }
            return parameter;
        }

        //Converts a query text that is written as an expression into a text with placeholders.
        //The string literals are unquoted and put into the text as is, everything between them is an expression and becomes a placeholder.
        //A plain text without the leading "=" is returned unchanged.
        private string ConvertCommandTextExpression(string commandText)
        {
            if (!IsExpression(commandText))
                return commandText;

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
                    result.Append(text[index++]);
                }
            }
            AppendCommandTextExpression(result, expression);

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
        private StiFiltersCollection ProcessFiltersType(XmlNode baseNode, string dataSet)
        {
            StiFiltersCollection filters = new StiFiltersCollection();

            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Filter":
                        ProcessFilterType(node, filters, dataSet);
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            return filters;
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

            op = ConvertFilterOperator(op);

            return expression + " " + op + " " + values[0];
        }

        private string CreateInFilterExpression(string expression, List<string> values)
        {
            if (values == null || values.Count == 0)
                return string.Empty;

            if (values.Count == 1)
                return expression + (ConvertSyntaxToCSharp ? " == " : " = ") + values[0];

            var sb = new StringBuilder("(");
            for (var index = 0; index < values.Count; index++)
            {
                if (index > 0)
                    sb.Append(ConvertSyntaxToCSharp ? " || " : " OrElse ");

                sb.Append(expression);
                sb.Append(ConvertSyntaxToCSharp ? " == " : " = ");
                sb.Append(values[index]);
            }
            sb.Append(")");

            return sb.ToString();
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

            return op;
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
