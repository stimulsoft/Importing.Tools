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
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using RasData = CrystalDecisions.ReportAppServer.DataDefModel;

namespace Import.CrystalReports
{
    /// <summary>
    /// The databases and the SQL of the data sources: SQL Server through OLE DB becomes an MS SQL database, ODBC
    /// becomes an ODBC database, a Command keeps its SQL with the parameters replaced by the variables, and a stored
    /// procedure is called as a stored procedure with its parameters bound to the variables.
    /// </summary>
    public sealed partial class StiCrystalReportsHelper
    {
        #region Consts.Data
        private static readonly string[] SqlServerProviders = { "SQLOLEDB", "SQLNCLI", "MSOLEDBSQL" };

        private static readonly Regex CommandParameter = new Regex(@"\{\?([^}]+)\}", RegexOptions.Compiled);
        #endregion

        #region Methods.Data
        private static bool IsSqlServerProvider(string provider)
        {
            return provider != null &&
                SqlServerProviders.Any(p => provider.StartsWith(p, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// The connection string of MS SQL: the user of the report, or the integrated security if the report keeps
        /// no user. The password is never saved in a Crystal report.
        /// </summary>
        private static string GetSqlServerConnectionString(string server, string database, Hashtable logon,
            string user)
        {
            var catalog = logon["Initial Catalog"] as string;
            if (string.IsNullOrEmpty(catalog))
                catalog = database;

            var security = string.IsNullOrEmpty(user)
                ? "Integrated Security=True"
                : $"User ID={user};Password=";

            return $"Data Source={server};Initial Catalog={catalog};{security}";
        }

        private static string GetOdbcConnectionString(string server, Hashtable logon, string user)
        {
            var dsn = logon["DSN"] as string;
            if (string.IsNullOrEmpty(dsn))
                dsn = server;

            var connectionString = $"Dsn={dsn}";
            var database = logon["Database"] as string;
            if (!string.IsNullOrEmpty(database))
                connectionString += $";Database={database}";

            if (!string.IsNullOrEmpty(user))
                connectionString += $";Uid={user};Pwd=";

            return connectionString;
        }

        /// <summary>
        /// Sets the SQL of a data source by the table of the RAS: the text of a Command, the call of a stored procedure
        /// or the selection of a table. It runs after the parameters became the variables of the dictionary.
        /// </summary>
        private void AddSqlCommand(StiSqlSource source, Table table)
        {
            pendingTranslations.Add(() => ApplySqlCommand(source, table));
        }

        private void ApplySqlCommand(StiSqlSource source, Table table)
        {
            var rasTable = GetRasObject(table, "RasTable");

            var command = rasTable as RasData.CommandTable;
            if (command != null)
            {
                var text = command.CommandText ?? string.Empty;
                source.SqlCommand = CommandParameter.Replace(text, ReplaceCommandParameter);
                return;
            }

            var procedure = rasTable as RasData.Procedure;
            if (procedure != null && source.GetType() == typeof(StiSqlSource))
            {
                source.Type = StiSqlSourceType.StoredProcedure;
                source.SqlCommand = GetProcedureName(procedure, table);
                foreach (RasData.ParameterField parameter in procedure.Parameters)
                {
                    source.Parameters.Add(CreateProcedureParameter(parameter));
                }

                return;
            }

            var name = table.Location;
            if (string.IsNullOrEmpty(name))
                name = table.Name;

            source.SqlCommand = source is StiOdbcSource || source is StiOleDbSource
                ? $"select * from \"{name}\""
                : $"select * from [{name}]";
        }

        /// <summary>
        /// A parameter of a Command, {?Name}, becomes the variable of the parameter, which Stimulsoft inserts into
        /// the text of the SQL as Crystal does.
        /// </summary>
        private string ReplaceCommandParameter(Match match)
        {
            var variable = FindParameterVariable(match.Groups[1].Value);
            if (variable == null)
            {
                warningList.Add(StiWarning.GetWarningFormula("SQL", $"Unknown parameter {match.Value}"));
                return match.Value;
            }

            return "{" + variable + "}";
        }

        /// <summary>
        /// Finds the variable of a parameter by its name; the names of Crystal are case insensitive, a Command
        /// may write {?@lcWoNo} for the parameter @lcWono.
        /// </summary>
        private string FindParameterVariable(string parameter)
        {
            var key = "?" + parameter;
            foreach (DictionaryEntry entry in htVariableConversion)
            {
                if (string.Equals((string)entry.Key, key, StringComparison.OrdinalIgnoreCase))
                    return (string)entry.Value;
            }

            return null;
        }

        /// <summary>
        /// The name of a stored procedure without the number of the version, like dbo.rptOrders;1.
        /// </summary>
        private static string GetProcedureName(RasData.Procedure procedure, Table table)
        {
            var name = procedure.QualifiedName;
            if (string.IsNullOrEmpty(name))
                name = table.Location;

            var versionIndex = name.IndexOf(';');
            return versionIndex > 0 ? name.Substring(0, versionIndex) : name;
        }

        private StiDataParameter CreateProcedureParameter(RasData.ParameterField parameter)
        {
            var name = parameter.Name.StartsWith("@") ? parameter.Name : "@" + parameter.Name;
            var variable = FindParameterVariable(parameter.Name) ?? FindParameterVariable(name) ?? ReplaceSymbols(name);

            return new StiDataParameter(name, variable, (int)GetSqlDbType(parameter.Type), 0);
        }

        private static SqlDbType GetSqlDbType(RasData.CrFieldValueTypeEnum type)
        {
            switch (type)
            {
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeBooleanField:
                    return SqlDbType.Bit;

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeCurrencyField:
                    return SqlDbType.Money;

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeDateField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeDateTimeField:
                    return SqlDbType.DateTime;

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt8sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt8uField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt16sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt16uField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt32sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt32uField:
                    return SqlDbType.Int;

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeNumberField:
                    return SqlDbType.Float;

                default:
                    return SqlDbType.NVarChar;
            }
        }

        /// <summary>
        /// The prompt and the default values of a parameter: the prompt becomes the description of the variable,
        /// several default values become the list of the dialog, the first one becomes the value of the variable.
        /// Only strings and numbers are taken: the variable checks its value, a date in a wrong format would stop
        /// the import.
        /// </summary>
        private static void ApplyParameterValues(StiVariable variable, ParameterFieldDefinition parameter)
        {
            if (!string.IsNullOrWhiteSpace(parameter.PromptText))
                variable.Description = parameter.PromptText;

            var isText = variable.Type == typeof(string) || variable.Type == typeof(StringList);
            var isNumber = variable.Type == typeof(double) || variable.Type == typeof(DoubleList);
            if (!isText && !isNumber)
                return;

            var keys = new List<string>();
            var labels = new List<string>();
            foreach (var value in parameter.DefaultValues.OfType<ParameterDiscreteValue>())
            {
                if (value.Value == null)
                    continue;

                keys.Add(System.Convert.ToString(value.Value, CultureInfo.InvariantCulture));
                labels.Add(value.Description ?? string.Empty);
            }

            if (keys.Count == 0)
                return;

            if (keys.Count > 1)
            {
                variable.DialogInfo.ItemsInitializationType = StiItemsInitializationType.Items;
                variable.DialogInfo.Keys = keys.ToArray();
                variable.DialogInfo.Values = labels.ToArray();
                variable.DialogInfo.AllowUserValues = parameter.EnableAllowEditingDefaultValue;
            }

            if (variable.Type == typeof(string) || variable.Type == typeof(double))
                variable.Value = keys[0];
        }

        /// <summary>
        /// Returns the object of the RAS behind an object of the engine. The properties are private in the base
        /// classes, so they are searched through the base types: the reflection of a derived type doesn't see them.
        /// </summary>
        private static object GetRasObject(object engineObject, string property)
        {
            const BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public |
                BindingFlags.DeclaredOnly;

            for (var type = engineObject?.GetType(); type != null; type = type.BaseType)
            {
                var info = type.GetProperty(property, flags);
                if (info != null)
                    return info.GetValue(engineObject, null);
            }

            return null;
        }
        #endregion
    }
}
