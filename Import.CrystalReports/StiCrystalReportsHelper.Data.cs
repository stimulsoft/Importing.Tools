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
using Stimulsoft.Report.Components;
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

        private static readonly string[] OracleProviders = { "OraOLEDB", "MSDAORA" };

        private static readonly string[] AccessProviders = { "Microsoft.ACE.OLEDB", "Microsoft.Jet.OLEDB" };

        private const string AccessProvider = "Microsoft.ACE.OLEDB.12.0";

        private const string ConnectionAccess = "Access/Excel (DAO)";

        private const string ConnectionOracle = "Oracle";

        private const string ConnectionXBase = "xBase";

        private const string ConnectionXml = "XML";

        private const string ConnectionOdbc = "ODBC (RDO)";

        private static readonly Regex CommandParameter = new Regex(@"\{\?([^}]+)\}|[{}]", RegexOptions.Compiled);

        private static readonly Dictionary<string, RasData.CrDateConditionEnum> DatePeriods =
            new Dictionary<string, RasData.CrDateConditionEnum>(StringComparer.OrdinalIgnoreCase)
            {
                { "daily", RasData.CrDateConditionEnum.crDateConditionDaily },
                { "weekly", RasData.CrDateConditionEnum.crDateConditionWeekly },
                { "monthly", RasData.CrDateConditionEnum.crDateConditionMonthly },
                { "quarterly", RasData.CrDateConditionEnum.crDateConditionQuarterly },
                { "semiannually", RasData.CrDateConditionEnum.crDateConditionSemiannually },
                { "annually", RasData.CrDateConditionEnum.crDateConditionAnnually },
                { "by hour", RasData.CrDateConditionEnum.crDateConditionHour }
            };
        #endregion

        #region Methods.Data
        private static bool IsSqlServerProvider(string provider)
        {
            return IsProvider(provider, SqlServerProviders);
        }

        /// <summary>
        /// The sources of the tables which Crystal leaves out of its query: a table whose fields the report doesn't
        /// use and which links no used table to the others, so the parents and the children of such tables too.
        /// G1082 links Customer, Orders and Orders_Detail and prints the 2191 orders; the import took Orders_Detail,
        /// the last child, as the main source and printed 3684 rows of details. G1447 reads only a Command, its four
        /// linked tables without fields are left out. The synthetic data of the checks has one row in every table but
        /// the main one, so only the databases showed it. A report which uses no field keeps all the tables.
        /// A table between two tables of the query stays: film_actor links film and actor in the Polish samples
        /// of sakila (G2664, G2637), and without it actor became the main source and printed one row of twelve.
        /// </summary>
        private HashSet<StiDataSource> GetUnusedSources(ReportDocument report, List<StiDataRelation> relations)
        {
            var unused = new HashSet<StiDataSource>();
            var used = new HashSet<StiDataSource>();
            var sources = new List<StiDataSource>();
            foreach (Table table in report.Database.Tables)
            {
                var source = htTableNameToDataSource[table.Name] as StiDataSource;
                if (source == null)
                    continue;

                sources.Add(source);
                if (table.Fields.Cast<DatabaseFieldDefinition>().Any(f => f.UseCount > 0))
                    used.Add(source);
            }

            if (used.Count == 0)
                return unused;

            var changed = true;
            while (changed)
            {
                changed = false;
                foreach (var source in sources)
                {
                    if (used.Contains(source) || unused.Contains(source))
                        continue;

                    // A parent of a table which stays in the query links it to the others and stays too
                    if (relations.Any(r => r.ParentSource == source && !unused.Contains(r.ChildSource)))
                        continue;

                    // A child of two tables which stay, like a table of a many-to-many link, joins them; a parent
                    // linked only to this table doesn't count, it would stay just because of it
                    var linked = relations
                        .Where(r => r.ChildSource == source && !unused.Contains(r.ParentSource))
                        .Select(r => r.ParentSource).Distinct()
                        .Count(p => used.Contains(p) || HasOtherLinks(p, source, relations, unused));
                    if (linked > 1)
                        continue;

                    unused.Add(source);
                    changed = true;
                }
            }

            return unused;
        }

        /// <summary>
        /// Whether the source links a table which stays, other than the given one.
        /// </summary>
        private static bool HasOtherLinks(StiDataSource source, StiDataSource other, List<StiDataRelation> relations,
            HashSet<StiDataSource> unused)
        {
            return relations.Any(r =>
                (r.ParentSource == source && r.ChildSource != other && !unused.Contains(r.ChildSource)) ||
                (r.ChildSource == source && r.ParentSource != other && !unused.Contains(r.ParentSource)));
        }

        private static bool IsProvider(string provider, string[] providers)
        {
            return provider != null && providers.Any(p => provider.StartsWith(p, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// The type of a connection as QE_DatabaseType names it. The native drivers of Oracle, DAO, xBase and XML
        /// leave it empty and are known by their library, Oracle, Access and ODBC through OLE DB by the provider.
        /// Such tables could not be identified before: Access through DAO in 182 reports, Oracle in 97, xBase in 25.
        /// </summary>
        private static string GetConnectionType(string library, string type, string provider)
        {
            if (string.Equals(library, "crdb_dao.dll", StringComparison.OrdinalIgnoreCase))
                return ConnectionAccess;

            if (string.Equals(library, "crdb_oracle.dll", StringComparison.OrdinalIgnoreCase))
                return ConnectionOracle;

            if (string.Equals(library, "crdb_p2bxbse.dll", StringComparison.OrdinalIgnoreCase))
                return ConnectionXBase;

            if (string.Equals(library, "crdb_xml.dll", StringComparison.OrdinalIgnoreCase))
                return ConnectionXml;

            // The provider of OLE DB for ODBC: its data source is the DSN
            if (type == "OLE DB (ADO)" && IsProvider(provider, new[] { "MSDASQL" }))
                return ConnectionOdbc;

            if (type == "OLE DB (ADO)" && IsProvider(provider, OracleProviders))
                return ConnectionOracle;

            if (type == "OLE DB (ADO)" && IsProvider(provider, AccessProviders))
                return ConnectionAccess;

            return type;
        }

        /// <summary>
        /// The name of the data which the application gives at run time, null for a table of a database: a DataSet
        /// of ADO.NET without a file is registered by its name, a file of field definitions (TTX) describes the
        /// columns of a table only. Such a table has no database, the report reads the registered data.
        /// </summary>
        private static string GetApplicationDataName(Table table, string library, string type, string server,
            Hashtable logon)
        {
            if (string.Equals(library, "crdb_fielddef.dll", StringComparison.OrdinalIgnoreCase))
                return table.Name;

            if ((type == "ADO.NET" || type == "ADO.NET (XML)") && string.IsNullOrEmpty(logon["File Path"] as string))
                return $"{System.IO.Path.GetFileName(server).Replace('.', '_')}.{table.Name}";

            return null;
        }

        private static string GetOracleConnectionString(string server, string user)
        {
            return $"Data Source={server};User Id={user};Password=;";
        }

        private static string GetAccessConnectionString(string provider, string path, string user)
        {
            if (!IsProvider(provider, AccessProviders))
                provider = AccessProvider;

            var connectionString = $"Provider={provider};Data Source={path};";
            if (!string.IsNullOrEmpty(user))
                connectionString += $"User ID={user};Password=;";

            return connectionString;
        }

        /// <summary>
        /// The connection string of a connection without a DSN, which Crystal keeps whole (SAP Business One:
        /// DRIVER={B1CRHPROXY32};DATABASE=...;SERVERNODE=...), without the password. The logon keeps the current
        /// values of its keys apart, the string may hold the old ones (B0001: DATABASE=SBODEMOUS in the string,
        /// Database=SBODEMODK in the logon), so a key of the logon replaces the key of the same name.
        /// </summary>
        private static string GetOdbcDriverConnectionString(string connectionString, Hashtable logon)
        {
            var parts = new List<string>();
            foreach (var part in connectionString.Split(';'))
            {
                var separator = part.IndexOf('=');
                var key = separator > 0 ? part.Substring(0, separator).Trim() : null;
                if (key == null)
                {
                    parts.Add(part);
                    continue;
                }

                if (key.Equals("PWD", StringComparison.OrdinalIgnoreCase) ||
                    key.Equals("Password", StringComparison.OrdinalIgnoreCase))
                    continue;

                var value = GetLogonValue(logon, key);
                parts.Add(value != null ? $"{key}={value}" : part);
            }

            return string.Join(";", parts);
        }

        /// <summary>
        /// The value of a key of the logon by its name in any case, except the connection string itself.
        /// </summary>
        private static string GetLogonValue(Hashtable logon, string key)
        {
            foreach (DictionaryEntry entry in logon)
            {
                var name = entry.Key as string;
                if (name != "Connection String" && string.Equals(name, key, StringComparison.OrdinalIgnoreCase))
                    return entry.Value as string;
            }

            return null;
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
            var driverConnectionString = logon["Connection String"] as string;
            if (string.IsNullOrEmpty(logon["DSN"] as string) && !string.IsNullOrEmpty(driverConnectionString))
                return GetOdbcDriverConnectionString(driverConnectionString, logon);

            var dsn = logon["DSN"] as string;
            if (string.IsNullOrEmpty(dsn))
                dsn = server;

            var connectionString = $"Dsn={dsn}";
            var database = logon["Database"] as string ?? logon["Initial Catalog"] as string;
            if (!string.IsNullOrEmpty(database))
                connectionString += $";Database={database}";

            if (!string.IsNullOrEmpty(user))
                connectionString += $";Uid={user};Pwd=";

            return connectionString;
        }

        /// <summary>
        /// What a table reads: the SQL of a Command or the table of the database. The tables of the report and of
        /// its subreports with one name and one query are one data source.
        /// </summary>
        private static string GetTableQuery(Table table)
        {
            var connection = table.LogOnInfo.ConnectionInfo;
            var command = GetRasObject(table, "RasTable") as RasData.CommandTable;
            var query = command != null ? command.CommandText : table.Location;

            return $"{connection.ServerName}|{connection.DatabaseName}|{query}";
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

            source.SqlCommand = source is StiOdbcSource || source is StiOleDbSource || source is StiOracleSource
                ? $"select * from \"{name}\""
                : $"select * from [{name}]";
        }

        /// <summary>
        /// A parameter of a Command, {?Name}, becomes the variable of the parameter, which Stimulsoft inserts into
        /// the text of the SQL as Crystal does. Any other brace is a character of the SQL, like the escape of ODBC
        /// {oj ...} (G0146): Stimulsoft would take it for an expression, so it is printed by an expression.
        /// </summary>
        private string ReplaceCommandParameter(Match match)
        {
            if (match.Groups[1].Success)
            {
                var variable = FindParameterVariable(match.Groups[1].Value);
                if (variable != null)
                    return "{" + variable + "}";

                warningList.Add(StiWarning.GetWarningFormula("SQL", $"Unknown parameter {match.Value}"));
            }

            return EscapeBraces(match.Value);
        }

        /// <summary>
        /// Finds the variable of a parameter by its name; the names of Crystal are case insensitive, a Command
        /// may write {?@lcWoNo} for the parameter @lcWono and {?SampleID} for the parameter @SampleID (G0589).
        /// </summary>
        private string FindParameterVariable(string parameter)
        {
            var key = "?" + parameter;
            foreach (DictionaryEntry entry in htVariableConversion)
            {
                if (string.Equals((string)entry.Key, key, StringComparison.OrdinalIgnoreCase))
                    return (string)entry.Value;
            }

            var name = parameter.TrimStart('@');
            foreach (DictionaryEntry entry in htVariableConversion)
            {
                var entryKey = (string)entry.Key;
                if (entryKey.StartsWith("?") &&
                    string.Equals(entryKey.Substring(1).TrimStart('@'), name, StringComparison.OrdinalIgnoreCase))
                {
                    return (string)entry.Value;
                }
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
        /// The order of a group and the period of a group on a date: the sort of the group gives the direction
        /// (the original order of Crystal is no sorting), the date condition of the group turns the date into
        /// the start of its period, so all the dates of a month, a week and so on make one group.
        /// </summary>
        private void ApplyGroupOptions(StiGroupHeaderBand band, ReportDocument report, int index)
        {
            if (index >= report.DataDefinition.Groups.Count)
                return;

            var group = report.DataDefinition.Groups[index];
            band.SortDirection = GetGroupSortDirection(report, group);

            var groupOptions = (GetRasObject(group, "RasGroup") as RasData.ISCRGroup)?.Options;
            var options = groupOptions as RasData.DateGroupOptions;
            var condition = band.Condition.Value;
            if (condition == null || !condition.StartsWith("{") || !condition.EndsWith("}"))
                return;

            var specified = groupOptions as RasData.SpecifiedGroupOptions;
            if (specified != null)
            {
                ApplySpecifiedOrder(band, specified, condition.Substring(1, condition.Length - 2),
                    group.ConditionField?.FormulaName?.Trim('{', '}'));
                return;
            }

            if (options == null)
                return;

            var date = condition.Substring(1, condition.Length - 2);
            var period = GetDatePeriod(date, options.DateCondition);
            if (period == null)
            {
                warningList.Add(StiWarning.GetWarningFormula(band.Name, $"The date group {options.DateCondition} " +
                    "is not supported"));
                return;
            }

            band.Condition.Value = "{" + period + "}";
        }

        /// <summary>
        /// A group in a specified order: every named group has a condition, a formula of Crystal, and the other
        /// values are merged into one group, kept in their own groups or discarded. The condition of the band numbers
        /// the named groups, so they keep their order and the others go last; GroupName prints the names. G0678,
        /// G2591 and others printed the values of the field instead of "Less than 30 days", "0..7" and "Others".
        /// </summary>
        private void ApplySpecifiedOrder(StiGroupHeaderBand band, RasData.SpecifiedGroupOptions options,
            string value, string field)
        {
            var conditions = new List<string>();
            var names = new List<string>();
            foreach (RasData.ISCRFilter filter in options.SpecifiedValueFilters)
            {
                var expression = TranslateFormula(filter.FreeEditingText, band.Name);
                if (expression == null)
                    return;

                if (expression.Type != typeof(bool))
                {
                    warningList.Add(StiWarning.GetWarningFormula(band.Name,
                        $"The condition of the group '{filter.Name}' is not boolean"));
                    return;
                }

                conditions.Add(expression.Code);
                names.Add(filter.Name);
            }

            var unspecified = options.UnspecifiedValuesType;
            var others = unspecified == RasData.CrUnspecifiedValuesTypeEnum.crUnspecifiedValuesTypeSeparateValues;
            if (unspecified == RasData.CrUnspecifiedValuesTypeEnum.crUnspecifiedValuesTypeDiscardValues)
            {
                warningList.Add(StiWarning.GetWarningFormula(band.Name,
                    "The values out of the specified groups are not discarded, they make one group"));
            }

            // The interpretation mode doesn't concatenate a string with a number, ToString works in both modes
            var key = others ? $"\"999\" + ({value}).ToString()" : "\"999\"";
            var name = others ? $"({value}).ToString()" : ToStringLiteral(options.UnspecifiedValuesName ?? string.Empty);
            for (var index = conditions.Count - 1; index >= 0; index--)
            {
                var number = index.ToString("D3", CultureInfo.InvariantCulture);
                // The parser of the interpretation needs the parentheses around a nested conditional
                key = $"({conditions[index]}) ? \"{number}\" : ({key})";
                name = $"({conditions[index]}) ? {ToStringLiteral(names[index])} : ({name})";
            }

            band.Condition.Value = "{" + key + "}";
            band.SortDirection = StiGroupSortDirection.Ascending;

            if (field != null)
                specifiedGroupNames[field] = name;
        }

        private static string ToStringLiteral(string text)
        {
            return "\"" + text.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
        }

        private static StiGroupSortDirection GetGroupSortDirection(ReportDocument report,
            CrystalDecisions.CrystalReports.Engine.Group group)
        {
            var field = group.ConditionField?.FormulaName;
            foreach (SortField sort in report.DataDefinition.SortFields)
            {
                if (sort.SortType != SortFieldType.GroupSortField || sort.Field?.FormulaName != field)
                    continue;

                switch (sort.SortDirection)
                {
                    case SortDirection.DescendingOrder:
                        return StiGroupSortDirection.Descending;

                    case SortDirection.AscendingOrder:
                        return StiGroupSortDirection.Ascending;

                    default:
                        return StiGroupSortDirection.None;
                }
            }

            return StiGroupSortDirection.Ascending;
        }

        /// <summary>
        /// The start of the period of a date: a week starts on Sunday as in Crystal, it is counted from the Sunday
        /// of 7 January 1900 without the enumeration DayOfWeek, which the interpretation mode can't cast.
        /// </summary>
        private static string GetDatePeriod(string date, RasData.CrDateConditionEnum condition)
        {
            switch (condition)
            {
                case RasData.CrDateConditionEnum.crDateConditionDaily:
                    return $"{date}.Date";

                case RasData.CrDateConditionEnum.crDateConditionWeekly:
                    return $"{date}.Date.AddDays(-(({date}.Date - DateSerial(1900, 1, 7)).Days % 7))";

                case RasData.CrDateConditionEnum.crDateConditionMonthly:
                    return $"DateSerial({date}.Year, {date}.Month, 1)";

                case RasData.CrDateConditionEnum.crDateConditionQuarterly:
                    return $"DateSerial({date}.Year, ({date}.Month - 1) / 3 * 3 + 1, 1)";

                case RasData.CrDateConditionEnum.crDateConditionSemiannually:
                    return $"DateSerial({date}.Year, ({date}.Month - 1) / 6 * 6 + 1, 1)";

                case RasData.CrDateConditionEnum.crDateConditionAnnually:
                    return $"DateSerial({date}.Year, 1, 1)";

                case RasData.CrDateConditionEnum.crDateConditionHour:
                    return $"{date}.Date.AddHours({date}.Hour)";

                default:
                    return null;
            }
        }

        /// <summary>
        /// The date condition of a group name like GroupName ({Orders.Date}, "monthly"), null for a group name
        /// without a period.
        /// </summary>
        private static RasData.CrDateConditionEnum? GetGroupNameCondition(string groupName)
        {
            var match = Regex.Match(groupName ?? string.Empty, @",\s*""([^""]+)""\s*\)\s*$");
            if (!match.Success)
                return null;

            RasData.CrDateConditionEnum condition;
            if (!DatePeriods.TryGetValue(match.Groups[1].Value, out condition))
                return null;

            return condition;
        }

        /// <summary>
        /// The period of a group by a date as the formulas of Crystal name it, like "monthly"; null for a group
        /// without a period or with a period which the import doesn't support.
        /// </summary>
        private string GetGroupPeriod(CrystalDecisions.CrystalReports.Engine.Group group)
        {
            var options = (GetRasObject(group, "RasGroup") as RasData.ISCRGroup)?.Options as RasData.DateGroupOptions;
            if (options == null)
                return null;

            return DatePeriods.FirstOrDefault(p => p.Value == options.DateCondition).Key;
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
