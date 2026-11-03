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
using System.Linq;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;

namespace Import.CrystalReports
{
    /// <summary>
    /// The names of the report: the data sources, the variables, the pages and the components become the members
    /// of one class of the compiled report, so they must be identifiers of C# and differ from each other and from
    /// the members which the expressions call without a prefix (Totals, Format, Sum and others). Crystal keeps
    /// these names apart, so a formula Title1 and a text object Title1 or a table and a subreport of one name are
    /// usual in the reports. The names are compared ignoring the case: the dictionary of Stimulsoft finds them so.
    /// </summary>
    public class StiCrystalNameTable
    {
        #region Fields
        private static readonly string[] ReservedNames = CreateReservedNames();

        private readonly HashSet<string> names = new HashSet<string>(ReservedNames, StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> parameterNames =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Methods
        private static string[] CreateReservedNames()
        {
            var reserved = new List<string>
            {
                "Report", "Dictionary", "Totals", "Page", "Pages", "Line", "Column", "Today", "Time", "PageNumber",
                "TotalPageCount", "PageNofM", "IsFirstPage", "IsLastPage", "Name", "Variables", "DataSources",
                "Format", "IIF", "IsNull", "ToString", "CheckExcelValue", "Abs", "DateSerial", "Left", "Right", "Mid",
                "Round", "Trim", "ParseDouble", "StrToDateTime", "ToProperCase"
            };

            foreach (var function in new[] { "Sum", "Avg", "Min", "Max", "Count", "CountDistinct", "Median", "Mode",
                "First", "Last", "StDev", "StDevP" })
            {
                reserved.Add(function);
                reserved.Add(function + "If");
                reserved.Add(function + "Running");
                reserved.Add(function + "IfRunning");
            }

            return reserved.ToArray();
        }

        /// <summary>
        /// Takes a name which is already given, like the name of the first page of a new report.
        /// </summary>
        public void Add(string name)
        {
            names.Add(name);
        }

        /// <summary>
        /// Reserves the name of the data source of a table before the other members: the expressions show the tables
        /// most often, so the other members which meet such a name get a suffix.
        /// </summary>
        public void ReserveTable(string tableName)
        {
            names.Add(GetDataSourceName(tableName));
        }

        /// <summary>
        /// The name of the data source of a table; the same table of the report and of a subreport is one source.
        /// </summary>
        public static string GetDataSourceName(string tableName)
        {
            var name = StiNameValidator.CorrectName(tableName);
            if (ReservedNames.Contains(name, StringComparer.OrdinalIgnoreCase))
                name += "Source";

            return name;
        }

        /// <summary>
        /// A new member of the report class: an identifier of C# which no other member has.
        /// </summary>
        public string Create(string name)
        {
            var baseName = StiNameValidator.CorrectName(name);
            if (string.IsNullOrEmpty(baseName))
                baseName = "Item";

            var memberName = baseName;
            var index = 1;
            while (names.Contains(memberName))
            {
                memberName = $"{baseName}_{index}";
                index++;
            }

            names.Add(memberName);
            return memberName;
        }

        /// <summary>
        /// The variable of a parameter; a parameter of a subreport with the name of a parameter of the report is the
        /// same variable, as it was before: the dictionary keeps one variable of a name.
        /// </summary>
        public string GetParameterName(string parameterName)
        {
            string name;
            if (!parameterNames.TryGetValue(parameterName, out name))
            {
                name = Create(parameterName);
                parameterNames[parameterName] = name;
            }

            return name;
        }

        /// <summary>
        /// The name of a column of a data source: an identifier which differs from the other columns. The spaces are
        /// removed with the next letter capitalized, as the import always did, the validator corrects the rest:
        /// a digit at the beginning, a dot and the other symbols.
        /// </summary>
        public static string CreateColumnName(StiDataSource dataSource, string fieldName)
        {
            var baseName = StiNameValidator.CorrectName(RemoveSpaces(fieldName));
            if (string.IsNullOrEmpty(baseName))
                baseName = "Column";

            var columnName = baseName;
            var index = 1;
            while (HasColumn(dataSource, columnName))
            {
                columnName = $"{baseName}_{index}";
                index++;
            }

            return columnName;
        }

        /// <summary>
        /// The name of the column of a field of a table, as CreateColumnName gave it.
        /// </summary>
        public static string GetColumnName(StiDataSource dataSource, string fieldName)
        {
            var column = dataSource.Columns.Cast<StiDataColumn>().FirstOrDefault(c => c.NameInSource == fieldName);
            return column != null ? column.Name : CreateColumnName(dataSource, fieldName);
        }

        /// <summary>
        /// The name of a relation is a property of the class of the child source, beside its columns: a column of the
        /// name of the parent table, like Company in a table linked to Company, would give two members of one name.
        /// Another relation of the same name must link the same sources, otherwise the classes of the relations
        /// clash; the relations of a subreport repeat the relations of the report and keep their names.
        /// The suffix is Link, not Relation: the code of the report names the field of a relation X ParentX and its
        /// class ParentXRelation, so a relation XRelation would give the field ParentXRelation beside that class.
        /// </summary>
        public static string CreateRelationName(StiDictionary dictionary, StiDataRelation relation,
            string parentTableName)
        {
            var baseName = StiNameValidator.CorrectName(parentTableName);
            var relationName = baseName;
            var index = 1;
            while (IsRelationNameTaken(dictionary, relation, relationName))
            {
                relationName = $"{baseName}Link{(index > 1 ? index.ToString() : string.Empty)}";
                index++;
            }

            return relationName;
        }

        private static bool IsRelationNameTaken(StiDictionary dictionary, StiDataRelation relation, string name)
        {
            var child = relation.ChildSource;
            if (child != null && HasColumn(child, name))
                return true;

            return dictionary.Relations.Cast<StiDataRelation>().Any(r =>
                string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase) &&
                (r.ChildSource?.Name != child?.Name || r.ParentSource?.Name != relation.ParentSource?.Name));
        }

        /// <summary>
        /// A table of a subreport which reads the same table of the database as the report is one source with the
        /// fields of both: the copy of the subreport takes the names of the columns of the source, and its other
        /// fields are added to the source.
        /// </summary>
        public static void MergeColumns(StiDataSource source, StiDataSource copy)
        {
            foreach (StiDataColumn column in copy.Columns)
            {
                var sourceColumn = source.Columns.Cast<StiDataColumn>()
                    .FirstOrDefault(c => c.NameInSource == column.NameInSource);

                if (sourceColumn == null)
                {
                    sourceColumn = new StiDataColumn(column.NameInSource, CreateColumnName(source, column.NameInSource),
                        column.Alias, column.Type);
                    source.Columns.Add(sourceColumn);
                }

                // The setter of Name renames NameInSource and Alias too when they are equal to the old name
                var nameInSource = column.NameInSource;
                var alias = column.Alias;
                column.Name = sourceColumn.Name;
                column.NameInSource = nameInSource;
                column.Alias = alias;
            }
        }

        private static bool HasColumn(StiDataSource dataSource, string name)
        {
            return dataSource.Columns.Cast<StiDataColumn>()
                .Any(c => string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Removes the spaces and capitalizes the letter after each of them: Order Date becomes OrderDate.
        /// </summary>
        public static string RemoveSpaces(string name)
        {
            var chars = new List<char>(name.Length);
            var upper = false;
            foreach (var ch in name)
            {
                if (ch == ' ')
                {
                    upper = true;
                    continue;
                }

                chars.Add(upper ? char.ToUpper(ch) : ch);
                upper = false;
            }

            return new string(chars.ToArray());
        }
        #endregion
    }
}
