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

using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Import.CrystalReports
{
    /// <summary>
    /// The direction of a link of Crystal. A link goes from a source table to a destination table and both
    /// directions occur: Orders_Detail -> Orders -> Customer in G1082, Customer -> Orders in G2899 and the other
    /// samples of the Xtreme database. A relation of Stimulsoft needs the parent, the table of "one", and the import
    /// always took the destination: in G2899 Customer became the main source and every customer showed one order of
    /// its 2192. The report keeps no keys of the tables, so the names decide: the side whose fields look like the key
    /// of their own table (Customer.Customer ID, film.film_id, Account.Account Number) is the parent. On two samples
    /// (539 links) 76 links were turned, all from the table of "one"; no link had a key on both sides.
    /// </summary>
    public static class StiCrystalTableLinks
    {
        #region Consts
        private static readonly string[] KeySuffixes = { "id", "no", "nr", "num", "number", "code", "key" };
        #endregion

        #region Methods
        /// <summary>
        /// True if the source table of the link is the parent: its fields look like its key and the fields of the
        /// destination table don't. Otherwise the destination is the parent, as the import took it before.
        /// </summary>
        public static bool IsSourceParent(string sourceTable, IList<string> sourceFields, string destinationTable,
            IList<string> destinationFields)
        {
            return IsKeyOfTable(sourceTable, sourceFields) && !IsKeyOfTable(destinationTable, destinationFields);
        }

        /// <summary>
        /// True if the only field of a link is named as the key of its table: the name of the table, also without
        /// the ending of the plural, and a suffix like ID or Number, or just ID. Spaces, underscores and the case
        /// don't matter.
        /// </summary>
        public static bool IsKeyOfTable(string table, IList<string> fields)
        {
            if (string.IsNullOrEmpty(table) || fields == null || fields.Count != 1)
                return false;

            var field = Normalize(fields[0]);
            if (field == "id")
                return true;

            var name = Normalize(table.Substring(table.LastIndexOf('.') + 1));
            return GetSingularNames(name).Any(n => KeySuffixes.Any(s => field == n + s));
        }

        private static IEnumerable<string> GetSingularNames(string name)
        {
            yield return name;

            if (name.EndsWith("ies") && name.Length > 5)
                yield return name.Substring(0, name.Length - 3) + "y";

            if (name.EndsWith("es") && name.Length > 4)
                yield return name.Substring(0, name.Length - 2);

            if (name.EndsWith("s") && name.Length > 3)
                yield return name.Substring(0, name.Length - 1);
        }

        private static string Normalize(string text)
        {
            var builder = new StringBuilder();
            foreach (var c in text ?? string.Empty)
            {
                if (char.IsLetterOrDigit(c))
                    builder.Append(char.ToLowerInvariant(c));
            }

            return builder.ToString();
        }
        #endregion
    }
}
