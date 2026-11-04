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

namespace Import.CrystalReports
{
    /// <summary>
    /// The parts of a .NET date pattern built from the date format of Crystal.
    /// </summary>
    public static class StiCrystalDatePattern
    {
        #region Methods
        /// <summary>
        /// Splits a pattern into the parts before the long month, the long month and the parts after it; null if
        /// the pattern has no long month outside the quoted separators. .NET writes a long month in the genitive
        /// case when the pattern has a day, so the month formatted alone takes the nominative case, as in Crystal.
        /// </summary>
        public static List<string> SplitAtLongMonth(string pattern)
        {
            var parts = new List<string>();
            var start = 0;
            var inQuote = false;
            var index = 0;

            while (index < pattern.Length)
            {
                var c = pattern[index];
                if (c == '\\')
                {
                    index += 2;
                    continue;
                }

                if (c == '\'')
                    inQuote = !inQuote;

                if (!inQuote && string.CompareOrdinal(pattern, index, "MMMM", 0, 4) == 0)
                {
                    if (index > start)
                        parts.Add(pattern.Substring(start, index - start));

                    parts.Add("MMMM");
                    index += 4;
                    start = index;
                    continue;
                }

                index++;
            }

            if (parts.Count == 0)
                return null;

            if (start < pattern.Length)
                parts.Add(pattern.Substring(start));

            return parts;
        }
        #endregion
    }
}
