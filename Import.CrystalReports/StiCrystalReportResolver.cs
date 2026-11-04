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
using Import.CrystalReports.Formulas;

namespace Import.CrystalReports
{
    /// <summary>
    /// Keeps the names of the converted report for the translator of formulas: the fields of the databases,
    /// the formulas, the parameters, the running totals and the groups. The names of Crystal are case insensitive.
    /// </summary>
    public class StiCrystalReportResolver : IStiCrystalResolver
    {
        #region IStiCrystalResolver
        public StiCrystalReference ResolveReference(string name)
        {
            ReferenceResolving?.Invoke(name);

            StiCrystalReference reference;
            return references.TryGetValue(name, out reference) ? reference : null;
        }

        public StiCrystalReference ResolveCurrentFieldValue()
        {
            return CurrentFieldValue;
        }

        public string ResolveGroupBand(string name, string period)
        {
            string band;
            if (period != null && groupBands.TryGetValue(GetGroupKey(name, period), out band))
                return band;

            return groupBands.TryGetValue(name, out band) ? band : null;
        }

        public StiCrystalReference ResolveVariable(string name, Type type)
        {
            return VariableResolving?.Invoke(name, type);
        }
        #endregion

        #region Fields
        private readonly Dictionary<string, StiCrystalReference> references =
            new Dictionary<string, StiCrystalReference>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> groupBands =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Properties
        /// <summary>
        /// The value of the formatted object for the formulas of conditional formatting; null for other formulas.
        /// </summary>
        public StiCrystalReference CurrentFieldValue { get; set; }

        /// <summary>
        /// Creates or finds the variable of the dictionary for a global or shared variable of a formula.
        /// </summary>
        public Func<string, Type, StiCrystalReference> VariableResolving { get; set; }
        #endregion

        #region Events
        /// <summary>
        /// Occurs before a reference is returned to the translator; the importer translates the referenced formula
        /// at this moment, so the reference already knows whether the formula has summaries.
        /// </summary>
        public event Action<string> ReferenceResolving;
        #endregion

        #region Methods
        /// <summary>
        /// Adds a reference by the content of the braces, for example Orders.Amount or @Total.
        /// </summary>
        public void AddReference(string name, StiCrystalReference reference)
        {
            reference.Name = name;
            references[name] = reference;
        }

        public StiCrystalReference GetReference(string name)
        {
            StiCrystalReference reference;
            return references.TryGetValue(name, out reference) ? reference : null;
        }

        public void RemoveReference(string name)
        {
            references.Remove(name);
        }

        /// <summary>
        /// Adds the band of a group. A report can group by one date several times with different periods: G2508
        /// groups by the quarter, the month and the week of SoldDate, and all its summaries took one band.
        /// </summary>
        public void AddGroupBand(string name, string band, string period)
        {
            groupBands[name] = band;

            if (period != null)
                groupBands[GetGroupKey(name, period)] = band;
        }

        private static string GetGroupKey(string name, string period)
        {
            return name + "|" + period.ToLowerInvariant();
        }
        #endregion
    }
}
