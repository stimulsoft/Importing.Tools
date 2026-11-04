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

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// Gives the translator of formulas the names of the report: the fields, the formulas, the parameters,
    /// the groups and the value of the formatted object.
    /// </summary>
    public interface IStiCrystalResolver
    {
        /// <summary>
        /// Resolves the content of a reference without the braces, for example Orders.Amount, @Total or ?Date.
        /// Returns null if the reference is unknown.
        /// </summary>
        StiCrystalReference ResolveReference(string name);

        /// <summary>
        /// Resolves CurrentFieldValue of a formula of conditional formatting. Returns null outside such a formula.
        /// </summary>
        StiCrystalReference ResolveCurrentFieldValue();

        /// <summary>
        /// Returns the name of the band of the group whose condition is the specified reference, for the summaries
        /// like Sum({Orders.Amount}, {Orders.Customer}). The period, like "monthly", picks one of the groups by
        /// the same date; null takes the group by the field. Returns null if the group is unknown.
        /// </summary>
        string ResolveGroupBand(string name, string period);

        /// <summary>
        /// Returns the variable of the dictionary for a global or shared variable of a formula; the same name gives
        /// the same variable in all the formulas of the report. Returns null if such variables are not supported.
        /// </summary>
        StiCrystalReference ResolveVariable(string name, System.Type type);
    }
}
