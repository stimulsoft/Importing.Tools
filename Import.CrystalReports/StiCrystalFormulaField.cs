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
using Stimulsoft.Report.Dictionary;

namespace Import.CrystalReports
{
    /// <summary>
    /// A formula field of the Crystal report and the variable which it becomes; the formula is translated once,
    /// when it or a formula using it is translated first.
    /// </summary>
    public class StiCrystalFormulaField
    {
        #region Properties
        /// <summary>
        /// The name of the formula as it is written in a reference, for example @Total.
        /// </summary>
        public string Name { get; }

        public string Text { get; }

        public StiVariable Variable { get; }

        /// <summary>
        /// The type of the value of the formula in Crystal.
        /// </summary>
        public Type DeclaredType { get; }

        public bool IsTranslating { get; set; }

        public bool IsTranslated { get; set; }

        /// <summary>
        /// The translated expression when the formula has summaries; such a formula is placed into the text of
        /// a component instead of the reference to the variable. Null for other formulas.
        /// </summary>
        public string InlineExpression { get; set; }

        /// <summary>
        /// True if Crystal calculates the formula with the default values instead of NULL, by the option of the
        /// formula or of the report; otherwise a NULL field which the formula reads gives NULL for the formula.
        /// </summary>
        public bool NullsAsDefault { get; set; }
        #endregion

        public StiCrystalFormulaField(string name, string text, StiVariable variable, Type declaredType)
        {
            Name = name;
            Text = text;
            Variable = variable;
            DeclaredType = declaredType;
        }
    }
}
