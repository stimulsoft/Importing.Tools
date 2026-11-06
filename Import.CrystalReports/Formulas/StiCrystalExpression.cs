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

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// A translated part of a formula: the C# code and its type. An array literal keeps its items and a range keeps
    /// its bounds instead of the code, because C# has no such values and they are expanded by the operator In.
    /// </summary>
    public class StiCrystalExpression
    {
        #region Properties
        public string Code { get; }

        public Type Type { get; }

        /// <summary>
        /// True if the code needs no parentheses when it is an operand: a literal, a name, a call, a member access.
        /// </summary>
        public bool IsAtomic { get; }

        /// <summary>
        /// The reference of the formula if the expression is just a reference, for the functions like IsNull.
        /// </summary>
        public StiCrystalReference Reference { get; set; }

        /// <summary>
        /// The text of a string literal, for the arguments which must be constant, like the interval of DateAdd.
        /// </summary>
        public string Literal { get; set; }

        /// <summary>
        /// False for a decimal which is a number of Crystal, not a currency: a summary of a number field or a ratio
        /// of two currencies. Null when it is not known, so a decimal stays a currency, as a currency field is.
        /// </summary>
        public bool? IsCurrency { get; set; }

        /// <summary>
        /// The name of the variable of the dictionary: for a global variable it is the name to assign,
        /// for an assignment it is the assigned variable.
        /// </summary>
        public string AssignedVariable { get; set; }

        /// <summary>
        /// True for an assignment of a global variable, Variable = value; it can be only the whole expression.
        /// </summary>
        public bool IsAssignment { get; set; }

        public List<StiCrystalExpression> Items { get; private set; }

        public StiCrystalExpression From { get; private set; }

        public StiCrystalExpression To { get; private set; }

        public bool IncludeFrom { get; private set; }

        public bool IncludeTo { get; private set; }

        public bool IsArray
        {
            get
            {
                return Items != null;
            }
        }

        public bool IsRange { get; private set; }
        #endregion

        #region Methods
        /// <summary>
        /// Returns the code in parentheses if the code is not atomic.
        /// </summary>
        public string ToOperand()
        {
            return IsAtomic ? Code : "(" + Code + ")";
        }

        public static StiCrystalExpression CreateArray(List<StiCrystalExpression> items)
        {
            return new StiCrystalExpression(null, typeof(object), isAtomic: true)
            {
                Items = items
            };
        }

        /// <summary>
        /// Creates a range; a missing bound is null, like in upfrom 5.
        /// </summary>
        public static StiCrystalExpression CreateRange(StiCrystalExpression from, bool includeFrom,
            StiCrystalExpression to, bool includeTo)
        {
            return new StiCrystalExpression(null, (from ?? to).Type, isAtomic: true)
            {
                IsRange = true,
                From = from,
                To = to,
                IncludeFrom = includeFrom,
                IncludeTo = includeTo
            };
        }

        public override string ToString()
        {
            return Code;
        }
        #endregion

        public StiCrystalExpression(string code, Type type, bool isAtomic = false)
        {
            Code = code;
            Type = type;
            IsAtomic = isAtomic;
        }
    }
}
