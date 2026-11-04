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

namespace Import.CrystalReports
{
    /// <summary>
    /// A running total of Crystal translated to a running aggregate function of Stimulsoft, like SumIfRunning.
    /// </summary>
    public class StiCrystalRunningTotal
    {
        #region Properties
        /// <summary>
        /// The aggregate function without the suffix, for example Sum or CountDistinct.
        /// </summary>
        public string Function { get; }

        /// <summary>
        /// The band which the function goes over: the band of the group which resets the total, or the data band.
        /// </summary>
        public string Band { get; }

        /// <summary>
        /// The summarized value; Count takes no value.
        /// </summary>
        public string Value { get; }

        /// <summary>
        /// The condition of the evaluation, null if the total takes every row.
        /// </summary>
        public string Condition { get; }

        public Type Type { get; }

        /// <summary>
        /// The expression of the running total; the running functions of the interpretation return object,
        /// so the value is cast to its type.
        /// </summary>
        public string Code
        {
            get
            {
                var arguments = new List<string> { Band };

                // Count of Stimulsoft counts the rows, it takes only the band and the condition
                if (Function != "Count")
                    arguments.Add(Value);

                if (Condition != null)
                    arguments.Add(Condition);

                var suffix = Condition != null ? "IfRunning" : "Running";
                return $"(({TypeName}){Function}{suffix}({string.Join(", ", arguments)}))";
            }
        }

        /// <summary>
        /// True if the total can be shown in a group header by GetHeaderCode.
        /// </summary>
        public bool CanBeShownInHeader => Function == "Sum" || Function == "Count";

        private string TypeName => Type == typeof(int) ? "int" : Type == typeof(decimal) ? "decimal"
            : Type == typeof(string) ? "string" : "DateTime";
        #endregion

        #region Methods
        /// <summary>
        /// The expression of the total in a group header. Crystal shows there the value at the first row of the
        /// group, Stimulsoft calculates a running function in a header up to the end of the group. So the total of
        /// the group is subtracted, and the first row, which the header reads, is added again.
        /// </summary>
        public string GetHeaderCode(string header)
        {
            var conditionSuffix = Condition != null ? "If" : "";
            var condition = Condition != null ? $", {Condition}" : "";

            if (Function == "Count")
            {
                var row = Condition != null ? $"(({Condition}) ? 1 : 0)" : "1";
                return $"((int)({Code} - ((int)Count{conditionSuffix}({header}{condition})) + {row}))";
            }

            var value = $"((decimal)({Value}))";
            var rowValue = Condition != null ? $"(({Condition}) ? {value} : ((decimal)0))" : value;
            var total = $"((decimal)Sum{conditionSuffix}({header}, {Value}{condition}))";
            return $"((decimal)({Code} - {total} + {rowValue}))";
        }
        #endregion

        public StiCrystalRunningTotal(string function, string band, string value, string condition, Type type)
        {
            Function = function;
            Band = band;
            Value = value;
            Condition = condition;
            Type = type;
        }
    }
}
