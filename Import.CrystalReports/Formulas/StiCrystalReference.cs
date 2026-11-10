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
    /// What a reference of a formula, like {Table.Field}, {@Formula} or {?Parameter}, means in the Stimulsoft report.
    /// </summary>
    public class StiCrystalReference
    {
        #region Properties
        /// <summary>
        /// The content of the reference of the formula without the braces, for example Orders.Amount.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// The expression of the value in the Stimulsoft report, for example Orders.Customers.Name.
        /// </summary>
        public string Expression { get; set; }

        public Type Type { get; set; }

        /// <summary>
        /// The name of the data source and of the column of a field of a database; they are used by the functions
        /// which check the value of a column, like IsNull, Previous and Next. Empty for other references.
        /// </summary>
        public string DataSourceName { get; set; }

        public string ColumnName { get; set; }

        /// <summary>
        /// True for a parameter with a range of values: the bounds are the properties From and To of the value.
        /// </summary>
        public bool IsRange { get; set; }

        /// <summary>
        /// True for a parameter with several values: the value is a list with the method Contains.
        /// </summary>
        public bool IsMultiValue { get; set; }

        /// <summary>
        /// True if the expression contains a summary, like Sum or Count; the formulas which use it get this flag too.
        /// </summary>
        public bool HasSummaries { get; set; }

        /// <summary>
        /// False if the decimal value of a formula is a number of Crystal, not a currency; see
        /// StiCrystalExpression.IsCurrency.
        /// </summary>
        public bool? IsCurrency { get; set; }

        /// <summary>
        /// The fields of a database which a formula reads in any case: Crystal gives NULL for the formula when one
        /// of them is NULL. Null for the other references.
        /// </summary>
        public List<StiCrystalReference> NullFields { get; set; }
        #endregion

        #region Methods
        /// <summary>
        /// The check of NULL of the column of a field. IsNull of a data source reads the current row of the source;
        /// a column of a parent table is read through the relations, the source itself doesn't move with the band,
        /// so such a column is checked by its path: StiNullValuesHelper.IsNull(this, "Orders.Customer.Name"). G2899
        /// filtered by IsNull(Customer, "CustomerName") and printed nothing in the interpretation once Customer
        /// became the parent of Orders.
        /// </summary>
        public string GetIsNullCode()
        {
            if (string.IsNullOrEmpty(Expression) || Expression == $"{DataSourceName}.{ColumnName}")
                return $"IsNull({DataSourceName}, \"{ColumnName}\")";

            return $"StiNullValuesHelper.IsNull(this, \"{Expression}\")";
        }
        #endregion

        public StiCrystalReference(string expression, Type type)
        {
            Expression = expression;
            Type = type;
        }
    }
}
