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

namespace Import.CrystalReports
{
    /// <summary>
    /// A value of a chart: the expression of the summarized field with its type, the function of the virtual source
    /// which summarizes it, the title of the series and the type of the result.
    /// </summary>
    internal sealed class StiCrystalChartValue
    {
        #region Properties
        public string Expression { get; }

        public Type Type { get; }

        public string Function { get; }

        public string Title { get; }

        public Type ResultType { get; }

        /// <summary>
        /// The column of the result in the virtual source.
        /// </summary>
        public string Column { get; set; }
        #endregion

        public StiCrystalChartValue(string expression, Type type, string function, string title, Type resultType)
        {
            Expression = expression;
            Type = type;
            Function = function;
            Title = title;
            ResultType = resultType;
        }
    }
}
