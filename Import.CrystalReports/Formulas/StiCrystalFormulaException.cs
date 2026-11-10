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

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// The exception thrown when a Crystal formula can't be parsed or translated.
    /// </summary>
    public class StiCrystalFormulaException : Exception
    {
        #region Properties
        /// <summary>
        /// The position in the formula where the problem was found, or -1 if unknown.
        /// </summary>
        public int Position { get; }

        /// <summary>
        /// The name of the unknown field, formula or parameter if the problem is an unknown reference, otherwise null.
        /// </summary>
        public string Reference { get; set; }
        #endregion

        public StiCrystalFormulaException(string message, int position = -1) :
            base(position >= 0 ? $"{message} (position {position})" : message)
        {
            Position = position;
        }
    }
}
