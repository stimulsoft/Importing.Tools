#region Copyright (C) 2003-2025 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports  											}
{																	}
{	Copyright (C) 2003-2025 Stimulsoft     							}
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
#endregion Copyright (C) 2003-2025 Stimulsoft

using System;

namespace Import.CrystalReports
{
    public class StiWarning
    {
        #region Fields
        public string Message = null;
        public string Description = null;
        #endregion

        #region Static methods
        public static StiWarning GetWarningPicture(string objectName)
        {
            return new StiWarning(
                string.Format("Image '{0}': Impossible to get content", objectName),
                "When using CrystalReports from .Net, it is possible to get only some properties of report object." +
                " For example, it is possible to get only image size and location, it is impossible to get image content.");
        }

        public static StiWarning GetWarningDatabase(string objectName)
        {
            return new StiWarning(
                string.Format("Database '{0}': Could not identify the type", objectName),
                "When connecting to data bases CrystalReports often uses its own internal libraries." +
                " It is possible to get only some properties of connection from .Net and impossible to get ConnectionString." +
                " Therefore, not all data bases can be identified." +
                " By default, for not identified bases StiOleDbDatabase and ConnectionString is used without specifying provider.");
        }

        public static StiWarning GetWarningFormulaField(string objectName)
        {
            return new StiWarning(
                string.Format("FormulaField '{0}': Expression can be incorrect", objectName),
                "CrystalReports allows using expressions and formulas in FormulaFields." +
                " On the current moment parsing and syntax converting cannot be done, expressions is written 'as is'." +
                " Therefore, in many cases it is required further manual correction of expressions.");
        }

        public static StiWarning GetWarningFormula(string place, string reason)
        {
            return new StiWarning(
                string.Format("Formula '{0}': {1}", place, reason),
                "The formula can't be translated to an expression of Stimulsoft Reports, its text is kept 'as is'." +
                " Please correct the expression manually.");
        }

        public static StiWarning GetWarningFormat(string objectName, string property, string formula)
        {
            return new StiWarning(
                string.Format("Format of '{0}': {1} is calculated by the formula {2}", objectName, property, formula),
                "A format of Stimulsoft Reports has the fixed settings, so the value of the formula is not used and" +
                " the static setting of the field is kept. Please correct the format manually if it is needed.");
        }
        #endregion

        #region this
        public StiWarning()
            : this(string.Empty, string.Empty)
        {
        }

        public StiWarning(string message, string description)
        {
            this.Message = message;
            this.Description = description;
        }
        #endregion
    }
}
