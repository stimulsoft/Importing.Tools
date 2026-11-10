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

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// A custom function of the report: Function (StringVar key, optional NumberVar n := 1) and the body, a formula
    /// in Crystal syntax which uses the parameters as variables. The translator puts the body instead of the call.
    /// </summary>
    public class StiCrystalCustomFunction
    {
        #region Properties
        public string Name { get; }

        /// <summary>
        /// The parameters as the nodes Declare: the name, the type of the variable and the default value of an
        /// optional parameter as the only child.
        /// </summary>
        public List<StiCrystalNode> Parameters { get; }

        public StiCrystalNode Body { get; }
        #endregion

        public StiCrystalCustomFunction(string name, List<StiCrystalNode> parameters, StiCrystalNode body)
        {
            Name = name;
            Parameters = parameters;
            Body = body;
        }
    }
}
