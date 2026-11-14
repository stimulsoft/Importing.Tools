#region Copyright (C) 2003-2026 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports												}
{	                         										}
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

using Stimulsoft.Report.CodeDom;
using Stimulsoft.Report.Dictionary;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Gives the names to the components and the variables of the imported report. The report is
    /// compiled in Visual Basic, where a name is the same in any case, and every component, data source
    /// and variable is a member of the class of the report. So two items named Textbox11 and TextBox11,
    /// or a data set and a parameter both named Company, cannot live in one report. Reporting Services
    /// keeps the items, the data sets and the parameters apart, so such names are legal there.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts
        private const string ParametersCategory = "ReportParameters";
        #endregion

        #region Fields
        //the names already given to the components, in any case
        private HashSet<string> usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        //all the names of the report definition, a renamed item does not take a name of an item to come
        private HashSet<string> definitionNames;

        //the names of the data sets, a parameter of the same name gets another name
        private HashSet<string> dataSetNames;

        //the names of the public members of the report, a component of the same name hides the member
        private static HashSet<string> reportMemberNames;
        #endregion

        #region Methods.Names
        private string CheckComponentName(StiReport report, string baseName)
        {
            var name = StiNameCreation.CreateName(report, baseName, false, true, true);

            if (name != baseName || IsNameTaken(report, name, baseName))
            {
                var root = StiNameValidator.CorrectName(baseName, report);
                var counter = 2;
                do
                {
                    name = root + counter++;
                }
                while (IsNameTaken(report, name, baseName) || StiNameCreation.Exists(report, name));
            }

            usedNames.Add(name);

            return name;
        }

        /// <summary>
        /// Tells whether the name belongs to another component, a data source, a variable or a member of
        /// the report, in any case. A component named PageNumber hides the page number of the report, so
        /// Globals!PageNumber, which is PageNumber in the report, reads the component. A name of the
        /// definition other than the own one is taken as well, so the item that has it keeps it when its
        /// turn comes.
        /// </summary>
        private bool IsNameTaken(StiReport report, string name, string baseName)
        {
            if (usedNames.Contains(name) || GetReportMemberNames().Contains(name))
                return true;

            if (report.Dictionary.DataSources.Cast<StiDataSource>().Any(d => IsSameName(d.Name, name)))
                return true;

            if (report.Dictionary.Variables.Cast<StiVariable>().Any(v => IsSameName(v.Name, name)))
                return true;

            return !IsSameName(name, baseName) && GetDefinitionNames().Contains(name);
        }

        /// <summary>
        /// Returns the name of the variable of a parameter: the name of the parameter, or the name with
        /// the Parameter suffix when a data set has the name, as a data source of Stimulsoft, or when the
        /// name is a keyword of Visual Basic, as Object, which is no name of a member there.
        /// </summary>
        private string GetParameterVariableName(StiReport report, string name)
        {
            var isKeyword = StiVBCodeGenerator.IsKeywordExist(name.ToLowerInvariant());
            if (!isKeyword && !GetDataSetNames().Contains(name))
                return name;

            var root = name + "Parameter";
            var newName = root;
            var counter = 2;
            while (GetDataSetNames().Contains(newName) || GetDefinitionNames().Contains(newName) ||
                report.Dictionary.Variables.Cast<StiVariable>().Any(v => IsSameName(v.Name, newName)))
            {
                newName = root + counter++;
            }

            ThrowError(null, null, isKeyword
                ? $"Parameter renamed, its name is a keyword of Visual Basic: {name} -> {newName}"
                : $"Parameter renamed, a data set has its name: {name} -> {newName}");

            return newName;
        }

        private static HashSet<string> GetReportMemberNames()
        {
            if (reportMemberNames == null)
            {
                var members = typeof(StiReport).GetMembers(BindingFlags.Public | BindingFlags.Instance |
                    BindingFlags.Static | BindingFlags.FlattenHierarchy);

                reportMemberNames = new HashSet<string>(members.Select(m => m.Name), StringComparer.OrdinalIgnoreCase);
            }

            return reportMemberNames;
        }

        private HashSet<string> GetDefinitionNames()
        {
            if (definitionNames == null)
                definitionNames = CollectNames("//*[@Name]");

            return definitionNames;
        }

        private HashSet<string> GetDataSetNames()
        {
            if (dataSetNames == null)
                dataSetNames = CollectNames("//*[local-name()='DataSet'][@Name]");

            return dataSetNames;
        }

        private HashSet<string> CollectNames(string xpath)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var document = reportNode as XmlDocument ?? reportNode?.OwnerDocument;
            if (document == null)
                return names;

            foreach (XmlNode node in document.SelectNodes(xpath))
            {
                names.Add(node.Attributes["Name"].Value);
            }

            return names;
        }

        private static bool IsSameName(string name1, string name2)
        {
            return string.Equals(name1, name2, StringComparison.OrdinalIgnoreCase);
        }
        #endregion
    }
}
