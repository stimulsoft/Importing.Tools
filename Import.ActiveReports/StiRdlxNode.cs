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

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// A node of an RDLX report definition. The XML file of ActiveReports.NET and the JSON file of ActiveReportsJS
    /// are both read into a tree of these nodes, so the import does not depend on the form of the file.
    /// An XML element and a JSON object become a node, their attributes, child elements and properties become
    /// child nodes, an item of a JSON array becomes a child node of the node of the array.
    /// </summary>
    public class StiRdlxNode
    {
        #region Properties
        public string Name { get; }

        /// <summary>
        /// The type of a report item: the name of the XML element or the value of Type of the JSON object.
        /// </summary>
        public string ItemType { get; set; }

        public string Value { get; set; }

        public List<StiRdlxNode> Nodes { get; } = new List<StiRdlxNode>();
        #endregion

        #region Methods
        public StiRdlxNode Child(string name)
        {
            foreach (var node in Nodes)
            {
                if (string.Equals(node.Name, name, StringComparison.OrdinalIgnoreCase))
                    return node;
            }

            return null;
        }

        /// <summary>
        /// Returns the child with the first name or, when there is no such child, with the second name.
        /// The two forms of the file name some nodes differently.
        /// </summary>
        public StiRdlxNode Child(string name, string otherName)
        {
            return Child(name) ?? Child(otherName);
        }

        public bool Has(string name)
        {
            return Child(name) != null;
        }

        public string Get(string name)
        {
            var node = Child(name);
            if (node == null)
                return null;

            return node.Value;
        }

        public string Get(string name, string defaultValue)
        {
            var value = Get(name);
            if (string.IsNullOrEmpty(value))
                return defaultValue;

            return value;
        }

        public bool GetBool(string name)
        {
            return string.Equals(Get(name), "true", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Returns the items of a child collection, an empty list when there is no such child.
        /// </summary>
        public List<StiRdlxNode> Items(string name)
        {
            var node = Child(name);
            if (node == null)
                return new List<StiRdlxNode>();

            return node.Nodes;
        }

        public List<StiRdlxNode> Items(string name, string otherName)
        {
            var node = Child(name) ?? Child(otherName);
            if (node == null)
                return new List<StiRdlxNode>();

            return node.Nodes;
        }

        public bool IsType(string type)
        {
            return string.Equals(ItemType, type, StringComparison.OrdinalIgnoreCase);
        }

        public override string ToString()
        {
            return Value == null ? Name : $"{Name} = {Value}";
        }
        #endregion

        public StiRdlxNode(string name)
        {
            Name = name;
            ItemType = name;
        }
    }
}
