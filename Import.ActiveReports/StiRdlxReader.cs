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
using System.Globalization;
using System.IO;
using System.Xml;
using Stimulsoft.Base.Json;
using Stimulsoft.Base.Json.Linq;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Reads an RDLX report definition, XML or JSON, into a tree of StiRdlxNode.
    /// </summary>
    public static class StiRdlxReader
    {
        #region Methods
        public static StiRdlxNode ReadXml(XmlElement element)
        {
            var node = new StiRdlxNode(element.LocalName);

            foreach (XmlAttribute attribute in element.Attributes)
            {
                if (attribute.Prefix == "xmlns" || attribute.Name == "xmlns")
                    continue;

                node.Nodes.Add(new StiRdlxNode(attribute.LocalName) { Value = attribute.Value });
            }

            var hasElements = false;

            foreach (XmlNode child in element.ChildNodes)
            {
                var childElement = child as XmlElement;
                if (childElement == null)
                    continue;

                hasElements = true;
                node.Nodes.Add(ReadXml(childElement));
            }

            if (!hasElements)
                node.Value = element.InnerText;

            return node;
        }

        public static StiRdlxNode ReadJson(string text)
        {
            return ReadJson("Report", ParseJson(text));
        }

        /// <summary>
        /// Parses JSON and keeps the strings that look like dates as strings.
        /// </summary>
        public static JToken ParseJson(string text)
        {
            using (var reader = new JsonTextReader(new StringReader(text)))
            {
                reader.DateParseHandling = DateParseHandling.None;
                return JToken.Load(reader);
            }
        }

        private static StiRdlxNode ReadJson(string name, JToken token)
        {
            var node = new StiRdlxNode(name);

            var obj = token as JObject;
            if (obj != null)
            {
                foreach (var property in obj.Properties())
                {
                    if (property.Value == null || property.Value.Type == JTokenType.Null)
                        continue;

                    node.Nodes.Add(ReadJson(property.Name, property.Value));
                }

                var type = node.Get("Type");
                if (!string.IsNullOrEmpty(type))
                    node.ItemType = type;

                return node;
            }

            var array = token as JArray;
            if (array != null)
            {
                var itemName = GetItemName(name);

                foreach (var item in array)
                {
                    if (item == null || item.Type == JTokenType.Null)
                        continue;

                    node.Nodes.Add(ReadJson(itemName, item));
                }

                return node;
            }

            node.Value = GetText(token);
            return node;
        }

        /// <summary>
        /// The name of an item of a collection: TableRows gives TableRow, Properties gives Property.
        /// </summary>
        private static string GetItemName(string name)
        {
            if (name.EndsWith("ies", StringComparison.Ordinal))
                return name.Substring(0, name.Length - 3) + "y";

            if (name.EndsWith("s", StringComparison.Ordinal) && name.Length > 1)
                return name.Substring(0, name.Length - 1);

            return name;
        }

        private static string GetText(JToken token)
        {
            switch (token.Type)
            {
                case JTokenType.Boolean:
                    return token.Value<bool>() ? "true" : "false";

                case JTokenType.Integer:
                    return token.Value<long>().ToString(CultureInfo.InvariantCulture);

                case JTokenType.Float:
                    return token.Value<double>().ToString("R", CultureInfo.InvariantCulture);

                default:
                    return token.ToString();
            }
        }
        #endregion
    }
}
