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

using Stimulsoft.Report.Dictionary;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Makes the shared data sources and the shared data sets of Reporting Services. A report keeps only the
    /// name of such a source or data set, the definition lives on the report server or in the project, as
    /// an .rds or an .rsd file next to the report. When the folder of these files is given to the import,
    /// the connection of a source and the query of a data set are read from them; otherwise the source is
    /// a database without the connection and the data set is named in the list of errors.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts
        private const string SharedSourceExtension = ".rds";

        private const string SharedDataSetExtension = ".rsd";
        #endregion

        #region Fields.Shared
        //the folder of the .rds and .rsd files, null when they are not read
        private string sharedFolder;

        //the databases of the shared data sources, by the name of the definition
        private Dictionary<string, StiDatabase> sharedDatabases =
            new Dictionary<string, StiDatabase>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Methods.SharedSources
        /// <summary>
        /// Makes the database of a shared data source. The reference is the path of the definition on the
        /// server, its last part is the name of the .rds file. Without the file the database has no
        /// connection, so the data sets keep their queries.
        /// </summary>
        private StiDatabase CreateSharedDatabase(string name, string reference)
        {
            var definitionName = GetDefinitionName(reference);
            var definition = LoadDefinition(definitionName, SharedSourceExtension);
            var database = definition != null ? CreateDatabaseOfDefinition(name, definition) : null;

            if (database == null)
            {
                ThrowError(null, null, $"Shared data source not imported, set the connection: {name} ({reference})");
                database = new StiSqlDatabase(name, string.Empty);
            }

            if (definitionName != null && !sharedDatabases.ContainsKey(definitionName))
                sharedDatabases[definitionName] = database;

            return database;
        }

        /// <summary>
        /// Makes the database of an .rds file. Report Builder and SSDT write the definition with the
        /// extension and the connection on the top level and the user and the password of the stored
        /// credentials; a project of Visual Studio writes them in ConnectionProperties, as a report does.
        /// </summary>
        private StiDatabase CreateDatabaseOfDefinition(string name, XmlElement definition)
        {
            var properties = GetChild(definition, "ConnectionProperties") ?? definition;
            var provider = GetChild(properties, "Extension")?.InnerText.Trim();
            var connectString = GetChild(properties, "ConnectString")?.InnerText.Trim() ?? string.Empty;
            var baseConnectString = connectString;

            if (IsTrue(GetChild(properties, "IntegratedSecurity")?.InnerText) ||
                string.Equals(GetChild(definition, "CredentialRetrieval")?.InnerText, "Integrated",
                    StringComparison.OrdinalIgnoreCase))
            {
                connectString = AddConnectionPart(connectString, "Integrated Security=True");
            }

            var userName = GetChild(definition, "UserName")?.InnerText;
            if (!string.IsNullOrEmpty(userName) && !IsTrue(GetChild(definition, "WindowsCredentials")?.InnerText))
            {
                connectString = AddConnectionPart(connectString, $"User ID={userName}");

                var password = GetChild(definition, "Password")?.InnerText;
                if (!string.IsNullOrEmpty(password))
                    connectString = AddConnectionPart(connectString, $"Password={password}");
            }

            return CreateDatabase(name, provider, connectString, baseConnectString);
        }

        private static string AddConnectionPart(string connectString, string part)
        {
            connectString = connectString.TrimEnd();
            if (connectString.Length > 0 && !connectString.EndsWith(";"))
                connectString += ";";

            return connectString + part;
        }
        #endregion

        #region Methods.SharedDataSets
        /// <summary>
        /// Makes the data source of a shared data set. The query, its type and its data source are read
        /// from the .rsd file, the values of the parameters come from the report, the fields too. The query
        /// is given to the import of a query of the report, so a shared data set becomes the same data source
        /// as a data set written into the report.
        /// </summary>
        private StiDataSource ProcessSharedDataSetType(XmlNode baseNode, StiReport report, ArrayList columns)
        {
            var dataSetName = baseNode.ParentNode.Attributes["Name"].Value;
            var reference = GetChild(baseNode, "SharedDataSetReference")?.InnerText;
            var definition = LoadDefinition(GetDefinitionName(reference), SharedDataSetExtension);
            var definitionQuery = GetChild(GetChild(definition, "DataSet"), "Query");

            if (definitionQuery == null)
            {
                ThrowError(null, null, $"Shared data set not imported: {dataSetName}");
                return null;
            }

            var document = baseNode.OwnerDocument;
            var dataSet = document.CreateElement("DataSet", baseNode.NamespaceURI);
            dataSet.SetAttribute("Name", dataSetName);

            var query = document.CreateElement("Query", baseNode.NamespaceURI);
            dataSet.AppendChild(query);
            AddQueryChild(query, "DataSourceName", GetSharedDataSetDatabase(report, definitionQuery)?.Name);
            AddQueryChild(query, "CommandText", GetChild(definitionQuery, "CommandText")?.InnerText);
            AddQueryChild(query, "CommandType", GetChild(definitionQuery, "CommandType")?.InnerText);

            var parameters = GetChild(baseNode, "QueryParameters");
            if (parameters != null)
                query.AppendChild(parameters.CloneNode(true));

            return ProcessQueryType(query, report, columns);
        }

        /// <summary>
        /// Returns the database of the shared data source of a shared data set: the one the report already
        /// has for this definition, else a new one added to the report.
        /// </summary>
        private StiDatabase GetSharedDataSetDatabase(StiReport report, XmlElement definitionQuery)
        {
            var reference = GetChild(definitionQuery, "DataSourceReference")?.InnerText;
            var definitionName = GetDefinitionName(reference);
            if (definitionName == null)
                return null;

            if (sharedDatabases.TryGetValue(definitionName, out var database))
                return database;

            database = CreateSharedDatabase(definitionName, reference);
            if (report.Dictionary.Databases[database.Name] == null)
                report.Dictionary.Databases.Add(database);

            return database;
        }

        private static void AddQueryChild(XmlElement query, string name, string value)
        {
            if (value == null)
                return;

            var child = query.OwnerDocument.CreateElement(name, query.NamespaceURI);
            child.InnerText = value;
            query.AppendChild(child);
        }
        #endregion

        #region Methods.Definitions
        /// <summary>
        /// Returns the name of the file of a definition: the last part of its path on the server, without the
        /// folders, so the reference cannot lead out of the folder of the definitions.
        /// </summary>
        private static string GetDefinitionName(string reference)
        {
            if (string.IsNullOrWhiteSpace(reference))
                return null;

            var name = reference.Trim().Split('/', '\\').Last().Trim();
            var isValid = name.Length > 0 && name != ".." && name.IndexOfAny(Path.GetInvalidFileNameChars()) == -1;

            return isValid ? name : null;
        }

        /// <summary>
        /// Reads the definition with the given name and extension from the folder of the definitions. Null
        /// when the folder is not given, the file is absent or is not XML.
        /// </summary>
        private XmlElement LoadDefinition(string name, string extension)
        {
            if (sharedFolder == null || name == null)
                return null;

            var path = Path.Combine(sharedFolder, name + extension);
            if (!File.Exists(path))
                return null;

            try
            {
                var document = new XmlDocument { XmlResolver = null };
                using (var reader = XmlReader.Create(path, GetSafeXmlSettings()))
                {
                    document.Load(reader);
                }

                return document.DocumentElement;
            }
            catch (Exception)
            {
                ThrowError(null, null, $"Definition not read: {name}{extension}");
                return null;
            }
        }

        private static XmlElement GetChild(XmlNode node, string localName)
        {
            return node?.ChildNodes.OfType<XmlElement>().FirstOrDefault(e => e.LocalName == localName);
        }
        #endregion
    }
}
