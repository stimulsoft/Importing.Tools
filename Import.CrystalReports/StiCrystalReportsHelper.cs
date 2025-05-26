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
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components;
using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Runtime.ExceptionServices;
//using CrystalDecisions.ReportAppServer.ReportDefModel;


namespace Import.CrystalReports
{
    /// <summary>
    /// Class helps converts Crystal Reports templates to Stimulsoft Reports templates.
    /// </summary>
    [System.Security.SecuritySafeCritical]
    [System.Security.SuppressUnmanagedCodeSecurity]
	public sealed class StiCrystalReportsHelper
	{
        #region Enums
        public enum CrTextFormatEnum
        {
            crTextFormatStandardText,
            crTextFormatRTFText,
            crTextFormatHTMLText
        }

        public enum CrTextRotationAngleEnum
        {
            crRotationAngleRotate0,
            crRotationAngleRotate90,
            crRotationAngleRotate270
        }
        #endregion

        #region Options
        //if true - use StiPrimitives, else use Shapes
        private bool usePrimitives = true;
        //if true - use variables as functions, else use variables as string
        private bool useFunctions = false;
        #endregion

        #region Constants
        //private const float unit = 0.0017761989342806394316163410301954f; //cm
        //private const float unit = 1f / 1440 * 2.54;     //cm
        private const float unit = 1f / 1440 * 100;     //hi
        #endregion

        #region Variables
        private StiDataSource mainDataSource = null;
		private Hashtable htDataBaseNames = null;
		private Hashtable htTableNameToDatabase = null;
		private Hashtable htTableNameToDataSource = null;
		private Hashtable htDataSourceToNameWithRelation = null;
		private Hashtable htDataSourceParent = null;
		private Hashtable htFieldNameConversion = null;
        private Hashtable htVariableConversion = null;
        private Hashtable textConversion = null;
        private ArrayList textConversionSorted = null;
        private bool flagTextConversionSorted = false;
        private string[] sortFields = null;
        private ArrayList groupFields = null;
        private int groupFieldsCounter = 0;
        private ArrayList summaryFieldsList = null;
        private ArrayList subreportsList = null;
        private StiReport stimulReport = null;
        private Hashtable htNames = null;
        private ArrayList dataSourcesOnThisPage = null;
        private List<string> methods = null;
        #endregion

        public List<StiWarning> warningList = new List<StiWarning>();
        private static string internalVersion = StiReport.GetReportVersion();

        #region Utils
        private void MakeTableNameRecursive(StiDataSource table, string baseName)
		{
			string currentName = baseName + ".";
			htDataSourceToNameWithRelation[table] = currentName;
			StiDataRelationsCollection drc = table.GetParentRelations();
			foreach (StiDataRelation relation in drc)
			{
				MakeTableNameRecursive(relation.ParentSource, currentName + relation.Name);
			}
		}

		private static string FieldNameCorrection(string baseName)
		{
			StringBuilder sb = new StringBuilder(baseName);
			int charIndex = 0;
			for (int pos = 0; pos < baseName.Length; pos++)
			{
				if (baseName[pos] == ' ')
				{
					sb.Remove(charIndex, 1);
					if (charIndex < sb.Length && Char.IsLetter(sb[charIndex]))
					{
						sb[charIndex] = Char.ToUpper(sb[charIndex]);
					}
				}
				else charIndex++;
			}
			for (int pos = 0; pos < sb.Length; pos++)
			{
				if (!(Char.IsLetterOrDigit(sb[pos]) || sb[pos] == '_' || sb[pos] == '.')) sb[pos] = '_';
			}
			baseName = sb.ToString();
			return baseName;
		}

        public class myComparerClass : IComparer    //sort list by strings length in revese order
        {
            int IComparer.Compare(object x, object y)
            {
                return (((DictionaryEntry)y).Key as string).Length.CompareTo((((DictionaryEntry)x).Key as string).Length);
            }
        }

        private string ReplaceFieldsAndVariables(string text)
        {
            if (textConversionSorted == null || textConversionSorted.Count != textConversion.Count || !flagTextConversionSorted)
            {
                textConversionSorted = new ArrayList();
                foreach (DictionaryEntry de in textConversion)
                {
                    textConversionSorted.Add(de);
                }
                IComparer myComparer = new myComparerClass();
                textConversionSorted.Sort(myComparer);
                flagTextConversionSorted = true;
            }

            StringBuilder sb = new StringBuilder(text);
            //for (int index = 0; index < textConversion.Length; index++)
            //{
            //    sb.Replace((string)textConversion[index].Key, (string)textConversion[index].Value);
            //}
            //foreach (DictionaryEntry de in textConversion)
            //{
            //    sb.Replace((string)de.Key, (string)de.Value);
            //}
            foreach (DictionaryEntry de in textConversionSorted)
            {
                sb.Replace((string)de.Key, (string)de.Value);
            }
            return sb.ToString();
        }

        private void AddRangeToHashtable(Hashtable main, Hashtable range)
        {
            foreach (DictionaryEntry de in range)
            {
                main[de.Key] = de.Value;
            }
        }

        private string CheckComponentsNames(string name)
        {
            if (htNames.Contains(name))
            {
                int nameIndex = 1;
                while (htNames.Contains(string.Format("{0}_{1}", name, nameIndex)))
                {
                    nameIndex++;
                }
                name = string.Format("{0}_{1}", name, nameIndex);
            }
            htNames[name] = name;
            return name;
        }

        public static string ReplaceSymbols(string str)
        {
            StringBuilder sb = new StringBuilder(str);
            for (int pos = 0; pos < sb.Length; pos++)
            {
                //if ((!(Char.IsLetterOrDigit(sb[pos]) || sb[pos] == '_')) & (!(sb[pos] == '.'))) sb[pos] = '_';
                if (!(Char.IsLetterOrDigit(sb[pos]) || sb[pos] == '_')) sb[pos] = '_';
            }
            return sb.ToString();
        }

        private bool CheckFunction(StiVariable variable, string value)
        {
            bool flag = true;
            string formula = value.Trim();
            formula = NormalizeEndOfLineSymbols(formula).Replace("\r\n", " \r\n ");

            try
            {
                formula = formula.Replace("ToNumber", "double.Parse");
                formula = formula.Replace(" {", " ").Replace("{", " ");
                formula = formula.Replace("} ", " ").Replace("}", " ");
                formula = formula.Replace(" AND ", " && ");
                formula = formula.Replace(" and ", " && ");
                formula = formula.Replace(" OR ", " || ");
                formula = formula.Replace(" or ", " || ");

                formula = formula.Replace("IsNull", "IsNull1");
                formula = formula.Replace("isnull", "IsNull1");
                formula = formula.Replace("trim", "Trim");
                formula = formula.Replace("length", "Length");
                formula = formula.Replace("Pagenumber", "PageNumber");

                formula = formula.Replace("Chr(13)", "\"\\r\"");
                formula = formula.Replace("CHR(13)", "\"\\r\"");
                formula = formula.Replace("ChrW(13)", "\"\\r\"");
                formula = formula.Replace("Chr(10)", "\"\\n\"");
                formula = formula.Replace("CHR(10)", "\"\\n\"");
                formula = formula.Replace("ChrW(10)", "\"\\n\"");

                formula = formula.Replace("<=", "\x18\x01");
                formula = formula.Replace(">=", "\x18\x02");
                formula = formula.Replace(":=", "\x18\x03");
                formula = formula.Replace("=", "==");
                formula = formula.Replace("\x18\x01", "<=");
                formula = formula.Replace("\x18\x02", ">=");
                formula = formula.Replace("\x18\x03", ":=");
                formula = formula.Replace("<>", "!=");

                if (formula.Contains(":="))
                {
                    formula = formula.Replace(" If ", " if(");
                    formula = formula.Replace(" if ", " if(");
                    formula = formula.Replace(" then ", " ) ");
                    formula = formula.Replace(" Then ", " ) ");
                    formula = formula.Replace(" Else ", " else ");

                    formula = formula.Replace(":=", "=");

                    //identation
                    List<string> lines = SplitString(formula);
                    for (int index = 0; index < lines.Count; index++)
                    {
                        lines[index] = "\t" + lines[index] + " ";
                    }

                    //make result
                    string result = "null";
                    string lastLine = lines[lines.Count - 1].Trim();
                    if (!lastLine.Contains(" ") && !lastLine.Contains(":="))
                    {
                        result = lastLine;
                        lines.RemoveAt(lines.Count - 1);
                    }

                    //add method description
                    lines.Insert(0, string.Format("private {0} mt_{1}()", variable.Type.Name.ToString(), variable.Name.ToString()));
                    lines.Insert(1, "{");
                    lines.Add("\treturn " + result + ";");
                    lines.Add("}");
                    lines.Add(string.Empty);

                    //add to methods list
                    foreach (string line in lines)
                    {
                        methods.Add(line);
                    }

                    formula = string.Format("mt_{0}()", variable.Name.ToString());
                }
                else
                {
                    if (formula.StartsWith("if ")) formula = " " + formula;
                    formula = formula.Replace(" if ", " ");
                    formula = formula.Replace(" If ", " ");
                    formula = formula.Replace(" then ", " ? ");
                    formula = formula.Replace(" Then ", " ? ");
                    formula = formula.Replace(" else ", " : ");
                    formula = formula.Replace(" Else ", " : ");

                    formula = formula.Replace("UpperCase", "ToUpperCase");
                    formula = formula.Replace("LowerCase", "ToLowerCase");

                    //formula = formula.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ");
                }

                variable.InitBy = StiVariableInitBy.Expression;
                variable.Value = formula;
            }
            catch
            {
                variable.InitBy = StiVariableInitBy.Value;
                variable.Type = typeof(string);
                variable.Value = value;
                flag = false;
            }
            return flag;
        }

        private static string NormalizeEndOfLineSymbols(string inputText)
        {
            if (inputText == null || inputText.Length < 2 || (inputText.IndexOf('\r') == -1 && inputText.IndexOf('\n') == -1))
                return inputText;

            var sb = new StringBuilder();
            for (var index = 0; index < inputText.Length; index++)
            {
                var ch = inputText[index];
                if (ch == '\r' || ch == '\n')
                {
                    if (index + 1 < inputText.Length)
                    {
                        var ch2 = inputText[index + 1];
                        if ((ch2 == '\r' || ch2 == '\n') && (ch2 != ch))
                            index++;
                    }
                    sb.Append("\r\n");
                }
                else
                {
                    sb.Append(ch);
                }
            }

            return sb.ToString();
        }

        public static List<string> SplitString(string inputString)
        {
            var stringList = new List<string>();
            if (inputString == null) inputString = string.Empty;

            var st = new StringBuilder();
            foreach (char ch in inputString)
            {
                if (ch == '\n') continue;
                if (ch == '\r')
                {
                    stringList.Add(st.ToString().TrimEnd());
                    st.Length = 0;
                }
                else
                {
                    st.Append(ch);
                }
            }
            if (st.Length > 0) stringList.Add(st.ToString().TrimEnd());
            if (stringList.Count == 0) stringList.Add(string.Empty);

            return stringList;
        }

        private Color GetBackgroundColorFromBorder(Border border)
        {
            Color color = border.BackgroundColor;

            object obj1 = GetPropertyFromHiddenObject(border, "RasBorder.BackgroundColor");
            if (obj1 != null)
            {
                uint backColor = (uint)obj1;
                backColor = backColor ^ 0xFF000000;
                color = Color.FromArgb((int)backColor);
            }
            return color;
        }

        private float GetTextRotationAngle(ReportObject obj, string section)
        {
            CrTextRotationAngleEnum rotationAngle = CrTextRotationAngleEnum.crRotationAngleRotate0;
            object objj = GetPropertyFromHiddenComObject(obj, section + ".Format.TextRotationAngle");
            if (objj != null)
            {
                rotationAngle = (CrTextRotationAngleEnum)objj;
            }

            float angle = 0;
            switch (rotationAngle)
            {
                case CrTextRotationAngleEnum.crRotationAngleRotate90:
                    angle = 90;
                    break;

                case CrTextRotationAngleEnum.crRotationAngleRotate270:
                    angle = 270;
                    break;

                default:
                    angle = 0;
                    break;
            }
            return angle;
        }
        #endregion

        #region Get hidden properties
        private static object GetPropertyFromHiddenObject(object parentObj, string property)
        {
            if (parentObj == null) return null;
            string main = property;
            string subs = null;
            int posDot = property.IndexOf('.');
            if (posDot != -1)
            {
                main = property.Substring(0, posDot);
                subs = property.Substring(posDot + 1);
            }

            object obj = null;
            Type typeFormat1 = parentObj.GetType();
            System.Reflection.PropertyInfo infoFormat1 = typeFormat1.GetProperty(main, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (infoFormat1 != null)
            {
                obj = infoFormat1.GetValue(parentObj, null);  //get object
            }

            if (subs != null)
            {
                return GetPropertyFromHiddenObject(obj, subs);
            }
            return obj;
        }

        private static object GetPropertyFromHiddenComObject(object parentObj, string property)
        {
            if (parentObj == null) return null;
            string main = property;
            string subs = null;
            int posDot = property.IndexOf('.');
            if (posDot != -1)
            {
                main = property.Substring(0, posDot);
                subs = property.Substring(posDot + 1);
            }

            object obj = null;
            try
            {
                obj = parentObj.GetType().InvokeMember(
                    main,
                    BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Instance |
                     BindingFlags.IgnoreCase | BindingFlags.GetProperty | BindingFlags.GetField,
                    null,
                    parentObj,
                    null);
            }
            catch
            {
                obj = null;
            }

            if (subs != null)
            {
                return GetPropertyFromHiddenComObject(obj, subs);
            }
            return obj;
        }
        #endregion

        public StiReport Convert(ReportDocument crystalReport)
        {
            return Convert(crystalReport, null, true, false);
        }
    
		public StiReport Convert(ReportDocument crystalReport, IStiTreeLog log, bool usePrimitives, bool useFunctions)
		{
            this.usePrimitives = usePrimitives;
            this.useFunctions = useFunctions;
            log.OpenLog("Report");

            methods = new List<string>();

			stimulReport = new StiReport();
            stimulReport.MetaTags.Add("InternalVersion", internalVersion);
			try
            {
                #region Report Parameters
                log.OpenNode("Report Parameters");

				stimulReport.ReportUnit = StiReportUnitType.HundredthsOfInch;

				stimulReport.ReportAuthor =	crystalReport.SummaryInfo.ReportAuthor;
                log.WriteNode(" Report Author: ", stimulReport.ReportAuthor);

				stimulReport.ReportDescription = crystalReport.SummaryInfo.ReportComments;
                log.WriteNode(" Report Description: ", stimulReport.ReportDescription);

				stimulReport.ReportAlias = crystalReport.SummaryInfo.ReportTitle;
                log.WriteNode(" Report Alias: ", stimulReport.ReportAlias);

                log.CloseNode();
                #endregion

                stimulReport.Info.ShowHeaders = false;
                htNames = new Hashtable();

                if ((crystalReport.Subreports != null) && (crystalReport.Subreports.Count > 0))
                {
                    for (int indexPage = 0; indexPage < crystalReport.Subreports.Count; indexPage++)
                    {
                        StiPage page = new StiPage(stimulReport);
                        page.Name = crystalReport.Subreports[indexPage].Name;
                        stimulReport.Pages.Add(page);
                    }
                }

                AddPage(stimulReport.Pages[0], crystalReport, false, log);

                if ((crystalReport.Subreports != null) && (crystalReport.Subreports.Count > 0))
                {
                    for (int indexPage = 0; indexPage < crystalReport.Subreports.Count; indexPage++)
                    {
                        log.OpenNode(string.Format("Subreport: {0}", crystalReport.Subreports[indexPage].Name));
                        //stimulReport.Pages.Add(new StiPage(stimulReport));
                        AddPage(stimulReport.Pages[1 + indexPage], crystalReport.Subreports[indexPage], true, log);
                        log.CloseNode();
                    }
                }


                log.OpenNode("Warnings");
                foreach (StiWarning warning in warningList)
                {
                    log.WriteNode(warning.Message);
                }
                log.CloseNode();
            }
			catch (Exception e)
			{
                StiExceptionProvider.Show(e);
            }

            #region Postprocess report
            stimulReport.Info.ShowHeaders = true;
            stimulReport.Info.ForceDesigningMode = true;    //force DockToContainer for components with Enabled=false
            foreach (StiPage page in stimulReport.Pages)
            {
                page.DockToContainer();
                page.Correct();
            }
            stimulReport.Info.ForceDesigningMode = false;

            foreach (StiComponent comp in stimulReport.GetComponents())
            {
                comp.Linked = false;
            }

            foreach (StiDataSource ds in stimulReport.Dictionary.DataSources)
            {
                ds.Columns.Sort(StiSortOrder.Asc);
            }
            #endregion

            #region Add methods
            //if (methods.Count > 0)
            {
                int methodsPos = stimulReport.Script.IndexOf("#region StiReport");
                if (methodsPos > 0)
                {
                    methodsPos -= 9;

                    methods.AddRange(new string[] {
                        "private string ToText(string x) { return x; }",
                        "private string ToText(double x) { return x.ToString(); }",
                        "private string ToText(double x, int y) { return x.ToString(\"F\" + y.ToString()); }",
                        "private string ToText(DateTime x, string y) { return x.ToString(y); }",
                        "private string Trim(string st) { return st.Trim(); }",
                        "private string TrimRight(string st) { return st.TrimEnd(); }",
                        "private string TrimLeft(string st) { return st.TrimStart(); }",
                        "private string UpperCase(string st) { return st.ToUpper(); }",
                        "private string LowerCase(string st) { return st.ToLower(); }",
                        "private bool IsNull1(object obj) { return obj == null || obj == DBNull.Value; }",
                        "private DateTime DateValue(object obj) { return (DateTime)Stimulsoft.Base.StiConvert.ChangeType(obj, typeof(DateTime), true); }"
                    });

                    foreach (string st in methods)
                    {
                        string st2 = "\t\t" + st + "\r\n";
                        stimulReport.Script = stimulReport.Script.Insert(methodsPos, st2);
                        methodsPos += st2.Length;
                    }
                }
            }
            #endregion

            log.CloseLog();

			return stimulReport;
		}

        #region Adds
        private void AddPage(StiPage page, ReportDocument crystalReport, bool isSubReport, IStiTreeLog log)
        {
            #region Page Parameters
            page.TitleBeforeHeader = true;
            //page.LargeHeight = true;

            if (!isSubReport)
            {
	            log.OpenNode("Page Parameters");

                page.PaperSize = (System.Drawing.Printing.PaperKind)crystalReport.PrintOptions.PaperSize;
                bool flag = false;
                try
                {
                    string paperKindName = Enum.GetName(typeof (PaperKind), page.PaperSize);
                    if (paperKindName == null) flag = true;
                }
                catch
                {
                    flag = true;
                }
                if (flag)
                {
                    page.PaperSize = PaperKind.Custom;
                }
                log.WriteNode("Paper Size: ", page.PaperSize.ToString());

	            double marginsBottom = Math.Round(crystalReport.PrintOptions.PageMargins.bottomMargin * unit, 2);
	            log.WriteNode("Page Margin Bottom: ", marginsBottom);
	
	            double marginsLeft = Math.Round(crystalReport.PrintOptions.PageMargins.leftMargin * unit, 2);
	            log.WriteNode("Page Margin Left: ", marginsLeft);
	
	            double marginsRight = Math.Round(crystalReport.PrintOptions.PageMargins.rightMargin * unit, 2);
	            log.WriteNode("Page Margin Right: ", marginsRight);
	
	            double marginsTop = Math.Round(crystalReport.PrintOptions.PageMargins.topMargin * unit, 2);
	            log.WriteNode("Page Margin Top: ", marginsTop);

                page.Margins = new StiMargins(marginsLeft, marginsRight, marginsTop, marginsBottom);

                if (crystalReport.PrintOptions.PaperOrientation == PaperOrientation.Landscape)
                {
                    page.Orientation = StiPageOrientation.Landscape;
                }
                else
                {
                    page.Orientation = StiPageOrientation.Portrait;
                }
	            log.WriteNode("Page Orientation: ", page.Orientation);

	            log.CloseNode();
            }
            #endregion

            #region DataBases
            log.OpenNode("Databases");

            htDataBaseNames = new Hashtable();
            htTableNameToDatabase = new Hashtable();
            htTableNameToDataSource = new Hashtable();
            dataSourcesOnThisPage = new ArrayList();

            #region for CrystalReport v10 managed 2.5
            //foreach (IConnectionInfo conInfo in crystalReport.DataSourceConnections)
            //{
            //    string DataBaseName = conInfo.DatabaseName;
            //    string DataBaseDll = null;
            //    string QEDataBaseName = null;
            //    string QEDataBaseType = null;
            //    string QELogonDataBaseName = null;
            //    string QELogonDataBaseType = null;
            //    string QEServerDescription = null;
            //    bool QESqlDB = false;
            //    string foundNewNameValue = null;

            //    foreach (NameValuePair2 nameValue in conInfo.Attributes.Collection)
            //    {
            //        #region Get Connection Info Attributes
            //        switch (nameValue.Name as string)
            //        {
            //            case "Database DLL":
            //                DataBaseDll = (string)nameValue.Value;
            //                break;

            //            case "QE_DatabaseName":
            //                QEDataBaseName = (string)nameValue.Value;
            //                break;

            //            case "QE_DatabaseType":
            //                QEDataBaseType = (string)nameValue.Value;
            //                break;

            //            case "QE_LogonProperties":
            //                #region LogonProperties
            //                DbConnectionAttributes logonAttributes = (DbConnectionAttributes)nameValue.Value;
            //                foreach (NameValuePair2 nameValue2 in logonAttributes.Collection)
            //                {
            //                    switch (nameValue2.Name as string)
            //                    {
            //                        case "Database Name":
            //                            QELogonDataBaseName = (string)nameValue2.Value;
            //                            break;

            //                        case "Database Type":
            //                            QELogonDataBaseType = (string)nameValue2.Value;
            //                            break;

            //                        default:
            //                            foundNewNameValue = "Yes";
            //                            break;
            //                    }
            //                }
            //                #endregion
            //                break;

            //            case "QE_ServerDescription":
            //                QEServerDescription = (string)nameValue.Value;
            //                break;

            //            case "QE_SQLDB":
            //                QESqlDB = (bool)nameValue.Value;
            //                break;

            //            default:
            //                foundNewNameValue = "Yes";
            //                break;
            //        }
            //        #endregion
            //    }

            //    htDataBaseNames[DataBaseName] = DataBaseName;
            //    string databaseConnectionString = null;
            //    bool dataBaseIsOleDB = false;
            //    switch (QELogonDataBaseType)
            //    {
            //        case "Access":
            //            databaseConnectionString = string.Format(
            //                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};User Id=admin;Password=;",
            //                conInfo.ServerName);
            //            dataBaseIsOleDB = true;
            //            break;

            //        case "DBase":
            //            databaseConnectionString = "Provider=VFPOLEDB.1;Data Source=C:\\";
            //            dataBaseIsOleDB = true;
            //            break;
            //    }

            //    if (dataBaseIsOleDB)
            //    {
            //        StiOleDbDatabase database = new StiOleDbDatabase(DataBaseName, databaseConnectionString);
            //        bool founded = false;
            //        foreach (StiDatabase tempDataBase in stimulReport.Dictionary.Databases)
            //        {
            //            if (tempDataBase.Name == database.Name) founded = true;
            //        }
            //        if (!founded)
            //        {
            //            stimulReport.Dictionary.Databases.Add(database);
            //        }
            //    }
            //}
            #endregion

            #region for CrystalReport v9 managed 1.1
            //foreach (Table table in crystalReport.Database.Tables)
            //{
            //    string crDatabaseName = table.LogOnInfo.ConnectionInfo.DatabaseName;
            //    string crServerName = table.LogOnInfo.ConnectionInfo.ServerName;
            //    if (crDatabaseName == string.Empty) crDatabaseName = crServerName;
            //    if (crDatabaseName == string.Empty) crDatabaseName = table.Name;

            //    if (htDataBaseNames[crDatabaseName] == null)
            //    {
            //        #region make new database
            //        string databaseName = string.Empty;
            //        string databaseConnectionString = string.Empty;
            //        StiDatabase database = null;

            //        //database Access/Excell
            //        if (crServerName.ToLower().EndsWith(".mdb") ||
            //            crServerName.ToLower().EndsWith(".xls"))
            //        {
            //            databaseConnectionString = string.Format(
            //                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};User Id={1};Password={2};",
            //                crServerName,
            //                table.LogOnInfo.ConnectionInfo.UserID,
            //                table.LogOnInfo.ConnectionInfo.Password);
            //            databaseName = crServerName.Replace(':', '*');
            //            int symIndex = databaseName.LastIndexOf('\\');
            //            if ((symIndex != -1) && (symIndex < databaseName.Length - 1))
            //            {
            //                databaseName = databaseName.Substring(symIndex + 1);
            //            }
            //            database = new StiOleDbDatabase(databaseName, databaseConnectionString);
            //        }

            //        //database SQL
            //        if ((crServerName != crDatabaseName) && (table.Location.ToLower().IndexOf(".dbo.") != -1))
            //        {
            //            databaseConnectionString = string.Format(
            //                "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=True;",
            //                crServerName,
            //                crDatabaseName,
            //                table.LogOnInfo.ConnectionInfo.UserID,
            //                table.LogOnInfo.ConnectionInfo.Password);
            //            databaseName = crDatabaseName;
            //            database = new StiSqlDatabase(databaseName, databaseConnectionString);
            //        }

            //        //another databases
            //        if (database == null)
            //        {
            //            databaseConnectionString = string.Format(
            //                "Provider=  ;Data Source={0};User Id={1};Password={2};",
            //                crServerName,
            //                table.LogOnInfo.ConnectionInfo.UserID,
            //                table.LogOnInfo.ConnectionInfo.Password);
            //            databaseName = crServerName.Replace(':', '*');
            //            database = new StiOleDbDatabase(databaseName, databaseConnectionString);
            //            warningList.Add(StiWarning.GetWarningDatabase(databaseName));
            //        }

            //        bool founded = false;
            //        foreach (StiDatabase tempDataBase in stimulReport.Dictionary.Databases)
            //        {
            //            if (tempDataBase.Name == database.Name) founded = true;
            //        }
            //        if (!founded)
            //        {
            //            stimulReport.Dictionary.Databases.Add(database);
            //        }
            //        htDataBaseNames[table.LogOnInfo.ConnectionInfo.DatabaseName] = databaseName;
            //        #endregion

            //        htTableNameToDatabase[table.Name] = database;
            //        htTableNameToDatabase["database:" + crDatabaseName] = database;

            //        log.WriteNode("Database: ", (crDatabaseName == crServerName ? crDatabaseName : crServerName + "." + crDatabaseName));
            //    }
            //    else
            //    {
            //        htTableNameToDatabase[table.Name] = htTableNameToDatabase["database:" + crDatabaseName];
            //    }
            //    //log.WriteNode("Database: ", (crDatabaseName == crServerName ? crDatabaseName : crServerName + "." + crDatabaseName));
            //}
            #endregion

            #region for CrystalReport v11 managed 2.7
            foreach (Table table in crystalReport.Database.Tables)
            {
                #region Get DataBase name
                string DataBaseName = table.LogOnInfo.ConnectionInfo.DatabaseName;
                string ServerName = table.LogOnInfo.ConnectionInfo.ServerName;
                string DataBaseDll = null;
                string QEDataBaseName = null;
                string QEDataBaseType = null;
                string QEServerDescription = null;
                Hashtable htLogon = new Hashtable();

                foreach (NameValuePair2 nameValue in table.LogOnInfo.ConnectionInfo.Attributes.Collection)
                {
                    #region Get Connection Info Attributes
                    switch (nameValue.Name as string)
                    {
                        case "Database DLL":
                            DataBaseDll = (string)nameValue.Value;
                            break;

                        case "QE_DatabaseName":
                            QEDataBaseName = (string)nameValue.Value;
                            break;

                        case "QE_DatabaseType":
                            QEDataBaseType = (string)nameValue.Value;
                            break;

                        case "QE_LogonProperties":
                            DbConnectionAttributes logonAttributes = (DbConnectionAttributes)nameValue.Value;
                            foreach (NameValuePair2 nameValue2 in logonAttributes.Collection)
                            {
                                string name = (nameValue2.Name as string).Trim();
                                htLogon[name] = nameValue2.Value;
                            }
                            break;

                        case "QE_ServerDescription":
                            QEServerDescription = (string)nameValue.Value;
                            break;
                    }
                    #endregion
                }

                if ((DataBaseName == null) || (DataBaseName.Length == 0)) DataBaseName = QEDataBaseName;
                if ((DataBaseName == null) || (DataBaseName.Length == 0)) DataBaseName = ServerName;
                if ((DataBaseName == null) || (DataBaseName.Length == 0)) DataBaseName = QEServerDescription;
                if ((DataBaseName == null) || (DataBaseName.Length == 0)) DataBaseName = table.Name;
                #endregion

                StiDatabase database = null;
                if (htDataBaseNames[DataBaseName] == null)
                {
                    #region Make new database
                    string databaseName = DataBaseName;
                    string databaseConnectionString = null;
                    switch (QEDataBaseType)
                    {
                        case "Access/Excel (DAO)":
                            databaseConnectionString = string.Format(
                                "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};User Id={1};Password={2};",
                                ServerName,
                                table.LogOnInfo.ConnectionInfo.UserID,
                                table.LogOnInfo.ConnectionInfo.Password);
                            databaseName = System.IO.Path.GetFileName(ServerName).Replace('.', '_');
                            database = new StiOleDbDatabase(databaseName, databaseConnectionString);
                            //DataBaseName = databaseName;
                            break;

                        case "OLE DB (ADO)":
                            string provider = "Unknown";
                            if (htLogon.ContainsKey("Provider")) provider = (string)htLogon["Provider"];
                            if (provider == "SQLOLEDB")
                            {
                                databaseConnectionString = string.Format(
                                    "Provider={0};Server={1};Database={2};User Id={3};Password={4};",
                                    provider,
                                    ServerName,
                                    DataBaseName,
                                    table.LogOnInfo.ConnectionInfo.UserID,
                                    table.LogOnInfo.ConnectionInfo.Password);
                                databaseName = DataBaseName;
                            }
                            else
                            {
                                databaseConnectionString = string.Format(
                                    "Provider={0};Data Source={1};User Id={2};Password={3};",
                                    provider,
                                    ServerName,
                                    table.LogOnInfo.ConnectionInfo.UserID,
                                    table.LogOnInfo.ConnectionInfo.Password);
                                databaseName = System.IO.Path.GetFileName(ServerName).Replace('.', '_');
                            }
                            database = new StiOleDbDatabase(databaseName, databaseConnectionString);
                            //DataBaseName = databaseName;
                            break;

                        case "ADO.NET":
                            string filePath = string.Empty;
                            if (htLogon.ContainsKey("File Path")) filePath = (string)htLogon["File Path"];
                            databaseName = System.IO.Path.GetFileName(ServerName).Replace('.', '_');
                            if (filePath.EndsWith(".xsd"))
                            {
                                database = new StiXmlDatabase(
                                    databaseName,
                                    string.Empty,
                                    filePath);
                            }
                            break;

                        case "ADO.NET (XML)":
                            string filePath2 = string.Empty;
                            if (htLogon.ContainsKey("File Path")) filePath2 = (string)htLogon["File Path"];
                            databaseName = System.IO.Path.GetFileName(ServerName).Replace('.', '_');
                            database = new StiXmlDatabase(
                                databaseName,
                                string.Empty,
                                filePath2);
                            break;

                        case "XML":
                            string dataPath = string.Empty;
                            string schemaPath = string.Empty;
                            if (htLogon.ContainsKey("Local XML File")) dataPath = (string)htLogon["Local XML File"];
                            if (htLogon.ContainsKey("Local Schema File")) schemaPath = (string)htLogon["Local Schema File"];
                            databaseName = System.IO.Path.GetFileName(ServerName).Replace('.', '_');
                            database = new StiXmlDatabase(
                                databaseName,
                                schemaPath,
                                dataPath);
                            break;

                        //    //database SQL
                        //    if ((crServerName != crDatabaseName) && (table.Location.ToLower().IndexOf(".dbo.") != -1))
                        //    {
                        //        databaseConnectionString = string.Format(
                        //            "Server={0};Database={1};User ID={2};Password={3};Trusted_Connection=True;",
                        //            crServerName,
                        //            crDatabaseName,
                        //            table.LogOnInfo.ConnectionInfo.UserID,
                        //            table.LogOnInfo.ConnectionInfo.Password);
                        //        databaseName = crDatabaseName;
                        //        database = new StiSqlDatabase(databaseName, databaseConnectionString);
                        //    }
                    }
                    if (database == null)
                    {
                        databaseConnectionString = string.Format(
                            "Provider=Unknown;Data Source={0};User Id={1};Password={2};",
                            ServerName,
                            table.LogOnInfo.ConnectionInfo.UserID,
                            table.LogOnInfo.ConnectionInfo.Password);
                        databaseName = System.IO.Path.GetFileName(ServerName).Replace('.', '_');
                        database = new StiOleDbDatabase(databaseName, databaseConnectionString);
                        warningList.Add(StiWarning.GetWarningDatabase(databaseName));
                    }
                    #endregion

                    #region Add database to dictionary
                    bool founded = false;
                    foreach (StiDatabase tempDataBase in stimulReport.Dictionary.Databases)
                    {
                        if (tempDataBase.Name == databaseName) founded = true;
                    }
                    if (!founded)
                    {
                        stimulReport.Dictionary.Databases.Add(database);
                    }
                    htDataBaseNames[DataBaseName] = DataBaseName;
                    #endregion

                    htTableNameToDatabase[table.Name] = database;
                    htTableNameToDatabase["database:" + DataBaseName] = database;
                    log.WriteNode("Database: ", DataBaseName);
                }
                else
                {
                    database = (StiDatabase)htTableNameToDatabase["database:" + DataBaseName];
                    htTableNameToDatabase[table.Name] = database;
                }

                #region Make DataSource
                string tableNameValidated = (table.Name == StiNameValidator.CorrectName(table.Name) ? table.Name : "\"" + table.Name + "\"");
                StiDataSource dataSource = null;
                if ((htTableNameToDatabase[table.Name] as StiOleDbDatabase) != null)
                {
                    string nameInSource = database.Name;
                    dataSource = new StiOleDbSource(nameInSource, StiNameValidator.CorrectName(table.Name), table.Name);
                    (dataSource as StiOleDbSource).SqlCommand = string.Format("select * from {0}", tableNameValidated);
                }
                if ((htTableNameToDatabase[table.Name] as StiXmlDatabase) != null)
                {
                    string nameInSource = database.Name + "." + table.Name;
                    dataSource = new StiDataTableSource(nameInSource, StiNameValidator.CorrectName(table.Name), table.Name);
                }
                if (dataSource == null)
                {
                    string nameInSource = database.Name + "." + table.Name;
                    dataSource = new StiSqlSource(nameInSource, StiNameValidator.CorrectName(table.Name), table.Name);
                    (dataSource as StiSqlSource).SqlCommand = string.Format("select * from {0}", tableNameValidated);
                }

                foreach (DatabaseFieldDefinition field in table.Fields)
                {
                    //dataSource.Columns.Add(FieldNameCorrection(field.Name), field.Name, typeof(object));
                    StiDataColumn column = new StiDataColumn(field.Name, FieldNameCorrection(field.Name), field.Name, ConvertType(field.ValueType));
                    dataSource.Columns.Add(column);
                }

                //stimulReport.Dictionary.DataSources.Add(dataSource);
                bool founded2 = false;
                foreach (StiDataSource tempDataSource in stimulReport.Dictionary.DataSources)
                {
                    if (tempDataSource.Name == dataSource.Name) founded2 = true;
                }
                if (!founded2)
                {
                    stimulReport.Dictionary.DataSources.Add(dataSource);
                    log.WriteNode("Datasource: ", dataSource.Name);
                }
                dataSourcesOnThisPage.Add(dataSource);

                htTableNameToDataSource[table.Name] = dataSource;
                #endregion
            }
            #endregion

            log.CloseNode();

            #endregion

            #region DataSources for v9 and v10
            //log.OpenNode("Data Sources");
            //htTableNameToDataSource = new Hashtable();
            //foreach (Table table in crystalReport.Database.Tables)
            //{
            //    string dataName = (string)htDataBaseNames[table.LogOnInfo.ConnectionInfo.DatabaseName];
            //    if (dataName == null) dataName = string.Empty;

            //    StiSqlSource dataSource = null;
            //    if ((htTableNameToDatabase[table.Name] as StiOleDbDatabase) != null)
            //    {
            //        dataSource = new StiOleDbSource(dataName, StiNameValidator.CorrectName(table.Name), table.Name);
            //    }
            //    else
            //    {
            //        dataSource = new StiSqlSource(dataName, StiNameValidator.CorrectName(table.Name), table.Name);
            //    }
            //    if (StiNameValidator.CorrectName(table.Name) == table.Name)
            //    {
            //        dataSource.SqlCommand = string.Format("select * from {0}", table.Name);
            //    }
            //    else
            //    {
            //        dataSource.SqlCommand = string.Format("select * from \"{0}\"", table.Name);
            //    }

            //    foreach (DatabaseFieldDefinition field in table.Fields)
            //    {
            //        //dataSource.Columns.Add(FieldNameCorrection(field.Name), field.Name, typeof(object));
            //        StiDataColumn column = new StiDataColumn(field.Name, FieldNameCorrection(field.Name), field.Name, typeof(object));
            //        dataSource.Columns.Add(column);
            //    }

            //    //stimulReport.Dictionary.DataSources.Add(dataSource);
            //    bool founded = false;
            //    foreach (StiDataSource tempDataSource in stimulReport.Dictionary.DataSources)
            //    {
            //        if (tempDataSource.Name == dataSource.Name) founded = true;
            //    }
            //    if (!founded)
            //    {
            //        stimulReport.Dictionary.DataSources.Add(dataSource);
            //    }

            //    htTableNameToDataSource[table.Name] = dataSource;
            //    log.WriteNode("Data Source: ", table.Name);
            //}
            //log.CloseNode();
            #endregion

            #region DataRelations
            log.OpenNode("Data Relations");
            htDataSourceParent = new Hashtable();
            foreach (TableLink link in crystalReport.Database.Links)
            {
                StiDataRelation relation = new StiDataRelation();
                relation.Alias = string.Format("{0} to {1}", link.SourceTable.Name, link.DestinationTable.Name);
                relation.Name = ReplaceSymbols(link.DestinationTable.Name);
                relation.NameInSource = ReplaceSymbols(link.DestinationTable.Name);
                relation.ParentSource = (StiDataSource)htTableNameToDataSource[link.DestinationTable.Name];
                relation.ParentColumns = new string[link.DestinationFields.Count];
                for (int fieldIndex = 0; fieldIndex < link.DestinationFields.Count; fieldIndex++)
                {
                    relation.ParentColumns[fieldIndex] = link.DestinationFields[fieldIndex].Name;
                }
                relation.ChildSource = (StiDataSource)htTableNameToDataSource[link.SourceTable.Name];
                relation.ChildColumns = new string[link.SourceFields.Count];
                for (int fieldIndex = 0; fieldIndex < link.SourceFields.Count; fieldIndex++)
                {
                    relation.ChildColumns[fieldIndex] = link.SourceFields[fieldIndex].Name;
                }
                stimulReport.Dictionary.Relations.Add(relation);
                htDataSourceParent[relation.ParentSource] = relation.ParentSource;

                log.WriteNode("Data Relation: ", link.DestinationTable.Name);
            }
            log.CloseNode();
            #endregion

            #region Prepare table for convert relations
            mainDataSource = null;
            htDataSourceToNameWithRelation = new Hashtable();
            //foreach (StiDataSource dataSource in stimulReport.Dictionary.DataSources)
            foreach (StiDataSource dataSource in dataSourcesOnThisPage)
            {
                if (!htDataSourceParent.ContainsKey(dataSource))
                {
                    if (mainDataSource == null)
                    {
                        mainDataSource = dataSource;
                    }
                    MakeTableNameRecursive(dataSource, dataSource.Name);
                }
            }

            htFieldNameConversion = new Hashtable();
            foreach (Table table in crystalReport.Database.Tables)
            {
                StiDataSource ds = (StiDataSource)htTableNameToDataSource[table.Name];
                string dsRelationName = (string)htDataSourceToNameWithRelation[ds];
                foreach (DatabaseFieldDefinition field in table.Fields)
                {
                    string fieldNameCR = field.TableName + "." + field.Name;
                    string fieldNameSR = dsRelationName + FieldNameCorrection(field.Name);
                    htFieldNameConversion[fieldNameCR] = fieldNameSR;
                }
            }
            #endregion

            #region SortFields
            log.OpenNode("Sort Fields");
            ArrayList sortFieldsArray = new ArrayList();
            foreach (SortField field in crystalReport.DataDefinition.SortFields)
            {
                string fieldNameCR = field.Field.FormulaName.Substring(1, field.Field.FormulaName.Length - 2);
                string fieldNameSR = (string)htFieldNameConversion[fieldNameCR];
                if (fieldNameSR != null)
                {
                    string[] parts = fieldNameSR.Split(new char[] { '.' });
                    if (parts.Length > 1)
                    {
                        for (int indexPart = 0; indexPart < parts.Length - 1; indexPart++)
                        {
                            foreach (StiDataRelation relation in stimulReport.Dictionary.Relations)
                            {
                                if (parts[indexPart] == relation.Name)
                                {
                                    parts[indexPart] = relation.NameInSource;
                                    break;
                                }
                            }
                        }
                    }
                    string sortDirection = "ASC";
                    if (field.SortDirection == SortDirection.DescendingOrder)
                    {
                        sortDirection = "DESC";
                    }
                    parts[0] = sortDirection;
                    for (int indexPart = 0; indexPart < parts.Length; indexPart++)
                    {
                        sortFieldsArray.Add(parts[indexPart]);
                    }
                }
                log.WriteNode("Sort Field: ", fieldNameCR);
            }
            if (sortFieldsArray.Count > 0)
            {
                sortFields = new string[sortFieldsArray.Count];
                for (int indexField = 0; indexField < sortFieldsArray.Count; indexField++)
                {
                    sortFields[indexField] = (string)sortFieldsArray[indexField];
                }
            }
            log.CloseNode();
            #endregion

            //textConversion = new DictionaryEntry[htFieldNameConversion.Count];
            //htFieldNameConversion.CopyTo(textConversion, 0);
            textConversion = new Hashtable(htFieldNameConversion);
            flagTextConversionSorted = false;

            #region Fields
            htVariableConversion = new Hashtable();

			log.OpenNode("Formula Fields");
            foreach (FormulaFieldDefinition formula in crystalReport.DataDefinition.FormulaFields)
            {
                string formulaName = formula.FormulaName;
                if (formulaName != null && formulaName.Trim().Length > 0)
                {
                    formulaName = formulaName.Substring(1, formulaName.Length - 2);
                    string formulaNameNormalized = ReplaceSymbols(formulaName).Substring(1);
                    string value = formula.Text != null ? ReplaceFieldsAndVariables(formula.Text) : string.Empty;
                    value = value.Replace("\r", "");

                    StiVariable variable = new StiVariable("Formula Fields");
                    variable.Name = formulaNameNormalized;
                    variable.Alias = formulaNameNormalized;
                    variable.Type = ConvertType(formula.ValueType);
                    variable.ReadOnly = true;
                    if (useFunctions)
                    {
                        //value = ReplaceFieldsAndVariables(value);
                        CheckFunction(variable, value);
                        warningList.Add(StiWarning.GetWarningFormulaField(formulaName));
                    }
                    else
                    {
                        value = value.Replace("\n", " ");

                        variable.InitBy = StiVariableInitBy.Value;
                        variable.Type = typeof(string);
                        variable.Value = value;
                    }
                    stimulReport.Dictionary.Variables.Add(variable);

                    //htVariableConversion.Add(formulaName, ReplaceSymbols(formulaName));
                    htVariableConversion[formulaName] = formulaNameNormalized;
                    textConversion[formulaName] = formulaNameNormalized; 

                    log.WriteNode("Formula Field: ", formulaName);
                }
            }
            log.CloseNode();

            log.OpenNode("Running Total Fields");
            foreach (RunningTotalFieldDefinition formula in crystalReport.DataDefinition.RunningTotalFields)
            {
                string formulaName = formula.FormulaName;
                if (formulaName != null && formulaName.Trim().Length > 0)
                {
                    formulaName = formulaName.Substring(1, formulaName.Length - 2);
                    string formulaNameNormalized = ReplaceSymbols(formulaName).Substring(1);
                    string value = formula.SummarizedField.FormulaName != null ? formula.SummarizedField.FormulaName : string.Empty;
                    value = value.Substring(1, value.Length - 2);
                    //value = ReplaceFieldsAndVariables(value.Replace("\r", "").Replace("\n", ""));
                    string stValue = value.Replace("\r\n", " ").Replace("\r", " ").Replace("\n", " ");
                    bool flag = true;
                    if (htFieldNameConversion.ContainsKey(stValue))
                    {
                        value = (string)htFieldNameConversion[stValue];
                        value = "Totals.Sum(" + value.Substring(0, value.IndexOf('.')) + ", " + value + ")";
                    }
                    else
                    {
                        value = (string)textConversion[stValue];
                        value = "Totals.Sum(" + value + ")";
                        flag = false;
                    }

                    StiVariable variable = new StiVariable("RunningTotal Fields");
                    variable.Name = formulaNameNormalized;
                    variable.Alias = formulaNameNormalized;
                    if (flag)
                    {
                        variable.Type = ConvertType(formula.ValueType);
                        variable.ReadOnly = true;
                        variable.InitBy = StiVariableInitBy.Expression;
                    }
                    else
                    {
                        variable.ReadOnly = false;
                        variable.InitBy = StiVariableInitBy.Value;
                        variable.Type = typeof(string);
                    }
                    variable.Value = value;
                    stimulReport.Dictionary.Variables.Add(variable);

                    //htVariableConversion.Add(formulaName, ReplaceSymbols(formulaName));
                    htVariableConversion[formulaName] = formulaNameNormalized;
                    textConversion[formulaName] = formulaNameNormalized; 

                    log.WriteNode("Running Total: ", formulaName);
                }
            }
            log.CloseNode();

            log.OpenNode("Parameter Fields");
            foreach (ParameterFieldDefinition parameter in crystalReport.DataDefinition.ParameterFields)
            {
                StiVariable variable = new StiVariable("ParameterField");
                string parameterName = ReplaceSymbols(parameter.ParameterFieldName);
                variable.Name = parameterName;
                variable.Alias = parameterName;
                switch (parameter.ParameterValueKind)
                {
                    case ParameterValueKind.BooleanParameter:
                        variable.Type = typeof(bool);
                        break;
                    case ParameterValueKind.DateParameter:
                    case ParameterValueKind.DateTimeParameter:
                    case ParameterValueKind.TimeParameter:
                        variable.Type = typeof(DateTime);
                        break;
                    case ParameterValueKind.NumberParameter:
                        variable.Type = typeof(double);
                        break;
                    case ParameterValueKind.StringParameter:
                        variable.Type = typeof(string);
                        break;
                    default:
                        variable.Type = typeof(string);
                        break;
                }
                variable.ReadOnly = false;
                variable.InitBy = StiVariableInitBy.Value;
                variable.RequestFromUser = true;
                //variable.Value = string.Empty;
                stimulReport.Dictionary.Variables.Add(variable);

                //htVariableConversion.Add("?" + parameter.ParameterFieldName, parameterName);
                htVariableConversion["?" + parameter.ParameterFieldName] = parameterName;

                log.WriteNode("Parameter Field: ", parameterName);
            }
            log.CloseNode();

            AddRangeToHashtable(textConversion, htVariableConversion);
            flagTextConversionSorted = false;

            foreach (StiVariable var in stimulReport.Dictionary.Variables)
            {
                if (var.Category == "Formula Fields" && useFunctions)
                {
                    //var.Value = ReplaceFieldsAndVariables(var.Value);
                    foreach (DictionaryEntry de in htVariableConversion)
                    {
                        if (var.Value.Contains((string)de.Key)) var.Value = var.Value.Replace((string)de.Key, (string)de.Value);
                    }
                }
            }
            #endregion

            //textConversion = new DictionaryEntry[htFieldNameConversion.Count + htVariableConversion.Count];
            //htFieldNameConversion.CopyTo(textConversion, 0);
            //htVariableConversion.CopyTo(textConversion, htFieldNameConversion.Count);

            #region Groups conditions
            groupFields = new ArrayList();
            foreach (Group group in crystalReport.DataDefinition.Groups)
            {
                DictionaryEntry dicEntry = new DictionaryEntry();
                if (group != null)
                {
                    if (group.ConditionField.Kind == FieldKind.DatabaseField)
                    {
                        //string field = FieldNameCorrection(group.ConditionField.FormulaName.Substring(1, group.ConditionField.FormulaName.Length - 2));
                        string field = group.ConditionField.FormulaName.Substring(1, group.ConditionField.FormulaName.Length - 2);
                        dicEntry.Key = field;
                        //dicEntry.Value = "{" + (string)htFieldNameConversion[field] + "}";
                        dicEntry.Value = null;
                        groupFields.Add(dicEntry);
                        continue;
                    }
                    else if (group.ConditionField.Kind == FieldKind.FormulaField)
                    {
                        string field = group.ConditionField.FormulaName.Substring(1, group.ConditionField.FormulaName.Length - 2);
                        dicEntry.Key = field;
                        //dicEntry.Value = "{" + (string)htFieldNameConversion[field] + "}";
                        dicEntry.Value = null;
                        groupFields.Add(dicEntry);
                        continue;
                    }
                }
                dicEntry.Key = null;
                dicEntry.Value = null;
                groupFields.Add(dicEntry);
            }
            groupFieldsCounter = 0;
            #endregion

            #region Subreports
            if (!isSubReport)
            {
                subreportsList = new ArrayList();
                if ((crystalReport.Subreports != null) && (crystalReport.Subreports.Count > 0))
                {
                    for (int indexSubreport = 0; indexSubreport < crystalReport.Subreports.Count; indexSubreport++)
                    {
                        subreportsList.Add(crystalReport.Subreports[indexSubreport].Name);
                    }
                }
            }
            #endregion

            summaryFieldsList = new ArrayList();

            log.OpenNode("Bands");
            AddBands(page, crystalReport, mainDataSource, log);
            log.CloseNode();

            #region Fill SummaryFields with GroupBands names
            if (summaryFieldsList.Count > 0)
            {
                foreach (DictionaryEntry dicEntry in summaryFieldsList)
                {
                    for (int indexGroup = 0; indexGroup < groupFields.Count; indexGroup++)
                    {
                        DictionaryEntry dicEntryGroup = (DictionaryEntry)groupFields[indexGroup];
                        if ((string)dicEntry.Value == (string)dicEntryGroup.Key)
                        {
                            StiComponent component = (StiComponent)dicEntry.Key;
                            StiText text = component as StiText;
                            text.Text = text.Text.Value.Replace("{},", (string)dicEntryGroup.Value + ",");
                        }
                    }
                }
            }
            #endregion

            stimulReport.Info.ForceDesigningMode = true;
            page.DockToContainer();
            page.Correct();
            stimulReport.Info.ForceDesigningMode = false;

        }

        private void AddBands(StiPage page, ReportDocument crystalReport, StiDataSource mainDataSource, IStiTreeLog log)
		{
			double bandPosition = 0;
            for (int sectionIndex = 0; sectionIndex < crystalReport.ReportDefinition.Sections.Count; sectionIndex++)
			{
                Section section = crystalReport.ReportDefinition.Sections[sectionIndex];
                int subSectionIndex = GetSubSectionIndex(section);

				StiBand band = TypeSection(section);
				if (band != null)
				{
                    if (band is StiGroupHeaderBand || band is StiGroupFooterBand)
                    {
                        int subGroupIndex = GetSubSectionIndex(section, true);
                        bool nextSect = CheckNextSectionIsSecondPart(sectionIndex, crystalReport) || subSectionIndex > 0;
                        band.Name = CheckComponentsNames(string.Format("{0}{1}{2}", band.Name, subGroupIndex + 1, nextSect ? "" + (char)((int)'a' + subSectionIndex) : ""));
                    }
                    else
                    {
                        band.Name = CheckComponentsNames(string.Format("{0}{1}", band.Name, subSectionIndex + 1));
                    }

                    band.Height = section.Height * unit;
                    band.Brush = new StiSolidBrush(section.SectionFormat.BackgroundColor);

                    #region Suppress
                    //band.Enabled = !section.SectionFormat.EnableSuppress;

                    bool hasSuppressExpression = false;
                    object objCondition = GetPropertyFromHiddenObject(section.SectionFormat, "RasSectionFormat.ConditionFormulas");
                    if (objCondition != null)
                    {
                        var formulaClass = objCondition as CrystalDecisions.ReportAppServer.ReportDefModel.SectionAreaFormatConditionFormulasClass;
                        var conditionFormula = formulaClass[CrystalDecisions.ReportAppServer.ReportDefModel.CrSectionAreaFormatConditionFormulaTypeEnum.crSectionAreaConditionFormulaTypeEnableSuppress];
                        if ((conditionFormula != null) && !string.IsNullOrWhiteSpace(conditionFormula.Text))
                        {
                            var text = ReplaceFieldsAndVariables(conditionFormula.Text);
                            var tempVariable = new StiVariable();
                            CheckFunction(tempVariable, text);

                            StiCondition condition = new StiCondition();
                            condition.Item = StiFilterItem.Expression;
                            condition.Expression = tempVariable.Value;
                            band.Conditions.Add(condition);

                            hasSuppressExpression = true;
                        }
                    }

                    if (section.SectionFormat.EnableSuppress && !hasSuppressExpression)
                    {
                        band.Brush = new StiHatchBrush(HatchStyle.BackwardDiagonal, Color.DarkGray, Color.White);
                        band.Enabled = false;
                    }

                    if (hasSuppressExpression || !band.Enabled)
                    {
                        if (band is StiDataBand)
                        {
                            (band as StiDataBand).CalcInvisible = true;
                        }
                    }
                    #endregion

                    band.Parent = page;
                    band.Page = page;
					band.Top = bandPosition;
                    if (section.SectionFormat.EnableNewPageBefore)
                    {
                        if ((band as IStiPageBreak) != null)
                        {
                            (band as IStiPageBreak).NewPageBefore = true;
                            (band as IStiPageBreak).SkipFirst = false;
                        }
                    }
                    if (section.SectionFormat.EnableNewPageAfter)
                    {
                        if ((band as IStiPageBreak) != null) (band as IStiPageBreak).NewPageAfter = true;
                    }
                    if (section.SectionFormat.EnablePrintAtBottomOfPage)
                    {
                        if ((band as IStiPrintAtBottom) != null) (band as IStiPrintAtBottom).PrintAtBottom = true;
                    }

					//bandPosition += band.Height + 2;
                    bandPosition += band.Height;

					if ((band is StiDataBand) && (mainDataSource != null))
					{
						StiDataBand dataBand = band as StiDataBand;
						dataBand.DataSourceName = mainDataSource.Name;
						dataBand.Sort = sortFields;
					}
                    if (band is StiGroupHeaderBand || band is StiGroupFooterBand || band is StiDataBand)
                    {
                        //bool needChild = false;
                        if (subSectionIndex == 0)
                        {
                            if (band is StiGroupHeaderBand)
                            {
                                #region Set condition
                                int conditionIndex = GetSubSectionIndex(section, true);
                                if (conditionIndex < groupFields.Count)
                                {
                                    StiGroupHeaderBand groupHeaderBand = band as StiGroupHeaderBand;
                                    groupHeaderBand.SortDirection = StiGroupSortDirection.Ascending;
                                    DictionaryEntry de = (DictionaryEntry)groupFields[conditionIndex];
                                    string field = "{}";
                                    if (de.Key != null)
                                    {
                                        if (htFieldNameConversion[(string)de.Key] != null)
                                        {
                                            field = "{" + (string)htFieldNameConversion[(string)de.Key] + "}";
                                        }
                                        else if (htVariableConversion[(string)de.Key] != null)
                                        {
                                            field = "{" + (string)htVariableConversion[(string)de.Key] + "}";
                                        }
                                        else
                                        {
                                            field = (string)de.Key;
                                        }
                                    }
                                    groupHeaderBand.Condition.Value = field;
                                    if (de.Value == null)
                                    {
                                        de.Value = band.Name;
                                        groupFields[groupFieldsCounter] = de;
                                    }
                                    groupFieldsCounter = conditionIndex + 1;
                                }
                                #endregion
                            }

                            if (CheckNextSectionIsSecondPart(sectionIndex, crystalReport))
                            {
                                StiChildBand child = new StiChildBand();
                                child.Name = band.Name;
                                child.Height = band.Height;
                                child.Brush = band.Brush;
                                child.Parent = band.Parent;
                                child.Page = band.Page;
                                child.Top = band.Top;

                                band.Height = 0;
                                band.Brush = new StiEmptyBrush();
                                band.Name = CheckComponentsNames(band.Name.Substring(0, band.Name.Length - 1));
                                page.Components.Add(band);
                                band = child;
                            }
                        }
                        else
                        {
                            StiChildBand child = new StiChildBand();
                            child.Name = band.Name;
                            child.Height = band.Height;
                            child.Brush = band.Brush;
                            child.Parent = band.Parent;
                            child.Page = band.Page;
                            child.Top = band.Top;

                            band = child;

                            if ((band as IStiPageBreak) != null)
                            {
                                child.NewPageBefore = (band as IStiPageBreak).NewPageBefore;
                                child.SkipFirst = (band as IStiPageBreak).SkipFirst;
                            }
                            if ((band as IStiPageBreak) != null) child.NewPageAfter = (band as IStiPageBreak).NewPageAfter;
                            if ((band as IStiPrintAtBottom) != null) child.PrintAtBottom = (band as IStiPrintAtBottom).PrintAtBottom;
                        }
                    }

					page.Components.Add(band);

                    log.OpenNode(band.Name);

					AddComponent(band, section, log);

                    log.CloseNode();
				}
			}
		}

        private int GetSubSectionIndex(Section section, bool group = false)
        {
            int sectionCode = -1;
            object obj1 = GetPropertyFromHiddenObject(section, "RasSection.SectionCode");
            if (obj1 != null)
            {
                sectionCode = (int)obj1;
            }

            while (sectionCode >= 1000)
            {
                sectionCode -= 1000;
            }

            int subSectionIndex = sectionCode / 25;
            int groupIndex = sectionCode % 25;

            return group ? groupIndex : subSectionIndex;
        }

        private bool CheckNextSectionIsSecondPart(int sectionIndex, ReportDocument crystalReport)
        {
            if (sectionIndex + 1 >= crystalReport.ReportDefinition.Sections.Count) return false;

            Section section1 = crystalReport.ReportDefinition.Sections[sectionIndex];
            int sectionCode1 = -1;
            object obj1 = GetPropertyFromHiddenObject(section1, "RasSection.SectionCode");
            if (obj1 != null)
            {
                sectionCode1 = (int)obj1;
            }

            Section section2 = crystalReport.ReportDefinition.Sections[sectionIndex + 1];
            int sectionCode2 = -1;
            object obj2 = GetPropertyFromHiddenObject(section2, "RasSection.SectionCode");
            if (obj2 != null)
            {
                sectionCode2 = (int)obj2;
            }

            if (sectionCode1 / 1000 != sectionCode2 / 1000) return false;
            while (sectionCode1 >= 1000) sectionCode1 -= 1000;
            while (sectionCode2 >= 1000) sectionCode2 -= 1000;

            return (sectionCode1 / 25 + 1) == (sectionCode2 / 25);
        }

		
		private void AddComponent(StiBand band, Section section, IStiTreeLog log)
		{
			foreach (ReportObject obj in section.ReportObjects)
			{
				StiComponent comp = TypeComponent(obj);
				if (comp != null)
				{
                    comp.Name = CheckComponentsNames(obj.Name);

                    if ((comp is StiVerticalLinePrimitive) || (comp is StiRectanglePrimitive))
                    {
                        StiStartPointPrimitive start = new StiStartPointPrimitive();
                        start.Left = comp.Left;
                        start.Top = comp.Top;
                        start.ReferenceToGuid = comp.Guid;
                        start.Parent = band;
                        band.Components.Add(start);
                        start.Linked = true;

                        StiEndPointPrimitive end = new StiEndPointPrimitive();
                        end.Left = comp.Right;
                        end.Top = comp.Bottom;
                        end.ReferenceToGuid = comp.Guid;
                        end.Parent = band;
                        band.Components.Add(end);
                        end.Linked = true;

                        comp.Top += band.Top;
                        comp.Parent = band.Page;
                        band.Page.Components.Add(comp);
                    }
                    else
                    {
                        comp.Parent = band;
                        band.Components.Add(comp);
                    }

                    comp.Linked = true;

                    log.WriteNode(comp.Name);
				}
			}
        }
        #endregion

        #region Get types
        private static StiBand TypeSection(Section s)
		{
			switch (s.Kind)
			{
				case AreaSectionKind.ReportHeader:
				{
                    return new StiReportTitleBand() { Name = "ReportHeader" };
				}
				case AreaSectionKind.PageHeader:
				{
                    return new StiPageHeaderBand() { Name = "PageHeader" };
				}
				case AreaSectionKind.GroupHeader:
				{
                    return new StiGroupHeaderBand() { Name = "GroupHeader" };
				}
				case AreaSectionKind.GroupFooter:
				{
                    return new StiGroupFooterBand() { Name = "GroupFooter" };
				}
				case AreaSectionKind.PageFooter:
				{
                    return new StiPageFooterBand() { Name = "PageFooter" };
				}
				case AreaSectionKind.ReportFooter:
				{
                    //return new StiFooterBand() { Name = "ReportFooter" };
                    return new StiReportSummaryBand() { Name = "ReportFooter" };
				}
				case AreaSectionKind.Detail:
				{
                    return new StiDataBand() { Name = "Details" };
				}
			}
			return null;
		}


		private StiComponent TypeComponent(ReportObject obj)
		{
			switch (obj.Kind)
			{
				case ReportObjectKind.FieldObject:
				{
					return ImportFieldObject(obj);
				}
				case ReportObjectKind.TextObject:
				{
					return ImportTextObject(obj);
				}
				case ReportObjectKind.LineObject:
				{
					return ImportLineObject(obj);
				}
				case ReportObjectKind.BoxObject:
				{
					return ImportBoxObject(obj);
				}
				case ReportObjectKind.PictureObject:
				{
					return ImportPictureObject(obj);
				}
                case ReportObjectKind.FieldHeadingObject:
                {
                    return ImportTextObject(obj);
                }
                case ReportObjectKind.SubreportObject:
                {
                    return ImportSubreportObject(obj);
                }
				default:
				{
					return null;
				}						
			}
		}

        [HandleProcessCorruptedStateExceptions]
        private string TypeField(FieldObject obj, StiComponent component)
		{
            try
            {
                switch (obj.DataSource.Kind)
                {
                    case FieldKind.DatabaseField:
                        {
                            return GetDataBaseField(obj);
                        }
                    case FieldKind.FormulaField:
                        {
                            return GetFormulaField(obj);
                        }
                    case FieldKind.GroupNameField:
                        {
                            return GetGroupNameField(obj);
                        }
                    case FieldKind.ParameterField:
                        {
                            return GetParameterField(obj);
                        }
                    case FieldKind.RunningTotalField:
                        {
                            return GetRunningTotalField(obj);
                        }
                    case FieldKind.SpecialVarField:
                        {
                            return GetSpecialVarField(obj);
                        }
                    case FieldKind.SQLExpressionField:
                        {
                            return GetSQLExpressionField(obj);
                        }
                    case FieldKind.SummaryField:
                        {
                            return GetSummaryField(obj, component);
                        }
                    default:
                        {
                            return null;
                        }
                }
            }
            catch
            {
                return null;
            }
        }
        #endregion

        #region Get fields
        private string GetDataBaseField(FieldObject obj)
		{
			//string field = FieldNameCorrection(obj.DataSource.FormulaName.Substring(1, obj.DataSource.FormulaName.Length - 2));
            string field = obj.DataSource.FormulaName.Substring(1, obj.DataSource.FormulaName.Length - 2);
			return "{" + (string)htFieldNameConversion[field] + "}";
		}


		private string GetFormulaField(FieldObject obj)
		{
            return "{" + ReplaceFieldsAndVariables(obj.DataSource.FormulaName.Substring(1, obj.DataSource.FormulaName.Length - 2)) + "}";
		}


		private string GetGroupNameField(FieldObject obj)
		{
            string field = string.Empty;
            string formula = obj.DataSource.FormulaName;
            int index1 = formula.IndexOf('{');
            int index2 = formula.LastIndexOf('}');
            if ((index1 != -1) && (index2 != -1))
            {
                formula = formula.Substring(index1 + 1, index2 - index1 - 1);
                //field = FieldNameCorrection(formula);
                field = formula;
            }
            return "{" + (string)htFieldNameConversion[field] + "}";
		}


		private string GetParameterField(FieldObject obj)
		{
			//return obj.DataSource.FormulaName;
            string field = obj.DataSource.FormulaName.Substring(1, obj.DataSource.FormulaName.Length - 2);
            return "{" + (string)htVariableConversion[field] + "}";
		}


		private string GetRunningTotalField(FieldObject obj)
		{
            return ReplaceFieldsAndVariables(obj.DataSource.FormulaName);
		}


		private static string GetSpecialVarField(FieldObject obj)
		{
			switch(obj.DataSource.FormulaName)
			{
				case "PageNofM":
				{
					return "{PageNofM}";
				}
				case "FileCreationDate":
				{
					return "{ReportCreated}";
				}
				case "FileAuthor":
				{
					return "{ReportAuthor}";
				}
				case "ReportTitle":
				{
					return "{ReportAlias}";
				}
				case "ReportComents":
				{
					return "{ReportDescription}";
				}
				case "TotalPageCount":
				{
					return "{TotalPageCount}";
				}
				case "PageNumber":
				{
					return "{PageNumber}";
				}
				case "DataTime":
				{
                    if (obj.FieldFormat.TimeFormat.SecondFormat == SecondFormat.NumericSecond)
                    {
                        return "{Format(\"{0:T}\", Time)}";
                    }
                    else
                    {
                        return "{Format(\"{0:t}\", Time)}";
                    }
                }
				case "DataDate":
				{
					return "{Format(\"{0:d}\", Today)}";
                }
                case "PrintTime":
                {
                    if (obj.FieldFormat.TimeFormat.SecondFormat == SecondFormat.NumericSecond)
                    {
                        return "{Format(\"{0:T}\", Time)}";
                    }
                    else
                    {
                        return "{Format(\"{0:t}\", Time)}";
                    }
                }
                case "PrintDate":
                {
                    return "{Format(\"{0:d}\", Today)}";
                }
                case "ModificationTime":
				{
                    if (obj.FieldFormat.TimeFormat.SecondFormat == SecondFormat.NumericSecond)
                    {
                        return "{Format(\"{0:T}\", ReportChanged)}";
                    }
                    else
                    {
                        return "{Format(\"{0:t}\", ReportChanged)}";
                    }
                }
				case "ModificationDate":
				{
                    return "{Format(\"{0:d}\", ReportChanged)}";
                }
				default:
				{
					return "";
				}
			}
		}


		private static string GetSQLExpressionField(FieldObject obj)
		{
			return obj.DataSource.FormulaName;
		}


		private string GetSummaryField(FieldObject obj, StiComponent component)
		{
			//return obj.DataSource.FormulaName;
            SummaryFieldDefinition sfd = obj.DataSource as SummaryFieldDefinition;
            string field = sfd.SummarizedField.FormulaName.Substring(1, sfd.SummarizedField.FormulaName.Length - 2);
            string operation = string.Empty;
            switch (sfd.Operation)
            {
                case SummaryOperation.Average:
                    operation = "Avg";
                    break;
                case SummaryOperation.Count:
                    operation = "Count";
                    break;
                case SummaryOperation.Maximum:
                    operation = "Max";
                    break;
                case SummaryOperation.Median:
                    operation = "Median";
                    break;
                case SummaryOperation.Minimum:
                    operation = "Min";
                    break;
                case SummaryOperation.Mode:
                    operation = "Mode";
                    break;
                case SummaryOperation.Sum:
                    operation = "Sum";
                    break;
                default:
                    StringBuilder opSb = new StringBuilder();
                    int pos = 0;
                    while ((pos < sfd.FormulaName.Length) && (char.IsLetter(sfd.FormulaName[pos])))
                    {
                        opSb.Append(sfd.FormulaName[pos]);
                        pos++;
                    }
                    operation = opSb.ToString();
                    break;
            }
            if (sfd.Group != null)
            {
                DictionaryEntry dicEntry = new DictionaryEntry();
                dicEntry.Key = component;
                string groupField = sfd.Group.ConditionField.FormulaName.Substring(1, sfd.Group.ConditionField.FormulaName.Length - 2);
                dicEntry.Value = groupField;
                summaryFieldsList.Add(dicEntry);
                return "{" + operation + "({}," + ReplaceFieldsAndVariables(field) + ")}";
            }
            else
            {
                return "{" + operation + "(" + ReplaceFieldsAndVariables(field) + ")}";
            }
        }
        #endregion

        #region Import objects
        private StiComponent ImportSubreportObject(ReportObject o)
        {
            StiSubReport subReport = new StiSubReport();
            SubreportObject obj = o as SubreportObject;

            int pageIndex = 0;
            for (int index = 0; index < subreportsList.Count; index++)
            {
                if (obj.SubreportName == (string)subreportsList[index])
                {
                    pageIndex = index + 1;
                }
            }
            if (pageIndex != 0)
            {
                subReport.SubReportPage = stimulReport.Pages[pageIndex];
            }

            subReport.Border = ConvertBorder(obj.Border);
            //subReport.Brush = new StiSolidBrush(obj.Border.BackgroundColor);
            subReport.Brush = new StiSolidBrush(GetBackgroundColorFromBorder(obj.Border));

            ConvertProperties(subReport, obj);

            return subReport;
        }


		private StiComponent ImportPictureObject(ReportObject o)
		{
			StiImage image = new StiImage();
			PictureObject obj = o as PictureObject;

            ConvertProperties(image, obj);

            warningList.Add(StiWarning.GetWarningPicture(obj.Name));

			//image = null;
			return image;
		}


        private StiComponent ImportTextObject(ReportObject o)
        {
            StiText text = new StiText();
            TextObject obj = o as TextObject;

            text.Border = ConvertBorder(obj.Border);
            //text.Text = ReplaceFieldsAndVariables(obj.Text.Replace("\n", ""));
            text.Text = ReplaceFieldsAndVariables(obj.Text);
            text.Font = obj.Font;

            ConvertProperties(text, obj);

            text.TextBrush = new StiSolidBrush(obj.Color);
            //text.Brush = new StiSolidBrush(obj.Border.BackgroundColor);
            text.Brush = new StiSolidBrush(GetBackgroundColorFromBorder(obj.Border));
            text.HorAlignment = ConvertTextAligment(obj.ObjectFormat.HorizontalAlignment);

            //text.WordWrap = obj.ObjectFormat.EnableCanGrow;   //??? only in old versions?
            text.CanGrow = obj.ObjectFormat.EnableCanGrow;
            text.WordWrap = true;

            text.Angle = GetTextRotationAngle(obj, "RasTextObject");

            return text;
        }


		private StiComponent ImportFieldObject(ReportObject o)
		{
			FieldObject obj = o as FieldObject;
            CrTextFormatEnum textFormat = CrTextFormatEnum.crTextFormatStandardText;
            StiComponent component = null;

            object objj = GetPropertyFromHiddenObject(obj.FieldFormat, "RasFieldFormat.StringFormat.TextFormat");
            if (objj != null)
            {
                textFormat = (CrTextFormatEnum)objj;
            }

            if (textFormat == CrTextFormatEnum.crTextFormatRTFText)
            {
                StiRichText text = new StiRichText();
                text.Border = ConvertBorder(obj.Border);

                ConvertProperties(text, obj);

                text.Text = TypeField(obj, text);

                component = text;
            }
            else
            {
                StiText text = new StiText();
                text.Border = ConvertBorder(obj.Border);
                text.Font = obj.Font;

                ConvertProperties(text, obj);

                text.TextBrush = new StiSolidBrush(obj.Color);
                //text.Brush = new StiSolidBrush(obj.Border.BackgroundColor);
                text.Brush = new StiSolidBrush(GetBackgroundColorFromBorder(obj.Border)); 
                text.HorAlignment = ConvertTextAligment(obj.ObjectFormat.HorizontalAlignment);
                text.WordWrap = obj.ObjectFormat.EnableCanGrow;

                text.Angle = GetTextRotationAngle(obj, "RasFieldObject");

                text.Text = TypeField(obj, text);

                component = text;
            }

			return component;
		}


		private StiComponent ImportLineObject(ReportObject o)
		{
            if (usePrimitives)
            {
                LineObject obj = o as LineObject;

                //float lineSize = obj.LineThickness / 20f;
                float lineSize = obj.LineThickness * unit;

                StiLinePrimitive line = null;
                if (obj.Width <= 0)
                {
                    line = new StiVerticalLinePrimitive();
                    line.Height = obj.Height * unit;
                    line.Left += lineSize / 2;
                }
                else
                {
                    line = new StiHorizontalLinePrimitive();
                    line.Width = obj.Width * unit;
                    line.Top += lineSize / 2;
                }
                line.Size = lineSize;
                line.Style = ConvertLineStyle(obj.LineStyle);
                line.Color = obj.LineColor;
                line.Left += obj.Left * unit;
                line.Top += obj.Top * unit;

                return line;
            }
            else
            {
                StiShape shape = new StiShape();
                LineObject obj = o as LineObject;

                shape.Height = obj.Height * unit;
                shape.Width = obj.Width * unit;
                shape.Left = obj.Left * unit;
                shape.Top = obj.Top * unit;

                if (obj.Width <= 0)
                {
                    shape.ShapeType = new Stimulsoft.Report.Components.ShapeTypes.StiVerticalLineShapeType();
                    shape.Width = obj.LineThickness * unit;
                }
                else
                {
                    shape.ShapeType = new Stimulsoft.Report.Components.ShapeTypes.StiHorizontalLineShapeType();
                    shape.Height = obj.LineThickness * unit;
                }

                shape.Style = ConvertLineStyle(obj.LineStyle);
                shape.BorderColor = obj.LineColor;
                shape.Size = obj.LineThickness * unit;

                return shape;
            }
		}
 

		private StiComponent ImportBoxObject(ReportObject o)
		{
            BoxObject obj = o as BoxObject;

            #region Get round value
            int valueCornerEllipseHeight = 0;
            int valueCornerEllipseWidth = 0;

            object obj1 = GetPropertyFromHiddenComObject(obj, "RasBoxObject.CornerEllipseHeight");
            if (obj1 != null)
            {
                valueCornerEllipseHeight = (int)obj1;
            }
            object obj2 = GetPropertyFromHiddenComObject(obj, "RasBoxObject.CornerEllipseWidth");
            if (obj2 != null)
            {
                valueCornerEllipseWidth = (int)obj2;
            }

            float roundValue = ((valueCornerEllipseWidth + valueCornerEllipseHeight) / 2f) * unit / (float)(obj.Width > obj.Height ? obj.Height : obj.Width);
            roundValue = Math.Min(roundValue, 0.49f);
            #endregion

            if (usePrimitives)
            {
                StiRectanglePrimitive rect = new StiRectanglePrimitive();
                if (roundValue > 0.001)
                {
                    rect = new StiRoundedRectanglePrimitive();
                    (rect as StiRoundedRectanglePrimitive).Round = roundValue;
                }
                rect.Left = obj.Left * unit;
                rect.Top = obj.Top * unit;
                rect.Width = obj.Width * unit;
                rect.Height = obj.Height * unit;
                rect.Size = obj.LineThickness / 20f;
                rect.Style = ConvertLineStyle(obj.LineStyle);
                rect.Color = obj.LineColor;

                return rect;
            }
            else
            {
                StiShape shape = new StiShape();

                shape.ShapeType = new Stimulsoft.Report.Components.ShapeTypes.StiRectangleShapeType();
                if (roundValue > 0.001)
                {
                    shape.ShapeType = new Stimulsoft.Report.Components.ShapeTypes.StiRoundedRectangleShapeType();
                    (shape.ShapeType as Stimulsoft.Report.Components.ShapeTypes.StiRoundedRectangleShapeType).Round = roundValue;
                }

                shape.Height = obj.Height * unit;
                shape.Width = obj.Width * unit;
                shape.Left = obj.Left * unit;
                shape.Top = obj.Top * unit;

                shape.BorderColor = obj.Border.BorderColor;
                shape.Style = ConvertLineStyle(obj.LineStyle);
                shape.Size = obj.LineThickness * unit;

                return shape;
            }
        }
        #endregion

        #region Converters
        private static Type ConvertType(FieldValueType valueType)
        {
            switch (valueType)
            {
                case FieldValueType.BooleanField:
                    return typeof(bool);

                case FieldValueType.CurrencyField:
                    return typeof(decimal);

                case FieldValueType.DateField:
                case FieldValueType.DateTimeField:
                    return typeof(DateTime);

                case FieldValueType.Int16sField:
                case FieldValueType.Int16uField:
                case FieldValueType.Int32sField:
                case FieldValueType.Int32uField:
                case FieldValueType.Int8sField:
                case FieldValueType.Int8uField:
                    return typeof(int);

                case FieldValueType.NumberField:
                    return typeof(double);

                case FieldValueType.StringField:
                    return typeof(string);

                case FieldValueType.BitmapField:
                case FieldValueType.IconField:
                case FieldValueType.PictureField:
                    return typeof(string);

                case FieldValueType.TimeField:
                    return typeof(TimeSpan);

                default:
                    return typeof(object);
            }
        }
        
        
        private static StiTextHorAlignment ConvertTextAligment(Alignment alignment)
        {
            switch (alignment)
            {
                case Alignment.LeftAlign:
                    return StiTextHorAlignment.Left;

                case Alignment.RightAlign:
                    return StiTextHorAlignment.Right;

                case Alignment.HorizontalCenterAlign:
                    return StiTextHorAlignment.Center;

                default:
                    return (StiTextHorAlignment)0;
            }
        }


        private static StiPenStyle ConvertLineStyle(LineStyle lineStyle)
		{
			switch (lineStyle)
			{
				case LineStyle.SingleLine:
				{
					return StiPenStyle.Solid;
				}
				case LineStyle.DotLine:
				{
					return StiPenStyle.Dot;
				}
				case LineStyle.DashLine:
				{
					return StiPenStyle.Dash;
				}
				default:
				{
					return StiPenStyle.None;
				}
			}
		}
		
		
		private static StiBorder ConvertBorder(Border border)
		{
			StiBorderSides sides = StiBorderSides.None;

			bool solid = false;

			#region Get sides
			if (ConvertLineStyle(border.TopLineStyle) != StiPenStyle.None)
			{
				sides |= StiBorderSides.Top;
				if (ConvertLineStyle(border.TopLineStyle) == StiPenStyle.Solid)
				{
					solid = true;
				}
			}
			if (ConvertLineStyle(border.RightLineStyle) != StiPenStyle.None)
			{
				sides |= StiBorderSides.Right;
				if (ConvertLineStyle(border.RightLineStyle) == StiPenStyle.Solid)
				{
					solid = true;
				}
			}
			if (ConvertLineStyle(border.BottomLineStyle) != StiPenStyle.None)
			{
				sides |= StiBorderSides.Bottom;
				if (ConvertLineStyle(border.BottomLineStyle) == StiPenStyle.Solid)
				{
					solid = true;
				}
			}
			if (ConvertLineStyle(border.LeftLineStyle) != StiPenStyle.None)
			{
				sides |= StiBorderSides.Left;
				if (ConvertLineStyle(border.LeftLineStyle) == StiPenStyle.Solid)
				{
					solid = true;
				}
			}
			#endregion

			StiBorder brd = new StiBorder();
            brd.Side = sides;

			if (solid)
			{
				brd.Style = StiPenStyle.Solid;
			}
			brd.Color = border.BorderColor;

			return brd;
		}


		private static void ConvertProperties(StiComponent comp, ReportObject obj)
		{
			comp.Height = obj.Height * unit;
			comp.Width = obj.Width * unit;
			comp.Left = obj.Left * unit;
			comp.Top = obj.Top * unit;

			comp.CanGrow = obj.ObjectFormat.EnableCanGrow;
			comp.Enabled = !(obj.ObjectFormat.EnableSuppress);
        }
        #endregion

	}
}
