#region Copyright (C) 2003-2017 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports  											}
{																	}
{	Copyright (C) 2003-2017 Stimulsoft     							}
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
#endregion Copyright (C) 2003-2017 Stimulsoft

using System;
using System.Collections;
using System.Collections.Generic;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;

namespace Stimulsoft.Report.Import
{
    public partial class StiVisualFoxProParser
    {
        #region Class StiFoxProKeywordInfo
        public class StiFoxProKeywordInfo
        {
            public string Name;
            public string StiName;

            public StiFoxProKeywordInfo(string name, string stiName)
            {
                this.Name = name;
                this.StiName = stiName;
            }

            public override string ToString()
            {
                return Name + " -> " + StiName;
            }
        }
        #endregion

        #region Properties
        private static Hashtable typesList = null;
        private static Hashtable TypesList
        {
            get
            {
                if (typesList == null)
                {
                    typesList = new Hashtable();
                    typesList["bool"] = TypeCode.Boolean;
                    typesList["Boolean"] = TypeCode.Boolean;
                    typesList["byte"] = TypeCode.Byte;
                    typesList["Byte"] = TypeCode.Byte;
                    typesList["sbyte"] = TypeCode.SByte;
                    typesList["Sbyte"] = TypeCode.SByte;
                    typesList["char"] = TypeCode.Char;
                    typesList["Char"] = TypeCode.Char;
                    typesList["decimal"] = TypeCode.Decimal;
                    typesList["Decimal"] = TypeCode.Decimal;
                    typesList["double"] = TypeCode.Double;
                    typesList["Double"] = TypeCode.Double;
                    typesList["float"] = TypeCode.Single;
                    typesList["Single"] = TypeCode.Single;
                    typesList["int"] = TypeCode.Int32;
                    typesList["uint"] = TypeCode.UInt32;
                    typesList["long"] = TypeCode.Int64;
                    typesList["ulong"] = TypeCode.UInt64;
                    typesList["short"] = TypeCode.Int16;
                    typesList["Int16"] = TypeCode.Int16;
                    typesList["Int32"] = TypeCode.Int32;
                    typesList["Int64"] = TypeCode.Int64;
                    typesList["ushort"] = TypeCode.UInt16;
                    typesList["UInt16"] = TypeCode.UInt16;
                    typesList["UInt32"] = TypeCode.UInt32;
                    typesList["UInt64"] = TypeCode.UInt64;
                    typesList["object"] = TypeCode.Object;
                    typesList["string"] = TypeCode.String;
                    typesList["String"] = TypeCode.String;
                    typesList["DateTime"] = TypeCode.DateTime;
                    typesList["DateTime"] = TypeCode.DateTime;
                }
                return typesList;
            }
        }


        public static List<StiFoxProKeywordInfo> FoxProSystemVariables = new List<StiFoxProKeywordInfo>()
        {
            new StiFoxProKeywordInfo("_pageno", "PageNumber"),
            new StiFoxProKeywordInfo("_pagetotal", "TotalPageCount")
        };

        private static Hashtable systemVariablesList = null;
        private static Hashtable SystemVariablesList
        {
            get
            {
                if (systemVariablesList == null)
                {
                    systemVariablesList = new Hashtable();

                    for (int index = 0; index < FoxProSystemVariables.Count; index++)
                    {
                        var sysVariable = FoxProSystemVariables[index];
                        systemVariablesList[sysVariable.Name.ToLowerInvariant()] = (StiSystemVariableType)(index + 1);
                    }
                }
                return systemVariablesList;
            }
        }

        private static Hashtable propertiesList = null;
        private static Hashtable PropertiesList
        {
            get
            {
                if (propertiesList == null)
                {
                    propertiesList = new Hashtable();
                    propertiesList["Year"] = StiPropertyType.Year;
                    propertiesList["Month"] = StiPropertyType.Month;
                    propertiesList["Day"] = StiPropertyType.Day;
                    propertiesList["Hour"] = StiPropertyType.Hour;
                    propertiesList["Minute"] = StiPropertyType.Minute;
                    propertiesList["Second"] = StiPropertyType.Second;
                    propertiesList["Length"] = StiPropertyType.Length;
                    propertiesList["From"] = StiPropertyType.From;
                    propertiesList["To"] = StiPropertyType.To;
                    propertiesList["FromDate"] = StiPropertyType.FromDate;
                    propertiesList["ToDate"] = StiPropertyType.ToDate;
                    propertiesList["FromTime"] = StiPropertyType.FromTime;
                    propertiesList["ToTime"] = StiPropertyType.ToTime;
                    propertiesList["SelectedLine"] = StiPropertyType.SelectedLine;
                    propertiesList["Name"] = StiPropertyType.Name;
                    propertiesList["TagValue"] = StiPropertyType.TagValue;

                    propertiesList["Days"] = StiPropertyType.Days;
                    propertiesList["Hours"] = StiPropertyType.Hours;
                    propertiesList["Milliseconds"] = StiPropertyType.Milliseconds;
                    propertiesList["Minutes"] = StiPropertyType.Minutes;
                    propertiesList["Seconds"] = StiPropertyType.Seconds;
                    propertiesList["Ticks"] = StiPropertyType.Ticks;

                    propertiesList["Count"] = StiPropertyType.Count;
                }
                return propertiesList;
            }
        }


        public static List<StiFoxProKeywordInfo> FoxProFunctions = new List<StiFoxProKeywordInfo>()
        {
            new StiFoxProKeywordInfo("ASC", "*ASC"),
            new StiFoxProKeywordInfo("ALLTRIM", "Trim"),
            new StiFoxProKeywordInfo("ALLTR", "Trim"),
            new StiFoxProKeywordInfo("ALLT", "Trim"),
            new StiFoxProKeywordInfo("AT", "*AT"),
            new StiFoxProKeywordInfo("ATC", "*ATC"),
            new StiFoxProKeywordInfo("CHR", "*CHR"),
            new StiFoxProKeywordInfo("CHRTRAN", "*CHRTRAN"),
            new StiFoxProKeywordInfo("CTOBIN", "*CTOBIN"),
            new StiFoxProKeywordInfo("CURSORTOXML", "*CURSORTOXML"),
            new StiFoxProKeywordInfo("CURVAL", "*CURVAL"),
            new StiFoxProKeywordInfo("FILETOSTR", "*FILETOSTR"),
            new StiFoxProKeywordInfo("GETDATE", "*GETDATE"),
            new StiFoxProKeywordInfo("GETPEM", "*GETPEM"),
            new StiFoxProKeywordInfo("GETWORDCOUNT", "*GETWORDCOUNT"),
            new StiFoxProKeywordInfo("GETWORDNUM", "*GETWORDNUM"),
            new StiFoxProKeywordInfo("LEFT", "*LEFT"),
            new StiFoxProKeywordInfo("LEN", "*LEN"),
            new StiFoxProKeywordInfo("LOWER", "ToLowerCase"),
            new StiFoxProKeywordInfo("LTRIM", "*LTRIM"),
            new StiFoxProKeywordInfo("MAX", "*MAX"),
            new StiFoxProKeywordInfo("MIN", "*MIN"),
            new StiFoxProKeywordInfo("OCCURS", "*OCCURS"),
            new StiFoxProKeywordInfo("OEMTOANSI", "*OEMTOANSI"),
            new StiFoxProKeywordInfo("OLDVAL", "*OLDVAL"),
            new StiFoxProKeywordInfo("PADC", "*PADC"),
            new StiFoxProKeywordInfo("PADL", "*PADL"),
            new StiFoxProKeywordInfo("PADR", "*PADR"),
            new StiFoxProKeywordInfo("PEMSTATUS", "*PEMSTATUS"),
            new StiFoxProKeywordInfo("PROPER", "*PROPER"),
            new StiFoxProKeywordInfo("RAT", "*RAT"),
            new StiFoxProKeywordInfo("REPLICATE", "*REPLICATE"),
            new StiFoxProKeywordInfo("RIGHT", "*RIGHT"),
            new StiFoxProKeywordInfo("RTRIM", "*RTRIM"),
            new StiFoxProKeywordInfo("SOUNDEX", "*SOUNDEX"),
            new StiFoxProKeywordInfo("SPACE", "*SPACE"),
            new StiFoxProKeywordInfo("STR", "fnc_STR"),
            new StiFoxProKeywordInfo("STREXTRACT", "*STREXTRACT"),
            new StiFoxProKeywordInfo("STRTRAN", "*STRTRAN"),
            new StiFoxProKeywordInfo("STUFF", "*STUFF"),
            new StiFoxProKeywordInfo("SUBSTR", "*SUBSTR"),
            new StiFoxProKeywordInfo("TRANSFORM", "*TRANSFORM"),
            new StiFoxProKeywordInfo("TRIM", "*TRIM"),
            new StiFoxProKeywordInfo("TYPE", "*TYPE"),
            new StiFoxProKeywordInfo("UPPER", "ToUpperCase"),
            new StiFoxProKeywordInfo("AT_C", "*AT_C"),
            new StiFoxProKeywordInfo("ATCC", "*ATCC"),
            new StiFoxProKeywordInfo("CHRTRANC", "*CHRTRANC"),
            new StiFoxProKeywordInfo("LEFTC", "*LEFTC"),
            new StiFoxProKeywordInfo("LENC", "*LENC"),
            new StiFoxProKeywordInfo("RATC", "*RATC"),
            new StiFoxProKeywordInfo("RIGHTC", "*RIGHTC"),
            new StiFoxProKeywordInfo("TEXTMERGE", "*TEXTMERGE"),
            new StiFoxProKeywordInfo("STRCONV", "*STRCONV"),
            new StiFoxProKeywordInfo("STUFFC", "*STUFFC"),
            new StiFoxProKeywordInfo("SUBSTRC", "*SUBSTRC"),

            new StiFoxProKeywordInfo("ABS", "*ABS"),
            new StiFoxProKeywordInfo("ACOS", "*ACOS"),
            new StiFoxProKeywordInfo("ASIN", "*ASIN"),
            new StiFoxProKeywordInfo("ATAN", "*ATAN"),
            new StiFoxProKeywordInfo("ATN2", "*ATN2"),
            new StiFoxProKeywordInfo("AVG", "*AVG"),
            new StiFoxProKeywordInfo("BINTOC", "*BINTOC"),
            new StiFoxProKeywordInfo("BITAND", "*BITAND"),
            new StiFoxProKeywordInfo("BITCLEAR", "*BITCLEAR"),
            new StiFoxProKeywordInfo("BITLSHIFT", "*BITLSHIFT"),
            new StiFoxProKeywordInfo("BITRSHIFT", "*BITRSHIFT"),
            new StiFoxProKeywordInfo("BITSET", "*BITSET"),
            new StiFoxProKeywordInfo("BITTEST", "*BITTEST"),
            new StiFoxProKeywordInfo("BITXOR", "*BITXOR"),
            new StiFoxProKeywordInfo("CEILING", "*CEILING"),
            new StiFoxProKeywordInfo("COS", "*COS"),
            new StiFoxProKeywordInfo("COUNT", "*COUNT"),
            new StiFoxProKeywordInfo("DTOR", "*DTOR"),
            new StiFoxProKeywordInfo("EXP", "*EXP"),
            new StiFoxProKeywordInfo("FLOOR", "*FLOOR"),
            new StiFoxProKeywordInfo("FV", "*FV"),
            new StiFoxProKeywordInfo("INT", "*INT"),
            new StiFoxProKeywordInfo("LOG", "*LOG"),
            new StiFoxProKeywordInfo("LOG10", "*LOG10"),
            //new StiFoxProKeywordInfo("MAX", "*MAX"),
            //new StiFoxProKeywordInfo("MIN", "*MIN"),
            new StiFoxProKeywordInfo("MOD", "*MOD"),
            new StiFoxProKeywordInfo("MTON", "*MTON"),
            new StiFoxProKeywordInfo("NTOM", "*NTOM"),
            new StiFoxProKeywordInfo("PAYMENT", "*PAYMENT"),
            new StiFoxProKeywordInfo("PI", "*PI"),
            new StiFoxProKeywordInfo("PV", "*PV"),
            new StiFoxProKeywordInfo("RAND", "*RAND"),
            new StiFoxProKeywordInfo("ROUND", "Math.Round"),
            new StiFoxProKeywordInfo("RECCOUNT", "*RECCOUNT"),
            new StiFoxProKeywordInfo("RECNO", "*RECNO"),
            new StiFoxProKeywordInfo("RTOD", "*RTOD"),
            new StiFoxProKeywordInfo("SIGN", "*SIGN"),
            new StiFoxProKeywordInfo("SIN", "*SIN"),
            new StiFoxProKeywordInfo("SQRT", "*SQRT"),
            new StiFoxProKeywordInfo("SUM", "*SUM"),
            new StiFoxProKeywordInfo("TAN", "*TAN"),
            new StiFoxProKeywordInfo("VAL", "*VAL"),

            new StiFoxProKeywordInfo("BETWEEN", "*BETWEEN"),
            new StiFoxProKeywordInfo("DELETED", "*DELETED"),
            new StiFoxProKeywordInfo("EMPTY", "fnc_EMPTY"),
            new StiFoxProKeywordInfo("IIF", "IIF"),
            new StiFoxProKeywordInfo("INLIST", "*INLIST"),
            new StiFoxProKeywordInfo("NVL", "*NVL"),
            new StiFoxProKeywordInfo("SEEK", "*SEEK"),

            new StiFoxProKeywordInfo("CDOW", "*CDOW"),
            new StiFoxProKeywordInfo("CMONTH", "*CMONTH"),
            new StiFoxProKeywordInfo("CTOD", "*CTOD"),
            new StiFoxProKeywordInfo("CTOT", "*CTOT"),
            new StiFoxProKeywordInfo("DATE", "Today^"),
            new StiFoxProKeywordInfo("DATETIME", "*DATETIME"),
            new StiFoxProKeywordInfo("DAY", "*DAY"),
            new StiFoxProKeywordInfo("DMY", "*DMY"),
            new StiFoxProKeywordInfo("DOW", "*DOW"),
            new StiFoxProKeywordInfo("DTOC", "*DTOC"),
            new StiFoxProKeywordInfo("DTOS", "*DTOS"),
            new StiFoxProKeywordInfo("DTOT", "*DTOT"),
            new StiFoxProKeywordInfo("GOMONTH", "*GOMONTH"),
            new StiFoxProKeywordInfo("HOUR", "*HOUR"),
            //new StiFoxProKeywordInfo("MAX", "*MAX"),
            new StiFoxProKeywordInfo("MDY", "*MDY"),
            //new StiFoxProKeywordInfo("MIN", "*MIN"),
            new StiFoxProKeywordInfo("MINUTE", "*MINUTE"),
            new StiFoxProKeywordInfo("MONTH", "*MONTH"),
            new StiFoxProKeywordInfo("QUARTER", "*QUARTER"),
            new StiFoxProKeywordInfo("SEC", "*SEC"),
            new StiFoxProKeywordInfo("SECONDS", "*SECONDS"),
            new StiFoxProKeywordInfo("TIME", "Time^"),
            new StiFoxProKeywordInfo("TTOC", "*TTOC"),
            new StiFoxProKeywordInfo("TTOD", "*TTOD"),
            new StiFoxProKeywordInfo("WEEK", "*WEEK"),
            new StiFoxProKeywordInfo("YEAR", "*YEAR"),

            new StiFoxProKeywordInfo("ICASE", "Switch"),
            new StiFoxProKeywordInfo("SYS", "*SYS"),
            new StiFoxProKeywordInfo("TRANS", "*TRANS"),

            new StiFoxProKeywordInfo("ISNULL", "*ISNULL"),
            new StiFoxProKeywordInfo("EVALUATE", "*EVALUATE"),
            new StiFoxProKeywordInfo("field", "*field")
        };

        private static Hashtable functionsList = null;
        private static Hashtable FunctionsList
        {
            get
            {
                if (functionsList == null)
                {
                    functionsList = new Hashtable();

                    for (int index = 0; index < FoxProFunctions.Count; index++)
                    {
                        var func = FoxProFunctions[index];
                        functionsList[func.Name.ToLowerInvariant()] = (StiFunctionType)(index + 1);
                    }
                }
                return functionsList;
            }
        }



        private static Hashtable methodsList = null;
        private static Hashtable MethodsList
        {
            get
            {
                if (methodsList == null)
                {
                    methodsList = new Hashtable();
                    //methodsList[""] = StiFunctionType.FunctionNameSpace;
                    methodsList["Substring"] = StiMethodType.Substring;
                    methodsList["ToString"] = StiMethodType.ToString;
                    methodsList["ToLower"] = StiMethodType.ToLower;
                    methodsList["ToUpper"] = StiMethodType.ToUpper;
                    methodsList["IndexOf"] = StiMethodType.IndexOf;
                    methodsList["StartsWith"] = StiMethodType.StartsWith;
                    methodsList["EndsWith"] = StiMethodType.EndsWith;

                    methodsList["Parse"] = StiMethodType.Parse;
                    methodsList["Contains"] = StiMethodType.Contains;
                    methodsList["GetData"] = StiMethodType.GetData;
                    methodsList["ToQueryString"] = StiMethodType.ToQueryString;

                    methodsList["AddYears"] = StiMethodType.AddYears;
                    methodsList["AddMonths"] = StiMethodType.AddMonths;
                    methodsList["AddDays"] = StiMethodType.AddDays;
                    methodsList["AddHours"] = StiMethodType.AddHours;
                    methodsList["AddMinutes"] = StiMethodType.AddMinutes;
                    methodsList["AddSeconds"] = StiMethodType.AddSeconds;
                    methodsList["AddMilliseconds"] = StiMethodType.AddMilliseconds;
                }
                return methodsList;
            }
        }

        // список параметров функций, которые не надо вычислять сразу, 
        // а оставлять в виде кода для последующего вычисления
        private static Hashtable parametersList = null;
        private static Hashtable ParametersList
        {
            get
            {
                if (parametersList == null)
                {
                    parametersList = new Hashtable();

                    #region Fill parameters list
                                      
                    //parametersList[StiFunctionType.Rank] = StiParameterNumber.Param2;
                    #endregion
                }
                return parametersList;
            }
        }


        private Hashtable componentsList = null;
        private Hashtable ComponentsList
        {
            get
            {
                if (componentsList == null)
                {
                    componentsList = new Hashtable();
                    StiComponentsCollection comps = report.GetComponents();
                    foreach (StiComponent comp in comps)
                    {
                        componentsList[comp.Name] = comp;
                    }
                    componentsList["this"] = report;
                }
                return componentsList;
            }
        }

        
        private static Hashtable constantsList = null;
        private static Hashtable ConstantsList
        {
            get
            {
                if (constantsList == null)
                {
                    constantsList = new Hashtable();
                    constantsList[".t."] = true;
                    constantsList[".f."] = false;
                    constantsList["null"] = null;
                    constantsList["DBNull"] = namespaceObj;
                    constantsList["DBNull.Value"] = DBNull.Value;

                    constantsList["MidpointRounding"] = namespaceObj;
                    constantsList["MidpointRounding.ToEven"] = MidpointRounding.ToEven;
                    constantsList["MidpointRounding.AwayFromZero"] = MidpointRounding.AwayFromZero;

                    //constantsList["StiRankOrder"] = namespaceObj;
                    //constantsList["StiRankOrder.Asc"] = StiRankOrder.Asc;
                    //constantsList["StiRankOrder.Desc"] = StiRankOrder.Desc;

                    //constantsList[""] = ;
                }
                return constantsList;
            }
        }


        private static object namespaceObj = new object();
        //private static Hashtable namespacesList = null;
        //private static Hashtable NamespacesList
        //{
        //    get
        //    {
        //        if (namespacesList == null)
        //        {
        //            namespacesList = new Hashtable();
        //            namespacesList["MidpointRounding"] = namespaceObj;
        //            namespacesList["System"] = namespaceObj;
        //            namespacesList["System.Convert"] = namespaceObj;
        //            namespacesList["Math"] = namespaceObj;
        //            namespacesList["Totals"] = namespaceObj;
        //        }
        //        return constantsList;
        //    }
        //}
        #endregion
    }
}
