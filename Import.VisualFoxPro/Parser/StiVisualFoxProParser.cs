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
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Text;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Engine;

namespace Stimulsoft.Report.Import
{
    public partial class StiVisualFoxProParser
    {
        #region Structures
        public class StiParserData
        {
            public object Data = null;
            public List<StiAsmCommand> AsmList = null;
            public List<StiAsmCommand> ConditionAsmList = null;
            public StiVisualFoxProParser Parser = null;

            public StiParserData(object data, List<StiAsmCommand> asmList, StiVisualFoxProParser parser)
            {
                this.Data = data;
                this.AsmList = asmList;
                this.Parser = parser;
                this.ConditionAsmList = null;
            }

            public StiParserData(object data, List<StiAsmCommand> asmList, StiVisualFoxProParser parser, List<StiAsmCommand> conditionAsmList)
            {
                this.Data = data;
                this.AsmList = asmList;
                this.Parser = parser;
                this.ConditionAsmList = conditionAsmList;
            }
        }

        public class StiFilterParserData
        {
            public StiComponent Component;
            public string Expression;

            public StiFilterParserData(StiComponent component, string expression)
            {
                this.Component = component;
                this.Expression = expression;
            }
        }

        public class StiToken
        {
            public StiTokenType Type = StiTokenType.Empty;
            public string Value;
            public object ValueObject;
            public int Position = -1;
            public int Length = -1;

            public StiToken(StiTokenType type, int position, int length)
            {
                this.Type = type;
                this.Position = position;
                this.Length = length;
            }
            public StiToken(StiTokenType type, int position)
            {
                this.Type = type;
                this.Position = position;
            }
            public StiToken(StiTokenType type)
            {
                this.Type = type;
            }
            public StiToken()
            {
                this.Type = StiTokenType.Empty;
            }

            public override string ToString()
            {
                return string.Format("TokenType={0}{1}", Type.ToString(), Value != null ? string.Format(", value=\"{0}\"", Value) : "");
            }
        }

        public class StiAsmCommand
        {
            public StiAsmCommandType Type;
            public object Parameter1;
            public object Parameter2;
            public int Position = -1;
            public int Length = -1;

            public StiAsmCommand(StiAsmCommandType type)
                : this(type, null, null)
            {
            }

            public StiAsmCommand(StiAsmCommandType type, object parameter)
                : this(type, parameter, null)
            {
            }

            public StiAsmCommand(StiAsmCommandType type, object parameter1, object parameter2)
            {
                this.Type = type;
                this.Parameter1 = parameter1;
                this.Parameter2 = parameter2;
            }

            public override string ToString()
            {
                return string.Format("{0}({1},{2})", Type.ToString(),
                    Parameter1 != null ? ParamToString(Parameter1) : "null",
                    Parameter2 != null ? ParamToString(Parameter2) : "null");
            }

            private string ParamToString(object param)
            {
                if (param is StiFunctionType)
                {
                    int index = (int)param;
                    if (index <= FoxProFunctions.Count)
                    {
                        return FoxProFunctions[index - 1].Name;
                    }
                }
                else if (param is StiSystemVariableType)
                {
                    int index = (int)param;
                    if (index <= FoxProSystemVariables.Count)
                    {
                        return FoxProSystemVariables[index - 1].Name;
                    }
                }
                return param.ToString();
            }
        }

        public class StiParserMethodInfo
        {
            public StiFunctionType Name;
            public int Number;
            public Type[] Arguments;
            public Type ReturnType;

            public StiParserMethodInfo(StiFunctionType name, int number, Type[] arguments)
            {
                this.Name = name;
                this.Number = number;
                this.Arguments = arguments;
                this.ReturnType = typeof(string);
            }

            public StiParserMethodInfo(StiFunctionType name, int number, Type[] arguments, Type returnType)
            {
                this.Name = name;
                this.Number = number;
                this.Arguments = arguments;
                this.ReturnType = returnType;
            }

        }
        #endregion

        #region Fields
        private StiReport report = null;
        private string inputExpression = string.Empty;
        private StiComponent component = null;
        private object sender = null;

        private int position = 0;
        private List<StiToken> tokensList = null;
        private StiToken currentToken = null;
        private int tokenPos = 0;
        private List<StiAsmCommand> asmList = null;
        private Hashtable hashAliases = null;

        private int expressionPosition = 0;
        #endregion
 
        #region ParseTextValue
        public static object ParseTextValue(string inputExpression, StiComponent component)
        {
            bool storeToPrint = false;
            return ParseTextValue(inputExpression, component, component, ref storeToPrint, true, false, null);
        }

        public static object ParseTextValue(string inputExpression, StiComponent component, ref bool storeToPrint, bool executeIfStoreToPrint)
        {
            return ParseTextValue(inputExpression, component, component, ref storeToPrint, executeIfStoreToPrint, false, null);
        }

        public static object ParseTextValue(string inputExpression, StiComponent component, ref bool storeToPrint, bool executeIfStoreToPrint, bool returnAsmList)
        {
            return ParseTextValue(inputExpression, component, component, ref storeToPrint, executeIfStoreToPrint, returnAsmList, null);
        }

        public static object ParseTextValue(string inputExpression, StiComponent component, object sender)
        {
            bool storeToPrint = false;
            return ParseTextValue(inputExpression, component, sender, ref storeToPrint, true, false, null);
        }
        public static object ParseTextValue(string inputExpression, StiComponent component, object sender, ref bool storeToPrint, bool executeIfStoreToPrint)
        {
            return ParseTextValue(inputExpression, component, sender, ref storeToPrint, executeIfStoreToPrint, false, null);
        }

        public static object ParseTextValue(string inputExpression, StiComponent component, object sender, ref bool storeToPrint, bool executeIfStoreToPrint, bool returnAsmList)
        {
            return ParseTextValue(inputExpression, component, sender, ref storeToPrint, executeIfStoreToPrint, returnAsmList, null);
        }

        public static object ParseTextValue(string inputExpression, StiComponent component, object sender, ref bool storeToPrint, bool executeIfStoreToPrint, bool returnAsmList, StiVisualFoxProParser parser)
        {
            if (string.IsNullOrEmpty(inputExpression)) return null;
            if (parser == null)
            {
                parser = new StiVisualFoxProParser();
            }
            parser.report = component.Report;
            parser.component = component;
            parser.sender = sender;
            List<StiAsmCommand> list = null;
            StiEngine engine = component.Report.Engine;
            string expressionId = inputExpression + component.Name;

            //if (engine != null)
            //{
            //    if (engine.parserConversionStore == null)
            //    {
            //        engine.parserConversionStore = new Hashtable();
            //    }
            //    if (engine.parserConversionStore.Contains(expressionId))
            //    {
            //        list = (List<StiAsmCommand>)engine.parserConversionStore[expressionId];
            //    }
            //}

            if (list == null)
            {
                try
                {
                    list = new List<StiAsmCommand>();
                    int counter = 0;
                    int pos = 0;
                    while (pos < inputExpression.Length)
                    {
                        #region Plain text
                        int posBegin = pos;
                        while (pos < inputExpression.Length && inputExpression[pos] != '{') pos++;
                        if (pos != posBegin)
                        {
                            list.Add(new StiAsmCommand(StiAsmCommandType.PushValue, inputExpression.Substring(posBegin, pos - posBegin)));
                            counter++;
                            if (counter > 1) list.Add(new StiAsmCommand(StiAsmCommandType.Add));
                        }
                        #endregion

                        #region Expression
                        if (pos < inputExpression.Length && inputExpression[pos] == '{')
                        {
                            pos++;
                            posBegin = pos;
                            bool flag = false;
                            while (pos < inputExpression.Length)
                            {
                                if (inputExpression[pos] == '"')
                                {
                                    pos++;
                                    int pos2 = pos;
                                    while (pos2 < inputExpression.Length)
                                    {
                                        if (inputExpression[pos2] == '"') break;
                                        if (inputExpression[pos2] == '\\') pos2++;
                                        pos2++;
                                    }
                                    pos = pos2 + 1;
                                    continue;
                                }
                                if (inputExpression[pos] == '}')
                                {
                                    string currentExpression = inputExpression.Substring(posBegin, pos - posBegin);
                                    if (currentExpression != null && currentExpression.Length > 0)
                                    {
                                        parser.expressionPosition = posBegin;
                                        list.AddRange(parser.ParseToAsm(currentExpression));
                                        counter++;
                                        if (counter > 1)
                                        {
                                            list.Add(new StiAsmCommand(StiAsmCommandType.Cast, TypeCode.String));
                                            list.Add(new StiAsmCommand(StiAsmCommandType.Add));
                                        }
                                    }
                                    flag = true;
                                    pos++;
                                    break;
                                }
                                pos++;
                            }
                            if (!flag)
                            {
                                parser.expressionPosition = posBegin;
                                list.Add(new StiAsmCommand(StiAsmCommandType.PushValue, inputExpression.Substring(posBegin)));
                                counter++;
                                if (counter > 1) list.Add(new StiAsmCommand(StiAsmCommandType.Add));
                            }
                        }
                        #endregion
                    }
                }
                catch
                {
                    //if (engine != null)
                    //{
                    //    engine.parserConversionStore[expressionId] = new List<StiAsmCommand>();
                    //}
                    //throw ex;
                    return inputExpression;
                }
                //if (engine != null)
                //{
                //    engine.parserConversionStore[expressionId] = list;
                //}
            }

            if (returnAsmList) return list;

            string result = parser.MakeOutput(list);

            return result;
        }

        private List<StiAsmCommand> ParseToAsm(string inputExpression)
        {
            this.inputExpression = inputExpression;
            MakeTokensList();
            asmList = new List<StiAsmCommand>();
            eval_exp();
            return asmList;
        }
        #endregion

        #region Check for DataBandsUsedInPageTotals
        internal static void CheckForDataBandsUsedInPageTotals(StiText stiText)
        {
            try
            {
                StiReport report = stiText.Report;
                bool storeToPrint = false;
                object result = ParseTextValue(stiText.Text.Value, stiText, ref storeToPrint, false, true);
            }
            catch (Exception ex)
            {
                string str = string.Format("Expression in Text property of '{0}' can't be evaluated! {1}", stiText.Name, ex.Message);
                StiLogService.Write(stiText.GetType(), str);
                StiLogService.Write(stiText.GetType(), ex.Message);
                stiText.Report.WriteToReportRenderingMessages(str);
            }
        }
        #endregion

        #region Prepare variable value
        internal static object PrepareVariableValue(StiVariable var, StiReport report, StiText textBox = null)
        {
            if (textBox == null)
            {
                textBox = new StiText
                {
                    Name = "**ReportVariables**",
                    Page = report.Pages[0]
                };
            }

            object obj = null;

            if (var.Type.IsValueType || var.Type == typeof(string))
            {
                if (var.InitBy == StiVariableInitBy.Value)
                {
                    obj = var.ValueObject;
                }
                else
                {
                    obj = StiVisualFoxProParser.ParseTextValue("{" + var.Value + "}", textBox);
                }
            }
            else
            {
                obj = Activator.CreateInstance(var.Type);
                if (obj is Range)
                {
                    var range = obj as Range;
                    if (var.InitBy == StiVariableInitBy.Value)
                    {
                        range.FromObject = (var.ValueObject as Range).FromObject;
                        range.ToObject = (var.ValueObject as Range).ToObject;
                    }
                    else
                    {
                        range.FromObject = StiVisualFoxProParser.ParseTextValue("{" + var.InitByExpressionFrom + "}", textBox);
                        range.ToObject = StiVisualFoxProParser.ParseTextValue("{" + var.InitByExpressionTo + "}", textBox);
                    }
                }
            }

            //store calculated value in variables hash
            report[var.Name] = obj;

            return obj;
        }
        #endregion

        public string defaultDataSourceName = "dataTable";

    }
}
