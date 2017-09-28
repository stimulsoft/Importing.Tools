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
using System.Text;
using Stimulsoft.Report.Dictionary;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;

namespace Stimulsoft.Report.Import
{
    public partial class StiVisualFoxProParser
    {

        #region MakeOutput
        public string MakeOutput(object objectAsmList)
        {
            List<StiAsmCommand> asmList = objectAsmList as List<StiAsmCommand>;
            if (asmList == null || asmList.Count == 0) return string.Empty;
            Stack stack = new Stack();
            ArrayList argsList = null;
            string par1 = string.Empty;
            string par2 = string.Empty;
            foreach (StiAsmCommand asmCommand in asmList)
            {
                switch (asmCommand.Type)
                {
                    case StiAsmCommandType.PushValue:
                        if (asmCommand.Parameter1 is string)
                        {
                            stack.Push(string.Format("\"{0}\"", asmCommand.Parameter1));
                        }
                        else
                        {
                            stack.Push(Convert.ToString(asmCommand.Parameter1));
                        }
                        break;
                    case StiAsmCommandType.PushVariable:
                        stack.Push(Convert.ToString(asmCommand.Parameter1));
                        break;
                    case StiAsmCommandType.PushSystemVariable:
                        stack.Push(FoxProSystemVariables[Convert.ToInt32(asmCommand.Parameter1) - 1].StiName);
                        break;
                    //case StiAsmCommandType.PushComponent:
                    //    stack.Push(asmCommand.Parameter1);
                    //    break;

                    //case StiAsmCommandType.CopyToVariable:
                    //    //report.Dictionary.Variables[(string)asmCommand.Parameter1].ValueObject = stack.Peek();
                    //    report[(string)asmCommand.Parameter1] = stack.Peek();
                    //    break;

                    case StiAsmCommandType.PushFunction:
                        #region Push function value
                        argsList = new ArrayList();
                        for (int index = 0; index < (int)asmCommand.Parameter2; index++)
                        {
                            argsList.Add(stack.Pop());
                        }
                        argsList.Reverse();
                        int indexFunc = (int)asmCommand.Parameter1 - 1;
                        StringBuilder funcResult = new StringBuilder();
                        string stiFuncName = FoxProFunctions[indexFunc].StiName;
                        if (stiFuncName.EndsWith("^"))  //SystemVariable
                        {
                            funcResult.Append(stiFuncName.Substring(0, stiFuncName.Length - 1));
                        }
                        else
                        {
                            funcResult.Append(stiFuncName + "(");
                            for (int index = 0; index < argsList.Count; index++)
                            {
                                if (index > 0) funcResult.Append(",");
                                funcResult.Append(Convert.ToString(argsList[index]));
                            }
                            funcResult.Append(")");
                        }
                        stack.Push(funcResult.ToString());
                        #endregion
                        break;

                    //case StiAsmCommandType.PushMethod:
                    //    #region Push method value
                    //    argsList = new ArrayList();
                    //    for (int index = 0; index < (int)asmCommand.Parameter2; index++)
                    //    {
                    //        argsList.Add(stack.Pop());
                    //    }
                    //    argsList.Reverse();
                    //    stack.Push(call_method(asmCommand.Parameter1, argsList));
                    //    #endregion
                    //    break;

                    //case StiAsmCommandType.PushProperty:
                    //    argsList = new ArrayList();
                    //    argsList.Add(stack.Pop());
                    //    stack.Push(call_property(asmCommand.Parameter1, argsList));
                    //    break;

                    case StiAsmCommandType.PushDataSourceField:
                        stack.Push(Convert.ToString(asmCommand.Parameter1));
                        break;

                    case StiAsmCommandType.Bracers:
                        stack.Push("(" + Convert.ToString(stack.Pop()) + ")");
                        break;

                    //case StiAsmCommandType.PushArrayElement:
                    //    #region Push array value
                    //    argsList = new ArrayList();
                    //    for (int index = 0; index < (int)asmCommand.Parameter1; index++)
                    //    {
                    //        argsList.Add(stack.Pop());
                    //    }
                    //    argsList.Reverse();
                    //    stack.Push(call_arrayElement(argsList));
                    //    #endregion
                    //    break;

                    case StiAsmCommandType.Add:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " + " + par2);
                        break;
                    case StiAsmCommandType.Sub:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " - " + par2);
                        break;

                    case StiAsmCommandType.Mult:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " * " + par2);
                        break;
                    case StiAsmCommandType.Div:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " / " + par2);
                        break;
                    //case StiAsmCommandType.Mod:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_Mod(par1, par2));
                    //    break;

                    //case StiAsmCommandType.Power:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_Pow(par1, par2));
                    //    break;

                    //case StiAsmCommandType.Neg:
                    //    par1 = stack.Pop();
                    //    stack.Push(op_Neg(par1));
                    //    break;

                    //case StiAsmCommandType.Cast:
                    //    par1 = stack.Pop();
                    //    par2 = asmCommand.Parameter1;
                    //    stack.Push(op_Cast(par1, par2));
                    //    break;

                    case StiAsmCommandType.Not:
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push("!" + par1);
                        break;

                    //case StiAsmCommandType.CompareLeft:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_CompareLeft(par1, par2));
                    //    break;
                    //case StiAsmCommandType.CompareLeftEqual:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_CompareLeftEqual(par1, par2));
                    //    break;
                    //case StiAsmCommandType.CompareRight:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_CompareRight(par1, par2));
                    //    break;
                    //case StiAsmCommandType.CompareRightEqual:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_CompareRightEqual(par1, par2));
                    //    break;

                    case StiAsmCommandType.CompareEqual:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " == " + par2);
                        break;
                    case StiAsmCommandType.CompareNotEqual:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " != " + par2);
                        break;

                    //case StiAsmCommandType.Shl:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_Shl(par1, par2));
                    //    break;
                    //case StiAsmCommandType.Shr:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_Shr(par1, par2));
                    //    break;

                    case StiAsmCommandType.And:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " & " + par2);
                        break;
                    case StiAsmCommandType.Or:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " | " + par2);
                        break;
                    //case StiAsmCommandType.Xor:
                    //    par2 = stack.Pop();
                    //    par1 = stack.Pop();
                    //    stack.Push(op_Xor(par1, par2));
                    //    break;

                    case StiAsmCommandType.And2:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " && " + par2);
                        break;
                    case StiAsmCommandType.Or2:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par1 + " || " + par2);
                        break;

                    case StiAsmCommandType.Contains:
                        par2 = Convert.ToString(stack.Pop());
                        par1 = Convert.ToString(stack.Pop());
                        stack.Push(par2 + ".Contains(" + par1 + ")");
                        break;

                }
            }
            return "{" + Convert.ToString(stack.Pop()) + "}";
        }
 
        #endregion

    }
}
