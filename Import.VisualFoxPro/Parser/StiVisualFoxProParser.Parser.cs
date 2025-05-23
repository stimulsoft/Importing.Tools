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
using System.Reflection;
using Stimulsoft.Report.Components;

namespace Stimulsoft.Report.Import
{
    public partial class StiVisualFoxProParser
    {
        #region Parser

        //----------------------------------------
        // Entry point
        //----------------------------------------
        private void eval_exp()
        {
            tokenPos = 0;
            if (tokensList.Count == 0)
            {
                ThrowError(ParserErrorCode.ExpressionIsEmpty);
                return;
            }
            eval_exp0();
            if (tokenPos <= tokensList.Count)
            {
                ThrowError(ParserErrorCode.UnprocessedLexemesRemain);
            }
        }

        private void eval_exp0()
        {
            get_token();
            eval_exp01();
        }


        //----------------------------------------
        // Assignment
        //----------------------------------------
        private void eval_exp01()
        {
            if (currentToken.Type == StiTokenType.Variable)
            {
                StiToken variableToken = currentToken;
                get_token();
                if (currentToken.Type != StiTokenType.Assign)
                {
                    tokenPos--;
                    currentToken = tokensList[tokenPos - 1];
                }
                else
                {
                    get_token();
                    eval_exp1();
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.CopyToVariable, variableToken.Value));
                    return;
                }
            }
            eval_exp1();
        }

        //----------------------------------------
        // Conditional Operator  ? : 
        //----------------------------------------
        private void eval_exp1()
        {
            eval_exp10();
            if (currentToken.Type == StiTokenType.Question)
            {
                get_token();
                eval_exp10();
                if (currentToken.Type != StiTokenType.Colon) ThrowError(ParserErrorCode.SyntaxError, currentToken);
                get_token();
                eval_exp10();
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushFunction, StiFunctionType.IIF, 3));
            }
        }


        //----------------------------------------
        // Logical OR
        //----------------------------------------
        private void eval_exp10()
        {
            eval_exp11();
            while (currentToken.Type == StiTokenType.DoubleOr)
            {
                get_token();
                eval_exp11();
                asmList.Add(new StiAsmCommand(StiAsmCommandType.Or2));
            }
        }

        //----------------------------------------
        // Logical AND
        //----------------------------------------
        private void eval_exp11()
        {
            eval_exp12();
            while (currentToken.Type == StiTokenType.DoubleAnd)
            {
                get_token();
                eval_exp12();
                asmList.Add(new StiAsmCommand(StiAsmCommandType.And2));
            }
        }

        //----------------------------------------
        // Binary OR
        //----------------------------------------
        private void eval_exp12()
        {
            eval_exp14();
            if (currentToken.Type == StiTokenType.Or)
            {
                get_token();
                eval_exp14();
                asmList.Add(new StiAsmCommand(StiAsmCommandType.Or));
            }
        }

        //----------------------------------------
        // Binary XOR
        //----------------------------------------
        private void eval_exp14()
        {
            eval_exp15();
            if (currentToken.Type == StiTokenType.Xor)
            {
                get_token();
                eval_exp15();
                asmList.Add(new StiAsmCommand(StiAsmCommandType.Xor));
            }
        }

        //----------------------------------------
        // Binary AND
        //----------------------------------------
        private void eval_exp15()
        {
            eval_exp16();
            if (currentToken.Type == StiTokenType.And)
            {
                get_token();
                eval_exp16();
                asmList.Add(new StiAsmCommand(StiAsmCommandType.And));
            }
        }


        //----------------------------------------
        // Equality (==, !=)
        //----------------------------------------
        private void eval_exp16()
        {
            eval_exp17();
            if (currentToken.Type == StiTokenType.Equal || currentToken.Type == StiTokenType.NotEqual)
            {
                StiAsmCommand command = new StiAsmCommand(StiAsmCommandType.CompareEqual);
                if (currentToken.Type == StiTokenType.NotEqual) command.Type = StiAsmCommandType.CompareNotEqual;
                get_token();
                eval_exp17();
                asmList.Add(command);
            }
        }

        //----------------------------------------
        // Relational and type testing (<, >, <=, >=, is, as)
        //----------------------------------------
        private void eval_exp17()
        {
            eval_exp18();
            if (currentToken.Type == StiTokenType.Left || currentToken.Type == StiTokenType.LeftEqual ||
                currentToken.Type == StiTokenType.Right || currentToken.Type == StiTokenType.RightEqual)
            {
                StiAsmCommand command = null;
                if (currentToken.Type == StiTokenType.Left) command = new StiAsmCommand(StiAsmCommandType.CompareLeft);
                if (currentToken.Type == StiTokenType.LeftEqual) command = new StiAsmCommand(StiAsmCommandType.CompareLeftEqual);
                if (currentToken.Type == StiTokenType.Right) command = new StiAsmCommand(StiAsmCommandType.CompareRight);
                if (currentToken.Type == StiTokenType.RightEqual) command = new StiAsmCommand(StiAsmCommandType.CompareRightEqual);
                get_token();
                eval_exp18();
                asmList.Add(command);
            }
        }


        //----------------------------------------
        // Shift (<<, >>)
        //----------------------------------------
        private void eval_exp18()
        {
            eval_exp2();
            if ((currentToken.Type == StiTokenType.Shl) || (currentToken.Type == StiTokenType.Shr))
            {
                StiAsmCommand command = new StiAsmCommand(StiAsmCommandType.Shl);
                if (currentToken.Type == StiTokenType.Shr) command.Type = StiAsmCommandType.Shr;
                get_token();
                eval_exp2();
                asmList.Add(command);
            }
        }


        //----------------------------------------
        // Addition or subtraction of two terms
        //----------------------------------------
        private void eval_exp2()
        {
            eval_exp3();
            while ((currentToken.Type == StiTokenType.Plus) || (currentToken.Type == StiTokenType.Minus))
            {
                StiToken operation = currentToken;
                get_token();
                eval_exp3();
                if (operation.Type == StiTokenType.Minus)
                {
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.Sub));
                }
                else if (operation.Type == StiTokenType.Plus)
                {
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.Add));
                }
            }
        }


        //----------------------------------------
        // Multiplication or division of two factors
        //----------------------------------------
        private void eval_exp3()
        {
            eval_exp34();
            while (currentToken.Type == StiTokenType.Mult || currentToken.Type == StiTokenType.Div || currentToken.Type == StiTokenType.Percent)
            {
                StiToken operation = currentToken;
                get_token();
                eval_exp34();
                if (operation.Type == StiTokenType.Mult)
                {
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.Mult));
                }
                else if (operation.Type == StiTokenType.Div)
                {
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.Div));
                }
                if (operation.Type == StiTokenType.Percent)
                {
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.Mod));
                }
            }
        }


        //----------------------------------------
        // Operation $
        //----------------------------------------
        private void eval_exp34()
        {
            eval_exp4();
            if (currentToken.Type == StiTokenType.Dollar)
            {
                get_token();
                eval_exp4();
                asmList.Add(new StiAsmCommand(StiAsmCommandType.Contains));
            }
        }


        //----------------------------------------
        // Exponentiation
        //----------------------------------------
        private void eval_exp4()
        {
            eval_exp5();
            //if (currentToken.Type == StiTokenType.Xor)
            //{
            //    get_token();
            //    eval_exp4();
            //    asmList.Add(new StiAsmCommand(StiAsmCommandType.Power));
            //}
        }


        //----------------------------------------
        // Calculating unary sign + and -
        //----------------------------------------
        private void eval_exp5()
        {
            StiAsmCommand command = null;
            if (currentToken.Type == StiTokenType.Plus || currentToken.Type == StiTokenType.Minus || currentToken.Type == StiTokenType.Not)
            {
                if (currentToken.Type == StiTokenType.Minus) command = new StiAsmCommand(StiAsmCommandType.Neg);
                if (currentToken.Type == StiTokenType.Not) command = new StiAsmCommand(StiAsmCommandType.Not);
                get_token();
            }
            eval_exp6();
            if (command != null)
            {
                asmList.Add(command);
            }
        }


        //----------------------------------------
        // Processing expression in parentheses
        //----------------------------------------
        private void eval_exp6()
        {
            if (currentToken.Type == StiTokenType.LParenthesis)
            {
                get_token();
                if (currentToken.Type == StiTokenType.Cast)
                {
                    TypeCode typeCode = (TypeCode)TypesList[currentToken.Value];
                    get_token();
                    if (currentToken.Type != StiTokenType.RParenthesis)
                    {
                        ThrowError(ParserErrorCode.RightParenthesisExpected);  //unbalanced parentheses
                    }
                    get_token();
                    eval_exp5();
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.Cast, typeCode));
                }
                else
                {
                    eval_exp1();
                    if (currentToken.Type != StiTokenType.RParenthesis)
                    {
                        ThrowError(ParserErrorCode.RightParenthesisExpected);  //unbalanced parentheses
                    }
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.Bracers));  //for info only
                    get_token();
                    if (currentToken.Type == StiTokenType.Dot)
                    {
                        get_token();
                        eval_exp7();
                    }
                    if (currentToken.Type == StiTokenType.LBracket)
                    {
                        eval_exp62();
                    }
                }
            }
            else
            {
                eval_exp62();
            }
        }


        //----------------------------------------
        // Process indexes
        //----------------------------------------
        private void eval_exp62()
        {
            if (currentToken.Type == StiTokenType.LBracket)
            {
                int argsCount = 0;
                while (argsCount == 0 || currentToken.Type == StiTokenType.Comma)
                {
                    get_token();
                    eval_exp1();
                    argsCount++;
                }
                if (currentToken.Type != StiTokenType.RBracket)
                {
                    ThrowError(ParserErrorCode.SyntaxError, currentToken);  //unbalanced square brackets  //!!!
                }
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushArrayElement, argsCount + 1));
                get_token();
                if (currentToken.Type == StiTokenType.LBracket)
                {
                    eval_exp62();
                }
                if (currentToken.Type == StiTokenType.Dot)
                {
                    get_token();
                    eval_exp7();
                }
            }
            else
            {
                eval_exp7();
            }
        }


        //----------------------------------------
        // Computing methods and properties
        //----------------------------------------
        private void eval_exp7()
        {
            atom();
            if (currentToken.Type == StiTokenType.Dot)
            {
                get_token();
                eval_exp7();
            }
            if (currentToken.Type == StiTokenType.LBracket)
            {
                eval_exp62();
            }
        }


        //----------------------------------------
        // Getting the value of a number or variable
        //----------------------------------------
        private void atom()
        {
            if (currentToken.Type == StiTokenType.Variable)
            {
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushVariable, currentToken.Value));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.SystemVariable)
            {
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushSystemVariable, SystemVariablesList[currentToken.Value]));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.Function)
            {
                StiToken function = currentToken;
                StiFunctionType functionType = (StiFunctionType)FunctionsList[function.Value];
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushFunction, functionType, get_args_count(functionType)));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.Method)
            {
                StiToken method = currentToken;
                StiMethodType methodType = (StiMethodType)MethodsList[method.Value];
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushMethod, methodType, get_args_count(methodType) + 1));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.Property)
            {
                StiToken function = currentToken;
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushProperty, PropertiesList[function.Value]));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.DataSourceField)
            {
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushDataSourceField, currentToken.Value));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.BusinessObjectField)
            {
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushBusinessObjectField, currentToken.Value));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.Component)
            {
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushComponent, ComponentsList[currentToken.Value]));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.Number)
            {
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushValue, currentToken.ValueObject));
                get_token();
                return;
            }
            else if (currentToken.Type == StiTokenType.String)
            {
                asmList.Add(new StiAsmCommand(StiAsmCommandType.PushValue, currentToken.ValueObject));
                get_token();
                return;
            }
            else
            {
                if (currentToken.Type == StiTokenType.Empty)
                {
                    ThrowError(ParserErrorCode.UnexpectedEndOfExpression);
                }
                ThrowError(ParserErrorCode.SyntaxError, currentToken);
            }
        }


        //----------------------------------------
        // Getting function arguments
        //----------------------------------------
        private int get_args_count(object name)
        {
            var args = get_args();

            //If aggregateComponent is not specified, search it
            var func = (StiFunctionType)name;

            //we check which parameters should be left as expressions
            int bitsValue = 0;
            if (ParametersList.Contains(name)) bitsValue = (int)ParametersList[name];
            int bitsCounter = 1;
            foreach (var arg in args)
            {
                if ((bitsValue & bitsCounter) > 0)
                {
                    asmList.Add(new StiAsmCommand(StiAsmCommandType.PushValue, arg));
                }
                else
                {
                    asmList.AddRange(arg);
                }
                bitsCounter = bitsCounter << 1;
            }

            return args.Count;
        }

        private List<List<StiAsmCommand>> get_args()
        {
            var args = new List<List<StiAsmCommand>>();

            get_token();
            if (currentToken.Type != StiTokenType.LParenthesis) ThrowError(ParserErrorCode.LeftParenthesisExpected);
            get_token();
            if (currentToken.Type == StiTokenType.RParenthesis)
            {
                //empty list
                return args;
            }
            else
            {
                tokenPos--;
                currentToken = tokensList[tokenPos - 1];
            }

            //process list of values
            List<StiAsmCommand> tempAsmList = asmList;
            do
            {
                asmList = new List<StiAsmCommand>();
                eval_exp0();
                args.Add(asmList);
            }
            while (currentToken.Type == StiTokenType.Comma);
            asmList = tempAsmList;

            if (currentToken.Type != StiTokenType.RParenthesis) ThrowError(ParserErrorCode.RightParenthesisExpected);
            return args;
        }


        private void get_token()
        {
            if (tokenPos < tokensList.Count)
            {
                currentToken = tokensList[tokenPos];
            }
            else
            {
                currentToken = new StiToken();
            }
            tokenPos++;
        }

        #endregion
    }
}
