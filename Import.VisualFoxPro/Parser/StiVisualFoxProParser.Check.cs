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

namespace Stimulsoft.Report.Import
{
    public partial class StiVisualFoxProParser
    {
        #region Exception provider

        private static string[] errorsList = {
            "Syntax error",        //0
            "Integral constant is too large",    //1        значение целочисленной константы слишком велико
            "The expression is empty",             //2
            "Division by zero",              //3
            "Unexpected end of expression",  //4
            "The name '{0}' does not exist in the current context",   //5
            "Syntax error - unprocessed lexemes remain",  //6
            "( expected",         //7
            ") expected",         //8
            "Field, method, or property is not found: '{0}'", //9
            "Operator '{0}' cannot be applied to operands of type '{1}' and type '{2}'",    //10
            "The function is not found: '{0}'",                     //11
            "No overload for method '{0}' takes '{1}' arguments",   //12
            "The '{0}' function has invalid argument '{1}': cannot convert from '{2}' to '{3}'",   //13
            "The '{0}' function is not yet implemented",            //14
            "The '{0}' method has invalid argument '{1}': cannot convert from '{2}' to '{3}'",  //15
            "'{0}' does not contain a definition for '{1}'",   //16
            "There is no matching overloaded method for '{0}({1})'"};   //17

        private enum ParserErrorCode
        {
            SyntaxError = 0,
            IntegralConstantIsTooLarge = 1,
            ExpressionIsEmpty = 2,
            DivisionByZero = 3,
            UnexpectedEndOfExpression = 4,
            NameDoesNotExistInCurrentContext = 5,
            UnprocessedLexemesRemain = 6,
            LeftParenthesisExpected = 7,
            RightParenthesisExpected = 8,
            FieldMethodOrPropertyNotFound = 9,
            OperatorCannotBeAppliedToOperands = 10,
            FunctionNotFound = 11,
            NoOverloadForMethodTakesNArguments = 12,
            FunctionHasInvalidArgument = 13,
            FunctionNotYetImplemented = 14,
            MethodHasInvalidArgument = 15,
            ItemDoesNotContainDefinition = 16,
            NoMatchingOverloadedMethod = 17
        }


        // Отображение сообщения о синтаксической ошибке
        private void ThrowError(ParserErrorCode code)
        {
            ThrowError(code, null, string.Empty, string.Empty, string.Empty, string.Empty);
        }

        private void ThrowError(ParserErrorCode code, string message1)
        {
            ThrowError(code, null, message1, string.Empty, string.Empty, string.Empty);
        }

        private void ThrowError(ParserErrorCode code, string message1, string message2)
        {
            ThrowError(code, null, message1, message2, string.Empty, string.Empty);
        }

        private void ThrowError(ParserErrorCode code, string message1, string message2, string message3)
        {
            ThrowError(code, null, message1, message2, message3, string.Empty);
        }

        private void ThrowError(ParserErrorCode code, string message1, string message2, string message3, string message4)
        {
            ThrowError(code, null, message1, message2, message3, message4);
        }


        private void ThrowError(ParserErrorCode code, StiToken token)
        {
            ThrowError(code, token, string.Empty, string.Empty, string.Empty, string.Empty);
        }

        private void ThrowError(ParserErrorCode code, StiToken token, string message1)
        {
            ThrowError(code, token, message1, string.Empty, string.Empty, string.Empty);
        }

        private void ThrowError(ParserErrorCode code, StiToken token, string message1, string message2)
        {
            ThrowError(code, token, message1, message2, string.Empty, string.Empty);
        }

        private void ThrowError(ParserErrorCode code, StiToken token, string message1, string message2, string message3)
        {
            ThrowError(code, token, message1, message2, message3, string.Empty);
        }

        private void ThrowError(ParserErrorCode code, StiToken token, string message1, string message2, string message3, string message4)
        {
            string errorMessage = "Unknown error";
            int errorCode = (int)code;
            if (errorCode < errorsList.Length)
            {
                errorMessage = string.Format(errorsList[errorCode], message1, message2, message3, message4);
            }
            string fullMessage = "Parser error: " + errorMessage;
            StiParserException ex = new StiParserException(fullMessage);
            ex.BaseMessage = errorMessage;
            if (token != null)
            {
                ex.Position = expressionPosition + token.Position;
                ex.Length = token.Length;
            }
            throw ex;
        }

        public class StiParserException : Exception
        {
            public string BaseMessage = null;
            public int Position = -1;
            public int Length = -1;

            public StiParserException(string message)
                : base(message)
            {
            }

            public StiParserException()
                : base()
            {
            }
        }

        #endregion

        #region GetTypeName
        private string GetTypeName(object value)
        {
            if (value == null)
            {
                return "null";
            }
            else
            {
                return value.GetType().ToString();
            }
        }
        #endregion

    }
}
