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

namespace Stimulsoft.Report.Import
{
    public partial class StiVisualFoxProParser
    {
        #region Enums
        public enum StiTokenType
        {
            /// <summary>
            /// Empty token
            /// </summary>
            Empty = 0,
            Delimiter,
            Variable,
            SystemVariable,
            DataSourceField,
            BusinessObjectField,
            Number,
            Function,   //args
            Method,     //parent + args
            Property,   //parent
            Component,
            Cast,
            String,

            /// <summary>
            /// .
            /// </summary>
            Dot,
            /// <summary>
            /// ,
            /// </summary>
            Comma,
            /// <summary>
            /// :
            /// </summary>
            Colon,
            /// <summary>
            /// ;
            /// </summary>
            SemiColon,
            /// <summary>
            /// Shift to the left Token.
            /// </summary>
            Shl,
            /// <summary>
            /// Shift to the right Token.
            /// </summary>
            Shr,
            /// <summary>
            /// = Assign Token.
            /// </summary>
            Assign,
            /// <summary>
            /// Equal Token.
            /// </summary>
            Equal,
            /// <summary>
            /// NotEqual Token.
            /// </summary>
            NotEqual,
            /// <summary>
            /// LeftEqual Token.
            /// </summary>
            LeftEqual,
            /// <summary>
            /// Left Token.
            /// </summary>
            Left,
            /// <summary>
            /// RightEqual Token.
            /// </summary>
            RightEqual,
            /// <summary>
            /// Right Token.
            /// </summary>
            Right,
            /// <summary>
            /// Logical NOT Token.
            /// </summary>
            Not,
            /// <summary>
            /// Logical OR Token.
            /// </summary>
            Or,
            /// <summary>
            /// Logical AND Token.
            /// </summary>
            And,
            /// <summary>
            /// ^
            /// </summary>
            Xor,
            /// <summary>
            /// Double logical OR Token.
            /// </summary>
            DoubleOr,
            /// <summary>
            /// Double logical AND Token.
            /// </summary>
            DoubleAnd,
            ///// <summary>
            ///// Copyright
            ///// </summary>
            //Copyright,
            /// <summary>
            /// ?
            /// </summary>
            Question,
            /// <summary>
            /// +
            /// </summary>
            Plus,
            /// <summary>
            /// -
            /// </summary>
            Minus,
            /// <summary>
            /// *
            /// </summary>
            Mult,
            /// <summary>
            /// /
            /// </summary>
            Div,
            ///// <summary>
            ///// \
            ///// </summary>
            //Splash,
            /// <summary>
            /// %
            /// </summary>
            Percent,
            ///// <summary>
            ///// @
            ///// </summary>
            //Ampersand,
            ///// <summary>
            ///// #
            ///// </summary>
            //Sharp,
            /// <summary>
            /// $
            /// </summary>
            Dollar,
            ///// <summary>
            ///// €
            ///// </summary>
            //Euro,
            ///// <summary>
            ///// ++
            ///// </summary>
            //DoublePlus,
            ///// <summary>
            ///// --
            ///// </summary>
            //DoubleMinus,
            /// <summary>
            /// (
            /// </summary>
            LParenthesis,
            /// <summary>
            /// )
            /// </summary>
            RParenthesis,
            ///// <summary>
            ///// {
            ///// </summary>
            //LBrace,
            ///// <summary>
            ///// }
            ///// </summary>
            //RBrace,
            /// <summary>
            /// [
            /// </summary>
            LBracket,
            /// <summary>
            /// ]
            /// </summary>
            RBracket,
            ///// <summary>
            ///// Token contains value.
            ///// </summary>
            //Value,
            /// <summary>
            /// Token contains identifier.
            /// </summary>
            Identifier,
            /// <summary>
            /// 
            /// </summary>
            Unknown,
            ///// <summary>
            ///// EOF Token.
            ///// </summary>
            //EOF
        }

        public enum StiAsmCommandType
        {
            PushValue       = 2000,
            PushVariable,
            PushSystemVariable,
            PushDataSourceField,
            PushBusinessObjectField,
            PushFunction,
            PushMethod,
            PushProperty,
            PushComponent,
            PushArrayElement,
            CopyToVariable,
            Add             = 2020,
            Sub,
            Mult,
            Div,
            Mod,
            Power,
            Neg,
            Cast,
            Not,
            CompareLeft,
            CompareLeftEqual,
            CompareRight,
            CompareRightEqual,
            CompareEqual,
            CompareNotEqual,
            Shl,
            Shr,
            And,
            And2,
            Or,
            Or2,
            Xor,
            Contains,
            Bracers
        }

        public enum StiSystemVariableType
        {
            PAGENO
        }

        public enum StiPropertyType
        {
            Year,
            Month,
            Day,
            Hour,
            Minute,
            Second,
            Length,
            From,
            To,
            FromDate,
            ToDate,
            FromTime,
            ToTime,
            SelectedLine,
            Name,
            TagValue,
            Days,
            Hours,
            Milliseconds,
            Minutes,
            Seconds,
            Ticks,
            Count
        }

        //[Flags]
        //public enum StiParserDataType
        //{
        //    None = 0x0000,
        //    Object = 0x0001,
        //    Object = 0x7FFFFFFF,
        //    zFloat = 0x0002,
        //    zDouble = 0x0004,
        //    zDecimal = 0x0008,
        //    Byte = 0x0010,
        //    SByte = 0x0020,
        //    Int16 = 0x0040,
        //    UInt16 = 0x0080,
        //    Int32 = 0x0100,
        //    UInt32 = 0x0200,
        //    Int64 = 0x0400,
        //    UInt64 = 0x0800,
        //    Bool = 0x1000,
        //    Char = 0x2000,
        //    String = 0x4000,
        //    DateTime = 0x8000,
        //    TimeSpan = 0x00010000,
        //    Image = 0x00020000,

        //    Short = Byte | SByte | Int16,
        //    UShort = Byte | UInt16 | Char,
        //    Int = Short | UInt16 | Int32,
        //    UInt = UShort | UInt32,
        //    Long = Int | UInt32 | Int64,
        //    ULong = UInt | UInt64,
        //    Float = zFloat | Long | ULong,
        //    Double = zDouble | Long | ULong | zFloat,
        //    Decimal = zDecimal | Long | ULong,

        //    BasedType = 0x10000000,
        //    FixedType = 0x20000000,
        //    Nullable = 0x40000000
        //}

        public enum StiFunctionType
        {
            NameSpace = 0,

            IIF, 

            //methods

            //m_Substring = 1000,
            //m_ToString,
            //m_ToLower,
            //m_ToUpper,
            //m_IndexOf,
            //m_StartsWith,
            //m_EndsWith,
            
            //m_Parse,
            //m_Contains,
            //m_GetData,
            //m_ToQueryString,
            
            //m_AddYears,
            //m_AddMonths,
            //m_AddDays,
            //m_AddHours,
            //m_AddMinutes,
            //m_AddSeconds,
            //m_AddMilliseconds,
            
            //m_MethodNameSpace,


            //operators

            //op_Add = 2020,
            //op_Sub,
            //op_Mult,
            //op_Div,
            //op_Mod,
            //op_Power,
            //op_Neg,
            //op_Cast,
            //op_Not,
            //op_CompareLeft,
            //op_CompareLeftEqual,
            //op_CompareRight,
            //op_CompareRightEqual,
            //op_CompareEqual,
            //op_CompareNotEqual,
            //op_Shl,
            //op_Shr,
            //op_And,
            //op_And2,
            //op_Or,
            //op_Or2,
            //op_Xor

        }

        public enum StiMethodType
        {
            Substring = 1000,
            ToString,
            ToLower,
            ToUpper,
            IndexOf,
            StartsWith,
            EndsWith,

            Parse,
            Contains,
            GetData,
            ToQueryString,

            AddYears,
            AddMonths,
            AddDays,
            AddHours,
            AddMinutes,
            AddSeconds,
            AddMilliseconds,

            MethodNameSpace
        }

        [Flags]
        public enum StiParameterNumber
        {
            Param1 = 1,
            Param2 = 2,
            Param3 = 4,
            Param4 = 8
        }

        #endregion
    }
}
