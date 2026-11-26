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
using System.IO;
using System.Text;
using System.Globalization;

namespace Stimulsoft.Report.Import
{
	/// <summary>
	/// Summary description for StiRtfLex.
	/// </summary>
	public class StiRtfLex
    {
        #region Variables
        private TextReader rtf;
		private const int Eof = -1;
        #endregion

        #region Methods
        public StiRtfToken NextToken()
		{
			int c;
			StiRtfToken token = new StiRtfToken();

			c = rtf.Read();
            //a tab of the file is a tab of the text, the same as the keyword "tab"
            while (c == '\r' || c == '\n' || c == '\0')
            {
                c = rtf.Read();
            }

			if (c != Eof)
			{
				switch (c)
				{
					case '{':
						token.Type = StiRtfTokenType.GroupStart;
						break;
					case '}':
						token.Type = StiRtfTokenType.GroupEnd;
						break;
					case '\\':
						ParseKeyword(token);
						break;
					default:
						token.Type = StiRtfTokenType.Text;
						ParseText((char)c, token);
						break;
				}
			}
			else
			{
				token.Type = StiRtfTokenType.Eof;
			}

			return token;
        }

        private void ParseKeyword(StiRtfToken token)
		{
			int c = rtf.Peek();

            #region Special character or control symbol
            if (!Char.IsLetter((char)c))
            {
                rtf.Read();
                if (c == '\\' || c == '{' || c == '}')
                {
                    //Special character
                    token.Type = StiRtfTokenType.Text;
                    token.Key = ((char)c).ToString();
                }
                else
                {
                    //Control symbol
                    token.Type = StiRtfTokenType.Control;
                    token.Key = ((char)c).ToString();
                    if (token.Key == "\'")
                    {
                        string code = string.Format("{0}{1}", (char)rtf.Read(), (char)rtf.Read());
                        //a damaged code is skipped, it must not stop the whole import
                        if (int.TryParse(code, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var charCode))
                        {
                            token.HasParameter = true;
                            token.Parameter = charCode;
                        }
                    }
                }
                return;
            }
            #endregion

            #region Keyword
            StringBuilder keyWord = new StringBuilder();
            c = rtf.Peek();
			while (Char.IsLetter((char)c))
			{
				rtf.Read();
                keyWord.Append((char)c);
				c = rtf.Peek();
			}
			token.Type = StiRtfTokenType.Keyword;
            token.Key = keyWord.ToString();
            #endregion

            #region Parameter
            if (Char.IsDigit((char)c) || c == '-')
			{
				token.HasParameter = true;

                bool negative = false;
                if (c == '-')
				{
					negative = true;
					rtf.Read();
				}
                StringBuilder paramStr = new StringBuilder();
                c = rtf.Peek();
				while (Char.IsDigit((char)c))
				{
					rtf.Read();
					paramStr.Append((char)c);
					c = rtf.Peek();
				}
				//a parameter without digits or a too long one is zero
				int.TryParse(paramStr.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out var paramInt);
                if (negative)
                {
                    paramInt = -paramInt;
                }
				token.Parameter = paramInt;
            }
            #endregion

            if (c == ' ')
			{
				rtf.Read();
			}

            //the binary data is not a text: its braces and backslashes must not be read as tokens
            if (token.Key == "bin" && token.HasParameter)
            {
                for (int index = 0; index < token.Parameter && rtf.Peek() != Eof; index++)
                {
                    rtf.Read();
                }
            }
        }

        private void ParseText(char ch, StiRtfToken token)
		{
			StringBuilder text = new StringBuilder(ch.ToString());

			int c = rtf.Peek();
			while (c == '\r' || c == '\n' || c == '\0')
			{
				rtf.Read();
				c = rtf.Peek();
			}

			while (c != '\\' && c != '}' && c != '{' && c != Eof)
			{
				rtf.Read();
				text.Append((char)c);
				c = rtf.Peek();
				while (c == '\r' || c == '\n' || c == '\0')
				{
					rtf.Read();
					c = rtf.Peek();
				}
			}

			token.Key = text.ToString();
        }
        #endregion

		public StiRtfLex(TextReader rtfReader)
		{
			rtf = rtfReader;
        }
    }
}
