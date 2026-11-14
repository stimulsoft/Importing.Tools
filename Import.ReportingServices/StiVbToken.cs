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

namespace Stimulsoft.Report.Import.ReportingServices
{
    /// <summary>
    /// A token of an expression or of the code of Visual Basic with its text exactly as in the source.
    /// </summary>
    internal class StiVbToken
    {
        #region Properties
        public StiVbTokenType Type { get; }

        public string Text { get; }

        /// <summary>
        /// The token means something for the code: it is not a space, an end of a line or a comment.
        /// </summary>
        public bool IsSignificant
        {
            get
            {
                return Type != StiVbTokenType.Whitespace && Type != StiVbTokenType.NewLine &&
                    Type != StiVbTokenType.Comment;
            }
        }

        /// <summary>
        /// The text of a name without the square brackets around it.
        /// </summary>
        public string Name
        {
            get
            {
                return Text.Length > 1 && Text[0] == '[' && Text[Text.Length - 1] == ']'
                    ? Text.Substring(1, Text.Length - 2)
                    : Text;
            }
        }

        /// <summary>
        /// The value of a string without the quotes, a doubled quote inside turns into one.
        /// </summary>
        public string StringValue
        {
            get
            {
                var text = Text;

                if (text.EndsWith("c", StringComparison.OrdinalIgnoreCase) && text.Length > 1 &&
                    text[text.Length - 2] == '"')
                    text = text.Substring(0, text.Length - 1);

                var end = text.Length > 1 && text[text.Length - 1] == '"' ? text.Length - 1 : text.Length;

                return text.Substring(1, end - 1).Replace("\"\"", "\"");
            }
        }
        #endregion

        #region Methods
        /// <summary>
        /// Checks that the token is the given name or keyword, in any case as Visual Basic does.
        /// </summary>
        public bool IsName(string name)
        {
            return Type == StiVbTokenType.Identifier && string.Equals(Name, name, StringComparison.OrdinalIgnoreCase);
        }

        public bool IsOperator(string text)
        {
            return Type == StiVbTokenType.Operator && Text == text;
        }

        public override string ToString()
        {
            return $"{Type}: {Text}";
        }
        #endregion

        public StiVbToken(StiVbTokenType type, string text)
        {
            Type = type;
            Text = text;
        }
    }
}
