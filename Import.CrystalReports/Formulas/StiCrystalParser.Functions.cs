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

using System.Collections.Generic;
using System.Linq;

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// The custom functions of a report in Crystal syntax: Function ([Optional] TypeVar [Array | Range] name
    /// [:= default], ...) and the body.
    /// </summary>
    public partial class StiCrystalParser
    {
        #region Methods.Functions
        public static StiCrystalCustomFunction ParseFunction(string name, string text)
        {
            var parser = new StiCrystalParser(new StiCrystalLexer(text).Tokenize());
            parser.ExpectWord("function");
            parser.Expect(StiCrystalTokenType.LeftParenthesis);

            var parameters = new List<StiCrystalNode>();
            while (parser.Current.Type != StiCrystalTokenType.RightParenthesis)
            {
                parameters.Add(parser.ParseFunctionParameter());
                if (parser.Current.Type != StiCrystalTokenType.Comma)
                    break;

                parser.index++;
            }

            parser.Expect(StiCrystalTokenType.RightParenthesis);
            var body = parser.ParseSequence();
            if (parser.Current.Type != StiCrystalTokenType.End)
                throw new StiCrystalFormulaException($"Unexpected '{parser.Current.Text}'", parser.Current.Position);

            return new StiCrystalCustomFunction(name, parameters, body);
        }

        private StiCrystalNode ParseFunctionParameter()
        {
            var position = Current.Position;
            if (Current.IsWord("optional"))
                index++;

            var type = Current.Text.ToLowerInvariant();
            if (Current.Type != StiCrystalTokenType.Identifier || !VariableTypes.Contains(type))
                throw new StiCrystalFormulaException("The type of the parameter expected", Current.Position);

            index++;
            var isArray = Current.IsWord("array");
            var isRange = Current.IsWord("range");
            if (isArray || isRange)
                index++;

            var name = Current.Text;
            Expect(StiCrystalTokenType.Identifier);

            var parameter = new StiCrystalNode(StiCrystalNodeKind.Declare, name, position)
            {
                Scope = "local",
                VariableType = type,
                IsArrayVariable = isArray,
                IsRangeVariable = isRange
            };

            if (Current.IsOperator(":="))
            {
                index++;
                parameter.Add(ParseImp());
            }

            return parameter;
        }
        #endregion
    }
}
