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
using System.Reflection;
using System.Text;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report;

namespace Stimulsoft.Report.Import
{
    public partial class StiVisualFoxProParser
    {
        #region Lexer

        #region GetNextLexem
        // Получение очередной лексемы.
        private StiToken GetNextLexem()
        {
            //пропустить пробелы, символы табуляции и другие незначащие символы
            while (position < inputExpression.Length && isWhiteSpace(inputExpression[position])) position++;
            if (position >= inputExpression.Length) return null;

            StiToken token = null;
            char ch = inputExpression[position];
            if (char.IsLetter(ch) || (ch == '_'))
            {
                int pos2 = position + 1;
                while ((pos2 < inputExpression.Length) && (char.IsLetterOrDigit(inputExpression[pos2]) || inputExpression[pos2] == '_')) pos2++;
                token = new StiToken();
                token.Value = inputExpression.Substring(position, pos2 - position);
                token.Type = StiTokenType.Identifier;
                token.Position = position;
                token.Length = pos2 - position;
                position = pos2;

                string alias = token.Value;
                if ((token.Position > 0) && (inputExpression[token.Position - 1] == '.')) alias = "." + alias;
                if (hashAliases.ContainsKey(alias))
                {
                    token.Value = (string)hashAliases[alias];
                }

                token.Value = token.Value.ToLowerInvariant();   //FoxPro only

                return token;
            }
            else if (char.IsDigit(ch))
            {
                token = new StiToken();
                token.Type = StiTokenType.Number;
                token.Position = position;
                token.ValueObject = ScanNumber();
                token.Length = position - token.Position;
                return token;
            }
            else if ((ch == '"' || ch == '\'') || ((ch == '@' || ch == '\\') && (position < inputExpression.Length - 1) && (inputExpression[position + 1] == '"')))
            {
                #region "String"
                bool needReplaceBackslash = true;
                if (ch == '@')
                {
                    needReplaceBackslash = false;
                    position++;
                }

                position++;
                int pos2 = position;
                while (pos2 < inputExpression.Length)
                {
                    if (inputExpression[pos2] == '"') break;
                    if (inputExpression[pos2] == '\'') break;
                    if (inputExpression[pos2] == '\\')
                    {
                        pos2++;
                        if (inputExpression[pos2] == '"') break;
                    }
                    pos2++;
                }
                token = new StiToken();
                token.Type = StiTokenType.String;
                //token.ValueObject = inputExpression.Substring(position, pos2 - position).Replace("\\r", "\r").Replace("\\n", "\n").Replace("\\t", "\t").Replace("\\\"", "\"").Replace("\\'", "'").Replace("\\\\", "\\");
                string st = inputExpression.Substring(position, pos2 - position);
                if (needReplaceBackslash)
                {
                    token.ValueObject = ReplaceBackslash(st);
                }
                else
                {
                    token.ValueObject = st;
                }

                token.Position = position - 1;
                position = pos2 + 1;
                token.Length = position - token.Position;
                return token;
                #endregion
            }
            else
            {
                #region check for alias bracket
                if (ch == '[')
                {
                    int pos2 = inputExpression.IndexOf(']', position);
                    if (pos2 != -1)
                    {
                        pos2++;
                        string alias = inputExpression.Substring(position, pos2 - position);
                        if ((position > 0) && (inputExpression[position - 1] == '.'))
                        {
                            alias = "." + alias; 
                        }
                        if (hashAliases.ContainsKey(alias))
                        {
                            token = new StiToken();
                            token.Value = ((string)hashAliases[alias]).ToLowerInvariant();
                            token.Type = StiTokenType.Identifier;
                            token.Position = position;
                            token.Length = pos2 - position;
                            position = pos2;
                            return token;
                        }
                    }
                }
                #endregion

                #region Delimiters
                int tPos = position;
                position++;
                char ch2 = ' ';
                if (position < inputExpression.Length) ch2 = inputExpression[position];
                switch (ch)
                {
                    case '.': return new StiToken(StiTokenType.Dot, tPos, 1);
                    case '(': return new StiToken(StiTokenType.LParenthesis, tPos, 1);
                    case ')': return new StiToken(StiTokenType.RParenthesis, tPos, 1);
                    case '[': return new StiToken(StiTokenType.LBracket, tPos, 1);
                    case ']': return new StiToken(StiTokenType.RBracket, tPos, 1);
                    case '+': return new StiToken(StiTokenType.Plus, tPos, 1);
                    case '-': return new StiToken(StiTokenType.Minus, tPos, 1);
                    case '*': return new StiToken(StiTokenType.Mult, tPos, 1);
                    case '/': return new StiToken(StiTokenType.Div, tPos, 1);
                    case '%': return new StiToken(StiTokenType.Percent, tPos, 1);
                    case '^': return new StiToken(StiTokenType.Xor, tPos, 1);
                    case ',': return new StiToken(StiTokenType.Comma, tPos, 1);
                    case ':': return new StiToken(StiTokenType.Colon, tPos, 1);
                    case ';': return new StiToken(StiTokenType.SemiColon, tPos, 1);
                    case '?': return new StiToken(StiTokenType.Question, tPos, 1);
                    case '$': return new StiToken(StiTokenType.Dollar, tPos, 1);
                    case '|':
                        if (ch2 == '|')
                        {
                            position++;
                            return new StiToken(StiTokenType.DoubleOr, tPos, 2);
                        }
                        else return new StiToken(StiTokenType.Or, tPos, 1);
                    case '&':
                        if (ch2 == '&')
                        {
                            position++;
                            return new StiToken(StiTokenType.DoubleAnd, tPos, 2);
                        }
                        else return new StiToken(StiTokenType.And, tPos, 1);
                    case '!':
                        if (ch2 == '=')
                        {
                            position++;
                            return new StiToken(StiTokenType.NotEqual, tPos, 2);
                        }
                        else return new StiToken(StiTokenType.Not, tPos, 1);
                    case '=':
                        return new StiToken(StiTokenType.Equal, tPos, 1);
                        //if (ch2 == '=')
                        //{
                        //    position++;
                        //    return new StiToken(StiTokenType.Equal, tPos, 2);
                        //}
                        //else return new StiToken(StiTokenType.Assign, tPos, 1);
                    case '<':
                        if (ch2 == '<')
                        {
                            position++;
                            return new StiToken(StiTokenType.Shl, tPos, 2);
                        }
                        else if (ch2 == '=')
                        {
                            position++;
                            return new StiToken(StiTokenType.LeftEqual, tPos, 2);
                        }
                        else if (ch2 == '>')
                        {
                            position++;
                            return new StiToken(StiTokenType.NotEqual, tPos, 2);
                        }
                        else return new StiToken(StiTokenType.Left, tPos, 1);
                    case '>':
                        if (ch2 == '>')
                        {
                            position++;
                            return new StiToken(StiTokenType.Shr, tPos, 2);
                        }
                        else if (ch2 == '=')
                        {
                            position++;
                            return new StiToken(StiTokenType.RightEqual, tPos, 2);
                        }
                        else return new StiToken(StiTokenType.Right, tPos, 1);


                    default:
                        token = new StiToken(StiTokenType.Unknown);
                        token.ValueObject = ch;
                        token.Position = tPos;
                        token.Length = 1;
                        return token;
                }
                #endregion
            }
        }

        private static bool isWhiteSpace(char ch)
        {
            return char.IsWhiteSpace(ch) || ch < 0x20;
        }
        #endregion

        #region BuildAliases
        private void BuildAliases()
        {
            if (hashAliases != null) return;

            hashAliases = new Hashtable();

            foreach (StiDataSource dataSource in report.Dictionary.DataSources)
            {
                string dataSourceName = dataSource.Name;
                string dataSourceAlias = GetCorrectedAlias(dataSource.Alias);
                if (dataSourceAlias != dataSourceName)
                {
                    hashAliases[dataSourceAlias] = dataSourceName;
                }

                foreach (StiDataColumn dataColumn in dataSource.Columns)
                {
                    string dataColumnName = dataColumn.Name;
                    string dataColumnAlias = GetCorrectedAlias(dataColumn.Alias);
                    if (dataColumnAlias != dataColumnName)
                    {
                        hashAliases["." + dataColumnAlias] = dataColumnName;
                    }
                }
            }

            foreach (StiDataRelation dataRelation in report.Dictionary.Relations)
            {
                string dataRelationName = dataRelation.Name;
                string dataRelationAlias = GetCorrectedAlias(dataRelation.Alias);
                if (dataRelationAlias != dataRelationName)
                {
                    hashAliases["." + dataRelationAlias] = dataRelationName;
                }
            }

            foreach (StiVariable variable in report.Dictionary.Variables)
            {
                string variableName = variable.Name;
                string variableAlias = GetCorrectedAlias(variable.Alias);
                if (variableAlias != variableName)
                {
                    hashAliases[variableAlias] = variableName;
                }
            }
        }

        private static bool IsValidName(string name)
        {
            if (string.IsNullOrEmpty(name) || !(Char.IsLetter(name[0]) || name[0] == '_'))
                return false;

            for (int pos = 0; pos < name.Length; pos++)
                if (!(Char.IsLetterOrDigit(name[pos]) || (name[pos] == '_'))) return false;

            return true;
        }

        private static string GetCorrectedAlias(string alias)
        {
            if (IsValidName(alias)) return alias;
            return string.Format("[{0}]", alias);
        }
        #endregion

        #region ReplaceBackslash
        private static string ReplaceBackslash(string input)
        {
            StringBuilder output = new StringBuilder();
            for (int index = 0; index < input.Length; index++)
            {
                if ((input[index] == '\\') && (index < input.Length - 1))
                {
                    index++;
                    char ch = input[index];
                    switch (ch)
                    {
                        case '\\':
                            output.Append("\\");
                            break;

                        case '\'':
                            output.Append('\'');
                            break;

                        case '"':
                            output.Append('"');
                            break;

                        case '0':
                            output.Append('\0');
                            break;

                        case 'n':
                            output.Append('\n');
                            break;

                        case 'r':
                            output.Append('\r');
                            break;

                        case 't':
                            output.Append('\t');
                            break;

                        default:
                            output.Append("\\" + ch);
                            break;
                    }
                }
                else
                {
                    output.Append(input[index]);
                }
            }

            return output.ToString();
        }
        #endregion

        #region ScanNumber
        private object ScanNumber()
        {
            TypeCode typecode = TypeCode.Int32;
            int posBegin = position;
            int posBeginAll = position;
            //integer part
            while (position != inputExpression.Length && Char.IsDigit(inputExpression[position]))
            {
                position++;
            }
            if (position != inputExpression.Length && inputExpression[position] == '.' &&
                position + 1 != inputExpression.Length && Char.IsDigit(inputExpression[position + 1]))
            {
                //fractional part
                position++;
                while (position != inputExpression.Length && Char.IsDigit(inputExpression[position]))
                {
                    position++;
                }
                typecode = TypeCode.Double;
            }
            string nm = inputExpression.Substring(posBegin, position - posBegin);
            nm = nm.Replace(".", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);

            if (position != inputExpression.Length && Char.IsLetter(inputExpression[position]))
            {
                //postfix
                posBegin = position;
                while (position != inputExpression.Length && Char.IsLetter(inputExpression[position]))
                {
                    position++;
                }
                string postfix = inputExpression.Substring(posBegin, position - posBegin).ToLower();
                if (postfix == "f") typecode = TypeCode.Single;
                if (postfix == "d") typecode = TypeCode.Double;
                if (postfix == "m") typecode = TypeCode.Decimal;
                if (postfix == "l") typecode = TypeCode.Int64;
                if (postfix == "u" || postfix == "ul" || postfix == "lu") typecode = TypeCode.UInt64;
            }

            if ((typecode == TypeCode.Int32) && (nm.Length > 9)) typecode = TypeCode.Int64;

            object result = null;
            try
            {
                result = Convert.ChangeType(nm, typecode);
            }
            catch
            {
                if (typecode == TypeCode.Int32 || typecode == TypeCode.Int64 || typecode == TypeCode.UInt32 || typecode == TypeCode.UInt64)
                {
                    ThrowError(ParserErrorCode.IntegralConstantIsTooLarge, new StiToken(StiTokenType.Number, posBeginAll, position - posBeginAll));
                }
            }
            return result;
        }
        #endregion

        #region PostProcessTokensList
        private List<StiToken> PostProcessTokensList(List<StiToken> tokensList)
        {
            List<StiToken> newList = new List<StiToken>();
            tokenPos = 0;
            while (tokenPos < tokensList.Count)
            {
                StiToken token = tokensList[tokenPos];
                tokenPos++;
                if (token.Type == StiTokenType.Identifier)
                {
                    StiDataSource ds = report.Dictionary.DataSources[token.Value];
                    StiBusinessObject bos = report.Dictionary.BusinessObjects[token.Value];

                    #region check for DataSource field
                    if (ds != null)
                    {
                        StringBuilder fieldPath = new StringBuilder(StiNameValidator.CorrectName(token.Value));
                        while (tokenPos + 1 < tokensList.Count && tokensList[tokenPos].Type == StiTokenType.Dot)
                        {
                            token = tokensList[tokenPos + 1];
                            string nextName = StiNameValidator.CorrectName(token.Value);

                            StiDataRelation dr = GetDataRelationByName(nextName, ds);
                            if (dr != null)
                            {
                                ds = dr.ParentSource;
                                tokenPos += 2;
                                fieldPath.Append(".");
                                fieldPath.Append(dr.NameInSource);
                                continue;
                            }
                            if (ds.Columns.Contains(nextName))
                            {
                                tokenPos += 2;
                                fieldPath.Append(".");
                                fieldPath.Append(nextName);
                                break;
                            }
                            foreach (StiDataColumn column in ds.Columns)
                            {
                                if (StiNameValidator.CorrectName(column.Name) == nextName)
                                {
                                    tokenPos += 2;
                                    fieldPath.Append(".");
                                    fieldPath.Append(column.NameInSource);
                                    break;
                                }
                            }

                            CheckDataSourceField(ds.Name, nextName);
                            tokenPos += 2;
                            fieldPath.Append(".");
                            fieldPath.Append(nextName);

                            //token = tokensList[tokenPos - 1];
                            break;
                        }
                        token.Type = StiTokenType.DataSourceField;
                        //надо оптимизировать и сохранять сразу массив строк !!!!!
                        token.Value = fieldPath.ToString();
                    }
                    #endregion

                    #region check for BusinessObject field
                    else if (bos != null)
                    {
                        StringBuilder fieldPath = new StringBuilder(token.Value);
                        while (tokenPos + 1 < tokensList.Count && tokensList[tokenPos].Type == StiTokenType.Dot)
                        //while (inputExpression[pos2] == '.')
                        {
                            token = tokensList[tokenPos + 1];
                            string nextName = token.Value;

                            if (bos.Columns.Contains(nextName))
                            {
                                tokenPos += 2;
                                fieldPath.Append(".");
                                fieldPath.Append(nextName);
                                break;
                            }
                            bos = bos.BusinessObjects[nextName];
                            if (bos != null)
                            {
                                tokenPos += 2;
                                fieldPath.Append(".");
                                fieldPath.Append(bos.Name);
                                continue;
                            }
                            break;
                        }
                        token.Type = StiTokenType.BusinessObjectField;
                        //надо оптимизировать и сохранять сразу массив строк !!!!!
                        token.Value = fieldPath.ToString();
                    }
                    #endregion

                    else if ((newList.Count > 0) && (newList[newList.Count - 1].Type == StiTokenType.Dot))
                    {
                        if (MethodsList.Contains(token.Value))
                        {
                            token.Type = StiTokenType.Method;
                        }
                        else if (PropertiesList.Contains(token.Value))
                        {
                            token.Type = StiTokenType.Property;
                        }
                        else
                        {
                            ThrowError(ParserErrorCode.FieldMethodOrPropertyNotFound, token, token.Value);
                        }
                    }

                    else if (TypesList.Contains(token.Value))
                    {
                        token.Type = StiTokenType.Cast;

                        if ((tokenPos + 1 < tokensList.Count) && (tokensList[tokenPos].Type == StiTokenType.Dot))
                        {
                            string tempName = token.Value + "." + tokensList[tokenPos + 1].Value;
                            if (FunctionsList.Contains(tempName))
                            {
                                token.Type = StiTokenType.Function;
                                token.Value = tempName;
                                tokenPos += 2;
                            }
                            if (SystemVariablesList.Contains(tempName))
                            {
                                token.Type = StiTokenType.SystemVariable;
                                token.Value = tempName;
                                tokenPos += 2;
                            }
                        }
                    }

                    else if (ComponentsList.Contains(token.Value))
                    {
                        token.Type = StiTokenType.Component;
                        if ((tokenPos + 1 < tokensList.Count) && (tokensList[tokenPos].Type == StiTokenType.Colon) && ComponentsList.Contains(tokensList[tokenPos + 1].Value))
                        {
                            StiComponent comp = (StiComponent)ComponentsList[tokensList[tokenPos + 1].Value];
                            if (comp != null && comp is StiDataBand)
                            {
                                token.Value = (comp as StiDataBand).DataSourceName;
                                token.Type = StiTokenType.DataSourceField;
                                tokenPos += 2;
                            }
                        }
                    }

                    else if (FunctionsList.Contains(token.Value))
                    {
                        while ((StiFunctionType)FunctionsList[token.Value] == StiFunctionType.NameSpace)
                        {
                            if (tokenPos + 1 >= tokensList.Count) ThrowError(ParserErrorCode.UnexpectedEndOfExpression);
                            token.Value += "." + tokensList[tokenPos + 1].Value;
                            tokenPos += 2;
                            if (!FunctionsList.Contains(token.Value)) ThrowError(ParserErrorCode.FunctionNotFound, token, token.Value);
                        }
                        token.Type = StiTokenType.Function;
                    }

                    else if (SystemVariablesList.Contains(token.Value) && (token.Value != "value" || component is Stimulsoft.Report.CrossTab.StiCrossCell))
                    {
                        token.Type = StiTokenType.SystemVariable;
                    }

                    //else if (token.Value.ToLowerInvariant() == "true" || token.Value.ToLowerInvariant() == "false")
                    //{
                    //    if (token.Value.ToLowerInvariant() == "true")
                    //        token.ValueObject = true;
                    //    else
                    //        token.ValueObject = false;
                    //    token.Type = StiTokenType.Number;
                    //}
                    //else if (token.Value.ToLowerInvariant() == "null")
                    //{
                    //    token.ValueObject = null;
                    //    token.Type = StiTokenType.Number;
                    //}

                    else if (ConstantsList.Contains(token.Value))
                    {
                        while (ConstantsList[token.Value] == namespaceObj)
                        {
                            if (tokenPos + 1 >= tokensList.Count) ThrowError(ParserErrorCode.UnexpectedEndOfExpression);
                            string oldTokenValue = token.Value;
                            token.Value += "." + tokensList[tokenPos + 1].Value;
                            tokenPos += 2;
                            if (!ConstantsList.Contains(token.Value)) ThrowError(ParserErrorCode.ItemDoesNotContainDefinition, token, oldTokenValue, tokensList[tokenPos + 1].Value);
                        }
                        token.ValueObject = ConstantsList[token.Value];
                        token.Type = StiTokenType.Number;
                    }

                    else if (report.Dictionary.Variables.Contains(token.Value))
                    {
                        token.Type = StiTokenType.Variable;
                    }
                    else if (token.Value == "or" || token.Value == "and" || token.Value == "not")
                    {
                        if (token.Value == "or") token.Type = StiTokenType.DoubleOr;
                        if (token.Value == "and") token.Type = StiTokenType.DoubleAnd;
                        if (token.Value == "not") token.Type = StiTokenType.Not;
                    }
                    else
                    {
                        if ((tokenPos < tokensList.Count) && (tokensList[tokenPos].Type != StiTokenType.Dot) || (tokenPos == tokensList.Count))
                        {
                            CheckDataSourceField(defaultDataSourceName, token.Value);

                            token.Type = StiTokenType.DataSourceField;
                            token.Value = defaultDataSourceName + "." + token.Value;
                        }
                        else
                        {
                            if ((tokenPos + 1 < tokensList.Count) && (tokensList[tokenPos].Type == StiTokenType.Dot))
                            {
                                CheckDataSourceField(token.Value, tokensList[tokenPos + 1].Value);

                                token.Type = StiTokenType.DataSourceField;
                                token.Value = token.Value + "." + tokensList[tokenPos + 1].Value;
                                tokenPos += 2;
                            }
                            else
                            {
                                ThrowError(ParserErrorCode.NameDoesNotExistInCurrentContext, token, token.Value);
                            }
                        }
                    }
                }
                newList.Add(token);
            }
            return newList;
        }

        //private StiDataSource GetDataSourceByName(string name)
        //{
        //    foreach (StiDataSource ds in report.Dictionary.DataSources)
        //    {
        //        if (ds.Alias == name)
        //        {
        //            return ds;
        //        }
        //    }
        //    return null;
        //}

        private StiDataRelation GetDataRelationByName(string name, StiDataSource ds)
        {
            StiDataRelation dr = null;
            foreach (StiDataRelation drTemp in report.Dictionary.Relations)
            {
                if ((drTemp.Name == name || drTemp.NameInSource == name) && (drTemp.ChildSource == ds))
                {
                    dr = drTemp;
                    break;
                }
            }
            return dr;
        }

        private void CheckDataSourceField(string dataSourceName, string fieldName)
        {
            StiDataSource ds = report.Dictionary.DataSources[dataSourceName];
            if (ds == null)
            {
                ds = new StiSqlSource();
                (ds as StiSqlSource).NameInSource = string.Format("Connection1.{0}", dataSourceName);
                ds.Name = dataSourceName;
                ds.Alias = dataSourceName;
                (ds as StiSqlSource).SqlCommand = string.Format("select * from {0}", dataSourceName);

                report.Dictionary.DataSources.Add(ds);
            }
            StiDataColumn dc = ds.Columns[fieldName];
            if (dc == null)
            {
                dc = new StiDataColumn(fieldName, fieldName, fieldName, typeof(string));
                ds.Columns.Add(dc);
            }
        }
        #endregion

        private void MakeTokensList()
        {
            BuildAliases();
            tokensList = new List<StiToken>();
            position = 0;
            while (true)
            {
                StiToken token = GetNextLexem();
                if (token == null) break;
                tokensList.Add(token);
            }
            tokensList = PostProcessTokensList(tokensList);
        }

        #endregion
    }
}
