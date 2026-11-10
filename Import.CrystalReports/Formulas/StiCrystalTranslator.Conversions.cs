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
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// The functions of Crystal which convert values between the kinds of the formulas: Join of a parameter with
    /// several values, Picture by a constant mask, Time and TimeValue, ShiftDateTime between UTC and the local
    /// time. On the corpus of 2026-11-10 they failed 13 to 57 formulas each.
    /// </summary>
    public partial class StiCrystalTranslator
    {
        #region Consts.Conversions
        /// <summary>
        /// The length which takes the rest of a string in Mid.
        /// </summary>
        private const int RestLength = 1000000;
        #endregion

        #region Methods.Conversions
        private StiCrystalExpression TranslateConversionFunction(string name, List<StiCrystalExpression> args,
            int position)
        {
            switch (name)
            {
                case "join":
                    return TranslateJoin(args);

                case "picture":
                    return TranslatePicture(args);

                case "time":
                case "timevalue":
                case "ctime":
                    return TranslateTime(args);

                case "shiftdatetime":
                    return TranslateShiftDateTime(args);

                // The time zone of the printing and of the user is the local time zone of the computer
                case "printtimezone":
                case "currentceusertimezone":
                    if (args.Count == 0)
                        return new StiCrystalExpression("\"\"", typeof(string), isAtomic: true) { Literal = "" };

                    break;
            }

            return null;
        }

        /// <summary>
        /// Join of the values of a parameter: ToQueryString of the list without quotes joins them by ", ".
        /// </summary>
        private static StiCrystalExpression TranslateJoin(List<StiCrystalExpression> args)
        {
            var list = args.Count > 0 ? args[0] : null;
            if (list?.Reference == null || !list.Reference.IsMultiValue)
                return null;

            var delimiter = args.Count > 1 ? args[1].Literal : " ";
            if (delimiter == null)
                return null;

            var code = $"{list.ToOperand()}.ToQueryString(\"\")";
            if (delimiter != ", ")
                code += $".Replace(\", \", {ToCSharpString(delimiter)})";

            return new StiCrystalExpression(code, typeof(string), isAtomic: true);
        }

        /// <summary>
        /// Picture(text, mask) with a constant mask: every x takes the next character of the text, the other
        /// characters of the mask are copied, the characters of the text beyond the mask go to the end.
        /// </summary>
        private static StiCrystalExpression TranslatePicture(List<StiCrystalExpression> args)
        {
            if (args.Count != 2 || args[0].Type != typeof(string) || string.IsNullOrEmpty(args[1].Literal))
                return null;

            var mask = args[1].Literal;
            var parts = new List<string>();
            var literal = new StringBuilder();
            var position = 0;
            var index = 0;
            while (index < mask.Length)
            {
                if (char.ToLowerInvariant(mask[index]) != 'x')
                {
                    literal.Append(mask[index++]);
                    continue;
                }

                if (literal.Length > 0)
                {
                    parts.Add(ToCSharpString(literal.ToString()));
                    literal.Clear();
                }

                var start = index;
                while (index < mask.Length && char.ToLowerInvariant(mask[index]) == 'x')
                {
                    index++;
                }

                parts.Add($"Mid({args[0].Code}, {position}, {index - start})");
                position += index - start;
            }

            if (literal.Length > 0)
                parts.Add(ToCSharpString(literal.ToString()));

            parts.Add($"Mid({args[0].Code}, {position}, {RestLength})");
            return new StiCrystalExpression(string.Join(" + ", parts), typeof(string));
        }

        /// <summary>
        /// Time(hours, minutes, seconds) and the time of a date and time; a time is TimeSpan, as the time fields.
        /// </summary>
        private static StiCrystalExpression TranslateTime(List<StiCrystalExpression> args)
        {
            if (args.Count == 3 && args.All(a => IsNumber(a.Type)))
                return Call("TimeSerial", typeof(TimeSpan), args.Select(a => Coerce(a, typeof(long))).ToArray());

            if (args.Count == 1 && args[0].Type == typeof(TimeSpan))
                return args[0];

            if (args.Count == 1 && IsDate(args[0].Type))
            {
                var value = args[0].ToOperand();
                return new StiCrystalExpression($"TimeSerial({value}.Hour, {value}.Minute, {value}.Second)",
                    typeof(TimeSpan), isAtomic: true);
            }

            return null;
        }

        /// <summary>
        /// ShiftDateTime(value, from, to) with constant time zones. Only UTC and the local time zone of the
        /// computer, the empty string, PrintTimeZone or CurrentCEUserTimeZone, are known: a zone of a parameter,
        /// like "Pacific,480", has no equivalent, the function stays unsupported.
        /// </summary>
        private static StiCrystalExpression TranslateShiftDateTime(List<StiCrystalExpression> args)
        {
            if (args.Count != 3 || !IsDate(args[0].Type))
                return null;

            var from = GetTimeZoneKind(args[1]);
            var to = GetTimeZoneKind(args[2]);
            if (from == null || to == null)
                return null;

            if (from == to)
                return args[0];

            return from == "utc"
                ? Call("UtcToLocalTime", typeof(DateTime), args[0])
                : Call("LocalTimeToUtc", typeof(DateTime), args[0]);
        }

        /// <summary>
        /// "utc" for the zone UTC, "local" for the local time zone, null for another zone.
        /// </summary>
        private static string GetTimeZoneKind(StiCrystalExpression zone)
        {
            if (zone.Literal == null)
                return null;

            var name = zone.Literal.Trim();
            if (name.Length == 0)
                return "local";

            var comma = name.IndexOf(',');
            var zoneName = (comma < 0 ? name : name.Substring(0, comma)).Trim();
            var offset = comma < 0 ? "0" : name.Substring(comma + 1).Split(',')[0].Trim();
            return (zoneName.Equals("UTC", StringComparison.OrdinalIgnoreCase) ||
                zoneName.Equals("GMT", StringComparison.OrdinalIgnoreCase)) && offset == "0"
                ? "utc"
                : null;
        }
        #endregion
    }
}
