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
using System.Text;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.TextFormats;
using Ras = CrystalDecisions.ReportAppServer.ReportDefModel;
using RasData = CrystalDecisions.ReportAppServer.DataDefModel;

namespace Import.CrystalReports
{
    /// <summary>
    /// The formats of the field objects. A field which uses the system default of Crystal shows its value by
    /// the settings of the computer, so it gets a format with the local settings, and only the number of decimals
    /// is fixed; the engine reports the separators and the currency of the computer of the import, which must not
    /// be fixed in the report. A field with its own format keeps its separators, its currency and its patterns.
    /// </summary>
    public sealed partial class StiCrystalReportsHelper
    {
        #region Consts.Formats
        /// <summary>
        /// The patterns of the negative numbers of .NET: -n, n- and (n).
        /// </summary>
        private const int NegativeLeadingMinus = 1;

        private const int NegativeTrailingMinus = 3;

        private const int NegativeBracketed = 0;
        #endregion

        #region Methods.Formats
        private void ApplyFieldFormat(StiText text, ReportObject obj)
        {
            var field = GetRasReportObject(obj) as Ras.ISCRFieldObject;
            var format = field?.FieldFormat;
            if (format == null)
                return;

            WarnFormatFormulas(obj.Name, format.NumericFormat?.ConditionFormulas);

            var isDefault = format.CommonFormat == null || format.CommonFormat.EnableSystemDefault;
            switch (field.FieldValueType)
            {
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeNumberField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeCurrencyField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt8sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt8uField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt16sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt16uField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt32sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt32uField:
                    ApplyNumberFormat(text, format.NumericFormat, isDefault,
                        field.FieldValueType == RasData.CrFieldValueTypeEnum.crFieldValueTypeCurrencyField);
                    break;

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeDateField:
                    text.TextFormat = new StiDateFormatService(GetDatePattern(format.DateFormat, isDefault), "");
                    break;

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeDateTimeField:
                    text.TextFormat = new StiDateFormatService(GetDateTimePattern(format, isDefault), "");
                    break;

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeTimeField:
                    text.TextFormat = new StiDateFormatService(GetTimePattern(format.TimeFormat, isDefault), "");
                    break;

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeBooleanField:
                    ApplyBooleanFormat(text, format.BooleanFormat);
                    break;
            }
        }

        /// <summary>
        /// Crystal prints nothing for a NULL value of a database field, Stimulsoft converts it to 0 or to the empty
        /// date. NullValue of a text of a data column shows instead of NULL in both calculation modes; it must not be
        /// empty, so it is a space. A string is empty anyway, and a report which converts NULL to the default value
        /// prints 0 like Stimulsoft.
        /// </summary>
        private void ApplyNullValue(StiText text, FieldObject obj)
        {
            if (obj.DataSource.Kind != FieldKind.DatabaseField || obj.DataSource.ValueType == FieldValueType.StringField)
                return;

            if (ConvertsNullsToDefault(mainDocument))
                return;

            var value = text.Text?.Value;
            if (string.IsNullOrEmpty(value) || !value.StartsWith("{") || !value.EndsWith("}"))
                return;

            text.Type = StiSystemTextType.DataColumn;
            text.NullValue = " ";
        }

        /// <summary>
        /// Crystal may calculate the settings of a number format by formulas, for example the number of decimals
        /// from the data; a format of Stimulsoft is fixed, so such a field only gets a warning.
        /// </summary>
        private void WarnFormatFormulas(string objectName, Ras.NumericFieldFormatConditionFormulas formulas)
        {
            if (formulas == null)
                return;

            foreach (Ras.CrNumericFieldFormatConditionFormulaTypeEnum kind in
                Enum.GetValues(typeof(Ras.CrNumericFieldFormatConditionFormulaTypeEnum)))
            {
                var formula = formulas[kind];
                if (formula == null || string.IsNullOrWhiteSpace(formula.Text))
                    continue;

                var property = kind.ToString().Replace("crNumericFieldFormatConditionFormulaType", string.Empty);
                warningList.Add(StiWarning.GetWarningFormat(objectName, property, formula.Text));
            }
        }

        private static void ApplyNumberFormat(StiText text, Ras.NumericFieldFormat number, bool isDefault,
            bool isCurrencyField)
        {
            if (number == null)
                return;

            if (number.EnableSuppressIfZero)
                text.HideZeros = true;

            var hasSymbol = number.CurrencySymbolFormat != Ras.CrCurrencySymbolTypeEnum.crCurrencySymbolTypeNoSymbol;
            if (isDefault)
            {
                text.TextFormat = isCurrencyField || hasSymbol
                    ? (StiFormatService)new StiCurrencyFormatService
                    {
                        UseLocalSetting = true,
                        State = StiTextFormatState.DecimalDigits,
                        DecimalDigits = number.NDecimalPlaces
                    }
                    : new StiNumberFormatService
                    {
                        UseLocalSetting = true,
                        State = StiTextFormatState.DecimalDigits,
                        DecimalDigits = number.NDecimalPlaces
                    };
                return;
            }

            if (!hasSymbol)
            {
                text.TextFormat = new StiNumberFormatService
                {
                    UseLocalSetting = false,
                    DecimalDigits = number.NDecimalPlaces,
                    DecimalSeparator = number.DecimalSymbol,
                    GroupSeparator = number.ThousandSymbol,
                    GroupSize = 3,
                    UseGroupSeparator = number.ThousandsSeparator,
                    NegativePattern = GetNegativePattern(number.NegativeFormat)
                };
                return;
            }

            var leading = number.CurrencyPosition ==
                Ras.CrCurrencyPositionFormatEnum.crCurrencyPositionFormatLeadingCurrencyInsideNegative ||
                number.CurrencyPosition ==
                Ras.CrCurrencyPositionFormatEnum.crCurrencyPositionFormatLeadingCurrencyOutsideNegative;

            text.TextFormat = new StiCurrencyFormatService
            {
                UseLocalSetting = false,
                DecimalDigits = number.NDecimalPlaces,
                DecimalSeparator = number.DecimalSymbol,
                GroupSeparator = number.ThousandSymbol,
                GroupSize = 3,
                UseGroupSeparator = number.ThousandsSeparator,
                Symbol = number.CurrencySymbol,
                PositivePattern = leading ? 0 : 1,
                NegativePattern = GetCurrencyNegativePattern(number.NegativeFormat, leading)
            };
        }

        private static int GetNegativePattern(Ras.CrNegativeTypeEnum negative)
        {
            switch (negative)
            {
                case Ras.CrNegativeTypeEnum.crNegativeTypeTrailingMinus:
                    return NegativeTrailingMinus;

                case Ras.CrNegativeTypeEnum.crNegativeTypeBracketed:
                    return NegativeBracketed;

                default:
                    return NegativeLeadingMinus;
            }
        }

        /// <summary>
        /// The patterns of the negative currency of .NET: ($n) (n$) -$n -n$ $n- n$-.
        /// </summary>
        private static int GetCurrencyNegativePattern(Ras.CrNegativeTypeEnum negative, bool leading)
        {
            switch (negative)
            {
                case Ras.CrNegativeTypeEnum.crNegativeTypeBracketed:
                    return leading ? 0 : 4;

                case Ras.CrNegativeTypeEnum.crNegativeTypeTrailingMinus:
                    return leading ? 3 : 7;

                default:
                    return leading ? 1 : 5;
            }
        }

        /// <summary>
        /// The .NET pattern of a date: the order, the parts, the separators and the day of the week of Crystal.
        /// </summary>
        private static string GetDatePattern(Ras.DateFieldFormat date, bool isDefault)
        {
            if (date == null || isDefault ||
                date.SystemDefaultType != Ras.CrDateSystemDefaultTypeEnum.crDateSystemDefaultTypeNotUsingDefaults)
            {
                return date != null &&
                    date.SystemDefaultType == Ras.CrDateSystemDefaultTypeEnum.crDateSystemDefaultTypeUseLongDate
                    ? "D"
                    : "d";
            }

            var year = date.YearFormat == Ras.CrYearFormatEnum.crYearFormatLongYear ? "yyyy"
                : date.YearFormat == Ras.CrYearFormatEnum.crYearFormatShortYear ? "yy" : null;

            var month = GetMonthPattern(date.MonthFormat);
            var day = date.DayFormat == Ras.CrDayFormatEnum.crDayFormatLeadingZeroNumericDay ? "dd"
                : date.DayFormat == Ras.CrDayFormatEnum.crDayFormatNumericDay ? "d" : null;

            string[] parts;
            switch (date.DateOrder)
            {
                case Ras.CrDateOrderEnum.crDateOrderYearMonthDay:
                    parts = new[] { year, month, day };
                    break;

                case Ras.CrDateOrderEnum.crDateOrderDayMonthYear:
                    parts = new[] { day, month, year };
                    break;

                default:
                    parts = new[] { month, day, year };
                    break;
            }

            var separators = new[] { date.DateFirstSeparator, date.DateSecondSeparator };
            var pattern = new StringBuilder();
            var index = 0;
            foreach (var part in parts)
            {
                if (part == null)
                    continue;

                if (pattern.Length > 0)
                    pattern.Append(Quote(separators[index++ % 2]));

                pattern.Append(part);
            }

            return AddDayOfWeek(pattern.ToString(), date);
        }

        private static string GetMonthPattern(Ras.CrMonthFormatEnum month)
        {
            switch (month)
            {
                case Ras.CrMonthFormatEnum.crMonthFormatNumericMonth:
                    return "M";

                case Ras.CrMonthFormatEnum.crMonthFormatLeadingZeroNumericMonth:
                    return "MM";

                case Ras.CrMonthFormatEnum.crMonthFormatShortMonth:
                    return "MMM";

                case Ras.CrMonthFormatEnum.crMonthFormatLongMonth:
                    return "MMMM";

                default:
                    return null;
            }
        }

        private static string AddDayOfWeek(string pattern, Ras.DateFieldFormat date)
        {
            if (date.DayOfWeekType == Ras.CrDayOfWeekTypeEnum.crDayOfWeekTypeNoDayOfWeek)
                return pattern;

            var dayOfWeek = date.DayOfWeekType == Ras.CrDayOfWeekTypeEnum.crDayOfWeekTypeLong ? "dddd" : "ddd";
            var separator = Quote(date.DayOfWeekSeparator);
            return date.DayOfWeekPosition == Ras.CrDayOfWeekPositionEnum.crDayOfWeekPositionTrailing
                ? pattern + separator + dayOfWeek
                : dayOfWeek + separator + pattern;
        }

        private static string GetTimePattern(Ras.TimeFieldFormat time, bool isDefault)
        {
            if (time == null || isDefault)
                return "T";

            var is12Hour = time.TimeBase == Ras.CrTimeBaseEnum.crTimeBase12Hour;
            var hour = time.HourFormat == Ras.CrHourFormatEnum.crHourFormatNoHour ? null
                : time.HourFormat == Ras.CrHourFormatEnum.crHourFormatNumericHour
                    ? (is12Hour ? "hh" : "HH")
                    : (is12Hour ? "h" : "H");

            var minute = time.MinuteFormat == Ras.CrMinuteFormatEnum.crMinuteFormatNoMinute ? null
                : time.MinuteFormat == Ras.CrMinuteFormatEnum.crMinuteFormatNumericMinute ? "mm" : "m";

            var second = time.SecondFormat == Ras.CrSecondFormatEnum.crSecondFormatNoSecond ? null
                : time.SecondFormat == Ras.CrSecondFormatEnum.crSecondFormatNumericSecond ? "ss" : "s";

            var pattern = hour ?? string.Empty;
            if (minute != null)
                pattern += (pattern.Length > 0 ? Quote(time.HourMinuteSeparator) : string.Empty) + minute;

            if (second != null)
                pattern += (pattern.Length > 0 ? Quote(time.MinuteSecondSeparator) : string.Empty) + second;

            if (is12Hour)
                pattern = time.AMPMFormat == Ras.CrAMPMFormatEnum.crAMPMFormatBefore ? "tt " + pattern : pattern + " tt";

            return pattern.Length > 0 ? pattern : "T";
        }

        private static string GetDateTimePattern(Ras.FieldFormat format, bool isDefault)
        {
            var order = format.DateTimeFormat?.DateTimeOrder ?? Ras.CrDateTimeOrderEnum.crDateTimeOrderDateThenTime;
            switch (order)
            {
                case Ras.CrDateTimeOrderEnum.crDateTimeOrderDateOnly:
                    return GetDatePattern(format.DateFormat, isDefault);

                case Ras.CrDateTimeOrderEnum.crDateTimeOrderTimeOnly:
                    return GetTimePattern(format.TimeFormat, isDefault);
            }

            if (isDefault)
                return "G";

            var separator = Quote(format.DateTimeFormat?.DateTimeSeparator ?? " ");
            var datePattern = GetDatePattern(format.DateFormat, false);
            var timePattern = GetTimePattern(format.TimeFormat, false);
            return order == Ras.CrDateTimeOrderEnum.crDateTimeOrderTimeThenDate
                ? timePattern + separator + datePattern
                : datePattern + separator + timePattern;
        }

        private static void ApplyBooleanFormat(StiText text, Ras.BooleanFieldFormat format)
        {
            if (format == null)
                return;

            string trueText;
            string falseText;
            switch (format.OutputFormat)
            {
                case Ras.CrBooleanOutputFormatEnum.crBooleanOutputFormatTOrF:
                    trueText = "T";
                    falseText = "F";
                    break;

                case Ras.CrBooleanOutputFormatEnum.crBooleanOutputFormatYesOrNo:
                    trueText = "Yes";
                    falseText = "No";
                    break;

                case Ras.CrBooleanOutputFormatEnum.crBooleanOutputFormatYOrN:
                    trueText = "Y";
                    falseText = "N";
                    break;

                case Ras.CrBooleanOutputFormatEnum.crBooleanOutputFormatOneOrZero:
                    trueText = "1";
                    falseText = "0";
                    break;

                default:
                    trueText = "True";
                    falseText = "False";
                    break;
            }

            text.TextFormat = new StiBooleanFormatService
            {
                TrueDisplay = trueText,
                FalseDisplay = falseText
            };
        }

        /// <summary>
        /// Quotes the literal text of a date pattern, so the letters of a separator are not taken as a pattern.
        /// </summary>
        private static string Quote(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;

            return "'" + text.Replace("'", "\\'") + "'";
        }
        #endregion
    }
}
