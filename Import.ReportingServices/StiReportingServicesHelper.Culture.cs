#region Copyright (C) 2003-2026 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports												}
{	                         										}
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

using Stimulsoft.Report.Components;
using Stimulsoft.Report.Components.TextFormats;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts the language of a report and of its text boxes. Reporting Services formats the values
    /// of the texts by the language, Stimulsoft by the culture of the report.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts
        private static readonly Regex UserLanguageRegex = new Regex(@"^=\s*User\s*!\s*Language\s*$",
            RegexOptions.IgnoreCase);
        #endregion

        #region Fields.Culture
        /// <summary>
        /// The culture of the report when its language is a constant, null when it is the language
        /// of the viewer or an expression.
        /// </summary>
        private CultureInfo reportCulture;

        /// <summary>
        /// The languages of the text boxes. The format of a text box may come with the style of its
        /// text run after the language, so the language is applied when the report is built.
        /// </summary>
        private Dictionary<StiText, string> textLanguages = new Dictionary<StiText, string>();
        #endregion

        #region Culture
        /// <summary>
        /// Sets the culture of the report from its language. User!Language is the language of the
        /// viewer, which Stimulsoft renders in when the culture of the report is empty; another
        /// expression is computed when the report is rendered.
        /// </summary>
        private void ProcessReportLanguage(XmlNode node, StiReport report)
        {
            var language = GetNodeTextValue(node)?.Trim();
            if (string.IsNullOrEmpty(language) || UserLanguageRegex.IsMatch(language))
                return;

            if (IsExpression(language))
            {
                report.Culture = ConvertExpression(language, null).Trim();
                return;
            }

            reportCulture = GetCulture(language);
            if (reportCulture == null)
            {
                ThrowError(null, null, $"Language not supported: {language}");
                return;
            }

            report.Culture = reportCulture.Name;
        }

        /// <summary>
        /// Processes the style of a text box. Its text runs take the language of their own style or
        /// of the report, Reporting Services does not use the language of the text box for them.
        /// </summary>
        private void ProcessTextboxStyle(XmlNode node, XmlNode textboxNode, StiText text, string dataSet)
        {
            if (textboxNode["Paragraphs"] == null)
            {
                ProcessStyleType(node, text, dataSet);
                return;
            }

            var hasRunLanguage = textLanguages.TryGetValue(text, out var runLanguage);

            ProcessStyleType(node, text, dataSet);

            if (hasRunLanguage)
                textLanguages[text] = runLanguage;

            else
                textLanguages.Remove(text);
        }

        /// <summary>
        /// Applies the languages of the text boxes. Only a number, currency or percentage format can
        /// carry a culture of its own: it takes the separators, the patterns and the symbol of the
        /// language, a standard format of a date takes its pattern. A text box in the language of the
        /// report needs nothing.
        /// </summary>
        private void ApplyTextLanguages()
        {
            foreach (var pair in textLanguages)
            {
                ApplyTextLanguage(pair.Key, pair.Value);
            }

            textLanguages.Clear();
        }

        private void ApplyTextLanguage(StiText text, string language)
        {
            if (string.IsNullOrEmpty(language) || UserLanguageRegex.IsMatch(language))
                return;

            if (IsExpression(language))
            {
                ThrowError(null, null, $"Language given by an expression not applied: {text.Name}");
                return;
            }

            var culture = GetCulture(language);
            if (culture == null)
            {
                ThrowError(null, null, $"Language not supported: {language}");
                return;
            }

            if (reportCulture != null && culture.Name == reportCulture.Name)
                return;

            //a constant text is not formatted
            if (!IsSingleExpression(text.Text.Value))
                return;

            if (text.TextFormat is StiNumberFormatService format)
                ApplyCulture(format, culture.NumberFormat);

            else if (!ApplyDateCulture(text, culture))
                ThrowError(null, null, $"Language of a text box without a number format not applied: {text.Name}");
        }

        /// <summary>
        /// Writes the pattern of a standard format of a date in the language, d of de-DE turns into
        /// dd.MM.yyyy, so the order and the separators do not depend on the culture of the report. The
        /// names of the months and of the days still come in the language of the report.
        /// </summary>
        private bool ApplyDateCulture(StiText text, CultureInfo culture)
        {
            if (!(text.TextFormat is StiCustomFormatService format) || format.StringFormat == null)
                return false;

            var type = GetExpressionType(text.Text.Value.Substring(1, text.Text.Value.Length - 2));
            if (type != typeof(DateTime) && type != typeof(DateTime?))
                return false;

            var pattern = GetDatePattern(format.StringFormat.Trim(), culture.DateTimeFormat);
            if (pattern == null)
                return false;

            text.TextFormat = new StiCustomFormatService(pattern);

            if (pattern.Contains("MMM") || pattern.Contains("ddd"))
                ThrowError(null, null, $"Names of months and days use the language of the report: {text.Name}");

            return true;
        }

        private static string GetDatePattern(string format, DateTimeFormatInfo info)
        {
            switch (format)
            {
                case "d":
                    return info.ShortDatePattern;

                case "D":
                    return info.LongDatePattern;

                case "t":
                    return info.ShortTimePattern;

                case "T":
                    return info.LongTimePattern;

                case "f":
                    return $"{info.LongDatePattern} {info.ShortTimePattern}";

                case "F":
                    return info.FullDateTimePattern;

                case "g":
                    return $"{info.ShortDatePattern} {info.ShortTimePattern}";

                case "G":
                    return $"{info.ShortDatePattern} {info.LongTimePattern}";

                case "M":
                case "m":
                    return info.MonthDayPattern;

                case "Y":
                case "y":
                    return info.YearMonthPattern;

                default:
                    return null;
            }
        }

        private static void ApplyCulture(StiNumberFormatService format, NumberFormatInfo info)
        {
            int decimalDigits;

            if (format is StiPercentageFormatService percentage)
            {
                percentage.DecimalSeparator = info.PercentDecimalSeparator;
                percentage.GroupSeparator = info.PercentGroupSeparator;
                percentage.GroupSize = info.PercentGroupSizes[0];
                percentage.NegativePattern = info.PercentNegativePattern;
                percentage.PositivePattern = info.PercentPositivePattern;
                percentage.Symbol = info.PercentSymbol;
                percentage.State |= StiTextFormatState.PositivePattern | StiTextFormatState.PercentageSymbol;
                decimalDigits = info.PercentDecimalDigits;
            }

            else if (format is StiCurrencyFormatService currency)
            {
                currency.DecimalSeparator = info.CurrencyDecimalSeparator;
                currency.GroupSeparator = info.CurrencyGroupSeparator;
                currency.GroupSize = info.CurrencyGroupSizes[0];
                currency.NegativePattern = info.CurrencyNegativePattern;
                currency.PositivePattern = info.CurrencyPositivePattern;
                currency.Symbol = info.CurrencySymbol;
                currency.State |= StiTextFormatState.PositivePattern | StiTextFormatState.CurrencySymbol;
                decimalDigits = info.CurrencyDecimalDigits;
            }

            else
            {
                format.DecimalSeparator = info.NumberDecimalSeparator;
                format.GroupSeparator = info.NumberGroupSeparator;
                format.GroupSize = info.NumberGroupSizes[0];
                format.NegativePattern = info.NumberNegativePattern;
                decimalDigits = info.NumberDecimalDigits;
            }

            format.State |= StiTextFormatState.DecimalSeparator | StiTextFormatState.GroupSeparator |
                StiTextFormatState.GroupSize | StiTextFormatState.NegativePattern;

            if ((format.State & StiTextFormatState.DecimalDigits) == 0)
            {
                format.DecimalDigits = decimalDigits;
                format.State |= StiTextFormatState.DecimalDigits;
            }
        }

        /// <summary>
        /// The calendar of a text box. Stimulsoft formats the dates by the Gregorian calendar of the
        /// culture, another calendar is named in the list of errors. Old reports write 0 for the default.
        /// </summary>
        private void CheckCalendar(XmlNode baseNode, XmlNode node)
        {
            var calendar = GetNodeTextValue(node)?.Trim();
            if (string.IsNullOrEmpty(calendar) || calendar == "0" ||
                calendar.Equals("Default", StringComparison.OrdinalIgnoreCase) ||
                calendar.Equals("Gregorian", StringComparison.OrdinalIgnoreCase))
                return;

            ThrowError(baseNode.Name, node.Name);
        }

        /// <summary>
        /// The digits of a text box. Variant 1, the default, keeps the Latin digits; the national
        /// digits of the other variants are named in the list of errors.
        /// </summary>
        private void CheckNumeralVariant(XmlNode baseNode, XmlNode node)
        {
            var variant = GetNodeTextValue(node)?.Trim();
            if (string.IsNullOrEmpty(variant) || variant == "1")
                return;

            ThrowError(baseNode.Name, node.Name);
        }

        private static CultureInfo GetCulture(string name)
        {
            try
            {
                return CultureInfo.GetCultureInfo(name);
            }
            catch (CultureNotFoundException)
            {
                return null;
            }
        }
        #endregion
    }
}
