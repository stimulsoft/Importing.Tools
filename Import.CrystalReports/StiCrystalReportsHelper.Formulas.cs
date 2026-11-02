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
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Import.CrystalReports.Formulas;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Ras = CrystalDecisions.ReportAppServer.ReportDefModel;

namespace Import.CrystalReports
{
    /// <summary>
    /// The translation of the formulas of Crystal to the expressions of Stimulsoft. The formulas are collected while
    /// the page is converted and are translated at the end, when the bands of the groups are known.
    /// </summary>
    public sealed partial class StiCrystalReportsHelper
    {
        #region Consts.Formulas
        private const string FormatColor = "Color";
        private const string FormatBackColor = "BackColor";
        private const string FormatStyle = "Style";
        private const string FormatName = "FontName";
        private const string FormatSize = "FontSize";
        private const string FormatUnderline = "Underline";
        private const string FormatStrikeout = "Strikeout";
        #endregion

        #region Fields.Formulas
        private StiCrystalReportResolver resolver;
        private Dictionary<string, StiCrystalFormulaField> formulaFields;
        private List<Action> pendingTranslations;
        private Dictionary<string, StiCrystalReference> globalVariables;
        private StiPage formulaPage;
        #endregion

        #region Methods.Formulas
        /// <summary>
        /// Starts the formulas of a page: the main report and every subreport have their own fields and formulas.
        /// </summary>
        private void InitFormulas(StiPage page)
        {
            formulaPage = page;
            runningTotals = null;
            runningTotalExpressions = null;
            runningTotalVariables = null;
            resolver = new StiCrystalReportResolver();
            resolver.ReferenceResolving += Resolver_ReferenceResolving;
            resolver.VariableResolving = GetGlobalVariable;
            formulaFields = new Dictionary<string, StiCrystalFormulaField>(StringComparer.OrdinalIgnoreCase);
            pendingTranslations = new List<Action>();
        }

        /// <summary>
        /// Adds a formula field. Until the formula is translated its variable is a string with the text of the formula:
        /// the variable checks its value, so the declared type is set together with the translated expression.
        /// </summary>
        private void AddFormulaField(string name, string text, StiVariable variable, Type declaredType)
        {
            formulaFields[name] = new StiCrystalFormulaField(name, text, variable, declaredType);
            resolver.AddReference(name, new StiCrystalReference(variable.Name, declaredType));
        }

        /// <summary>
        /// Translates all the collected formulas: first the formula fields, then the conditions and the filters.
        /// </summary>
        private void ApplyPendingTranslations()
        {
            TranslateRunningTotals();

            foreach (var formula in formulaFields.Values)
            {
                TranslateFormulaField(formula);
            }

            foreach (var action in pendingTranslations)
            {
                action();
            }

            pendingTranslations.Clear();
        }

        /// <summary>
        /// Translates a formula field to the expression of its variable. A formula which can't be translated keeps
        /// its text as the value of the variable, and its reference is removed, so the formulas which use it are not
        /// translated either: an expression with a wrong reference would break the compilation of the report.
        /// </summary>
        private void TranslateFormulaField(StiCrystalFormulaField formula)
        {
            if (formula.IsTranslated || formula.IsTranslating)
                return;

            formula.IsTranslating = true;
            var variable = formula.Variable;
            var reference = resolver.GetReference(formula.Name);
            try
            {
                var translator = new StiCrystalTranslator(resolver);
                var declaredType = formula.DeclaredType == typeof(object) ? null : formula.DeclaredType;
                var expression = translator.Translate(formula.Text, declaredType);
                if (expression == null)
                    throw new StiCrystalFormulaException("The formula is empty");

                reference.Type = expression.Type;

                // A running function is calculated only in a text, so the variable keeps the text of the formula
                if (IsRunningExpression(expression.Code))
                {
                    variable.InitBy = StiVariableInitBy.Value;
                    variable.Type = typeof(string);
                    variable.Value = formula.Text.Replace("\r", "").Replace("\n", " ");
                }

                else
                {
                    variable.InitBy = StiVariableInitBy.Expression;
                    variable.Value = expression.Code;
                    variable.Type = expression.Type;
                }

                if (translator.HasSummaries)
                {
                    formula.InlineExpression = expression.Code;
                    reference.Expression = "(" + expression.Code + ")";
                    reference.HasSummaries = true;
                }
            }
            catch (StiCrystalFormulaException e)
            {
                resolver.RemoveReference(formula.Name);
                variable.InitBy = StiVariableInitBy.Value;
                variable.Type = typeof(string);
                variable.Value = formula.Text.Replace("\r", "").Replace("\n", " ");
                warningList.Add(StiWarning.GetWarningFormula(formula.Name, e.Message));
            }
            finally
            {
                formula.IsTranslating = false;
                formula.IsTranslated = true;
            }
        }

        /// <summary>
        /// Translates a formula which is not a formula field: a condition or a filter. Returns null and adds a warning
        /// if the formula can't be translated; returns null without a warning for an empty formula.
        /// </summary>
        private StiCrystalExpression TranslateFormula(string text, string place,
            StiCrystalReference currentFieldValue = null)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            resolver.CurrentFieldValue = GetInlineFieldValue(currentFieldValue);
            try
            {
                var expression = new StiCrystalTranslator(resolver).Translate(text);

                if (expression != null)
                    CheckNotRunning(expression);

                return expression;
            }
            catch (StiCrystalFormulaException e)
            {
                warningList.Add(StiWarning.GetWarningFormula(place, e.Message));
                return null;
            }
            finally
            {
                resolver.CurrentFieldValue = null;
            }
        }

        /// <summary>
        /// A running function works only in a text: a condition or a filter with it would not be calculated.
        /// </summary>
        private static void CheckNotRunning(StiCrystalExpression expression)
        {
            if (IsRunningExpression(expression.Code))
                throw new StiCrystalFormulaException("A running total can be used only in the text of a field");
        }

        /// <summary>
        /// The value of a field object with a formula which has summaries is the expression of the formula: the
        /// variable of such a formula doesn't keep its value.
        /// </summary>
        private StiCrystalReference GetInlineFieldValue(StiCrystalReference currentFieldValue)
        {
            if (currentFieldValue == null)
                return null;

            var formula = formulaFields.Values.FirstOrDefault(f => f.Variable.Name == currentFieldValue.Expression);
            if (formula?.InlineExpression == null)
                return currentFieldValue;

            var reference = resolver.GetReference(formula.Name);
            var type = reference?.Type ?? currentFieldValue.Type;
            return new StiCrystalReference("(" + formula.InlineExpression + ")", type);
        }

        /// <summary>
        /// The formula of the suppression becomes a condition which disables the component when it is true. The formula
        /// of a section or of an object overrides its own checkbox Suppress, so the component is enabled; the formula
        /// of an area doesn't override the checkbox of the section.
        /// </summary>
        private void AddSuppressCondition(StiComponent component, string text, string place,
            StiCrystalReference currentFieldValue = null, bool overridesSuppress = true)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            pendingTranslations.Add(() =>
                ApplySuppressCondition(component, text, place, currentFieldValue, overridesSuppress));
        }

        /// <summary>
        /// Crystal suppresses a section by the section itself and by its area: a band is hidden by the checkbox of
        /// either of them, and it gets the conditions of the formulas of both. A formula overrides the checkbox of its
        /// own level only.
        /// </summary>
        private void ApplySectionSuppress(StiBand band, Section section, Area area)
        {
            var sectionFormula = GetSuppressFormula(
                GetPropertyFromHiddenObject(section.SectionFormat, "RasSectionFormat") as Ras.ISCRSectionFormat);
            var areaFormula = area == null
                ? null
                : GetSuppressFormula(
                    GetPropertyFromHiddenObject(area.AreaFormat, "RasAreaFormat") as Ras.ISCRAreaFormat);

            var sectionSuppressed = section.SectionFormat.EnableSuppress && sectionFormula == null;
            var areaSuppressed = area != null && area.AreaFormat.EnableSuppress && areaFormula == null;
            if (sectionSuppressed || areaSuppressed)
            {
                band.Brush = new StiHatchBrush(HatchStyle.BackwardDiagonal, Color.DarkGray, Color.White);
                band.Enabled = false;

                var dataBand = band as StiDataBand;
                if (dataBand != null)
                    dataBand.CalcInvisible = true;

                return;
            }

            AddSuppressCondition(band, sectionFormula, band.Name);
            AddSuppressCondition(band, areaFormula, area?.Name, overridesSuppress: false);
        }

        private static Area GetArea(ReportDocument report, Section section)
        {
            foreach (Area area in report.ReportDefinition.Areas)
            {
                foreach (Section item in area.Sections)
                {
                    if (item.Name == section.Name)
                        return area;
                }
            }

            return null;
        }

        private static string GetSuppressFormula(Ras.ISCRSectionFormat format)
        {
            return GetSuppressFormula(format?.ConditionFormulas);
        }

        private static string GetSuppressFormula(Ras.ISCRAreaFormat format)
        {
            return GetSuppressFormula(format?.ConditionFormulas);
        }

        private static string GetSuppressFormula(Ras.SectionAreaFormatConditionFormulas formulas)
        {
            var text = formulas?[Ras.CrSectionAreaFormatConditionFormulaTypeEnum
                .crSectionAreaConditionFormulaTypeEnableSuppress]?.Text;
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        private void ApplySuppressCondition(StiComponent component, string text, string place,
            StiCrystalReference currentFieldValue, bool overridesSuppress)
        {
            var expression = TranslateFormula(text, place, currentFieldValue);
            if (expression == null)
                return;

            if (expression.Type != typeof(bool))
            {
                warningList.Add(StiWarning.GetWarningFormula(place, "The condition is not boolean"));
                return;
            }

            // When the formula is set, Crystal ignores the checkbox Suppress, but the engine still reports
            // the checkbox: the component is enabled and the condition decides
            if (overridesSuppress)
                component.Enabled = true;

            component.Conditions.Add(new StiCondition
            {
                Item = StiFilterItem.Expression,
                Expression = expression.Code,
                Enabled = false,
                Permissions = StiConditionPermissions.None
            });

            var dataBand = component as StiDataBand;
            if (dataBand != null)
                dataBand.CalcInvisible = true;
        }

        /// <summary>
        /// The record selection formula of Crystal becomes the filter of the data band.
        /// </summary>
        private void AddRecordSelection(StiDataBand dataBand, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            pendingTranslations.Add(() =>
            {
                var expression = TranslateFormula(text, "Record Selection");
                if (expression == null || expression.Type != typeof(bool))
                    return;

                dataBand.Filters.Add(new StiFilter(expression.Code));
                dataBand.FilterOn = true;
            });
        }

        /// <summary>
        /// A field object with a formula which has summaries shows the expression itself, not the variable.
        /// </summary>
        private void AddFormulaFieldText(StiText text, string formulaName)
        {
            pendingTranslations.Add(() =>
            {
                StiCrystalFormulaField formula;
                if (formulaFields.TryGetValue(formulaName, out formula) && formula.InlineExpression != null)
                    text.Text = "{" + formula.InlineExpression + "}";
            });
        }

        /// <summary>
        /// CurrentFieldValue of a formula of a field object is the expression of the field.
        /// </summary>
        private static StiCrystalReference GetCurrentFieldValue(ReportObject obj, StiComponent component)
        {
            var fieldObject = obj as FieldObject;
            var text = (component as StiText)?.Text?.Value;
            if (fieldObject == null || text == null || !text.StartsWith("{") || !text.EndsWith("}") ||
                text.IndexOf('{', 1) >= 0)
                return null;

            try
            {
                var expression = text.Substring(1, text.Length - 2);
                return new StiCrystalReference(expression, ConvertType(fieldObject.DataSource.ValueType));
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// The type of the variable of a parameter with a range of values.
        /// </summary>
        private static Type GetRangeType(Type type)
        {
            if (type == typeof(DateTime))
                return typeof(DateTimeRange);

            if (type == typeof(double))
                return typeof(DoubleRange);

            if (type == typeof(decimal))
                return typeof(DecimalRange);

            return typeof(StringRange);
        }

        /// <summary>
        /// The type of the variable of a parameter with several values.
        /// </summary>
        private static Type GetListType(Type type)
        {
            if (type == typeof(DateTime))
                return typeof(DateTimeList);

            if (type == typeof(double))
                return typeof(DoubleList);

            if (type == typeof(decimal))
                return typeof(DecimalList);

            if (type == typeof(bool))
                return typeof(BoolList);

            return typeof(StringList);
        }

        /// <summary>
        /// A global or shared variable of the formulas becomes a variable of the dictionary which the formulas assign;
        /// the same name gives the same variable in the report and in its subreports.
        /// </summary>
        private StiCrystalReference GetGlobalVariable(string name, Type type)
        {
            if (globalVariables == null)
                globalVariables = new Dictionary<string, StiCrystalReference>(StringComparer.OrdinalIgnoreCase);

            StiCrystalReference reference;
            if (globalVariables.TryGetValue(name, out reference))
                return reference;

            var variableName = StiNameValidator.CorrectName(name);
            var baseName = variableName;
            var index = 1;
            while (stimulReport.Dictionary.Variables[variableName] != null ||
                stimulReport.Dictionary.DataSources[variableName] != null)
            {
                variableName = baseName + "Var" + (index > 1 ? index.ToString() : string.Empty);
                index++;
            }

            var variable = new StiVariable("Crystal Variables")
            {
                Name = variableName,
                Alias = name,
                Type = type,
                ReadOnly = false,
                InitBy = StiVariableInitBy.Value
            };

            if (type == typeof(double) || type == typeof(decimal))
                variable.Value = "0";

            else if (type == typeof(bool))
                variable.Value = "False";

            else if (type == typeof(string))
                variable.Value = string.Empty;

            stimulReport.Dictionary.Variables.Add(variable);

            reference = new StiCrystalReference(variableName, type);
            globalVariables[name] = reference;
            return reference;
        }

        /// <summary>
        /// The formulas of the color, the font and the background of a field object become the conditions which
        /// apply the values of the formula.
        /// </summary>
        private void AddFormatConditions(StiComponent component, ReportObject obj, StiCrystalReference currentFieldValue)
        {
            var text = component as StiText;
            var rasField = GetRasReportObject(obj) as Ras.ISCRFieldObject;
            if (text == null || rasField == null)
                return;

            var fontFormulas = rasField.FontColor?.ConditionFormulas;
            if (fontFormulas != null)
            {
                AddFormatCondition(text, FormatColor,
                    fontFormulas[Ras.CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeColor]?.Text,
                    currentFieldValue);

                AddFormatCondition(text, FormatStyle,
                    fontFormulas[Ras.CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeStyle]?.Text,
                    currentFieldValue);

                AddFormatCondition(text, FormatName,
                    fontFormulas[Ras.CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeName]?.Text,
                    currentFieldValue);

                AddFormatCondition(text, FormatSize,
                    fontFormulas[Ras.CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeSize]?.Text,
                    currentFieldValue);

                AddFormatCondition(text, FormatUnderline,
                    fontFormulas[Ras.CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeUnderline]?.Text,
                    currentFieldValue);

                AddFormatCondition(text, FormatStrikeout,
                    fontFormulas[Ras.CrFontColorConditionFormulaTypeEnum.crFontColorConditionFormulaTypeStrikeout]?.Text,
                    currentFieldValue);
            }

            var borderFormulas = rasField.Border?.ConditionFormulas;
            if (borderFormulas != null)
            {
                var backColor = Ras.CrBorderConditionFormulaTypeEnum.crBorderConditionFormulaTypeBackgroundColor;
                AddFormatCondition(text, FormatBackColor, borderFormulas[backColor]?.Text, currentFieldValue);
            }
        }

        private void AddFormatCondition(StiText text, string kind, string formula, StiCrystalReference currentFieldValue)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return;

            pendingTranslations.Add(() => ApplyFormatCondition(text, kind, formula, currentFieldValue));
        }

        /// <summary>
        /// A formula of formatting gives a value, a condition of Stimulsoft applies a fixed style. So every constant
        /// value of the formula gets its own condition "the formula gives this value", and the condition applies
        /// this value only, through the permissions of the condition. DefaultAttribute keeps the style of the text.
        /// </summary>
        private void ApplyFormatCondition(StiText text, string kind, string formula,
            StiCrystalReference currentFieldValue)
        {
            var place = $"{text.Name}.{kind}";
            resolver.CurrentFieldValue = GetInlineFieldValue(currentFieldValue);
            try
            {
                var node = StiCrystalParser.Parse(formula);
                var translator = new StiCrystalTranslator(resolver);
                if (kind == FormatUnderline || kind == FormatStrikeout)
                {
                    var expression = translator.Translate(node);
                    if (expression == null || expression.Type != typeof(bool))
                        throw new StiCrystalFormulaException("The formula is not boolean");

                    CheckNotRunning(expression);

                    var style = kind == FormatUnderline ? FontStyle.Underline : FontStyle.Strikeout;
                    text.Conditions.Add(new StiCondition
                    {
                        Item = StiFilterItem.Expression,
                        Expression = expression.Code,
                        Font = new Font(text.Font, text.Font.Style | style),
                        Permissions = kind == FormatUnderline
                            ? StiConditionPermissions.FontStyleUnderline
                            : StiConditionPermissions.FontStyleStrikeout
                    });
                    return;
                }

                var values = StiCrystalTranslator.GetResultConstants(node);
                if (values.Count == 0)
                    throw new StiCrystalFormulaException("The formula has no constant values");

                foreach (var value in values)
                {
                    if (!(value is string) && (double)value == StiCrystalFormatConstants.DefaultAttribute)
                        continue;

                    var condition = CreateFormatCondition(text, kind, value);
                    if (condition == null)
                        continue;

                    var expression = translator.TranslateEquals(node, value);
                    if (expression == null || expression.Code == "false")
                        continue;

                    CheckNotRunning(expression);
                    condition.Item = StiFilterItem.Expression;
                    condition.Expression = expression.Code;
                    text.Conditions.Add(condition);
                }
            }
            catch (StiCrystalFormulaException e)
            {
                warningList.Add(StiWarning.GetWarningFormula(place, e.Message));
            }
            finally
            {
                resolver.CurrentFieldValue = null;
            }
        }

        /// <summary>
        /// Creates the condition which applies one value of a formula of formatting.
        /// </summary>
        private static StiCondition CreateFormatCondition(StiText text, string kind, object value)
        {
            var font = text.Font;
            switch (kind)
            {
                case FormatColor:
                    return new StiCondition
                    {
                        TextColor = StiCrystalFormatConstants.ToColor((int)(double)value),
                        Permissions = StiConditionPermissions.TextColor
                    };

                case FormatBackColor:
                    return new StiCondition
                    {
                        BackColor = StiCrystalFormatConstants.ToColor((int)(double)value),
                        Permissions = StiConditionPermissions.BackColor
                    };

                case FormatStyle:
                    var style = StiCrystalFormatConstants.ToFontStyle((int)(double)value) |
                        (font.Style & (FontStyle.Underline | FontStyle.Strikeout));

                    return new StiCondition
                    {
                        Font = new Font(font, style),
                        Permissions = StiConditionPermissions.FontStyleBold | StiConditionPermissions.FontStyleItalic
                    };

                case FormatName:
                    var name = value as string;
                    if (string.IsNullOrWhiteSpace(name))
                        return null;

                    return new StiCondition
                    {
                        Font = new Font(name, font.Size, font.Style),
                        Permissions = StiConditionPermissions.Font
                    };

                case FormatSize:
                    var size = value is string ? 0 : (float)(double)value;
                    if (size <= 0)
                        return null;

                    return new StiCondition
                    {
                        Font = new Font(font.FontFamily, size, font.Style),
                        Permissions = StiConditionPermissions.FontSize
                    };
            }

            return null;
        }

        /// <summary>
        /// Returns the object of the RAS behind a report object; the property is private in the base class.
        /// </summary>
        private static Ras.ISCRReportObject GetRasReportObject(ReportObject obj)
        {
            return GetRasObject(obj, "RasReportObject") as Ras.ISCRReportObject;
        }

        /// <summary>
        /// Reads the formula of the suppression of a report object through the object model of the RAS.
        /// </summary>
        private static string GetSuppressFormula(ReportObject obj)
        {
            var rasObject = GetRasReportObject(obj);
            var formulas = rasObject?.Format?.ConditionFormulas;
            if (formulas == null)
                return null;

            var formula = formulas[
                Ras.CrObjectFormatConditionFormulaTypeEnum.crObjectFormatConditionFormulaTypeEnableSuppress];
            return formula?.Text;
        }
        #endregion

        #region Handlers.Formulas
        private void Resolver_ReferenceResolving(string name)
        {
            StiCrystalFormulaField formula;
            if (formulaFields.TryGetValue(name, out formula))
                TranslateFormulaField(formula);
        }
        #endregion
    }
}
