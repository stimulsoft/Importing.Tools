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
using System.Linq;
using System.Text.RegularExpressions;
using CrystalDecisions.CrystalReports.Engine;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report;
using Stimulsoft.Report.Chart;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;
using Ras = CrystalDecisions.ReportAppServer.ReportDefModel;
using RasData = CrystalDecisions.ReportAppServer.DataDefModel;

namespace Import.CrystalReports
{
    /// <summary>
    /// The charts of Crystal. A chart of Crystal summarizes the records on change of its condition fields, the chart
    /// of Stimulsoft takes a point from every row of its data source. So every chart gets its own virtual source,
    /// which groups the records of the main source by the conditions and computes the summaries; a formula of Crystal
    /// comes into the virtual source as a calculated column of the main source. A chart in a group section shows
    /// the records of the current group only: the virtual source keeps the conditions of the groups too, and the
    /// chart filters its rows by the values of the groups.
    /// </summary>
    public sealed partial class StiCrystalReportsHelper
    {
        #region Consts.Charts
        private const string FormulaFieldsCategory = "Formula Fields";
        private const string RunningTotalFieldsCategory = "RunningTotal Fields";
        private const string CrystalVariablesCategory = "Crystal Variables";

        /// <summary>
        /// A call of a summary of Stimulsoft: Sum, cCount, MaxDate, SumIfRunning and so on, not Math.Max.
        /// </summary>
        private static readonly Regex ChartSummaryCall = new Regex(
            @"(?<![\w.])(c|p|col)?(Sum|SumDistinct|Count|CountDistinct|Avg|AvgDate|AvgTime|Min|Max|MinDate|MaxDate|" +
            @"MinTime|MaxTime|MinStr|MaxStr|First|Last|Median|Mode|Rank)(If)?(Running)?\s*\(",
            RegexOptions.Compiled);
        #endregion

        #region Fields.Charts
        private Dictionary<StiVirtualSource, StiChart> charts;
        private Dictionary<int, StiCrystalChartGroup> chartGroups;
        #endregion

        #region Methods.Charts
        /// <summary>
        /// Starts the charts of a page: the main report and every subreport have their own groups.
        /// </summary>
        private void InitCharts()
        {
            charts = new Dictionary<StiVirtualSource, StiChart>();
            chartGroups = new Dictionary<int, StiCrystalChartGroup>();
        }

        /// <summary>
        /// Keeps the condition of a group band for the charts in the sections of the group.
        /// </summary>
        private void AddChartGroup(int index, string condition, Type type)
        {
            if (string.IsNullOrEmpty(condition) || type == null || type == typeof(object))
                return;

            chartGroups[index] = new StiCrystalChartGroup(UnwrapExpression(condition), type);
        }

        /// <summary>
        /// The type of the condition of a group: a group in a specified order compares the keys of its named groups.
        /// </summary>
        private Type GetChartGroupType(Groups groups, int index, string field)
        {
            if (specifiedGroupNames.ContainsKey(field))
                return typeof(string);

            return index < groups.Count ? ConvertType(groups[index].ConditionField.ValueType) : null;
        }

        /// <summary>
        /// The record selection of Crystal filters the records of the charts too.
        /// </summary>
        private void AddChartRecordSelection(string expression)
        {
            foreach (var source in charts.Keys)
            {
                source.Filters.Add(new StiFilter(expression));
                source.FilterOn = true;
            }
        }

        private StiComponent ImportChartObject(ReportObject obj)
        {
            var rasChart = GetRasReportObject(obj) as Ras.ISCRChartObject;
            var definition = rasChart?.ChartDefinition;
            if (definition == null || mainDataSource == null)
                return SkipChart(obj, "Chart");

            var isRecordChart = definition.ChartType == Ras.CrChartTypeEnum.crChartTypeDetail &&
                (definition.ConditionFields == null || definition.ConditionFields.Count == 0);

            if (definition.ChartType != Ras.CrChartTypeEnum.crChartTypeDetail &&
                definition.ChartType != Ras.CrChartTypeEnum.crChartTypeGroup)
                return SkipChart(obj, definition.ChartType.ToString());

            var styleType = rasChart.ChartStyle?.Type ?? Ras.CrChartStyleTypeEnum.crChartStyleTypeBar;
            if (CreateChartSeries(styleType) == null)
                return SkipChart(obj, styleType.ToString());

            // The name of the object is only the prefix of the columns, the component takes the name itself
            var chartName = StiNameValidator.CorrectName(obj.Name);
            var source = new StiVirtualSource(mainDataSource.Name, CreateMemberName(chartName + "Data"));
            var groupColumns = new List<string>();
            var results = new List<string>();
            var filters = new List<StiFilter>();

            #region Groups
            var level = GetChartGroupLevel(rasChart);
            for (var index = 0; index <= level; index++)
            {
                StiCrystalChartGroup group;
                if (!chartGroups.TryGetValue(index, out group))
                    continue;

                var column = AddChartColumn(source, chartName + "Group" + (index + 1), group.Expression, group.Type);
                if (isRecordChart)
                    AddChartResult(results, column, string.Empty);

                else
                    groupColumns.Add(column);

                // The condition of a group in a specified order is a chain of ?:, which needs the parentheses
                filters.Add(new StiFilter($"{source.Name}.{column} == ({group.Expression})"));
            }
            #endregion

            #region Conditions
            string argumentColumn = null;
            string seriesColumn = null;

            if (isRecordChart)
            {
                argumentColumn = AddChartColumn(source, chartName + "Record", "Line", typeof(int));
                AddChartResult(results, argumentColumn, string.Empty);
            }
            else
            {
                var conditions = definition.ConditionFields.Cast<RasData.ISCRField>().ToList();
                var period = GetChartPeriod(definition);

                for (var index = 0; index < conditions.Count && index < 2; index++)
                {
                    var condition = conditions[index];
                    var expression = GetChartFieldExpression(condition.FormulaForm);
                    if (expression == null)
                        return SkipChart(obj, "Chart on " + condition.FormulaForm);

                    var type = GetChartFieldType(condition.Type);
                    if (index == 0 && period != null && type == typeof(DateTime))
                        expression = GetDatePeriod(expression, period.Value);

                    var column = AddChartColumn(source, chartName + (index == 0 ? "Argument" : "Series"),
                        expression, type);
                    groupColumns.Add(column);

                    if (index == 0)
                        argumentColumn = column;

                    else
                        seriesColumn = column;
                }
            }
            #endregion

            #region Values
            var pointValueCount = GetPointValueCount(styleType);
            var values = new List<StiCrystalChartValue>();
            foreach (RasData.ISCRField field in definition.DataFields)
            {
                var value = GetChartValue(field, isRecordChart);
                if (value == null)
                {
                    warningList.Add(StiWarning.GetWarningObject(obj.Name, "Chart value " + field.FormulaForm));
                    continue;
                }

                var name = $"{chartName}Value{values.Count + 1}";
                var column = GetChartMasterColumn(name, value.Expression, value.Type);
                AddChartResult(results, column, value.Function, name);
                source.Columns.Add(new StiDataColumn(name, value.ResultType));

                value.Column = name;
                values.Add(value);

                // A pie shows one value; a second condition divides the bars by its values, not by several values;
                // a point of a scatter takes its coordinates from the first values
                var isLastValue = pointValueCount > 0
                    ? values.Count == pointValueCount
                    : seriesColumn != null || IsPieSeries(styleType);

                if (isLastValue)
                    break;
            }

            if (values.Count == 0)
                return SkipChart(obj, "Chart without values");

            if (values.Count < pointValueCount)
                return SkipChart(obj, $"{styleType} with {values.Count} values");
            #endregion

            source.GroupColumns = groupColumns.ToArray();
            source.Results = results.ToArray();
            stimulReport.Dictionary.DataSources.Add(source);

            var chart = new StiChart
            {
                DataSourceName = source.Name
            };
            ConvertProperties(chart, obj);

            foreach (var filter in filters)
            {
                chart.Filters.Add(filter);
            }

            chart.FilterOn = chart.Filters.Count > 0;

            // The values of a scatter or a bubble are the coordinates of one series: X, Y and the size
            var seriesValues = pointValueCount > 0 ? values.Skip(1).Take(1).ToList() : values;
            foreach (var value in seriesValues)
            {
                var series = CreateChartSeries(styleType);
                series.ArgumentDataColumn = pointValueCount > 0
                    ? $"{source.Name}.{values[0].Column}"
                    : $"{source.Name}.{argumentColumn}";
                series.ValueDataColumn = $"{source.Name}.{value.Column}";
                series.CoreTitle = value.Title;

                var bubble = series as StiBubbleSeries;
                if (bubble != null)
                    bubble.WeightDataColumn = $"{source.Name}.{values[2].Column}";

                if (seriesColumn != null)
                {
                    series.AutoSeriesKeyDataColumn = $"{source.Name}.{seriesColumn}";
                    series.AutoSeriesTitleDataColumn = $"{source.Name}.{seriesColumn}";
                }

                chart.Series.Add(series);
            }

            // The legend of Crystal names the slices of a pie, Stimulsoft writes their values by default
            if (IsPieSeries(styleType))
                chart.SeriesLabels.LegendValueType = StiSeriesLabelsValueType.Argument;

            var textOptions = rasChart.ChartStyle?.TextOptions;
            if (!string.IsNullOrWhiteSpace(textOptions?.Title))
            {
                chart.Title.Text = textOptions.Title;
                chart.Title.Visible = true;
                ApplyChartTitleFont(chart.Title, textOptions.TitleFont);
            }

            charts[source] = chart;
            return chart;
        }

        /// <summary>
        /// Completes the charts when the formulas are translated. A calculated column with the variable of a formula
        /// takes the type of the variable: the type of the field in Crystal may differ from the translated one, and
        /// a formula which isn't translated keeps its text. A formula with summaries, a running total or a variable
        /// of Crystal has no value for a single record, a calculated column can't compute it, so such a chart is
        /// removed with a warning.
        /// </summary>
        private void CompleteCharts()
        {
            foreach (var pair in charts)
            {
                var source = pair.Key;
                var masterSource = stimulReport.Dictionary.DataSources[source.NameInSource];
                if (masterSource == null || CompleteChartColumns(source, masterSource))
                    continue;

                RemoveChart(pair.Value, source, masterSource);
            }
        }

        private bool CompleteChartColumns(StiVirtualSource source, StiDataSource masterSource)
        {
            foreach (var column in source.GroupColumns)
            {
                if (!CompleteChartColumn(source, masterSource, column, null, column))
                    return false;
            }

            for (var index = 0; index + 2 < source.Results.Length; index += 3)
            {
                if (!CompleteChartColumn(source, masterSource, source.Results[index], source.Results[index + 1],
                    source.Results[index + 2]))
                    return false;
            }

            return true;
        }

        /// <summary>
        /// Checks a column of the master source which a virtual source reads: a group or a record when the name of
        /// the result is the name of the column, otherwise a summarized value.
        /// </summary>
        private bool CompleteChartColumn(StiVirtualSource source, StiDataSource masterSource, string column,
            string function, string name)
        {
            var calcColumn = masterSource.Columns[column] as StiCalcDataColumn;
            if (calcColumn == null)
                return true;

            if (!IsRecordExpression(calcColumn.Expression))
                return false;

            var variable = stimulReport.Dictionary.Variables[calcColumn.Expression];
            if (variable != null)
                calcColumn.Type = variable.Type;

            var isValue = name != column;
            if (isValue && function != "Count" && function != "CountDistinct" && !IsNumericType(calcColumn.Type))
                return false;

            // A group and a value of a record are copied as they are, a summary keeps the type of its function
            var resultColumn = source.Columns[name];
            if (resultColumn != null && (!isValue || string.IsNullOrEmpty(function)))
                resultColumn.Type = calcColumn.Type;

            return true;
        }

        /// <summary>
        /// True if the expression has a value for a single record: no summaries in it and in the formulas which it
        /// uses, every formula is translated, no running totals and variables of Crystal. A formula with a summary
        /// inside another formula isn't marked, so the expressions of the formulas are checked too (G2032: MaxDate).
        /// </summary>
        private bool IsRecordExpression(string expression)
        {
            return IsRecordExpression(expression, new HashSet<StiVariable>());
        }

        private bool IsRecordExpression(string expression, HashSet<StiVariable> passedVariables)
        {
            if (ChartSummaryCall.IsMatch(expression))
                return false;

            foreach (StiVariable variable in stimulReport.Dictionary.Variables)
            {
                if (variable.Category != FormulaFieldsCategory && variable.Category != RunningTotalFieldsCategory &&
                    variable.Category != CrystalVariablesCategory)
                    continue;

                if (!Regex.IsMatch(expression, $@"(?<![\w.]){Regex.Escape(variable.Name)}(?!\w)"))
                    continue;

                if (variable.Category != FormulaFieldsCategory || variable.InitBy != StiVariableInitBy.Expression)
                    return false;

                if (passedVariables.Add(variable) && !IsRecordExpression(variable.Value ?? string.Empty,
                    passedVariables))
                    return false;
            }

            return true;
        }

        private void RemoveChart(StiChart chart, StiVirtualSource source, StiDataSource masterSource)
        {
            chart.Parent?.Components.Remove(chart);
            stimulReport.Dictionary.DataSources.Remove(source);

            var columns = source.GroupColumns.Concat(source.Results.Where((value, index) => index % 3 == 0));
            foreach (var column in columns.Distinct().ToList())
            {
                if (masterSource.Columns[column] is StiCalcDataColumn calcColumn)
                    masterSource.Columns.Remove(calcColumn);
            }

            warningList.Add(StiWarning.GetWarningObject(chart.Name, "Chart on a formula without a value for a record"));
        }

        /// <summary>
        /// The font and the color of the title as Crystal prints them, the style of the chart would change them.
        /// </summary>
        private static void ApplyChartTitleFont(IStiChartTitle title, Ras.FontColor fontColor)
        {
            var font = fontColor?.Font;
            if (font == null || string.IsNullOrEmpty(font.Name) || font.Size <= 0)
                return;

            var style = FontStyle.Regular;
            if (font.Bold)
                style |= FontStyle.Bold;

            if (font.Italic)
                style |= FontStyle.Italic;

            if (font.Underline)
                style |= FontStyle.Underline;

            title.AllowApplyStyle = false;
            title.Font = new Font(font.Name, (float)font.Size, style);
            title.Brush = new StiSolidBrush(ColorTranslator.FromWin32((int)fontColor.Color));
        }

        private StiComponent SkipChart(ReportObject obj, string kind)
        {
            warningList.Add(StiWarning.GetWarningObject(obj.Name, kind));
            return null;
        }

        /// <summary>
        /// The level of the group whose section holds the chart, -1 for a chart outside the groups.
        /// </summary>
        private static int GetChartGroupLevel(Ras.ISCRChartObject chart)
        {
            switch (chart.ChartReportArea)
            {
                case Ras.CrAreaSectionKindEnum.crAreaSectionKindGroupHeader:
                case Ras.CrAreaSectionKindEnum.crAreaSectionKindGroupFooter:
                    return chart.ChartGroupIndex;

                default:
                    return -1;
            }
        }

        /// <summary>
        /// The period of a chart on a date: the summaries name it, like Sum ({Orders.Amount}, {Orders.Date}, "daily").
        /// </summary>
        private static RasData.CrDateConditionEnum? GetChartPeriod(Ras.ISCRChartDefinition definition)
        {
            foreach (RasData.ISCRField field in definition.DataFields)
            {
                var period = GetGroupNameCondition(field.FormulaForm);
                if (period != null)
                    return period;
            }

            return null;
        }

        /// <summary>
        /// The value of a chart: the summarized field and the function of the virtual source. A chart for each record
        /// shows the field of every record as it is.
        /// </summary>
        private StiCrystalChartValue GetChartValue(RasData.ISCRField field, bool isRecordChart)
        {
            var summary = field as RasData.ISCRSummaryField;
            var summarized = summary?.SummarizedField ?? field;

            var expression = GetChartFieldExpression(summarized.FormulaForm);
            if (expression == null)
                return null;

            var type = GetChartFieldType(summarized.Type);
            var title = summarized.ShortName ?? summarized.Name;

            if (isRecordChart || summary == null)
                return IsNumericType(type)
                    ? new StiCrystalChartValue(expression, type, string.Empty, title, type)
                    : null;

            var function = GetChartFunction(summary.Operation);
            if (function == null)
                return null;

            if (function == "Count" || function == "CountDistinct")
                return new StiCrystalChartValue(expression, type, function, title, typeof(int));

            return IsNumericType(type)
                ? new StiCrystalChartValue(expression, type, function, title, typeof(decimal))
                : null;
        }

        private static string GetChartFunction(RasData.CrSummaryOperationEnum operation)
        {
            switch (operation)
            {
                case RasData.CrSummaryOperationEnum.crSummaryOperationSum:
                    return "Sum";

                case RasData.CrSummaryOperationEnum.crSummaryOperationAverage:
                    return "Avg";

                case RasData.CrSummaryOperationEnum.crSummaryOperationCount:
                    return "Count";

                case RasData.CrSummaryOperationEnum.crSummaryOperationDistinctCount:
                    return "CountDistinct";

                case RasData.CrSummaryOperationEnum.crSummaryOperationMinimum:
                    return "Min";

                case RasData.CrSummaryOperationEnum.crSummaryOperationMaximum:
                    return "Max";

                case RasData.CrSummaryOperationEnum.crSummaryOperationMedian:
                    return "Median";

                default:
                    return null;
            }
        }

        /// <summary>
        /// The expression of a field of a chart: a column with the relations from the main source or the variable of
        /// a formula, a parameter or a running total.
        /// </summary>
        private string GetChartFieldExpression(string formulaForm)
        {
            var name = formulaForm?.Trim();
            if (name == null || name.Length < 3 || name[0] != '{' || name[name.Length - 1] != '}')
                return null;

            name = name.Substring(1, name.Length - 2);
            return (string)htFieldNameConversion[name] ?? (string)htVariableConversion[name];
        }

        /// <summary>
        /// A column of the main source which the virtual source takes as it is, a group or a value of a record.
        /// </summary>
        private string AddChartColumn(StiVirtualSource source, string name, string expression, Type type)
        {
            var column = GetChartMasterColumn(name, expression, type);
            if (source.Columns[column] == null)
                source.Columns.Add(new StiDataColumn(column, mainDataSource.Columns[column].Type));

            return column;
        }

        /// <summary>
        /// The column of the main source for the virtual source: a column of the main source itself, otherwise a new
        /// calculated column. The virtual source reads its columns from the main source by name only.
        /// </summary>
        private string GetChartMasterColumn(string name, string expression, Type type)
        {
            var prefix = mainDataSource.Name + ".";
            if (expression.StartsWith(prefix, StringComparison.Ordinal))
            {
                var columnName = expression.Substring(prefix.Length);
                var column = mainDataSource.Columns[columnName];
                if (column != null && !(column is StiCalcDataColumn))
                    return columnName;
            }

            // A field of a related table: the column gives the type, a time of Crystal may be a date in the source
            var dataColumn = StiDataColumn.GetDataColumnFromColumnName(stimulReport.Dictionary, expression,
                allowRelationName: true);
            if (dataColumn != null)
                type = dataColumn.Type;

            var calcName = CreateColumnName(mainDataSource, name);
            mainDataSource.Columns.Add(new StiCalcDataColumn(calcName, calcName, type, expression));
            return calcName;
        }

        private static void AddChartResult(List<string> results, string column, string function)
        {
            AddChartResult(results, column, function, column);
        }

        private static void AddChartResult(List<string> results, string column, string function, string name)
        {
            results.Add(column);
            results.Add(function);
            results.Add(name);
        }

        private static StiSeries CreateChartSeries(Ras.CrChartStyleTypeEnum type)
        {
            switch (type)
            {
                case Ras.CrChartStyleTypeEnum.crChartStyleTypeBar:
                case Ras.CrChartStyleTypeEnum.crChartStyleType3DRiser:
                    return new StiClusteredColumnSeries();

                case Ras.CrChartStyleTypeEnum.crChartStyleTypeLine:
                    return new StiLineSeries();

                case Ras.CrChartStyleTypeEnum.crChartStyleTypeArea:
                    return new StiAreaSeries();

                case Ras.CrChartStyleTypeEnum.crChartStyleTypePie:
                    return new StiPieSeries();

                case Ras.CrChartStyleTypeEnum.crChartStyleTypeDoughnut:
                    return new StiDoughnutSeries();

                // The numeric axis of Crystal is a line over dates or numbers (G2031, G2737)
                case Ras.CrChartStyleTypeEnum.crChartStyleTypeNumeric:
                    return new StiLineSeries();

                case Ras.CrChartStyleTypeEnum.crChartStyleTypeRadar:
                    return new StiRadarLineSeries();

                case Ras.CrChartStyleTypeEnum.crChartStyleTypeFunnel:
                    return new StiFunnelSeries();

                case Ras.CrChartStyleTypeEnum.crChartStyleTypeXYScatter:
                    return new StiScatterSeries();

                case Ras.CrChartStyleTypeEnum.crChartStyleTypeBubble:
                    return new StiBubbleSeries();

                default:
                    return null;
            }
        }

        /// <summary>
        /// A pie and a funnel show one value of each argument.
        /// </summary>
        private static bool IsPieSeries(Ras.CrChartStyleTypeEnum type)
        {
            return type == Ras.CrChartStyleTypeEnum.crChartStyleTypePie ||
                type == Ras.CrChartStyleTypeEnum.crChartStyleTypeDoughnut ||
                type == Ras.CrChartStyleTypeEnum.crChartStyleTypeFunnel;
        }

        /// <summary>
        /// The number of the values of a point: X and Y of a scatter, X, Y and the size of a bubble; 0 when every
        /// value is a series of its own over the argument.
        /// </summary>
        private static int GetPointValueCount(Ras.CrChartStyleTypeEnum type)
        {
            return type == Ras.CrChartStyleTypeEnum.crChartStyleTypeXYScatter ? 2
                : type == Ras.CrChartStyleTypeEnum.crChartStyleTypeBubble ? 3
                : 0;
        }

        private static Type GetChartFieldType(RasData.CrFieldValueTypeEnum type)
        {
            switch (type)
            {
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeBooleanField:
                    return typeof(bool);

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeCurrencyField:
                    return typeof(decimal);

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeDateField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeDateTimeField:
                    return typeof(DateTime);

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt8sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt8uField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt16sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt16uField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt32sField:
                case RasData.CrFieldValueTypeEnum.crFieldValueTypeInt32uField:
                    return typeof(int);

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeNumberField:
                    return typeof(double);

                case RasData.CrFieldValueTypeEnum.crFieldValueTypeTimeField:
                    return typeof(TimeSpan);

                default:
                    return typeof(string);
            }
        }

        private static bool IsNumericType(Type type)
        {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte) ||
                type == typeof(double) || type == typeof(float) || type == typeof(decimal);
        }

        private static string UnwrapExpression(string expression)
        {
            if (expression != null && expression.Length >= 2 && expression.StartsWith("{") && expression.EndsWith("}"))
                return expression.Substring(1, expression.Length - 2);

            return expression;
        }
        #endregion
    }
}
