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

using Stimulsoft.Base;
using System.Collections.Generic;
using System.Threading;
using System;
using System.Collections;
using Stimulsoft.Report.Components;
using System.Drawing;
using System.Text;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report.Components.ShapeTypes;
using Stimulsoft.Report.BarCodes;
using Stimulsoft.Report.CrossTab;
using Stimulsoft.Report.Chart;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Gauge;
using Stimulsoft.Report.Gauge.Helpers;
using Stimulsoft.Report.Components.TextFormats;

namespace Stimulsoft.Report.Import
{
    public class StiListAndLabelHelper
    {
        #region Fields
        private StiPage page = null;
        private Hashtable keyPairs = new Hashtable();
        private List<Hashtable> objects = new List<Hashtable>();
        private List<Hashtable> layers = new List<Hashtable>();
        private Hashtable objectsKeyPairs = new Hashtable();
        private Hashtable layersKeyPairs = new Hashtable();
        private List<Hashtable> pageLayoutsCollection = new List<Hashtable>();
        private List<Color> foregroundScheme = new List<Color>();
        private List<Color> backgroundScheme = new List<Color>();
        private Font defaultFont = new Font("Arial", 12);
        private Hashtable crossTabs = new Hashtable();
        private int crossTabIndex = 0;
        private int panelCrossTabIndex = 1;
        private int braceCounter = 0;
        private bool isNestedParameters = false;
        private bool isRow = false;
        private bool isColumn = false;
        private bool isData = false;
        private List<StiText> crossTabRows = new List<StiText>();
        private List<StiText> crossTabColumns = new List<StiText>();
        private List<StiText> crossTabDatas = new List<StiText>();
        private string outputFormatter = "";
        private string outputFormatterValue = "";
        private bool isCorrectOutputFormatter = false;
        private bool isProperties = false;
        private string currentRoot = "";
        private string previousRoot = "";
        private string parsingWord = "";
        private bool isTargetObject = false;
        private List<string> targetObject = new List<string>();
        private bool isTable = false;
        private List<List<string>> nestedPartsInStrings = new List<List<string>>();
        private List<List<string>> panelNestedObjects = new List<List<string>>();
        private Hashtable panelCrossTabs = new Hashtable();
        private List<List<string>> containerNestedPartsInStrings = new List<List<string>>();
        private int containerIndex = 0;
        private Hashtable panelTables = new Hashtable();
        private int tableIndex = 1;
        private int dataBandIndex = 1;
        private int headerBandIndex = 1;
        private int footerBandIndex = 1;
        private int groupHeaderBandIndex = 1;
        private int groupFooterBandIndex = 1;
        private int chartIndex = 1;
        private int gaugeIndex = 1;
        #endregion

        #region Root node
        public void ProcessFile(string[] iniFile, StiReport report)
        {
            IniParser(iniFile, report);
            crossTabIndex = 0;
            report.ReportUnit = StiReportUnitType.Millimeters;
            page = report.Pages[0];
            page.UnlimitedBreakable = false;
            page.TitleBeforeHeader = true;

            SetPageSettings(page);

            for (int objIndex = 0; objIndex < objects.Count; objIndex++)
            {
                var type = GetSetting(objects[objIndex], "Object", "ObjectName");

                if (type == "Text")
                    ProcessText(page, objects[objIndex]);

                if (type == "Line")
                    ProcessLine(page, objects[objIndex]);

                if (type == "Rectangle")
                    ProcessRectangle(page, objects[objIndex]);

                if (type == "Ellipse")
                    ProcessEllipse(page, objects[objIndex]);

                if (type == "Picture")
                    ProcessPicture(page, objects[objIndex]);

                if (type == "Barcode")
                    ProcessBarcode(page, objects[objIndex]);

                if (type == "Formatted Text")
                    ProcessRichText(page, objects[objIndex]);

                if (type == "Checkbox")
                    ProcessCheckBox(page, objects[objIndex]);

                if (type == "Chart")
                    ProcessChart(page, objects[objIndex]);

                if (type == "Gantt Chart")
                    ProcessGanttChart(page, objects[objIndex]);

                if (type == "Crosstab (Pivot Table)")
                    ProcessCrossTab(page, objects[objIndex]);

                if (type == "Report Container")
                {
                    ProcessPanel(page, objects[objIndex], containerIndex);
                    containerIndex++;
                }

                if (type == "Gauge")
                    ProcessGauge(page, objects[objIndex]);
            }
        }
        #endregion

        #region Items
        private void ProcessText(StiPage page, Hashtable obj)
        {
            var text = new StiText();
            text.WordWrap = true;
            ClientRectangle(obj, text);
            Font(obj, text);

            var textString = GetSetting(obj, "Object", "Text");
            var textFormat = GetSetting(obj, "Object", "OutputFormatter");

            TextFormat(text, textFormat);

            if (!string.IsNullOrEmpty(textString))
                text.Text = textString;

            var horAlignment = GetSetting(obj, "Object", "Align");
            text.HorAlignment = TextHorAlignment(horAlignment);

            var backColor = ParseStyleColor(GetSetting(obj, "Object", "BkColor"));

            if (backColor != null)
                text.Brush = new StiSolidBrush(backColor);

            page.Components.Add(text);
        }

        private void ProcessLine(StiPage page, Hashtable obj)
        {
            var lineDirection = GetSetting(obj, "Object", "LineDirection");
            var line = ShapeLineComponent(lineDirection);
            ClientRectangle(obj, line);

            var lineSize = ToHi(GetSetting(obj, "Object", "Width"));

            if (lineSize > 0)
                line.Size = lineSize;

            else
                line.Size = 1;

            var lineStyle = GetSetting(obj, "Object", "LineType");
            line.Style = ParseBorderStyle(lineStyle);

            var foreColor = ParseStyleColor(GetSetting(obj, "Object", "FgColor"));
            line.BorderColor = foreColor;

            page.Components.Add(line);
        }

        private void ProcessRectangle(StiPage page, Hashtable obj)
        {
            var rect = new StiShape { ShapeType = new StiRectangleShapeType() };
            ClientRectangle(obj, rect);

            var lineSize = ToHi(GetSetting(obj, "Object", "Width"));

            if (lineSize > 0)
                rect.Size = lineSize;

            else
                rect.Size = 1;

            var lineStyle = GetSetting(obj, "Object", "LineType");
            rect.Style = ParseBorderStyle(lineStyle);

            var backColor = ParseStyleColor(GetSetting(obj, "Object", "BkColor"));
            rect.Brush = new StiSolidBrush(backColor);

            var foreColor = ParseStyleColor(GetSetting(obj, "Object", "FgColor"));
            rect.BorderColor = foreColor;

            page.Components.Add(rect);
        }

        private void ProcessEllipse(StiPage page, Hashtable obj)
        {
            var rect = new StiShape { ShapeType = new StiOvalShapeType() };
            ClientRectangle(obj, rect);

            var lineSize = ToHi(GetSetting(obj, "Object", "Width"));

            if (lineSize > 0)
                rect.Size = lineSize;

            else
                rect.Size = 1;

            var lineStyle = GetSetting(obj, "Object", "LineType");
            rect.Style = ParseBorderStyle(lineStyle);

            var backColor = ParseStyleColor(GetSetting(obj, "Object", "BkColor"));
            rect.Brush = new StiSolidBrush(backColor);

            var foreColor = ParseStyleColor(GetSetting(obj, "Object", "FgColor"));
            rect.BorderColor = foreColor;

            page.Components.Add(rect);
        }

        private void ProcessPicture(StiPage page, Hashtable obj)
        {
            var image = new StiImage();
            ClientRectangle(obj, image);

            var content = GetSetting(obj, "Object", "Contents");

            /*if (content != null)
            {
                if (content.StartsWith("@2@"))
                {
                    var correctString = content.Remove(0, 3);
                    content = correctString;
                }

                var bytes = Convert.FromBase64String(content);
                image.ImageBytes = bytes;
            }
            else
            {
                var imagePath = GetSetting(obj, "Object", "Filename");

                if (!string.IsNullOrEmpty(imagePath))
                {
                    try
                    {
                        var newImage = Image.FromFile(imagePath);
                        image.Image = newImage;
                    }
                    catch
                    {
                        //maybe imagePath = project path + fileName;
                    }
                }
            }*/

            page.Components.Add(image);
        }

        private void ProcessBarcode(StiPage page, Hashtable obj)
        {
            var barcode = new StiBarCode();
            ClientRectangle(obj, barcode);

            var type = GetSetting(obj, "Object", "Method");
            BarcodeType(barcode, type);

            var code = GetSetting(obj, "Object", "FixedText");
            barcode.Code = new StiBarCodeExpression(code);

            var backColor = ParseStyleColor(GetSetting(obj, "Object", "BkColor"));
            barcode.BackColor = backColor;

            var foreColor = ParseStyleColor(GetSetting(obj, "Object", "FgColor"));
            barcode.ForeColor = foreColor;

            page.Components.Add(barcode);
        }

        private void ProcessRichText(StiPage page, Hashtable obj)
        {
            var text = new StiRichText();
            ClientRectangle(obj, text);

            var rtfString = GetSetting(obj, "Object", "Contents");

            if (!string.IsNullOrEmpty(rtfString))
                text.RtfText = rtfString;

            page.Components.Add(text);
        }

        private void ProcessCheckBox(StiPage page, Hashtable obj)
        {
            var checkBox = new StiCheckBox();
            ClientRectangle(obj, checkBox);

            var width = 5d;
            var height = 5d;
            var x = checkBox.ClientRectangle.X + (checkBox.ClientRectangle.Width - width) / 2;
            var y = checkBox.ClientRectangle.Y + (checkBox.ClientRectangle.Height - height) / 2;

            checkBox.ClientRectangle = new RectangleD(x, y, width, height);
            checkBox.CheckStyleForTrue = StiCheckStyle.CrossRectangle;
            checkBox.Enabled = true;

            var isChecked = GetSetting(obj, "Object", "Contents");

            if (isChecked != null && isChecked == "True")
            {
                checkBox.Checked = new StiCheckedExpression();
                checkBox.Checked.Value = "true";
            }

            page.Components.Add(checkBox);
        }

        private void ProcessCrossTab(StiPage page, Hashtable obj)
        {
            var crossTable = new StiCrossTab();
            crossTable.Name = $"CrossTab{crossTabIndex}";

            var mainDataSource = GetSetting(obj, "obj", "TableID");

            if (string.IsNullOrEmpty(mainDataSource))
                mainDataSource = GetSetting(obj, "Object", "ID");

            if (!string.IsNullOrEmpty(mainDataSource))
                crossTable.DataSourceName = mainDataSource;

            ClientRectangle(obj, crossTable);

            var crossTab = (List<List<StiText>>)crossTabs[crossTabIndex];

            var columnFields = new List<StiText>();
            var rowFields = new List<StiText>();
            var dataFields = new List<StiText>();

            if (crossTab != null && crossTab.Count > 0)
                columnFields = crossTab[0];

            if (crossTab != null && crossTab.Count > 1)
                rowFields = crossTab[1];

            if (crossTab != null && crossTab.Count > 2)
                dataFields = crossTab[2];

            CrossTabBuilder(page, obj, crossTable, rowFields, columnFields, dataFields);

            page.Components.Add(crossTable);

            crossTabIndex++;
        }

        private void ProcessPanel(StiPage page, Hashtable obj, int index)
        {
            panelCrossTabs = new Hashtable();
            panelNestedObjects = new List<List<string>>();

            var childY = 0d;
            var panelContainer = new StiPanel();
            panelContainer.Name = $"Container{index + 1}";

            ClientRectangle(obj, panelContainer);

            panelContainer.CanGrow = false;
            panelContainer.CanShrink = false;
            panelContainer.Page = page;

            foreach (var str in containerNestedPartsInStrings[index])
            {
                NestedObjectParse(str);
            }

            for (int objectIndex = 0; objectIndex < panelNestedObjects.Count; objectIndex++)
            {
                if (panelNestedObjects[objectIndex].Count > 3 && panelNestedObjects[objectIndex][2] == "TableType=Crosstab")
                {
                    var mainDataSource = "";
                    var dataSourceArray = panelNestedObjects[objectIndex][3].Split('=');

                    if (dataSourceArray.Length == 2)
                        mainDataSource = dataSourceArray[1];

                    foreach (var str in panelNestedObjects[objectIndex])
                    {
                        if (!string.IsNullOrEmpty(str))
                            CrossTabLineParse(str, objectIndex);
                    }

                    foreach (var crossObj in panelCrossTabs.Values)
                    {
                        var crossTab = crossObj as List<List<StiText>>;
                        var crossTable = new StiCrossTab();

                        if (!string.IsNullOrEmpty(mainDataSource))
                            crossTable.DataSourceName = mainDataSource;

                        crossTable.Name = $"{panelContainer.Name}_CrossTab{panelCrossTabIndex}";
                        panelCrossTabIndex++;

                        var columnFields = new List<StiText>();
                        var rowFields = new List<StiText>();
                        var dataFields = new List<StiText>();

                        if (crossTab != null && crossTab.Count > 0)
                            columnFields = crossTab[0];

                        if (crossTab != null && crossTab.Count > 1)
                            rowFields = crossTab[1];

                        if (crossTab != null && crossTab.Count > 2)
                            dataFields = crossTab[2];

                        CrossTabBuilder(page, obj, crossTable, rowFields, columnFields, dataFields, true);

                        if (crossTable.Width > panelContainer.Width)
                            crossTable.Width = panelContainer.Width;

                        if (crossTable.ClientRectangle.Y + childY + crossTable.Height >
                            panelContainer.Height)
                            crossTable.Height = panelContainer.Height -
                                (crossTable.ClientRectangle.Y + childY);

                        crossTable.ClientRectangle = new RectangleD(
                            crossTable.ClientRectangle.X,
                            crossTable.ClientRectangle.Y + childY,
                            crossTable.ClientRectangle.Width,
                            crossTable.ClientRectangle.Height);

                        childY += crossTable.Height;

                        panelContainer.Components.Add(crossTable);
                    }

                    panelCrossTabs.Clear();
                }
                else if (panelNestedObjects[objectIndex].Count > 2 && panelNestedObjects[objectIndex][2] == "TableType=List")
                {
                    TableLineParse(panelNestedObjects[objectIndex], objectIndex);

                    foreach (var tableObj in panelTables.Values)
                    {
                        var table = tableObj as StiPanel;
                        var tablePanel = TableBuilder(page, table, panelContainer, 0, childY);
                        panelContainer.Components.Add(tablePanel);

                        childY += tablePanel.ClientRectangle.Height;
                    }

                    panelTables.Clear();
                }
                else if (panelNestedObjects[objectIndex].Count > 2 && panelNestedObjects[objectIndex][2] == "TableType=Chart")
                {
                    var chart = new StiChart();
                    chart.Name = $"Chart{chartIndex}";
                    chartIndex++;

                    chart.CanGrow = false;
                    chart.CanShrink = false;
                    chart.Page = page;
                    chart.Width = panelContainer.ClientRectangle.Width;

                    ChartLineParse(chart, panelNestedObjects[objectIndex], childY,
                        panelContainer.ClientRectangle.Width);

                    childY += chart.ClientRectangle.Height;

                    panelContainer.Components.Add(chart);
                }
            }

            page.Components.Add(panelContainer);
        }

        private void ProcessChart(StiPage page, Hashtable obj)
        {
            var chart = new StiChart();
            chart.Name = $"Chart{chartIndex}";
            chartIndex++;

            ClientRectangle(obj, chart);

            var type = GetSetting(obj, "Object", "Type");
            var dataMode = GetSetting(obj, "Object", "DataMode");
            var visMode = GetSetting(obj, "Object", "VisMode");
            ChartType(chart, type, visMode, dataMode);

            chart.CanGrow = true;
            chart.CanShrink = false;
            chart.Page = page;

            page.Components.Add(chart);
        }

        private void ProcessGanttChart(StiPage page, Hashtable obj)
        {
            var chart = new StiChart();
            chart.Name = $"Chart{chartIndex}";
            chartIndex++;

            ClientRectangle(obj, chart);

            var chartArea = new StiGanttArea();
            var chartSeries = new StiGanttSeries();
            chart.Area = chartArea;
            chart.Series.Add(chartSeries);

            chart.CanGrow = true;
            chart.CanShrink = false;
            chart.Page = page;

            page.Components.Add(chart);
        }

        private void ProcessGauge(StiPage page, Hashtable obj)
        {
            var gauge = new StiGauge();
            gauge.Name = $"Gauge{gaugeIndex}";
            gaugeIndex++;

            gauge.CanGrow = true;
            gauge.CanShrink = false;
            gauge.Page = page;

            //var frame = GetSetting(obj, "Object", "Frame");
            var glass = GetSetting(obj, "Object", "Glass");
            //var pointer = GetSetting(obj, "Object", "Pointer");

            if (!string.IsNullOrEmpty(glass))
            {
                var glassAray = glass.Split('.');

                if (glassAray.Length == 5)
                {
                    var gaugeType = glassAray[3];

                    if (gaugeType == "ROUND")
                        gauge.Type = StiGaugeType.FullCircular;

                    else if (gaugeType == "HALFROUND")
                        gauge.Type = StiGaugeType.HalfCircular;
                }
                else if (glassAray.Length == 4)
                {
                    var gaugeType = glassAray[3];

                    if (gaugeType.EndsWith("H"))
                        gauge.Type = StiGaugeType.HorizontalLinear;

                    else if (gaugeType.EndsWith("V"))
                        gauge.Type = StiGaugeType.Linear;
                }
            }

            StiGaugeV2InitHelper.Init(gauge, gauge.Type, true, false);

            gauge.Style = new StiGaugeStyleXF26();
            gauge.Minimum = 0;
            gauge.Maximum = 100;

            ClientRectangle(obj, gauge);

            page.Components.Add(gauge);
        }
        #endregion

        #region Methods Ini
        private struct SectionPair
        {
            internal string Section;
            internal string Key;
        }

        private void IniParser(string[] iniFile, StiReport report)
        {
            var objectInStrings = new List<string>();
            var objectsInStrings = new List<List<string>>();
            var nestedPartInStrings = new List<string>();
            var isNestedPartInside = false;

            #region Brake Init
            try
            {
                foreach (var line in iniFile)
                {
                    var str = line.Trim();

                    if (str != "")
                    {
                        if (str.StartsWith("[") && str.EndsWith("]") && !isNestedParameters)
                        {
                            var newRoot = str.Substring(1, str.Length - 2);

                            if (newRoot.StartsWith("@"))
                            {
                                nestedPartInStrings.Add(str);

                                braceCounter = 0;
                                isNestedParameters = true;
                                isNestedPartInside = true;
                            }
                            else
                            {
                                var newObj = new List<string>();

                                if (previousRoot != null)
                                    newObj.Add(previousRoot);

                                if (objectInStrings.Count > 0)
                                {
                                    foreach (var strr in objectInStrings)
                                        newObj.Add(strr);
                                }

                                if (newObj.Count > 0)
                                    objectsInStrings.Add(newObj);

                                objectInStrings.Clear();

                                previousRoot = str;

                                if (!isNestedPartInside)
                                    nestedPartsInStrings.Add(new List<string>());

                                else
                                    isNestedPartInside = false;
                            }
                        }
                        else if (str.StartsWith("[") && str.EndsWith("]") && isNestedParameters)
                        {
                            nestedPartInStrings.Add(str);
                        }
                        else
                        {
                            var keyPair = str.Split(new char[] { '=' }, 2);

                            SectionPair sectionPair;
                            string value = null;

                            sectionPair.Section = currentRoot;
                            sectionPair.Key = keyPair[0];

                            if (keyPair.Length > 1)
                                value = keyPair[1];

                            if (isNestedParameters && (sectionPair.Key == "{" || sectionPair.Key == "}"))
                            {
                                if (sectionPair.Key == "{")
                                    braceCounter++;

                                if (sectionPair.Key == "}")
                                    braceCounter--;

                                nestedPartInStrings.Add(str);
                            }
                            else if (isNestedParameters && sectionPair.Key != "{" && sectionPair.Key != "}")
                            {
                                nestedPartInStrings.Add(str);
                            }
                            else
                            {
                                objectInStrings.Add(str);
                            }

                            if (isNestedParameters && braceCounter == 0)
                            {
                                isNestedParameters = false;

                                var newPartsCollection = new List<string>();

                                foreach (var partStr in nestedPartInStrings)
                                    newPartsCollection.Add(partStr);

                                nestedPartsInStrings.Add(newPartsCollection);

                                if (nestedPartsInStrings.Count > objectsInStrings.Count)
                                {
                                    if (objectInStrings.Count > 0 || previousRoot != null)
                                    {
                                        var newObj = new List<string>();

                                        if (previousRoot != null)
                                            newObj.Add(previousRoot);

                                        foreach (var strr in objectInStrings)
                                            newObj.Add(strr);

                                        if (newObj.Count > 0)
                                            objectsInStrings.Add(newObj);

                                        objectInStrings.Clear();
                                        previousRoot = null;
                                    }

                                    var differense = nestedPartsInStrings.Count - objectsInStrings.Count;

                                    while (differense > 0)
                                    {
                                        var emptyObj = new List<string>();
                                        objectsInStrings.Add(emptyObj);

                                        differense--;
                                    }
                                }

                                nestedPartInStrings.Clear();
                            }
                        }
                    }
                }

                var obj = new List<string>();

                if (previousRoot != null)
                    obj.Add(previousRoot);

                if (objectInStrings.Count > 0)
                {
                    foreach (var strr in objectInStrings)
                        obj.Add(strr);
                }

                if (obj.Count > 0)
                    objectsInStrings.Add(obj);

                objectInStrings.Clear();

                if (!isNestedPartInside)
                    nestedPartsInStrings.Add(new List<string>());

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
            }

            var tempContainerNestedPartsInStrings = new List<List<string>>();

            for (int index = 0; index < objectsInStrings.Count; index++)
            {
                if (objectsInStrings[index].Count > 2 && objectsInStrings[index][2] == "ObjectName=Report Container")
                {
                    tempContainerNestedPartsInStrings.Add(nestedPartsInStrings[index]);
                }
                else
                {
                    foreach (var nestedStr in nestedPartsInStrings[index])
                    {
                        objectsInStrings[index].Add(nestedStr);
                    }
                }
            }

            if (tempContainerNestedPartsInStrings.Count > 0)
            {
                foreach (var nestedObjectStr in tempContainerNestedPartsInStrings)
                    containerNestedPartsInStrings.Add(nestedObjectStr);

                tempContainerNestedPartsInStrings.Clear();
            }

            #region Color Scheme
            var isForeground = false;
            var isBackground = false;

            foreach (var obj in objectsInStrings)
            {
                if (obj.Count > 0 && obj[0] == "[ExtendedInfo]")
                {
                    foreach (var line in obj)
                    {
                        if (isForeground && line == "}")
                            isForeground = false;

                        if (isBackground && line == "}")
                            isBackground = false;

                        if (isForeground)
                        {
                            var colorInString = line.Replace("(", "").Replace(")", "").Split('=');

                            if (colorInString.Length == 2)
                            {
                                var color = ParseStyleColor(colorInString[1]);

                                if (color != null)
                                    foregroundScheme.Add(color);
                            }
                        }

                        if (isBackground)
                        {
                            var colorInString = line.Replace("(", "").Replace(")", "").Split('=');

                            if (colorInString.Length == 2)
                            {
                                var color = ParseStyleColor(colorInString[1]);

                                if (color != null)
                                    backgroundScheme.Add(color);
                            }
                        }

                        if (line == "[Foreground]")
                            isForeground = true;

                        if (line == "[Background]")
                            isBackground = true;
                    }
                }
            }
            #endregion

            #region Page Layouts
            var isPageLayout = false;
            var pageLayout = new List<string>();
            var pageLayots = new List<List<string>>();

            foreach (var obj in objectsInStrings)
            {
                if (obj.Count > 0 && obj[0] == "[PageLayouts]")
                {
                    foreach (var line in obj)
                    {
                        if (isPageLayout && line == "}")
                        {
                            isPageLayout = false;
                            pageLayots.Add(pageLayout);
                        }

                        if (isPageLayout && line != "{" && line != "}")
                            pageLayout.Add(line);

                        if (line == "[PageLayout]")
                        {
                            isPageLayout = true;
                            pageLayout = new List<string>();
                        }
                    }
                }
            }

            foreach (var layout in pageLayots)
            {
                var layoutHash = new Hashtable();

                foreach (var line in layout)
                {
                    var keyPair = line.Split(new char[] { '=' }, 2);

                    if (keyPair.Length > 1)
                        layoutHash.Add(keyPair[0], keyPair[1]);
                }

                pageLayoutsCollection.Add(layoutHash);
            }
            #endregion

            #region DataSources
            var isDataSource = false;
            var dataSource = new List<string>();
            var dataSources = new List<List<string>>();
            var dataSourseBraces = 0;

            foreach (var obj in objectsInStrings)
            {
                if (obj.Count > 0 && obj[0] == "[Description]")
                {
                    foreach (var line in obj)
                    {
                        if (isDataSource)
                            dataSource.Add(line);

                        if (line == "{")
                            dataSourseBraces++;

                        if (line == "}")
                            dataSourseBraces--;

                        if (isDataSource && dataSourseBraces == 0)
                        {
                            isDataSource = false;
                            dataSources.Add(dataSource);
                        }

                        if (line.Split(':')[0] == "[Table")
                        {
                            dataSourseBraces = 0;
                            isDataSource = true;
                            dataSource = new List<string>();
                            dataSource.Add(line);
                        }
                    }
                }
            }

            foreach (var source in dataSources)
            {
                var newDataSource = new StiDataTableSource();
                var isFields = false;

                if (source.Count > 0)
                {
                    var nameLines = source[0].Split(':');

                    if (nameLines.Length == 2)
                    {
                        var name = nameLines[1].Replace("[", "").Replace("]", "");
                        newDataSource.Name = name;
                        newDataSource.NameInSource = name;
                        newDataSource.Alias = name;
                    }
                }

                foreach (var line in source)
                {
                    if (isFields && line == "}")
                        isFields = false;

                    if (isFields && line != "{")
                    {
                        var newLine = line.Split('.');

                        if (newLine.Length == 2)
                        {
                            var columnName = newLine[1].Split('=')[0];

                            var newColumn = new StiDataColumn(columnName);
                            newColumn.Name = columnName;

                            Type newType = Type.GetType("System.Object");
                            newColumn.Type = newType;

                            newDataSource.Columns.Add(newColumn);
                        }
                    }

                    if (line == "[UsedIdentifiers]")
                        isFields = true;
                }

                report.Dictionary.DataSources.Add(newDataSource);
            }
            #endregion
            #endregion

            isNestedParameters = false;
            braceCounter = 0;
            currentRoot = null;
            previousRoot = null;

            try
            {
                for (int objectIndex = 0; objectIndex < objectsInStrings.Count; objectIndex++)
                {
                    foreach (var str in objectsInStrings[objectIndex])
                    {
                        if (!string.IsNullOrEmpty(str))
                        {
                            IniLineParse(str);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }

        private void IniLineParse(string str)
        {
            var strLine = str.Trim();

            if (strLine.StartsWith("[") && strLine.EndsWith("]") && !isNestedParameters)
            {
                var newRoot = strLine.Substring(1, strLine.Length - 2);

                if (newRoot.StartsWith("@"))
                {
                    braceCounter = 0;
                    isNestedParameters = true;
                }
                else
                {
                    currentRoot = strLine.Substring(1, strLine.Length - 2);

                    if (previousRoot == "Object")
                    {
                        objects.Add(objectsKeyPairs);
                        objectsKeyPairs = new Hashtable();

                        if (crossTabColumns.Count > 0 && crossTabRows.Count > 0)
                        {
                            var columns = new List<StiText>();
                            var rows = new List<StiText>();
                            var data = new List<StiText>();

                            foreach (var column in crossTabColumns)
                                columns.Add(column);

                            foreach (var row in crossTabRows)
                                rows.Add(row);

                            foreach (var dataCell in crossTabDatas)
                                data.Add(dataCell);

                            var crossTab = new List<List<StiText>>();
                            crossTab.Add(columns);
                            crossTab.Add(rows);
                            crossTab.Add(data);

                            crossTabColumns.Clear();
                            crossTabRows.Clear();
                            crossTabDatas.Clear();

                            crossTabs.Add(crossTabIndex, crossTab);
                            crossTabIndex++;
                        }
                    }
                    else if (previousRoot == "Layer")
                    {
                        layers.Add(layersKeyPairs);
                        layersKeyPairs = new Hashtable();
                    }

                    previousRoot = currentRoot;
                }
            }
            else if (strLine.StartsWith("[") && strLine.EndsWith("]") && isNestedParameters)
            {
                var newRoot = strLine.Substring(1, strLine.Length - 2);

                if (newRoot == "Coord.Row")
                {
                    isRow = true;
                    isColumn = false;
                    isData = false;
                }

                if (newRoot == "Coord.Col")
                {
                    isRow = false;
                    isColumn = true;
                    isData = false;
                }

                if (newRoot == "Properties.Cell" && crossTabDatas.Count == 0)
                {
                    isRow = false;
                    isColumn = false;
                    isProperties = true;
                }

                if (isCorrectOutputFormatter && newRoot == outputFormatter)
                {
                    isData = true;
                }

                if (newRoot == "Properties.Label.Col")
                {
                    isRow = false;
                    isColumn = false;
                    isData = false;
                    isProperties = false;
                }
            }
            else
            {
                var keyPair = strLine.Split(new char[] { '=' }, 2);

                SectionPair sectionPair;
                string value = null;

                if (currentRoot == null)
                    currentRoot = "Root";

                sectionPair.Section = currentRoot;
                sectionPair.Key = keyPair[0];

                if (keyPair.Length > 1)
                    value = keyPair[1];

                if (isRow && sectionPair.Key == "Formula")
                {
                    isRow = false;

                    var newText = new StiText();
                    newText.Text = value;
                    TextFormat(newText, outputFormatterValue);
                    outputFormatterValue = "";

                    crossTabRows.Add(newText);
                }

                if (isColumn && sectionPair.Key == "Formula")
                {
                    isColumn = false;

                    var newText = new StiText();
                    newText.Text = value;
                    TextFormat(newText, outputFormatterValue);
                    outputFormatterValue = "";

                    crossTabColumns.Add(newText);
                }

                if (isProperties && sectionPair.Key == "OutputFormatter")
                {
                    outputFormatter = $"OutputFormatter.{value}";
                    outputFormatterValue = value;
                    isCorrectOutputFormatter = true;
                }

                if (isProperties && isData && sectionPair.Key == "Formula")
                {
                    isData = false;

                    var newText = new StiText();
                    newText.Text = value;
                    TextFormat(newText, outputFormatterValue);
                    outputFormatterValue = "";

                    crossTabDatas.Add(newText);
                }

                if (isNestedParameters && (sectionPair.Key == "{" || sectionPair.Key == "}"))
                {
                    if (sectionPair.Key == "{")
                        braceCounter++;

                    if (sectionPair.Key == "}")
                        braceCounter--;
                }
                else
                {
                    if (!keyPairs.Contains(sectionPair) && sectionPair.Key != "}" && currentRoot != "Object" && currentRoot != "Layer")
                    {
                        keyPairs.Add(sectionPair, value);
                    }
                    else if (currentRoot == "Object")
                    {
                        if (!objectsKeyPairs.Contains(sectionPair))
                            objectsKeyPairs.Add(sectionPair, value);
                    }
                    else if (currentRoot == "Layer")
                    {
                        if (!layersKeyPairs.Contains(sectionPair))
                            layersKeyPairs.Add(sectionPair, value);
                    }
                }

                if (isNestedParameters && braceCounter == 0)
                    isNestedParameters = false;
            }
        }

        private void CrossTabLineParse(string str, int crossTabIndex)
        {
            var strLine = str.Trim();

            if (strLine.StartsWith("[") && strLine.EndsWith("]"))
            {
                var newRoot = strLine.Substring(1, strLine.Length - 2);
                currentRoot = newRoot;

                if (newRoot == "Table" && !isTable)
                {
                    isTable = true;
                    braceCounter = 0;
                }

                if (newRoot == "Coord.Row")
                {
                    isRow = true;
                    isColumn = false;
                    isData = false;
                }

                if (newRoot == "Coord.Col")
                {
                    isRow = false;
                    isColumn = true;
                    isData = false;
                }

                if (newRoot == "Properties.Cell")
                {
                    isRow = false;
                    isColumn = false;
                    isProperties = true;
                }

                if (isCorrectOutputFormatter && newRoot == outputFormatter)
                {
                    isData = true;
                }

                if (newRoot == "Properties.Label.Col" || newRoot == "Properties.Label.Row")
                {
                    isRow = false;
                    isColumn = false;
                    isData = false;
                    isProperties = false;
                }
            }
            else
            {
                var keyPair = strLine.Split(new char[] { '=' }, 2);

                SectionPair sectionPair;
                string value = null;

                if (currentRoot == null)
                    currentRoot = "Root";

                sectionPair.Section = currentRoot;
                sectionPair.Key = keyPair[0];

                if (keyPair.Length > 1)
                    value = keyPair[1];

                if (isRow && sectionPair.Key == "Formula")
                {
                    isRow = false;

                    var newText = new StiText();
                    newText.Text = value;
                    TextFormat(newText, outputFormatterValue);
                    outputFormatterValue = "";

                    crossTabRows.Add(newText);
                }

                if (isColumn && sectionPair.Key == "Formula")
                {
                    isColumn = false;

                    var newText = new StiText();
                    newText.Text = value;
                    TextFormat(newText, outputFormatterValue);
                    outputFormatterValue = "";

                    crossTabColumns.Add(newText);
                }

                if (isProperties && sectionPair.Key == "OutputFormatter")
                {
                    outputFormatter = $"OutputFormatter.{value}";
                    outputFormatterValue = value;
                    isCorrectOutputFormatter = true;
                    isData = true;
                }

                if (isProperties && isData && sectionPair.Key == "Formula")
                {
                    isData = false;

                    var newText = new StiText();
                    newText.Text = value;
                    TextFormat(newText, outputFormatterValue);
                    outputFormatterValue = "";

                    crossTabDatas.Add(newText);
                }

                if (isTable && (sectionPair.Key == "{" || sectionPair.Key == "}"))
                {
                    if (sectionPair.Key == "{")
                        braceCounter++;

                    if (sectionPair.Key == "}")
                        braceCounter--;
                }
                else
                {
                    if (!keyPairs.Contains(sectionPair) && sectionPair.Key != "}")
                        keyPairs.Add(sectionPair, value);
                }

                if (isTable && braceCounter == 0)
                {
                    isTable = false;

                    if (crossTabColumns.Count > 0 && crossTabRows.Count > 0)
                    {
                        var columns = new List<StiText>();
                        var rows = new List<StiText>();
                        var data = new List<StiText>();

                        foreach (var column in crossTabColumns)
                            columns.Add(column);

                        foreach (var row in crossTabRows)
                            rows.Add(row);

                        foreach (var dataCell in crossTabDatas)
                            data.Add(dataCell);

                        var crossTab = new List<List<StiText>>();
                        crossTab.Add(columns);
                        crossTab.Add(rows);
                        crossTab.Add(data);

                        crossTabColumns.Clear();
                        crossTabRows.Clear();
                        crossTabDatas.Clear();

                        panelCrossTabs.Add(crossTabIndex, crossTab);
                    }

                    objects.Add(objectsKeyPairs);
                    objectsKeyPairs = new Hashtable();
                }
            }
        }

        private void TableLineParse(List<string> tableLineCollection, int tableIndex)
        {
            var row = new List<StiComponent>();
            var tableCellText = new StiText();
            var tableCellWidth = 0d;
            var tableCellHeight = 0d;
            var tableCellX = 0d;
            var tableCellY = 0d;
            var isTableRow = false;
            var isTableCell = false;
            var isTableCellProperties = false;
            var isCellBorder = false;
            var isLeftBorder = false;
            var hasLeftBorder = false;
            var isTopBorder = false;
            var hasTopBorder = false;
            var isRightBorder = false;
            var hasRightBorder = false;
            var isBottomBorder = false;
            var hasBottomBorder = false;
            var cellBorder = new StiBorder(StiBorderSides.None, Color.Black, 1, StiPenStyle.Solid);
            var tableBands = new List<StiBand>();
            var cellBraceCounter = 0;
            var borderBraceCounter = 0;
            var isCellFont = false;
            var fontBraceCounter = 0;
            var cellFont = new Hashtable();
            var isDataBand = false;
            var isHeaderBand = false;
            var isFooterBand = false;
            var isGroupHeaderBand = false;
            var isGroupFooterBand = false;
            var allTables = new List<StiPanel>();
            var childTableLinesCollection = ChildTables(tableLineCollection);
            var dataSource = "";

            foreach (var table in childTableLinesCollection)
            {
                tableBands = new List<StiBand>();

                foreach (var str in table)
                {
                    if (!string.IsNullOrEmpty(str))
                    {
                        var strLine = str.Trim();

                        if (strLine.StartsWith("[") && strLine.EndsWith("]"))
                        {
                            var newRoot = strLine.Substring(1, strLine.Length - 2);
                            currentRoot = newRoot;

                            if (newRoot == "Table" && !isTable)
                            {
                                isTable = true;
                                braceCounter = 0;
                            }

                            if ((newRoot == "Body"
                                || newRoot == "Header"
                                || newRoot == "Footer"
                                || newRoot == "GroupHeader"
                                || newRoot == "GroupFooter")
                                && isTableRow && row.Count > 0)
                            {
                                isTableRow = true;

                                TableBandsFill(isHeaderBand, isGroupHeaderBand, isDataBand, isGroupFooterBand, isFooterBand,
                                        row, tableBands, dataSource);

                                TableBandsSort(tableBands);

                                row.Clear();
                                tableCellX = 0;
                                tableCellY = 0;
                            }

                            if (newRoot == "Line" && !isTableCell && row.Count > 0)
                            {
                                tableCellX = 0;
                                tableCellY = row[0].Height;
                            }

                            if (isTable && newRoot == "Body")
                            {
                                isDataBand = true;
                                isHeaderBand = false;
                                isFooterBand = false;
                                isGroupHeaderBand = false;
                                isGroupFooterBand = false;
                            }

                            if (isTable && newRoot == "Header")
                            {
                                isHeaderBand = true;
                                isDataBand = false;
                                isFooterBand = false;
                                isGroupHeaderBand = false;
                                isGroupFooterBand = false;
                            }

                            if (isTable && newRoot == "Footer")
                            {
                                isFooterBand = true;
                                isDataBand = false;
                                isHeaderBand = false;
                                isGroupHeaderBand = false;
                                isGroupFooterBand = false;
                            }

                            if (isTable && newRoot == "GroupHeader")
                            {
                                isGroupHeaderBand = true;
                                isDataBand = false;
                                isHeaderBand = false;
                                isFooterBand = false;
                                isGroupFooterBand = false;
                            }

                            if (isTable && newRoot == "GroupFooter")
                            {
                                isGroupFooterBand = true;
                                isDataBand = false;
                                isHeaderBand = false;
                                isFooterBand = false;
                                isGroupHeaderBand = false;
                            }

                            if (newRoot == "Columns" && !isTableRow)
                                isTableRow = true;

                            if (isTableRow && newRoot == "Column")
                            {
                                cellBraceCounter = 0;
                                isTableCell = true;
                                isTableCellProperties = true;
                            }

                            if (isTableCell && newRoot == "Frame")
                            {
                                isTableCellProperties = false;
                                isCellBorder = true;
                                borderBraceCounter = 0;
                            }

                            if (isCellBorder && newRoot == "Left")
                            {
                                isLeftBorder = true;
                                isBottomBorder = false;
                            }

                            if (isCellBorder && newRoot == "Top")
                            {
                                isTopBorder = true;
                                isLeftBorder = false;
                            }

                            if (isCellBorder && newRoot == "Right")
                            {
                                isRightBorder = true;
                                isTopBorder = false;
                            }

                            if (isCellBorder && newRoot == "Bottom")
                            {
                                isBottomBorder = true;
                                isRightBorder = false;
                            }

                            if (isTableCell && newRoot == "Font")
                            {
                                cellFont = new Hashtable();
                                isCellFont = true;
                                fontBraceCounter = 0;
                            }
                        }
                        else
                        {
                            var keyPair = strLine.Split(new char[] { '=' }, 2);

                            SectionPair sectionPair;
                            string value = null;

                            if (currentRoot == null)
                                currentRoot = "Root";

                            sectionPair.Section = currentRoot;
                            sectionPair.Key = keyPair[0];

                            if (keyPair.Length > 1)
                                value = keyPair[1];

                            #region Properties
                            if (isTableCell && isTableCellProperties && sectionPair.Key == "Text")
                                tableCellText.Text = value;

                            if (isTableCell && isTableCellProperties && sectionPair.Key == "Width")
                                tableCellWidth = ToHi(value);

                            if (isTableCell && isTableCellProperties && sectionPair.Key == "Height")
                                tableCellHeight = ToHi(value);

                            if (isTableCell && isTableCellProperties && sectionPair.Key == "BkColor")
                            {
                                var background = value.Split('.');

                                if (background.Length > 0)
                                {
                                    if (background[1] == "Scheme")
                                    {
                                        var colorIndex = int.Parse(background[background.Length - 1].Remove(0, 15));//remove "BackgroundColor"

                                        if (colorIndex < backgroundScheme.Count)
                                            tableCellText.Brush = new StiSolidBrush(backgroundScheme[colorIndex]);
                                    }
                                    if (background[1] == "Color")
                                    {
                                        var color = ParseStyleColor(background[background.Length - 1]);
                                        tableCellText.Brush = new StiSolidBrush(color);
                                    }
                                }
                            }

                            if (isTableCell && isTableCellProperties && sectionPair.Key == "Align")
                                tableCellText.HorAlignment = TextHorAlignment(value);

                            if (isTableCell && isTableCellProperties && sectionPair.Key == "OutputFormatter")
                            {
                                TextFormat(tableCellText, value);
                            }

                            if (isTable && sectionPair.Key == "TableID")
                                dataSource = value;
                            #endregion

                            #region Border
                            if (isLeftBorder && sectionPair.Key == "Line")
                            {
                                if (value == "True")
                                    hasLeftBorder = true;
                            }

                            if (isTopBorder && sectionPair.Key == "Line")
                            {
                                if (value == "True")
                                    hasTopBorder = true;
                            }

                            if (isRightBorder && sectionPair.Key == "Line")
                            {
                                if (value == "True")
                                    hasRightBorder = true;
                            }

                            if (isBottomBorder && sectionPair.Key == "Line")
                            {
                                if (value == "True")
                                    hasBottomBorder = true;
                            }
                            #endregion

                            #region Font
                            if (isCellFont)
                            {
                                sectionPair.Section = "Object";
                                cellFont.Add(sectionPair, value);
                            }
                            #endregion

                            if (isTable && (sectionPair.Key == "{" || sectionPair.Key == "}"))
                            {
                                if (sectionPair.Key == "{")
                                {
                                    braceCounter++;

                                    if (isTableCell)
                                        cellBraceCounter++;

                                    if (isCellBorder)
                                        borderBraceCounter++;

                                    if (isCellFont)
                                        fontBraceCounter++;
                                }

                                if (sectionPair.Key == "}")
                                {
                                    braceCounter--;

                                    if (isTableCell)
                                        cellBraceCounter--;

                                    if (isCellBorder)
                                        borderBraceCounter--;

                                    if (isCellFont)
                                        fontBraceCounter--;
                                }
                            }

                            if (isCellFont && fontBraceCounter == 0)
                            {
                                Font(cellFont, tableCellText);
                                isCellFont = false;
                            }

                            if (isCellBorder && borderBraceCounter == 0)
                            {
                                isBottomBorder = false;
                                isCellBorder = false;

                                if (hasLeftBorder && hasTopBorder && hasRightBorder && hasBottomBorder)
                                    cellBorder.Side = StiBorderSides.All;

                                if (hasLeftBorder && !hasTopBorder && !hasRightBorder && !hasBottomBorder)
                                    cellBorder.Side = StiBorderSides.Left;

                                if (!hasLeftBorder && hasTopBorder && !hasRightBorder && !hasBottomBorder)
                                    cellBorder.Side = StiBorderSides.Top;

                                if (!hasLeftBorder && !hasTopBorder && hasRightBorder && !hasBottomBorder)
                                    cellBorder.Side = StiBorderSides.Right;

                                if (!hasLeftBorder && !hasTopBorder && !hasRightBorder && hasBottomBorder)
                                    cellBorder.Side = StiBorderSides.Bottom;

                                if (!hasLeftBorder && !hasTopBorder && !hasRightBorder && !hasBottomBorder)
                                    cellBorder.Side = StiBorderSides.None;

                                hasLeftBorder = false;
                                hasTopBorder = false;
                                hasRightBorder = false;
                                hasBottomBorder = false;
                            }

                            if (isTable && isTableCell && cellBraceCounter == 0)
                            {
                                tableCellText.ClientRectangle = new RectangleD(tableCellX, tableCellY, tableCellWidth, tableCellHeight);
                                tableCellText.Border = cellBorder;
                                tableCellText.WordWrap = true;
                                row.Add(tableCellText);
                                isTableCell = false;
                                tableCellX += tableCellWidth;

                                tableCellText = new StiText();
                                tableCellWidth = 0d;
                                tableCellHeight = 0d;
                            }

                            if (isTable && braceCounter == 0)
                            {
                                isTable = false;

                                if (row.Count > 0)
                                {
                                    TableBandsFill(isHeaderBand, isGroupHeaderBand, isDataBand, isGroupFooterBand, isFooterBand,
                                        row, tableBands, dataSource);

                                    TableBandsSort(tableBands);

                                    row.Clear();
                                    tableCellX = 0;
                                    tableCellY = 0;
                                }

                                var tablePanel = new StiPanel();

                                foreach (var band in tableBands)
                                {
                                    tablePanel.Components.Add(band);
                                }

                                allTables.Add(tablePanel);
                            }
                        }
                    }
                }
            }

            var tempTableCollection = new List<StiPanel>();

            for (int ind = allTables.Count - 1; ind >= 0; ind--)
            {
                tempTableCollection.Add(allTables[ind]);
            }

            allTables = tempTableCollection;

            for (int tableInd = allTables.Count - 1; tableInd >= 0; tableInd--)
            {
                var headerBands = new List<StiHeaderBand>();
                var groupHeaderBands = new List<StiGroupHeaderBand>();
                var dataBands = new List<StiDataBand>();
                var groupFooterBands = new List<StiGroupFooterBand>();
                var footerBands = new List<StiFooterBand>();

                foreach (var band in allTables[tableInd].Components)
                {
                    if (band is StiHeaderBand)
                        headerBands.Add(band as StiHeaderBand);

                    if (band is StiGroupHeaderBand)
                        groupHeaderBands.Add(band as StiGroupHeaderBand);

                    if (band is StiDataBand)
                        dataBands.Add(band as StiDataBand);

                    if (band is StiGroupFooterBand)
                        groupFooterBands.Add(band as StiGroupFooterBand);

                    if (band is StiFooterBand)
                        footerBands.Add(band as StiFooterBand);
                }

                allTables[tableInd].Components.Clear();

                if (headerBands.Count == 1)
                    allTables[tableInd].Components.Add(headerBands[0]);

                if (groupHeaderBands.Count == 1)
                    allTables[tableInd].Components.Add(groupHeaderBands[0]);

                if (dataBands.Count == 1)
                    allTables[tableInd].Components.Add(dataBands[0]);

                if (tableInd <= allTables.Count - 2)
                {
                    foreach (var band in allTables[tableInd + 1].Components)
                    {
                        allTables[tableInd].Components.Add(band as StiComponent);
                    }
                }

                if (groupFooterBands.Count == 1)
                    allTables[tableInd].Components.Add(groupFooterBands[0]);

                if (footerBands.Count == 1)
                    allTables[tableInd].Components.Add(footerBands[0]);
            }

            panelTables.Add(tableIndex, allTables[0]);
        }

        private void ChartLineParse(StiChart chart, List<string> chartLineCollection,
            double y, double width)
        {
            var type = "";
            var dataMode = "";
            var visMode = "";
            var height = "0";

            foreach (var line in chartLineCollection)
            {
                if (line.StartsWith("Type="))
                {
                    var typeArray = line.Split('=');

                    if (typeArray.Length == 2)
                        type = typeArray[1];
                }
                else if (line.StartsWith("DataMode="))
                {
                    var dataModeArray = line.Split('=');

                    if (dataModeArray.Length == 2)
                        dataMode = dataModeArray[1];
                }
                else if (line.StartsWith("VisMode="))
                {
                    var visModeArray = line.Split('=');

                    if (visModeArray.Length == 2)
                        visMode = visModeArray[1];
                }
                else if (line.StartsWith("Height="))
                {
                    var heightArray = line.Split('=');

                    if (heightArray.Length == 2)
                        height = heightArray[1];
                }
            }

            ChartType(chart, type, visMode, dataMode);

            chart.ClientRectangle = new RectangleD(0d, y, width, ToHi(height));
        }

        private void NestedObjectParse(string str)
        {
            if (str.StartsWith("[") && str.EndsWith("]") && !isTargetObject)
            {
                var newRoot = str.Substring(1, str.Length - 2);

                if (newRoot.StartsWith("@"))
                {
                    if (newRoot == "@Tables")
                        parsingWord = "Table";
                }

                if (newRoot == parsingWord)
                {
                    targetObject.Add(str);

                    braceCounter = 0;
                    isTargetObject = true;
                }
            }
            else if (str.StartsWith("[") && str.EndsWith("]") && isTargetObject)
            {
                targetObject.Add(str);
            }
            else
            {
                var keyPair = str.Split(new char[] { '=' }, 2);

                SectionPair sectionPair;

                sectionPair.Section = currentRoot;
                sectionPair.Key = keyPair[0];

                if (isTargetObject && (sectionPair.Key == "{" || sectionPair.Key == "}"))
                {
                    if (sectionPair.Key == "{")
                        braceCounter++;

                    if (sectionPair.Key == "}")
                        braceCounter--;

                    targetObject.Add(str);
                }
                else if (isTargetObject)
                {
                    targetObject.Add(str);
                }

                if (isTargetObject && braceCounter == 0)
                {
                    isTargetObject = false;

                    var newPartsCollection = new List<string>();

                    foreach (var partStr in targetObject)
                        newPartsCollection.Add(partStr);

                    panelNestedObjects.Add(newPartsCollection);

                    targetObject.Clear();
                }
            }
        }

        /// <summary>
        /// Returns the value for the given section, key pair.
        /// </summary>
        /// <param name="sectionName">Section name.</param>
        /// <param name="settingName">Key name.</param>
        private string GetSetting(Hashtable table, string sectionName, string settingName)
        {
            SectionPair sectionPair;
            sectionPair.Section = sectionName;
            sectionPair.Key = settingName;

            return (string)table[sectionPair];
        }
        #endregion

        #region Utils
        private void TextFormat(StiText text, string textFormat)
        {
            if (!string.IsNullOrEmpty(textFormat))
            {
                if (textFormat == "DATETIME")
                    text.TextFormat = new StiDateFormatService();

                else if (textFormat == "DOUBLE")
                    text.TextFormat = new StiNumberFormatService();

                else if (textFormat == "CURRENCY")
                    text.TextFormat = new StiCurrencyFormatService();
            }
        }

        private void ChartType(StiChart chart, string type, string visMode, string dataMode)
        {
            var chartType = "12";

            if (type == "0" || type == "5")
                chartType = $"{type}{visMode}";

            else
                chartType = $"{type}{dataMode}";

            switch (chartType)
            {
                case "00":
                    var pieArea = new StiPieArea();
                    var pieSeries = new StiPieSeries();
                    chart.Area = pieArea;
                    chart.Series.Add(pieSeries);
                    break;

                case "016":
                    var donutArea = new StiDoughnutArea();
                    var donutSeries = new StiDoughnutSeries();
                    chart.Area = donutArea;
                    chart.Series.Add(donutSeries);
                    break;

                case "01":
                    var pie3dArea = new StiPie3dArea();
                    var pie3dSeries = new StiPie3dSeries();
                    chart.Area = pie3dArea;
                    chart.Series.Add(pie3dSeries);
                    break;

                case "10":
                    var histogramAreaColumnArea = new StiHistogramArea();
                    var histogramAreaSeries = new StiHistogramSeries();
                    chart.Area = histogramAreaColumnArea;
                    chart.Series.Add(histogramAreaSeries);
                    break;

                case "12":
                    var clusteredColumnArea = new StiClusteredColumnArea();
                    var clusteredColumnSeries = new StiClusteredColumnSeries();
                    chart.Area = clusteredColumnArea;
                    chart.Series.Add(clusteredColumnSeries);
                    break;

                case "13":
                    var stackedColumnArea = new StiStackedColumnArea();
                    var stackedColumnSeries = new StiStackedColumnSeries();
                    chart.Area = stackedColumnArea;
                    chart.Series.Add(stackedColumnSeries);
                    break;

                case "20":
                    var lineArea = new StiLineArea();
                    var lineSeries = new StiLineSeries();
                    chart.Area = lineArea;
                    chart.Series.Add(lineSeries);
                    break;

                case "21":
                    var multiLineArea = new StiLineArea();
                    var multiLineSeries = new StiLineSeries();
                    chart.Area = multiLineArea;
                    chart.Series.Add(multiLineSeries);
                    break;

                case "23":
                    var fullStackedLineArea = new StiFullStackedLineArea();
                    var fullStackedLineSeries = new StiFullStackedLineSeries();
                    chart.Area = fullStackedLineArea;
                    chart.Series.Add(fullStackedLineSeries);
                    break;

                case "30":
                    var areaArea = new StiAreaArea();
                    var areaSeries = new StiAreaSeries();
                    chart.Area = areaArea;
                    chart.Series.Add(areaSeries);
                    break;

                case "33":
                    var stackedAreaArea = new StiStackedAreaArea();
                    var stackedAreaSeries = new StiStackedAreaSeries();
                    chart.Area = stackedAreaArea;
                    chart.Series.Add(stackedAreaSeries);
                    break;

                case "34":
                    var fullStackedAreaArea = new StiFullStackedAreaArea();
                    var fullStackedAreaSeries = new StiFullStackedAreaSeries();
                    chart.Area = fullStackedAreaArea;
                    chart.Series.Add(fullStackedAreaSeries);
                    break;

                case "45":
                    var BubbleArea = new StiBubbleArea();
                    var BubbleSeries = new StiBubbleSeries();
                    chart.Area = BubbleArea;
                    chart.Series.Add(BubbleSeries);
                    break;

                case "50":
                    var funnelArea = new StiFunnelArea();
                    var funnelSeries = new StiFunnelSeries();
                    chart.Area = funnelArea;
                    chart.Series.Add(funnelSeries);
                    break;

                case "70":
                    var radarLineArea = new StiRadarLineArea();
                    var radarLineSeries = new StiRadarLineSeries();
                    chart.Area = radarLineArea;
                    chart.Series.Add(radarLineSeries);
                    break;

                case "80":
                    var treemapArea = new StiTreemapArea();
                    var treemapSeries = new StiTreemapSeries();
                    chart.Area = treemapArea;
                    chart.Series.Add(treemapSeries);
                    break;
            }
        }

        private StiTextHorAlignment TextHorAlignment(string value)
        {
            var align = StiTextHorAlignment.Left;

            switch (value)
            {
                case "0":
                    align = StiTextHorAlignment.Left;
                    break;

                case "1":
                    align = StiTextHorAlignment.Center;
                    break;

                case "2":
                    align = StiTextHorAlignment.Right;
                    break;

                case "3":
                    align = StiTextHorAlignment.Width;
                    break;
            }

            return align;
        }

        private List<List<string>> ChildTables(List<string> tableLineCollection)
        {
            var childTableLinesCollection = new List<List<string>>();
            var fullTable = tableLineCollection;
            var firstIndexes = new List<int>();
            var lastIndexes = new List<int>();
            var isChildTable = false;
            var childTableLines = new List<string>();
            var childTableBraceCounter = 0;

            for (int index = 0; index < fullTable.Count; index++)
            {
                if (fullTable[index] == "[ChildTables]" && !isChildTable)
                    firstIndexes.Add(index);
            }

            for (int ind = firstIndexes.Count - 1; ind >= 0; ind--)
            {
                for (int index = firstIndexes[ind]; index < fullTable.Count; index++)
                {
                    if (isChildTable && (fullTable[index] == "{" || fullTable[index] == "}"))
                    {
                        if (fullTable[index] == "{")
                            childTableBraceCounter++;

                        if (fullTable[index] == "}")
                            childTableBraceCounter--;
                    }

                    if (isChildTable && childTableBraceCounter == 0)
                    {
                        lastIndexes.Add(index);
                        isChildTable = false;
                    }

                    if (fullTable[index] == "[ChildTables]" && !isChildTable)
                    {
                        isChildTable = true;
                        childTableBraceCounter = 0;
                        childTableLines = new List<string>();
                    }

                    if (isChildTable)
                        childTableLines.Add(fullTable[index]);
                }

                if (lastIndexes.Count > 0)
                    fullTable.RemoveRange(firstIndexes[ind], lastIndexes[lastIndexes.Count - 1] - firstIndexes[ind]);

                childTableLinesCollection.Add(childTableLines);
            }
            childTableLinesCollection.Add(fullTable);

            return childTableLinesCollection;
        }

        private void TableBandsFill(bool isHeaderBand, bool isGroupHeaderBand, bool isDataBand, bool isGroupFooterBand,
            bool isFooterBand, List<StiComponent> row, List<StiBand> tableBands, string dataSource)
        {
            var linePanel = new StiPanel();
            var lines = new List<StiPanel>();
            var cellY = 0d;
            var linePanelWidth = 0d;
            var linePanelHeight = 0d;
            var cellIndex = 1;

            if (row.Count > 0)
                cellY = row[0].ClientRectangle.Y;

            foreach (var cell in row)
            {
                cell.Page = page;

                if (cell.ClientRectangle.Y > cellY)
                {
                    linePanel.CanGrow = false;
                    linePanel.CanShrink = false;
                    linePanel.ClientRectangle = new RectangleD(
                        0,
                        cellY,
                        linePanelWidth,
                        linePanelHeight);

                    lines.Add(linePanel);
                    linePanel = new StiPanel();
                    linePanelWidth = 0d;
                    linePanelHeight = 0d;

                    cellY = cell.ClientRectangle.Y;

                    cell.Name = $"{cellIndex}";
                    cellIndex++;
                    linePanel.Components.Add(cell);
                    linePanelWidth += cell.ClientRectangle.Width;
                    linePanelHeight = cell.ClientRectangle.Height;
                }
                else
                {
                    cell.Name = $"{cellIndex}";
                    cellIndex++;
                    linePanel.Components.Add(cell);
                    linePanelWidth += cell.ClientRectangle.Width;
                    linePanelHeight = cell.ClientRectangle.Height;
                }
            }

            if (linePanel.Components.Count > 0)
            {
                linePanel.CanGrow = false;
                linePanel.CanShrink = false;
                linePanel.ClientRectangle = new RectangleD(
                    0,
                    cellY,
                    linePanelWidth,
                    linePanelHeight);

                lines.Add(linePanel);
            }

            if (lines.Count > 1)
            {
                foreach (var line in lines)
                {
                    foreach (StiComponent cell in line.Components)
                    {
                        cell.GrowToHeight = true;
                        cell.ClientRectangle = new RectangleD(
                            cell.ClientRectangle.X,
                            0,
                            cell.ClientRectangle.Width,
                            cell.ClientRectangle.Height);
                    }
                }

                row.Clear();

                foreach (var line in lines)
                {
                    row.Add(line);
                }
            }

            if (isHeaderBand)
            {
                var headerBand = new StiHeaderBand();
                headerBand.Name = $"HeaderBand{headerBandIndex}";
                headerBandIndex++;

                BandFill(headerBand, row);

                tableBands.Add(headerBand);
            }
            else if (isGroupHeaderBand)
            {
                var groupHeaderBand = new StiGroupHeaderBand();
                groupHeaderBand.Name = $"GroupHeaderBand{groupHeaderBandIndex}";
                groupHeaderBandIndex++;

                BandFill(groupHeaderBand, row);

                tableBands.Add(groupHeaderBand);
            }
            else if (isDataBand)
            {
                var dataBand = new StiDataBand();
                dataBand.Name = $"DataBand{dataBandIndex}";
                dataBandIndex++;

                if (!string.IsNullOrEmpty(dataSource))
                    dataBand.DataSourceName = dataSource;

                BandFill(dataBand, row);

                tableBands.Add(dataBand);
            }
            else if (isGroupFooterBand)
            {
                var groupFooterBand = new StiGroupFooterBand();
                groupFooterBand.Name = $"GroupFooterBand{groupFooterBandIndex}";
                groupFooterBandIndex++;

                BandFill(groupFooterBand, row);

                tableBands.Add(groupFooterBand);
            }
            else if (isFooterBand)
            {
                var footerBand = new StiFooterBand();
                footerBand.Name = $"FooterBand{footerBandIndex}";
                footerBandIndex++;

                BandFill(footerBand, row);

                tableBands.Add(footerBand);
            }
        }

        private void TableBandsSort(List<StiBand> tableBands)
        {
            var headerBands = new List<StiHeaderBand>();
            var groupHeaderBands = new List<StiGroupHeaderBand>();
            var dataBands = new List<StiDataBand>();
            var groupFooterBands = new List<StiGroupFooterBand>();
            var footerBands = new List<StiFooterBand>();

            foreach (var band in tableBands)
            {
                if (band is StiHeaderBand)
                    headerBands.Add(band as StiHeaderBand);

                if (band is StiGroupHeaderBand)
                    groupHeaderBands.Add(band as StiGroupHeaderBand);

                if (band is StiDataBand)
                    dataBands.Add(band as StiDataBand);

                if (band is StiGroupFooterBand)
                    groupFooterBands.Add(band as StiGroupFooterBand);

                if (band is StiFooterBand)
                    footerBands.Add(band as StiFooterBand);
            }

            tableBands.Clear();

            if (headerBands.Count == 1)
                tableBands.Add(headerBands[0]);

            if (groupHeaderBands.Count == 1)
                tableBands.Add(groupHeaderBands[0]);

            if (dataBands.Count == 1)
                tableBands.Add(dataBands[0]);

            if (groupFooterBands.Count == 1)
                tableBands.Add(groupFooterBands[0]);

            if (footerBands.Count == 1)
                tableBands.Add(footerBands[0]);
        }

        private StiPanel TableBuilder(StiPage page, StiPanel table, StiPanel container, double x = 0d, double y = 0d)
        {
            var childY = 0d;
            var tableWidth = 0d;
            var tableHeight = 0d;
            var tablePanel = new StiPanel();
            tablePanel.Name = $"Table{tableIndex}";
            tableIndex++;
            tablePanel.CanGrow = false;
            tablePanel.CanShrink = false;
            tablePanel.Page = page;

            for (int bandIndex = 0; bandIndex < table.Components.Count; bandIndex++)
            {
                var bandHeight = 0d;
                var bandWidth = 0d;

                for (var cellIndex = 0; cellIndex < (table.Components[bandIndex] as StiBand).Components.Count; cellIndex++)
                {
                    var cell = (table.Components[bandIndex] as StiBand).Components[cellIndex];
                    cell.Name = $"{table.Components[bandIndex].Name}_Cell{cellIndex + 1}";
                    cell.Page = page;
                    bandHeight = cell.ClientRectangle.Y + cell.ClientRectangle.Height;

                    if (cell is StiPanel)
                        bandWidth = cell.ClientRectangle.Width;

                    else
                        bandWidth += cell.ClientRectangle.Width;
                }

                if (bandWidth > container.ClientRectangle.Width)
                    bandWidth = container.ClientRectangle.Width;

                table.Components[bandIndex].ClientRectangle = new RectangleD(0, childY, bandWidth, bandHeight);
                table.Components[bandIndex].Linked = true;
                tablePanel.Components.Add(table.Components[bandIndex]);

                if (tableWidth < table.Components[bandIndex].ClientRectangle.Width)
                    tableWidth = table.Components[bandIndex].ClientRectangle.Width;

                tableHeight += table.Components[bandIndex].ClientRectangle.Height;
                childY += bandHeight;
            }

            tablePanel.ClientRectangle = new RectangleD(x, y, tableWidth, tableHeight);

            return tablePanel;
        }

        private void BandFill(StiBand band, List<StiComponent> row)
        {
            var cellHeight = 0d;

            foreach (var cell in row)
            {
                if (cell.ClientRectangle.Height > cellHeight)
                    cellHeight = cell.ClientRectangle.Height;
            }

            foreach (var cell in row)
            {
                cell.ClientRectangle = new RectangleD(
                    cell.ClientRectangle.X,
                    cell.ClientRectangle.Y,
                    cell.ClientRectangle.Width,
                    cellHeight);

                cell.Linked = true;
                band.Components.Add(cell);
            }

            band.Height = cellHeight;
        }

        private void CrossTabBuilder(StiPage page, Hashtable obj, StiCrossTab crossTable, List<StiText> rowFields,
            List<StiText> columnFields, List<StiText> dataFields, bool isEmbeddedCrossTab = false)
        {
            var rowTotalMaxYCoordinate = 0d;
            var crossRowX = 0d;
            var defaultRowHeight = 5d;
            var defaultCellMargin = 0.001d;
            var rowWidthCollection = new List<double>();
            var allRowsWidth = 0d;
            var measureCoeff = 0.35d;
            var cells = new string[rowFields.Count, columnFields.Count];

            for (int index = 0; index < rowFields.Count; index++)
            {
                var width = MeasureSize(rowFields[index].Text, defaultFont).Width * measureCoeff;

                if (width < 30)
                    width = 30;

                allRowsWidth += width;
                rowWidthCollection.Add(width);
            }

            var leftTitleText = "";

            if (rowFields.Count > 0 && columnFields.Count > 0 && cells[0, 0] != null)
                leftTitleText = cells[0, 0].Replace("{", "").Replace("}", "");

            var leftTitleWidth = MeasureSize(leftTitleText, defaultFont);

            if (allRowsWidth < leftTitleWidth.Width * measureCoeff)
            {
                var difference = (leftTitleWidth.Width * measureCoeff - allRowsWidth) / rowWidthCollection.Count;

                allRowsWidth = leftTitleWidth.Width * measureCoeff;

                for (int index = 0; index < rowWidthCollection.Count; index++)
                {
                    rowWidthCollection[index] += difference;
                }
            }

            #region Rows
            for (int index = 0; index < rowFields.Count; index++)
            {
                var rowWidth = ParseDouble(rowWidthCollection[index].ToString());
                var rowHeight = defaultRowHeight;

                if (dataFields.Count > 0)
                    rowHeight = defaultRowHeight * dataFields.Count;

                var rowTotal = new StiCrossRowTotal();
                rowTotal.Font = defaultFont;
                rowTotal.Name = $"{crossTable.Name}_RowTotal{index}";
                rowTotal.Page = page;

                var row = new StiCrossRow();

                row.Alias = rowFields[index].Text.ToString().Replace("[", "").Replace("]", "").Replace("\"", "");
                row.ClientRectangle = new RectangleD
                    (
                    crossRowX,
                    10 + defaultCellMargin + defaultCellMargin + defaultRowHeight * columnFields.Count,
                    rowWidth,
                    rowHeight * (rowFields.Count - index)
                    );
                row.DisplayValue = new StiDisplayCrossValueExpression(rowFields[index].Text.ToString().Replace(" ", ""));
                row.Font = defaultFont;
                row.Name = $"{crossTable.Name}_Row{index}";
                row.Page = page;
                row.Value = new StiCrossValueExpression(rowFields[index].Text.ToString().Replace(" ", ""));
                row.TextFormat = rowFields[index].TextFormat;
                row.Guid = Guid.NewGuid().ToString();
                row.TotalGuid = rowTotal.Guid;

                var rowTotalWidth = 0d;

                if (index > 0 && index != rowWidthCollection.Count - 1)
                    rowTotalWidth = allRowsWidth - ParseDouble(rowWidthCollection[index - 1].ToString());

                else if (index > 0 && index == rowWidthCollection.Count - 1)
                    rowTotalWidth = row.ClientRectangle.Width;

                else
                    rowTotalWidth = allRowsWidth;

                rowTotal.ClientRectangle = new RectangleD
                    (
                    crossRowX,
                    row.ClientRectangle.Y + row.ClientRectangle.Height,
                    rowTotalWidth,
                    rowHeight
                    );

                if (rowTotal.ClientRectangle.Y + rowTotal.ClientRectangle.Height > rowTotalMaxYCoordinate)
                    rowTotalMaxYCoordinate = rowTotal.ClientRectangle.Y + rowTotal.ClientRectangle.Height;

                crossTable.Components.Add(rowTotal);
                crossTable.Components.Add(row);

                var crossTitle = new StiCrossTitle();
                crossTitle.Name = $"{crossTable.Name}_Row{index}_Title";
                crossTitle.ClientRectangle = new RectangleD
                    (
                    crossRowX,
                    10 + defaultCellMargin,
                    rowWidth,
                    defaultRowHeight * columnFields.Count
                    );
                crossTitle.TypeOfComponent = $"Row:{crossTable.Name}_Row{index}";
                crossTitle.Font = defaultFont;
                crossTitle.Page = page;

                var rowValue = row.Value.ToString().Split('.');

                if (rowValue.Length > 0)
                    crossTitle.Text = rowValue[rowValue.Length - 1].Replace("{", "").Replace("}", "");

                crossTitle.Guid = Guid.NewGuid().ToString();

                crossRowX = crossTitle.ClientRectangle.X + crossTitle.ClientRectangle.Width;

                crossTable.Components.Add(crossTitle);
            }
            #endregion

            var leftTitle = new StiCrossTitle();
            leftTitle.Name = $"{crossTable.Name}_LeftTitle";
            leftTitle.TypeOfComponent = "LeftTitle";
            leftTitle.ClientRectangle = new RectangleD
                (
                0,
                0,
                crossRowX,
                10
                );
            leftTitle.Font = defaultFont;

            leftTitle.Page = page;

            if (rowFields.Count > 0 && columnFields.Count > 0 && cells[0, 0] != null)
                leftTitle.Text = cells[0, 0].Replace("{", "").Replace("}", "");

            leftTitle.Guid = Guid.NewGuid().ToString();

            var rightTitle = new StiCrossTitle();
            rightTitle.Name = $"{crossTable.Name}_RightTitle";
            rightTitle.TypeOfComponent = "RightTitle";
            rightTitle.Font = defaultFont;
            rightTitle.Page = page;
            rightTitle.Guid = Guid.NewGuid().ToString();

            var columnTotalMaxXCoordinate = 0d;
            var lastColumnWidthForData = 20d;
            var columnWidthCollection = new List<double>();

            for (int index = 0; index < columnFields.Count; index++)
            {
                var value = columnFields[index].Text.ToString().Replace("[", "").Replace("]", "").Replace("\"", "");
                var width = MeasureSize(value, defaultFont).Width * measureCoeff;
                columnWidthCollection.Add(width);
            }

            var columnTotalWidthCollection = new double[columnWidthCollection.Count];

            for (int ind = columnWidthCollection.Count - 1; ind >= 0; ind--)
            {
                var width = columnWidthCollection[ind];
                var totalWidth = 0d;

                if (ind > 0)
                {
                    var nextWidth = columnWidthCollection[ind - 1];

                    if (nextWidth > width)
                    {
                        totalWidth = nextWidth - width;

                        if (totalWidth > 30)
                        {
                            columnWidthCollection[ind] = width + (totalWidth - 30);
                            totalWidth = 30;
                        }
                    }
                    else
                    {
                        totalWidth = 30;
                        columnWidthCollection[ind - 1] = width + totalWidth;
                    }
                }
                else
                {
                    totalWidth = 30;
                }

                columnTotalWidthCollection[ind] = totalWidth;
            }

            var allColumnTextsForRightTitle = "";

            #region Columns
            for (int index = 0; index < columnFields.Count; index++)
            {
                var columnWidth = ParseDouble(columnWidthCollection[index].ToString());
                var totalWidth = ParseDouble(columnTotalWidthCollection[index].ToString());

                var columnTotal = new StiCrossColumnTotal();
                columnTotal.Guid = Guid.NewGuid().ToString();

                var column = new StiCrossColumn();
                column.Alias = columnFields[index].Text.ToString().Replace("[", "").Replace("]", "").Replace("\"", "");

                if (index < columnFields.Count - 1)
                    allColumnTextsForRightTitle += $"{column.Alias}, ";

                else
                    allColumnTextsForRightTitle += $"{column.Alias}";

                column.ClientRectangle = new RectangleD
                    (
                    leftTitle.ClientRectangle.X + leftTitle.ClientRectangle.Width + defaultCellMargin,
                    10 + defaultCellMargin + defaultRowHeight * index,
                    columnWidth,
                    defaultRowHeight
                    );
                column.DisplayValue = new StiDisplayCrossValueExpression(columnFields[index].Text.ToString().Replace(" ", ""));
                column.Font = defaultFont;
                column.Name = $"{crossTable.Name}_Column{index}";
                column.Page = page;
                column.Value = new StiCrossValueExpression(columnFields[index].Text.ToString().Replace(" ", ""));
                column.TextFormat = columnFields[index].TextFormat;
                column.Guid = Guid.NewGuid().ToString();
                column.TotalGuid = columnTotal.Guid;

                crossTable.Components.Add(column);

                columnTotal.ClientRectangle = new RectangleD
                    (
                    leftTitle.ClientRectangle.X + leftTitle.ClientRectangle.Width + column.ClientRectangle.Width + defaultCellMargin,
                    column.ClientRectangle.Y,
                    totalWidth,
                    defaultRowHeight * (columnFields.Count - index)
                    );
                columnTotal.Font = defaultFont;
                columnTotal.Name = $"{crossTable.Name}_ColTotal{index}";
                columnTotal.Page = page;

                if (columnTotal.ClientRectangle.X + columnTotal.ClientRectangle.Width > columnTotalMaxXCoordinate)
                    columnTotalMaxXCoordinate = columnTotal.ClientRectangle.X + columnTotal.ClientRectangle.Width;

                crossTable.Components.Add(columnTotal);

                if (index == columnFields.Count - 1)
                    lastColumnWidthForData = column.ClientRectangle.Width;
            }
            #endregion

            rightTitle.Text = allColumnTextsForRightTitle;

            if (columnWidthCollection.Count > 0)
            {
                rightTitle.ClientRectangle = new RectangleD
                (
                leftTitle.ClientRectangle.Width + defaultCellMargin,
                0,
                columnWidthCollection[0] + 30,
                10
                );
            }

            #region Data
            for (int index = 0; index < dataFields.Count; index++)
            {
                var summary = new StiCrossSummary();

                if (dataFields.Count > 0)
                {
                    summary.Alias = dataFields[index].Text.ToString().Replace("[", "").Replace("]", "").Replace("\"", "");
                    summary.ClientRectangle = new RectangleD
                        (
                        leftTitle.ClientRectangle.Width + defaultCellMargin,
                        10 + defaultCellMargin + defaultCellMargin + defaultRowHeight * columnFields.Count + defaultRowHeight * index,
                        lastColumnWidthForData,
                        defaultRowHeight
                        );
                    summary.Font = defaultFont;
                    summary.Name = $"{crossTable.Name}_Sum{index}";
                    summary.Page = page;
                    summary.Value = new StiCrossValueExpression(dataFields[index].Text.ToString().Replace(" ", ""));
                    summary.TextFormat = dataFields[index].TextFormat;
                    summary.Guid = Guid.NewGuid().ToString();

                    /*if (summaryTypeCollection.Count == 1)
                    {
                        if (summaryTypeCollection[0] == "Count")
                            summary.Summary = CrossTab.Core.StiSummaryType.Count;

                        else if (summaryTypeCollection[0] == "Avg")
                            summary.Summary = CrossTab.Core.StiSummaryType.Average;

                        else if (summaryTypeCollection[0] == "Sum")
                            summary.Summary = CrossTab.Core.StiSummaryType.Sum;

                        else if (summaryTypeCollection[0] == "Min")
                            summary.Summary = CrossTab.Core.StiSummaryType.Min;

                        else if (summaryTypeCollection[0] == "Max")
                            summary.Summary = CrossTab.Core.StiSummaryType.Max;
                    }
                    else if (summaryTypeCollection.Count == 0)
                    {
                        summary.Summary = CrossTab.Core.StiSummaryType.Sum;
                    }*/
                }
                else if (dataFields.Count == 0)
                {
                    summary.Alias = $"{crossTable.Name}_Sum1";
                    summary.ClientRectangle = new RectangleD
                        (
                        32,
                        34,
                        lastColumnWidthForData,
                        defaultRowHeight);
                    summary.Name = $"{crossTable.Name}_Sum1";
                    summary.Guid = Guid.NewGuid().ToString();
                }

                crossTable.Components.Add(summary);
            }
            #endregion

            crossTable.Components.Add(leftTitle);
            crossTable.Components.Add(rightTitle);

            if (!isEmbeddedCrossTab)
                ClientRectangle(obj, crossTable);

            if (crossTable.Height == 0)
                crossTable.Height = rowTotalMaxYCoordinate;

            if (crossTable.Width == 0)
                crossTable.Width = columnTotalMaxXCoordinate;
        }

        private SizeF MeasureSize(string text, Font font)
        {
            Bitmap img = new Bitmap(1, 1);
            Graphics g = Graphics.FromImage(img);
            StringFormat sf = new StringFormat(StringFormat.GenericTypographic);
            g.PageUnit = GraphicsUnit.Point;
            var result = g.MeasureString(text, font, 10000000, sf);

            return result;
        }

        private void BarcodeType(StiBarCode component, string typeName)
        {
            if (typeName == "17") component.BarCodeType = new StiCode93BarCodeType();
            else if (typeName == "18") component.BarCodeType = new StiMsiBarCodeType();
            else if (typeName == "1") component.BarCodeType = new StiEAN8BarCodeType();
            else if (typeName == "0") component.BarCodeType = new StiEAN13BarCodeType();
            else if (typeName == "12") component.BarCodeType = new StiEAN128AutoBarCodeType();
            else if (typeName == "9") component.BarCodeType = new StiPostnetBarCodeType();
            else if (typeName == "2") component.BarCodeType = new StiUpcABarCodeType();
            else if (typeName == "26") component.BarCodeType = new StiCode39BarCodeType();
            else if (typeName == "29") component.BarCodeType = new StiCode39ExtBarCodeType();
            else if (typeName == "11") component.BarCodeType = new StiCodabarBarCodeType();
            else if (typeName == "19") component.BarCodeType = new StiCode11BarCodeType();
            else if (typeName == "6") component.BarCodeType = new StiInterleaved2of5BarCodeType();
            else if (typeName == "64") component.BarCodeType = new StiPdf417BarCodeType();
            else if (typeName == "68") component.BarCodeType = new StiDataMatrixBarCodeType();
            else if (typeName == "70") component.BarCodeType = new StiQRCodeBarCodeType();
            else if (typeName == "41") component.BarCodeType = new StiIntelligentMail4StateBarCodeType();
            else if (typeName == "46") component.BarCodeType = new StiPharmacodeBarCodeType();
            else if (typeName == "13") component.BarCodeType = new StiCode128AutoBarCodeType();
            else if (typeName == "8") component.BarCodeType = new StiStandard2of5BarCodeType();
            else component.BarCodeType = new StiCode93BarCodeType();
        }

        private StiPenStyle ParseBorderStyle(string styleName)
        {
            switch (styleName)
            {
                case "None":
                    return StiPenStyle.None;
                case "2":
                    return StiPenStyle.Dash;
                case "1":
                    return StiPenStyle.Dot;
                case "3":
                    return StiPenStyle.DashDot;
                case "4":
                    return StiPenStyle.DashDotDot;
            }
            return StiPenStyle.Solid;
        }

        private StiShape ShapeLineComponent(string lineDirection)
        {
            var component = new StiShape { ShapeType = new StiHorizontalLineShapeType() };

            if (lineDirection == "2")
                component = new StiShape { ShapeType = new StiHorizontalLineShapeType() };

            else if (lineDirection == "1")
                component = new StiShape { ShapeType = new StiDiagonalUpLineShapeType() };

            else if (lineDirection == "0")
                component = new StiShape { ShapeType = new StiDiagonalDownLineShapeType() };

            else if (lineDirection == "3")
                component = new StiShape { ShapeType = new StiVerticalLineShapeType() };

            return component;
        }

        private Hashtable HtmlNameToColor = null;

        private Color ParseStyleColor(string colorAttribute)
        {
            Color color = Color.Transparent;
            if (colorAttribute != null && colorAttribute.Length > 1)
            {
                if (colorAttribute.StartsWith("LL.Color."))
                {
                    var correctColor = colorAttribute.Replace("LL.Color.", "");
                    colorAttribute = correctColor;
                }

                if (colorAttribute.StartsWith("RGB"))
                {
                    var correctColor = colorAttribute.Remove(0, 3).Replace("(", "").Replace(")", "");
                    colorAttribute = correctColor;
                }

                if (colorAttribute[0] == '#')
                {
                    #region Parse RGB value in hexadecimal notation
                    string colorSt = colorAttribute.Substring(1).ToLowerInvariant();
                    StringBuilder sbc = new StringBuilder();
                    foreach (char ch in colorSt)
                    {
                        if (ch == '0' || ch == '1' || ch == '2' || ch == '3' || ch == '4' || ch == '5' || ch == '6' || ch == '7' ||
                            ch == '8' || ch == '9' || ch == 'a' || ch == 'b' || ch == 'c' || ch == 'd' || ch == 'e'
                            || ch == 'f') sbc.Append(ch);
                    }
                    if (sbc.Length == 3)
                    {
                        colorSt = string.Format("{0}{0}{1}{1}{2}{2}", sbc[0], sbc[1], sbc[2]);
                    }
                    else
                    {
                        colorSt = sbc.ToString();
                    }
                    if (colorSt.Length == 6)
                    {
                        int colorInt = Convert.ToInt32(colorSt, 16);
                        color = Color.FromArgb(0xFF, colorInt >> 16 & 0xFF, colorInt >> 8 & 0xFF, colorInt & 0xFF);
                    }
                    #endregion
                }
                else if (colorAttribute.StartsWith("rgb"))
                {
                    #region Parse RGB function
                    string[] colors = colorAttribute.Substring(4, colorAttribute.Length - 5).Split(new char[] { ',' });
                    color = Color.FromArgb(0xFF, int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]));
                    #endregion
                }
                else if (colorAttribute.IndexOf(',') != -1)
                {
                    #region Parse RGB values
                    string[] colors = colorAttribute.Split(new char[] { ',' });

                    if (colors.Length == 3)
                        color = Color.FromArgb(0xFF, int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]));

                    else if (colors.Length == 4)
                        color = Color.FromArgb(int.Parse(colors[0]), int.Parse(colors[1]), int.Parse(colors[2]), int.Parse(colors[3]));
                    #endregion
                }
                else
                {
                    #region Parse color keywords
                    if (HtmlNameToColor == null)
                    {
                        #region Init hashtable
                        string[,] initData = {
                        {"AliceBlue",       "#F0F8FF"},
                        {"AntiqueWhite",    "#FAEBD7"},
                        {"Aqua",        "#00FFFF"},
                        {"Aquamarine",  "#7FFFD4"},
                        {"Azure",       "#F0FFFF"},
                        {"Beige",       "#F5F5DC"},
                        {"Bisque",      "#FFE4C4"},
                        {"Black",       "#000000"},
                        {"BlanchedAlmond",  "#FFEBCD"},
                        {"Blue",        "#0000FF"},
                        {"BlueViolet",  "#8A2BE2"},
                        {"Brown",       "#A52A2A"},
                        {"BurlyWood",   "#DEB887"},
                        {"CadetBlue",   "#5F9EA0"},
                        {"Chartreuse",  "#7FFF00"},
                        {"Chocolate",   "#D2691E"},
                        {"Coral",       "#FF7F50"},
                        {"CornflowerBlue",  "#6495ED"},
                        {"Cornsilk",    "#FFF8DC"},
                        {"Crimson",     "#DC143C"},
                        {"Cyan",        "#00FFFF"},
                        {"DarkBlue",    "#00008B"},
                        {"DarkCyan",    "#008B8B"},
                        {"DarkGoldenRod",   "#B8860B"},
                        {"DarkGray",    "#A9A9A9"},
                        {"DarkGrey",    "#A9A9A9"},
                        {"DarkGreen",   "#006400"},
                        {"DarkKhaki",   "#BDB76B"},
                        {"DarkMagenta", "#8B008B"},
                        {"DarkOliveGreen",  "#556B2F"},
                        {"Darkorange",  "#FF8C00"},
                        {"DarkOrchid",  "#9932CC"},
                        {"DarkRed",     "#8B0000"},
                        {"DarkSalmon",  "#E9967A"},
                        {"DarkSeaGreen",    "#8FBC8F"},
                        {"DarkSlateBlue",   "#483D8B"},
                        {"DarkSlateGray",   "#2F4F4F"},
                        {"DarkSlateGrey",   "#2F4F4F"},
                        {"DarkTurquoise",   "#00CED1"},
                        {"DarkViolet",  "#9400D3"},
                        {"DeepPink",    "#FF1493"},
                        {"DeepSkyBlue", "#00BFFF"},
                        {"DimGray",     "#696969"},
                        {"DimGrey",     "#696969"},
                        {"DodgerBlue",  "#1E90FF"},
                        {"FireBrick",   "#B22222"},
                        {"FloralWhite", "#FFFAF0"},
                        {"ForestGreen", "#228B22"},
                        {"Fuchsia",     "#FF00FF"},
                        {"Gainsboro",   "#DCDCDC"},
                        {"GhostWhite",  "#F8F8FF"},
                        {"Gold",        "#FFD700"},
                        {"GoldenRod",   "#DAA520"},
                        {"Gray",        "#808080"},
                        {"Grey",        "#808080"},
                        {"Green",       "#008000"},
                        {"GreenYellow", "#ADFF2F"},
                        {"HoneyDew",    "#F0FFF0"},
                        {"HotPink",     "#FF69B4"},
                        {"IndianRed",   "#CD5C5C"},
                        {"Indigo",      "#4B0082"},
                        {"Ivory",       "#FFFFF0"},
                        {"Khaki",       "#F0E68C"},
                        {"Lavender",    "#E6E6FA"},
                        {"LavenderBlush",   "#FFF0F5"},
                        {"LawnGreen",   "#7CFC00"},
                        {"LemonChiffon",    "#FFFACD"},
                        {"LightBlue",   "#ADD8E6"},
                        {"LightCoral",  "#F08080"},
                        {"LightCyan",   "#E0FFFF"},
                        {"LightGoldenRodYellow",    "#FAFAD2"},
                        {"LightGray",   "#D3D3D3"},
                        {"LightGrey",   "#D3D3D3"},
                        {"LightGreen",  "#90EE90"},
                        {"LightPink",   "#FFB6C1"},
                        {"LightSalmon", "#FFA07A"},
                        {"LightSeaGreen",   "#20B2AA"},
                        {"LightSkyBlue",    "#87CEFA"},
                        {"LightSlateGray",  "#778899"},
                        {"LightSlateGrey",  "#778899"},
                        {"LightSteelBlue",  "#B0C4DE"},
                        {"LightYellow", "#FFFFE0"},
                        {"Lime",        "#00FF00"},
                        {"LimeGreen",   "#32CD32"},
                        {"Linen",       "#FAF0E6"},
                        {"Magenta",     "#FF00FF"},
                        {"Maroon",      "#800000"},
                        {"MediumAquaMarine",    "#66CDAA"},
                        {"MediumBlue",  "#0000CD"},
                        {"MediumOrchid",    "#BA55D3"},
                        {"MediumPurple",    "#9370D8"},
                        {"MediumSeaGreen",  "#3CB371"},
                        {"MediumSlateBlue", "#7B68EE"},
                        {"MediumSpringGreen",   "#00FA9A"},
                        {"MediumTurquoise", "#48D1CC"},
                        {"MediumVioletRed", "#C71585"},
                        {"MidnightBlue",    "#191970"},
                        {"MintCream",   "#F5FFFA"},
                        {"MistyRose",   "#FFE4E1"},
                        {"Moccasin",    "#FFE4B5"},
                        {"NavajoWhite", "#FFDEAD"},
                        {"Navy",        "#000080"},
                        {"OldLace",     "#FDF5E6"},
                        {"Olive",       "#808000"},
                        {"OliveDrab",   "#6B8E23"},
                        {"Orange",      "#FFA500"},
                        {"OrangeRed",   "#FF4500"},
                        {"Orchid",      "#DA70D6"},
                        {"PaleGoldenRod",   "#EEE8AA"},
                        {"PaleGreen",   "#98FB98"},
                        {"PaleTurquoise",   "#AFEEEE"},
                        {"PaleVioletRed",   "#D87093"},
                        {"PapayaWhip",  "#FFEFD5"},
                        {"PeachPuff",   "#FFDAB9"},
                        {"Peru",        "#CD853F"},
                        {"Pink",        "#FFC0CB"},
                        {"Plum",        "#DDA0DD"},
                        {"PowderBlue",  "#B0E0E6"},
                        {"Purple",      "#800080"},
                        {"Red",         "#FF0000"},
                        {"RosyBrown",   "#BC8F8F"},
                        {"RoyalBlue",   "#4169E1"},
                        {"SaddleBrown", "#8B4513"},
                        {"Salmon",      "#FA8072"},
                        {"SandyBrown",  "#F4A460"},
                        {"SeaGreen",    "#2E8B57"},
                        {"SeaShell",    "#FFF5EE"},
                        {"Sienna",      "#A0522D"},
                        {"Silver",      "#C0C0C0"},
                        {"SkyBlue",     "#87CEEB"},
                        {"SlateBlue",   "#6A5ACD"},
                        {"SlateGray",   "#708090"},
                        {"SlateGrey",   "#708090"},
                        {"Snow",        "#FFFAFA"},
                        {"SpringGreen", "#00FF7F"},
                        {"SteelBlue",   "#4682B4"},
                        {"Tan",         "#D2B48C"},
                        {"Teal",        "#008080"},
                        {"Thistle",     "#D8BFD8"},
                        {"Tomato",      "#FF6347"},
                        {"Turquoise",   "#40E0D0"},
                        {"Violet",      "#EE82EE"},
                        {"Wheat",       "#F5DEB3"},
                        {"White",       "#FFFFFF"},
                        {"WhiteSmoke",  "#F5F5F5"},
                        {"Yellow",      "#FFFF00"},
                        {"YellowGreen", "#9ACD32"}};

                        HtmlNameToColor = new Hashtable();

                        for (int index = 0; index < initData.GetLength(0); index++)
                        {
                            string key = initData[index, 0].ToLowerInvariant();
                            int colorInt = Convert.ToInt32(initData[index, 1].Substring(1), 16);
                            Color value = Color.FromArgb(0xFF, colorInt >> 16 & 0xFF, colorInt >> 8 & 0xFF, colorInt & 0xFF);
                            HtmlNameToColor[key] = value;
                        }
                        #endregion
                    }

                    string colorSt = colorAttribute.ToLowerInvariant();

                    if (HtmlNameToColor.ContainsKey(colorSt))
                        color = (Color)HtmlNameToColor[colorSt];
                    #endregion
                }
            }
            return color;
        }

        private void Font(Hashtable obj, IStiFont component)
        {
            var font = defaultFont;

            var useDefaultFont = GetSetting(obj, "Object", "Default");
            var fontFamily = GetSetting(obj, "Object", "FaceName").Replace("\"", "");
            var fontSize = 12f;
            var fontSizeString = GetSetting(obj, "Object", "Size");

            if (fontSizeString != "Null()")
                fontSize = (float)ParseDouble(fontSizeString);

            var isBold = GetSetting(obj, "Object", "Bold");
            var isItalic = GetSetting(obj, "Object", "Italic");
            var isUnderline = GetSetting(obj, "Object", "Underline");
            var isStrikeout = GetSetting(obj, "Object", "Strikeout");

            var colorString = GetSetting(obj, "Object", "Color");
            var color = Color.Black;

            if (colorString != "Null()")
                color = ParseStyleColor(colorString);

            if (useDefaultFont == "False")
                font = new Font(fontFamily, fontSize);

            if (fontFamily != "Null()")
                font = new Font(fontFamily, font.Size);

            if (fontSizeString != "Null()")
                font = new Font(font.FontFamily, fontSize);

            if (isBold == "True")
                font = new Font(font.FontFamily, font.Size, font.Style | FontStyle.Bold);

            if (isItalic == "True")
                font = new Font(font.FontFamily, font.Size, font.Style | FontStyle.Italic);

            if (isUnderline == "True")
                font = new Font(font.FontFamily, font.Size, font.Style | FontStyle.Underline);

            if (isStrikeout == "True")
                font = new Font(font.FontFamily, font.Size, font.Style | FontStyle.Strikeout);

            component.Font = font;

            if (component is IStiTextBrush comp)
                comp.TextBrush = new StiSolidBrush(color);
        }

        private int ToHi(string size)
        {
            if (size.Length > 0 && size[size.Length - 1] != ')')
            {
                if (size.Length > 3)
                {
                    var correctSize = size.Insert(size.Length - 3, ".");
                    size = correctSize;
                }
                else if (size.Length == 3)
                {
                    var correctSize = size.Insert(size.Length - 3, "0.");
                    size = correctSize;
                }
            }

            var result = (int)Math.Round(ParseDouble(size), 0);

            return result;
        }

        private void ClientRectangle(Hashtable obj, StiComponent component)
        {
            var x = ToHi(GetSetting(obj, "Object", "Position/Left"));
            var y = ToHi(GetSetting(obj, "Object", "Position/Top"));
            var width = ToHi(GetSetting(obj, "Object", "Position/Width"));
            var height = ToHi(GetSetting(obj, "Object", "Position/Height"));

            component.ClientRectangle = new Base.Drawing.RectangleD(x, y, width, height);
        }

        private void SetPageSettings(StiPage page)
        {
            page.Margins = new StiMargins(0);

            PageSettings(page);
            DefaultFont();
        }

        private void DefaultFont()
        {
            var fontString = GetSetting(keyPairs, "Layout", "DefFont");

            if (fontString != null)
            {
                var fontArray = fontString.Split(',');

                if (fontArray.Length > 4)
                {
                    var fontSize = (float)ParseDouble(fontArray[3].Replace("}", "").Replace("{", ""));
                    var fontFamily = fontArray[fontArray.Length - 1].Replace("}", "").Replace("{", "");

                    defaultFont = new Font(fontFamily, fontSize);
                }
            }
        }

        private void PageSettings(StiPage page)
        {
            foreach (var layout in pageLayoutsCollection)
            {
                var layoutType = layout["DisplayName"].ToString();

                if (layoutType == "Standard-Layout" || layoutType == "")
                {
                    var pageOrientation = layout["PaperFormat.Orientation"].ToString();
                    var pageFormat = layout["PaperFormat"].ToString();
                    var pageWidth = layout["PaperFormat.cx"].ToString();
                    var pageHeight = layout["PaperFormat.cy"].ToString();

                    page.Orientation = PageOrientation(pageOrientation);
                    page.PaperSize = PageFormat(pageFormat);

                    if (string.IsNullOrEmpty(pageWidth))
                        pageWidth = "209982"; //210mm A4

                    if (string.IsNullOrEmpty(pageHeight))
                        pageHeight = "297028"; //297mm A4

                    page.Width = ToHi(pageWidth);
                    page.Height = ToHi(pageHeight);

                    if (page.Orientation == StiPageOrientation.Landscape && page.Height > page.Width)
                    {
                        page.Width = ToHi(pageHeight);
                        page.Height = ToHi(pageWidth);
                    }
                }
            }
        }

        private StiPageOrientation PageOrientation(string pageOrientation)
        {
            var orientation = StiPageOrientation.Portrait;

            switch (pageOrientation)
            {
                case "1":
                    orientation = StiPageOrientation.Portrait;
                    break;

                case "2":
                    orientation = StiPageOrientation.Landscape;
                    break;

                default:
                    orientation = StiPageOrientation.Portrait;
                    break;
            }

            return orientation;
        }

        private System.Drawing.Printing.PaperKind PageFormat(string pageFormat)
        {
            var paperKind = System.Drawing.Printing.PaperKind.A4;

            switch (pageFormat)
            {
                case "66":
                    paperKind = System.Drawing.Printing.PaperKind.A2;
                    break;

                case "8":
                    paperKind = System.Drawing.Printing.PaperKind.A3;
                    break;

                case "9":
                    paperKind = System.Drawing.Printing.PaperKind.A4;
                    break;

                case "11":
                    paperKind = System.Drawing.Printing.PaperKind.A5;
                    break;

                default:
                    paperKind = System.Drawing.Printing.PaperKind.A4;
                    break;
            }

            return paperKind;
        }

        private double ParseDouble(string strValue)
        {
            var newStringValue = strValue.Trim().Replace(",", ".").Replace(".", ApplicationDecimalSeparator);
            var result = 0d;
            var success = double.TryParse(newStringValue, out result);

            if (!success)
            {
                var newStrings = newStringValue.Split('(');

                if (newStrings.Length == 2 && page != null)
                {
                    newStringValue = newStrings[1].Replace(")", "");

                    if (newStringValue.Length > 3)
                    {
                        var correctSize = newStringValue.Insert(newStringValue.Length - 3, ".");
                        newStringValue = correctSize;
                    }
                    else if (newStringValue.Length == 3)
                    {
                        var correctSize = newStringValue.Insert(newStringValue.Length - 3, "0.");
                        newStringValue = correctSize;
                    }

                    var superiorSizeItem = newStrings[0].Split('-')[0].Replace(" ", "");
                    var superiorValue = 0d;
                    var inferiorValue = 0d;

                    if (superiorSizeItem == "#LL.Device.Page.Size.cx")
                        superiorValue = page.ClientRectangle.Width;

                    if (superiorSizeItem == "#LL.Device.Page.Size.cy")
                        superiorValue = page.ClientRectangle.Height;

                    var succesInferiorValue = double.TryParse(newStringValue, out inferiorValue);

                    if (succesInferiorValue && superiorValue >= 0 && inferiorValue >= 0)
                        result = superiorValue - inferiorValue;
                }
            }

            return result;
        }

        private static string ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;
        #endregion

        #region Methods.Import
        public static StiImportResult Import(byte[] bytes)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = StiCultureInfo.GetEN(false);

                ApplicationDecimalSeparator = Thread.CurrentThread.CurrentCulture.NumberFormat.NumberDecimalSeparator;

                var resultString = System.Text.Encoding.UTF8.GetString(bytes)
                    .Replace("\0", "")
                    .Replace("\r\n", "\n")
                    .Replace("\r", "\n");

                var resultArray = resultString.Split('\n');

                var report = new StiReport();
                var errors = new List<string>();

                var helper = new StiListAndLabelHelper();
                helper.ProcessFile(resultArray, report);

                return new StiImportResult(report, errors);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }
        #endregion
    }
}
