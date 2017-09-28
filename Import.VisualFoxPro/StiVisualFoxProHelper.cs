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
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Printing;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components;
using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Xml;
using Stimulsoft.Report.Components.TextFormats;
using Import.VisualFoxPro;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Class helps converts Visual FoxPro templates to Stimulsoft Reports templates.
    /// </summary>
    public sealed class StiVisualFoxProHelper
	{
        #region Enums
        [Flags]
        private enum FFontStyle
        {
            Normal = 0,
            Bold = 1,
            Italic = 2,
            Underlined = 4,
            Strikethrough = 128
        }

        private enum ObjType
        {
            Comment = 0,
            Report = 1,
            Workarea = 2,   //(2.x reports)
            Index = 3,      //(2.x reports)
            Relation = 4,   //(2.x reports)
            Label = 5,
            Line = 6,
            RectangleOrShape = 7,
            Field = 8,
            Bandinfo = 9,
            Group = 10,
            PictureOrOle = 17,
            Variable = 18,
            PrinterDriverSetup = 21, //(2.x reports)
            Font = 23,
            DataEnvironment = 25,
            CursorRelationOrCursorAdapter = 26
        }
        #endregion

        #region Constants
        const double headerHeight = 2082.5;
        #endregion

        #region Variables
        private StiReport report = null;
        private DataRow dataRow = null;
        List<StiBand> bands = null;
        List<double> bandsTop = null;
        private List<string> methods = null;
        #endregion

        public List<StiWarning> warningList = new List<StiWarning>();
        private static string internalVersion = "2015.2 from 07-Oct-2015";

        #region Utils
        private static string FieldNameCorrection(string baseName)
        {
            StringBuilder sb = new StringBuilder(baseName);
            int charIndex = 0;
            for (int pos = 0; pos < baseName.Length; pos++)
            {
                if (baseName[pos] == ' ')
                {
                    sb.Remove(charIndex, 1);
                    if (charIndex < sb.Length && Char.IsLetter(sb[charIndex]))
                    {
                        sb[charIndex] = Char.ToUpper(sb[charIndex]);
                    }
                }
                else charIndex++;
            }
            for (int pos = 0; pos < sb.Length; pos++)
            {
                if (!(Char.IsLetterOrDigit(sb[pos]) || sb[pos] == '_' || sb[pos] == '.')) sb[pos] = '_';
            }
            baseName = sb.ToString();
            return baseName;
        }

        //private string CheckComponentsNames(string name)
        //{
        //    if (htNames.Contains(name))
        //    {
        //        int nameIndex = 1;
        //        while (htNames.Contains(string.Format("{0}_{1}", name, nameIndex)))
        //        {
        //            nameIndex++;
        //        }
        //        name = string.Format("{0}_{1}", name, nameIndex);
        //    }
        //    htNames[name] = name;
        //    return name;
        //}

        public static string ReplaceSymbols(string str)
        {
            StringBuilder sb = new StringBuilder(str);
            for (int pos = 0; pos < sb.Length; pos++)
            {
                //if ((!(Char.IsLetterOrDigit(sb[pos]) || sb[pos] == '_')) & (!(sb[pos] == '.'))) sb[pos] = '_';
                if (!(Char.IsLetterOrDigit(sb[pos]) || sb[pos] == '_')) sb[pos] = '_';
            }
            return sb.ToString();
        }


        private int GetInt(string columnName)
        {
            object obj = dataRow[columnName];
            if ((obj == null) || (obj == DBNull.Value)) return 0;
            return System.Convert.ToInt32(obj);
        }
        private double GetDouble(string columnName)
        {
            object obj = dataRow[columnName];
            if ((obj == null) || (obj == DBNull.Value)) return 0;
            var val = System.Convert.ToDouble(obj);
            return val;
        }
        private bool GetBool(string columnName)
        {
            object obj = dataRow[columnName];
            if ((obj == null) || (obj == DBNull.Value)) return false;
            return System.Convert.ToBoolean(obj);
        }
        private string GetString(string columnName)
        {
            object obj = dataRow[columnName];
            if ((obj == null) || (obj == DBNull.Value)) return null;
            return System.Convert.ToString(obj);
        }

        private double fru(double x)
        {
            if (GetString("PLATFORM").Trim().ToLowerInvariant() == "windows")
            {
                return x/100;
            }
            return x;
        }

        private Font GetFont()
        {
            FontStyle fs = FontStyle.Regular;
            int fontStyle = GetInt("FONTSTYLE");
            if ((fontStyle & (int)FFontStyle.Bold) > 0) fs |= FontStyle.Bold;
            if ((fontStyle & (int)FFontStyle.Italic) > 0) fs |= FontStyle.Italic;
            if ((fontStyle & (int)FFontStyle.Underlined) > 0) fs |= FontStyle.Underline;
            if ((fontStyle & (int)FFontStyle.Strikethrough) > 0) fs |= FontStyle.Strikeout;
            string fontFamily = GetString("FONTFACE");
            double fontSize = GetInt("FONTSIZE");
            if (!string.IsNullOrWhiteSpace(fontFamily) && (fontSize > 0))
                return new Font(fontFamily, (float)fontSize, fs);
            return new Font("Microsoft Sans Serif", 8);
        }

        private Color GetPenColor()
        {
            int r = GetInt("PENRED");
            int g = GetInt("PENGREEN");
            int b = GetInt("PENBLUE");
            Color penColor = Color.Black;
            if ((r != -1) && (g != -1) && (b != -1))
            {
                penColor = Color.FromArgb(r, g, b);
            }
            return penColor;
        }

        private Color GetFillColor()
        {
            int r = GetInt("FILLRED");
            int g = GetInt("FILLGREEN");
            int b = GetInt("FILLBLUE");
            Color fillColor = Color.White;
            if ((r != -1) && (g != -1) && (b != -1))
            {
                fillColor = Color.FromArgb(r, g, b);
            }
            int mode = GetInt("MODE");
            if (mode == 1)
            {
                fillColor = Color.Transparent;
            }
            return fillColor;
        }

        private StiPenStyle GetPenStyle()
        {
            int penPat = GetInt("PENPAT");
            StiPenStyle penStyle = StiPenStyle.None;
            switch (penPat)
            {
                case 0:
                    penStyle = StiPenStyle.None;
                    break;
                case 1:
                case 5:
                case 6:
                case 7:
                    penStyle = StiPenStyle.Dot;
                    break;
                case 2:
                    penStyle = StiPenStyle.Dash;
                    break;
                case 3:
                    penStyle = StiPenStyle.DashDot;
                    break;
                case 4:
                    penStyle = StiPenStyle.DashDotDot;
                    break;
                case 8:
                    penStyle = StiPenStyle.Solid;
                    break;
            }
            return penStyle;
        }

        private int GetPenSize()
        {
            return GetInt("PENSIZE");
        }

        private Pen GetPen()
        {
            Color penColor = GetPenColor();
            int penSize = GetPenSize();
            StiPenStyle penStyle = GetPenStyle();

            Pen pen = new Pen(penColor, penSize);
            pen.DashStyle = StiPenUtils.GetPenStyle(penStyle);

            return pen;
        }

        private StiBrush GetFillBrush()
        {
            Color fillColor = GetFillColor();
            int fillPat = GetInt("FILLPAT");

            //0 = None 1 = Solid 2 = Horizontal lines 3 = Vertical lines 4 = Diagonal lines, leaning left 5 = Diagonal lines, leaning right 6 = Grid (horizontal and vertical lines) 7 = Hatch (left and right diagonal lines)
            if (fillPat == 0)
            {
                return new StiEmptyBrush();
            }
            if (fillPat == 1)
            {
                return new StiSolidBrush(fillColor);
            }
            return new StiHatchBrush(HatchStyle.DiagonalCross, Color.Black, fillColor);
        }

        #endregion

        public StiReport Convert(DataTable visualFoxPro)
        {
            //log.OpenLog("Report");

            methods = new List<string>();

            report = new StiReport();
            report.MetaTags.Add("InternalVersion", internalVersion);
            try
            {
                #region Report Parameters
                //log.OpenNode("Report Parameters");

                report.ReportUnit = StiReportUnitType.HundredthsOfInch;

                //stimulReport.ReportAuthor =	crystalReport.SummaryInfo.ReportAuthor;
                //log.WriteNode(" Report Author: ", stimulReport.ReportAuthor);

                //stimulReport.ReportDescription = crystalReport.SummaryInfo.ReportComments;
                //log.WriteNode(" Report Description: ", stimulReport.ReportDescription);

                //stimulReport.ReportAlias = crystalReport.SummaryInfo.ReportTitle;
                //log.WriteNode(" Report Alias: ", stimulReport.ReportAlias);

                //log.CloseNode();
                #endregion

                report.Info.ShowHeaders = false;

                AddPage(report.Pages[0], visualFoxPro);

                //log.OpenNode("Warnings");
                //foreach (StiWarning warning in warningList)
                //{
                //    log.WriteNode(warning.Message);
                //}
                //log.CloseNode();
            }
            catch (Exception e)
            {
                StiExceptionProvider.Show(e);
            }

            report.Info.ShowHeaders = true;
            foreach (StiPage page in report.Pages)
            {
                page.DockToContainer();
                page.Correct();
            }

            foreach (StiComponent comp in report.GetComponents())
            {
                comp.Linked = false;
            }

            report.Dictionary.DataSources.Sort(StiSortOrder.Asc, true);

            #region Methods
            //if (methods.Count > 0)
            {
                int methodsPos = report.Script.IndexOf("#region StiReport");
                if (methodsPos > 0)
                {
                    methodsPos -= 9;

                    methods.AddRange(new string[] { 
                        "private string fnc_STR(double x, int y = 10, int z = 0) { return Math.Round(x, z).ToString().PadLeft(y); }",
                        "private bool fnc_EMPTY(string obj) { return string.IsNullOrEmpty(obj); }",
                        "private bool fnc_EMPTY(double obj) { return obj == 0; }",
                        " "
                    });

                    foreach (string st in methods)
                    {
                        string st2 = "\t\t" + st + "\r\n";
                        report.Script = report.Script.Insert(methodsPos, st2);
                        methodsPos += st2.Length;
                    }
                }
            }
            #endregion

            //log.CloseLog();

            return report;
        }


        #region Adds
        private void AddPage(StiPage page, DataTable dataTable)
        {
            #region Page Parameters
            page.TitleBeforeHeader = true;
            page.Orientation = StiPageOrientation.Landscape;

            //log.OpenNode("Page Parameters");

            //log.CloseNode();
            #endregion

            #region Components
            //log.OpenNode("Page Components");

            #region First pass
            bands = new List<StiBand>();
            bandsTop = new List<double>();

            foreach (DataRow dr in dataTable.Rows)
            {
                dataRow = dr;
                ObjType objType = (ObjType)GetInt("OBJTYPE");
                int objCode = GetInt("OBJCODE");
                double height = GetDouble("HEIGHT");
                double width = GetDouble("WIDTH");

                switch (objType)
                {
                    case ObjType.Bandinfo:
                        var band = TypeBand(objCode);
                        band.Height = fru(height);
                        page.Components.Add(band);
                        bands.Add(band);
                        bandsTop.Add(((bandsTop.Count == 0) ? 0 : bandsTop[bandsTop.Count - 1]) + height + headerHeight);
                        break;

                    case ObjType.Variable:
                        var variable = new StiVariable();
                        variable.Name = GetString("NAME");
                        variable.Alias = variable.Name;
                        variable.InitBy = StiVariableInitBy.Expression;
                        variable.Value = GetString("EXPR");
                        variable.Type = typeof(object);
                        variable.RequestFromUser = true;
                        page.Report.Dictionary.Variables.Add(variable);
                        break;

                }
            }
            #endregion

            #region Second pass
            foreach (DataRow dr in dataTable.Rows)
            {
                dataRow = dr;
                ObjType objType = (ObjType)GetInt("OBJTYPE");
                int objCode = GetInt("OBJCODE");
                double height = GetDouble("HEIGHT");
                double width = GetDouble("WIDTH");

                switch (objType)
                {
                    case ObjType.Label:
                        var stiText5 = new StiText();
                        AddToBandAndGetCommonProperties(stiText5);
                        stiText5.TextBrush = new StiSolidBrush(GetPenColor());
                        stiText5.Brush = GetFillBrush();
                        string stiText5Expr = GetString("EXPR");
                        stiText5.Text = stiText5Expr.Substring(1, stiText5Expr.Length - 2);
                        stiText5.Font = GetFont();
                        break;

                    case ObjType.Field:
                        var stiText8 = new StiText();
                        AddToBandAndGetCommonProperties(stiText8);
                        stiText8.TextBrush = new StiSolidBrush(GetPenColor());
                        stiText8.Brush = GetFillBrush();
                        stiText8.Text = ConvertExpression("{" + GetString("EXPR") + "}", stiText8);
                        stiText8.Font = GetFont();

                        #region TextFormat
                        string fillChar = GetString("FILLCHAR");
                        if (fillChar == "N")
                        {
                            StiNumberFormatService format = new StiNumberFormatService();
                            string picture = GetString("PICTURE");
                            if (!string.IsNullOrEmpty(picture))
                            {
                                int pos = picture.IndexOf(".");
                                if (pos == -1)
                                {
                                    format.DecimalDigits = 0;
                                }
                                else
                                {
                                    format.DecimalDigits = picture.Substring(pos + 1).Length;
                                }
                            }
                            stiText8.TextFormat = format;
                        }
                        if (fillChar == "D")
                        {
                            StiDateFormatService format = new StiDateFormatService();
                            stiText8.TextFormat = format;
                        }
                        #endregion

                        #region Alignment
                        int alignment = GetInt("OFFSET");
                        if (alignment == 1) stiText8.HorAlignment = StiTextHorAlignment.Right;
                        if (alignment == 2) stiText8.HorAlignment = StiTextHorAlignment.Center;
                        #endregion
                        break;

                    case ObjType.Line:
                        StiLinePrimitive line = null;
                        if (width > height)
                        {
                            line = new StiHorizontalLinePrimitive();
                        }
                        else
                        {
                            line = new StiVerticalLinePrimitive();
                        }
                        AddToBandAndGetCommonProperties(line);
                        line.Color = GetPenColor();
                        line.Style = GetPenStyle();
                        line.Size = GetPenSize();
                        break;

                    case ObjType.RectangleOrShape:
                        //проверять - если есть заливка - то надо Shape
                        StiRectanglePrimitive rectangle = null;
                        int offset = GetInt("OFFSET");
                        if (offset == 0)
                        {
                            rectangle = new StiRectanglePrimitive();
                        }
                        else
                        {
                            rectangle = new StiRoundedRectanglePrimitive();
                            (rectangle as StiRoundedRectanglePrimitive).Round = offset;
                        }
                        AddToBandAndGetCommonProperties(rectangle);
                        rectangle.Color = GetPenColor();
                        rectangle.Style = GetPenStyle();
                        rectangle.Size = GetPenSize();
                        break;

                    case ObjType.PictureOrOle:
                        var image = new StiImage();
                        AddToBandAndGetCommonProperties(image);
                        int sourceType = GetInt("OFFSET");
                        if (sourceType == 0)
                        {
                            image.ImageURL.Value = GetString("PICTURE").Trim();
                        }
                        if (sourceType == 1)
                        {
                            image.DataColumn = ConvertExpression("{" + GetString("NAME").Trim() + "}", image);
                        }
                        if (sourceType == 2)
                        {
                            string imageName = GetString("NAME").Trim();
                            if (imageName.StartsWith("(") && imageName.EndsWith(")")) imageName = imageName.Substring(1, imageName.Length - 2);
                            image.ImageData = new StiImageDataExpression(ConvertExpression("{" + imageName + "}", image));
                        }
                        break;

                    case ObjType.Variable:
                        var variable = new StiVariable();
                        variable.Name = GetString("NAME");
                        variable.Alias = variable.Name;
                        variable.InitBy = StiVariableInitBy.Expression;
                        string varValue = ConvertExpression("{" + GetString("EXPR") + "}", null);
                        variable.Value = varValue.Substring(1, varValue.Length - 2);
                        variable.Type = typeof(object);
                        variable.RequestFromUser = true;
                        page.Report.Dictionary.Variables.Add(variable);
                        break;

                    case ObjType.Bandinfo:
                    case ObjType.Font:
                        //none, skip
                        break;


                    default:
                        //log.WriteNode(string.Format("Unsupported OBJTYPE={0}", objType));
                        break;
                }
            }
            #endregion

            //log.CloseNode();
            #endregion

            page.DockToContainer();
            page.Correct();
        }

        private void AddToBandAndGetCommonProperties(StiComponent comp)
        {
            double height = GetDouble("HEIGHT");
            double width = GetDouble("WIDTH");
            double hpos = GetDouble("HPOS");
            double vpos = GetDouble("VPOS");

            if ((comp is StiVerticalLinePrimitive) || (comp is StiRectanglePrimitive))
            {
                StiStartPointPrimitive start = new StiStartPointPrimitive();
                start.ReferenceToGuid = comp.Guid;
                AddToBand(start, hpos, vpos, 0, 0);

                StiEndPointPrimitive end = new StiEndPointPrimitive();
                end.ReferenceToGuid = comp.Guid;
                AddToBand(end, hpos + width, vpos + height, 0, 0);

                bands[0].Page.Components.Add(comp);
            }
            else
            {
                AddToBand(comp, hpos, vpos, width, height);
            }

            if (GetString("PLATFORM").Trim().ToLowerInvariant() != "windows")
            {
                comp.Enabled = false;
            }

        }

        private void AddToBand(StiComponent comp, double hpos, double vpos, double width, double height)
        {
            int bandIndex = 0;
            double bandPos = 0;
            while (vpos >= bandsTop[bandIndex] - 10)
            {
                bandPos = bandsTop[bandIndex];
                bandIndex++;
            }
            vpos -= bandPos;
            if (vpos < 10) vpos = 0;

            comp.ClientRectangle = new RectangleD(fru(hpos), fru(vpos), fru(width), fru(height));
            bands[bandIndex].Components.Add(comp);
            comp.Linked = true;
        }

        private StiBand TypeBand(int objCode)
        {
            switch (objCode)
            {
                case 0: return new StiReportTitleBand();
                case 1: return new StiPageHeaderBand();
                case 2: return new StiColumnHeaderBand();
                case 3: return new StiGroupHeaderBand();
                case 4: return new StiDataBand();
                case 5: return new StiGroupFooterBand();
                case 6: return new StiColumnFooterBand();
                case 7: return new StiPageFooterBand();
                case 8: return new StiReportSummaryBand();
            }
            return new StiHeaderBand();
        }

        #endregion

        #region ExpressionConverter
        private string ConvertExpression(string input, StiComponent comp)
        {
            if (comp == null)
            {
                comp = new StiText();
                comp.Name = "*TextBox*";
                comp.Page = report.Pages[0];
            }

            var storeToPrint = false;
            //object aa = StiVisualFoxProParser.ParseTextValue(input, comp, comp, ref storeToPrint, false, true);
            object aa = StiVisualFoxProParser.ParseTextValue(input, comp, comp, ref storeToPrint, false);
            if (aa != null)
            {
                return aa.ToString();
            }

            return input;
        }
        #endregion


        #region Methods.Import
        public static StiImportResult Import(byte[] dataBytes, byte[] memoBytes)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US", false);

                DataTable visualFoxPro;
                using (var dataStream = new MemoryStream(dataBytes))
                {
                    if (memoBytes != null)
                    {
                        using (var memoStream = new MemoryStream(memoBytes))
                        {
                            visualFoxPro = StiDBaseHelper.GetTable(dataStream, memoStream, false, 0);
                        }
                    }
                    else
                    {
                        visualFoxPro = StiDBaseHelper.GetTable(dataStream, null, false, 0);
                    }
                }

                var convertor = new StiVisualFoxProHelper();
                var report = convertor.Convert(visualFoxPro);

                return new StiImportResult(report, null);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }
        #endregion

    }
}
