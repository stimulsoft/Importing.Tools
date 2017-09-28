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
using System.IO;
using System.Text;
using System.Drawing;
using System.Collections;
using System.Globalization;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components;
using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;

namespace Stimulsoft.Report.Import
{
    public class StiRtfHelper
    {
        #region Variables
        private const double TwipsToHi	= 1 / 14.4;
		private const double defaultTextBoxHeight = 10;
		private const double defaultTextBoxWidth = 748;

        private const int codePagesTableSize = 14;
		private int[,] codePagesTable = null;
		private int mainCodePage = 1252;

		private Hashtable langToCodepage = null;
		private Hashtable codepageToEncodingTables = null;

        private int bandsCounter = 0;
        #endregion

        #region Structures
        private class StiRtfCell
		{
			public StiText Content = null;
			public bool FixedHeight = false;
			public bool Merged = false;
			public int MergedCellsCount = 0;
			public int MergedX = 0;
			public bool UsedOnLastPass = false;

            public StiRtfCell()
			{
				Content = new StiText();
				FixedHeight = false;
				Merged = false;
				MergedCellsCount = 0;
				MergedX = 0;
				UsedOnLastPass = false;
			}

			public StiRtfCell Clone()
			{
				StiRtfCell newCell = (StiRtfCell)this.MemberwiseClone();
				newCell.Content = (StiText)this.Content.Clone();
				return newCell;
			}
		}

		private struct StiRtfState
		{
			public int FontNumber;
            public int AFontNumber;
            public double FontSize;
			public bool Bold;
			public bool Italic;
			public bool Underline;
			public int FontColor;
			public char[] Encoding;
		}

		private class StiRtfFontData
		{
			public int Number;
			public int Charset;
			public string Name;
			public char[] Encoding;
		}

		private class StiRtfPageData
		{
			public int PageWidth;
			public int PageHeight;
			public int PageMarginLeft;
			public int PageMarginRight;
			public int PageMarginTop;
			public int PageMarginBottom;
			public bool PageLandscape;
			public int SectWidth;
			public int SectHeight;
			public int SectMarginLeft;
			public int SectMarginRight;
			public int SectMarginTop;
			public int SectMarginBottom;
			public bool SectLandscape;

			public StiRtfPageData()
			{
				PageWidth = 11906;
				PageHeight = 16838;
				PageMarginLeft = 567;
				PageMarginRight = 567;
				PageMarginTop = 567;
				PageMarginBottom = 567;
				PageLandscape = false;
				SectWidth = PageWidth;
				SectHeight = PageHeight;
				SectMarginLeft = PageMarginLeft;
				SectMarginRight = PageMarginRight;
				SectMarginTop = PageMarginTop;
				SectMarginBottom = PageMarginBottom;
				SectLandscape = false;
			}
        }
        #endregion

        #region Utils
        private double Convert(double twips)
		{
			return Math.Round(twips * TwipsToHi, 1);
        }

		private int GetCodepageFromLang(int lang)
		{
            if ((lang == 0x0000) || (lang == 0x007F) || (lang == 0x00FF) || (lang == 0x0400) || (lang == 0x0800))
			{
				return -1;
			}
			if (!langToCodepage.ContainsKey(lang))
			{
				CultureInfo ci = new CultureInfo(lang, false);
				int codePage = ci.TextInfo.ANSICodePage;
				langToCodepage[lang] = codePage;
			}
			return (int)langToCodepage[lang];
		}

		private char[] GetEncodingTable(int codepage)
		{
			if (!codepageToEncodingTables.ContainsKey(codepage))
			{
				char[] table = new char[256];
				Encoding enc = Encoding.GetEncoding(codepage);
				byte[] encodeBuf = new byte[256];
				for (int index = 0; index < 256; index++) encodeBuf[index] = (byte)(index);
				enc.GetChars(encodeBuf, 0, 256, table, 0);
				codepageToEncodingTables[codepage] = table;
			}
			return (char[])codepageToEncodingTables[codepage];
		}

		private StiRtfFontData GetFontData(int fontNumber, ArrayList fontTable)
		{
			StiRtfFontData fontData = null;
			foreach (StiRtfFontData data in fontTable)
			{
				if (data.Number == fontNumber)
				{
					fontData = data;
					break;
				}
			}
			if (fontData != null)
			{
				if (fontData.Encoding == null)
				{
					int codePage = mainCodePage;
					for (int index = 0; index < codePagesTableSize; index++)
					{
						if (codePagesTable[index, 2] == fontData.Charset) codePage = codePagesTable[index, 1];
					}
					fontData.Encoding = GetEncodingTable(codePage);
				}
			}
			return fontData;
		}

		private void UpdatePage(StiPage page, StiRtfPageData pageData)
		{
			page.PageWidth = Convert(pageData.SectWidth);
			page.PageHeight = Convert(pageData.SectHeight);
            page.Margins = new StiMargins(
                Convert(pageData.SectMarginLeft),
                Convert(pageData.SectMarginRight),
                Convert(pageData.SectMarginTop),
                Convert(pageData.SectMarginBottom));
			if (pageData.SectLandscape) page.Orientation = StiPageOrientation.Landscape;
        }
        #endregion

        #region Links
        //ms-help://MS.VSCC.2003/MS.MSDNQTR.2004APR.1033/intl/nls_2jzn.htm
		//ms-help://MS.VSCC.2003/MS.MSDNQTR.2004APR.1033/intl/nls_34rz.htm
		//ms-help://MS.VSCC.2003/MS.MSDNQTR.2004APR.1033/intl/nls_8rse.htm
        //ms-help://MS.VSCC.2003/MS.MSDNQTR.2004APR.1033/intl/unicode_81rn.htm
        #endregion

        #region Methods.Import
        public static StiImportResult Import(byte[] inputData)
        {
            var report = new StiRtfHelper().ImportInternal(inputData);

            return new StiImportResult(report, null);
        }

        private StiReport ImportInternal(byte[] inputData)
		{
            var report = new StiReport();

            try
            {
                bandsCounter = 1;

                StringBuilder sb = new StringBuilder();
                for (int index = 0; index < inputData.Length; index++)
                {
                    sb.Append((char)inputData[index]);
                }


                StiRtfTree tree = new StiRtfTree();
                tree.LoadRtfText(sb.ToString());
                ArrayList bandsList = tree.GetBandsData();

                #region Scan header
                mainCodePage = 1252;
                ArrayList fontTable = new ArrayList();
                ArrayList colorTable = new ArrayList();
                StiRtfNodeCollection main = tree.rootNode.ChildNodes[0].ChildNodes;
                int pos = 0;
                //bool afterGroup = false;
                while (true)
                {
                    StiRtfTreeNode currentNode = main[pos];
                    if (currentNode.NodeType == StiRtfNodeType.Group)
                    {
                        //afterGroup = true;
                        StiRtfTreeNode node = currentNode.ChildNodes[0];
                        if (node.NodeType == StiRtfNodeType.Keyword)
                        {
                            if (node.NodeKey == "fonttbl")
                            {
                                #region parse font table
                                node = node.NextToken;
                                while (node.NodeType == StiRtfNodeType.Group)
                                {
                                    StiRtfTreeNode nextNode = node.ChildNodes[0];
                                    StiRtfFontData fontData = new StiRtfFontData();
                                    while (nextNode.NodeType != StiRtfNodeType.GroupEnd)
                                    {
                                        if ((nextNode.NodeType == StiRtfNodeType.Keyword) && (nextNode.NodeKey == "f"))
                                        {
                                            fontData.Number = nextNode.Parameter;
                                        }
                                        if ((nextNode.NodeType == StiRtfNodeType.Keyword) && (nextNode.NodeKey == "fcharset"))
                                        {
                                            fontData.Charset = nextNode.Parameter;
                                        }
                                        if (nextNode.NodeType == StiRtfNodeType.Text)
                                        {
                                            string name = nextNode.NodeKey;
                                            if (name.EndsWith(";", StringComparison.InvariantCulture)) name = name.Substring(0, name.Length - 1);
                                            fontData.Name = name;
                                        }
                                        nextNode = nextNode.NextToken;
                                    }
                                    fontTable.Add(fontData);
                                    node = node.NextToken;
                                }
                                #endregion
                            }
                            if (node.NodeKey == "colortbl")
                            {
                                #region parse color table
                                node = node.NextToken;
                                int vRed = 0;
                                int vGreen = 0;
                                int vBlue = 0;
                                while (node.NodeType != StiRtfNodeType.GroupEnd)
                                {
                                    if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "red"))
                                    {
                                        vRed = node.Parameter;
                                    }
                                    if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "green"))
                                    {
                                        vGreen = node.Parameter;
                                    }
                                    if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "blue"))
                                    {
                                        vBlue = node.Parameter;
                                    }
                                    if ((node.NodeType == StiRtfNodeType.Text) && (node.NodeKey.IndexOf(';') != -1))
                                    {
                                        for (int indexKey = 0; indexKey < node.NodeKey.Length; indexKey++)
                                        {
                                            if (node.NodeKey[indexKey] == ';')
                                            {
                                                Color color = Color.FromArgb(vRed, vGreen, vBlue);
                                                colorTable.Add(color);
                                                vRed = 0;
                                                vGreen = 0;
                                                vBlue = 0;
                                            }
                                        }
                                    }
                                    node = node.NextToken;
                                }
                                #endregion
                            }
                        }
                    }
                    else
                    {
                        if (currentNode.NodeType == StiRtfNodeType.Keyword)
                        {
                            if (currentNode.NodeKey == "ansicpg")
                            {
                                mainCodePage = currentNode.Parameter;
                            }
                        }
                        //find next group
                        int pos2 = pos;
                        while ((pos2 < main.Count) && (main[pos2].NodeType != StiRtfNodeType.Group))
                        {
                            pos2++;
                        }
                        //check for header or info
                        if (pos2 < main.Count)
                        {
                            StiRtfTreeNode tempNode = main[pos2].ChildNodes[0];
                            if (((tempNode.NodeType == StiRtfNodeType.Keyword) &&
                                (StiRtfTree.TokensHash.ContainsKey(tempNode.NodeKey) && (int)StiRtfTree.TokensHash[tempNode.NodeKey] == 0)) ||
                                ((tempNode.NodeType == StiRtfNodeType.Control) && (tempNode.NodeKey == "*")))
                            {
                                pos++;
                                continue;
                            }
                        }
                        break;
                    }
                    pos++;
                }
                #endregion

                #region Prepare data
                langToCodepage = new Hashtable();
                codepageToEncodingTables = new Hashtable();

                codePagesTable = new int[codePagesTableSize, 3]
			    {
				    // langid, codepage, charset
				    {0x409, 1252,   1},
				    {0x419, 1251, 204},
				    {0x415, 1250, 238},
				    {0x408, 1253, 161},
				    {0x41F, 1254, 162},
				    {0x40D, 1255, 177},
				    {0x401, 1256, 178},
				    {0x425, 1257, 186},
				    {0x42A, 1258, 163},
				    {0x41E,  874, 222},
				    {0x411,  932, 128},
				    {0x804,  936, 134},
				    {0x412,  949, 129},
				    {0x404,  950, 136}
			    };

                report.Pages.Clear();
                report.Pages.Add(new StiPage(report));
                report.Unit = Stimulsoft.Report.Units.StiUnit.HundredthsOfInch;
                StiPage currentPage = report.Pages[0];
                StiRtfPageData pageData = new StiRtfPageData();

                Stack states = new Stack();
                StiRtfState currentState = new StiRtfState();
                StiRtfState beginState = new StiRtfState();
                currentState.Encoding = GetEncodingTable(mainCodePage);
                int skipGroupCount = 0;
                #endregion

                #region parse text
                for (int indexBand = 0; indexBand < bandsList.Count; indexBand++)
                {
                    StiRtfTree.StiRtfBandData data = (StiRtfTree.StiRtfBandData)bandsList[indexBand];
                    StiRtfTreeNode node = data.start;
                    if (!data.isTableRow)
                    {
                        #region plain text
                        StringBuilder res = new StringBuilder();
                        bool endLoop = false;
                        while (!endLoop)
                        {
                            if (res.Length == 0) beginState = currentState;
                            if (node.NodeType == StiRtfNodeType.Group)
                            {
                                bool needSkipGroup = false;
                                StiRtfTreeNode nextNode = node.NextFlateToken;
                                if ((nextNode.NodeType == StiRtfNodeType.Control) && (nextNode.NodeKey == "*"))
                                {
                                    needSkipGroup = true;
                                }
                                else
                                {
                                    if (nextNode.NodeType == StiRtfNodeType.Keyword)
                                    {
                                        if (nextNode.NodeKey == "field")
                                        {
                                            #region keyword "field"
                                            while (nextNode.NodeType != StiRtfNodeType.GroupEnd)
                                            {
                                                if ((nextNode.NodeType == StiRtfNodeType.Group) && (nextNode.NextFlateToken.NodeKey == "fldrslt"))
                                                {
                                                    nextNode = nextNode.NextToken;
                                                }
                                                else
                                                {
                                                    StiRtfTreeNode tempNode = nextNode;
                                                    nextNode = nextNode.NextToken;
                                                    tempNode.RemoveThisToken();
                                                }
                                            }
                                            #endregion
                                        }
                                        if (nextNode.NodeKey == "object")
                                        {
                                            needSkipGroup = true;
                                        }
                                        if (nextNode.NodeKey == "nonshppict")
                                        {
                                            needSkipGroup = true;
                                        }
                                        if (nextNode.NodeKey == "pict")
                                        {
                                            needSkipGroup = true;
                                        }
                                    }
                                }
                                if (needSkipGroup || (skipGroupCount > 0))
                                {
                                    skipGroupCount++;
                                }
                                states.Push(currentState);
                            }
                            if (node.NodeType == StiRtfNodeType.GroupEnd)
                            {
                                currentState = (StiRtfState)states.Pop();
                                if (skipGroupCount > 0) skipGroupCount--;
                            }

                            if (skipGroupCount == 0)
                            {
                                if (node.NodeType == StiRtfNodeType.Keyword)
                                {
                                    switch (node.NodeKey)
                                    {
                                        case "sect":
                                            currentPage = new StiPage(report);
                                            report.Pages.Add(currentPage);
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "page":
                                            currentPage = new StiPage(report);
                                            report.Pages.Add(currentPage);
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "sectd":
                                            pageData.SectWidth = pageData.PageWidth;
                                            pageData.SectHeight = pageData.PageHeight;
                                            pageData.SectMarginLeft = pageData.PageMarginLeft;
                                            pageData.SectMarginRight = pageData.PageMarginRight;
                                            pageData.SectMarginTop = pageData.PageMarginTop;
                                            pageData.SectMarginBottom = pageData.PageMarginBottom;
                                            pageData.SectLandscape = pageData.PageLandscape;
                                            UpdatePage(currentPage, pageData);
                                            break;

                                        case "paperw":
                                            pageData.PageWidth = node.Parameter;
                                            pageData.SectWidth = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "paperh":
                                            pageData.PageHeight = node.Parameter;
                                            pageData.SectHeight = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "margl":
                                            pageData.PageMarginLeft = node.Parameter;
                                            pageData.SectMarginLeft = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "margr":
                                            pageData.PageMarginRight = node.Parameter;
                                            pageData.SectMarginRight = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "margt":
                                            pageData.PageMarginTop = node.Parameter;
                                            pageData.SectMarginTop = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "margb":
                                            pageData.PageMarginBottom = node.Parameter;
                                            pageData.SectMarginBottom = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "landscape":
                                            pageData.PageLandscape = true;
                                            pageData.SectLandscape = true;
                                            UpdatePage(currentPage, pageData);
                                            break;

                                        case "pgwsxn":
                                            pageData.SectWidth = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "pghsxn":
                                            pageData.SectHeight = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "marglsxn":
                                            pageData.SectMarginLeft = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "margrsxn":
                                            pageData.SectMarginRight = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "margtsxn":
                                            pageData.SectMarginTop = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "margbsxn":
                                            pageData.SectMarginBottom = node.Parameter;
                                            UpdatePage(currentPage, pageData);
                                            break;
                                        case "lndscpsxn":
                                            pageData.SectLandscape = true;
                                            UpdatePage(currentPage, pageData);
                                            break;

                                        case "lang":
                                            int codepage = GetCodepageFromLang(node.Parameter);
                                            if (codepage >= 0) currentState.Encoding = GetEncodingTable(codepage);
                                            break;
                                        case "cf":
                                            currentState.FontColor = node.Parameter;
                                            break;
                                        case "f":
                                            currentState.FontNumber = node.Parameter;
                                            currentState.Encoding = GetFontData(currentState.FontNumber, fontTable).Encoding;
                                            break;
                                        case "af":
                                            currentState.AFontNumber = node.Parameter;
                                            if (currentState.FontNumber == 0)
                                            {
                                                currentState.Encoding = GetFontData(currentState.AFontNumber, fontTable).Encoding;
                                            }
                                            break;
                                        case "fs":
                                            currentState.FontSize = node.Parameter / 2f;
                                            break;
                                        case "b":
                                            currentState.Bold = !node.HasParameter;
                                            break;
                                        case "i":
                                            currentState.Italic = !node.HasParameter;
                                            break;
                                        case "ul":
                                            currentState.Underline = !node.HasParameter;
                                            break;
                                        case "par":
                                            res.Append("\r\n");
                                            break;
                                        case "line":
                                            res.Append("\n");
                                            break;
                                        case "tab":
                                            res.Append("\t");
                                            break;

                                        case "u":
                                            res.Append((char)node.Parameter);
                                            if (node.NextToken.NodeType == StiRtfNodeType.Text)
                                            {
                                                StiRtfTreeNode nextNode = node.NextToken;
                                                if (nextNode.NodeKey.StartsWith("?", StringComparison.InvariantCulture))
                                                {
                                                    nextNode.NodeKey = nextNode.NodeKey.Substring(1);
                                                }
                                            }
                                            break;
                                    }
                                }
                                if (node.NodeType == StiRtfNodeType.Text)
                                {
                                    //res.Append(node.NodeKey);
                                    for (int indexChar = 0; indexChar < node.NodeKey.Length; indexChar++)
                                    {
                                        res.Append(currentState.Encoding[node.NodeKey[indexChar]]);
                                    }
                                }
                                if ((node.NodeType == StiRtfNodeType.Control) && (node.HasParameter))
                                {
                                    res.Append(currentState.Encoding[node.Parameter]);
                                }
                            }
                            if (node == data.end) endLoop = true;
                            node = node.NextFlateToken;
                        }
                        if (res.Length > 0)
                        {
                            StiDataBand band = new StiDataBand();
                            band.Name = string.Format("DataBand{0}", bandsCounter++);
                            band.Height = defaultTextBoxHeight;
                            band.CountData = 1;
                            band.Top = defaultTextBoxHeight * 2 * indexBand;
                            band.CanBreak = true;
                            currentPage.Components.Add(band);

                            StiText text = new StiText();
                            text.Text = res.ToString();
                            //text.Width = defaultTextBoxWidth;
                            text.Width = Convert(pageData.SectWidth - pageData.SectMarginLeft - pageData.SectMarginRight);
                            text.Height = defaultTextBoxHeight;
                            FontStyle fs = FontStyle.Regular;
                            if (beginState.Bold) fs |= FontStyle.Bold;
                            if (beginState.Italic) fs |= FontStyle.Italic;
                            if (beginState.Underline) fs |= FontStyle.Underline;
                            text.Font = new Font(GetFontData(beginState.FontNumber, fontTable).Name, (float)beginState.FontSize, fs);
                            text.TextBrush = new StiSolidBrush((Color)colorTable[beginState.FontColor]);

                            text.Parent = band;
                            text.Page = currentPage;
                            text.CanGrow = true;
                            text.WordWrap = true;
                            SizeD textSize = text.GetActualSize();
                            if (textSize.Height > text.Height) text.Height = textSize.Height;
                            text.CanBreak = true;
                            band.Components.Add(text);
                            band.Height = text.Height;
                        }

                        #endregion
                    }
                    else
                    {
                        #region table rows
                        ArrayList previousCells = null;
                        ArrayList table = new ArrayList();
                        ArrayList mergedCells = null;

                        #region scan table row
                        while ((indexBand < bandsList.Count) && (((StiRtfTree.StiRtfBandData)bandsList[indexBand]).isTableRow))
                        {
                            data = (StiRtfTree.StiRtfBandData)bandsList[indexBand];
                            node = data.start;

                            if (mergedCells == null) mergedCells = new ArrayList();
                            if (data.isIntbl)
                            {
                                foreach (StiRtfCell cell in mergedCells)
                                {
                                    cell.MergedCellsCount++;
                                    cell.UsedOnLastPass = true;
                                }
                            }
                            ArrayList cells = new ArrayList();
                            StiRtfCell currentCell = new StiRtfCell();
                            if (previousCells != null)
                            {
                                for (int indexComp = 0; indexComp < previousCells.Count; indexComp++)
                                {
                                    cells.Add(((StiRtfCell)previousCells[indexComp]).Clone());
                                }
                                currentCell = (StiRtfCell)cells[0];
                            }
                            int cellIndex = 0;
                            double posX = 0;
                            StringBuilder res = new StringBuilder();
                            bool endLoop = false;
                            bool skipLoop = false;
                            bool cellContentPresent = false;
                            StiBorderSides borderSide = StiBorderSides.None;
                            double lineHeight = defaultTextBoxHeight;
                            while (!endLoop)
                            {
                                if (res.Length == 0) beginState = currentState;
                                if (node.NodeType == StiRtfNodeType.Group)
                                {
                                    bool needSkipGroup = false;
                                    StiRtfTreeNode nextNode = node.NextFlateToken;
                                    if ((nextNode.NodeType == StiRtfNodeType.Control) && (nextNode.NodeKey == "*"))
                                    {
                                        needSkipGroup = true;
                                    }
                                    else
                                    {
                                        if (nextNode.NodeType == StiRtfNodeType.Keyword)
                                        {
                                            if (nextNode.NodeKey == "field")
                                            {
                                                #region keyword "field"
                                                while (nextNode.NodeType != StiRtfNodeType.GroupEnd)
                                                {
                                                    if ((nextNode.NodeType == StiRtfNodeType.Group) && (nextNode.NextFlateToken.NodeKey == "fldrslt"))
                                                    {
                                                        nextNode = nextNode.NextToken;
                                                    }
                                                    else
                                                    {
                                                        StiRtfTreeNode tempNode = nextNode;
                                                        nextNode = nextNode.NextToken;
                                                        tempNode.RemoveThisToken();
                                                    }
                                                }
                                                #endregion
                                            }
                                            if (nextNode.NodeKey == "object")
                                            {
                                                needSkipGroup = true;
                                            }
                                            if (nextNode.NodeKey == "nonshppict")
                                            {
                                                needSkipGroup = true;
                                            }
                                            if (nextNode.NodeKey == "pict")
                                            {
                                                needSkipGroup = true;
                                            }
                                        }
                                    }
                                    if (needSkipGroup || (skipGroupCount > 0))
                                    {
                                        skipGroupCount++;
                                    }
                                    states.Push(currentState);
                                }
                                if (node.NodeType == StiRtfNodeType.GroupEnd)
                                {
                                    currentState = (StiRtfState)states.Pop();
                                    if (skipGroupCount > 0) skipGroupCount--;
                                }

                                if ((!skipLoop) && (skipGroupCount == 0))
                                {
                                    if (node.NodeType == StiRtfNodeType.Text)
                                    {
                                        //res.Append(node.NodeKey);
                                        for (int indexChar = 0; indexChar < node.NodeKey.Length; indexChar++)
                                        {
                                            res.Append(currentState.Encoding[node.NodeKey[indexChar]]);
                                        }
                                    }
                                    if (node.NodeType == StiRtfNodeType.Keyword)
                                    {
                                        switch (node.NodeKey)
                                        {
                                            #region //tbldef
                                            case "trowd":
                                                if (node != data.start)
                                                {
                                                    skipLoop = true;
                                                    break;
                                                }
                                                cells = new ArrayList();
                                                currentCell = new StiRtfCell();
                                                cellIndex = 0;
                                                break;

                                            case "trrh":
                                                lineHeight = Convert(node.Parameter);
                                                break;

                                            case "clvmgf":
                                                currentCell.UsedOnLastPass = true;
                                                mergedCells.Add(currentCell);
                                                break;
                                            case "clvmrg":
                                                currentCell.Merged = true;
                                                foreach (StiRtfCell cell in mergedCells)
                                                {
                                                    if (currentCell.MergedX == cell.MergedX)
                                                    {
                                                        cell.MergedCellsCount++;
                                                        cell.UsedOnLastPass = true;
                                                    }
                                                }
                                                break;

                                            case "clvertalt":
                                                currentCell.Content.VertAlignment = StiVertAlignment.Top;
                                                break;
                                            case "clvertalc":
                                                currentCell.Content.VertAlignment = StiVertAlignment.Center;
                                                break;
                                            case "clvertalb":
                                                currentCell.Content.VertAlignment = StiVertAlignment.Bottom;
                                                break;

                                            case "clbrdrt":
                                                borderSide = StiBorderSides.Top;
                                                break;
                                            case "clbrdrb":
                                                borderSide = StiBorderSides.Bottom;
                                                break;
                                            case "clbrdrl":
                                                borderSide = StiBorderSides.Left;
                                                break;
                                            case "clbrdrr":
                                                borderSide = StiBorderSides.Right;
                                                break;

                                            case "brdrnone":
                                                currentCell.Content.Border.Side |= borderSide;
                                                currentCell.Content.Border.Side ^= borderSide;
                                                break;
                                            case "brdrs":
                                                currentCell.Content.Border.Side |= borderSide;
                                                currentCell.Content.Border.Style = StiPenStyle.Solid;
                                                break;
                                            case "brdrdb":
                                                currentCell.Content.Border.Side |= borderSide;
                                                currentCell.Content.Border.Style = StiPenStyle.Double;
                                                break;
                                            case "brdrdot":
                                                currentCell.Content.Border.Side |= borderSide;
                                                currentCell.Content.Border.Style = StiPenStyle.Dot;
                                                break;
                                            case "brdrdash":
                                                currentCell.Content.Border.Side |= borderSide;
                                                currentCell.Content.Border.Style = StiPenStyle.Dash;
                                                break;

                                            case "brdrw":
                                                currentCell.Content.Border.Side |= borderSide;
                                                currentCell.Content.Border.Size = Convert(node.Parameter);
                                                break;

                                            case "brdrcf":
                                                currentCell.Content.Border.Side |= borderSide;
                                                currentCell.Content.Border.Color = (Color)colorTable[node.Parameter];
                                                break;

                                            case "clcbpat":
                                                currentCell.Content.Brush = new StiSolidBrush((Color)colorTable[node.Parameter]);
                                                break;

                                            case "cellx":
                                                double previousPosX = posX;
                                                posX = Convert(node.Parameter);
                                                currentCell.Content.Width = posX - previousPosX;
                                                currentCell.Content.Height = Math.Abs(lineHeight);
                                                if (lineHeight < 0) currentCell.FixedHeight = true;
                                                cells.Add(currentCell);
                                                currentCell = new StiRtfCell();
                                                currentCell.Content.Left = posX;
                                                currentCell.MergedX = node.Parameter;
                                                cellIndex++;
                                                break;
                                            #endregion

                                            #region //txtpar
                                            case "pard":
                                                //переключаемся на параметры текста
                                                if (!cellContentPresent)
                                                {
                                                    cellIndex = 0;
                                                    cellContentPresent = true;
                                                    currentCell = (StiRtfCell)cells[0];
                                                }
                                                break;

                                            case "ql":
                                                currentCell.Content.HorAlignment = StiTextHorAlignment.Left;
                                                break;
                                            case "qr":
                                                currentCell.Content.HorAlignment = StiTextHorAlignment.Right;
                                                break;
                                            case "qj":
                                                currentCell.Content.HorAlignment = StiTextHorAlignment.Width;
                                                break;
                                            case "qc":
                                                currentCell.Content.HorAlignment = StiTextHorAlignment.Center;
                                                break;

                                            case "lang":
                                                int codepage = GetCodepageFromLang(node.Parameter);
                                                if (codepage >= 0) currentState.Encoding = GetEncodingTable(codepage);
                                                break;
                                            case "cf":
                                                currentState.FontColor = node.Parameter;
                                                break;
                                            case "f":
                                                currentState.FontNumber = node.Parameter;
                                                currentState.Encoding = GetFontData(currentState.FontNumber, fontTable).Encoding;
                                                break;
                                            case "af":
                                                currentState.AFontNumber = node.Parameter;
                                                if (currentState.FontNumber == 0)
                                                {
                                                    currentState.Encoding = GetFontData(currentState.AFontNumber, fontTable).Encoding;
                                                }
                                                break;
                                            case "fs":
                                                currentState.FontSize = node.Parameter / 2f;
                                                break;
                                            case "b":
                                                currentState.Bold = !node.HasParameter;
                                                break;
                                            case "i":
                                                currentState.Italic = !node.HasParameter;
                                                break;
                                            case "ul":
                                                currentState.Underline = !node.HasParameter;
                                                break;

                                            case "par":
                                                res.Append("\r\n");
                                                break;
                                            case "line":
                                                res.Append("\n");
                                                break;
                                            case "tab":
                                                res.Append("\t");
                                                break;

                                            case "u":
                                                res.Append((char)node.Parameter);
                                                if (node.NextToken.NodeType == StiRtfNodeType.Text)
                                                {
                                                    StiRtfTreeNode nextNode = node.NextToken;
                                                    if (nextNode.NodeKey.StartsWith("?", StringComparison.InvariantCulture))
                                                    {
                                                        nextNode.NodeKey = nextNode.NodeKey.Substring(1);
                                                    }
                                                }
                                                break;

                                            case "cell":
                                                currentCell.Content.Text = res.ToString();
                                                FontStyle fs = FontStyle.Regular;
                                                if (beginState.Bold) fs |= FontStyle.Bold;
                                                if (beginState.Italic) fs |= FontStyle.Italic;
                                                if (beginState.Underline) fs |= FontStyle.Underline;
                                                var fontData = GetFontData(beginState.FontNumber, fontTable);
                                                if (fontData == null) fontData = GetFontData(beginState.AFontNumber, fontTable);
                                                currentCell.Content.Font = new Font(fontData.Name, (beginState.FontSize > 0 ? (float)beginState.FontSize : 2), fs);    //!!! fontSize
                                                currentCell.Content.TextBrush = new StiSolidBrush((Color)colorTable[beginState.FontColor]);
                                                res = new StringBuilder();
                                                cellIndex++;
                                                if (cellIndex < cells.Count)
                                                {
                                                    currentCell = (StiRtfCell)cells[cellIndex];
                                                }
                                                else
                                                {
                                                    currentCell = new StiRtfCell();
                                                }
                                                break;
                                            #endregion
                                        }
                                    }
                                    if (node.NodeType == StiRtfNodeType.Control)
                                    {
                                        if (node.HasParameter)
                                        {
                                            res.Append(currentState.Encoding[node.Parameter]);
                                        }
                                    }
                                }
                                if (node == data.end) endLoop = true;
                                node = node.NextFlateToken;
                            }
                            if ((cells != null) && (cells.Count > 0) && cellContentPresent)
                            {
                                table.Add(cells);
                                if (mergedCells.Count > 0)
                                {
                                    ArrayList tempMerged = new ArrayList();
                                    foreach (StiRtfCell cell in mergedCells)
                                    {
                                        if (cell.UsedOnLastPass) tempMerged.Add(cell);
                                        cell.UsedOnLastPass = false;
                                    }
                                    mergedCells = tempMerged;
                                }
                            }
                            previousCells = cells;
                            indexBand++;
                        }
                        indexBand--;
                        #endregion

                        #region create new band
                        StiDataBand band = new StiDataBand();
                        band.Name = string.Format("DataBand{0}", bandsCounter++);
                        band.CountData = 1;
                        band.Top = defaultTextBoxHeight * 2 * indexBand;
                        band.Brush = new StiSolidBrush(Color.FromArgb(32, Color.Green));
                        band.CanBreak = true;
                        currentPage.Components.Add(band);
                        #endregion

                        #region recalculate cells height
                        double[] linesHeight = new double[table.Count];
                        StiRtfCell[,] matrix = new StiRtfCell[table.Count, 64];
                        for (int indexLine = 0; indexLine < table.Count; indexLine++)
                        {
                            ArrayList currentLine = (ArrayList)table[indexLine];
                            for (int indexCell = 0; indexCell < currentLine.Count; indexCell++)
                            {
                                matrix[indexLine, indexCell] = (StiRtfCell)currentLine[indexCell];
                            }
                        }
                        //first pass
                        for (int indexLine = 0; indexLine < table.Count; indexLine++)
                        {
                            //calculate real height of each cell, then set line height
                            for (int indexCell = 0; indexCell < 64; indexCell++)
                            {
                                StiRtfCell cell = matrix[indexLine, indexCell];
                                if (cell != null)
                                {
                                    if (cell.FixedHeight)
                                    {
                                        linesHeight[indexLine] = cell.Content.Height;
                                        break;
                                    }
                                    cell.Content.Parent = band;
                                    cell.Content.Page = currentPage;
                                    cell.Content.CanGrow = true;
                                    cell.Content.WordWrap = true;
                                    if (cell.MergedCellsCount == 0)
                                    {
                                        SizeD cellSize = cell.Content.GetActualSize();
                                        if (cellSize.Height > cell.Content.Height) cell.Content.Height = cellSize.Height;
                                    }
                                    if (cell.Content.Height > linesHeight[indexLine]) linesHeight[indexLine] = cell.Content.Height;
                                }
                            }
                        }
                        #endregion

                        #region add components to band
                        double cellPosY = 0;
                        for (int indexLine = 0; indexLine < table.Count; indexLine++)
                        {
                            double cellPosX = 0;
                            for (int indexCell = 0; indexCell < 64; indexCell++)
                            {
                                StiRtfCell cell = matrix[indexLine, indexCell];
                                if (cell != null)
                                {
                                    //set real height for each cell in line
                                    if (cell.Content.Height < linesHeight[indexLine]) cell.Content.Height = linesHeight[indexLine];
                                    if (cell.MergedCellsCount > 0)
                                    {
                                        for (int indexCount = 0; indexCount < cell.MergedCellsCount; indexCount++)
                                        {
                                            cell.Content.Height += linesHeight[indexLine + indexCount + 1];
                                        }
                                        cell.Content.CanBreak = true;
                                    }

                                    if (!cell.Merged)
                                    {
                                        cell.Content.Left = cellPosX;
                                        cell.Content.Top = cellPosY;
                                        cell.Content.Parent = band;
                                        StiText text = cell.Content;
                                        band.Components.Add(text);
                                    }
                                    cellPosX += cell.Content.Width;
                                }
                            }
                            cellPosY += linesHeight[indexLine];
                        }
                        band.Height = cellPosY;
                        #endregion

                        #endregion
                    }
                }

                int pagesCounter = 1;
                foreach (StiPage page in report.Pages)
                {
                    page.LargeHeight = true;
                    page.LargeHeightFactor = 2;
                    page.Name = string.Format("Page{0}", pagesCounter++);
                }
                #endregion
            }
            finally
            {
                langToCodepage = null;
                codePagesTable = null;
                codepageToEncodingTables = null;
            }
			return report;
        }
        #endregion
    }

}
