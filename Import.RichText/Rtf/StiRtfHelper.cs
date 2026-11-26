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

#if NETCORE
using Font = Stimulsoft.Drawing.Font;
#endif

namespace Stimulsoft.Report.Import
{
    public class StiRtfHelper
    {
        #region Variables
        private const double TwipsToHi	= 1 / 14.4;
		private const double defaultTextBoxHeight = 10;
		private const double defaultTextBoxWidth = 748;
		private const double defaultFontSize = 12;
		private const string defaultFontName = "Arial";

        private const int codePagesTableSize = 14;
		private int[,] codePagesTable = null;
		private int mainCodePage = 1252;

		private Hashtable langToCodepage = null;
		private Hashtable codepageToEncodingTables = null;

        private int bandsCounter = 0;
        private int defaultFontNumber;
        //the count of the characters after "\u" which repeat it for the old readers and are skipped
        private int fallbackCount;
        //the first byte of a double byte character which waits for the second one
        private int leadByte = -1;
        //the depth of the groups inside the definition of a row of a nested table, 0 outside of it
        private int nestedLevel;
        //the text which is collected now has an expression of the report: the number of the page
        private bool textHasExpression;
        //the last field gave an expression, its result (the text which the document shows) is skipped
        private bool skipFieldResult;
        //the bands of the headers and of the footers by their pages
        private Hashtable pageHeaders = new Hashtable();
        private Hashtable pageFooters = new Hashtable();
        private StiRtfFontData defaultFontData = new StiRtfFontData { Name = defaultFontName };
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
			public int CodePage;
			public bool FontHasCharset;
			public int UnicodeSkip;
			public bool Hidden;
			}

		private class StiRtfPageBand
		{
			public StiBand Band;
			public string Key;
		}

		private class StiRtfFontData
		{
			public int Number;
			public int Charset;
			public string Name;
			public char[] Encoding;
			public int CodePage;
			public bool HasCharset;
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
            if ((lang == 0x0000) || (lang == 0x007F) || (lang == 0x00FF) || (lang == 0x0400) || (lang == 0x0800) || (lang == 0x0C00))
			{
				return -1;
			}
			if (!langToCodepage.ContainsKey(lang))
			{
				int codePage = -1;
				try
				{
					codePage = new CultureInfo(lang, false).TextInfo.ANSICodePage;
					if (codePage == 0) codePage = -1;
				}
				catch (ArgumentException)
				{
					//an unknown language keeps the current code page
				}
				langToCodepage[lang] = codePage;
			}
			return (int)langToCodepage[lang];
		}

		private char[] GetEncodingTable(int codepage)
		{
			if (!codepageToEncodingTables.ContainsKey(codepage))
			{
				char[] table = new char[256];
				Encoding enc;
				try
				{
					enc = Encoding.GetEncoding(codepage);
				}
				catch (Exception e) when (e is ArgumentException || e is NotSupportedException)
				{
					enc = Encoding.GetEncoding(1252);
				}
				byte[] encodeBuf = new byte[256];
				for (int index = 0; index < 256; index++) encodeBuf[index] = (byte)(index);
				enc.GetChars(encodeBuf, 0, 256, table, 0);
				codepageToEncodingTables[codepage] = table;
			}
			return (char[])codepageToEncodingTables[codepage];
		}

		private StiRtfFontData GetFontData(StiRtfState state, ArrayList fontTable)
		{
            int fontNumber = state.FontNumber;

            StiRtfFontData fontData = null;
			foreach (StiRtfFontData data in fontTable)
			{
				if (data.Number == fontNumber)
				{
					fontData = data;
					break;
				}
			}
            if (fontData == null)
            {
                fontNumber = state.AFontNumber;
                foreach (StiRtfFontData data in fontTable)
                {
                    if (data.Number == fontNumber)
                    {
                        fontData = data;
                        break;
                    }
                }
            }

            //a font which the table doesn't have is replaced: the first font of the table or the default one
            if (fontData == null) fontData = fontTable.Count > 0 ? (StiRtfFontData)fontTable[0] : defaultFontData;

            if (fontData.Encoding == null)
            {
                int codePage = mainCodePage;
                for (int index = 0; index < codePagesTableSize; index++)
                {
                    if (codePagesTable[index, 2] != fontData.Charset) continue;

                    codePage = codePagesTable[index, 1];
                    //the charsets 0 and 1 are "ANSI" and "default": they don't name a code page
                    fontData.HasCharset = fontData.Charset > 1;
                }
                fontData.CodePage = codePage;
                fontData.Encoding = GetEncodingTable(codePage);
            }
            return fontData;
		}

		private static Color GetColor(ArrayList colorTable, int index)
		{
			//a file without the color table, or a wrong index, gives the default color
			return index >= 0 && index < colorTable.Count ? (Color)colorTable[index] : Color.Black;
		}

		private static void ParseFontTable(StiRtfTreeNode node, ArrayList fontTable)
		{
			//a font is either a group "{\f0\froman Times;}" or a part of the plain sequence
			//"\f0\fswiss Helvetica;\f1\fnil Menlo;", as macOS writes it
			StiRtfFontData fontData = null;
			int level = 0;
			int skipLevel = 0;
			node = node.NextFlateToken;
			while (node != null && level >= 0)
			{
				if (node.NodeType == StiRtfNodeType.Group)
				{
					level++;
					StiRtfTreeNode firstNode = node.ChildNodes[0];
					bool ignorable = firstNode.NodeType == StiRtfNodeType.Control && firstNode.NodeKey == "*";
					if (ignorable || skipLevel > 0) skipLevel++;
				}
				else if (node.NodeType == StiRtfNodeType.GroupEnd)
				{
					level--;
					if (skipLevel > 0) skipLevel--;
				}
				else if (skipLevel == 0)
				{
					if (node.NodeType == StiRtfNodeType.Keyword && node.NodeKey == "f")
					{
						fontData = new StiRtfFontData();
						fontData.Number = node.Parameter;
						fontTable.Add(fontData);
					}
					else if (fontData != null && node.NodeType == StiRtfNodeType.Keyword && node.NodeKey == "fcharset")
					{
						fontData.Charset = node.Parameter;
					}
					else if (fontData != null && fontData.Name == null && node.NodeType == StiRtfNodeType.Text)
					{
						string name = node.NodeKey.Trim().TrimEnd(';').Trim();
						if (name.Length > 0) fontData.Name = name;
					}
				}
				node = node.NextFlateToken;
			}

			foreach (StiRtfFontData data in fontTable)
			{
				if (data.Name == null) data.Name = defaultFontName;
			}
		}

		private static bool HasCellDefinitionAhead(StiRtfTreeNode node, StiRtfTreeNode end)
		{
			//true while the definition of the row goes on: one more "cellx" stands before the content of the cells
			while (node != null && node != end)
			{
				node = node.NextFlateToken;
				if (node == null || node.NodeType != StiRtfNodeType.Keyword) continue;

				if (node.NodeKey == "cellx") return true;
				if (node.NodeKey == "cell" || node.NodeKey == "row" || node.NodeKey == "trowd") return false;
			}
			return false;
		}

		private void SetFont(ref StiRtfState state, ArrayList fontTable)
		{
			StiRtfFontData fontData = GetFontData(state, fontTable);
			state.Encoding = fontData.Encoding;
			state.CodePage = fontData.CodePage;
			state.FontHasCharset = fontData.HasCharset;
		}

		private static bool IsLeadByte(int codePage, int value)
		{
			switch (codePage)
			{
				case 932:
					return (value >= 0x81 && value <= 0x9F) || (value >= 0xE0 && value <= 0xFC);

				case 936:
				case 949:
				case 950:
					return value >= 0x81 && value <= 0xFE;

				default:
					return false;
			}
		}

		private bool IsSkippedDestination(string key)
		{
			if (key == "fldrslt" && skipFieldResult)
			{
				skipFieldResult = false;
				return true;
			}
			switch (key)
			{
				case "object":
				case "nonshppict":
				case "pict":
				//a footnote is not a part of the body
				case "footnote":
				//the copy of a nested table for the readers which don't know them
				case "nonesttables":
					return true;

				default:
					return StiRtfTree.TokensHash.ContainsKey(key) && (int)StiRtfTree.TokensHash[key] == 0;
			}
		}

		//A nested table is kept as the text of its cell: a tab between the cells, a new line after a row.
		private static void AppendNestedRowEnd(StringBuilder res, StiRtfState state)
		{
			if (state.Hidden) return;

			if (res.Length > 0 && res[res.Length - 1] == '\t') res.Length--;
			//the end of the row can come twice: as "nestrow" and as the end of its group
			if (res.Length > 0 && res[res.Length - 1] == '\n') return;
			res.Append("\r\n");
		}

		private static void AppendChar(StringBuilder res, char value, StiRtfState state)
		{
			if (!state.Hidden) res.Append(value);
		}

		private void AppendByte(StringBuilder res, int value, StiRtfState state)
		{
			if (value < 0 || value > 255) return;

			if (fallbackCount > 0)
			{
				fallbackCount--;
				return;
			}

			if (leadByte >= 0)
			{
				int firstByte = leadByte;
				leadByte = -1;
				if (value >= 0x40)
				{
					byte[] bytes = new byte[] { (byte)firstByte, (byte)value };
					if (!state.Hidden) res.Append(Encoding.GetEncoding(state.CodePage).GetChars(bytes));
					return;
				}
				//not a pair: both bytes are single characters
				AppendChar(res, state.Encoding[firstByte], state);
			}

			if (IsLeadByte(state.CodePage, value))
			{
				leadByte = value;
				return;
			}
			AppendChar(res, state.Encoding[value], state);
		}

		//The text of the document and the keywords of the character formatting,
		//the same for a paragraph and for a cell.
		private void ProcessTextToken(StiRtfTreeNode node, StringBuilder res, ref StiRtfState state,
			ArrayList fontTable)
		{
			if (node.NodeType == StiRtfNodeType.Text)
			{
				for (int index = 0; index < node.NodeKey.Length; index++)
				{
					AppendByte(res, node.NodeKey[index], state);
				}
				return;
			}

			if (node.NodeType == StiRtfNodeType.Control)
			{
				switch (node.NodeKey)
				{
					case "'":
						if (node.HasParameter) AppendByte(res, node.Parameter, state);
						break;
					case "~":
						AppendChar(res, (char)0xA0, state);
						break;
					case "_":
						AppendChar(res, '-', state);
						break;
					//a backslash at the end of a line is a paragraph
					case "\r":
					case "\n":
						if (!state.Hidden) res.Append("\r\n");
						break;
				}
				return;
			}

			if (node.NodeType != StiRtfNodeType.Keyword) return;

			//a keyword ends the characters which repeat "\u" and breaks a pair of bytes
			fallbackCount = 0;
			leadByte = -1;

			switch (node.NodeKey)
			{
				case "lang":
					//the language sets the code page only for a font which doesn't name it
					int codepage = GetCodepageFromLang(node.Parameter);
					if (codepage >= 0 && !state.FontHasCharset)
					{
						state.Encoding = GetEncodingTable(codepage);
						state.CodePage = codepage;
					}
					break;
				case "cf":
					state.FontColor = node.Parameter;
					break;
				case "f":
					state.FontNumber = node.Parameter;
					SetFont(ref state, fontTable);
					break;
				case "af":
					state.AFontNumber = node.Parameter;
					if (state.FontNumber == 0) SetFont(ref state, fontTable);
					break;
				case "fs":
					state.FontSize = node.Parameter / 2f;
					break;
				case "b":
					state.Bold = !node.HasParameter;
					break;
				case "i":
					state.Italic = !node.HasParameter;
					break;
				case "ul":
					state.Underline = !node.HasParameter;
					break;
				case "v":
					state.Hidden = !node.HasParameter || node.Parameter != 0;
					break;
				case "plain":
					state.Bold = false;
					state.Italic = false;
					state.Underline = false;
					state.Hidden = false;
					state.FontColor = 0;
					state.FontSize = defaultFontSize;
					state.FontNumber = defaultFontNumber;
					SetFont(ref state, fontTable);
					break;

				case "par":
					if (!state.Hidden) res.Append("\r\n");
					break;
				case "line":
					if (!state.Hidden) res.Append("\n");
					break;
				case "tab":
				case "nestcell":
					if (!state.Hidden) res.Append("\t");
					break;

				case "nestrow":
					AppendNestedRowEnd(res, state);
					break;

				case "field":
					//a field in the middle of a group; a field at the start of a group is read with the group
					string expression = GetFieldExpression(node);
					skipFieldResult = expression != null;
					if (expression != null && !state.Hidden)
					{
						res.Append(expression);
						textHasExpression = true;
					}
					break;
				case "chpgn":
					if (!state.Hidden)
					{
						res.Append("{PageNumber}");
						textHasExpression = true;
					}
					break;

				case "uc":
					state.UnicodeSkip = Math.Max(0, node.Parameter);
					break;
				case "u":
					AppendChar(res, (char)(node.Parameter & 0xFFFF), state);
					fallbackCount = state.UnicodeSkip;
					break;

				case "lquote":
					AppendChar(res, (char)0x2018, state);
					break;
				case "rquote":
					AppendChar(res, (char)0x2019, state);
					break;
				case "ldblquote":
					AppendChar(res, (char)0x201C, state);
					break;
				case "rdblquote":
					AppendChar(res, (char)0x201D, state);
					break;
				case "endash":
					AppendChar(res, (char)0x2013, state);
					break;
				case "emdash":
					AppendChar(res, (char)0x2014, state);
					break;
				case "bullet":
					AppendChar(res, (char)0x2022, state);
					break;
				case "emspace":
				case "enspace":
				case "qmspace":
					AppendChar(res, ' ', state);
					break;
			}
		}

		//A document can have several headers: for all the pages, for the right, the left and the first ones.
		//A page of the report has one, the most common kind wins.
		private static int GetPageBandPriority(string key)
		{
			switch (key)
			{
				case "header":
				case "footer":
					return 4;

				case "headerr":
				case "footerr":
					return 3;

				case "headerl":
				case "footerl":
					return 2;

				default:
					return 1;
			}
		}

		private StiBand AddPageBand(StiPage page, bool isHeader)
		{
			StiBand band;
			if (isHeader)
			{
				band = new StiPageHeaderBand();
				band.Name = string.Format("PageHeaderBand{0}", bandsCounter++);
				page.Components.Insert(0, band);
			}
			else
			{
				band = new StiPageFooterBand();
				band.Name = string.Format("PageFooterBand{0}", bandsCounter++);
				page.Components.Add(band);
			}
			band.Height = 0;
			band.Page = page;
			return band;
		}

		//Returns the band for the content of a header or a footer, null when the page has a better kind of it.
		private StiBand GetPageBand(StiPage page, string key)
		{
			bool isHeader = key.StartsWith("header", StringComparison.Ordinal);
			Hashtable bands = isHeader ? pageHeaders : pageFooters;
			StiRtfPageBand pageBand = (StiRtfPageBand)bands[page];

			if (pageBand != null && pageBand.Key != key)
			{
				if (GetPageBandPriority(pageBand.Key) >= GetPageBandPriority(key)) return null;

				page.Components.Remove(pageBand.Band);
				pageBand = null;
			}

			if (pageBand == null)
			{
				pageBand = new StiRtfPageBand();
				pageBand.Key = key;
				pageBand.Band = AddPageBand(page, isHeader);
				bands[page] = pageBand;
			}
			return pageBand.Band;
		}

		//The band which takes the next text or table: a new data band of the body,
		//or the header or footer of the page.
		private StiBand GetContentBand(StiPage page, string pageBandKey, int indexBand)
		{
			StiDataBand band = new StiDataBand();
			if (pageBandKey != null)
			{
				//the band of a dropped kind of header stays out of the page
				StiBand pageBand = GetPageBand(page, pageBandKey);
				return pageBand != null ? pageBand : band;
			}

			band.Name = string.Format("DataBand{0}", bandsCounter++);
			band.CountData = 1;
			band.Top = defaultTextBoxHeight * 2 * indexBand;
			band.CanBreak = true;
			page.Components.Add(band);
			return band;
		}

		//A page without its own header or footer takes them from the previous page: the header of a section
		//works for all its pages and for the next sections.
		private void InheritPageBand(Hashtable bands, StiPage previousPage, StiPage page, bool isHeader)
		{
			if (previousPage == null || bands[page] != null) return;

			StiRtfPageBand previous = (StiRtfPageBand)bands[previousPage];
			if (previous == null) return;

			StiRtfPageBand pageBand = new StiRtfPageBand();
			pageBand.Key = previous.Key;
			pageBand.Band = AddPageBand(page, isHeader);
			pageBand.Band.Height = previous.Band.Height;
			foreach (StiComponent component in previous.Band.Components)
			{
				StiComponent clone = (StiComponent)component.Clone();
				//the copy gets its own name when it is added to the band
				clone.Name = string.Empty;
				clone.Page = page;
				clone.Parent = pageBand.Band;
				pageBand.Band.Components.Add(clone);
			}
			bands[page] = pageBand;
		}

		private static bool IsSameTable(ArrayList bandsList, int index, string pageBand)
		{
			if (index >= bandsList.Count) return false;

			StiRtfTree.StiRtfBandData data = (StiRtfTree.StiRtfBandData)bandsList[index];
			return data.isTableRow && data.pageBand == pageBand;
		}

		//The fields with the number of the page become the system variables of the report.
		private static string GetFieldExpression(StiRtfTreeNode fieldNode)
		{
			StringBuilder instruction = new StringBuilder();
			StiRtfTreeNode node = fieldNode.NextToken;
			while (node != null && node.NodeType != StiRtfNodeType.GroupEnd)
			{
				if (node.NodeType == StiRtfNodeType.Group && node.ChildNodes.Count > 1 &&
					node.ChildNodes[1].NodeType == StiRtfNodeType.Keyword && node.ChildNodes[1].NodeKey == "fldinst")
				{
					StiRtfTreeNode last = node.LastChild;
					StiRtfTreeNode item = node.NextFlateToken;
					while (item != null && item != last)
					{
						if (item.NodeType == StiRtfNodeType.Text) instruction.Append(item.NodeKey);
						item = item.NextFlateToken;
					}
				}
				node = node.NextToken;
			}

			string[] words = instruction.ToString().Trim().Split(' ');
			switch (words[0].ToUpperInvariant())
			{
				case "PAGE":
					return "{PageNumber}";

				case "NUMPAGES":
				case "SECTIONPAGES":
					return "{TotalPageCount}";

				default:
					return null;
			}
		}

		private void SetText(StiText text, string value)
		{
			StringBuilder clean = new StringBuilder(value.Length);
			for (int index = 0; index < value.Length; index++)
			{
				char symbol = value[index];
				if (char.IsHighSurrogate(symbol) && index + 1 < value.Length && char.IsLowSurrogate(value[index + 1]))
				{
					clean.Append(symbol);
					clean.Append(value[index + 1]);
					index++;
				}
				else if (char.IsSurrogate(symbol))
				{
					//a half of a pair can't be saved to the report
					clean.Append((char)0xFFFD);
				}
				else if (symbol >= ' ' || symbol == '\t' || symbol == '\r' || symbol == '\n')
				{
					clean.Append(symbol);
				}
			}
			text.Text = clean.ToString();

			//a brace of the document is a text, not an expression of the report;
			//a text with the number of the page keeps its expression
			if (value.IndexOf('{') >= 0 && !textHasExpression) text.OnlyText = true;
			textHasExpression = false;
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

                //not an RTF document: there is nothing to import
                if (tree.rootNode.ChildNodes.Count == 0 || tree.rootNode.ChildNodes[0].NodeType != StiRtfNodeType.Group)
                {
                    return report;
                }

                ArrayList bandsList = tree.GetBandsData();

                #region Scan header
                mainCodePage = 1252;
                ArrayList fontTable = new ArrayList();
                ArrayList colorTable = new ArrayList();
                defaultFontNumber = 0;
                fallbackCount = 0;
                leadByte = -1;
                nestedLevel = 0;
                textHasExpression = false;
                skipFieldResult = false;
                pageHeaders.Clear();
                pageFooters.Clear();
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
                                ParseFontTable(node, fontTable);
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
                            if (currentNode.NodeKey == "deff")
                            {
                                defaultFontNumber = currentNode.Parameter;
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
                currentState.CodePage = mainCodePage;
                currentState.UnicodeSkip = 1;
                //the defaults of RTF for a text without "\f" and "\fs"
                currentState.FontNumber = defaultFontNumber;
                currentState.FontSize = defaultFontSize;
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
                                if (nestedLevel > 0) nestedLevel++;
                                if ((nextNode.NodeType == StiRtfNodeType.Control) && (nextNode.NodeKey == "*"))
                                {
                                    //the definition of a nested row ends the row; some writers
                                    //put the text of its cells inside, so it is read, not skipped
                                    bool nestedRow = StiRtfTree.IsNestedRowDefinition(node);
                                    if (nestedRow && nestedLevel == 0 && skipGroupCount == 0)
                                    {
                                        nestedLevel = 1;
                                    }
                                    else if (nestedLevel == 0)
                                    {
                                        needSkipGroup = true;
                                    }
                                }
                                else
                                {
                                    if (nextNode.NodeType == StiRtfNodeType.Keyword)
                                    {
                                        if (nextNode.NodeKey == "field")
                                        {
                                            #region keyword "field"
                                            string fieldExpression = null;
                                            if (skipGroupCount == 0) fieldExpression = GetFieldExpression(nextNode);
                                            skipFieldResult = false;
                                            while (nextNode.NodeType != StiRtfNodeType.GroupEnd)
                                            {
                                                bool isResult = (nextNode.NodeType == StiRtfNodeType.Group) &&
                                                    (nextNode.NextFlateToken.NodeKey == "fldrslt");
                                                //the result of the field is its text;
                                                //a known field gives an expression instead
                                                if (isResult && fieldExpression == null)
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
                                            if (fieldExpression != null && !currentState.Hidden)
                                            {
                                                res.Append(fieldExpression);
                                                textHasExpression = true;
                                            }
                                            #endregion
                                        }
                                        if (IsSkippedDestination(nextNode.NodeKey))
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
                                if (states.Count > 0) currentState = (StiRtfState)states.Pop();
                                if (skipGroupCount > 0) skipGroupCount--;
                                if (nestedLevel > 0)
                                {
                                    nestedLevel--;
                                    if (nestedLevel == 0) AppendNestedRowEnd(res, currentState);
                                }
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
                                    }
                                }
                                ProcessTextToken(node, res, ref currentState, fontTable);
                            }
                            if (node == data.end) endLoop = true;
                            node = node.NextFlateToken;
                            if (node == null) endLoop = true;
                        }
                        if (res.Length > 0)
                        {
                            StiBand band = GetContentBand(currentPage, data.pageBand, indexBand);
                            //the content of a header or of a footer is stacked in its band
                            double bandTop = band is StiDataBand ? 0 : band.Height;

                            StiText text = new StiText();
                            SetText(text, res.ToString());
                            //text.Width = defaultTextBoxWidth;
                            text.Width = Convert(pageData.SectWidth - pageData.SectMarginLeft - pageData.SectMarginRight);
                            text.Height = defaultTextBoxHeight;
                            FontStyle fs = FontStyle.Regular;
                            if (beginState.Bold) fs |= FontStyle.Bold;
                            if (beginState.Italic) fs |= FontStyle.Italic;
                            if (beginState.Underline) fs |= FontStyle.Underline;
                            float fontSize = (float)(beginState.FontSize > 0 ? beginState.FontSize : defaultFontSize);
                            text.Font = new Font(GetFontData(beginState, fontTable).Name, fontSize, fs);
                            text.TextBrush = new StiSolidBrush(GetColor(colorTable, beginState.FontColor));

                            text.Parent = band;
                            text.Page = currentPage;
                            text.CanGrow = true;
                            text.WordWrap = true;
                            SizeD textSize = text.GetActualSize();
                            if (textSize.Height > text.Height) text.Height = textSize.Height;
                            text.CanBreak = true;
                            text.Top = bandTop;
                            band.Components.Add(text);
                            band.Height = bandTop + text.Height;
                        }

                        #endregion
                    }
                    else
                    {
                        #region table rows
                        ArrayList previousCells = null;
                        //a table of a header doesn't join a table of the body which follows it
                        string tablePageBand = data.pageBand;
                        ArrayList table = new ArrayList();
                        ArrayList mergedCells = null;

                        #region scan table row
                        while (IsSameTable(bandsList, indexBand, tablePageBand))
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
                                    //the row takes the definition of the previous row, not its text
                                    StiRtfCell previousCell = ((StiRtfCell)previousCells[indexComp]).Clone();
                                    previousCell.Content.Text = string.Empty;
                                    cells.Add(previousCell);
                                }
                                if (cells.Count > 0) currentCell = (StiRtfCell)cells[0];
                            }
                            int cellIndex = 0;
                            double posX = 0;
                            StringBuilder res = new StringBuilder();
                            bool endLoop = false;
                            //a row without its own definition takes the cells of the previous row,
                            //its content starts at once
                            bool cellContentPresent = cells.Count > 0;
                            bool cellClosed = false;
                            //true inside a definition which stands in the row after its own one:
                            //a copy of the definition of this row, or the definition of a nested
                            //table. It is skipped up to its last "cellx"
                            bool ignoredDefinition = false;
                            //the cells which got their content before the definition of the row (SAS writes so)
                            ArrayList pendingCells = new ArrayList();
                            StiBorderSides borderSide = StiBorderSides.None;
                            double lineHeight = defaultTextBoxHeight;
                            while (!endLoop)
                            {
                                if (res.Length == 0) beginState = currentState;
                                if (node.NodeType == StiRtfNodeType.Group)
                                {
                                    bool needSkipGroup = false;
                                    StiRtfTreeNode nextNode = node.NextFlateToken;
                                    if (nestedLevel > 0) nestedLevel++;
                                    if ((nextNode.NodeType == StiRtfNodeType.Control) && (nextNode.NodeKey == "*"))
                                    {
                                        //the definition of a nested row ends the row; some writers
                                        //put the text of its cells inside, so it is read, not skipped
                                        bool nestedRow = StiRtfTree.IsNestedRowDefinition(node);
                                        if (nestedRow && nestedLevel == 0 && skipGroupCount == 0)
                                        {
                                            nestedLevel = 1;
                                        }
                                        else if (nestedLevel == 0)
                                        {
                                            needSkipGroup = true;
                                        }
                                    }
                                    else
                                    {
                                        if (nextNode.NodeType == StiRtfNodeType.Keyword)
                                        {
                                            if (nextNode.NodeKey == "field")
                                            {
                                                #region keyword "field"
                                                string fieldExpression = null;
                                                if (skipGroupCount == 0) fieldExpression = GetFieldExpression(nextNode);
                                                skipFieldResult = false;
                                                while (nextNode.NodeType != StiRtfNodeType.GroupEnd)
                                                {
                                                    bool isResult = (nextNode.NodeType == StiRtfNodeType.Group) &&
                                                        (nextNode.NextFlateToken.NodeKey == "fldrslt");
                                                    //the result of the field is its text;
                                                    //a known field gives an expression instead
                                                    if (isResult && fieldExpression == null)
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
                                                if (fieldExpression != null && !currentState.Hidden)
                                                {
                                                    res.Append(fieldExpression);
                                                    textHasExpression = true;
                                                }
                                                #endregion
                                            }
                                            if (IsSkippedDestination(nextNode.NodeKey))
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
                                    if (states.Count > 0) currentState = (StiRtfState)states.Pop();
                                    if (skipGroupCount > 0) skipGroupCount--;
                                    if (nestedLevel > 0)
                                    {
                                        nestedLevel--;
                                        if (nestedLevel == 0) AppendNestedRowEnd(res, currentState);
                                    }
                                }

                                if (skipGroupCount == 0)
                                {
                                    ProcessTextToken(node, res, ref currentState, fontTable);
                                    bool isRowKeyword = node.NodeType == StiRtfNodeType.Keyword && nestedLevel == 0;
                                    if (isRowKeyword)
                                    {
                                        if (node.NodeKey == "trowd" && node != data.start && cells.Count > 0)
                                        {
                                            ignoredDefinition = true;
                                        }
                                        isRowKeyword = !ignoredDefinition;
                                        if (ignoredDefinition && node.NodeKey == "cellx")
                                        {
                                            ignoredDefinition = HasCellDefinitionAhead(node, data.end);
                                        }
                                    }

                                    //the keywords of a nested definition belong to the nested table, not to this row
                                    if (isRowKeyword)
                                    {
                                        switch (node.NodeKey)
                                        {
                                            #region //tbldef
                                            case "trowd":
                                                //a definition in the middle of the row is read only when the row
                                                //had none yet: the content of a row, or a nested table at its
                                                //start, can stand before the definition
                                                cells = new ArrayList();
                                                currentCell = new StiRtfCell();
                                                cellIndex = 0;
                                                cellContentPresent = false;
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
                                                currentCell.Content.Border.Color = GetColor(colorTable, node.Parameter);
                                                break;

                                            case "clcbpat":
                                                if (node.Parameter >= 0 && node.Parameter < colorTable.Count)
                                                {
                                                    Color backColor = (Color)colorTable[node.Parameter];
                                                    currentCell.Content.Brush = new StiSolidBrush(backColor);
                                                }
                                                break;

                                            case "cellx":
                                                double previousPosX = posX;
                                                posX = Convert(node.Parameter);
                                                currentCell.Content.Width = posX - previousPosX;
                                                currentCell.Content.Height = Math.Abs(lineHeight);
                                                if (lineHeight < 0) currentCell.FixedHeight = true;
                                                if (cells.Count < pendingCells.Count)
                                                {
                                                    StiText content = ((StiRtfCell)pendingCells[cells.Count]).Content;
                                                    currentCell.Content.Text = content.Text;
                                                    currentCell.Content.OnlyText = content.OnlyText;
                                                    currentCell.Content.Font = content.Font;
                                                    currentCell.Content.TextBrush = content.TextBrush;
                                                    currentCell.Content.HorAlignment = content.HorAlignment;
                                                }
                                                cells.Add(currentCell);
                                                currentCell = new StiRtfCell();
                                                currentCell.Content.Left = posX;
                                                currentCell.MergedX = node.Parameter;
                                                cellIndex++;
                                                //the last "cellx" ends the definition of the row,
                                                //the content of the first cell follows; it doesn't have
                                                //to start with "pard", and "pard" can stand inside
                                                //the definition
                                                if (!cellContentPresent && !HasCellDefinitionAhead(node, data.end))
                                                {
                                                    cellContentPresent = true;
                                                    if (pendingCells.Count == 0)
                                                    {
                                                        cellIndex = 0;
                                                        currentCell = (StiRtfCell)cells[0];
                                                    }
                                                }
                                                break;
                                                #endregion

                                            #region //txtpar

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

                                            case "cell":
                                                cellClosed = true;
                                                SetText(currentCell.Content, res.ToString());
                                                FontStyle fs = FontStyle.Regular;
                                                if (beginState.Bold) fs |= FontStyle.Bold;
                                                if (beginState.Italic) fs |= FontStyle.Italic;
                                                if (beginState.Underline) fs |= FontStyle.Underline;
                                                var fontData = GetFontData(beginState, fontTable);
                                                currentCell.Content.Font = new Font(fontData.Name, (beginState.FontSize > 0 ? (float)beginState.FontSize : 2), fs);    //!!! fontSize
                                                Color textColor = GetColor(colorTable, beginState.FontColor);
                                                currentCell.Content.TextBrush = new StiSolidBrush(textColor);
                                                if (cellIndex >= cells.Count) pendingCells.Add(currentCell);
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
                                }
                                if (node == data.end) endLoop = true;
                                node = node.NextFlateToken;
                                if (node == null) endLoop = true;
                            }
                            //a tail of the table without any "cell" is not a row
                            if ((cells != null) && (cells.Count > 0) && cellContentPresent && cellClosed)
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
                        StiBand band = GetContentBand(currentPage, tablePageBand, indexBand);
                        if (band is StiDataBand) band.Brush = new StiSolidBrush(Color.FromArgb(32, Color.Green));
                        double bandTop = band is StiDataBand ? 0 : band.Height;
                        #endregion

                        #region recalculate cells height
                        double[] linesHeight = new double[table.Count];
                        int cellsCount = 0;
                        foreach (ArrayList line in table)
                        {
                            cellsCount = Math.Max(cellsCount, line.Count);
                        }
                        StiRtfCell[,] matrix = new StiRtfCell[table.Count, cellsCount];
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
                            for (int indexCell = 0; indexCell < cellsCount; indexCell++)
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
                        double cellPosY = bandTop;
                        for (int indexLine = 0; indexLine < table.Count; indexLine++)
                        {
                            double cellPosX = 0;
                            for (int indexCell = 0; indexCell < cellsCount; indexCell++)
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
                                            int mergedLine = indexLine + indexCount + 1;
                                            if (mergedLine < table.Count)
                                            {
                                                cell.Content.Height += linesHeight[mergedLine];
                                            }
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
                StiPage previousPage = null;
                foreach (StiPage page in report.Pages)
                {
                    InheritPageBand(pageHeaders, previousPage, page, true);
                    InheritPageBand(pageFooters, previousPage, page, false);
                    previousPage = page;
                    //page.LargeHeight = true;
                    //page.LargeHeightFactor = 2;
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
