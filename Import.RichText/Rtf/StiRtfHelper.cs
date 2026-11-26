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
using System.Drawing.Imaging;
using System.Collections;
using System.Globalization;
using Stimulsoft.Report;
using Stimulsoft.Report.Dictionary;
using Stimulsoft.Report.Components;
using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;

#if NETCORE
using Bitmap = Stimulsoft.Drawing.Bitmap;
using Font = Stimulsoft.Drawing.Font;
using ImageFormat = Stimulsoft.Drawing.Imaging.ImageFormat;
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
		//the names of one family for the scripts: "Arial CYR", "Times New Roman CE", "Arial (Hebrew)"
		private static readonly string[] FontNameScripts = { " CYR", " CE", " Greek", " Tur", " Baltic",
			" (Hebrew)", " (Arabic)", " (Vietnamese)", " (Thai)" };
		private static readonly char[] FontNameSeparators = { '-', ' ' };
		//the indents of a paragraph are dropped when they leave less than this width for the text
		private const double MinTextWidth = 50;
		private const float MinLineSpacing = 0.5f;
		private const float MaxLineSpacing = 5f;
		private static Hashtable fontNames;

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
        //the picture of "shppict" was read, so its copy in "nonshppict" is not needed
        private bool pictureCopyIsExtra;
        //the text of the shape was read from "shpinst", so its copy in "shprslt" is not needed
        private bool shapeResultIsExtra;
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
			//the pictures of the cell, null when it has none
			public ArrayList Pictures = null;
			//the cell continues the cell on its left
			public bool JoinedToLeft = false;
			//the sides of the border: each of them has its own line
			public StiAdvancedBorder Sides = new StiAdvancedBorder();
			//the offset and the alignment of the row, its first cell keeps them
			public double RowLeft;
			public StiTextHorAlignment RowAlignment;

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
				newCell.Sides = (StiAdvancedBorder)this.Sides.Clone();
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
			public StiTextHorAlignment Alignment;
			public bool SymbolFont;
			public int UnicodeSkip;
			public bool Hidden;
			//the paragraph: its indents and the space around it in twips, the height of a line
			public int LeftIndent;
			public int RightIndent;
			public int FirstIndent;
			public int SpaceBefore;
			public int SpaceAfter;
			public int LineHeight;
			public bool LineHeightIsMultiple;
			}

		private class StiRtfFormatCount
		{
			public StiRtfState State;
			public int Count;
		}

		//A part of a text with one formatting of its characters.
		private class StiRtfTextRun
		{
			public StiRtfState State;
			public int Start;
			public int Length;
		}

		private class StiRtfPicture
		{
			public byte[] Data;
			public double Width;
			public double Height;
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
			public bool Bold;
			public bool Italic;
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

		//The families which Windows doesn't have under the name of the file: the PostScript names of macOS,
		//the fonts of old Windows and their metric twins of Linux. GDI replaces a part of them by itself,
		//a report doesn't, and a text of a missing font is shown with the fallback one.
		private static Hashtable GetFontNames()
		{
			if (fontNames != null) return fontNames;

			Hashtable names = new Hashtable();
			string[] arial = { "arial", "arialmt", "helvetica", "helveticaneue", "helv", "albanyamt",
				"liberationsans", "nimbussansl" };
			string[] times = { "timesnewroman", "timesnewromanpsmt", "timesnewromanps", "times", "timesroman",
				"tmsrmn", "cgtimes", "msserif", "liberationserif", "thorndaleamt", "nimbusromanno9l" };
			string[] courier = { "couriernew", "couriernewpsmt", "couriernewps", "courier", "fixedsys",
				"liberationmono", "cumberlandamt", "nimbusmonol" };

			foreach (string name in arial)
			{
				names[name] = "Arial";
			}

			foreach (string name in times)
			{
				names[name] = "Times New Roman";
			}

			foreach (string name in courier)
			{
				names[name] = "Courier New";
			}

			names["menlo"] = "Consolas";
			names["monaco"] = "Consolas";
			names["calibri"] = "Calibri";
			names["verdana"] = "Verdana";
			names["tahoma"] = "Tahoma";
			names["arialunicodems"] = "Arial Unicode MS";
			names["trebuchetms"] = "Trebuchet MS";
			names["lucidagrande"] = "Lucida Sans Unicode";
			names["mssansserif"] = "Microsoft Sans Serif";
			names["palatino"] = "Palatino Linotype";
			names["centuryschoolbook"] = "Century Schoolbook";
			names["newcenturyschlbk"] = "Century Schoolbook";
			names["avantgarde"] = "Century Gothic";

			fontNames = names;
			return fontNames;
		}

		private static string GetKnownFontName(string name)
		{
			return (string)GetFontNames()[name.Replace(" ", string.Empty).ToLowerInvariant()];
		}

		//The style which stands in the name of a font: "Bold", "BoldItalicMT", "Oblique", "Roman".
		private static bool TryGetFontNameStyle(string word, ref bool bold, ref bool italic)
		{
			word = word.ToLowerInvariant();
			if (word.Length > 2 && word.EndsWith("mt", StringComparison.Ordinal)) word = word.Substring(0, word.Length - 2);

			switch (word)
			{
				case "regular":
				case "roman":
				case "light":
				case "medium":
					return true;

				case "bold":
					bold = true;
					return true;

				case "italic":
				case "oblique":
					italic = true;
					return true;

				case "bolditalic":
				case "boldoblique":
					bold = true;
					italic = true;
					return true;

				default:
					return false;
			}
		}

		//"Arial-BoldMT" is Arial with the bold style, "Times New Roman CYR" is Times New Roman,
		//"Helvetica" is Arial. A name which is not known stays as it is.
		private static void SetFontName(StiRtfFontData fontData, string name)
		{
			foreach (string script in FontNameScripts)
			{
				if (name.Length > script.Length && name.EndsWith(script, StringComparison.OrdinalIgnoreCase))
				{
					name = name.Substring(0, name.Length - script.Length).Trim();
				}
			}

			string family = name;
			string knownName = GetKnownFontName(family);
			bool bold = false;
			bool italic = false;
			bool postScript = false;

			//the words of the style stand at the end: "Helvetica-BoldOblique", "Arial Bold Italic"
			while (knownName == null)
			{
				int pos = family.LastIndexOfAny(FontNameSeparators);
				if (pos <= 0 || !TryGetFontNameStyle(family.Substring(pos + 1), ref bold, ref italic)) break;

				if (family[pos] == '-') postScript = true;
				family = family.Substring(0, pos).Trim();
				knownName = GetKnownFontName(family);
			}

			//"Arial Rounded MT Bold" is a family of its own, its last word is not a style
			if (knownName == null && !postScript)
			{
				fontData.Name = name;
				return;
			}

			fontData.Name = knownName ?? family;
			fontData.Bold = bold;
			fontData.Italic = italic;
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
						if (name.Length > 0) SetFontName(fontData, name);
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
			//the charset 2 is a font of symbols: Symbol, Wingdings
			//an old file gives the Symbol font without a charset
			state.SymbolFont = fontData.Charset == 2 || fontData.Name == "Symbol";
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

			if (key == "shprslt")
			{
				bool isExtra = shapeResultIsExtra;
				shapeResultIsExtra = false;
				return isExtra;
			}

			//Word can leave "shppict" without the data and keep the picture only in the copy
			if (key == "nonshppict")
			{
				bool isExtra = pictureCopyIsExtra;
				pictureCopyIsExtra = false;
				return isExtra;
			}
			switch (key)
			{
				case "object":
				case "pict":
				//a property of a shape: its name and value are not a text of the document
				case "sp":
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

		//The markers of the lists are characters of the fonts of symbols. A text component has one font
		//for the whole text, so the usual markers are replaced with the same signs of Unicode. The value
		//is the code of the character in the font: a byte, or the same byte in the private area U+F0xx.
		private static char GetSymbol(int value)
		{
			switch (value & 0xFF)
			{
				case 0xB7:
					return (char)0x2022;

				case 0xA7:
					return (char)0x25AA;

				case 0xD8:
					return (char)0x27A2;

				case 0xFC:
					return (char)0x2713;

				default:
					return (char)0;
			}
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

			char symbol = state.SymbolFont ? GetSymbol(value) : (char)0;
			//a file of Macintosh keeps the code of the marker in its own code page
			char encoded = state.Encoding[value];
			if (state.SymbolFont && symbol == 0 && encoded != value && encoded <= 0xFF) symbol = GetSymbol(encoded);
			AppendChar(res, symbol != 0 ? symbol : encoded, state);
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

				case "pard":
					state.Alignment = StiTextHorAlignment.Left;
					state.LeftIndent = 0;
					state.RightIndent = 0;
					state.FirstIndent = 0;
					state.SpaceBefore = 0;
					state.SpaceAfter = 0;
					state.LineHeight = 0;
					state.LineHeightIsMultiple = false;
					break;
				case "ql":
					state.Alignment = StiTextHorAlignment.Left;
					break;
				case "li":
					state.LeftIndent = node.Parameter;
					break;
				case "ri":
					state.RightIndent = node.Parameter;
					break;
				case "fi":
					state.FirstIndent = node.Parameter;
					break;
				case "sb":
					state.SpaceBefore = node.Parameter;
					break;
				case "sa":
					state.SpaceAfter = node.Parameter;
					break;
				case "sl":
					state.LineHeight = node.Parameter;
					break;
				case "slmult":
					state.LineHeightIsMultiple = node.Parameter == 1;
					break;
				case "qr":
					state.Alignment = StiTextHorAlignment.Right;
					break;
				case "qc":
					state.Alignment = StiTextHorAlignment.Center;
					break;
				case "qj":
					state.Alignment = StiTextHorAlignment.Width;
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
					int code = node.Parameter & 0xFFFF;
					char sign = code >= 0xF020 && code <= 0xF0FF ? GetSymbol(code) : (char)0;
					AppendChar(res, sign != 0 ? sign : (char)code, state);
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

		//The band for the rows of a table which are repeated on every page. It stands before the data band
		//with the other rows of the table.
		private StiBand AddHeaderBand(StiPage page)
		{
			StiHeaderBand band = new StiHeaderBand();
			band.Top = defaultTextBoxHeight * 2 * bandsCounter;
			band.Name = string.Format("HeaderBand{0}", bandsCounter++);
			band.PrintOnAllPages = true;
			page.Components.Add(band);
			return band;
		}

		//The band which takes the next text or table: a new data band of the body,
		//or the header or footer of the page.
		private StiBand GetContentBand(StiPage page, string pageBandKey)
		{
			StiDataBand band = new StiDataBand();
			if (pageBandKey != null)
			{
				//the band of a dropped kind of header stays out of the page
				StiBand pageBand = GetPageBand(page, pageBandKey);
				return pageBand != null ? pageBand : band;
			}

			//the bands keep the order of the document: a text can give several of them, with pictures between
			band.Top = defaultTextBoxHeight * 2 * bandsCounter;
			band.Name = string.Format("DataBand{0}", bandsCounter++);
			band.CountData = 1;
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

		private void AddTextBand(StiPage page, string pageBand, StringBuilder res, StiRtfState state,
			ArrayList runs, ArrayList fontTable, ArrayList colorTable, StiRtfPageData pageData)
		{
			StiBand band = GetContentBand(page, pageBand);
			//the content of a header or of a footer is stacked in its band
			double bandTop = band is StiDataBand ? 0 : band.Height;

			//the end of the last paragraph is not an empty line of the text
			int length = res.Length;
			if (length >= 2 && res[length - 2] == '\r' && res[length - 1] == '\n') res.Length = length - 2;

			string plainText = res.ToString();
			state = GetMainState(state, runs, plainText);
			string html = GetHtmlText(plainText, runs, ref state, fontTable, colorTable);
			runs.Clear();

			StiText text = new StiText();
			SetText(text, html ?? plainText);
			text.AllowHtmlTags = html != null;
			text.HorAlignment = state.Alignment;
			res.Length = 0;
			text.Width = Convert(pageData.SectWidth - pageData.SectMarginLeft - pageData.SectMarginRight);
			text.Height = defaultTextBoxHeight;

			//a hanging indent starts the first line on the left of the others: the text stands where it starts
			double leftIndent = Math.Max(0, Convert(state.LeftIndent + Math.Min(state.FirstIndent, 0)));
			double rightIndent = Math.Max(0, Convert(state.RightIndent));
			if (text.Width - leftIndent - rightIndent < MinTextWidth)
			{
				leftIndent = 0;
				rightIndent = 0;
			}

			//the indents and the space around the paragraph are the margins of the text: the component keeps
			//its place and its width, so several of them can be put on one band
			double spaceBefore = Math.Max(0, Convert(state.SpaceBefore));
			double spaceAfter = Math.Max(0, Convert(state.SpaceAfter));
			text.Margins = new StiMargins(leftIndent, rightIndent, spaceBefore, spaceAfter);

			//"sl240" with "slmult1" is the single spacing
			if (state.LineHeightIsMultiple && state.LineHeight > 0)
			{
				float lineSpacing = state.LineHeight / 240f;
				text.LineSpacing = Math.Max(MinLineSpacing, Math.Min(MaxLineSpacing, lineSpacing));
			}

			FontStyle fs = FontStyle.Regular;
			if (state.Bold) fs |= FontStyle.Bold;
			if (state.Italic) fs |= FontStyle.Italic;
			if (state.Underline) fs |= FontStyle.Underline;
			float fontSize = (float)(state.FontSize > 0 ? state.FontSize : defaultFontSize);
			StiRtfFontData fontData = GetFontData(state, fontTable);
			if (fontData.Bold) fs |= FontStyle.Bold;
			if (fontData.Italic) fs |= FontStyle.Italic;
			text.Font = new Font(fontData.Name, fontSize, fs);
			text.TextBrush = new StiSolidBrush(GetColor(colorTable, state.FontColor));

			text.Parent = band;
			text.Page = page;
			text.CanGrow = true;
			text.WordWrap = true;
			SizeD textSize = text.GetActualSize();
			if (textSize.Height > text.Height) text.Height = textSize.Height;
			text.CanBreak = true;
			text.Top = bandTop;
			band.Components.Add(text);
			band.Height = bandTop + text.Height;
		}

		private static byte[] HexToBytes(StringBuilder hex)
		{
			MemoryStream stream = new MemoryStream(hex.Length / 2);
			int high = -1;
			for (int index = 0; index < hex.Length; index++)
			{
				char symbol = hex[index];
				int digit;
				if (symbol >= '0' && symbol <= '9') digit = symbol - '0';
				else if (symbol >= 'a' && symbol <= 'f') digit = symbol - 'a' + 10;
				else if (symbol >= 'A' && symbol <= 'F') digit = symbol - 'A' + 10;
				else continue;

				if (high < 0)
				{
					high = digit;
				}
				else
				{
					stream.WriteByte((byte)(high * 16 + digit));
					high = -1;
				}
			}
			return stream.ToArray();
		}

		private static void WriteInt16(byte[] data, int offset, int value)
		{
			data[offset] = (byte)(value & 0xFF);
			data[offset + 1] = (byte)((value >> 8) & 0xFF);
		}

		//RTF keeps a metafile without the placeable header, and it is not an image file without it.
		//The size is in the units of the metafile: 0.01 mm for the usual mapping mode.
		private static byte[] AddMetafileHeader(byte[] data, int width, int height)
		{
			if (width <= 0 || height <= 0) return null;

			int unitsPerInch = 2540;
			while (width > 32767 || height > 32767)
			{
				width /= 2;
				height /= 2;
				unitsPerInch /= 2;
			}

			byte[] result = new byte[data.Length + 22];
			result[0] = 0xD7;
			result[1] = 0xCD;
			result[2] = 0xC6;
			result[3] = 0x9A;
			WriteInt16(result, 10, width);
			WriteInt16(result, 12, height);
			WriteInt16(result, 14, unitsPerInch);

			int checksum = 0;
			for (int index = 0; index < 20; index += 2)
			{
				checksum ^= result[index] | (result[index + 1] << 8);
			}
			WriteInt16(result, 20, checksum);

			Array.Copy(data, 0, result, 22, data.Length);
			return result;
		}

		//RTF keeps a bitmap as a DIB, the file of a bitmap has one more header before it.
		private static byte[] AddBitmapHeader(byte[] data)
		{
			if (data.Length < 40) return null;

			int headerSize = BitConverter.ToInt32(data, 0);
			int bitCount = BitConverter.ToInt16(data, 14);
			int compression = BitConverter.ToInt32(data, 16);
			int colors = BitConverter.ToInt32(data, 32);
			if (headerSize < 40 || headerSize > data.Length || colors < 0 || colors > 256) return null;

			if (colors == 0 && bitCount <= 8) colors = 1 << bitCount;
			int offset = 14 + headerSize + colors * 4;
			if (compression == 3 && headerSize == 40) offset += 12;

			byte[] result = new byte[data.Length + 14];
			result[0] = (byte)'B';
			result[1] = (byte)'M';
			Array.Copy(BitConverter.GetBytes(result.Length), 0, result, 2, 4);
			Array.Copy(BitConverter.GetBytes(offset), 0, result, 10, 4);
			Array.Copy(data, 0, result, 14, data.Length);
			return result;
		}

		private static bool StartsWith(byte[] data, int offset, params byte[] signature)
		{
			if (data.Length < offset + signature.Length) return false;

			for (int index = 0; index < signature.Length; index++)
			{
				if (data[offset + index] != signature[index]) return false;
			}
			return true;
		}

		//The kind of the image by its first bytes: the keyword of the format can be wrong, and a picture can
		//be an empty stub. Returns "file" for a ready image file, "wmf" and "dib" for the data which need
		//a header, null for the data which is not an image.
		private static string GetImageKind(byte[] data, string format)
		{
			if (data.Length < 20) return null;

			if (StartsWith(data, 0, 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A))
			{
				//the signature, the header and the end of the smallest PNG take 45 bytes
				return data.Length >= 45 ? "file" : null;
			}
			if (StartsWith(data, 0, 0xFF, 0xD8, 0xFF)) return "file";
			if (StartsWith(data, 0, 0x42, 0x4D)) return "file";
			if (StartsWith(data, 0, 0xD7, 0xCD, 0xC6, 0x9A)) return "file";
			if (StartsWith(data, 0, 1, 0, 0, 0) && StartsWith(data, 40, 0x20, 0x45, 0x4D, 0x46)) return "file";
			if ((data[0] == 1 || data[0] == 2) && StartsWith(data, 1, 0, 9, 0)) return "wmf";

			return format == "dibitmap" ? "dib" : null;
		}

		//A bitmap is not compressed: it makes the report several times bigger than the same image as PNG.
		private static byte[] ConvertBitmapToPng(byte[] data)
		{
			try
			{
				using (MemoryStream input = new MemoryStream(data))
				using (Bitmap bitmap = new Bitmap(input))
				using (MemoryStream output = new MemoryStream())
				{
					bitmap.Save(output, ImageFormat.Png);
					return output.ToArray();
				}
			}
			catch (Exception)
			{
				//a bitmap which can't be read stays as it is, the report decides what to do with it
				return data;
			}
		}

		//Reads the group "pict": the image and its size on the page. Returns null for a format without a reader.
		private StiRtfPicture ReadPicture(StiRtfTreeNode group)
		{
			string format = null;
			int width = 0;
			int height = 0;
			int goalWidth = 0;
			int goalHeight = 0;
			int scaleX = 100;
			int scaleY = 100;
			StringBuilder hex = new StringBuilder();

			for (int index = 0; index < group.ChildNodes.Count; index++)
			{
				StiRtfTreeNode node = group.ChildNodes[index];
				if (node.NodeType == StiRtfNodeType.Text) hex.Append(node.NodeKey);
				if (node.NodeType != StiRtfNodeType.Keyword) continue;

				switch (node.NodeKey)
				{
					case "pngblip":
					case "jpegblip":
					case "emfblip":
					case "wmetafile":
					case "dibitmap":
						format = node.NodeKey;
						break;
					case "picw":
						width = node.Parameter;
						break;
					case "pich":
						height = node.Parameter;
						break;
					case "picwgoal":
						goalWidth = node.Parameter;
						break;
					case "pichgoal":
						goalHeight = node.Parameter;
						break;
					case "picscalex":
						if (node.Parameter > 0) scaleX = node.Parameter;
						break;
					case "picscaley":
						if (node.Parameter > 0) scaleY = node.Parameter;
						break;
					//the binary data is not read by the lexer
					case "bin":
						return null;
				}
			}
			if (format == null) return null;

			bool isMetafile = format == "wmetafile" || format == "emfblip";
			//the size in twips: the wanted one, or the own size of the image in 0.01 mm or in pixels
			double twipsWidth = goalWidth > 0 ? goalWidth : (isMetafile ? width * 1440d / 2540 : width * 15d);
			double twipsHeight = goalHeight > 0 ? goalHeight : (isMetafile ? height * 1440d / 2540 : height * 15d);

			byte[] data = HexToBytes(hex);
			string kind = GetImageKind(data, format);
			if (kind == null) return null;

			if (kind == "wmf")
			{
				if (width <= 0) width = (int)(goalWidth * 2540d / 1440);
				if (height <= 0) height = (int)(goalHeight * 2540d / 1440);
				data = AddMetafileHeader(data, width, height);
			}
			if (kind == "dib") data = AddBitmapHeader(data);
			if (data == null) return null;
			if (StartsWith(data, 0, 0x42, 0x4D)) data = ConvertBitmapToPng(data);

			StiRtfPicture picture = new StiRtfPicture();
			picture.Data = data;
			picture.Width = twipsWidth > 0 ? Convert(twipsWidth * scaleX / 100) : 100;
			picture.Height = twipsHeight > 0 ? Convert(twipsHeight * scaleY / 100) : 100;
			return picture;
		}

		//The picture of the group which the reading has come to, null when the group is not a shown picture.
		private StiRtfPicture ReadPictureGroup(StiRtfTreeNode group, int skipGroupCount, StiRtfState state)
		{
			if (skipGroupCount > 0 || state.Hidden) return null;

			StiRtfTreeNode first = group.ChildNodes[0];
			if (first.NodeType != StiRtfNodeType.Keyword || first.NodeKey != "pict") return null;

			StiRtfPicture picture = ReadPicture(group);
			if (picture != null && group.ParentNode != null && IsPictureHolder(group.ParentNode))
			{
				pictureCopyIsExtra = true;
			}
			return picture;
		}

		//The starred groups which are read, not skipped. Word writes a picture twice: in "shppict" and,
		//for the old readers, in "nonshppict". A shape with a text (a text box) keeps it in "shptxt"
		//inside "shpinst"; a shape without a text stays skipped, its picture comes from "shprslt".
		private bool IsContentHolder(StiRtfTreeNode group)
		{
			if (group.ChildNodes.Count < 2) return false;

			StiRtfTreeNode second = group.ChildNodes[1];
			if (second.NodeType != StiRtfNodeType.Keyword) return false;
			if (second.NodeKey == "shppict") return true;
			if (second.NodeKey != "shpinst") return false;

			for (int index = 2; index < group.ChildNodes.Count; index++)
			{
				StiRtfTreeNode child = group.ChildNodes[index];
				if (child.NodeType != StiRtfNodeType.Group) continue;

				StiRtfTreeNode first = child.ChildNodes[0];
				if (first.NodeType == StiRtfNodeType.Keyword && first.NodeKey == "shptxt")
				{
					shapeResultIsExtra = true;
					return true;
				}
			}
			return false;
		}

		private static bool IsPictureHolder(StiRtfTreeNode group)
		{
			if (group.ChildNodes.Count < 2) return false;

			StiRtfTreeNode second = group.ChildNodes[1];
			return second.NodeType == StiRtfNodeType.Keyword && second.NodeKey == "shppict";
		}

		private static StiImage CreateImage(StiRtfPicture picture, double maxWidth, StiPage page)
		{
			double width = picture.Width;
			double height = picture.Height;
			if (maxWidth > 0 && width > maxWidth)
			{
				height = Math.Round(height * maxWidth / width, 1);
				width = maxWidth;
			}

			StiImage image = new StiImage();
			image.ImageBytes = picture.Data;
			image.Stretch = true;
			image.AspectRatio = true;
			image.Width = width;
			image.Height = height;
			image.Page = page;
			return image;
		}

		private void AddImageBand(StiPage page, string pageBand, StiRtfPicture picture, StiRtfPageData pageData)
		{
			StiBand band = GetContentBand(page, pageBand);
			double bandTop = band is StiDataBand ? 0 : band.Height;
			double maxWidth = Convert(pageData.SectWidth - pageData.SectMarginLeft - pageData.SectMarginRight);

			StiImage image = CreateImage(picture, maxWidth, page);
			image.Top = bandTop;
			image.Parent = band;
			band.Components.Add(image);
			band.Height = bandTop + image.Height;
		}

		private static double GetPicturesHeight(StiRtfCell cell)
		{
			if (cell.Pictures == null) return 0;

			double height = 0;
			foreach (StiRtfPicture picture in cell.Pictures)
			{
				double width = cell.Content.Width;
				bool isWide = width > 0 && picture.Width > width;
				height += isWide ? Math.Round(picture.Height * width / picture.Width, 1) : picture.Height;
			}
			return height;
		}

		//The pictures of a cell lie over its text component, one under another from the top of the cell.
		private static void AddCellPictures(StiBand band, StiRtfCell cell, StiPage page)
		{
			if (cell.Pictures == null) return;

			double top = cell.Content.Top;
			foreach (StiRtfPicture picture in cell.Pictures)
			{
				StiImage image = CreateImage(picture, cell.Content.Width, page);
				image.Left = cell.Content.Left;
				image.Top = top;
				image.Parent = band;
				band.Components.Add(image);
				top += image.Height;
			}
		}

		private static bool IsSameCharacterFormat(StiRtfState first, StiRtfState second)
		{
			return first.FontNumber == second.FontNumber && first.FontSize == second.FontSize &&
				first.Bold == second.Bold && first.Italic == second.Italic && first.Underline == second.Underline &&
				first.FontColor == second.FontColor;
		}

		//Remembers the formatting of the characters which were added to the text from the given position.
		private static void AddTextRun(ArrayList runs, StiRtfState state, StringBuilder res, int start)
		{
			if (res.Length <= start) return;

			StiRtfTextRun last = runs.Count > 0 ? (StiRtfTextRun)runs[runs.Count - 1] : null;
			if (last != null && last.Start + last.Length == start && IsSameCharacterFormat(last.State, state))
			{
				last.Length = res.Length - last.Start;
				return;
			}

			StiRtfTextRun run = new StiRtfTextRun();
			run.State = state;
			run.Start = start;
			run.Length = res.Length - start;
			runs.Add(run);
		}

		//The count of the characters of the run which are seen: the formatting of a space means nothing.
		private static int GetTextLength(string res, StiRtfTextRun run)
		{
			int count = 0;
			int end = Math.Min(res.Length, run.Start + run.Length);
			for (int index = run.Start; index < end; index++)
			{
				if (!char.IsWhiteSpace(res[index])) count++;
			}
			return count;
		}

		//A text component has one font for the whole text. It is the font of the most of its characters, not
		//of the first one: a paragraph may start with a marker of a list in the Symbol font or with a word
		//in bold.
		private static StiRtfState GetMainState(StiRtfState beginState, ArrayList runs, string res)
		{
			ArrayList formats = new ArrayList();
			StiRtfFormatCount main = null;
			foreach (StiRtfTextRun run in runs)
			{
				int count = GetTextLength(res, run);
				if (count == 0) continue;

				StiRtfFormatCount format = null;
				foreach (StiRtfFormatCount item in formats)
				{
					if (IsSameCharacterFormat(item.State, run.State)) format = item;
				}
				if (format == null)
				{
					format = new StiRtfFormatCount();
					format.State = run.State;
					formats.Add(format);
				}
				format.Count += count;
				if (main == null || format.Count > main.Count) main = format;
			}
			if (main == null) return beginState;

			StiRtfState state = beginState;
			state.FontNumber = main.State.FontNumber;
			state.AFontNumber = main.State.AFontNumber;
			state.FontSize = main.State.FontSize;
			state.Bold = main.State.Bold;
			state.Italic = main.State.Italic;
			state.Underline = main.State.Underline;
			state.FontColor = main.State.FontColor;
			return state;
		}

		private static void AppendHtml(StringBuilder html, string res, int start, int end)
		{
			for (int index = start; index < end; index++)
			{
				char symbol = res[index];
				switch (symbol)
				{
					case '&':
						html.Append("&amp;");
						break;

					case '<':
						html.Append("&lt;");
						break;

					case '>':
						html.Append("&gt;");
						break;

					//the end of a line is a tag: a text with tags does not break the line at the line feed
					case '\r':
						break;

					case '\n':
						html.Append("<br>");
						break;

					default:
						html.Append(symbol);
						break;
				}
			}
		}

		private static string EscapeHtml(string text)
		{
			StringBuilder html = new StringBuilder(text.Length);
			AppendHtml(html, text, 0, text.Length);
			return html.ToString();
		}

		//The part of the text in the tags of its own formatting. The font of the component gives the name,
		//the size and the color of the main formatting, so only the difference is written.
		private void AppendHtmlRun(StringBuilder html, string res, int start, int end, StiRtfState state,
			StiRtfState mainState, ArrayList fontTable, ArrayList colorTable)
		{
			StiRtfFontData fontData = GetFontData(state, fontTable);
			StiRtfFontData mainFontData = GetFontData(mainState, fontTable);

			StringBuilder font = new StringBuilder();
			//a marker in a font of symbols is a sign of Unicode already, it needs the font of the text
			bool isSymbolFont = fontData.Charset == 2 || fontData.Name == "Symbol";
			if (fontData.Name != mainFontData.Name && !isSymbolFont && !string.IsNullOrEmpty(fontData.Name))
			{
				font.Append(" face=\"").Append(fontData.Name.Replace("\"", string.Empty)).Append("\"");
			}
			if (state.FontSize > 0 && state.FontSize != mainState.FontSize)
			{
				font.Append(" size=\"").Append(state.FontSize.ToString(CultureInfo.InvariantCulture)).Append("\"");
			}
			if (state.FontColor != mainState.FontColor)
			{
				Color color = GetColor(colorTable, state.FontColor);
				font.Append(string.Format(" color=\"#{0:x2}{1:x2}{2:x2}\"", color.R, color.G, color.B));
			}

			bool bold = state.Bold || fontData.Bold;
			bool italic = state.Italic || fontData.Italic;
			if (font.Length > 0) html.Append("<font").Append(font).Append(">");
			if (bold) html.Append("<b>");
			if (italic) html.Append("<i>");
			if (state.Underline) html.Append("<u>");
			AppendHtml(html, res, start, end);
			if (state.Underline) html.Append("</u>");
			if (italic) html.Append("</i>");
			if (bold) html.Append("</b>");
			if (font.Length > 0) html.Append("</font>");
		}

		//A text component has one font, so a text with several ones is written with the tags of HTML.
		//Returns null for a text with one formatting: it stays a plain text. The main formatting loses its
		//bold, italic and underline then: a tag can switch them on, but not off.
		private string GetHtmlText(string res, ArrayList runs, ref StiRtfState mainState,
			ArrayList fontTable, ArrayList colorTable)
		{
			bool isMixed = false;
			foreach (StiRtfTextRun run in runs)
			{
				if (!IsSameCharacterFormat(run.State, mainState) && GetTextLength(res, run) > 0) isMixed = true;
			}
			if (!isMixed) return null;

			mainState.Bold = false;
			mainState.Italic = false;
			mainState.Underline = false;

			StringBuilder html = new StringBuilder(res.Length + runs.Count * 16);
			int position = 0;
			foreach (StiRtfTextRun run in runs)
			{
				int start = Math.Max(position, Math.Min(res.Length, run.Start));
				int end = Math.Max(start, Math.Min(res.Length, run.Start + run.Length));
				//a text between the runs is an expression of a field
				AppendHtml(html, res, position, start);
				//a run of spaces needs no tags; any other one does, the main formatting has no bold now
				if (GetTextLength(res, run) > 0)
				{
					AppendHtmlRun(html, res, start, end, run.State, mainState, fontTable, colorTable);
				}
				else
				{
					AppendHtml(html, res, start, end);
				}
				position = end;
			}
			AppendHtml(html, res, position, res.Length);
			return html.ToString();
		}

		//The formatting which a text component keeps for the whole text: a paragraph with another one
		//starts a new component.
		private static bool IsSameFormat(StiRtfState first, StiRtfState second)
		{
			return first.FontNumber == second.FontNumber && first.FontSize == second.FontSize &&
				first.Bold == second.Bold && first.Italic == second.Italic && first.Underline == second.Underline &&
				first.FontColor == second.FontColor && first.Alignment == second.Alignment &&
				IsSameParagraphFormat(first, second);
		}

		//The paragraphs with the same margins stay in one text, the space between them is lost: it is more
		//important to keep them together.
		private static bool IsSameParagraphFormat(StiRtfState first, StiRtfState second)
		{
			return first.SpaceBefore == second.SpaceBefore && first.SpaceAfter == second.SpaceAfter &&
				first.LeftIndent == second.LeftIndent && first.RightIndent == second.RightIndent &&
				first.FirstIndent == second.FirstIndent && first.LineHeight == second.LineHeight &&
				first.LineHeightIsMultiple == second.LineHeightIsMultiple;
		}

		private static bool IsContent(StiRtfTreeNode node)
		{
			if (node.NodeType == StiRtfNodeType.Text) return node.NodeKey.Trim().Length > 0;
			if (node.NodeType == StiRtfNodeType.Control) return node.NodeKey == "'";

			return node.NodeType == StiRtfNodeType.Keyword && node.NodeKey == "u";
		}

		//Returns false for a keyword which is not a kind of the line of a border.
		private static bool TryGetBorderStyle(string key, out StiPenStyle style)
		{
			style = StiPenStyle.Solid;
			if (!key.StartsWith("brdr", StringComparison.Ordinal)) return false;

			switch (key)
			{
				//the width, the color, the shadow and the space of a border, the borders of a paragraph
				case "brdrw":
				case "brdrcf":
				case "brdrsh":
				case "brdrsp":
				case "brdrart":
				case "brdrt":
				case "brdrb":
				case "brdrl":
				case "brdrr":
				case "brdrbtw":
				case "brdrbar":
					return false;

				case "brdrnone":
				case "brdrnil":
				case "brdrtbl":
					style = StiPenStyle.None;
					return true;

				case "brdrdot":
					style = StiPenStyle.Dot;
					return true;

				case "brdrdash":
				case "brdrdashsm":
					style = StiPenStyle.Dash;
					return true;

				case "brdrdashd":
				case "brdrdashdotstr":
					style = StiPenStyle.DashDot;
					return true;

				case "brdrdashdd":
					style = StiPenStyle.DashDotDot;
					return true;

				//a thick line
				case "brdrth":
					return true;
			}

			//the double and triple lines, the pairs of a thin and a thick line: "brdrtnthsg", "brdrthtnlg"
			bool isDouble = key == "brdrdb" || key == "brdrtriple" || key == "brdrwavydb" ||
				key.StartsWith("brdrtn", StringComparison.Ordinal) ||
				key.StartsWith("brdrth", StringComparison.Ordinal);
			if (isDouble) style = StiPenStyle.Double;

			return true;
		}

		//A width or a color of a side shows it with a solid line when the kind of the line is not set.
		private static void ShowBorderSide(StiBorderSide side)
		{
			if (side.Style == StiPenStyle.None) side.Style = StiPenStyle.Solid;
		}

		//The border of a cell: a simple one when all its sides have the same line.
		private static StiBorder GetBorder(StiAdvancedBorder sides)
		{
			StiBorderSide[] list = new StiBorderSide[]
			{
				sides.TopSide, sides.BottomSide, sides.LeftSide, sides.RightSide
			};
			StiBorderSide first = null;
			bool isSame = true;
			foreach (StiBorderSide side in list)
			{
				if (side.Style == StiPenStyle.None) continue;

				if (first == null)
				{
					first = side;
				}
				else if (side.Style != first.Style || side.Size != first.Size || side.Color != first.Color)
				{
					isSame = false;
				}
			}

			if (first == null) return new StiBorder();
			if (!isSame) return sides;

			return new StiBorder(sides.Side, first.Color, first.Size, first.Style);
		}

		//The leading rows of a table with "trhdr" are repeated on every page: they go to a header band.
		//It can't be done when a cell of these rows is merged with the cells below them.
		private static bool CanSplitHeader(ArrayList table, int headerLines)
		{
			if (headerLines <= 0 || headerLines >= table.Count) return false;

			for (int indexLine = 0; indexLine < headerLines; indexLine++)
			{
				foreach (StiRtfCell cell in (ArrayList)table[indexLine])
				{
					if (indexLine + cell.MergedCellsCount >= headerLines) return false;
				}
			}
			return true;
		}

		//The left position of a row: its offset ("trleft"), or the place of a centered or right aligned table.
		private static double GetRowLeft(StiRtfCell[,] matrix, int indexLine, int cellsCount, double pageWidth)
		{
			StiRtfCell firstCell = matrix[indexLine, 0];
			if (firstCell == null) return 0;

			double width = 0;
			for (int indexCell = 0; indexCell < cellsCount; indexCell++)
			{
				StiRtfCell cell = matrix[indexLine, indexCell];
				if (cell != null) width += cell.Content.Width;
			}

			double left = firstCell.RowLeft;
			if (firstCell.RowAlignment == StiTextHorAlignment.Center) left = (pageWidth - width) / 2;
			if (firstCell.RowAlignment == StiTextHorAlignment.Right) left = pageWidth - width;

			return Math.Max(0, Math.Round(left, 1));
		}

		//Joins a cell to the cell on its left. Some writers put the text of the merged cell
		//not to the first of its parts, so the text of the joined part is kept too.
		private static void JoinCell(StiRtfCell firstCell, StiRtfCell cell)
		{
			string text = cell.Content.Text.Value;
			if (!string.IsNullOrEmpty(text))
			{
				string firstText = firstCell.Content.Text.Value;
				//a plain text which is joined to a text with tags gets the same notation
				bool isHtml = cell.Content.AllowHtmlTags || firstCell.Content.AllowHtmlTags;
				if (isHtml && !cell.Content.AllowHtmlTags) text = EscapeHtml(text);
				if (isHtml && !firstCell.Content.AllowHtmlTags && !string.IsNullOrEmpty(firstText))
				{
					firstText = EscapeHtml(firstText);
				}
				firstCell.Content.AllowHtmlTags = isHtml;
				if (string.IsNullOrEmpty(firstText))
				{
					firstCell.Content.Text = text;
					firstCell.Content.Font = cell.Content.Font;
					firstCell.Content.TextBrush = cell.Content.TextBrush;
					firstCell.Content.HorAlignment = cell.Content.HorAlignment;
				}
				else
				{
					firstCell.Content.Text = firstText + " " + text;
				}
				if (cell.Content.OnlyText) firstCell.Content.OnlyText = true;
			}

			if (cell.Pictures != null)
			{
				if (firstCell.Pictures == null) firstCell.Pictures = new ArrayList();
				firstCell.Pictures.AddRange(cell.Pictures);
			}

			//the right side of the merged cell is the right side of its last part
			firstCell.Sides.RightSide = cell.Sides.RightSide;
			firstCell.Content.Width += cell.Content.Width;
			cell.Content.Width = 0;
			cell.Merged = true;
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
			//The page of the report swaps its width and height when the orientation changes, and RTF already keeps
			//the size of a landscape page as it is. So the orientation is set by the size and before it.
			StiPageOrientation orientation = pageData.SectWidth > pageData.SectHeight
				? StiPageOrientation.Landscape
				: StiPageOrientation.Portrait;
			if (page.Orientation != orientation) page.Orientation = orientation;

			page.PageWidth = Convert(pageData.SectWidth);
			page.PageHeight = Convert(pageData.SectHeight);
            page.Margins = new StiMargins(
                Convert(pageData.SectMarginLeft),
                Convert(pageData.SectMarginRight),
                Convert(pageData.SectMarginTop),
                Convert(pageData.SectMarginBottom));
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
                pictureCopyIsExtra = false;
                shapeResultIsExtra = false;
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
                            //the character set of the document: Macintosh, IBM PC, IBM PC of Western Europe
                            if (currentNode.NodeKey == "mac") mainCodePage = 10000;
                            if (currentNode.NodeKey == "pc") mainCodePage = 437;
                            if (currentNode.NodeKey == "pca") mainCodePage = 850;
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
                //the cells of the last row: a row without a definition takes them even after a text
                ArrayList lastRowCells = null;
                #endregion

                #region parse text
                for (int indexBand = 0; indexBand < bandsList.Count; indexBand++)
                {
                    StiRtfTree.StiRtfBandData data = (StiRtfTree.StiRtfBandData)bandsList[indexBand];
                    StiRtfTreeNode node = data.start;
                    if (!data.isTableRow)
                    {
                        #region plain text
                        bool paragraphEnded = false;
                        StringBuilder res = new StringBuilder();
                        ArrayList formats = new ArrayList();
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
                                    else if (nestedLevel == 0 && (skipGroupCount > 0 || !IsContentHolder(node)))
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
                                StiRtfPicture picture = ReadPictureGroup(node, skipGroupCount, currentState);
                                if (picture != null)
                                {
                                    //the picture stands in the text: the text before it gets its own band
                                    if (res.Length > 0)
                                    {
                                        AddTextBand(currentPage, data.pageBand, res, beginState, formats,
                                            fontTable, colorTable, pageData);
                                    }
                                    AddImageBand(currentPage, data.pageBand, picture, pageData);
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
                                if (node.NodeType == StiRtfNodeType.Keyword && node.NodeKey == "par")
                                {
                                    paragraphEnded = true;
                                }
                                else if (paragraphEnded && IsContent(node))
                                {
                                    //the next paragraph has its own font or alignment: the text before it is closed
                                    paragraphEnded = false;
                                    if (res.Length > 0 && !IsSameFormat(beginState, currentState))
                                    {
                                        AddTextBand(currentPage, data.pageBand, res, beginState, formats,
                                            fontTable, colorTable, pageData);
                                        beginState = currentState;
                                    }
                                }
                                int textLength = res.Length;
                                StiRtfState textState = currentState;
                                ProcessTextToken(node, res, ref currentState, fontTable);
                                AddTextRun(formats, textState, res, textLength);
                            }
                            if (node == data.end) endLoop = true;
                            node = node.NextFlateToken;
                            if (node == null) endLoop = true;
                        }
                        if (res.Length > 0)
                        {
                            AddTextBand(currentPage, data.pageBand, res, beginState, formats,
                                fontTable, colorTable, pageData);
                        }

                        #endregion
                    }
                    else
                    {
                        #region table rows
                        ArrayList previousCells = data.isIntbl ? lastRowCells : null;
                        //the count of the leading rows of the table which are repeated on every page
                        int headerLines = 0;
                        bool headerRow = false;
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
                                    previousCell.Content.AllowHtmlTags = false;
                                    previousCell.Pictures = null;
                                    cells.Add(previousCell);
                                }
                                if (cells.Count > 0) currentCell = (StiRtfCell)cells[0];
                            }
                            int cellIndex = 0;
                            double posX = 0;
                            StringBuilder res = new StringBuilder();
                            ArrayList formats = new ArrayList();
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
                            //the pictures which stand before the definition of the row, they go to its first cell
                            ArrayList contentPictures = null;
                            //the side of the cell which takes the keywords of a border,
                            //null for a border of another kind
                            StiBorderSide borderSide = null;
                            //the side has no line, so its width and color don't show it
                            bool borderHidden = false;
                            double rowLeft = 0;
                            StiTextHorAlignment rowAlignment = StiTextHorAlignment.Left;
                            double lineHeight = defaultTextBoxHeight;
                            double cellGap = 0;
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
                                        else if (nestedLevel == 0 && (skipGroupCount > 0 || !IsContentHolder(node)))
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
                                    StiRtfPicture picture = ReadPictureGroup(node, skipGroupCount, currentState);
                                    if (picture != null)
                                    {
                                        if (currentCell.Pictures == null) currentCell.Pictures = new ArrayList();
                                        currentCell.Pictures.Add(picture);
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
                                    int textLength = res.Length;
                                    StiRtfState textState = currentState;
                                    ProcessTextToken(node, res, ref currentState, fontTable);
                                    AddTextRun(formats, textState, res, textLength);
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

                                    StiPenStyle borderStyle;
                                    bool isBorderStyle = isRowKeyword && borderSide != null;
                                    if (isBorderStyle && TryGetBorderStyle(node.NodeKey, out borderStyle))
                                    {
                                        borderSide.Style = borderStyle;
                                        borderHidden = borderStyle == StiPenStyle.None;
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
                                                ArrayList rowPictures = currentCell.Pictures;
                                                if (rowPictures != null) contentPictures = rowPictures;
                                                cells = new ArrayList();
                                                currentCell = new StiRtfCell();
                                                cellIndex = 0;
                                                cellContentPresent = false;
                                                headerRow = false;
                                                break;

                                                case "trhdr":
                                                    headerRow = true;
                                                    break;

                                                case "trleft":
                                                    if (node.Parameter > 0 && cells.Count == 0)
                                                    {
                                                        rowLeft = Convert(node.Parameter);
                                                        posX = rowLeft;
                                                    }
                                                    break;

                                                case "trql":
                                                    rowAlignment = StiTextHorAlignment.Left;
                                                    break;
                                                case "trqc":
                                                    rowAlignment = StiTextHorAlignment.Center;
                                                    break;
                                                case "trqr":
                                                    rowAlignment = StiTextHorAlignment.Right;
                                                    break;

                                                case "trrh":
                                                lineHeight = Convert(node.Parameter);
                                                break;

                                            //the gap between the border of a cell and its text
                                            case "trgaph":
                                                cellGap = Convert(node.Parameter);
                                                break;

                                            case "clmrg":
                                                currentCell.JoinedToLeft = true;
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
                                                borderSide = currentCell.Sides.TopSide;
                                                borderHidden = false;
                                                break;
                                            case "clbrdrb":
                                                borderSide = currentCell.Sides.BottomSide;
                                                borderHidden = false;
                                                break;
                                            case "clbrdrl":
                                                borderSide = currentCell.Sides.LeftSide;
                                                borderHidden = false;
                                                break;
                                            case "clbrdrr":
                                                borderSide = currentCell.Sides.RightSide;
                                                borderHidden = false;
                                                break;

                                            //the borders of a paragraph and of the row, the diagonals:
                                            //they are not the sides of a cell
                                            case "brdrt":
                                            case "brdrb":
                                            case "brdrl":
                                            case "brdrr":
                                            case "brdrbtw":
                                            case "brdrbar":
                                            case "box":
                                            case "trbrdrt":
                                            case "trbrdrb":
                                            case "trbrdrl":
                                            case "trbrdrr":
                                            case "trbrdrh":
                                            case "trbrdrv":
                                            case "cldglu":
                                            case "cldgll":
                                                borderSide = null;
                                                break;

                                            case "brdrw":
                                                if (borderSide != null)
                                                {
                                                    borderSide.Size = Convert(node.Parameter);
                                                    if (!borderHidden) ShowBorderSide(borderSide);
                                                }
                                                break;

                                            case "brdrcf":
                                                if (borderSide != null)
                                                {
                                                    borderSide.Color = GetColor(colorTable, node.Parameter);
                                                    if (!borderHidden) ShowBorderSide(borderSide);
                                                }
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
                                                if (cells.Count == 0)
                                                {
                                                    currentCell.RowLeft = rowLeft;
                                                    currentCell.RowAlignment = rowAlignment;
                                                }
                                                if (cellGap > 0)
                                                {
                                                    StiMargins gap = new StiMargins(cellGap, cellGap, 0, 0);
                                                    currentCell.Content.Margins = gap;
                                                }
                                                currentCell.Content.Height = Math.Abs(lineHeight);
                                                if (lineHeight < 0) currentCell.FixedHeight = true;
                                                if (cells.Count < pendingCells.Count)
                                                {
                                                    StiText content = ((StiRtfCell)pendingCells[cells.Count]).Content;
                                                    currentCell.Content.Text = content.Text;
                                                    currentCell.Content.OnlyText = content.OnlyText;
                                                    currentCell.Content.AllowHtmlTags = content.AllowHtmlTags;
                                                    currentCell.Content.Font = content.Font;
                                                    currentCell.Content.TextBrush = content.TextBrush;
                                                    currentCell.Content.HorAlignment = content.HorAlignment;
                                                    StiRtfCell pendingCell = (StiRtfCell)pendingCells[cells.Count];
                                                    currentCell.Pictures = pendingCell.Pictures;
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
                                                        if (contentPictures != null)
                                                        {
                                                            currentCell.Pictures = contentPictures;
                                                            contentPictures = null;
                                                        }
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
                                                string plainText = res.ToString();
                                                beginState = GetMainState(beginState, formats, plainText);
                                                string html = GetHtmlText(plainText, formats, ref beginState,
                                                    fontTable, colorTable);
                                                formats.Clear();
                                                SetText(currentCell.Content, html ?? plainText);
                                                currentCell.Content.AllowHtmlTags = html != null;
                                                FontStyle fs = FontStyle.Regular;
                                                if (beginState.Bold) fs |= FontStyle.Bold;
                                                if (beginState.Italic) fs |= FontStyle.Italic;
                                                if (beginState.Underline) fs |= FontStyle.Underline;
                                                var fontData = GetFontData(beginState, fontTable);
                                                if (fontData.Bold) fs |= FontStyle.Bold;
                                                if (fontData.Italic) fs |= FontStyle.Italic;
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
                                if (headerRow && headerLines == table.Count) headerLines++;
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
                        if (previousCells != null) lastRowCells = previousCells;
                        #endregion

                        #region create new band
                        if (tablePageBand != null || !CanSplitHeader(table, headerLines)) headerLines = 0;
                        StiBand headerBand = headerLines > 0 ? AddHeaderBand(currentPage) : null;
                        StiBand band = GetContentBand(currentPage, tablePageBand);
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
                        //the cells which are merged in a row: the first one takes the width of the others
                        for (int indexLine = 0; indexLine < table.Count; indexLine++)
                        {
                            StiRtfCell firstCell = null;
                            for (int indexCell = 0; indexCell < cellsCount; indexCell++)
                            {
                                StiRtfCell cell = matrix[indexLine, indexCell];
                                if (cell == null) continue;

                                if (cell.JoinedToLeft && firstCell != null)
                                {
                                    JoinCell(firstCell, cell);
                                }
                                else
                                {
                                    firstCell = cell;
                                }
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
                                    cell.Content.Parent = indexLine < headerLines ? headerBand : band;
                                    cell.Content.Page = currentPage;
                                    cell.Content.CanGrow = true;
                                    cell.Content.WordWrap = true;
                                    //a cell higher than the page is continued on the next one; without it
                                    //the whole row moves to the next page and leaves an empty one
                                    cell.Content.CanBreak = true;
                                    if (cell.MergedCellsCount == 0)
                                    {
                                        SizeD cellSize = cell.Content.GetActualSize();
                                        if (cellSize.Height > cell.Content.Height) cell.Content.Height = cellSize.Height;
                                    }
                                    double picturesHeight = GetPicturesHeight(cell);
                                    if (picturesHeight > cell.Content.Height) cell.Content.Height = picturesHeight;
                                    if (cell.Content.Height > linesHeight[indexLine]) linesHeight[indexLine] = cell.Content.Height;
                                }
                            }
                        }
                        #endregion

                        #region add components to band
                        double cellPosY = bandTop;
                        for (int indexLine = 0; indexLine < table.Count; indexLine++)
                        {
                            //the rows after the repeated ones start the data band
                            if (headerBand != null && indexLine == headerLines)
                            {
                                headerBand.Height = cellPosY;
                                cellPosY = bandTop;
                            }
                            StiBand lineBand = indexLine < headerLines ? headerBand : band;
                            double cellPosX = GetRowLeft(matrix, indexLine, cellsCount, currentPage.Width);
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
                                        cell.Content.Parent = lineBand;
                                        cell.Content.Border = GetBorder(cell.Sides);
                                        StiText text = cell.Content;
                                        lineBand.Components.Add(text);
                                        AddCellPictures(lineBand, cell, currentPage);
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
