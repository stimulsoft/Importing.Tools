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

using System;
using System.IO;
using System.Text;
using System.Collections;

namespace Stimulsoft.Report.Import
{
	/// <summary>
	/// Summary description for StiRtfTree.
	/// </summary>
	public class StiRtfTree
    {
        #region Variables
        internal StiRtfTreeNode rootNode = null;
		private TextReader rtf;
		private StiRtfLex lex;
		private bool mergeSpecialCharacters;
        #endregion

        #region Structures
        public class StiRtfBandData
		{
			public StiRtfTreeNode start = null;
			public StiRtfTreeNode end = null;
			public bool isTableRow = false;
			public bool isIntbl = false;
			//the keyword of the header or footer group which holds the band ("header", "footerr"), null for the body
			public string pageBand = null;
        }
        #endregion

        #region Static properties
        private static Hashtable tokens = null;
        public static Hashtable TokensHash
        {
            get
            {
                if (tokens == null)
                {
                    tokens = new Hashtable();
                    //header and info
                    tokens["fonttbl"] = 0;
                    tokens["filetbl"] = 0;
                    tokens["colortbl"] = 0;
                    tokens["stylesheet"] = 0;
                    tokens["listtables"] = 0;
                    tokens["revtbl"] = 0;
                    tokens["rsidtable"] = 0;
                    tokens["generator"] = 0;
                    tokens["info"] = 0;
                    //bands
                    tokens["paperh"] = 1;
                    tokens["paperw"] = 1;
                    tokens["margl"] = 1;
                    tokens["margr"] = 1;
                    tokens["margt"] = 1;
                    tokens["margb"] = 1;
                    tokens["sectd"] = 1;
                    tokens["trowd"] = 1;
                    tokens["pard"] = 1;
                    tokens["plain"] = 1;
                    tokens["viewkind"] = 1;
                    tokens["nowidctlpar"] = 1;
                    tokens["widowctrl"] = 1;
                    tokens["headery"] = 1;
                    tokens["footery"] = 1;
                }
                return tokens;
            }
        }
        #endregion

        #region Methods.LoadRtfText
        public int LoadRtfText(string text)
		{
			int res = 0;
			rtf = new StringReader(text);
			lex = new StiRtfLex(rtf);
			res = ParseRtfTree();
			rtf.Close();
			return res;
        }

        private int ParseRtfTree()
		{
			int res = 0;
			int level = 0;
            int tokensCount = 0;
            int delayCount = 0;

			Encoding encoding = Encoding.Default;

			StiRtfTreeNode curNode = rootNode;
			StiRtfTreeNode newNode = null;

			StiRtfToken tok = lex.NextToken();

			while (tok.Type != StiRtfTokenType.Eof)
			{
                tokensCount++;
				switch (tok.Type)
				{
					case StiRtfTokenType.GroupStart:
						newNode = new StiRtfTreeNode(StiRtfNodeType.Group,"GROUP",false,0);
						curNode.AppendChild(newNode);
						curNode = newNode;
						level++;
						break;
					case StiRtfTokenType.GroupEnd:
						//an extra closing brace of a damaged file is skipped
						if (level == 0)
						{
							res = -1;
							break;
						}
						newNode = new StiRtfTreeNode(StiRtfNodeType.GroupEnd, "GROUPEND", false, 0);
						curNode.AppendChild(newNode);
						curNode = curNode.ParentNode;
                        //newNode = new StiRtfTreeNode(StiRtfNodeType.None, "NoneAfterGroup", false, 0);
                        //curNode.AppendChild(newNode);
						level--;
                        //show progress
                        if (delayCount == 0)
                        {
                            delayCount = 8;
                        }
                        else
                        {
                            delayCount--;
                        }
						break;
					case StiRtfTokenType.Keyword:
					case StiRtfTokenType.Control:
					case StiRtfTokenType.Text:
						if (mergeSpecialCharacters)
						{
							bool isText = tok.Type == StiRtfTokenType.Text || (tok.Type == StiRtfTokenType.Control && tok.Key == "'");
							if (curNode.LastChild != null && (curNode.LastChild.NodeType == StiRtfNodeType.Text && isText))
							{
								if (tok.Type == StiRtfTokenType.Text)
								{
									curNode.LastChild.NodeKey += tok.Key;
									break;
								}
								if (tok.Type == StiRtfTokenType.Control && tok.Key == "'")
								{
									curNode.LastChild.NodeKey += DecodeControlChar(tok.Parameter, encoding);
									break;
								}
							}
							else
							{
								if (tok.Type == StiRtfTokenType.Control && tok.Key == "'")
								{
									newNode = new StiRtfTreeNode(StiRtfNodeType.Text, DecodeControlChar(tok.Parameter, encoding), false, 0);
									curNode.AppendChild(newNode);
									break;
								}
							}
						}

						newNode = new StiRtfTreeNode(tok);
						curNode.AppendChild(newNode);

						if (mergeSpecialCharacters)
						{
							if (level == 1 && newNode.NodeType == StiRtfNodeType.Keyword && newNode.NodeKey == "ansicpg")
							{
								encoding = Encoding.GetEncoding(newNode.Parameter);
							}
						}

						break;
					default:
						res = -1;
						break;
				}
				tok = lex.NextToken();
			}

            //the groups of a truncated file are closed, the import relies on the closing nodes
            while (level > 0)
            {
                curNode.AppendChild(new StiRtfTreeNode(StiRtfNodeType.GroupEnd, "GROUPEND", false, 0));
                curNode = curNode.ParentNode;
                level--;
                res = -1;
            }

            rootNode.UpdateNodeIndex(true);
			return res;
        }

        private static string DecodeControlChar(int code, Encoding enc)
		{
			return enc.GetString(new byte[] {(byte)code});
        }
        #endregion

        #region Methods.ToStringEx
        public string ToStringEx()
        {
            return toStringInm(rootNode, 0, true);
        }

        private string toStringInm(StiRtfTreeNode curNode, int level, bool showNodeTypes)
        {
            StringBuilder res = new StringBuilder();
            StiRtfNodeCollection children = curNode.ChildNodes;
            for (int i = 0; i < level; i++)
                res.Append("  ");
            if (curNode.NodeType == StiRtfNodeType.Root)
                res.Append("ROOT\r\n");
            else if (curNode.NodeType == StiRtfNodeType.Group)
                res.Append("GROUP\r\n");
            else
            {
                if (showNodeTypes)
                {
                    res.Append(curNode.NodeType);
                    res.Append(": ");
                }
                res.Append(curNode.NodeKey);
                if (curNode.HasParameter)
                {
                    res.Append(" ");
                    res.Append(Convert.ToString(curNode.Parameter));
                }
                res.Append("\r\n");
            }
            foreach (StiRtfTreeNode node in children)
            {
                res.Append(toStringInm(node, level + 1, showNodeTypes));
            }
            return res.ToString();
        }
        #endregion

        #region Methods.ToText
        public string ToText()
        {
            StringBuilder res = new StringBuilder();
            //int headerLen = GetHeaderLength();
            int headerLen = 1;
            StiRtfTreeNode node = rootNode.ChildNodes[0].ChildNodes[headerLen];
            while (node != null)
            {
                if (node.NodeType == StiRtfNodeType.Text)
                {
                    res.Append(node.NodeKey);
                }
                if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "par"))
                {
                    res.Append("\r\n");
                }
                node = node.NextFlateToken;
            }
            return res.ToString();
        }
        #endregion

        #region Methods.GetBandsData
        //The starred group "nesttableprops" is the definition of a row of a nested table.
        internal static bool IsNestedRowDefinition(StiRtfTreeNode group)
        {
            if (group.NodeType != StiRtfNodeType.Group || group.ChildNodes.Count < 2) return false;

            StiRtfTreeNode first = group.ChildNodes[0];
            StiRtfTreeNode second = group.ChildNodes[1];
            return first.NodeType == StiRtfNodeType.Control && first.NodeKey == "*" &&
                second.NodeType == StiRtfNodeType.Keyword && second.NodeKey == "nesttableprops";
        }

        //The group of a header or of a footer of the page, with all their kinds: for the left, right and first pages.
        private static bool IsPageBandGroup(StiRtfTreeNode group)
        {
            if (group == null || group.NodeType != StiRtfNodeType.Group) return false;

            StiRtfTreeNode first = group.ChildNodes[0];
            if (first.NodeType != StiRtfNodeType.Keyword) return false;

            switch (first.NodeKey)
            {
                case "header":
                case "headerl":
                case "headerr":
                case "headerf":
                case "footer":
                case "footerl":
                case "footerr":
                case "footerf":
                    return true;

                default:
                    return false;
            }
        }

        private static string GetPageBand(StiRtfTreeNode node)
        {
            while (node != null)
            {
                if (IsPageBandGroup(node)) return node.ChildNodes[0].NodeKey;
                node = node.ParentNode;
            }
            return null;
        }

        private static bool IsInNestedRowDefinition(StiRtfTreeNode node)
        {
            StiRtfTreeNode parent = node.ParentNode;
            while (parent != null)
            {
                if (IsNestedRowDefinition(parent)) return true;
                parent = parent.ParentNode;
            }
            return false;
        }
        private static bool IsHeaderGroup(StiRtfTreeNode group)
        {
            StiRtfTreeNode node = group.ChildNodes[0];
            if (node.NodeType == StiRtfNodeType.Control) return node.NodeKey == "*";

            return node.NodeType == StiRtfNodeType.Keyword &&
                TokensHash.ContainsKey(node.NodeKey) && (int)TokensHash[node.NodeKey] == 0;
        }

        //The tokens which make a text band of their own: a text, the end of a paragraph, a picture.
        private static bool IsContentToken(StiRtfTreeNode node)
        {
            if (node.NodeType == StiRtfNodeType.Text) return node.NodeKey.Trim().Length > 0;
            if (node.NodeType == StiRtfNodeType.Control) return node.NodeKey == "'";
            if (node.NodeType != StiRtfNodeType.Keyword) return false;

            switch (node.NodeKey)
            {
                case "par":
                case "page":
                case "sect":
                case "line":
                case "tab":
                case "u":
                case "pict":
                case "object":
                case "field":
                    return true;

                default:
                    return false;
            }
        }

        //Writers put keywords between the end of a row and the definition of the next one: "pard", "itap0",
        //"ltrrow". They are not a text between two tables: as a band of their own they cut the table into
        //rows, and the cells which are merged in a column lose each other.
        private static bool IsRowGap(ArrayList bandsList, StiRtfTreeNode start, StiRtfTreeNode end)
        {
            if (bandsList.Count == 0) return false;

            StiRtfBandData previous = (StiRtfBandData)bandsList[bandsList.Count - 1];
            if (!previous.isTableRow || previous.pageBand != GetPageBand(end)) return false;

            StiRtfTreeNode node = start;
            while (node != null)
            {
                if (IsContentToken(node)) return false;
                if (node == end) return true;

                node = node.NextFlateToken;
            }
            return false;
        }

        public ArrayList GetBandsData()
		{
			StiRtfNodeCollection main = rootNode.ChildNodes[0].ChildNodes;
			StiRtfTreeNode lastRootNode = rootNode.ChildNodes[0].LastChild;

            #region Skip header and info
            int pos = 0;
            //find first group
			while ((pos < main.Count) && (main[pos].NodeType != StiRtfNodeType.Group))
			{
				pos++;
			}
            //a file without groups has no header, its text starts at once
            if (pos >= main.Count) pos = 0;
            //find last group in header or info
            while (pos < main.Count)
            {
                if (main[pos].NodeType == StiRtfNodeType.Group)
                {
                    //a group with the text of the document ends the header
                    if (!IsHeaderGroup(main[pos])) break;
                    pos++;
                }
                else
                {
                    //find next group
                    int pos2 = pos;
                    while ((pos2 < main.Count) && (main[pos2].NodeType != StiRtfNodeType.Group))
                    {
                        if ((main[pos2].NodeType == StiRtfNodeType.Keyword) &&
                            (TokensHash.ContainsKey(main[pos2].NodeKey)) && ((int)TokensHash[main[pos2].NodeKey] == 1))
                        {
                            pos = pos2;
                            pos2 = main.Count;  //for break loop
                            break;
                        }
                        pos2++;
                    }
                    //check for header or info
                    if (pos2 < main.Count)
                    {
                        if (IsHeaderGroup(main[pos2]))
                        {
                            pos = pos2;
                            pos++;
                            continue;
                        }
                    }
                    break;
                }
            }
            #endregion

			ArrayList bandsList = new ArrayList();
			if (pos >= main.Count) return bandsList;

			StiRtfTreeNode node = main[pos];
			StiRtfTreeNode lastNode = null;
			StiRtfBandData data = null;

			while ((node != null) && (node != lastRootNode))
			{
				data = new StiRtfBandData();
				data.start = node;
				bool tableFounded = false;

				while ((node != null) && (node != lastRootNode))
				{
					//a header or a footer is a band of its own: the text before it ends here
					if (IsPageBandGroup(node) && (node != data.start)) break;
					if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "page") && (node != data.start)) break;
					if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "sect") && (node != data.start)) break;
					//the "trowd" of a nested table stands inside a row of the outer table and doesn't start a new one
					if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "trowd") &&
						!IsInNestedRowDefinition(node))
					{
						tableFounded = true;
						break;
					}
					if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "intbl"))
					{
						node = data.start;
						tableFounded = true;
						break;
					}
						lastNode = node;
						node = node.NextFlateToken;
						//and the text after it starts a new band
						bool isGroupEnd = lastNode.NodeType == StiRtfNodeType.GroupEnd;
						if (isGroupEnd && IsPageBandGroup(lastNode.ParentNode)) break;
				}
				if (tableFounded && (node != data.start) && IsRowGap(bandsList, data.start, lastNode))
				{
					//the keywords between two rows stay with the first of them, the table goes on
					((StiRtfBandData)bandsList[bandsList.Count - 1]).end = lastNode;
				}
				else if (node != data.start)
				{
						data.end = lastNode;
						data.isTableRow = false;
						data.pageBand = GetPageBand(lastNode);
						bandsList.Add(data);
				}

				if ((tableFounded) && (node != null) && (node != lastRootNode))
				{
					data = new StiRtfBandData();
					data.start = node;
					data.isTableRow = true;
					if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "trowd"))
					{
						node = node.NextFlateToken;
					}
					else
					{
						data.isIntbl = true;
					}

					while ((node != null) && (node != lastRootNode))
					{
						if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "row"))
						{
							data.pageBand = GetPageBand(node);
							node = node.NextFlateToken;
							break;
						}
                        //if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "trowd")) break;
						lastNode = node;
						node = node.NextFlateToken;
					}
					while ((node != null) && (node != lastRootNode) && !IsPageBandGroup(node) &&
						((node.NodeType == StiRtfNodeType.Group) || (node.NodeType == StiRtfNodeType.GroupEnd)))
					{
						lastNode = node;
						node = node.NextFlateToken;
					}
					data.end = lastNode;
                    //if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "trowd"))
                    //{
                    //    data.end = node;	//
                    //}
					bandsList.Add(data);
                }
			}

			return bandsList;
        }
        #endregion

        public StiRtfTree()
        {
            rootNode = new StiRtfTreeNode(StiRtfNodeType.Root, "ROOT", false, 0);
            mergeSpecialCharacters = false;
        }

    }
}
