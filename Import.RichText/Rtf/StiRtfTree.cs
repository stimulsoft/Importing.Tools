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

            rootNode.UpdateNodeIndex(true);

			if (level != 0) res = -1;
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
            //find last group in header or info
            while (pos < main.Count)
            {
                if (main[pos].NodeType == StiRtfNodeType.Group)
                {
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
                        StiRtfTreeNode tempNode = main[pos2].ChildNodes[0];
                        if (((tempNode.NodeType == StiRtfNodeType.Keyword) &&
                            (TokensHash.ContainsKey(tempNode.NodeKey) && (int)TokensHash[tempNode.NodeKey] == 0)) ||
                            ((tempNode.NodeType == StiRtfNodeType.Control) && (tempNode.NodeKey == "*")))
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
					if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "page") && (node != data.start)) break;
					if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "sect") && (node != data.start)) break;
					if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "trowd"))
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
				}
				if (node != data.start)
				{
					data.end = lastNode;
					data.isTableRow = false;
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
							node = node.NextFlateToken;
							break;
						}
                        //if ((node.NodeType == StiRtfNodeType.Keyword) && (node.NodeKey == "trowd")) break;
						lastNode = node;
						node = node.NextFlateToken;
					}
					while ((node != null)&& (node != lastRootNode) && ((node.NodeType == StiRtfNodeType.Group) || (node.NodeType == StiRtfNodeType.GroupEnd)))
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
