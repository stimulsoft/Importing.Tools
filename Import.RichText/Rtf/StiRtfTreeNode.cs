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
using System.Collections;

namespace Stimulsoft.Report.Import
{
	#region Enum StiRtfNodeType
	public enum StiRtfNodeType
	{
		Root = 0,
		Keyword = 1,
		Control = 2,
		Text = 3,
		Group = 4,
		GroupEnd = 5,
		None = 6
	}
	#endregion

	/// <summary>
	/// Summary description for RtfTreeNode.
	/// </summary>
	public class StiRtfTreeNode
    {
        #region Variables
        private StiRtfNodeType type;
		private string key;
		private bool hasParam;
		private int param;
		private StiRtfNodeCollection children;
		private StiRtfTreeNode parent;
        private int index;
        #endregion

        #region Properties
        public StiRtfTreeNode ParentNode
		{
			get
			{
				return parent;
			}
			set
			{
				parent = value;
			}
		}

		public StiRtfNodeType NodeType
		{
			get
			{
				return type;
			}
			set
			{
				type = value;
			}
		}

		public string NodeKey
		{
			get
			{
				return key;
			}
			set
			{
				key = value;
			}
		}

		public bool HasParameter
		{
			get
			{
				return hasParam;
			}
			set
			{
				hasParam = value;
			}
		}

		public int Parameter
		{
			get
			{
				return param;
			}
			set
			{
				param = value;
			}
		}

        public int Index
        {
            get
            {
                return index;
            }
            set
            {
                index = value;
            }
        }

		public StiRtfNodeCollection ChildNodes
		{
			get
			{
				return children;
			}
        }
        #endregion

        #region Methods
        public void AppendChild(StiRtfTreeNode newNode)
		{
			if(newNode != null)
			{
				newNode.parent = this;
				children.Add(newNode);
			}
		}

		public StiRtfTreeNode LastChild
		{
			get
			{
				if (children.Count > 0)
					return children[children.Count - 1];
				else
					return null;
			}
		}

        internal void UpdateNodeIndex(bool recursive)
        {
            for (int index = 0; index < children.Count; index++)
            {
                children[index].index = index;
                if (recursive)
                {
                    children[index].UpdateNodeIndex(true);
                }
            }
        }

		public StiRtfTreeNode NextToken
		{
			get
			{
				//int currentIndex = parent.children.IndexOf(this);
                int currentIndex = index;
				if (currentIndex < parent.children.Count - 1)
				{
					return parent.children[currentIndex + 1];
				}
				else
				{
					return null;
				}
			}
		}

		public StiRtfTreeNode PreviousToken
		{
			get
			{
				//int currentIndex = parent.children.IndexOf(this);
                int currentIndex = index;
                if (currentIndex > 0)
				{
					return parent.children[currentIndex - 1];
				}
				else
				{
					return null;
				}
			}
		}

        public StiRtfTreeNode NextFlateToken
        {
            get
            {
                StiRtfTreeNode node = this;
                if (node.NodeType == StiRtfNodeType.Group)
                {
                    node = node.ChildNodes[0];
                }
                else
                {
                    if (node.NodeType == StiRtfNodeType.GroupEnd)
                    {
                        node = this.parent;
                        if (node.NodeType == StiRtfNodeType.Root)
                        {
                            node = null;
                        }
                        else
                        {
                            node = node.NextToken;
                        }
                    }
                    else
                    {
                        if (node.parent.NodeType == StiRtfNodeType.Root)
                        {
                            node = null;
                        }
                        else
                        {
                            //int currentIndex = node.parent.children.IndexOf(node);
                            int currentIndex = node.index;

                            node = node.parent.children[currentIndex + 1];
                        }
                    }
                }
                return node;
            }
        }

        public void RemoveThisToken()
        {
            //int currentIndex = parent.children.IndexOf(this);
            int currentIndex = index;
            parent.children.RemoveAt(currentIndex);
            parent.UpdateNodeIndex(false); 
            this.parent = null;
        }
        #endregion

        public StiRtfTreeNode(StiRtfNodeType nodeType, string key, bool hasParameter, int parameter)
        {
            this.children = new StiRtfNodeCollection();
            this.type = nodeType;
            this.key = key;
            this.hasParam = hasParameter;
            this.param = parameter;
        }

        internal StiRtfTreeNode(StiRtfToken token)
        {
            this.children = new StiRtfNodeCollection();
            this.type = (StiRtfNodeType)token.Type;
            this.key = token.Key;
            this.hasParam = token.HasParameter;
            this.param = token.Parameter;
        }
	}

}
