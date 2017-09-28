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

namespace Stimulsoft.Report.Import
{
	public class StiRtfNodeCollection : CollectionBase
    {
        #region Properties
        public StiRtfTreeNode this[int index]
        {
            get
            {
                return (StiRtfTreeNode)InnerList[index];
            }
            set
            {
                InnerList[index] = value;
            }
        }
        #endregion

        #region Methods
        public int Add(StiRtfTreeNode node)
		{
			InnerList.Add(node);
			return (InnerList.Count - 1);
		}

		public void Insert(int index, StiRtfTreeNode node)
		{
			InnerList.Insert(index, node);
		}

        public int IndexOf(StiRtfTreeNode node)
        {
            return InnerList.IndexOf(node);
        }

		public void AddRange(StiRtfNodeCollection collection)
		{
			InnerList.AddRange(collection);
		}

		public void RemoveRange(int index, int count)
		{
			InnerList.RemoveRange(index, count);
        }
        #endregion
    }
}
