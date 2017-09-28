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

namespace Stimulsoft.Report.Import
{
	#region Enum StiRtfTokenType
	public enum StiRtfTokenType
	{
		None = 0,
		Keyword = 1,
		Control = 2,
		Text = 3,
		Eof = 4,
		GroupStart = 5,
		GroupEnd = 6
	}
	#endregion

	/// <summary>
	/// Summary description for StiRtfToken.
	/// </summary>
	public class StiRtfToken
    {
        #region Variables
        private StiRtfTokenType type;
		private string key;
		private bool hasParam;
		private int param;
        #endregion

        #region Properties
        public StiRtfTokenType Type
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

		public string Key
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
        #endregion

        public StiRtfToken()
		{
			type = StiRtfTokenType.None;
			key = "";
		}
	}

}
