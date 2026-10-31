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

using System.Collections.Generic;
using System.Linq;

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// A node of the syntax tree of a Crystal formula. The meaning of the text and of the children depends on the kind:
    /// Binary and Unary keep the operator, Call keeps the name of the function, If keeps the condition, the branch
    /// Then and the optional branch Else, Range keeps the kind of the range (to, _to, upfrom and so on),
    /// Declare keeps the name of the variable and the optional initial value.
    /// </summary>
    public class StiCrystalNode
    {
        #region Properties
        public StiCrystalNodeKind Kind { get; }

        public string Text { get; }

        public List<StiCrystalNode> Children { get; } = new List<StiCrystalNode>();

        public int Position { get; }

        /// <summary>
        /// The scope of a declared variable: Local, Global or Shared.
        /// </summary>
        public string Scope { get; set; }

        /// <summary>
        /// The type of a declared variable as it is written in the formula, for example NumberVar.
        /// </summary>
        public string VariableType { get; set; }

        /// <summary>
        /// True for a declared variable of an array or range type.
        /// </summary>
        public bool IsArrayVariable { get; set; }

        public bool IsRangeVariable { get; set; }
        #endregion

        #region Methods
        public StiCrystalNode Add(StiCrystalNode child)
        {
            Children.Add(child);
            return this;
        }

        public override string ToString()
        {
            if (Children.Count == 0)
                return $"{Kind}:{Text}";

            return $"{Kind}:{Text}({string.Join(", ", Children.Select(c => c?.ToString() ?? "null"))})";
        }
        #endregion

        public StiCrystalNode(StiCrystalNodeKind kind, string text, int position, params StiCrystalNode[] children)
        {
            Kind = kind;
            Text = text;
            Position = position;
            Children.AddRange(children);
        }
    }
}
