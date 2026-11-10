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
using System.Drawing;

namespace Import.CrystalReports.Formulas
{
    /// <summary>
    /// The numbers of the constants of the formulas of conditional formatting: the colors, the font styles and
    /// DefaultAttribute. A color is red + green * 256 + blue * 65536, like the Color function of Crystal.
    /// </summary>
    public static class StiCrystalFormatConstants
    {
        #region Consts
        public const int DefaultAttribute = -1;

        public const int NoColor = -2;

        public const int Regular = 0;

        public const int Italic = 1;

        public const int Bold = 2;

        public const int BoldItalic = 3;
        #endregion

        #region Fields
        private static readonly Dictionary<string, int> constants = new Dictionary<string, int>
        {
            { "defaultattribute", DefaultAttribute },
            { "crnocolor", NoColor },
            { "crblack", 0x000000 },
            { "crmaroon", 0x000080 },
            { "crgreen", 0x008000 },
            { "crolive", 0x008080 },
            { "crnavy", 0x800000 },
            { "crpurple", 0x800080 },
            { "crteal", 0x808000 },
            { "crgray", 0x808080 },
            { "crsilver", 0xC0C0C0 },
            { "crred", 0x0000FF },
            { "crlime", 0x00FF00 },
            { "cryellow", 0x00FFFF },
            { "crblue", 0xFF0000 },
            { "crfuchsia", 0xFF00FF },
            { "craqua", 0xFFFF00 },
            { "crwhite", 0xFFFFFF },
            { "crregular", Regular },
            { "critalic", Italic },
            { "crbold", Bold },
            { "crbolditalic", BoldItalic }
        };
        private static readonly HashSet<string> fontStyles = new HashSet<string>
        {
            "regular",
            "italic",
            "bold",
            "bolditalic"
        };
        #endregion

        #region Methods
        /// <summary>
        /// Returns the number of a constant by its name in lower case. Crystal knows the colors without the prefix
        /// too, from its old versions: "if ... then Lime else White", about 130 formulas of the corpus.
        /// </summary>
        public static bool TryGetValue(string name, out int value)
        {
            if (constants.TryGetValue(name, out value))
                return true;

            return !name.StartsWith("cr") && !fontStyles.Contains(name) && constants.TryGetValue("cr" + name, out value);
        }

        /// <summary>
        /// Converts the number of a color of a formula to a color; NoColor gives the transparent color.
        /// </summary>
        public static Color ToColor(int value)
        {
            if (value == NoColor)
                return Color.Transparent;

            return Color.FromArgb(value & 0xFF, (value >> 8) & 0xFF, (value >> 16) & 0xFF);
        }

        public static FontStyle ToFontStyle(int value)
        {
            switch (value)
            {
                case Italic:
                    return FontStyle.Italic;

                case Bold:
                    return FontStyle.Bold;

                case BoldItalic:
                    return FontStyle.Bold | FontStyle.Italic;

                default:
                    return FontStyle.Regular;
            }
        }
        #endregion
    }
}
