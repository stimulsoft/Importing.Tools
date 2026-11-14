#region Copyright (C) 2003-2026 Stimulsoft
/*
{*******************************************************************}
{																	}
{	Stimulsoft Reports												}
{	                         										}
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
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// The named palettes of the charts of Reporting Services. The colors are taken from the charts rendered by
    /// the report viewer of Microsoft, a column per series over a white and a black background, so the alpha of
    /// the semitransparent palettes is known too. They match the palettes of Microsoft Chart Controls, except
    /// the palettes Pacific of Reporting Services 2016: Pacific has twenty colors, the ten own ones and the ten
    /// of PacificLight. Default is the palette Bright of Microsoft Chart Controls.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Fields.Palette
        private static readonly uint[] PacificColors =
        {
            0xFF01B8AA, 0xFF374649, 0xFFFD625E, 0xFFF2C80F, 0xFF5F6B6D, 0xFF8AD4EB, 0xFFFE9666, 0xFFA66999,
            0xFF3599B8, 0xFFDFBFBF, 0xFF4AC5BB, 0xFF5F6B6D, 0xFFFB8281, 0xFFF4D25A, 0xFF7F898A, 0xFFA4DDEE,
            0xFFFDAB89, 0xFFB687AC, 0xFF28738A, 0xFFA78F8F
        };

        private static readonly Dictionary<string, uint[]> ChartPalettes =
            new Dictionary<string, uint[]>(StringComparer.OrdinalIgnoreCase)
        {
            ["Default"] = new uint[]
            {
                0xFF008000, 0xFF0000FF, 0xFF800080, 0xFF00FF00, 0xFFFF00FF, 0xFF008080, 0xFFFFFF00, 0xFF808080,
                0xFF00FFFF, 0xFF000080, 0xFF800000, 0xFFFF0000, 0xFF808000, 0xFFC0C0C0, 0xFFFF6347, 0xFFFFE4B5
            },
            ["BrightPastel"] = new uint[]
            {
                0xFF418CF0, 0xFFFCB441, 0xFFE0400A, 0xFF056492, 0xFFBFBFBF, 0xFF1A3B69, 0xFFFFE382, 0xFF129CDD,
                0xFFCA6B4B, 0xFF005CDB, 0xFFF3D288, 0xFF506381, 0xFFF1B9A8, 0xFFE0830A, 0xFF7893BE
            },
            ["EarthTones"] = new uint[]
            {
                0xFFFF8000, 0xFFB8860B, 0xFFC04000, 0xFF6B8E23, 0xFFCD853F, 0xFFC0C000, 0xFF228B22, 0xFFD2691E,
                0xFF808000, 0xFF20B2AA, 0xFFF4A460, 0xFF00C000, 0xFF8FBC8B, 0xFFB22222, 0xFF8B4513, 0xFFC00000
            },
            ["Excel"] = new uint[]
            {
                0xFF9999FF, 0xFF993366, 0xFFFFFFCC, 0xFFCCFFFF, 0xFF660066, 0xFFFF8080, 0xFF0066CC, 0xFFCCCCFF,
                0xFF000080, 0xFFFF00FF, 0xFFFFFF00, 0xFF00FFFF, 0xFF800080, 0xFF800000, 0xFF008080, 0xFF0000FF
            },
            ["GrayScale"] = new uint[]
            {
                0xFFC8C8C8, 0xFFBDBDBD, 0xFFB2B2B2, 0xFFA7A7A7, 0xFF9C9C9C, 0xFF919191, 0xFF868686, 0xFF7B7B7B,
                0xFF707070, 0xFF656565, 0xFF5A5A5A, 0xFF4F4F4F, 0xFF444444, 0xFF393939, 0xFF2E2E2E, 0xFF232323
            },
            ["Light"] = new uint[]
            {
                0xFFE6E6FA, 0xFFFFF0F5, 0xFFFFDAB9, 0xFFFFFACD, 0xFFFFE4E1, 0xFFF0FFF0, 0xFFF0F8FF, 0xFFF5F5F5,
                0xFFFAEBD7, 0xFFE0FFFF
            },
            ["Pastel"] = new uint[]
            {
                0xFF87CEEB, 0xFF32CD32, 0xFFBA55D3, 0xFFF08080, 0xFF4682B4, 0xFF9ACD32, 0xFF40E0D0, 0xFFFF69B4,
                0xFFF0E68C, 0xFFD2B48C, 0xFF8FBC8B, 0xFF6495ED, 0xFFDDA0DD, 0xFF5F9EA0, 0xFFFFDAB9, 0xFFFFA07A
            },
            ["SemiTransparent"] = new uint[]
            {
                0x96FF0000, 0x9600FF00, 0x960000FF, 0x96FFFF00, 0x9600FFFF, 0x96FF00FF, 0x96AA7814, 0x50FF0000,
                0x5000FF00, 0x500000FF, 0x50FFFF00, 0x5000FFFF, 0x50FF00FF, 0x50AA7814, 0x96647832, 0x96285A96
            },
            ["Berry"] = new uint[]
            {
                0xFF8A2BE2, 0xFFBA55D3, 0xFF4169E1, 0xFFC71585, 0xFF0000FF, 0xFF8A2BE2, 0xFFDA70D6, 0xFF7B68EE,
                0xFFC000C0, 0xFF0000CD, 0xFF800080
            },
            ["Chocolate"] = new uint[]
            {
                0xFFA0522D, 0xFFD2691E, 0xFF8B0000, 0xFFCD853F, 0xFFA52A2A, 0xFFF4A460, 0xFF8B4513, 0xFFC04000,
                0xFFB22222, 0xFFB65C3A
            },
            ["Fire"] = new uint[]
            {
                0xFFFFD700, 0xFFFF0000, 0xFFFF1493, 0xFFDC143C, 0xFFFF8C00, 0xFFFF00FF, 0xFFFFFF00, 0xFFFF4500,
                0xFFC71585, 0xFFDDE221
            },
            ["SeaGreen"] = new uint[]
            {
                0xFF2E8B57, 0xFF66CDAA, 0xFF4682B4, 0xFF008B8B, 0xFF5F9EA0, 0xFF3CB371, 0xFF48D1CC, 0xFFB0C4DE,
                0xFF8FBC8B, 0xFF87CEEB
            },
            ["Pacific"] = PacificColors,
            ["PacificLight"] = PacificColors.Skip(10).ToArray(),

            //the colors of Pacific with the alpha 153
            ["PacificSemiTransparent"] = PacificColors.Select(color => color & 0x00FFFFFF | 0x99000000).ToArray()
        };
        #endregion

        #region Methods.Palette
        /// <summary>
        /// Returns the colors of a named palette of Reporting Services, or null for an unknown name.
        /// </summary>
        private static Color[] GetChartPaletteColors(string palette)
        {
            if (palette == null || !ChartPalettes.TryGetValue(palette.Trim(), out var colors))
                return null;

            return colors.Select(color => Color.FromArgb(unchecked((int)color))).ToArray();
        }
        #endregion
    }
}
