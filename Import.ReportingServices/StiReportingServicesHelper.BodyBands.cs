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

using Stimulsoft.Base.Drawing;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Import.ReportingServices;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Puts the bands of the data regions of the body on the page. The body of Reporting Services is a container,
    /// its items are placed freely and a data region below a growing one moves down; the import puts the body on
    /// one band and every data region on a panel with its own bands. Such a report is hard to edit, so the body
    /// is cut into strips that no item crosses: a strip with one data region gives its bands to the page, the
    /// other strips go to the report title, to the report summary or to a band that prints once. A strip with
    /// several data regions side by side keeps its panels. A rectangle around a data region is dissolved first,
    /// when it has no look of its own.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Consts.BodyBands
        //the items of Reporting Services often overlap a little, such an overlap does not join them in one strip
        private const double BodyOverlapTolerance = 1;
        #endregion

        #region Fields.BodyBands
        /// <summary>
        /// The bands of the bodies: the band of a body and the parts of it cut at the page breaks.
        /// </summary>
        private List<StiDataBand> bodyBands = new List<StiDataBand>();

        /// <summary>
        /// The rectangles that keep their items together on a page. A panel cannot break by default, so its
        /// CanBreak does not tell such a rectangle from the others.
        /// </summary>
        private HashSet<StiPanel> keptTogetherRectangles = new HashSet<StiPanel>();
        #endregion

        #region Properties.BodyBands
        /// <summary>
        /// How the items of the body are put on the page by the import, the bands of the data regions on the
        /// page as a rule.
        /// </summary>
        public static StiReportingServicesBodyLayout DefaultBodyLayout { get; set; }

        /// <summary>
        /// How the items of the body of this report are put on the page.
        /// </summary>
        internal StiReportingServicesBodyLayout BodyLayout { get; set; } = DefaultBodyLayout;
        #endregion

        #region Methods.BodyBands
        /// <summary>
        /// Puts the bands of the data regions of the bodies imported so far on their pages. It goes after all
        /// the passes that look for the bands inside the panel of a data region.
        /// </summary>
        private void ApplyBodyLayout()
        {
            if (BodyLayout == StiReportingServicesBodyLayout.Bands)
            {
                foreach (var group in bodyBands.Where(b => b.Parent != null).GroupBy(b => b.Parent))
                {
                    var bands = group.OrderBy(b => group.Key.Components.IndexOf(b)).ToList();
                    for (var index = 0; index < bands.Count; index++)
                    {
                        LiftBodyBands(bands[index], index == 0, index == bands.Count - 1);
                    }
                }
            }

            bodyBands.Clear();
        }

        /// <summary>
        /// Replaces the band of a body with the bands of its strips. The band stays when no strip has a data
        /// region alone or the band has its own look that the strips would lose.
        /// </summary>
        private void LiftBodyBands(StiDataBand band, bool isFirstPart, bool isLastPart)
        {
            if (!StiBrush.IsTransparent(band.Brush) || HasBorder(band.Border) || backgroundImages.ContainsValue(band))
                return;

            DissolveRectangles(band);

            var strips = GetBodyStrips(band);
            if (!strips.Any(IsBandStrip))
                return;

            //the report title and summary print across all the columns, the body of the columns flows in them
            var page = band.Page;
            var canTitle = isFirstPart && !band.NewPageBefore && page.Columns < 2;
            var canSummary = isLastPart && page.Columns < 2;

            var parent = band.Parent;
            var index = parent.Components.IndexOf(band);
            parent.Components.RemoveAt(index);

            var bands = new List<StiBand>();
            StiBand lastStaticBand = null;
            double bottom = 0;
            for (var stripIndex = 0; stripIndex < strips.Count; stripIndex++)
            {
                var strip = strips[stripIndex];
                var stripBottom = strip.Max(c => c.Bottom);

                if (IsBandStrip(strip))
                {
                    //the space above a data region stays, as in Reporting Services; the body shrinks, so the band
                    //that keeps the space does not
                    var space = strip[0].Top - bottom;
                    if (space > BodyOverlapTolerance && lastStaticBand != null)
                    {
                        lastStaticBand.Height += space;
                        lastStaticBand.CanShrink = false;
                    }
                    else if (space > BodyOverlapTolerance)
                    {
                        var spaceBand = CreateBodyStaticBand(new StiDataBand(), band, "Space", space);
                        spaceBand.CanShrink = false;
                        bands.Add(spaceBand);
                    }

                    bands.AddRange(LiftDataRegionBands(band, (StiPanel)strip[0]));
                    lastStaticBand = null;
                }
                else
                {
                    var isFirst = bands.Count == 0;
                    var isLast = stripIndex == strips.Count - 1;

                    StiBand staticBand;
                    if (isFirst && canTitle)
                        staticBand = CreateBodyStaticBand(new StiReportTitleBand(), band, "ReportTitle", 0);

                    else if (isLast && canSummary)
                        staticBand = CreateBodyStaticBand(new StiReportSummaryBand(), band, "ReportSummary", 0);

                    else
                        staticBand = CreateBodyStaticBand(new StiDataBand(), band, band.Name, 0);

                    //the strip keeps the space above it and the order of its items
                    foreach (var component in band.Components.Cast<StiComponent>().Where(strip.Contains).ToList())
                    {
                        band.Components.Remove(component);
                        component.Top -= bottom;
                        staticBand.Components.Add(component);
                    }

                    staticBand.Height = stripBottom - bottom;
                    bands.Add(staticBand);
                    lastStaticBand = staticBand;
                }

                bottom = Math.Max(bottom, stripBottom);
            }

            if (lastStaticBand != null && band.Height > bottom)
                lastStaticBand.Height += band.Height - bottom;

            //a part of the body starts a new page; the first band is a dynamic one, the title is not allowed there
            if (band.NewPageBefore && bands[0] is StiDynamicBand firstBand)
            {
                firstBand.NewPageBefore = true;
                firstBand.SkipFirst = false;
            }

            foreach (var newBand in bands)
            {
                parent.Components.Insert(index++, newBand);
            }
        }

        /// <summary>
        /// Puts the items of the rectangles on the band of the body when a data region of a rectangle gets a strip
        /// of its own that way. A rectangle only groups its items, so one without a look and alone in its strip
        /// is lost without a trace; a rectangle in a rectangle is taken the same way.
        /// </summary>
        private void DissolveRectangles(StiDataBand band)
        {
            var dissolved = true;
            while (dissolved)
            {
                dissolved = false;
                foreach (var strip in GetBodyStrips(band))
                {
                    if (strip.Count == 1 && IsLiftableRectangle(strip[0]))
                    {
                        DissolveRectangle(band, (StiPanel)strip[0]);
                        dissolved = true;
                    }
                }
            }
        }

        /// <summary>
        /// Returns true when the component is a rectangle without a look of its own and one of its data regions
        /// gets a strip alone when the rectangle is dissolved. A rectangle that keeps its items together on a page
        /// stays, the bands of the page do not keep them together.
        /// </summary>
        private bool IsLiftableRectangle(StiComponent component)
        {
            var panel = component as StiPanel;
            if (panel == null || !rectanglePanels.Contains(panel) || keptTogetherRectangles.Contains(panel) ||
                HasOwnLook(panel))
                return false;

            return GetBodyStrips(panel).Any(s => IsBandStrip(s) || s.Count == 1 && IsLiftableRectangle(s[0]));
        }

        /// <summary>
        /// Replaces the rectangle on the band with its items, in their order and at their places on the band.
        /// </summary>
        private static void DissolveRectangle(StiDataBand band, StiPanel panel)
        {
            var index = band.Components.IndexOf(panel);
            band.Components.RemoveAt(index);

            foreach (var component in panel.Components.Cast<StiComponent>().ToList())
            {
                panel.Components.Remove(component);
                component.Left += panel.Left;
                component.Top += panel.Top;
                band.Components.Insert(index++, component);
            }
        }

        /// <summary>
        /// Returns the items of the band of a body or of a container by the strips that no item crosses, from
        /// the top down.
        /// </summary>
        private static List<List<StiComponent>> GetBodyStrips(StiContainer band)
        {
            var strips = new List<List<StiComponent>>();
            double stripBottom = 0;
            foreach (var component in band.Components.Cast<StiComponent>().OrderBy(c => c.Top))
            {
                if (strips.Count == 0 || component.Top >= stripBottom - BodyOverlapTolerance)
                    strips.Add(new List<StiComponent>());

                strips[strips.Count - 1].Add(component);
                stripBottom = strips[strips.Count - 1].Count == 1
                    ? component.Bottom
                    : Math.Max(stripBottom, component.Bottom);
            }

            return strips;
        }

        /// <summary>
        /// Returns true when the strip is one data region that can give its bands to the page.
        /// </summary>
        private bool IsBandStrip(List<StiComponent> strip)
        {
            return strip.Count == 1 && IsBandPanel(strip[0]);
        }

        /// <summary>
        /// Returns true when the component is the panel of a data region that has only bands and nothing of its
        /// own: no visibility, no look, no action. Such a panel is lost without a trace when its bands go to the
        /// page.
        /// </summary>
        private bool IsBandPanel(StiComponent component)
        {
            var panel = component as StiPanel;
            if (panel == null || panel.GetType() != typeof(StiPanel) || panel.Components.Count == 0)
                return false;

            var bands = panel.Components.Cast<StiComponent>().ToList();
            if (!bands.All(c => c is StiDynamicBand) || !bands.Any(c => c is StiDataBand) || HasOwnLook(panel))
                return false;

            //a docked item keeps to the edge of the band, the left of the panel would be lost
            return panel.Left < BodyOverlapTolerance ||
                bands.All(b => ((StiBand)b).Components.Cast<StiComponent>().All(c => c.DockStyle == StiDockStyle.None));
        }

        /// <summary>
        /// Takes the bands out of the panel of a data region and moves their items by the left of the panel.
        /// </summary>
        private static List<StiBand> LiftDataRegionBands(StiDataBand band, StiPanel panel)
        {
            band.Components.Remove(panel);

            var bands = panel.Components.Cast<StiBand>().ToList();
            foreach (var dataRegionBand in bands)
            {
                panel.Components.Remove(dataRegionBand);
                dataRegionBand.Page = band.Page;

                foreach (StiComponent component in dataRegionBand.Components)
                {
                    component.Left += panel.Left;
                }
            }

            return bands;
        }

        /// <summary>
        /// Returns a band for the items of a strip of the body with the growth of the body. A data band prints
        /// once, also when the data regions of the strip have no rows.
        /// </summary>
        private StiBand CreateBodyStaticBand(StiBand staticBand, StiDataBand band, string name, double height)
        {
            staticBand.Name = CheckComponentName(band.Page.Report, name);
            staticBand.Page = band.Page;
            staticBand.CanGrow = band.CanGrow;
            staticBand.CanShrink = band.CanShrink;
            staticBand.ClientRectangle = new RectangleD(0, 0, band.Width, height);

            if (staticBand is StiDataBand dataBand)
            {
                dataBand.CountData = 1;
                dataBand.CanBreak = band.CanBreak;
                dataBand.PrintIfDetailEmpty = true;
            }

            return staticBand;
        }

        /// <summary>
        /// Returns true when the panel has something of its own that its items do not keep: the visibility, the
        /// fill, the border, a background image or an action.
        /// </summary>
        private bool HasOwnLook(StiPanel panel)
        {
            if (!panel.Enabled || panel.Expressions.Count > 0 || panel.Conditions.Count > 0 ||
                !string.IsNullOrEmpty(panel.BeforePrintEvent.Script))
                return true;

            if (!StiBrush.IsTransparent(panel.Brush) || HasBorder(panel.Border) || backgroundImages.ContainsValue(panel))
                return true;

            return !string.IsNullOrEmpty(panel.Bookmark?.Value) || !string.IsNullOrEmpty(panel.Hyperlink?.Value) ||
                !string.IsNullOrEmpty(panel.ToolTip?.Value);
        }

        private static bool HasBorder(StiBorder border)
        {
            return border != null && (border.IsTopBorderSidePresent || border.IsBottomBorderSidePresent ||
                border.IsLeftBorderSidePresent || border.IsRightBorderSidePresent || border.DropShadow);
        }
        #endregion
    }
}
