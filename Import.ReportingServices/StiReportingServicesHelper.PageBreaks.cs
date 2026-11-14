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
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Converts the page breaks of the report items of Reporting Services. A rectangle or a data region breaks
    /// the page before it, after it or both; Stimulsoft breaks the page only before or after a band. The items
    /// of the body are on one band, so the band is cut at the edges of the items that break the page, and every
    /// next part is a band that starts a new page.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Fields.PageBreaks
        /// <summary>
        /// The items that break the page and where: start, end or both.
        /// </summary>
        private Dictionary<StiComponent, string> itemPageBreaks = new Dictionary<StiComponent, string>();
        #endregion

        #region Methods.PageBreaks
        /// <summary>
        /// Reads the page break of a rectangle or a data region. Between breaks the page between the instances
        /// of a group; an item has one instance, so it does not break the page.
        /// </summary>
        private void ProcessItemPageBreak(XmlNode baseNode, StiComponent component)
        {
            var location = string.Empty;
            foreach (XmlNode node in baseNode.ChildNodes)
            {
                switch (node.Name)
                {
                    case "BreakLocation":
                        location = GetNodeTextValue(node)?.Trim().ToLowerInvariant();
                        break;

                    case "Disabled":
                        var disabled = GetNodeTextValue(node);
                        if (IsExpression(disabled))
                        {
                            ThrowError(baseNode.Name, node.Name);
                            return;
                        }

                        if (IsTrue(disabled))
                            return;
                        break;

                    case "ResetPageNumber":
                        if (IsTrue(GetNodeTextValue(node)))
                        {
                            ThrowError(null, null,
                                $"Reset of the page number by an item not imported: {component.Name}");
                        }
                        break;

                    default:
                        ThrowError(baseNode.Name, node.Name);
                        break;
                }
            }

            if (location == "start" || location == "end" || location == "startandend")
                itemPageBreaks[component] = location;
        }

        /// <summary>
        /// Cuts the band of the body at the page breaks of its items. A break before the first item or after
        /// the last one gives no empty page, as in Reporting Services. The page break of an item inside another
        /// item is not imported.
        /// </summary>
        private void SplitBodyByPageBreaks(StiDataBand band, StiPage page)
        {
            foreach (var component in itemPageBreaks.Keys.Where(c => c.Parent != band))
            {
                ThrowError(null, null, $"Page break of an item inside another item not imported: {component.Name}");
            }

            var components = band.Components.Cast<StiComponent>().ToList();
            var cuts = new SortedSet<double>();
            foreach (var component in components)
            {
                if (!itemPageBreaks.TryGetValue(component, out var location))
                    continue;

                if (location != "end")
                    cuts.Add(component.Top);

                if (location != "start")
                    cuts.Add(component.Bottom);
            }

            var contentTop = components.Count > 0 ? components.Min(c => c.Top) : 0;
            var contentBottom = components.Count > 0 ? components.Max(c => c.Bottom) : 0;
            var parts = cuts.Where(c => c > contentTop && c < contentBottom).ToList();
            if (parts.Count == 0)
                return;

            var index = band.Parent.Components.IndexOf(band);
            for (var partIndex = parts.Count - 1; partIndex >= 0; partIndex--)
            {
                var start = parts[partIndex];
                var partComponents = components.Where(c => c.Top >= start).ToList();
                components = components.Except(partComponents).ToList();
                if (partComponents.Count == 0)
                    continue;

                //the part starts with its first item, the space between the cut and the item is on the previous page
                var offset = partComponents.Min(c => c.Top);
                var partBand = CreateBodyPart(band, page, partIndex + 2);
                foreach (var component in partComponents)
                {
                    band.Components.Remove(component);
                    component.Top -= offset;
                    partBand.Components.Add(component);
                }

                partBand.Height = partComponents.Max(c => c.Bottom);
                band.Parent.Components.Insert(index + 1, partBand);
            }

            band.Height = components.Count > 0 ? components.Max(c => c.Bottom) : 0;
        }

        /// <summary>
        /// Returns a band for a part of the body: it prints once, starts a new page and has the style of the body.
        /// The band has one row, so the new page is not skipped for the first row.
        /// </summary>
        private StiDataBand CreateBodyPart(StiDataBand band, StiPage page, int number)
        {
            var partBand = new StiDataBand
            {
                Name = CheckComponentName(page.Report, band.Name + number),
                Page = page,
                CanBreak = band.CanBreak,
                CanGrow = band.CanGrow,
                CanShrink = band.CanShrink,
                CountData = 1,
                PrintIfDetailEmpty = true,
                NewPageBefore = true,
                SkipFirst = false,
                Brush = band.Brush?.Clone() as StiBrush,
                Border = band.Border?.Clone() as StiBorder
            };
            partBand.ClientRectangle = new RectangleD(0, 0, band.Width, 0);
            bodyBands.Add(partBand);

            return partBand;
        }
        #endregion
    }
}
