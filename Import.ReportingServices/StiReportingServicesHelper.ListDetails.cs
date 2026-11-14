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

using Stimulsoft.Report.Components;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Puts the details of a list onto the bands of the list. A list of Reporting Services repeats a rectangle for
    /// every group, and a table in the rectangle shows the rows of the group. The import puts the rectangle on the
    /// header of the group and the table on a panel inside it: the data band of the table either has no data
    /// source and prints one row, or prints all the rows of the data set in the first group. So the bands of the
    /// table go to the list: its header to the header of the group, its details to the data band of the list,
    /// its footer and the items below it to a footer of the group.
    /// </summary>
    public partial class StiReportingServicesHelper
    {
        #region Methods.ListDetails
        /// <summary>
        /// Puts the table of every list with one group and a rectangle onto the bands of the list.
        /// </summary>
        private void MergeListDetails()
        {
            foreach (var contents in rectanglePanels.ToList())
            {
                if (MergeListDetails(contents))
                    rectanglePanels.Remove(contents);
            }
        }

        /// <summary>
        /// Puts the table of the rectangle onto the bands of its list. The rectangle has to be the only item of
        /// the header of the innermost group, without a look of its own, with one table of the data set of the
        /// list and nothing beside the table; the list has no other bands.
        /// </summary>
        private bool MergeListDetails(StiPanel contents)
        {
            var groupHeader = contents.Parent as StiGroupHeaderBand;
            var list = groupHeader?.Parent as StiPanel;
            if (list == null || groupHeader.Components.Count != 1 || HasOwnLook(contents))
                return false;

            var listBands = list.Components.Cast<StiComponent>().ToList();
            var listDataBand = listBands.LastOrDefault() as StiDataBand;
            if (listDataBand == null || listBands.IndexOf(groupHeader) != listBands.Count - 2 ||
                !listBands.Take(listBands.Count - 1).All(b => b is StiGroupHeaderBand) || !IsEmptyBand(listDataBand))
                return false;

            var strips = GetBodyStrips(contents);
            var regionStrips = strips.Where(IsBandStrip).ToList();
            if (regionStrips.Count != 1)
                return false;

            var region = (StiPanel)regionStrips[0][0];
            var regionBands = region.Components.Cast<StiBand>().ToList();
            var dataBands = regionBands.OfType<StiDataBand>().ToList();
            if (dataBands.Count != 1 || !regionBands.All(IsListDetailsBand))
                return false;

            var dataBand = dataBands[0];
            var dataSourceName = dataBand.DataSourceName;
            if (!string.IsNullOrEmpty(dataSourceName) && dataSourceName != listDataBand.DataSourceName ||
                dataBand.FilterMode != listDataBand.FilterMode)
                return false;

            var regionIndex = strips.IndexOf(regionStrips[0]);
            var above = strips.Take(regionIndex).SelectMany(s => s).ToList();
            var below = strips.Skip(regionIndex + 1).SelectMany(s => s).ToList();
            var header = regionBands.OfType<StiHeaderBand>().FirstOrDefault();
            var footer = regionBands.OfType<StiFooterBand>().FirstOrDefault();
            var left = contents.Left + region.Left;
            var groupHeaderHeight = groupHeader.Height;

            //the items above the table and the header of the table are printed once for the group
            groupHeader.Components.Remove(contents);
            MoveListItems(above, groupHeader, contents.Left, contents.Top);

            var headerTop = contents.Top + region.Top;
            if (header != null)
                MoveListItems(header.Components.Cast<StiComponent>().ToList(), groupHeader, left, headerTop);

            groupHeader.Height = headerTop + (header?.Height ?? 0);

            if (keptTogetherRectangles.Contains(contents))
                groupHeader.KeepGroupTogether = true;

            //the details of the table are the rows of the group
            dataBand.DataSourceName = listDataBand.DataSourceName;
            dataBand.Sort = (listDataBand.Sort ?? new string[0]).Concat(dataBand.Sort ?? new string[0]).ToArray();
            dataBand.Filters = MergeFilters(listDataBand.Filters, dataBand.Filters) ?? new StiFiltersCollection();
            dataBand.FilterOn = listDataBand.FilterOn && dataBand.FilterOn;

            var index = list.Components.IndexOf(listDataBand);
            list.Components.Remove(listDataBand);
            foreach (var band in regionBands.Where(b => b != header && b != footer))
            {
                region.Components.Remove(band);
                band.Page = list.Page;
                list.Components.Insert(index++, band);

                foreach (StiComponent component in band.Components)
                {
                    component.Left += left;
                }
            }

            //the footer of the table and the items below it end the group
            var footerHeight = footer?.Height ?? 0;
            var height = footerHeight + groupHeaderHeight - contents.Top - region.Bottom;
            if (footer != null || below.Count > 0 || height > BodyOverlapTolerance)
            {
                var groupFooter = new StiGroupFooterBand
                {
                    Name = CheckComponentName(list.Report, $"{list.Name}_Gf{listBands.Count - 1}"),
                    Page = list.Page,
                    CanGrow = true
                };
                groupFooter.Height = height;

                if (footer != null)
                    MoveListItems(footer.Components.Cast<StiComponent>().ToList(), groupFooter, left, 0);

                MoveListItems(below, groupFooter, contents.Left, footerHeight - region.Bottom);
                list.Components.Insert(index, groupFooter);
            }

            ReplaceBandName(list, listDataBand.Name, dataBand.Name);
            return true;
        }

        /// <summary>
        /// Returns true when the band prints nothing: it has no items, no height and nothing of its own.
        /// </summary>
        private static bool IsEmptyBand(StiDataBand band)
        {
            return band.Components.Count == 0 && band.Height < BodyOverlapTolerance && band.Expressions.Count == 0 &&
                string.IsNullOrEmpty(band.Bookmark?.Value) && string.IsNullOrEmpty(band.BeforePrintEvent.Script);
        }

        /// <summary>
        /// Returns true when the band of a table can go to the list: a header, a footer, a group or the details.
        /// </summary>
        private static bool IsListDetailsBand(StiBand band)
        {
            return band is StiHeaderBand || band is StiFooterBand || band is StiGroupHeaderBand ||
                band is StiGroupFooterBand || band.GetType() == typeof(StiDataBand);
        }

        /// <summary>
        /// Moves the items to the band by the given offset.
        /// </summary>
        private static void MoveListItems(List<StiComponent> components, StiContainer band, double left, double top)
        {
            foreach (var component in components)
            {
                component.Parent.Components.Remove(component);
                component.Left += left;
                component.Top += top;
                band.Components.Add(component);
            }
        }

        /// <summary>
        /// Writes the new name of a band into the expressions of the components of the container.
        /// </summary>
        private static void ReplaceBandName(StiContainer container, string name, string newName)
        {
            var regex = new Regex($@"\b{Regex.Escape(name)}\b");
            foreach (StiComponent component in container.GetComponents())
            {
                ReplaceScopes(component, regex, m => newName);

                if (component.Interaction?.SortingColumn != null)
                    component.Interaction.SortingColumn = regex.Replace(component.Interaction.SortingColumn, newName);
            }
        }
        #endregion
    }
}
