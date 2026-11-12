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

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Xml;
using Stimulsoft.Base;
using Stimulsoft.Base.Drawing;
using Stimulsoft.Report.Components;
using Stimulsoft.Report.Dictionary;

namespace Stimulsoft.Report.Import
{
    /// <summary>
    /// Imports the RDLX reports of ActiveReports: the page and RDL reports of ActiveReports.NET (XML, .rdlx) and
    /// the reports of ActiveReportsJS (JSON, .rdlx-json). Both files hold the same model, so they are read into
    /// one tree of StiRdlxNode and imported by the same code.
    /// A section of the report becomes a page. The body is cut into horizontal strips: a data region that
    /// shares its rows with nothing else gives its bands straight to the page, everything else between such
    /// regions goes to a band printed once, where a data region is a panel with its own bands.
    /// </summary>
    public partial class StiActiveReportsRdlxHelper
    {
        #region Consts
        /// <summary>
        /// The marker of the data band of a total, replaced by the name of the band when all bands are made.
        /// </summary>
        private const string BandMarkerStart = "<<band:";
        private const string BandMarkerEnd = ">>";

        /// <summary>
        /// The difference of two coordinates, in hundredths of an inch, that is taken as none.
        /// </summary>
        private const double Tolerance = 0.5;
        #endregion

        #region Fields
        private StiReport report;
        private List<string> errors;
        private StiRdlxNode root;
        private bool isJson;
        private string dataFolder;
        private HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private HashSet<string> warnings = new HashSet<string>(StringComparer.Ordinal);
        private HashSet<string> regionNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, StiRdlxNode> embeddedImages =
            new Dictionary<string, StiRdlxNode>(StringComparer.OrdinalIgnoreCase);
        private Dictionary<string, string> parameters =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        #endregion

        #region Methods
        /// <summary>
        /// Fills the report from the root node of an RDLX definition.
        /// </summary>
        /// <param name="folder">The folder of the report file, the data files of the report are looked up
        /// in it. May be null.</param>
        public void ProcessRoot(StiRdlxNode rootNode, StiReport report, List<string> errorList, bool isJson,
            string folder)
        {
            this.report = report;
            this.errors = errorList;
            this.root = rootNode;
            this.isJson = isJson;
            this.dataFolder = folder;

            report.Pages.Clear();
            report.ReportUnit = StiReportUnitType.HundredthsOfInch;
            report.CalculationMode = StiCalculationMode.Interpretation;

            var reportName = root.Get("Name");
            if (!string.IsNullOrEmpty(reportName))
            {
                report.ReportName = reportName;
                report.ReportAlias = reportName;
            }

            // The formats of numbers and dates follow the language of the report.
            var language = root.Get("Language");
            if (!string.IsNullOrEmpty(language) && !language.StartsWith("="))
                report.Culture = language;

            CollectRegionNames(root);
            ProcessEmbeddedImages();
            ProcessParameters();
            ProcessData();

            var sections = root.Items("ReportSections");
            if (sections.Count == 0)
            {
                ProcessSection(root);
            }
            else
            {
                foreach (var section in sections)
                {
                    ProcessSection(section);
                }
            }

            if (report.Pages.Count == 0)
                report.Pages.Add(new StiPage(report));

            AddEmbeddedData();
            ResolveBandMarkers();
            ArrangeBands();
        }

        /// <summary>
        /// Collects the names of the data regions: a total may name one of them as its scope.
        /// </summary>
        private void CollectRegionNames(StiRdlxNode node)
        {
            var name = node.Get("Name");
            if (!string.IsNullOrEmpty(name) && (IsDataRegion(node) || node.IsType("Chart") || node.IsType("FixedPage")))
                regionNames.Add(name);

            foreach (var child in node.Nodes)
            {
                CollectRegionNames(child);
            }
        }

        private void ProcessEmbeddedImages()
        {
            foreach (var image in root.Items("EmbeddedImages"))
            {
                var name = image.Get("Name");
                if (!string.IsNullOrEmpty(name))
                    embeddedImages[name] = image;
            }
        }

        private void ProcessSection(StiRdlxNode section)
        {
            var body = section.Child("Body");
            var fixedPage = section.Child("FixedPage");

            if (fixedPage == null && body != null)
            {
                foreach (var item in body.Items("ReportItems"))
                {
                    if (item.IsType("FixedPage"))
                        fixedPage = item;
                }
            }

            if (fixedPage != null)
            {
                ProcessFixedPage(section, fixedPage);
                return;
            }

            var page = CreatePage(section, null);
            var scope = new StiRdlxScope();

            AddPageBand(section.Child("PageHeader"), new StiPageHeaderBand(), page, scope);
            AddPageBand(section.Child("PageFooter"), new StiPageFooterBand(), page, scope);

            if (body != null)
                BuildBody(body.Items("ReportItems"), page, scope);
        }

        /// <summary>
        /// A fixed page layout: every page of the layout becomes a page of the report. When the layout is
        /// grouped, a page is printed for each group of its data set.
        /// </summary>
        private void ProcessFixedPage(StiRdlxNode section, StiRdlxNode fixedPage)
        {
            var dataSet = FindDataSet(fixedPage.Get("DataSetName"));
            var group = fixedPage.Child("Group", "Grouping");

            foreach (var pageNode in fixedPage.Items("Pages"))
            {
                var page = CreatePage(section, pageNode);
                var items = pageNode.Items("ReportItems");
                var scope = new StiRdlxScope { DataSet = dataSet };

                if (dataSet != null && GetGroupExpression(group) != null)
                    BuildGroupedItems(items, group, dataSet, page, scope, isNewPage: true);

                else
                    BuildBody(items, page, scope);
            }
        }

        private StiPage CreatePage(StiRdlxNode section, StiRdlxNode fixedPageNode)
        {
            var page = new StiPage(report);
            page.Name = CreateName("Page");

            var settings = FindPageSettings(section);
            if (settings != null)
            {
                var width = ToHi(settings.Get("PageWidth"), 850);
                var height = ToHi(settings.Get("PageHeight"), 1100);

                page.Orientation = width > height ? StiPageOrientation.Landscape : StiPageOrientation.Portrait;
                page.PageWidth = width;
                page.PageHeight = height;
                page.Margins = new StiMargins(
                    ToHi(settings.Get("LeftMargin"), 0),
                    ToHi(settings.Get("RightMargin"), 0),
                    ToHi(settings.Get("TopMargin"), 0),
                    ToHi(settings.Get("BottomMargin"), 0));
            }

            if (fixedPageNode != null && fixedPageNode.Has("LeftMargin"))
            {
                page.Margins = new StiMargins(
                    ToHi(fixedPageNode.Get("LeftMargin"), 0),
                    ToHi(fixedPageNode.Get("RightMargin"), 0),
                    ToHi(fixedPageNode.Get("TopMargin"), 0),
                    ToHi(fixedPageNode.Get("BottomMargin"), 0));
            }

            report.Pages.Add(page);
            return page;
        }

        /// <summary>
        /// The page settings are in the Page node of the section or of the report in JSON and in the report
        /// node itself in XML.
        /// </summary>
        private StiRdlxNode FindPageSettings(StiRdlxNode section)
        {
            var candidates = new[] { section.Child("Page"), section, root.Child("Page"), root };

            foreach (var candidate in candidates)
            {
                if (candidate != null && candidate.Has("PageWidth"))
                    return candidate;
            }

            return null;
        }

        private void AddPageBand(StiRdlxNode node, StiBand band, StiPage page, StiRdlxScope scope)
        {
            if (node == null)
                return;

            band.Name = CreateName(band is StiPageHeaderBand ? "PageHeader" : "PageFooter");
            band.Height = ToHi(node.Get("Height"), 0);
            ApplyContainerStyle(node, band);
            page.Components.Add(band);

            AddItems(node.Items("ReportItems"), band, 0, 0, scope);
        }

        /// <summary>
        /// Places the items of a body on a page. A data region that no other item stands beside gives its bands
        /// to the page, the items between such regions go to bands printed once.
        /// </summary>
        private void BuildBody(List<StiRdlxNode> items, StiContainer container, StiRdlxScope scope)
        {
            var sorted = SortByTop(items);
            var strip = new List<StiRdlxNode>();
            var cursor = 0d;

            foreach (var node in sorted)
            {
                if (!IsFlowRegion(node, sorted))
                {
                    strip.Add(node);
                    continue;
                }

                var rect = GetRect(node);
                AddStrip(strip, cursor, rect.Top, container, scope);
                strip.Clear();

                BuildDataRegion(node, container, rect.Left, scope);
                cursor = rect.Bottom;
            }

            AddStrip(strip, cursor, cursor, container, scope);
        }

        /// <summary>
        /// Adds a band printed once with the items that stand between two data regions.
        /// </summary>
        private void AddStrip(List<StiRdlxNode> items, double top, double bottom, StiContainer container,
            StiRdlxScope scope)
        {
            foreach (var item in items)
            {
                bottom = Math.Max(bottom, GetRect(item).Bottom);
            }

            if (items.Count == 0 && bottom - top < 1)
                return;

            var band = CreateStaticBand(bottom - top);
            container.Components.Add(band);
            AddItems(items, band, 0, top, scope);
        }

        private StiDataBand CreateStaticBand(double height)
        {
            var band = new StiDataBand();
            band.Name = CreateName("StaticBand");
            band.CountData = 1;
            band.CanBreak = true;
            band.Height = Math.Max(0, height);
            return band;
        }

        private bool IsFlowRegion(StiRdlxNode node, List<StiRdlxNode> items)
        {
            if (!IsDataRegion(node))
                return false;

            var rect = GetRect(node);

            foreach (var other in items)
            {
                if (other == node)
                    continue;

                var otherRect = GetRect(other);
                if (otherRect.Top < rect.Bottom - Tolerance && rect.Top < otherRect.Bottom - Tolerance)
                    return false;
            }

            return true;
        }

        private static bool IsDataRegion(StiRdlxNode node)
        {
            return node.IsType("Table") || node.IsType("List") || node.IsType("BandedList") || node.IsType("Tablix");
        }

        private List<StiRdlxNode> SortByTop(List<StiRdlxNode> items)
        {
            var sorted = new List<StiRdlxNode>(items);
            var order = new Dictionary<StiRdlxNode, int>();

            for (var index = 0; index < items.Count; index++)
            {
                order[items[index]] = index;
            }

            sorted.Sort((a, b) =>
            {
                var result = GetRect(a).Top.CompareTo(GetRect(b).Top);
                return result != 0 ? result : order[a].CompareTo(order[b]);
            });

            return sorted;
        }

        /// <summary>
        /// Replaces the markers of the totals placed out of data regions by the names of the data bands.
        /// </summary>
        private void ResolveBandMarkers()
        {
            foreach (StiComponent component in report.GetComponents())
            {
                var text = component as StiText;
                if (text == null || text.Text.Value == null || !text.Text.Value.Contains(BandMarkerStart))
                    continue;

                text.Text.Value = ResolveBandMarkers(text.Text.Value);
            }
        }

        private string ResolveBandMarkers(string value)
        {
            var builder = new StringBuilder();
            var position = 0;

            while (true)
            {
                var start = value.IndexOf(BandMarkerStart, position, StringComparison.Ordinal);
                var end = start < 0 ? -1 : value.IndexOf(BandMarkerEnd, start, StringComparison.Ordinal);
                if (end < 0)
                    break;

                builder.Append(value, position, start - position);

                var name = value.Substring(start + BandMarkerStart.Length, end - start - BandMarkerStart.Length);
                var dataSet = FindDataSet(name);
                if (dataSet != null && dataSet.Band != null)
                    builder.Append(dataSet.Band.Name).Append(", ");

                position = end + BandMarkerEnd.Length;
            }

            builder.Append(value, position, value.Length - position);

            return builder.ToString().Replace(", )", ")");
        }

        /// <summary>
        /// Puts the bands of the pages and of the panels one under another, as the designer does.
        /// </summary>
        private void ArrangeBands()
        {
            foreach (StiPage page in report.Pages)
            {
                AssignPage(page, page);
                ArrangeBands(page);
            }
        }

        /// <summary>
        /// Sets the page of the components: a component added to a container that is not on a page yet does
        /// not get it.
        /// </summary>
        private void AssignPage(StiContainer container, StiPage page)
        {
            foreach (StiComponent component in container.Components)
            {
                component.Page = page;

                var child = component as StiContainer;
                if (child != null)
                    AssignPage(child, page);
            }
        }

        private void ArrangeBands(StiContainer container)
        {
            var hasBands = false;

            foreach (StiComponent component in container.Components)
            {
                var child = component as StiContainer;
                if (child != null)
                    ArrangeBands(child);

                if (component is StiBand)
                    hasBands = true;
            }

            if (hasBands)
                container.DockToContainer();
        }

        /// <summary>
        /// Makes a name that is unique in the report and is a valid identifier.
        /// </summary>
        private string CreateName(string preferred)
        {
            var builder = new StringBuilder();

            foreach (var ch in preferred ?? "")
            {
                if (char.IsLetterOrDigit(ch) || ch == '_')
                    builder.Append(ch);
            }

            if (builder.Length == 0 || char.IsDigit(builder[0]))
                builder.Insert(0, "Item");

            var baseName = builder.ToString();
            var name = baseName;
            var index = 1;

            if (IsGeneratedName(preferred))
                name = baseName + index++;

            while (!names.Add(name))
            {
                name = baseName + index++;
            }

            return name;
        }

        private static bool IsGeneratedName(string name)
        {
            switch (name)
            {
                case "Page":
                case "PageHeader":
                case "PageFooter":
                case "StaticBand":
                case "HeaderBand":
                case "FooterBand":
                case "GroupHeaderBand":
                case "GroupFooterBand":
                case "DataBand":
                case "Panel":
                case "Text":
                case "Item":
                    return true;

                default:
                    return false;
            }
        }

        private void Warn(string message)
        {
            if (warnings.Add(message))
                errors.Add(message);
        }
        #endregion

        #region Methods.Import
        /// <summary>
        /// Finds the format of an ActiveReports file by its content.
        /// </summary>
        public static StiActiveReportsFormat GetFormat(byte[] bytes)
        {
            var text = GetText(bytes, 4096);

            foreach (var ch in text)
            {
                if (char.IsWhiteSpace(ch) || ch == (char)0xFEFF)
                    continue;

                if (ch == '{')
                    return StiActiveReportsFormat.RdlxJson;

                break;
            }

            return text.Contains("ActiveReportsLayout") ? StiActiveReportsFormat.Rpx : StiActiveReportsFormat.RdlxXml;
        }

        public static StiImportResult Import(byte[] bytes)
        {
            return Import(bytes, null);
        }

        /// <summary>
        /// Imports an RDLX report, XML or JSON.
        /// </summary>
        /// <param name="folder">The folder of the report file, the data files of the report are looked up
        /// in it. May be null.</param>
        public static StiImportResult Import(byte[] bytes, string folder)
        {
            var currentCulture = Thread.CurrentThread.CurrentCulture;

            try
            {
                Thread.CurrentThread.CurrentCulture = StiCultureInfo.GetEN(false);

                var report = new StiReport();
                var helper = new StiActiveReportsRdlxHelper();
                var errors = new List<string>();
                var isJson = GetFormat(bytes) == StiActiveReportsFormat.RdlxJson;

                var rootNode = isJson ? StiRdlxReader.ReadJson(GetText(bytes, bytes.Length)) : ReadXml(bytes);
                helper.ProcessRoot(rootNode, report, errors, isJson, folder);

                return new StiImportResult(report, errors);
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = currentCulture;
            }
        }

        private static StiRdlxNode ReadXml(byte[] bytes)
        {
            var doc = new XmlDocument();

            using (var stream = new MemoryStream(bytes))
            {
                doc.Load(stream);
            }

            return StiRdlxReader.ReadXml(doc.DocumentElement);
        }

        private static string GetText(byte[] bytes, int count)
        {
            count = Math.Min(count, bytes.Length);

            if (count >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
                return Encoding.Unicode.GetString(bytes, 2, count - 2);

            if (count >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
                return Encoding.BigEndianUnicode.GetString(bytes, 2, count - 2);

            if (count >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
                return Encoding.UTF8.GetString(bytes, 3, count - 3);

            return Encoding.UTF8.GetString(bytes, 0, count);
        }
        #endregion
    }
}
