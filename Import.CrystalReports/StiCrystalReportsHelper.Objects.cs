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
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Stimulsoft.Report.Components;
using Ras = CrystalDecisions.ReportAppServer.ReportDefModel;

namespace Import.CrystalReports
{
    /// <summary>
    /// The report objects which the engine of Crystal describes only partly: the pictures and the blob fields.
    /// Their content is read through the object model of the RAS.
    /// </summary>
    public sealed partial class StiCrystalReportsHelper
    {
        #region Consts.Objects
        private static readonly Regex RenderedPicture = new Regex(
            "<div id=\"([^\"]+)\"[^>]*>\\s*<img [^>]*src=\"([^\"]+)\"", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        #endregion

        #region Fields.Objects
        private ReportDocument mainDocument;
        private Dictionary<string, byte[]> renderedPictures;
        #endregion

        #region Methods.Objects
        /// <summary>
        /// A picture keeps its image in the report; Crystal stretches it to the frame of the object. The formula
        /// of the graphic location becomes the URL of the image.
        /// </summary>
        private void ApplyPicture(StiImage image, ReportObject obj)
        {
            image.Stretch = true;

            var picture = GetRasReportObject(obj) as Ras.ISCRPictureObject;
            var bytes = picture?.PictureData?.ByteArray;
            if (bytes == null || bytes.Length == 0)
                bytes = GetRenderedPicture(obj.Name);

            if (bytes != null && bytes.Length > 0)
                image.ImageBytes = bytes;

            var location = picture?.GraphicLocationFormula?.Text;
            if (!string.IsNullOrWhiteSpace(location))
            {
                pendingTranslations.Add(() =>
                {
                    var expression = TranslateFormula(location, image.Name + ".GraphicLocation");
                    if (expression != null && expression.Type == typeof(string))
                        image.ImageURL.Value = "{" + expression.Code + "}";
                });
            }

            if (image.ImageBytes == null && string.IsNullOrWhiteSpace(location))
                warningList.Add(StiWarning.GetWarningPicture(obj.Name));
        }

        /// <summary>
        /// Returns the image of a picture object as Crystal renders it. The in-process RAS doesn't give the data of
        /// the pictures, so a copy of the report is exported to HTML once, on tables with one empty row, and the PNG
        /// files are taken by the identifiers of their blocks, which are the names of the objects. The images have
        /// the size of the objects on the screen, not the original resolution.
        /// </summary>
        private byte[] GetRenderedPicture(string name)
        {
            if (renderedPictures == null)
                renderedPictures = RenderPictures();

            byte[] bytes;
            return renderedPictures.TryGetValue(name, out bytes) ? bytes : null;
        }

        private Dictionary<string, byte[]> RenderPictures()
        {
            var pictures = new Dictionary<string, byte[]>();
            var path = mainDocument?.FilePath;
            if (string.IsNullOrEmpty(path) || !File.Exists(path))
                return pictures;

            var folder = Path.Combine(Path.GetTempPath(), "StiCrystalPictures" + Guid.NewGuid().ToString("N"));
            var copy = new ReportDocument();
            try
            {
                copy.Load(path);
                SetEmptyData(copy);
                foreach (ReportDocument subreport in copy.Subreports)
                {
                    SetEmptyData(subreport);
                }

                SetEmptyParameters(copy);

                Directory.CreateDirectory(folder);
                var html = Path.Combine(folder, "report.htm");
                copy.Export(new ExportOptions
                {
                    ExportFormatType = ExportFormatType.HTML40,
                    ExportDestinationType = ExportDestinationType.DiskFile,
                    DestinationOptions = new DiskFileDestinationOptions { DiskFileName = html },
                    FormatOptions = new HTMLFormatOptions
                    {
                        HTMLBaseFolderName = folder,
                        HTMLFileName = "report.htm",
                        HTMLEnableSeparatedPages = false,
                        HTMLHasPageNavigator = false
                    }
                });

                ReadRenderedPictures(folder, pictures);
            }
            catch (Exception e)
            {
                warningList.Add(new StiWarning("Pictures: the images can't be rendered by Crystal", e.Message));
            }
            finally
            {
                copy.Close();
                if (Directory.Exists(folder))
                    Directory.Delete(folder, true);
            }

            return pictures;
        }

        /// <summary>
        /// Every table gets one row of the default values, so the sections print without a database.
        /// </summary>
        private static void SetEmptyData(ReportDocument report)
        {
            foreach (Table table in report.Database.Tables)
            {
                var dataTable = new DataTable(table.Name);
                foreach (DatabaseFieldDefinition field in table.Fields)
                {
                    if (!dataTable.Columns.Contains(field.Name))
                        dataTable.Columns.Add(field.Name, ConvertType(field.ValueType));
                }

                var row = dataTable.NewRow();
                foreach (DataColumn column in dataTable.Columns)
                {
                    row[column] = GetEmptyValue(column.DataType);
                }

                dataTable.Rows.Add(row);
                table.SetDataSource(dataTable);
            }
        }

        /// <summary>
        /// The value of a field of the empty row: not an empty string and not zero, because the conditions like
        /// "the field is empty" usually hide the sections.
        /// </summary>
        private static object GetEmptyValue(Type type)
        {
            if (type == typeof(string))
                return "1";

            if (type == typeof(DateTime))
                return DateTime.Today;

            if (type == typeof(bool))
                return false;

            if (type.IsValueType)
                return System.Convert.ChangeType(1, type);

            return DBNull.Value;
        }

        /// <summary>
        /// Every parameter of the main report gets its first default value or a simple value of its type, so
        /// the export doesn't stop for a missing value.
        /// </summary>
        private static void SetEmptyParameters(ReportDocument report)
        {
            foreach (ParameterFieldDefinition parameter in report.DataDefinition.ParameterFields)
            {
                if (!string.IsNullOrEmpty(parameter.ReportName))
                    continue;

                object value;
                var defaultValue = parameter.DefaultValues.Count > 0
                    ? parameter.DefaultValues[0] as ParameterDiscreteValue
                    : null;

                if (defaultValue != null)
                    value = defaultValue.Value;

                else
                    value = GetEmptyParameterValue(parameter.ParameterValueKind);

                if (parameter.DiscreteOrRangeKind == DiscreteOrRangeKind.RangeValue)
                    value = new ParameterRangeValue { StartValue = value, EndValue = value };

                report.SetParameterValue(parameter.Name, value);
            }
        }

        private static object GetEmptyParameterValue(ParameterValueKind kind)
        {
            switch (kind)
            {
                case ParameterValueKind.BooleanParameter:
                    return true;

                case ParameterValueKind.NumberParameter:
                case ParameterValueKind.CurrencyParameter:
                    return 1d;

                case ParameterValueKind.DateParameter:
                case ParameterValueKind.DateTimeParameter:
                    return DateTime.Today;

                default:
                    return "1";
            }
        }

        /// <summary>
        /// Reads the images of the export: a block div id="Picture1" contains the img of the picture.
        /// </summary>
        private static void ReadRenderedPictures(string folder, Dictionary<string, byte[]> pictures)
        {
            foreach (var file in Directory.GetFiles(folder, "*.htm*", SearchOption.AllDirectories))
            {
                var html = File.ReadAllText(file);
                foreach (Match match in RenderedPicture.Matches(html))
                {
                    var name = match.Groups[1].Value;
                    var image = Path.Combine(Path.GetDirectoryName(file), match.Groups[2].Value.Replace('/', '\\'));
                    if (!pictures.ContainsKey(name) && File.Exists(image))
                        pictures[name] = File.ReadAllBytes(image);
                }
            }
        }

        /// <summary>
        /// A blob field shows the image of a column of the database.
        /// </summary>
        private StiComponent ImportBlobFieldObject(ReportObject obj)
        {
            var image = new StiImage();
            ConvertProperties(image, obj);
            image.Stretch = true;

            var blob = GetRasReportObject(obj) as Ras.ISCRBlobFieldObject;
            var field = blob?.DataSourceName;
            if (field != null && field.StartsWith("{") && field.EndsWith("}"))
            {
                var column = htFieldNameConversion[field.Substring(1, field.Length - 2)] as string;
                if (column != null)
                    image.DataColumn = column;
            }

            return image;
        }
        #endregion
    }
}
