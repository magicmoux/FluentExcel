﻿// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using NPOI.HPSF;
    using NPOI.HSSF.UserModel;
    using NPOI.HSSF.Util;
    using NPOI.SS.UserModel;
    using NPOI.SS.Util;
    using NPOI.XSSF.UserModel;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;

    /// <summary>
    /// Defines some extensions for <see cref="IEnumerable{T}"/> that using NPOI to provides excel functionality.
    /// </summary>
    public static class IEnumerableNpoiExtensions
    {
        private static IFormulaEvaluator _formulaEvaluator;

        public static byte[] ToExcelContent<T>(this IEnumerable<T> source, string sheetName = "sheet0", int maxRowsPerSheet = int.MaxValue, bool overwrite = false)
            where T : class
        {
            return ToExcel(source, null, s => sheetName, maxRowsPerSheet, overwrite);
        }

        public static void ToExcel<T>(this IEnumerable<T> source, string excelFile, string sheetName = "sheet0", int maxRowsPerSheet = int.MaxValue, bool overwrite = false)
            where T : class
        {
            //TODO check the file's path is valid
            ToExcel(source, excelFile, s => sheetName, maxRowsPerSheet, overwrite);
        }

        public static byte[] ToExcel<T>(this IEnumerable<T> source, string excelFile, Expression<Func<T, string>> sheetSelector, int maxRowsPerSheet = int.MaxValue, bool overwrite = false)
            where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            bool isVolatile = string.IsNullOrWhiteSpace(excelFile);
            if (!isVolatile)
            {
                var extension = Path.GetExtension(excelFile);
                if (extension.Equals(".xls"))
                {
                    Excel.Setting.UserXlsx = false;
                }
                else if (extension.Equals(".xlsx"))
                {
                    Excel.Setting.UserXlsx = true;
                }
                else
                {
                    throw new NotSupportedException($"not an excel file (*.xls | *.xlsx) extension: {extension}");
                }
            }
            else
            {
                excelFile = null;
            }

            IWorkbook book = InitializeWorkbook(excelFile);
            using (Stream ms = isVolatile ? (Stream)new MemoryStream() : new FileStream(excelFile, FileMode.OpenOrCreate, FileAccess.Write))
            {
                IEnumerable<byte> output = Enumerable.Empty<byte>();
                foreach (var sheet in source.AsQueryable().GroupBy(sheetSelector))
                {
                    int sheetIndex = 0;
                    var content = sheet.Select(row => row);
                    while (content.Any())
                    {
                        book = content.Take(maxRowsPerSheet).ToWorkbook(book, sheet.Key + (sheetIndex > 0 ? "_" + sheetIndex.ToString() : ""), overwrite);
                        sheetIndex++;
                        content = content.Skip(maxRowsPerSheet);
                    }
                }
                book.Write(ms);
                return isVolatile ? ((MemoryStream)ms).ToArray() : null;
            }
        }

        #region TODO relocate into a "Util" class

        internal static IWorkbook ToWorkbook<T>(this IEnumerable<T> source, IWorkbook workbook, string sheetName, bool overwrite = false)
        {
            #region TODO Handle a specific configuration parameter

            // can static properties or only instance properties?
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.GetProperty);

            bool fluentConfigEnabled = false;
            // get the fluent config for the sheet if it exists
            if (Excel.Setting.FluentConfigs.TryGetValue(typeof(T).FullName, out var fluentConfig))
            {
                fluentConfigEnabled = true;
            }

            // find out the configurations
            var propertyConfigurations = new PropertyConfiguration[properties.Length];
            for (var j = 0; j < properties.Length; j++)
            {
                var property = properties[j];

                // get the property config
                if (fluentConfigEnabled && fluentConfig.PropertyConfigurations.TryGetValue(property.Name, out var pc))
                {
                    propertyConfigurations[j] = pc;
                }
                else
                {
                    propertyConfigurations[j] = null;
                }
            }

            #endregion

            // new sheet
            //TODO check the sheet's name is valid
            var sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            else
            {
                // doesn't override the exist sheet if not required
                if (!overwrite) sheet = workbook.CreateSheet();
            }

            #region TODO make this configurable

            // cache cell styles
            var cellStyles = new Dictionary<int, ICellStyle>();

            // title row cell style
            var titleStyle = workbook.CreateCellStyle();
            titleStyle.Alignment = HorizontalAlignment.Center;
            titleStyle.VerticalAlignment = VerticalAlignment.Center;
            titleStyle.FillPattern = FillPattern.Bricks;
            titleStyle.FillBackgroundColor = HSSFColor.Grey40Percent.Index;
            titleStyle.FillForegroundColor = HSSFColor.White.Index;

            #endregion

            var titleRow = sheet.CreateRow(0);
            var rowIndex = 1;
            foreach (var item in source)
            {
                var row = sheet.CreateRow(rowIndex);
                for (var i = 0; i < properties.Length; i++)
                {
                    var property = properties[i];

                    int index = i;
                    var config = propertyConfigurations[i];
                    if (config != null)
                    {
                        if (config.IsExportIgnored)
                            continue;

                        index = config.Index;

                        if (index < 0 && !config.AutoIndex)
                            throw new Exception($"The excel cell index value hasn't been configured for the property: {property.Name}, see HasExcelIndex(int index) or AdjustAutoIndex() methods for more informations.");
                    }

                    // this is the first time.
                    if (rowIndex == 1)
                    {
                        // if not title, using property name as title.
                        var title = property.Name;
                        if (!string.IsNullOrEmpty(config?.Title))
                        {
                            title = config.Title;
                        }

                        if (!string.IsNullOrEmpty(config?.Formatter))
                        {
                            try
                            {
                                var style = workbook.CreateCellStyle();

                                var dataFormat = workbook.CreateDataFormat();

                                style.DataFormat = dataFormat.GetFormat(config.Formatter);

                                cellStyles[i] = style;
                            }
                            catch (Exception ex)
                            {
                                // the formatter isn't excel supported formatter
                                System.Diagnostics.Debug.WriteLine(ex.ToString());
                            }
                        }

                        var titleCell = titleRow.CreateCell(index);
                        titleCell.CellStyle = titleStyle;
                        titleCell.SetCellValue(title);
                    }

                    var unwrapType = property.PropertyType.UnwrapNullableType();

                    var value = property.GetValue(item, null);

                    // give a chance to the value converter even though value is null.
                    if (config?.ValueConverter != null)
                    {
                        value = config.ValueConverter(value);
                        if (value == null)
                            continue;

                        unwrapType = value.GetType().UnwrapNullableType();
                    }

                    if (value == null)
                        continue;

                    var cell = row.CreateCell(index);
                    if (cellStyles.TryGetValue(i, out var cellStyle))
                    {
                        cell.CellStyle = cellStyle;
                    }
                    else if (!string.IsNullOrEmpty(config?.Formatter) && value is IFormattable fv)
                    {
                        // the formatter isn't excel supported formatter, but it's a C# formatter.
                        // The result is the Excel cell data type become String.
                        cell.SetCellValue(fv.ToString(config.Formatter, CultureInfo.CurrentCulture));

                        continue;
                    }

                    if (unwrapType == typeof(bool))
                    {
                        cell.SetCellValue((bool)value);
                    }
                    else if (unwrapType == typeof(DateTime))
                    {
                        cell.SetCellValue(Convert.ToDateTime(value));
                    }
                    else if (unwrapType.IsInteger() ||
                             unwrapType == typeof(decimal) ||
                             unwrapType == typeof(double) ||
                             unwrapType == typeof(float))
                    {
                        cell.SetCellValue(Convert.ToDouble(value));
                    }
                    else
                    {
                        cell.SetCellValue(value.ToString());
                    }
                }

                rowIndex++;
            }

            // merge cells
            var mergableConfigs = propertyConfigurations.Where(c => c != null && c.AllowMerge).ToList();
            if (mergableConfigs.Any())
            {
                // merge cell style
                var vStyle = workbook.CreateCellStyle();
                vStyle.VerticalAlignment = VerticalAlignment.Center;

                foreach (var config in mergableConfigs)
                {
                    object previous = null;
                    int rowspan = 0, row = 1;
                    for (row = 1; row < rowIndex; row++)
                    {
                        var value = sheet.GetRow(row).GetCellValue(config.Index, _formulaEvaluator);
                        if (object.Equals(previous, value) && value != null)
                        {
                            rowspan++;
                        }
                        else
                        {
                            if (rowspan > 1)
                            {
                                sheet.GetRow(row - rowspan).Cells[config.Index].CellStyle = vStyle;
                                sheet.AddMergedRegion(new CellRangeAddress(row - rowspan, row - 1, config.Index, config.Index));
                            }
                            rowspan = 1;
                            previous = value;
                        }
                    }

                    // in what case? -> all rows need to be merged
                    if (rowspan > 1)
                    {
                        sheet.GetRow(row - rowspan).Cells[config.Index].CellStyle = vStyle;
                        sheet.AddMergedRegion(new CellRangeAddress(row - rowspan, row - 1, config.Index, config.Index));
                    }
                }
            }

            if (rowIndex > 1 && fluentConfigEnabled)
            {
                var statistics = fluentConfig.StatisticsConfigurations;
                var filterConfigs = fluentConfig.FilterConfigurations;
                var freezeConfigs = fluentConfig.FreezeConfigurations;

                // statistics row
                foreach (var item in statistics)
                {
                    var lastRow = sheet.CreateRow(rowIndex);
                    var cell = lastRow.CreateCell(0);
                    cell.SetCellValue(item.Name);
                    foreach (var column in item.Columns)
                    {
                        cell = lastRow.CreateCell(column);

                        // set the same cell style
                        cell.CellStyle = sheet.GetRow(rowIndex - 1)?.GetCell(column)?.CellStyle;

                        // set the cell formula
                        cell.CellFormula = $"{item.Formula}({GetCellPosition(1, column)}:{GetCellPosition(rowIndex - 1, column)})";
                    }

                    rowIndex++;
                }

                // set the freeze
                foreach (var freeze in freezeConfigs)
                {
                    sheet.CreateFreezePane(freeze.ColSplit, freeze.RowSplit, freeze.LeftMostColumn, freeze.TopRow);
                }

                // set the auto filter
                foreach (var filter in filterConfigs)
                {
                    sheet.SetAutoFilter(new CellRangeAddress(filter.FirstRow, filter.LastRow ?? rowIndex, filter.FirstCol, filter.LastCol));
                }
            }

            // autosize the all columns
            for (int i = 0; i < properties.Length; i++)
            {
                sheet.AutoSizeColumn(i);
            }

            return workbook;
        }

        private static IWorkbook InitializeWorkbook(string excelFile)
        {
            var setting = Excel.Setting;
            if (setting.UserXlsx)
            {
                if (!string.IsNullOrEmpty(excelFile) && File.Exists(excelFile))
                {
                    using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                    {
                        var workbook = new XSSFWorkbook(file);

                        _formulaEvaluator = new XSSFFormulaEvaluator(workbook);

                        return workbook;
                    }
                }
                else
                {
                    var workbook = new XSSFWorkbook();

                    _formulaEvaluator = new XSSFFormulaEvaluator(workbook);

                    return workbook;
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(excelFile) && File.Exists(excelFile))
                {
                    using (var file = new FileStream(excelFile, FileMode.Open, FileAccess.Read))
                    {
                        var workbook = new HSSFWorkbook(file);

                        _formulaEvaluator = new HSSFFormulaEvaluator(workbook);

                        return workbook;
                    }
                }
                else
                {
                    var workbook = new HSSFWorkbook();

                    _formulaEvaluator = new HSSFFormulaEvaluator(workbook);

                    var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                    dsi.Company = setting.Company;
                    workbook.DocumentSummaryInformation = dsi;

                    var si = PropertySetFactory.CreateSummaryInformation();
                    si.Author = setting.Author;
                    si.Subject = setting.Subject;
                    workbook.SummaryInformation = si;

                    return workbook;
                }
            }
        }

        private static string GetCellPosition(int row, int col)
        {
            col = Convert.ToInt32('A') + col;
            row = row + 1;
            return ((char)col) + row.ToString();
        }

        #endregion
    }
}