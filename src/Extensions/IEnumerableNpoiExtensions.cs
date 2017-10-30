// Copyright (c) rigofunc (xuyingting). All rights reserved.

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

        public static byte[] ToExcelContent<T>(this IEnumerable<T> source, string sheetName = "sheet0", int maxRowsPerSheet = int.MaxValue, bool overwrite = false, IFluentConfiguration configuration = null)
            where T : class
        {
            return ToExcel(source, null, s => sheetName, maxRowsPerSheet, overwrite, configuration);
        }

        [DefaultImplementation]
        public static byte[] ToExcel<T>(this IEnumerable<T> source, string excelFile, string sheetName = "sheet0", int maxRowsPerSheet = int.MaxValue, bool overwrite = false, IFluentConfiguration configuration = null)
            where T : class
        {
            //TODO check the file's path is valid
            //ToExcel(source, excelFile, s => sheetName, maxRowsPerSheet, overwrite, configuration);

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
                int sheetIndex = 0;
                var content = source.Where(i => i != null);
                while (sheetIndex == 0 || content.Any())
                {
                    book = content.Take(maxRowsPerSheet).ToWorkbook(book, sheetName + (sheetIndex > 0 ? "_" + sheetIndex.ToString() : ""), overwrite, configuration);
                    sheetIndex++;
                    content = content.Skip(maxRowsPerSheet);
                }
                book.Write(ms);
                return isVolatile ? ((MemoryStream)ms).ToArray() : null;
            }
        }

        public static byte[] ToExcel<T>(this IEnumerable<T> source, string excelFile, Expression<Func<T, string>> sheetSelector, int maxRowsPerSheet = int.MaxValue, bool overwrite = false, IFluentConfiguration configuration = null)
            where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            IEnumerable<byte> output = Enumerable.Empty<byte>();
            foreach (var sheet in source.AsQueryable().GroupBy(sheetSelector == null ? s => null : sheetSelector))
            {
                var result = ToExcel(sheet, excelFile, sheet.Key, maxRowsPerSheet, overwrite, configuration);
                if (result != null) output.Concat(result);
            }
            return output.ToArray();
        }

        #region TODO relocate into a "Util" class

        //TODO replace properties by the configuration and change overwrite defaulting to true
        internal static IWorkbook ToWorkbook<T>(this IEnumerable<T> source, IWorkbook workbook, string sheetName, bool overwrite = false, IFluentConfiguration configuration = null)
            where T : class
        {
            bool isConfigured = configuration != null;
            if (!isConfigured)
            {
                configuration = Excel.Setting.For<T>(false);
                // Get the fluent config for the type if it exists
                if (configuration != null)
                {
                    isConfigured = true;
                }
                else //TODO otherwise try creating is from annotations
                {
                    throw new ArgumentException($"No FluentExcel configuration found for type {typeof(T).FullName}");
                }
            }
            if (!isConfigured) throw new ArgumentException($"No FluentExcel configuration found for sheet {sheetName}");

            // new sheet
            //TODO check the sheet's name is valid
            var sheet = workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = workbook.CreateSheet(sheetName);
            }
            else
            {
                // doesn't override the existing sheet if not required
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

            var columns = configuration.ColumnConfigurations.Where(c => !c.IsExportIgnored).OrderBy(c => c.Index).ToList();
            var valueProviders = columns.Select(c => c.Expression.Compile()).ToList();

            var titleRow = sheet.CreateRow(0);
            for (var i = 0; i < columns.Count; i++)
            {
                var colConfig = columns[i];
                var title = colConfig.Title;
                if (!string.IsNullOrEmpty(colConfig?.Formatter))
                {
                    try
                    {
                        var style = workbook.CreateCellStyle();
                        var dataFormat = workbook.CreateDataFormat();
                        style.DataFormat = dataFormat.GetFormat(colConfig.Formatter);
                        cellStyles[i] = style;
                    }
                    catch (Exception ex)
                    {
                        // the formatter isn't excel supported formatter
                        System.Diagnostics.Debug.WriteLine(ex.ToString());
                    }
                }

                var titleCell = titleRow.CreateCell(i);
                titleCell.CellStyle = titleStyle;
                titleCell.SetCellValue(title);
            }
            var rowIndex = 1;
            foreach (var item in source)
            {
                var row = sheet.CreateRow(rowIndex);
                for (var i = 0; i < columns.Count; i++)
                {
                    int index = i;
                    var colConfig = columns[i];
                    var valueProvider = valueProviders[i];

                    //if (colConfig != null)
                    //{
                    //if (colConfig.IsExportIgnored)
                    //    continue;

                    // index = colConfig.Index;

                    //    //TODO check this
                    //    //if (index < 0 && !colConfig.AutoIndex)
                    //    //    throw new Exception($"The excel cell index value hasn't been configured for the property: {property.Name}, see HasExcelIndex(int index) or AdjustAutoIndex() methods for more informations.");
                    //}

                    var unwrapType = valueProvider.Method.ReturnType.UnwrapNullableType();
                    object value = null;
                    try
                    {
                        value = valueProvider.DynamicInvoke(item);
                    }
                    catch (TargetInvocationException)
                    {
                    }

                    // give a chance to the value converter even though value is null.
                    if (colConfig?.ValueConverter != null)
                    {
                        value = colConfig.ValueConverter(value);
                        if (value == null) continue;
                        unwrapType = value.GetType().UnwrapNullableType();
                    }
                    if (value == null) continue;

                    var cell = row.CreateCell(index);
                    if (cellStyles.TryGetValue(i, out var cellStyle))
                    {
                        cell.CellStyle = cellStyle;
                    }
                    else if (!string.IsNullOrEmpty(colConfig?.Formatter) && value is IFormattable fv)
                    {
                        // the formatter isn't excel supported formatter, but it's a C# formatter.
                        // The result is the Excel cell data type become String.
                        cell.SetCellValue(fv.ToString(colConfig.Formatter, CultureInfo.CurrentCulture));
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

            if (rowIndex > 1)
            {
                // merge cells
                var mergableConfigs = columns.Where(c => c != null && c.AllowMerge).ToList();
                if (mergableConfigs.Any())
                {
                    #region TODO make this configurable

                    // merge cell style

                    var vStyle = workbook.CreateCellStyle();
                    vStyle.VerticalAlignment = VerticalAlignment.Center;

                    #endregion

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

                var statistics = configuration.StatisticsConfigurations;
                var filterConfigs = configuration.FilterConfigurations;
                var freezeConfigs = configuration.FreezeConfigurations;

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
            for (int i = 0; i < columns.Count; i++)
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