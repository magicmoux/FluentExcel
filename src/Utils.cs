﻿using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace FluentExcel
{
    internal static class Utils
    {
        public static void ToWorksheet<T>(this IEnumerable<T> source, IWorkbook workbook, string sheetName, int maxRowsPerSheet = int.MaxValue, bool overwrite = true, IFluentConfiguration configuration = null)
            where T : class
        {
            int sheetIndex = 0;
            var content = source.Where(i => i != null);
            while (sheetIndex == 0 || content.Any())
            {
                content.Take(maxRowsPerSheet).BuildWorksheet(workbook, sheetName + (sheetIndex > 0 ? "_" + sheetIndex.ToString() : ""), overwrite, configuration);
                sheetIndex++;
                content = content.Skip(maxRowsPerSheet);
            }
        }

        private static void BuildWorksheet<T>(this IEnumerable<T> source, IWorkbook workbook, string sheetName, bool overwrite = true, IFluentConfiguration configuration = null)
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
                if (!string.IsNullOrEmpty(colConfig.Formatter))
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

                    var unwrapType = valueProvider.Method.ReturnType.UnwrapNullableType();
                    object value = null;
                    try
                    {
                        value = valueProvider.DynamicInvoke(item);
                    }
                    catch (TargetInvocationException)
                    {
                        // ignore null reference exceptions as empty cells
                    }

                    // give a chance to the value converter even though value is null.
                    if (colConfig.ValueConverter != null)
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
                    else if (!string.IsNullOrEmpty(colConfig.Formatter) && value is IFormattable fv)
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
                            var value = sheet.GetRow(row).GetCellValue(config.Index, workbook.GetCreationHelper().CreateFormulaEvaluator());
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
        }

        internal static IWorkbook InitializeWorkbook(string excelFile = null)
        {
            //TODO check the file's path is valid
            if (!string.IsNullOrWhiteSpace(excelFile) && File.Exists(excelFile))
            {
                var extension = Path.GetExtension(excelFile);
                var workbook = WorkbookFactory.Create(new FileStream(excelFile, FileMode.Open, FileAccess.Read));
                return workbook;
            }
            var setting = Excel.Setting;
            if (!string.IsNullOrWhiteSpace(excelFile))
            {
                if (Path.GetExtension(excelFile).ToLower() == ".xlsx")
                {
                    setting.UserXlsx = true;
                }
                else if (Path.GetExtension(excelFile).ToLower() == ".xls")
                {
                    setting.UserXlsx = false;
                }
                else
                {
                    throw new NotSupportedException($"Not a workbook : {excelFile}");
                }
            }
            if (setting.UserXlsx)
            {
                var workbook = new XSSFWorkbook();
                return workbook;
            }
            else
            {
                var workbook = new HSSFWorkbook();
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

        internal static string GetCellPosition(int row, int col)
        {
            col = Convert.ToInt32('A') + col;
            row = row + 1;
            return ((char)col) + row.ToString();
        }

        internal static object GetCellValue(this IRow row, int index, IFormulaEvaluator eval = null)
        {
            var cell = row.GetCell(index);
            if (cell == null)
            {
                return null;
            }

            return cell.GetCellValue(eval);
        }

        internal static object GetCellValue(this ICell cell, IFormulaEvaluator eval = null)
        {
            if (cell.IsMergedCell)
            {
                // what can I do here?
            }

            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    else
                    {
                        return cell.NumericCellValue;
                    }
                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Boolean:
                    return cell.BooleanCellValue;

                case CellType.Error:
                    return FormulaError.ForInt(cell.ErrorCellValue).String;

                case CellType.Formula:
                    if (eval != null)
                        return GetCellValue(eval.EvaluateInCell(cell));
                    else
                        return cell.CellFormula;

                case CellType.Blank:
                case CellType.Unknown:
                default:
                    return null;
            }
        }

        /// <summary>
        /// Builds the column title as a property name path from the expression
        /// </summary>
        /// <param name="expr">The expr.</param>
        /// <param name="separator">The separator.</param>
        /// <returns></returns>
        internal static string GetColumnTitle(LambdaExpression expr, string separator = " ")
        {
            var stack = new Stack<string>();

            MemberExpression me;
            switch (expr.Body.NodeType)
            {
                case ExpressionType.Convert:
                case ExpressionType.ConvertChecked:
                    var ue = expr.Body as UnaryExpression;
                    me = ((ue != null) ? ue.Operand : null) as MemberExpression;
                    break;

                default:
                    me = expr.Body as MemberExpression;
                    break;
            }

            while (me != null)
            {
                stack.Push(me.Member.GetCustomAttribute<DisplayAttribute>(true)?.Name ?? me.Member.Name);
                me = me.Expression as MemberExpression;
            }

            return string.Join(separator, stack.ToArray());
        }

        internal static PropertyInfo GetPropertyInfo<TModel, TProperty>(Expression<Func<TModel, TProperty>> propertyExpression)
        {
            if (propertyExpression.NodeType != ExpressionType.Lambda)
            {
                throw new ArgumentException($"{nameof(propertyExpression)} must be lambda expression", nameof(propertyExpression));
            }

            var lambda = (LambdaExpression)propertyExpression;

            var memberExpression = ExtractMemberExpression(lambda.Body);
            if (memberExpression == null)
            {
                throw new ArgumentException($"{nameof(propertyExpression)} must be lambda expression", nameof(propertyExpression));
            }

            if (memberExpression.Member.DeclaringType == null)
            {
                throw new InvalidOperationException("Property does not have declaring type");
            }

            return memberExpression.Member.DeclaringType.GetProperty(memberExpression.Member.Name);
        }

        private static MemberExpression ExtractMemberExpression(Expression expression)
        {
            if (expression.NodeType == ExpressionType.MemberAccess)
            {
                return ((MemberExpression)expression);
            }

            if (expression.NodeType == ExpressionType.Convert)
            {
                var operand = ((UnaryExpression)expression).Operand;
                return ExtractMemberExpression(operand);
            }

            return null;
        }

        internal static object GetDefault(this Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }

            return null;
        }
    }
}