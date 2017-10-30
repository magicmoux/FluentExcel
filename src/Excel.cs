﻿// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using NPOI.SS.UserModel;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Reflection;

    /// <summary>
    /// Provides some methods for loading <see cref="IEnumerable{T}"/> from excel.
    /// </summary>
    public static class Excel
    {
        private static IFormulaEvaluator _formulaEvaluator;

        /// <summary>
        /// Gets or sets the setting.
        /// </summary>
        /// <value>The setting.</value>
        public static WorkbookSettings Setting { get; set; } = new WorkbookSettings();

        /// <summary>
        /// Loading <see cref="IEnumerable{T}"/> from specified excel file. ///
        /// </summary>
        /// <typeparam name="T">The type of the model.</typeparam>
        /// <param name="excelFile">The excel file.</param>
        /// <param name="startRow">The row to start read.</param>
        /// <param name="sheetIndex">Which sheet to read.</param>
        /// <returns>The <see cref="IEnumerable{T}"/> loading from excel.</returns>
        public static IEnumerable<T> Load<T>(string excelFile, int startRow = 1, int sheetIndex = 0) where T : class, new()
        {
            if (!File.Exists(excelFile))
            {
                throw new FileNotFoundException();
            }

            var workbook = InitializeWorkbook(excelFile);

            // currently, only handle sheet one (or call side using foreach to support multiple sheet)
            var sheet = workbook.GetSheetAt(sheetIndex);

            // get the writable properties
            var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.SetProperty);

            bool fluentConfigEnabled = false;
            // get the fluent config
            if (Setting.FluentConfigs.TryGetValue(typeof(T).FullName, out var fluentConfig))
            {
                fluentConfigEnabled = true;
            }

            var propertyConfigurations = new ColumnConfiguration[properties.Length];
            for (var j = 0; j < properties.Length; j++)
            {
                var property = properties[j];
                if (fluentConfigEnabled && fluentConfig.PropertyConfigurations.TryGetValue(property.Name, out var pc))
                {
                    // fluent configure first(Hight Priority)
                    propertyConfigurations[j] = pc;
                }
                else
                {
                    propertyConfigurations[j] = null;
                }
            }

            var statistics = new List<StatisticsConfiguration>();
            if (fluentConfigEnabled)
            {
                statistics.AddRange(fluentConfig.StatisticsConfigurations);
            }

            var list = new List<T>();
            int idx = 0;

            IRow headerRow = null;

            // get the physical rows
            var rows = sheet.GetRowEnumerator();
            while (rows.MoveNext())
            {
                var row = rows.Current as IRow;

                if (idx == 0)
                    headerRow = row;
                idx++;

                if (row.RowNum < startRow)
                {
                    continue;
                }

                var item = new T();
                var itemIsValid = true;
                for (int i = 0; i < properties.Length; i++)
                {
                    var prop = properties[i];

                    int index = i;
                    var config = propertyConfigurations[i];
                    if (config != null)
                    {
                        if (config.IsImportIgnored)
                            continue;

                        index = config.Index;

                        // Try to autodiscover index from title and cache
                        if (index < 0 && config.AutoIndex && !string.IsNullOrEmpty(config.Title))
                        {
                            foreach (var cell in headerRow.Cells)
                            {
                                if (!string.IsNullOrEmpty(cell.StringCellValue))
                                {
                                    if (cell.StringCellValue.Equals(config.Title, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        index = cell.ColumnIndex;

                                        // cache
                                        config.Index = index;

                                        break;
                                    }
                                }
                            }
                        }

                        // check again
                        if (index < 0)
                        {
                            throw new ApplicationException("Please set the 'index' or 'autoIndex' by fluent api or attributes");
                        }
                    }

                    var value = row.GetCellValue(index, _formulaEvaluator);

                    // give a chance to the value converter.
                    if (config?.ValueConverter != null)
                    {
                        value = config.ValueConverter(value);
                    }

                    if (value == null)
                    {
                        continue;
                    }

                    // check whether is statics row
                    if (idx > startRow + 1 && index == 0
                        &&
                        statistics.Any(s => s.Name.Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase)))
                    {
                        var st = statistics.FirstOrDefault(s => s.Name.Equals(value.ToString(), StringComparison.InvariantCultureIgnoreCase));
                        var formula = row.GetCellValue(st.Columns.First()).ToString();
                        if (formula.StartsWith(st.Formula, StringComparison.InvariantCultureIgnoreCase))
                        {
                            itemIsValid = false;
                            break;
                        }
                    }

                    // property type
                    var propType = prop.PropertyType.UnwrapNullableType();

                    var safeValue = Convert.ChangeType(value, propType, CultureInfo.CurrentCulture);

                    prop.SetValue(item, safeValue, null);
                }

                if (itemIsValid)
                {
                    list.Add(item);
                }
            }

            return list;
        }

        #region TODO relocate into a "Util" class

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

        internal static object GetDefault(this Type type)
        {
            if (type.IsValueType)
            {
                return Activator.CreateInstance(type);
            }

            return null;
        }

        private static IWorkbook InitializeWorkbook(string excelFile)
        {
            var extension = Path.GetExtension(excelFile);
            var workbook = WorkbookFactory.Create(new FileStream(excelFile, FileMode.Open, FileAccess.Read));
            _formulaEvaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            return workbook;
        }

        #endregion
    }
}