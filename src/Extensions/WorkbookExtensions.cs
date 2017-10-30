using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

namespace FluentExcel.Extensions
{
    public static class WorkbookExtensions
    {
        private static Dictionary<Tuple<WorkbookSettings, string>, IEnumerable> worksheetsData = new Dictionary<Tuple<WorkbookSettings, string>, IEnumerable>();

        //TODO handle sheet groups when using GroupBy
        public static WorkbookSettings WithWorksheet<T>(this WorkbookSettings settings, IEnumerable<IGrouping<string, T>> worksheetData, string name = null, params Expression<Func<FluentConfiguration<T>, ColumnConfiguration>>[] columns)
            where T : class
        {
            foreach (var group in worksheetData)
            {
                WithWorksheet<T>(settings, group, group.Key, columns);
            }
            settings.CurrentSheetsSettings.UnionWith(settings.FluentConfigs.Where(cfg => worksheetData.Select(g => g.Key).Contains(cfg.Key)).Select(cfg => cfg.Value));
            return settings;
        }

        public static WorkbookSettings WithWorksheet<T>(this WorkbookSettings settings, IEnumerable<T> worksheetData, string name = null, params Expression<Func<FluentConfiguration<T>, ColumnConfiguration>>[] columns)
            where T : class
        {
            // Handles the sheet name automatically if none is provided
            if (string.IsNullOrWhiteSpace(name))
            {
                name = worksheetData.GetType().GetGenericArguments()[0].Name;
            }

            int sheetIndex = 0;
            string sheetName = name;
            while (settings.FluentConfigs.ContainsKey(sheetName))
            {
                sheetIndex++;
                sheetName = name + (sheetIndex > 0 ? sheetIndex.ToString() : "");
            }

            // Stores the datasource for the current sheet;
            worksheetsData[new Tuple<WorkbookSettings, string>(settings, sheetName)] = worksheetData;

            // Applies the columns definitions to the sheet
            int colIndex = 0;
            FluentConfiguration<T> worksheet;
            if (columns.Any())
            {
                worksheet = new FluentConfiguration<T>();
                settings.FluentConfigs[sheetName] = worksheet;

                foreach (var columnSettings in columns)
                {
                    ColumnConfiguration col = columnSettings.Compile().Invoke(worksheet);
                    col.HasExcelIndex(colIndex);
                    colIndex++;
                }
            }
            else
            {
                worksheet = Excel.Setting.For<T>();
                settings.FluentConfigs[sheetName] = worksheet;
            }
            // Sets the current configuration active sheet
            settings.CurrentSheetsSettings = new HashSet<IFluentConfiguration>() { worksheet };
            return settings;
        }

        public static WorkbookSettings AdjustAutoIndex(this WorkbookSettings settings)
        {
            settings.CurrentSheetsSettings.ToList().ForEach(s => s.GetType().GetMethod("AdjustAutoIndex").Invoke(s, null));
            return settings;
        }

        public static WorkbookSettings HasFilter(this WorkbookSettings settings, int firstColumn, int lastColumn, int firstRow, int? lastRow = null)
        {
            settings.CurrentSheetsSettings.ToList().ForEach(s => s.GetType().GetMethod("HasFilter").Invoke(s, new object[] { firstColumn, lastColumn, firstRow, lastRow }));
            return settings;
        }

        public static WorkbookSettings HasFreeze(this WorkbookSettings settings, int columnSplit, int rowSplit, int leftMostColumn, int topMostRow)
        {
            settings.CurrentSheetsSettings.ToList().ForEach(s => s.GetType().GetMethod("HasFreeze").Invoke(s, new object[] { columnSplit, rowSplit, leftMostColumn, topMostRow }));
            return settings;
        }

        public static WorkbookSettings HasIgnoredProperties<T>(this WorkbookSettings settings, params Expression<Func<T, object>>[] propertyExpressions)
            where T : class
        {
            settings.CurrentSheetsSettings.ToList().ForEach(s => s.GetType().GetMethod("HasIgnoredProperties").Invoke(s, propertyExpressions));
            return settings;
        }

        public static WorkbookSettings HasStatistics(this WorkbookSettings settings, string name, string formula, params int[] columnIndexes)
        {
            settings.CurrentSheetsSettings.ToList().ForEach(s => s.GetType().GetMethod("HasStatistics").Invoke(s, new object[] { name, formula, columnIndexes }));
            return settings;
        }

        //TODO create the ToExcelContent method as well
        public static void ToExcel(this WorkbookSettings settings, string excelFile, int maxRowsPerSheet = int.MaxValue, bool overwrite = false)
        {
            var worksheets = worksheetsData.Keys.Where(k => k.Item1 == settings).ToList();
            try
            {
                foreach (var key in worksheets)
                {
                    var sheetName = key.Item2;
                    var configuration = settings.FluentConfigs[sheetName];
                    var model = configuration.GetType().GetGenericArguments()[0];
                    var source = worksheetsData[key];
                    var toExcelImplementation = typeof(IEnumerableNpoiExtensions).GetMethods().Where(m => m.Name == "ToExcel" && m.GetCustomAttribute<DefaultImplementationAttribute>() != null).FirstOrDefault();
                    toExcelImplementation.MakeGenericMethod(model).Invoke(null, new object[] { source, excelFile, sheetName, maxRowsPerSheet, overwrite, configuration });
                }
            }
            finally
            {
                // Cleans up the worksheetsData dictionary
                foreach (var key in worksheets)
                {
                    worksheetsData.Remove(key);
                }
            }
        }
    }
}