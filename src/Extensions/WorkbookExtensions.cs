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

        //TODO implement this DRY
        public static WorkbookSettings WithWorksheets<T>(this WorkbookSettings settings, IGrouping<string, T> worksheetData, string name = null, params Expression<Func<FluentConfiguration<T>, ColumnConfiguration>>[] columns)
            where T : class
        {
            throw new NotSupportedException("Not supported yet");
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
            settings.ActiveWorksheetSettings = worksheet;
            return settings;
        }

        public static WorkbookSettings AdjustAutoIndex(this WorkbookSettings settings)
        {
            settings.ActiveWorksheetSettings.GetType().GetMethod("AdjustAutoIndex").Invoke(settings.ActiveWorksheetSettings, null);
            return settings;
        }

        public static WorkbookSettings HasFilter(this WorkbookSettings settings, int firstColumn, int lastColumn, int firstRow, int? lastRow = null)
        {
            settings.ActiveWorksheetSettings.GetType().GetMethod("HasFilter").Invoke(settings.ActiveWorksheetSettings, new object[] { firstColumn, lastColumn, firstRow, lastRow });
            return settings;
        }

        public static WorkbookSettings HasFreeze(this WorkbookSettings settings, int columnSplit, int rowSplit, int leftMostColumn, int topMostRow)
        {
            settings.ActiveWorksheetSettings.GetType().GetMethod("HasFreeze").Invoke(settings.ActiveWorksheetSettings, new object[] { columnSplit, rowSplit, leftMostColumn, topMostRow });
            return settings;
        }

        public static WorkbookSettings HasIgnoredProperties<T>(this WorkbookSettings settings, params Expression<Func<T, object>>[] propertyExpressions)
            where T : class
        {
            settings.ActiveWorksheetSettings.GetType().GetMethod("HasIgnoredProperties").Invoke(settings.ActiveWorksheetSettings, propertyExpressions);
            return settings;
        }

        public static WorkbookSettings HasStatistics(this WorkbookSettings settings, string name, string formula, params int[] columnIndexes)
        {
            settings.ActiveWorksheetSettings.GetType().GetMethod("HasStatistics").Invoke(settings.ActiveWorksheetSettings, new object[] { name, formula, columnIndexes });
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