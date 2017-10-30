// Copyright (c) rigofunc (xuyingting). All rights reserved.

namespace FluentExcel
{
    using NPOI.SS.UserModel;
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Linq.Expressions;

    /// <summary>
    /// Defines some extensions for <see cref="IEnumerable{T}"/> that using NPOI to provides excel functionality.
    /// </summary>
    public static class IEnumerableNpoiExtensions
    {
        public static byte[] ToExcelContent<T>(this IEnumerable<T> source, string sheetName = "sheet0", int maxRowsPerSheet = int.MaxValue, bool overwrite = false, IFluentConfiguration configuration = null)
            where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            IWorkbook book = Utils.InitializeWorkbook(null);
            using (MemoryStream ms = new MemoryStream())
            {
                source.ToWorksheet(book, sheetName, maxRowsPerSheet, overwrite, configuration);
                book.Write(ms);
                return ms.ToArray();
            }
        }

        public static byte[] ToExcelContent<T>(this IEnumerable<T> source, Expression<Func<T, string>> sheetSelector, int maxRowsPerSheet = int.MaxValue, bool overwrite = false, IFluentConfiguration configuration = null)
             where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            IWorkbook book = Utils.InitializeWorkbook(null);
            using (MemoryStream ms = new MemoryStream())
            {
                foreach (var sheet in source.AsQueryable().GroupBy(sheetSelector == null ? s => null : sheetSelector))
                {
                    sheet.ToWorksheet(book, sheet.Key, maxRowsPerSheet, overwrite, configuration);
                    book.Write(ms);
                }
                return ms.ToArray();
            }
        }

        public static void ToExcel<T>(this IEnumerable<T> source, string excelFile, string sheetName = "sheet0", int maxRowsPerSheet = int.MaxValue, bool overwrite = false, IFluentConfiguration configuration = null)
            where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            IWorkbook book = Utils.InitializeWorkbook(excelFile);
            using (Stream ms = new FileStream(excelFile, FileMode.OpenOrCreate, FileAccess.Write))
            {
                source.ToWorksheet(book, sheetName, maxRowsPerSheet, overwrite, configuration);
                book.Write(ms);
            }
        }

        public static void ToExcel<T>(this IEnumerable<T> source, string excelFile, Expression<Func<T, string>> sheetSelector, int maxRowsPerSheet = int.MaxValue, bool overwrite = false, IFluentConfiguration configuration = null)
            where T : class
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            IWorkbook book = Utils.InitializeWorkbook(excelFile);
            using (Stream ms = new FileStream(excelFile, FileMode.OpenOrCreate, FileAccess.Write))
            {
                foreach (var sheet in source.AsQueryable().GroupBy(sheetSelector == null ? s => null : sheetSelector))
                {
                    sheet.ToWorksheet(book, sheet.Key, maxRowsPerSheet, overwrite, configuration);
                }
                book.Write(ms);
            }
        }
    }
}