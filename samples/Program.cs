using FluentExcel;
using FluentExcel.Extensions;
using System;
using System.IO;
using System.Linq;

namespace samples
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // global call this
            FluentConfiguration();

            var len = 20;
            var reports = new Report[len];
            for (int i = 0; i < len; i++)
            {
                reports[i] = new Report
                {
                    City = "ningbo",
                    Building = "世茂首府",
                    HandleTime = DateTime.Now.AddDays(7 * i),
                    Broker = "rigofunc 18957139**7",
                    Customer = "yingting 18957139**7",
                    Room = "2#1703",
                    Brokerage = 125 * i,
                    Profits = 25 * i
                };
            }

            string path = Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).FullName;
            if (Environment.OSVersion.Version.Major >= 6)
            {
                path = Directory.GetParent(path).ToString();
            }
            var excelFile = path + "/Documents/sample.xls";

            // save to excel file with multiple sheets based on expression
            reports.ToExcel(excelFile, r => r.HandleTime.Date.ToString("yyyy-MM"), overwrite: true);

            // save to excel file with multiple sheets based on maxRows
            reports.ToExcel(excelFile, "reports", 7, overwrite: true);

            // save to excel file
            reports.ToExcel(excelFile, overwrite: true);

            // Build a adhoc configuration
            new WorkbookSettings()
                .WithWorksheet(reports.GroupBy(r => r.HandleTime.Date.ToString("yyyy-MM")), "Reports",
                    // TODO Créer la méthode d'extension Column au lieu de Property
                    f => f.Column(r => r.City).IsMergeEnabled(),
                    f => f.Column(r => r.Building).HasExcelTitle("Building").IsMergeEnabled(),
                    f => f.Column(r => r.Area).HasExcelTitle("Area").IsIgnored(false, true),
                    f => f.Column(r => r.CustomerObj.Id), // TODO trouver comment evaluer le titre de la colonne à partir de l'expression
                    f => f.Column(r => r.HandleTime).HasExcelTitle("HandleTime").HasDataFormatter("yyyy-MM-dd"),
                    f => f.Column(r => r.Brokerage).HasDataFormatter("￥0.00"),
                    f => f.Column(r => r.Profits).HasDataFormatter("￥0.00")
                )
                // Configuration de la feuille Reports
                .HasStatistics("合计", "SUM", 5, 6)
                    .HasFilter(firstColumn: 0, lastColumn: 2, firstRow: 0)
                    .HasFreeze(columnSplit: 2, rowSplit: 1, leftMostColumn: 2, topMostRow: 1)
                // Passage à la feuille Customers
                .WithWorksheet(reports.Select(r => r.CustomerObj).Distinct(), "Customers",
                    f => f.Column(c => c.Id),
                    f => f.Column(c => c.FirstName),
                    f => f.Column(c => c.LastName)
                )
                // Configuration de la feuille Customers
                .HasStatistics("合计", "SUM", 6, 7)
                    .HasFilter(firstColumn: 0, lastColumn: 2, firstRow: 0)
                    .HasFreeze(columnSplit: 2, rowSplit: 1, leftMostColumn: 2, topMostRow: 1)
                .ToExcel(path + "/Documents/adhoc-samples.xls", overwrite: true)
            ;

            // load from excel
            var loadFromExcel = Excel.Load<Report>(excelFile);
        }

        /// <summary>
        /// Use fluent configuration api. (doesn't poison your POCO)
        /// </summary>
        private static void FluentConfiguration()
        {
            var fc = Excel.Setting.For<Report>();

            fc.HasStatistics("合计", "SUM", 6, 7)
              .HasFilter(firstColumn: 0, lastColumn: 2, firstRow: 0)
              .HasFreeze(columnSplit: 2, rowSplit: 1, leftMostColumn: 2, topMostRow: 1);

            fc.Column(r => r.City)
              .HasExcelIndex(0)
              .HasExcelTitle("城市")
              .IsMergeEnabled();

            // or
            //fc.Property(r => r.City).HasExcelCell(0,"城市", allowMerge: true);

            fc.Column(r => r.Building)
              .HasExcelIndex(1)
              .HasExcelTitle("楼盘")
              .IsMergeEnabled();

            // configures the ignore when exporting or importing.
            fc.Column(r => r.Area)
              .HasExcelIndex(8)
              .HasExcelTitle("Area")
              .IsIgnored(exportingIsIgnored: false, importingIsIgnored: true);

            // or
            //fc.Property(r => r.Area).IsIgnored(8, "Area", formatter: null, exportingIsIgnored: false, importingIsIgnored: true);

            fc.Column(r => r.HandleTime)
              .HasExcelIndex(2)
              .HasExcelTitle("成交时间")
              .HasDataFormatter("yyyy-MM-dd");

            // or
            //fc.Property(r => r.HandleTime).HasExcelCell(2, "成交时间", formatter: "yyyy-MM-dd", allowMerge: false);
            // or
            //fc.Property(r => r.HandleTime).HasExcelCell(2, "成交时间", "yyyy-MM-dd");

            fc.Column(r => r.Broker)
              .HasExcelIndex(3)
              .HasExcelTitle("经纪人");

            fc.Column(r => r.Customer)
              .HasExcelIndex(4)
              .HasExcelTitle("客户");

            fc.Column(r => r.Room)
              .HasExcelIndex(5)
              .HasExcelTitle("房源");

            fc.Column(r => r.Brokerage)
              .HasExcelIndex(6)
              .HasDataFormatter("￥0.00")
              .HasExcelTitle("佣金(元)");

            fc.Column(r => r.Profits)
              .HasExcelIndex(7)
              .HasExcelTitle("收益(元)");
        }
    }
}