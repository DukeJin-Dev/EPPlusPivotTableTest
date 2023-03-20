using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace EPPlusPivotTableTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var table = new List<Shown>();
            for (int i = 0; i < 200; i++)
            {
                table.Add(new Shown { Date = DateTime.Today.AddDays(i),Amount =i%5==0?0: (decimal)10000 });
            }

            var path = Environment.CurrentDirectory + "\\excel\\";
            Directory.CreateDirectory(path);

            var filePath = path + DateTime.Now.ToString("yyyyMMdd_hhmmss") + ".xlsx";
            var aaa = File.Open(filePath, FileMode.OpenOrCreate);
            // var fileInfo = new FileInfo(filePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var pck = new ExcelPackage(aaa))
            {
                var sheet = pck.Workbook.Worksheets.Add("data");

                var aa = sheet.Cells["A1"].LoadFromCollection(table,true);

                sheet.Cells[2,2, aa.End.Row,2].Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;
                sheet.Cells[2, 1, aa.End.Row, 1].Style.Numberformat.Format = "#,##0.00";

                var dataRange = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];

                CreatePivotTableWithDataGrouping(pck, dataRange);

                pck.Save();
            }
        }

        private static void CreatePivotTableWithDataGrouping(ExcelPackage pck, ExcelRangeBase dataRange)
        {
            var wsPivot = pck.Workbook.Worksheets.Add("PivotDateGrp");
            var pt = wsPivot.PivotTables.Add(wsPivot.Cells["B3"], dataRange, "Report");

            //Add a rowfield
            var rowField = pt.RowFields.Add(pt.Fields["Date"]);
            rowField.AddDateGrouping(eDateGroupBy.Years | eDateGroupBy.Months);

            //Add the data fields and format them
            ExcelPivotTableDataField dataField = pt.DataFields.Add(pt.Fields["Amount"]);
            dataField.Format = "#,##0.00";
            dataField.Name = "Sum of Amount";

            //We want the datafields to appear in columns
            pt.DataOnRows = false;

        }


    }

    public class Shown
    {
        public decimal? Amount { get; set; }
        public DateTime? Date { get; set; }

    }

}
