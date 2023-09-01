using Lau.Net.Utils.Excel;
using Lau.Net.Utils.Excel.NpoiExtensions;
using NPOI.SS.UserModel;
using NPOI.SS.UserModel.Charts;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class NpoiUtilTest
    {
        [Test]
        public void ExcelToDataTableTest()
        { 
            var filePath = @"E:\\test\1.xlsx";
            var dt = NpoiStaticUtil.ExcelToDataTable(filePath);
            Assert.IsNotNull(dt);
        }

        [Test]
        public void DataTableToExcelTest()
        {
            var dt = CreateTable();
            var filePath = @"E:\\test\1.xls";
            NpoiStaticUtil.DataTableToExcel(filePath, dt, true, NpoiStaticUtil.ExcelType.Xls); 
            Assert.IsTrue(System.IO.File.Exists(filePath));
        }

        [Test]
        public void ModifyCellsStyle()
        {
            var dt = CreateTable();
            var npoiUtil = new NpoiUtil();
            var sheet = npoiUtil.DataTableToWorkbook(dt);
            var workbook = npoiUtil.Workbook;
            Action<ICellStyle> modifyCellStyle = cellStyle =>
            {
                var font = workbook.CreateFont(null).SetFontStyle(10, false, IndexedColors.Red.Index);
                cellStyle.SetFont(font);
            };
            sheet.ModifyCellsStyle( 2, -1, 0, -1, modifyCellStyle);
            var filePath = @"E:\\test\1.xls";
            npoiUtil.Workbook.SaveToExcel(filePath);
        }

        [Test]
        public void NpoiStyleTest()
        {
            var dt = CreateTable();
            var npoiUtil = new NpoiUtil();
            var sheet = npoiUtil.DataTableToWorkbook(dt);

            var style = npoiUtil.Workbook.CreateCellStyle();
            var font = npoiUtil.Workbook.CreateFont(null).SetFontStyle(14,true,IndexedColors.Red.Index);
            style.SetFont(font);
            //style.FillForegroundColor = IndexedColors.LightBlue.Index;
            //style.FillPattern = FillPattern.SolidForeground;
            //sheet.GetRow(2).RowStyle = style;
            //sheet.SetRowStyle(2, style); 
            sheet.InsertSheetByDataTable(dt, 12);
            var filePath = @"E:\\test\1.xls";
            npoiUtil.Workbook.SaveToExcel(filePath);
        }

        private DataTable CreateTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("月份");
            dt.Columns.Add("生产总数量", typeof(int));
            dt.Columns.Add("生产合格数", typeof(int));
            dt.Columns.Add("不良总数量", typeof(int));
            dt.Columns.Add("合格率" );
            Random random = new Random();
            for (int i = 0; i < 12; i++)
            {
                var row = dt.NewRow();
                int num = random.Next(0, 10);
                row["月份"] = (i + 1).ToString();
                var totalCount = random.Next(100, 1000);
                var goodCount = random.Next(1, totalCount);;
                row["生产总数量"] = totalCount;
                row["生产合格数"] = goodCount;
                row["不良总数量"] = totalCount - goodCount;
                row["合格率"] = string.Format("{0:0.00%}", (decimal)goodCount / totalCount);
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}
