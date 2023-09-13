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
            sheet.SetCellsStyle( 2, -1, 0, -1, modifyCellStyle);
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

        [Test]
        public void NpoiMergeCellsTest()
        {
            var dt = DataTableUtil.CreateTable("列1", "列2", "列3");
            dt.TableName = "asdf";
            dt.Rows.Add("M001", 2, 4);
            dt.Rows.Add("M001", 4, 24);
            dt.Rows.Add("M001", 4, 24);
            dt.Rows.Add("M001", 4, 24);
            dt.Rows.Add("M002", 2, 4);
            dt.Rows.Add("M002", 2, 4);
            dt.Rows.Add("M002", 2, 4);
            dt.Rows.Add("M003", 2, 4);
            dt.Rows.Add("M003", 2, 4);
            var counts = dt.AsEnumerable().GroupBy(r => r.Field<string>("列1")).Select(g => g.Count());
            var workbook = NpoiStaticUtil.DataTableToWorkBook(dt,true,0);
            var sheet = workbook.GetSheetAt(0);
              
            var currentRow = 1;
            var colorCellSytle = sheet.Workbook.CreateCellStyleWithBorder();
            colorCellSytle.SetCellBackgroundStyle(IndexedColors.LightGreen.Index);
  
            var cellStyle = sheet.Workbook.CreateCellStyleWithBorder(); ;
            cellStyle.SetCellAlignmentStyle(false, HorizontalAlignment.Left);
            //隔行设置行颜色
            var isSetRowStyle = false;
            foreach (var rowCount in counts)
            {
                if (isSetRowStyle)
                {
                    for (var i = 0; i < rowCount; i++)
                    {
                        sheet.SetRowStyle(currentRow + i, colorCellSytle);
                    }
                }
                if (rowCount > 1)
                { 
                    sheet.MergeCells(currentRow, currentRow + rowCount - 1, 0, 0, cellStyle);
                } 
                isSetRowStyle = !isSetRowStyle;
                currentRow += rowCount;
            }
            workbook.SaveToExcel(@"E:\\test\111.xls");
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
