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
using System.Drawing;
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
            var dt = NpoiUtil.ExcelToDataTable(filePath);
            Assert.IsNotNull(dt);
        }

        [Test]
        public void DataTableToExcelTest()
        {
            var dt = CreateTable();
            var filePath = @"E:\\test\1.xls";
            NpoiUtil.DataTableToExcel(filePath, dt, dateFormat:"yyyy-MM-dd HH:mm:ss", type:NpoiUtil.ExcelType.Xls);
            Assert.IsTrue(System.IO.File.Exists(filePath));
        }

        [Test]
        public void InsertSheetByDataTableTest()
        {
            var dt = CreateTable();
            var filePath = @"E:\\test\1.xlsx";
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            var workbook = NpoiUtil.CreateWorkbook();

            //公用样式定义在外面，即可复用样式对象
            var cellStyle = workbook.CreateCellStyleWithBorder();
            var centerStyle = workbook.CreateCellStyleWithBorder().SetCellAlignmentStyle(true);
            var shortDateStyle = workbook.CreateCellStyleWithBorder().SetCellDataFormat(workbook, "yyyy-MM-dd");
            var longDateStyle = workbook.CreateCellStyleWithBorder().SetCellDataFormat(workbook, "yyyy-MM-dd HH:mm:ss");
            var setBodyCellStyle = new Func<int, int, ICellStyle>((rowIndex, colIndex) =>
            {
                switch (colIndex)
                {
                    case 0:
                        return centerStyle;
                    case 5:
                        return shortDateStyle;
                    case 6:
                        return longDateStyle;
                    default:
                        return cellStyle;
                } 
            });
            var sheet = workbook.InsertSheetByDataTable(dt,setBodyCellStyle: setBodyCellStyle);
            //设置前2列和前1行冻结
            sheet.CreateFreezePane(2, 1);
            workbook.SaveToExcel(filePath);
            sw.Stop();
            var ms = sw.ElapsedMilliseconds;
            Assert.IsTrue(System.IO.File.Exists(filePath));
        }

        [Test]
        public void ModifyCellsStyle()
        {
            var dt = CreateTable();
            var workbook = NpoiUtil.CreateWorkbook(dt);
            var sheet = workbook.GetSheetAt(0); 
            Action<ICellStyle> modifyCellStyle = cellStyle =>
            {
                var font = workbook.CreateFont(null).SetFontStyle(10, false, IndexedColors.Red.Index);
                cellStyle.SetFont(font);
            };
            sheet.SetCellsStyle( 2, -1, 0, -1, modifyCellStyle);
            var filePath = @"E:\\test\1.xls";
            workbook.SaveToExcel(filePath);
        }

        [Test]
        public void NpoiStyleTest()
        {
            var dt = CreateTable();
            var workbook = NpoiUtil.CreateWorkbook(dt);
            var sheet = workbook.GetSheetAt(0);

            var style = workbook.CreateCellStyle();
            style.SetCellBackgroundStyle(IndexedColors.LightGreen.Index);
            style.SetCellDataFormat(sheet.Workbook, "[DbNum2][$-804]General");
            sheet.GetOrCreateCell(1, 3).CellStyle = style;
            //sheet.GetOrCreateCell(1, 4).CellStyle = style;
            sheet.GetOrCreateCell(2, 5).CellStyle = style;
            //var font = npoiUtil.Workbook.CreateFont(null).SetFontStyle(14,true,IndexedColors.Red.Index);
            //style.SetFont(font);
            //style.FillForegroundColor = IndexedColors.LightBlue.Index;
            //style.FillPattern = FillPattern.SolidForeground;
            //sheet.GetRow(2).RowStyle = style;
            //sheet.SetRowStyle(2, style);  
            var filePath = @"E:\\test\NpoiStyleTest.xlsx";
            workbook.SaveToExcel(filePath);
        }

        [Test]
        public void SetCellStyleByConditionTest()
        {
            var dt = CreateTable();
            var workbook = NpoiUtil.CreateWorkbook(dt);
            var sheet = workbook.GetSheetAt(0);
            var redCellStyle = workbook.CreateCellStyleWithBorder().SetCellBackgroundStyle(IndexedColors.Red.Index);
            sheet.SetCellStyleByCondition(1, 2, value => value.As<int>() > 554, redCellStyle);
            var filePath = @"E:\\test\SetCellStyleByConditionTest.xlsx";
            workbook.SaveToExcel(filePath);
        }

        [Test]
        public void NpoiMergeCellsTest()
        {
            var dt = DataTableUtil.CreateTable("列1", "列2|int", "测试自定义列宽阿斯蒂芬");
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
            var workbook = NpoiUtil.DataTableToWorkBook(dt,1);
            var sheet = workbook.GetSheetAt(0);

            //设置求和
            sheet.GetOrCreateCell(0, 0).SetCellValue("总数量");
            sheet.GetOrCreateCell(0, 1).SetCellFormulaForSum(2,dt.Rows.Count+1,1);

            // //设置列宽
            // int columnWidth = sheet.GetColumnWidth(2);
            // var adjustedWidth = columnWidth - (6 * 256);
            // sheet.SetColumnWidth(2, adjustedWidth);

            var currentRow = 2;
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
                    var sum = sheet.GetCellsValueWithSum(currentRow, currentRow + rowCount - 1, 1, 1);
                    sheet.MergeCells(currentRow, currentRow + rowCount - 1, 1, 1, cellStyle,sum);
                } 
                isSetRowStyle = !isSetRowStyle;
                currentRow += rowCount;
            }
            workbook.SaveToExcel(@"E:\\test\111.xls");
        }

        [Test]
        public void InsertImageTest()
        {
            var dt = CreateTable();
            var workbook = NpoiUtil.CreateWorkbook(dt);
            var sheet = workbook.GetSheetAt(0); 
            var img = Image.FromFile("E:\\test\\logo.png");
            var bytes = ImageUtil.ToBytes(img);
            sheet.InsertImage(bytes, 1, 3, 1, 4);
            var filePath = @"E:\\test\img.xls";
            workbook.SaveToExcel(filePath);
        }
        
        private DataTable CreateTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("月份");
            dt.Columns.Add("生产总数量", typeof(int));
            dt.Columns.Add("生产合格数", typeof(int));
            dt.Columns.Add("不良总数量", typeof(int));
            dt.Columns.Add("合格率",typeof(decimal) );
            dt.Columns.Add("日期",typeof(DateTime));
            dt.Columns.Add("日期2", typeof(DateTime));
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
                row["合格率"] =  (decimal)goodCount / totalCount;
                row["日期"] = DateTime.Now;
                row["日期2"] = DateTime.Now;
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}
