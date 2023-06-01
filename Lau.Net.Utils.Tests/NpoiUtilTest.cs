using Lau.Net.Utils.Excel;
using Lau.Net.Utils.Excel.NpoiExtensions;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
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
        public void NpoiTest()
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
            sheet.ModifyCellsStyle(workbook, 2, -1, 0, -1, modifyCellStyle);
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
        public void NpoiChartTest()
        {
            var dt = CreateTable();
            var npoiUtil = new NpoiUtil();
            var sheet = npoiUtil.DataTableToWorkbook(dt);



        }

        private DataTable CreateTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Category");
            dt.Columns.Add("Value1", typeof(int));
            dt.Columns.Add("Value2", typeof(int));
            dt.Columns.Add("Value3", typeof(int));
            Random random = new Random();
            int num = random.Next(0, 10); 
            for (int i = 0; i < 10; i++)
            {
                var row = dt.NewRow();
                row[0] = (char)('A' + num);
                row[1] = random.Next(0, 100 + 1);
                row[2] = random.Next(0, 100 + 1);
                row[3] = random.Next(0, 100 + 1);
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}
