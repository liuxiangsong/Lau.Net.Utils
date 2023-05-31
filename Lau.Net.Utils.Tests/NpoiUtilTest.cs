using Lau.Net.Utils.Excel;
using Lau.Net.Utils.Excel.NpoiExtensions;
using NPOI.SS.UserModel;
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
            


            sheet.DataTableToSheet(dt, 12);
            var filePath = @"E:\\test\1.xls";
            npoiUtil.Workbook.SaveToExcel(filePath);
        }

        private DataTable CreateTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("a");
            dt.Columns.Add("b");
            dt.Columns.Add("c");
            dt.Columns.Add("d", typeof(DateTime));
            for (int i = 0; i < 10; i++)
            {
                var row = dt.NewRow();
                row[0] = i;
                row[1] = "b" + i;
                row[2] = "c" + i;
                row[3] = DateTime.Now;
                dt.Rows.Add(row);
            }
            return dt;
        }
    }
}
