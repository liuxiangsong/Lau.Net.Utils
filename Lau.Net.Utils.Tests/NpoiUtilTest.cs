using Lau.Net.Utils.Excel;
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
            var filePath = @"E:\\MyCode\\test\1.xls";
            NpoiStaticUtil.DataTableToExcel(filePath, dt, true, NpoiStaticUtil.ExcelType.Xls);

            Assert.IsTrue(System.IO.File.Exists(filePath));
        }
    }
}
