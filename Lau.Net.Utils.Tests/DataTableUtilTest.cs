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
    public class DataTableUtilTest
    {
        [Test]
        public void CreateTableTest()
        {
            var cols = new string[] { "序号", "种类", "数量|int" };
            var dt = DataTableUtil.CreateTable(cols);
            Assert.AreEqual(dt.Columns.Count, 3);
        }

        [Test]
        public void InsertColumnsAfterTest()
        {
            var cols = new string[] { "序号", "种类", "数量|int" };
            var dt = DataTableUtil.CreateTable(cols);
            DataTableUtil.InsertColumnsAfter(dt, "序号", 3);
            Assert.IsTrue(dt.Columns.Count == 6);

            var newCols = new string[] { "A", "B" };
            DataTableUtil.InsertColumnsAfter(dt, "序号", newCols);
            Assert.IsTrue(dt.Columns[1].ColumnName == "A");

            var newCols2 = new string[] { "C", "D" };
            DataTableUtil.InsertColumnsAfter(dt, "种类", newCols2);
            Assert.IsTrue(dt.Columns[7].ColumnName == "C");
        }

        [Test]
        public void CreateSummaryRowTest()
        {
            // 创建 DataTable
            var cols = new string[] { "name", "种类|int", "数量|int" };
            var dt = DataTableUtil.CreateTable(cols);
            dt.Rows.Add("adsf", 2, 3);
            dt.Rows.Add("adf", 2, 3);
            dt.Rows.Add("adf", 1, 1);
            DataTableUtil.CreateSummaryRow(dt, true, null, "name = 'adf'");
        }

        [Test]
        public void CopyDataRowToTableTest()
        {
            var cols = new string[] { "name", "种类|int", "数量|int" };
            var dt = DataTableUtil.CreateTable(cols);
            dt.Rows.Add("adsf", 2, 3);
            dt.Rows.Add("adf", 2, 3);
            dt.Rows.Add("adf", 1, 1);

            var cols2 = new string[] { "name", "种类34|int", "数量|int" };
            var dt2 = DataTableUtil.CreateTable(cols2);
            DataTableUtil.CopyDataRowToTable(dt2, dt.Rows[0]);
            Assert.AreEqual(1, dt2.Rows.Count);
        }
    }
}
