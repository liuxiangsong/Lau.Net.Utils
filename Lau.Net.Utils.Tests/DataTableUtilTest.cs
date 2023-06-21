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
        public void AddSummaryRowTest()
        {
            // 创建 DataTable
            DataTable dt = new DataTable("MyTable");
            dt.Columns.Add("ID", typeof(int));
            dt.Columns.Add("Name", typeof(string));
            dt.Columns.Add("2012-1", typeof(decimal));

            // 添加数据行
            dt.Rows.Add(1, "Alice", 5000);
            dt.Rows.Add(2, "Bob", 6000);
            dt.Rows.Add(3, "Charlie", 7000);
             

            // 添加汇总行
            DataRow totalRow = dt.NewRow();
            totalRow["ID"] = DBNull.Value;
            totalRow["Name"] = "Total";
            totalRow["2012-1"] = dt.Compute("SUM([2012-1])", "");

            dt.Rows.Add(totalRow);
        }
    }
}
