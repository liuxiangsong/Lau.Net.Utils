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
        public static DataTable CreateTestTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("月份");
            dt.Columns.Add("生产总数量", typeof(int));
            dt.Columns.Add("生产合格数", typeof(int));
            dt.Columns.Add("不良总数量", typeof(int));
            dt.Columns.Add("合格率");
            Random random = new Random();
            for (int i = 0; i < 7; i++)
            {
                var row = dt.NewRow();
                int num = random.Next(0, 10);
                row["月份"] = (i + 1).ToString();
                var totalCount = random.Next(100, 1000);
                var goodCount = random.Next(-100, totalCount); ;
                row["生产总数量"] = totalCount;
                row["生产合格数"] = goodCount;
                row["不良总数量"] = totalCount - goodCount;
                row["合格率"] = string.Format("{0:0.00%}", (decimal)goodCount / totalCount);
                dt.Rows.Add(row);
            }
            var summaryRow = DataTableUtil.CreateSummaryRow(dt);
            summaryRow[0] = "总计";
            dt.Rows.Add(summaryRow);
            return dt;
        }

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
            Assert.AreEqual(dt.Rows.Count, 4);
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

        [Test]
        public void AddIdentityColumnTest()
        {
            var dt = CreateTestTable();
            DataTableUtil.AddIdentityColumn(dt, "序号");
            dt.Rows.Add(dt.NewRow());
            dt.Rows.Add(dt.NewRow());
            Assert.AreEqual("序号", dt.Columns[0].ColumnName);
            Assert.AreEqual(dt.Rows.Count, dt.Rows[dt.Rows.Count - 1][0]);
        }

        [Test]
        public void GetValueTest()
        {
            var dt = CreateTestTable();
            var value = dt.Rows[0].GetValue<string>("Sdf");
        }
    }
}
