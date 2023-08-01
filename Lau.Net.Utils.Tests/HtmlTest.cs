using HtmlAgilityPack;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks; 
using Lau.Net.Utils.Web.HtmlDocumentExtensions;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    internal class HtmlTest
    {
        [Test]
        public void GetOrCreateNodeByXpathTest()
        {
            var htmlDoc = new HtmlDocument();
            var node =  htmlDoc.DocumentNode.GetOrCreateNodeByXpath("//html/head/title");
            Assert.IsTrue(node.Name == "title");
            var html = htmlDoc.GetHtml();
        }

        [Test]
        public void InsertRowAndMergeCellTest()
        {
            var dt = CreateTestTable();
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetBodyNode().AppendDataTable(dt);
            //获取表头第一行
            var firstHeaderRow = tableNode.SelectSingleNode("//thead/tr");
            var copyRow = firstHeaderRow.CloneNode(true);
            firstHeaderRow.ParentNode.InsertAfter(copyRow, firstHeaderRow);
            //tableNode.GetTableCell(0, 0).SetAttributeValue("rowspan", "2");
            //tableNode.GetTableCell(0, 1).SetAttributeValue("rowspan", "2");


            tableNode.MergeTableHeaderCells(0, 0, 2, 0);
            tableNode.MergeTableHeaderCells(0, 1, 2, 0);
            tableNode.MergeTableHeaderCells(0, 2, 0, 3);
            //firstHeaderRow.CopyFrom(firstHeaderRow);
            var html = htmlDoc.GetHtml();
        }

        [Test]
        public void SetNodeStyleTest()
        {
            var dt = CreateTestTable();
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetBodyNode().AppendDataTable(dt);
            //第一列中包含总计字样的行设置字体加粗
            tableNode.SetNodeStyle(".//tr[contains(td[1], '总计')]", "font-weight:bold");
            //第一列中单元格内文本等于12的设置背景色
            tableNode.SetNodeStyle(".//tr[td[1]='12']", "background-color:#fce4d6");
            //第三列中单元格包含负数的设置为红色字段
            tableNode.SetNodeStyle(".//td[3][contains(text(), '-')]", "color:red");
            var html = htmlDoc.GetHtml();
        }

        private DataTable CreateTestTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("月份");
            dt.Columns.Add("生产总数量", typeof(int));
            dt.Columns.Add("生产合格数", typeof(int));
            dt.Columns.Add("不良总数量", typeof(int));
            dt.Columns.Add("合格率");
            Random random = new Random();
            for (int i = 0; i < 12; i++)
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
    }
}
