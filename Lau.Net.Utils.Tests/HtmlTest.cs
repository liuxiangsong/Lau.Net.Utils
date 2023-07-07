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
        public void HtmlDocumentTest()
        {
            var dt = CreateTestTable();
            var htmlDoc = new HtmlDocument(); 
            var tableNode = htmlDoc.GetBodyNode().AppendDataTable(dt); 
            //获取表头第一行
            var firstHeaderRow = tableNode.SelectSingleNode("//thead/tr");
            var headerCells = firstHeaderRow.SelectNodes("th");
            for (int i = 0; i < 4; i++)
            {  //给表头第一行前4个单元格添加合并行异性
                headerCells[i].SetAttributeValue("rowspan", "2");
            }
            var headNode = tableNode.SelectSingleNode("//thead");
            var newNode = tableNode.CreateNodesByHtml("<tr><th>测试</th><tr>")[0];
            //在表头的第一行后插入新的表头行
            headNode.InsertAfter(newNode, firstHeaderRow);
 
            
            tableNode.MergeTableCells(3, 1, 3, 0);
            var html = htmlDoc.GetHtml(); 
        }

        [Test]
        public void Test2()
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
                var goodCount = random.Next(1, totalCount); ;
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
