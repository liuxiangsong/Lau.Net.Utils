using HtmlAgilityPack;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lau.Net.Utils.Web.HtmlDocumentExtensions;
using Lau.Net.Utils.Web;
using Lau.Net.Utils.Excel;
using System.IO;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    internal class HtmlTest
    {
        [Test]
        public void SaveImageTest()
        {
            var dt = NpoiUtil.ExcelToDataTable(@"E:\test\test.xlsx");
            var html = HtmlUtil.ConvertToHtmlPage(dt);
            var bytes = HtmlUtil.ConvertHtmlToImageByte(html, 1024, 90);
            File.WriteAllBytes(@"E:\test\image1.jpg", bytes);
        }

        [Test]
        public void GetOrCreateNodeByXpathTest()
        {
            var htmlDoc = new HtmlDocument();
            var node = htmlDoc.DocumentNode.GetOrCreateNodeByXpath("//html/head/title");
            Assert.IsTrue(node.Name == "title");
            var html = htmlDoc.GetHtml();
        }

        [Test]
        public void ConvertToHtmlPageTest()
        {
            var dt = CreateTestTable();
            var html = HtmlUtil.ConvertToHtmlPage(dt, "标题文案");
            Assert.IsNotEmpty(html);
        }

        [Test]
        public void AddTitleForTableTest()
        {
            var dt = CreateTestTable();
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetOrCreateBodyNode().AppendDataTable(dt);
            var title = "<div>XX表单统计数据</div><div style='font-size: 12px;font-weight: normal;'>数据来源时间：2023-08-01</div>";
            tableNode.AddTitleForTable(title);
            tableNode.AddHtmlNodesAfterTableNode($"<div>test</div>");
            //var newNode = HtmlNode.CreateNode($"<div>test</div><div>testasdf</div>");
            //tableNode.ParentNode.InsertAfter(newNode, tableNode);
            var html = htmlDoc.GetHtml();
            Assert.IsNotEmpty(html);
        }

        [Test]
        public void InsertRowAndMergeCellTest()
        {
            var dt = CreateTestTable();
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetOrCreateBodyNode().AppendDataTable(dt);
            //获取表头第一行
            var firstHeaderRow = tableNode.SelectSingleNode("//thead/tr");
            var copyRow = firstHeaderRow.CloneNode(true);
            firstHeaderRow.ParentNode.InsertAfter(copyRow, firstHeaderRow);
            //tableNode.GetTableCell(0, 0).SetAttributeValue("rowspan", "2");
            //tableNode.GetTableCell(0, 1).SetAttributeValue("rowspan", "2");


            tableNode.MergeTableHeaderCells(0, 0, 2, 0);
            tableNode.MergeTableHeaderCells(0, 1, 2, 0);
            tableNode.MergeTableHeaderCells(0, 2, 0, 3, "合并表头");
            tableNode.MergeTableCells(0, 0, 2, 3);
            //firstHeaderRow.CopyFrom(firstHeaderRow);
            var html = htmlDoc.GetHtml();
        }

        [Test]
        public void SetTalbeColumnWidthTest()
        {
            var dt = CreateTestTable();
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetOrCreateBodyNode().AppendDataTable(dt);
            var dict = new Dictionary<int, int>
            {
                {0,40 },
                {1,60 },
                {3,80 }
            };
            tableNode.SetTalbeColumnWidth(htmlDoc,dict);
            var html = htmlDoc.GetHtml();
        }
        
        [Test]
        public void SetTableAlternateRowColorTest()
        {
            var dt = CreateTestTable();
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetOrCreateBodyNode().AppendDataTable(dt);
            //tableNode.SetTableAlternateRowColor("lightgreen");
            tableNode.SetTableAlternateRowColor("lightyellow", "lightgreen", new List<int> { 2,3,6});
            var html = htmlDoc.GetHtml();
        }

        [Test]
        public void SetNodeStyleTest()
        {
            var dt = CreateTestTable();
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetOrCreateBodyNode().AppendDataTable(dt);
            //设置第二、三【行】背景色
            tableNode.SetNodeStyle(".//tbody//tr[position() > 1 and position() <= 3]", "background-color:#fce4d6");
            //设置第一列中包含总计字样的【行】字体加粗
            tableNode.SetNodeStyle(".//tbody//tr[contains(td[1], '总计')]", "font-weight:bold");
            //设置第一列中文本等于12的【行】背景色
            tableNode.SetNodeStyle(".//tbody//tr[td[1]='12']", "background-color:#fce4d6");
            //设置第三列中包含负号的【单元格】为红色
            tableNode.SetNodeStyle(".//tbody//td[3][contains(text(), '-')]", "color:red");
            //设置第三列中包含数值大于0的【单元格】为蓝色
            tableNode.SetNodeStyle(".//tbody//td[3][number(.) > 0]", "color:blue");
            var html = htmlDoc.GetHtml();
            
        }

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
    }
}
