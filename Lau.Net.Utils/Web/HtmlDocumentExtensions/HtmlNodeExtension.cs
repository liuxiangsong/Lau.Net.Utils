using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Web.HtmlDocumentExtensions
{
    /// <summary>
    /// HtmlNode扩展方法
    /// </summary>
    public static class HtmlNodeExtension
    {
        /// <summary>
        /// 通过节点名称获取或创建节点
        /// </summary>
        /// <param name="parentNode">父节点</param>
        /// <param name="nodeName">节点名称</param>
        /// <returns>如果存在该节点则直接返回，否则创建节点</returns>
        public static HtmlNode GetOrCreateNodeByNodeName(this HtmlNode parentNode, string nodeName)
        {
            var xpath = $"//{nodeName}";
            var node = parentNode.SelectSingleNode(xpath);
            if (node == null)
            {
                node = parentNode.OwnerDocument.CreateElement(nodeName);
                parentNode.AppendChild(node);
            }
            return node;
        }

        /// <summary>
        /// 通过Xpath获取节点
        /// </summary>
        /// <param name="currentNode">当前节点</param>
        /// <param name="xpath">Xpath表达式</param>
        /// <returns>如果xpath中某些节点不存在则创建</returns>
        public static HtmlNode GetOrCreateNodeByXpath(this HtmlNode currentNode, string xpath)
        {
            var xpathParts = xpath.Split('/').Where(x=>!string.IsNullOrWhiteSpace(x));
            foreach (var xpathPart in xpathParts)
            {
                var nodeName = xpathPart;
                var attributeIndex = xpathPart.IndexOf("[@");
                if (attributeIndex >= 0)
                {
                    nodeName = xpathPart.Substring(0, attributeIndex);
                }

                var nextNode = currentNode.SelectSingleNode(nodeName);
                if (nextNode == null)
                {
                    var newNode = currentNode.OwnerDocument.CreateElement(nodeName);
                    currentNode.AppendChild(newNode);
                    if (attributeIndex >= 0)
                    {
                        var attributeValue = xpathPart.Substring(attributeIndex + 2, xpathPart.Length - attributeIndex - 3);
                        var attributeParts = attributeValue.Split('=');
                        var attributeName = attributeParts[0];
                        var attributeValueEscaped = attributeParts[1].Replace("\"", "");
                        newNode.Attributes.Add(attributeName, attributeValueEscaped);
                    }

                    nextNode = newNode;
                }
                currentNode = nextNode;
            }
            return currentNode;
        }

        /// <summary>
        /// 通过Html创建节点
        /// </summary>
        /// <param name="htmlNode"></param>
        /// <param name="html"></param>
        /// <returns></returns>
        public static HtmlNodeCollection CreateNodesByHtml(this HtmlNode htmlNode, string html)
        {
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);
            return htmlDoc.DocumentNode.ChildNodes;
        }

        #region Table相关方法

        /// <summary>
        /// 将DataTable添加至指定节点下
        /// </summary>
        /// <param name="parentNode">父节点</param>
        /// <param name="dataTable">数据表</param>
        /// <param name="columnPositonDict">列对齐配置：DataColumn和Postion(left、right)组成的字典</param>
        /// <param name="ignoreHeader">是否忽略dataTable的列头</param>
        /// <returns></returns>
        public static HtmlNode AppendDataTable(this HtmlNode parentNode, DataTable dataTable, Dictionary<string, string> columnPositonDict = null, bool ignoreHeader = false)
        {
            var tableHtml = HtmlUtil.ConvertToHtmlTable(dataTable, columnPositonDict, ignoreHeader);
            var tableNode = HtmlNode.CreateNode(tableHtml);
            parentNode.AppendChild(tableNode);
            return tableNode;
        }
         
        /// <summary>
        /// 合并Table表头单元格
        /// </summary>
        /// <param name="tableNode">table节点</param>
        /// <param name="rowIndex">要合并单元格的起始行，从0开始</param>
        /// <param name="colIndex">要合并单元格的起始列，从0开始</param>
        /// <param name="rowSpan">要合并的行数</param>
        /// <param name="colSpan">要合并的列数</param>
        /// <returns></returns>
        public static HtmlNode MergeTableHeaderCells(this HtmlNode tableNode, int rowIndex, int colIndex, int rowSpan, int colSpan)
        {
            MergeTableCells(tableNode, true, rowIndex, colIndex, rowSpan, colSpan);
            return tableNode;
        }
        /// <summary>
        /// 合并Table单元格
        /// </summary>
        /// <param name="tableNode">table节点</param>
        /// <param name="rowIndex">要合并单元格的起始行，从0开始</param>
        /// <param name="colIndex">要合并单元格的起始列，从0开始</param>
        /// <param name="rowSpan">要合并的行数</param>
        /// <param name="colSpan">要合并的列数</param>
        /// <returns></returns>
        public static HtmlNode MergeTableCells(this HtmlNode tableNode, int rowIndex, int colIndex, int rowSpan, int colSpan)
        {
            MergeTableCells(tableNode, false, rowIndex, colIndex, rowSpan, colSpan);
            return tableNode;
        }

        /// <summary>
        /// 为Table添加标题
        /// </summary>
        /// <param name="tableNode">table节点</param>
        /// <param name="title">标题内容：普通文本或者html文本</param>
        /// <param name="colspan">标题行占几列，默认为table第一行单元格的数量</param>
        /// <returns></returns>
        public static HtmlNode AddTitleForTable(this HtmlNode tableNode, string title, int colspan = 0)
        {
            if (!string.IsNullOrWhiteSpace(title))
            {
                //获取表头第一行
                var firstHeaderRow = tableNode.SelectSingleNode("//thead/tr");
                if (colspan == 0)
                {
                    var headerRow = tableNode.SelectSingleNode(".//tr[1]"); // 假设表头在第一行
                    colspan = headerRow.SelectNodes("th|td").Count;
                }
                var newNode = HtmlNode.CreateNode($"<th colspan='{colspan}'>{title}</th>");
                firstHeaderRow.ParentNode.InsertBefore(newNode, firstHeaderRow);
            }
            return tableNode;
        }

        /// <summary>
        /// 设置表格间隔行背景颜色
        /// </summary>
        /// <param name="tableNode">table节点</param>
        /// <param name="evenColor">奇数行的颜色，如果为空则不做处理</param>
        /// <param name="oddColor">偶数行的颜色，如果为空则不做处理</param>
        /// <param name="rowGroups">为null时，则间隔行设置颜色，否则按rowGroups里的数值按组间隔设置颜色
        /// 比如rowGroups=[2,3,5],则第1~2行为一组，第3~5行为一组，第6~11行为一组</param>
        /// <returns></returns>
        public static HtmlNode SetTableAlternateRowColor(this HtmlNode tableNode,string evenColor, string oddColor= "#fce4d6", List<int> rowGroups=null)
        {
            if (rowGroups == null)
            {
               var rowCount = tableNode.SelectNodes(".//tbody//tr").Count;
                rowGroups = Enumerable.Repeat(1, rowCount).ToList();
            }
            var rowIndex = 0;
            for (var i = 0; i < rowGroups.Count(); i++)
            {
                var xpath = $".//tbody//tr[position() > {rowIndex} and position() <= {rowIndex + rowGroups[i]}]";
                var isEvenRow = i % 2 == 0;
                if (isEvenRow && !string.IsNullOrEmpty(evenColor))
                {
                    tableNode.SetNodeStyle(xpath, $"background-color:{evenColor}");
                }
                else if(!isEvenRow && !string.IsNullOrEmpty(oddColor))
                {
                    tableNode.SetNodeStyle(xpath, $"background-color:{oddColor}");
                }
                rowIndex += rowGroups[i]; 
            }
            return tableNode;
        }

        /// <summary>
        /// 设置表格列宽度
        /// </summary>
        /// <param name="tableNode">表格HtmlNode</param>
        /// <param name="htmlDoc">HtmlDocument</param>
        /// <param name="colIndexWidth">表格列索引与宽度的键值对，宽度单位为像素</param>
        /// <returns></returns>
        public static HtmlNode SetTalbeColumnWidth(this HtmlNode tableNode,HtmlDocument htmlDoc, Dictionary<int,int> colIndexWidth)
        {
            var styleContent = "";
            foreach (var g in colIndexWidth.GroupBy(g => g.Value))
            {
               var styleName = string.Join(",", g.Select(x => $".col{x.Key}"));
               var style = $"{styleName} {{width:{g.Key}px;}} ";
               styleContent += style;
            }
            var headNode = htmlDoc.AddStyleNode(styleContent);            
            return tableNode;
        }
        #endregion

        /// <summary>
        /// 设置指定子节点的样式
        /// </summary>
        /// <param name="parentNode">父节点</param>
        /// <param name="xpath">XPath 表达式</param>
        /// <param name="style">css行内样式</param>
        /// <param name="overwrite ">是否覆盖之前的样式</param>
        /// <returns></returns>
        public static HtmlNode SetNodeStyle(this HtmlNode parentNode, string xpath,string style,bool overwrite = false)
        {
            var nodes = parentNode.SelectNodes(xpath);
            // 设置背景色
            if (nodes != null)
            {
                foreach (var node in nodes)
                {
                    var originalStyle = node.GetAttributeValue("style", "");
                    if (overwrite || string.IsNullOrEmpty(originalStyle))
                    {
                        node.SetAttributeValue("style", style);
                    }
                    else
                    {
                        node.SetAttributeValue("style", originalStyle + ";" + style);
                    }
                }
            }
            return parentNode;
        }

        #region  私有方法
        /// <summary>
        /// 合并表格单元格
        /// </summary>
        /// <param name="tableNode">表格节点</param>
        /// <param name="isMergeHeaderCell">是否是合并表头单元格</param>
        /// <param name="rowIndex"></param>
        /// <param name="colIndex"></param>
        /// <param name="rowSpan"></param>
        /// <param name="colSpan"></param>
        /// <returns></returns>
        private static HtmlNode MergeTableCells(HtmlNode tableNode, bool isMergeHeaderCell, int rowIndex, int colIndex, int rowSpan, int colSpan)
        {
            var rowXpath = "//tbody/tr";
            var cellTag = "td";
            if (isMergeHeaderCell)
            {
                rowXpath = "//thead/tr";
                cellTag = "th";
            }
            // 获取表格的所有行
            HtmlNodeCollection rows = tableNode.SelectNodes(rowXpath);
            if (rows == null || rows.Count <= rowIndex)
            {
                return tableNode;
            }
            // 获取要合并的单元格所在的行
            var rowNode = rows[rowIndex];
            // 获取该行的所有单元格
            HtmlNodeCollection cells = rowNode.SelectNodes($"{cellTag} | i[@data-deleted-tag='{cellTag}']");
            if (cells == null || cells.Count <= colIndex)
            {
                return tableNode;
            }
            // 获取要合并的第一个单元格
            HtmlNode firstCell = cells[colIndex];

            rowSpan = rowSpan > 0 ? rowSpan : 1;
            colSpan = colSpan > 0 ? colSpan : 1;
            if (rowSpan > rows.Count - rowIndex)
            {
                rowSpan = rows.Count - rowIndex;
            }
            if (colSpan > cells.Count - colIndex)
            {
                colSpan = cells.Count - colIndex;
            }

            // 修改第一个单元格的rowspan和colspan属性 
            firstCell.SetAttributeValue("rowspan", rowSpan.ToString());
            firstCell.SetAttributeValue("colspan", colSpan.ToString());

            // 移除其他要合并的单元格
            for (int i = 0; i < rowSpan; i++)
            {
                for (int j = 0; j < colSpan; j++)
                {
                    if (i == 0 && j == 0)
                    {
                        continue;  // 跳过第一个单元格
                    }
                    var currentRowIndex = rowIndex + i;
                    var currentColIndex = colIndex  + j;
                    HtmlNodeCollection curRowCells = rows[currentRowIndex].SelectNodes($"{cellTag} | i[@data-deleted-tag='{cellTag}']");
                    var cell = curRowCells[currentColIndex];
                    var replaceNode = HtmlNode.CreateNode($"<i data-deleted-tag='{cellTag}'></i>");
                    cell.ParentNode.ReplaceChild(replaceNode, cell);
                }
            }
            return tableNode;
        }

        ///// <summary>
        ///// 获取指定单元格
        ///// </summary>
        ///// <param name="tableNode"></param>
        ///// <param name="rowIndex">单元格所在行索引，从0开始</param>
        ///// <param name="colIndex">单元格所在列索引，从0开始</param>
        ///// <param name="isHeaderCell">是否是获取表头单元格</param>
        ///// <returns></returns>
        //public static HtmlNode GetTableCell(this HtmlNode tableNode, int rowIndex, int colIndex, bool isHeaderCell = false)
        //{
        //    var rowXpath = "//tr";
        //    var cellTag = "td";
        //    if (isHeaderCell)
        //    {
        //        rowXpath = "//thead/tr";
        //        cellTag = "th";
        //    }
        //    // 获取表格的所有行
        //    HtmlNodeCollection rows = tableNode.SelectNodes(rowXpath);
        //    if (rows == null || rows.Count <= rowIndex)
        //    {
        //        return null;
        //    }
        //    // 获取要合并的单元格所在的行
        //    HtmlNode rowNode = rows[rowIndex];
        //    // 获取该行的所有单元格
        //    HtmlNodeCollection cells = rowNode.SelectNodes(cellTag);
        //    if (cells == null || cells.Count <= colIndex)
        //    {
        //        return null;
        //    }
        //    return cells[colIndex];
        //}

        ///// <summary>
        ///// 合并表格单元格(针对所有合并的单元前面存在合并格的情况下有bug)
        ///// </summary>
        ///// <param name="tableNode">表格节点</param>
        ///// <param name="isMergeHeaderCell">是否是合并表头单元格</param>
        ///// <param name="rowIndex"></param>
        ///// <param name="colIndex"></param>
        ///// <param name="rowSpan"></param>
        ///// <param name="colSpan"></param>
        ///// <returns></returns>
        //private static HtmlNode MergeTableCells(HtmlNode tableNode, bool isMergeHeaderCell, int rowIndex, int colIndex, int rowSpan, int colSpan)
        //{
        //    var rowXpath = "//tr";
        //    var cellTag = "td";
        //    if (isMergeHeaderCell)
        //    {
        //        rowXpath = "//thead/tr";
        //        cellTag = "th";
        //    }
        //    // 获取表格的所有行
        //    HtmlNodeCollection rows = tableNode.SelectNodes(rowXpath);
        //    if (rows == null || rows.Count <= rowIndex)
        //    {
        //        return tableNode;
        //    }
        //    // 获取要合并的单元格所在的行
        //    HtmlNode rowNode = rows[rowIndex];
        //    // 获取该行的所有单元格
        //    HtmlNodeCollection cells = rowNode.SelectNodes(cellTag);
        //    if (cells == null || cells.Count <= colIndex)
        //    {
        //        return tableNode;
        //    }
        //    // 获取要合并的第一个单元格
        //    HtmlNode firstCell = cells[colIndex];

        //    rowSpan = rowSpan > 0 ? rowSpan : 1;
        //    colSpan = colSpan > 0 ? colSpan : 1;
        //    if (rowSpan > rows.Count - rowIndex)
        //    {
        //        rowSpan = rows.Count - rowIndex;
        //    }
        //    if (colSpan > cells.Count - colIndex)
        //    {
        //        colSpan = cells.Count - colIndex;
        //    }

        //    // 修改第一个单元格的rowspan和colspan属性 
        //    firstCell.SetAttributeValue("rowspan", rowSpan.ToString());
        //    firstCell.SetAttributeValue("colspan", colSpan.ToString());

        //    // 移除其他要合并的单元格
        //    for (int i = 0; i < rowSpan; i++)
        //    {
        //        var removeCells = new List<HtmlNode>();
        //        for (int j = 0; j < colSpan; j++)
        //        {
        //            if (i == 0 && j == 0)
        //            {
        //                continue;  // 跳过第一个单元格
        //            }
        //            var currentRowIndex = rowIndex + i;
        //            var currentColIndex = colIndex + 1 + j;
        //            var cell = rows[currentRowIndex].SelectSingleNode($"{cellTag}[{currentColIndex}]");
        //            removeCells.Add(cell);

        //        }
        //        removeCells.ForEach(cell => cell?.Remove());
        //    }
        //    return tableNode;
        //}
        #endregion
    }
}
