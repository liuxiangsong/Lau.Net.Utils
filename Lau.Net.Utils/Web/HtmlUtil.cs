using CoreHtmlToImage;
using HtmlAgilityPack;
using Lau.Net.Utils.Web.HtmlDocumentExtensions;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Web
{
    public static class HtmlUtil
    {
        /// <summary>
        /// 将DataTable转化为包含对应表格的html页面字符串
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="title">标题</param>
        /// <param name="ignoreHeader">是否忽略列头</param>
        /// <returns></returns>
        public static string ConvertToHtmlPage(DataTable dt,string title="", bool ignoreHeader = false)
        {
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetBodyNode().AppendDataTable(dt, ignoreHeader);
            if (!string.IsNullOrWhiteSpace(title))
            {
                //获取表头第一行
                var firstHeaderRow = tableNode.SelectSingleNode("//thead/tr");
                var newNode = HtmlNode.CreateNode($"<th colspan='{dt.Columns.Count}'>{title}</th>");
                firstHeaderRow.ParentNode.InsertBefore(newNode, firstHeaderRow);
            }
            var html = htmlDoc.GetHtml();
            return html;
        }

        /// <summary>
        /// 将DataTable转化为html 表格
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="title">标题</param>
        /// <param name="ignoreHeader">是否忽略dt中的列头</param>
        /// <returns></returns>
        public static string ConvertToHtmlTable(DataTable dt, bool ignoreHeader = false)
        {
            // 创建 HTML 表格元素
            var sb = new StringBuilder();
            sb.Append("<table border='1' cellspacing='0' cellpadding='5' style='width:100%;text-align: center;'>");
            if (!ignoreHeader)
            {
                // 创建表头行
                sb.Append("<thead><tr>");
                foreach (DataColumn column in dt.Columns)
                {
                    sb.AppendFormat("<th>{0}</th>", column.ColumnName);
                }
                sb.Append("</tr></thead>");
            }

            sb.Append("<tbody>");
            foreach (DataRow row in dt.Rows)
            {
                sb.Append("<tr>");
                foreach (DataColumn column in dt.Columns)
                {
                    sb.AppendFormat("<td>{0}</td>", row[column]);
                }
                sb.Append("</tr>");
            }
            sb.Append("</tbody>");
            sb.Append("</table>");
            return sb.ToString();
        }

        /// <summary>
        /// 将html转化为图片类型字节数组
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        public static byte[] ConvertHtmlToImageByte(string html)
        {
            var converter = new HtmlConverter();
            var bytes = converter.FromHtmlString(html);
            return bytes;
        }

        /// <summary>
        /// 将Datatable转化为图片字节
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="title">标题</param>
        /// <param name="ignoreHeader">是否忽略dt中的列头</param>
        /// <returns></returns>
        public static byte[] ConvertTableToImageByte(DataTable dt,string title="", bool ignoreHeader = false)
        {
            var html = ConvertToHtmlPage(dt, title, ignoreHeader);
            return ConvertHtmlToImageByte(html);
        }

        /// <summary>
        /// 将html转化为Image
        /// </summary>
        /// <param name="html"></param>
        /// <returns></returns>
        public static Image ConvertHtmlToImage(string html)
        {
            var bytes = ConvertHtmlToImageByte(html);
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                var image = Image.FromStream(ms);
                return image;
            }
        }
    }
}
