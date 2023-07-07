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
        /// <param name="htmlDoc"></param>
        /// <param name="dt">DataTable</param>
        /// <param name="ignoreHeader">是否忽略列头</param>
        /// <returns></returns>
        public static string ConvertToHtmlPage(DataTable dt, bool ignoreHeader = false)
        {
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetBodyNode().AppendDataTable(dt, ignoreHeader);
            var html = htmlDoc.GetHtml();
            return html;
        }

        /// <summary>
        /// 将DataTable转化为html 表格
        /// </summary>
        /// <param name="dataTable"></param>
        /// <param name="ignoreHeader">是否忽略列头</param>
        /// <returns></returns>
        public static string ConvertToHtmlTable(DataTable dataTable, bool ignoreHeader = false)
        {
            // 创建 HTML 表格元素
            var sb = new StringBuilder();
            sb.Append("<table border='1' cellspacing='0' cellpadding='5'>");
            if (!ignoreHeader)
            {
                // 创建表头行
                sb.Append("<thead><tr>");
                foreach (DataColumn column in dataTable.Columns)
                {
                    sb.AppendFormat("<th>{0}</th>", column.ColumnName);
                }
                sb.Append("</tr></thead>");
            }

            sb.Append("<tbody>");
            foreach (DataRow row in dataTable.Rows)
            {
                sb.Append("<tr>");
                foreach (DataColumn column in dataTable.Columns)
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
