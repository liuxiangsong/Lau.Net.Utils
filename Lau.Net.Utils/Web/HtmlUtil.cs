using CoreHtmlToImage;
using HtmlAgilityPack;
using Lau.Net.Utils.Net;
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
    /// <summary>
    /// Html工具类
    /// </summary>
    public static class HtmlUtil
    {
        /// <summary>
        /// 根据URL获取HtmlDocument对象
        /// </summary>
        /// <param name="url">网页URL</param>
        /// <returns>HtmlDocument对象</returns>
        public static async Task<HtmlDocument> GetHtmlDocumentAsync(string url)
        {
            //var client = new System.Net.WebClient() { Encoding = Encoding.UTF8 };
            //var html = client.DownloadString(url);
            var html = await RestSharpUtil.GetAsync<string>(url);
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(html);
            return htmlDoc;
        } 
        /// <summary>
        /// 将DataTable转化为包含对应表格的html页面字符串
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="title">标题</param>
        /// <param name="columnPositonDict">列对齐配置：DataColumn和Postion(left、right)组成的字典</param>
        /// <param name="ignoreHeader">是否忽略列头</param>
        /// <returns></returns>
        public static string ConvertToHtmlPage(DataTable dt, string title = "", Dictionary<string, string> columnPositonDict = null, bool ignoreHeader = false)
        {
            var htmlDoc = new HtmlDocument();
            var tableNode = htmlDoc.GetOrCreateBodyNode().AppendDataTable(dt, columnPositonDict, ignoreHeader);
            tableNode.AddTitleForTable(title, dt.Columns.Count);
            var html = htmlDoc.GetHtml();
            return html;
        }

        /// <summary>
        /// 将DataTable转化为html 表格
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="columnPositonDict">列对齐配置：DataColumn和Postion(left、right)组成的字典</param>
        /// <param name="ignoreHeader">是否忽略dt中的列头</param>
        /// <returns></returns>
        public static string ConvertToHtmlTable(DataTable dt, Dictionary<string, string> columnPositonDict = null, bool ignoreHeader = false)
        {
            // 创建 HTML 表格元素
            var sb = new StringBuilder();
            sb.Append("<table border='1' cellspacing='0' cellpadding='5' style='width:100%;text-align: center;'>");
            if (!ignoreHeader)
            {
                // 创建表头行
                sb.Append("<thead><tr>");
                for (var index = 0; index < dt.Columns.Count; index++)
                {
                    var column = dt.Columns[index];
                    sb.Append($"<th class=\"col{index}\">{column.ColumnName}</th>");
                }
                sb.Append("</tr></thead>");
            }

            sb.Append("<tbody>");
            foreach (DataRow row in dt.Rows)
            {
                sb.Append("<tr>");
                for (var index = 0; index < dt.Columns.Count; index++)
                {
                    var column = dt.Columns[index];
                    var positon = columnPositonDict.GetValue(column.ColumnName);
                    var style = "";
                    if (!string.IsNullOrEmpty(positon))
                    {
                        style = $" style=\"text-align: {positon};\"";
                    }
                    sb.Append($"<td  class=\"col{index}\"{style}>{row[column]}</td>");
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
        /// <param name="imgWidth">图片宽度：默认1024</param>
        /// <param name="imgQuality">图片的清晰度：默认为100</param>
        /// <returns></returns>
        public static byte[] ConvertHtmlToImageByte(string html, int imgWidth = 1024, int imgQuality = 100)
        {
            var converter = new HtmlConverter();
            var format = ImageFormat.Jpg;
            //if(!isPng)
            //{
            //    format = ImageFormat.Png;
            //}
            var bytes = converter.FromHtmlString(html, imgWidth, format, imgQuality);
            return bytes;
        }

        /// <summary>
        /// 将Datatable转化为图片字节
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="title">标题</param>
        /// <param name="ignoreHeader">是否忽略dt中的列头</param>
        /// <param name="imgWidth">图片宽度：默认1024</param>
        /// <param name="imgQuality">图片的清晰度：默认为100</param>
        /// <returns></returns>
        public static byte[] ConvertTableToImageByte(DataTable dt, string title = "", bool ignoreHeader = false, int imgWidth = 1024, int imgQuality = 100)
        {
            var html = ConvertToHtmlPage(dt, title, null, ignoreHeader);
            return ConvertHtmlToImageByte(html, imgWidth, imgQuality);
        }

        /// <summary>
        /// 将html转化为Image
        /// </summary>
        /// <param name="html"></param>
        /// <param name="imgWidth">图片宽度：默认1024</param>
        /// <param name="imgQuality">图片的清晰度：默认为100</param>
        /// <returns></returns>
        public static Image ConvertHtmlToImage(string html, int imgWidth = 1024, int imgQuality = 100)
        {
            var bytes = ConvertHtmlToImageByte(html, imgWidth, imgQuality);
            using (MemoryStream ms = new MemoryStream(bytes))
            {
                var image = Image.FromStream(ms);
                return image;
            }
        }
    }
}
