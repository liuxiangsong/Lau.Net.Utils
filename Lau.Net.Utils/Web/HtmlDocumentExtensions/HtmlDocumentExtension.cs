using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Web.HtmlDocumentExtensions
{
    public static class HtmlDocumentExtension
    {
        /// <summary>
        /// 获取html字符串
        /// </summary>
        /// <param name="htmlDoc"></param>
        /// <returns></returns>
        public static string GetHtml(this HtmlDocument htmlDoc)
        {
            var html = htmlDoc.DocumentNode.OuterHtml;
            return html;
        }
          
        public static HtmlNode GetBodyNode(this HtmlDocument htmlDoc)
        {
            var bodyNode = htmlDoc.DocumentNode.SelectSingleNode("body");
            if(bodyNode != null)
            {
                return bodyNode;
            }
            htmlDoc.OptionAutoCloseOnEnd = true;
            htmlDoc.OptionWriteEmptyNodes = true;
            var htmlNode = htmlDoc.DocumentNode.GetOrCreateHtmlNode("html");
            var headNode = htmlNode.GetOrCreateHtmlNode("head");
            var metaNode = htmlDoc.CreateElement("meta");
            metaNode.Attributes.Add("charset", "utf-8");
            headNode.AppendChild(metaNode);
            bodyNode = htmlNode.GetOrCreateHtmlNode("body");
            return bodyNode;
        }



    }
}
