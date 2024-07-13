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

        /// <summary>
        /// 获取或创建head节点
        /// </summary>
        /// <param name="htmlDoc"></param>
        /// <returns></returns>
        public static HtmlNode GetOrCreateHeadNode(this HtmlDocument htmlDoc)
        {
            var htmlNode = htmlDoc.DocumentNode.GetOrCreateNodeByNodeName("html");
            var headNode = htmlNode.GetOrCreateNodeByNodeName("head");
            return headNode; 
        }

        /// <summary>
        /// 获取或创建body节点
        /// </summary>
        /// <param name="htmlDoc"></param>
        /// <returns></returns>
        public static HtmlNode GetOrCreateBodyNode(this HtmlDocument htmlDoc)
        {
            var bodyNode = htmlDoc.DocumentNode.SelectSingleNode("body");
            if(bodyNode != null)
            {
                return bodyNode;
            }
            htmlDoc.OptionAutoCloseOnEnd = true;
            htmlDoc.OptionWriteEmptyNodes = true;
            var htmlNode = htmlDoc.DocumentNode.GetOrCreateNodeByNodeName("html");
            var headNode = htmlNode.GetOrCreateNodeByXpath("head");
            var metaNode = htmlDoc.CreateElement("meta");
            metaNode.Attributes.Add("charset", "utf-8");
            headNode.AppendChild(metaNode);
            bodyNode = htmlNode.GetOrCreateNodeByNodeName("body");
            return bodyNode;
        }

        /// <summary>
        /// 添加样式节点
        /// </summary>
        /// <param name="htmlDoc"></param>
        /// <param name="styleContent">样式内容</param>
        /// <returns>返回head节点</returns>
        public static HtmlNode AddStyleNode(this HtmlDocument htmlDoc,string styleContent)
        {
            var headNode = htmlDoc.GetOrCreateHeadNode();
            string styleNodeHtml = $"<style>{styleContent}</style>";
            var styleNode = HtmlNode.CreateNode(styleNodeHtml);
            headNode.AppendChild(styleNode);
            return headNode;
        }

    }
}
