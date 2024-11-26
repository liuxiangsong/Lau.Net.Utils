using Lau.Net.Utils.Web;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Lau.Net.Utils.Web.HtmlDocumentExtensions;
using NPOI.SS.Formula.Functions;
using System.IO;

namespace Lau.Net.Utils.Tests.Crawlers
{
    internal class NceCrawler
    {
        static string _saveDirectory;

        /// <summary>
        /// 爬取并下载新概念英语2
        /// </summary>
        /// <param name="saveDirectory">文件保存目录</param>
        /// <returns></returns>
        public static async Task Start(string saveDirectory)
        {
            _saveDirectory = saveDirectory;
            // 获取目录页
            var indexUrl = "http://www.newconceptenglish.com/index.php?id=nce-2";
            var indexDoc = await HtmlUtil.GetHtmlDocumentAsync(indexUrl);

            // 获取所有课程链接
            var links = indexDoc.DocumentNode.SelectNodes("//table//a")
                .Select(a => new
                {
                    Url = a.GetAttributeValue("href", ""),
                    Text = a.InnerText.Trim()
                })
                .Where(l => l.Url.Contains("course-2-"));

            foreach (var link in links)
            {
                var number = Regex.Match(link.Url, @"course-2-(\d{3})").Groups[1].Value;
                var title = GetEnglishTitle(link.Text);
                var filePath = GetFilePath($"{number}.{title}");
                if (File.Exists(filePath))
                {
                    continue;
                }
                var fullUrl = $"http://www.newconceptenglish.com/{link.Url}";
                await CrawlAndDownloadAsync(fullUrl, $"{number}.{title}");
            }
        }

        /// <summary>
        /// 获取英语标题
        /// </summary>
        /// <param name="fullTitle">完整标题</param>
        /// <returns></returns>
        private static string GetEnglishTitle(string fullTitle)
        {
            fullTitle = fullTitle.Trim().Substring(5).Trim();
            //查找第一个中文字符的位置
            var englishTitle = fullTitle;
            for (int i = 0; i < fullTitle.Length; i++)
            {
                if (char.GetUnicodeCategory(fullTitle[i]) == System.Globalization.UnicodeCategory.OtherLetter)
                {
                    englishTitle = fullTitle.Substring(0, i).Trim();
                    break;
                }
            }
            return englishTitle;
        }

        /// <summary>
        /// 获取文件路径
        /// </summary>
        /// <param name="title">标题</param>
        /// <returns></returns>
        private static string GetFilePath(string title)
        {
            var invalidChars = Path.GetInvalidFileNameChars()
            .Concat(Path.GetInvalidPathChars()).Distinct();
            // 移除名称中的非法字符
            title = new string(title.Where(c => !invalidChars.Contains(c)).ToArray());
            return Path.Combine(_saveDirectory, $@"{title}.md");
        }

        /// <summary>
        /// 爬取并下载新概念英语2
        /// </summary>
        /// <param name="url">链接</param>
        /// <param name="title">标题</param>
        /// <returns></returns>
        private static async Task CrawlAndDownloadAsync(string url, string title)
        {
            var htmlDoc = await HtmlUtil.GetHtmlDocumentAsync(url);

            // 获取标题元素
            var titleElement = htmlDoc.GetElementbyId("coursetitle");
            var fullTitle = titleElement?.InnerText.Trim() ?? string.Empty;

            var html = htmlDoc.GetHtml();
            var contentElement = htmlDoc.DocumentNode.SelectSingleNode("//h3[contains(text(), '新概念英语－课文')]/following-sibling::p");
            var content = contentElement?.InnerText.Trim() ?? string.Empty;

            // 将<br>和<p>标签转换为换行符
            content = Regex.Replace(content, @"<br\s*/?>\s*", "\n");
            content = Regex.Replace(content, @"</p>\s*", "\n\n");

            // 清理其他HTML标签
            content = Regex.Replace(content, @"<[^>]+>", "");
            content = Regex.Replace(content, @"&nbsp;", " ");

            // 清理多余空白，但保留换行
            content = Regex.Replace(content, @"[ \t]+", " ");
            content = Regex.Replace(content, @"\n{3,}", "\n\n");
            content = content.Trim();

            var path = GetFilePath(title);
            FileUtil.SaveFile(path, content);
        }

    }
}
