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

            var tasks = links.Select(async link =>
            {
                var number = Regex.Match(link.Url, @"course-2-(\d{3})").Groups[1].Value;
                var title = GetEnglishTitle(link.Text);
                var filePath = GetFilePath($"{number}.{title}");
                if (!File.Exists(filePath))
                {
                    var fullUrl = $"http://www.newconceptenglish.com/{link.Url}";
                    await CrawlAndDownloadAsync(fullUrl, $"{number}.{title}");
                }
            });
            await Task.WhenAll(tasks);
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
            try
            {
                var htmlDoc = await HtmlUtil.GetHtmlDocumentAsync(url);
                var nceDiv = htmlDoc.DocumentNode.SelectSingleNode("//div[contains(@class, 'nce')][.//h3[contains(text(), '新概念英语－课文')]]");

                if (nceDiv == null)
                {
                    throw new InvalidOperationException($"未找到课文内容: {url}");
                }

                // 移除所有 h3 标题
                nceDiv.SelectNodes(".//h3")?.ToList().ForEach(node => node.Remove());

                // 清理并格式化内容
                var content = CleanHtmlContent(nceDiv.InnerText);

                var path = GetFilePath(title);
                FileUtil.SaveFile(path, content);
            }
            catch (Exception ex)
            {
                // 可以根据需要添加日志记录
                throw new Exception($"处理页面 {url} 时发生错误", ex);
            }
        }

        /// <summary>
        /// 清理 HTML 内容
        /// </summary>
        /// <param name="content">内容</param>
        /// <returns></returns>
        private static string CleanHtmlContent(string content)
        {
            if (string.IsNullOrEmpty(content))
                return string.Empty;

            // 定义替换规则
            var cleaningRules = new Dictionary<string, string>
            {
                {@"<br\s*/?>\s*", "\n"}, // 替换 <br> 标签为换行符  
                {@"</p>\s*", "\n\n"}, // 替换 </p> 标签为两个换行符
                {@"<[^>]+>", ""}, // 移除所有 HTML 标签
                {@"&nbsp;", " "}, // 替换 &nbsp; 为空格
                {@"[ \t]+", " "}, // 替换多个空格为一个空格
                {@"\n{3,}", "\n\n"} // 替换三个或更多换行符为两个换行符
            };

            // 应用所有清理规则
            return cleaningRules.Aggregate(
                content.Trim(),
                (current, rule) => Regex.Replace(current, rule.Key, rule.Value)
            ).Trim();
        }

    }
}
