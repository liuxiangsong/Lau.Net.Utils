using HtmlAgilityPack;
using Lau.Net.Utils.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests.Crawlers
{
    internal static class NceAudioCrawler
    {
        private const string BaseUrl = "https://www.listeningexpress.com/nce-a/book2/";

        public static async Task Start(string downloadPath)
        {
            try
            {
                // 获取所有课程URL
                List<string> allLessonUrls = new List<string>();
                for (int page = 1; page <= 5; page++)
                {
                    var catalogUrl = page == 1 ? BaseUrl : $"{BaseUrl}index-{page}.html";
                    Console.WriteLine($"正在处理第{page}页: {catalogUrl}");

                    var lessonUrls = await ParseLessonUrlsAsync(catalogUrl);
                    allLessonUrls.AddRange(lessonUrls);
                }

                // 获取所有音频下载信息
                var downloadTasks = new List<(string Url, string FileName)>();
                var tasks = allLessonUrls.Select(async lessonUrl =>
                {
                    var audioInfo = await GetAudioDownloadInfoAsync(lessonUrl);
                    if (audioInfo.HasValue)
                    {
                        lock (downloadTasks)
                        {
                            downloadTasks.Add(audioInfo.Value);
                        }
                    }
                });
                await Task.WhenAll(tasks);

                // 批量下载
                var downloads = downloadTasks.ToDictionary(a => a.Url, a => Path.Combine(downloadPath, a.FileName));

                using (var downloadUtil = new DownloadUtil())
                {
                    await downloadUtil.DownloadBatchAsync(downloads);
                    //await downloadUtil.DownloadBatchAsync(downloads,
                    //    onStart: task => Console.WriteLine($"正在下载: {Path.GetFileName(task.FilePath)}"),
                    //    onComplete: task => Console.WriteLine($"下载完成: {Path.GetFileName(task.FilePath)}"),
                    //    onError: (task, ex) => Console.WriteLine($"下载文件 {Path.GetFileName(task.FilePath)} 时出错: {ex.Message}"),
                    //    onSkip: task => Console.WriteLine($"文件已存在，跳过下载: {Path.GetFileName(task.FilePath)}")
                    //);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"抓取过程出错: {ex.Message}");
            }
        }

        private static async Task<(string Url, string FileName)?> GetAudioDownloadInfoAsync(string lessonUrl)
        {
            try
            {
                Console.WriteLine($"正在解析课程: {lessonUrl}");
                var audioUrl = await ParseAudioUrlAsync(lessonUrl);
                if (!string.IsNullOrEmpty(audioUrl))
                {
                    var lessonName = ParseLessonName(lessonUrl);
                    return (audioUrl, $"{lessonName}.mp3");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"解析课程 {lessonUrl} 时出错: {ex.Message}");
            }
            return null;
        }

        private static async Task<List<string>> ParseLessonUrlsAsync(string url)
        {
            var urls = new List<string>();
            var doc = await HtmlUtil.GetHtmlDocumentAsync(url);

            // 查找所有课程链接
            var lessonNodes = doc.DocumentNode.SelectNodes("//div[@id='list']//a[@href]");
            if (lessonNodes != null)
            {
                foreach (var node in lessonNodes)
                {
                    var href = node.GetAttributeValue("href", "");
                    if (!string.IsNullOrEmpty(href) && href.Contains("book2") && href.EndsWith(".html"))
                    {
                        // 确保是完整的URL
                        if (!href.StartsWith("http"))
                        {
                            href = new Uri(new Uri(BaseUrl), href).ToString();
                        }
                        urls.Add(href);
                    }
                }
            }
            return urls;
        }

        private static async Task<string> ParseAudioUrlAsync(string lessonUrl)
        {
            var doc = await HtmlUtil.GetHtmlDocumentAsync(lessonUrl);

            var downloadNode = doc.DocumentNode.SelectSingleNode("//a[@id='btndl-mp3']");
            var href = downloadNode?.GetAttributeValue("href", "");

            // 确保返回完整的URL
            if (!string.IsNullOrEmpty(href))
            {
                if (!href.StartsWith("http"))
                {
                    href = new Uri(new Uri(lessonUrl), href).ToString();
                }
                return href;
            }
            return "";
        }

        private static string ParseLessonName(string lessonUrl)
        {
            // 从URL中提取课程名称
            var match = Regex.Match(lessonUrl, @"/([^/]+)\.html$");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return DateTime.Now.Ticks.ToString(); // 如果无法提取名称，使用时间戳
        }
    }
}
