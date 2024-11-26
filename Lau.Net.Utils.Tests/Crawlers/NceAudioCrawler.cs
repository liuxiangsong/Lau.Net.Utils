using HtmlAgilityPack;
using Lau.Net.Utils.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests.Crawlers
{
    internal class NceAudioCrawler
    {
        private readonly HttpClient _httpClient;
        private const string BaseUrl = "https://www.listeningexpress.com/nce-a/book2/";
        private readonly string _downloadPath;

        public NceAudioCrawler(string downloadPath)
        {
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36");
            _downloadPath = downloadPath;
            
            if (!Directory.Exists(_downloadPath))
            {
                Directory.CreateDirectory(_downloadPath);
            }
        }

        public async Task CrawlAllLessonsAsync()
        {
            try
            {
                List<string> allLessonUrls = new List<string>();
                for (int page = 1; page <= 5; page++)
                {
                    var catalogUrl = page == 1 ? BaseUrl : $"{BaseUrl}index-{page}.html";
                    Console.WriteLine($"正在处理第{page}页: {catalogUrl}");
                     
                    var lessonUrls = await ParseLessonUrlsAsync(catalogUrl);
                    allLessonUrls.AddRange(lessonUrls);
                }
                
                foreach (var lessonUrl in allLessonUrls)
                {
                    await CrawlLessonAudioAsync(lessonUrl);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"抓取过程出错: {ex.Message}");
            }
        }

        private async Task CrawlLessonAudioAsync(string lessonUrl)
        {
            try
            {
                Console.WriteLine($"正在处理课程: {lessonUrl}"); 
                
                var audioUrl = await ParseAudioUrlAsync(lessonUrl);
                if (!string.IsNullOrEmpty(audioUrl))
                {
                    var lessonName = ParseLessonName(lessonUrl);
                    await DownloadAudioFileAsync(audioUrl, lessonName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"处理课程 {lessonUrl} 时出错: {ex.Message}");
            }
        }

        private async Task<List<string>> ParseLessonUrlsAsync(string url)
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

        private async Task<string> ParseAudioUrlAsync(string lessonUrl)
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

        private string ParseLessonName(string lessonUrl)
        {
            // 从URL中提取课程名称
            var match = Regex.Match(lessonUrl, @"/([^/]+)\.html$");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }
            return DateTime.Now.Ticks.ToString(); // 如果无法提取名称，使用时间戳
        }

        private async Task DownloadAudioFileAsync(string audioUrl, string lessonName)
        {
            try
            {
                var fileName = $"{lessonName}.mp3";
                var filePath = Path.Combine(_downloadPath, fileName);

                // 检查文件是否已存在
                if (File.Exists(filePath))
                {
                    Console.WriteLine($"文件已存在，跳过下载: {fileName}");
                    return;
                }

                Console.WriteLine($"正在下载: {fileName}");
                var audioBytes = await _httpClient.GetByteArrayAsync(audioUrl);
                //await File.WriteAllBytesAsync(filePath, audioBytes);
                Console.WriteLine($"下载完成: {fileName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"下载音频文件时出错: {ex.Message}");
            }
        }
    }
}
