using HtmlAgilityPack;
using Lau.Net.Utils.Net;
using Lau.Net.Utils.Web;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    internal class AudioCrawler
    {
        private static string _saveDirectory;

        /// <summary>
        /// 爬取并下载音频
        /// </summary>
        /// <param name="saveDirectory">文件保存目录</param>    
        public static async Task Start(string saveDirectory)
        {
            _saveDirectory = saveDirectory;
            string baseUrl = "https://archive.org/download/ipodcast/";

            await CrawlAndDownloadAsync(baseUrl);
            Console.WriteLine("Download completed!");
        }

        static async Task CrawlAndDownloadAsync(string url)
        {
            var downloadDict = await GetDownloadItemsAsync(url);
            using (var downloadUtil = new DownloadUtil(3))
            {
                await downloadUtil.DownloadBatchAsync(downloadDict);
            }
        }

        /// <summary>
        /// 获取所有需要下载的文件信息
        /// </summary>
        /// <param name="url">起始URL</param>
        /// <returns>下载字典，Key为文件URL，Value为保存路径</returns>
        static async Task<Dictionary<string, string>> GetDownloadItemsAsync(string url)
        {
            var downloadDict = new Dictionary<string, string>();
            var htmlDoc = await HtmlUtil.GetHtmlDocumentAsync(url);

            var table = htmlDoc.DocumentNode.SelectSingleNode("//table[@class='directory-listing-table']");
            if (table == null)
            {
                Console.WriteLine($"No directory listing table found at {url}");
                return downloadDict;
            }

            var rows = table.SelectNodes(".//tr").Skip(1); // Skip header row

            foreach (var row in rows)
            {
                var cells = row.SelectNodes("td");
                if (cells?.Count < 2) continue;

                var linkCell = cells[0];
                var link = linkCell.SelectSingleNode(".//a");
                if (link == null) continue;

                string href = link.GetAttributeValue("href", string.Empty);
                string newPath = Path.Combine(_saveDirectory, href);

                if (href.EndsWith("/"))
                {
                    // 处理目录
                    Directory.CreateDirectory(newPath);
                    string newUrl = new Uri(new Uri(url), href).AbsoluteUri;
                    var subItems = await GetDownloadItemsAsync(newUrl);
                    foreach (var item in subItems)
                    {
                        downloadDict.Add(item.Key, item.Value);
                    }
                }
                else if (href.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase))
                {
                    // 处理MP3文件
                    string fileUrl = new Uri(new Uri(url), href).AbsoluteUri;
                    downloadDict.Add(fileUrl, newPath);
                }
            }

            return downloadDict;
        }
    }
}
