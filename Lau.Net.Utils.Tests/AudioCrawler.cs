using HtmlAgilityPack;
using Lau.Net.Utils.Net;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    internal class AudioCrawler
    {
        private static readonly HttpClient client = new HttpClient();
        private static readonly SemaphoreSlim semaphore = new SemaphoreSlim(5); // Limit to 5 concurrent downloads

        public static async Task Start()
        {
            string baseUrl = "https://archive.org/download/ipodcast/";
            string basePath = "downloaded_podcasts";

            await CrawlAndDownloadAsync(baseUrl, basePath);
            Console.WriteLine("Download completed!");
        }

        static async Task CrawlAndDownloadAsync(string url, string basePath)
        {
            var html = await client.GetStringAsync(url);
            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            var table = doc.DocumentNode.SelectSingleNode("//table[@class='directory-listing-table']");
            if (table == null)
            {
                Console.WriteLine($"No directory listing table found at {url}");
                return;
            }

            var rows = table.SelectNodes(".//tr").Skip(1); // Skip header row
            var downloadTasks = new ConcurrentBag<Task>();

            foreach (var row in rows)
            {
                var cells = row.SelectNodes("td");
                if (cells?.Count < 2) continue;

                var linkCell = cells[0];
                var link = linkCell.SelectSingleNode(".//a");
                if (link == null) continue;

                string href = link.GetAttributeValue("href", string.Empty);
                string newPath = Path.Combine(basePath, href);

                if (href.EndsWith("/"))
                {
                    // This is a directory
                    if (Directory.Exists(newPath))
                    {
                        Console.WriteLine($"Directory already exists: {newPath}");
                        continue;
                    }

                    Directory.CreateDirectory(newPath);
                    string newUrl = new Uri(new Uri(url), href).AbsoluteUri;
                    await CrawlAndDownloadAsync(newUrl, newPath);
                }
                else if (href.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase))
                {
                    // This is an MP3 file
                    if (File.Exists(newPath))
                    {
                        Console.WriteLine($"File already exists: {newPath}");
                        continue;
                    }

                    string fileUrl = new Uri(new Uri(url), href).AbsoluteUri;
                    downloadTasks.Add(DownloadFileAsync(fileUrl, newPath));
                }
            }

            await Task.WhenAll(downloadTasks);
        }

        static async Task DownloadFileAsync(string url, string filePath)
        {
            await semaphore.WaitAsync(); // Wait for a slot to be available
            try
            {
                Console.WriteLine($"Downloading: {url}");
                using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
                using (var stream = await response.Content.ReadAsStreamAsync())
                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await stream.CopyToAsync(fileStream);
                }
                Console.WriteLine($"Completed: {url}");
            }
            finally
            {
                semaphore.Release(); // Release the slot
            }
        }
    }
}
