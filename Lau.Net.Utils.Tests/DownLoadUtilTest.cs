using NUnit.Framework;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class DownLoadUtilTest
    {
        private DownloadUtil downloader;

        [SetUp]
        public void SetUp()
        {
            downloader = new DownloadUtil(3);
        }

        [Test]
        public async Task DownloadAsyncTest()
        {
            await downloader.DownloadAsync("https://example.com/file.zip", "C:\\Downloads\\file.zip");
        }

        [Test]
        public async Task DownloadWithProgressTest()
        {
            await downloader.DownloadAsync("https://example.com/file.zip", "C:\\Downloads\\file.zip", progress =>
            {
                Console.WriteLine($"当前下载进度：{progress:F2}%");
            });
        }

        [Test]
        public async Task DownloadBatchAsyncTest()
        {
            var downloadTasks = new Dictionary<string, string>
            {
                ["https://example.com/file1.zip"] = "C:\\Downloads\\file1.zip",
                ["https://example.com/file2.zip"] = "C:\\Downloads\\file2.zip",
                ["https://example.com/file3.zip"] = "C:\\Downloads\\file3.zip"
            };

            await downloader.DownloadBatchAsync(downloadTasks, (url, progress) =>
            {
                Console.WriteLine($"文件 {url} 下载进度：{progress:F2}%");
            });
        }

        [Test]
        public async Task DownloadWithCancelTest()
        {
            var cts = new CancellationTokenSource();
            var cancellationToken = cts.Token;

            // 启动下载任务
            var downloadTask = downloader.DownloadAsync("https://example.com/file.zip", "C:\\Downloads\\file.zip", cancellationToken);

            // 模拟 3 秒后取消
            _ = Task.Delay(3000).ContinueWith(_ => cts.Cancel());

            try
            {
                await downloadTask;
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("下载已被取消！");
            }
        }
    }
}
