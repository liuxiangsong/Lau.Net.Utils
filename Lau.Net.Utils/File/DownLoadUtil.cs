using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

/// <summary>
/// 下载工具类
/// </summary>
public class DownloadUtil : IDisposable
{
  private readonly HttpClient _httpClient;
  private readonly SemaphoreSlim _semaphore;

  /// <summary>
  /// 初始化下载工具
  /// </summary>
  /// <param name="maxConcurrentDownloads">并发下载数量</param> 
  public DownloadUtil(int maxConcurrentDownloads = 5)
  {
    if (maxConcurrentDownloads <= 0)
      throw new ArgumentException("并发数量必须大于 0");

    _httpClient = new HttpClient();
    _semaphore = new SemaphoreSlim(maxConcurrentDownloads);
  }

  /// <summary>
  /// 下载文件（无进度显示）
  /// </summary>
  /// <param name="url">链接</param>
  /// <param name="destinationPath">文件路径</param>
  /// <param name="cancellationToken">取消令牌</param>
  /// <returns></returns>
  public async Task DownloadAsync(string url, string destinationPath, CancellationToken cancellationToken = default(CancellationToken))
  {
    await _semaphore.WaitAsync(cancellationToken); // 控制并发下载数量

    try
    {
      using (var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken))
      {
        response.EnsureSuccessStatusCode();

        using (var stream = await response.Content.ReadAsStreamAsync())
        using (var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
        {
          await stream.CopyToAsync(fileStream, 8192, cancellationToken);
        }
      }
    }
    finally
    {
      _semaphore.Release(); // 释放信号量
    }
  }

  /// <summary>
  /// 下载文件（带进度显示）
  /// </summary>
  /// <param name="url">链接</param>
  /// <param name="destinationPath">文件路径</param>
  /// <param name="onProgress">进度回调</param>
  /// <param name="cancellationToken">取消令牌</param>
  /// <returns></returns>
  public async Task DownloadAsync(string url, string destinationPath, Action<double> onProgress, CancellationToken cancellationToken = default(CancellationToken))
  {
    await _semaphore.WaitAsync(cancellationToken);

    try
    {
      using (var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, cancellationToken))
      {
        response.EnsureSuccessStatusCode();

        var totalBytes = response.Content.Headers.ContentLength ?? -1L;
        var receivedBytes = 0L;

        using (var stream = await response.Content.ReadAsStreamAsync())
        using (var fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write, FileShare.None, 8192, true))
        {
          var buffer = new byte[8192];
          int bytesRead;
          while ((bytesRead = await stream.ReadAsync(buffer, 0, buffer.Length, cancellationToken)) > 0)
          {
            await fileStream.WriteAsync(buffer, 0, bytesRead, cancellationToken);
            receivedBytes += bytesRead;

            if (totalBytes > 0)
            {
              var progress = (double)receivedBytes / totalBytes * 100;
              onProgress?.Invoke(progress); // 调用进度回调
            }
          }
        }
      }
    }
    finally
    {
      _semaphore.Release();
    }
  }

  /// <summary>
  /// 批量下载
  /// </summary>
  /// <param name="downloadTasks">下载任务列表</param>
  /// <param name="onProgress">进度回调</param>
  /// <param name="cancellationToken">取消令牌</param>
  /// <returns></returns>
  public async Task DownloadBatchAsync(
      IEnumerable<KeyValuePair<string, string>> downloadTasks,
      Action<string, double> onProgress = null,
      CancellationToken cancellationToken = default)
  {
    var tasks = new List<Task>();

    foreach (var task in downloadTasks)
    {
      var url = task.Key;
      var destination = task.Value;

      tasks.Add(Task.Run(() =>
          DownloadAsync(url, destination, progress =>
              onProgress?.Invoke(url, progress),
              cancellationToken)));
    }

    await Task.WhenAll(tasks);
  }


  #region IDisposable
  private bool _disposed = false;

  public void Dispose()
  {
    Dispose(true);
    GC.SuppressFinalize(this);
  }

  protected virtual void Dispose(bool disposing)
  {
    if (!_disposed)
    {
      if (disposing)
      {
        _httpClient?.Dispose();
        _semaphore?.Dispose();
      }

      _disposed = true;
    }
  }

  ~DownloadUtil()
  {
    Dispose(false);
  }
  #endregion

}
