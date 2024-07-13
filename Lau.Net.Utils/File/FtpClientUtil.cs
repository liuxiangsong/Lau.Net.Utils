using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.FtpClient;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils
{
    public class FtpClientUtil
    {
        #region 私有属性
        string _ftpHost;
        string _userId;
        string _password;
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ftpHost">ftp服务器地址</param>
        /// <param name="userId">ftp用户名</param>
        /// <param name="password">ftp用户密码</param>
        public FtpClientUtil(string ftpHost, string userId, string password)
        {
            _ftpHost = ftpHost;
            _userId = userId;
            _password = password;
        }

        private FtpClient CreateFtpClient()
        {
            var ftpClient = new FtpClient();
            ftpClient.Host = _ftpHost;
            ftpClient.Credentials = new NetworkCredential(_userId, _password);
            return ftpClient;
        }

        /// <summary>
        /// 创建目录
        /// </summary>
        /// <param name="dirPath">目录相对路径</param>
        public void CreateDirectory(string dirPath)
        {
            using (FtpClient ftpClient = CreateFtpClient())
            {
                var isExists = ftpClient.DirectoryExists(dirPath);
                if (!isExists)
                {
                    ftpClient.CreateDirectory(dirPath);
                }
            }
        }

        /// <summary>
        /// 上传文件
        /// </summary>
        /// <param name="fileStream"></param>
        /// <param name="filePath"></param>
        public void UploadFile(Stream fileStream, string filePath)
        {
            using (FtpClient ftpClient = CreateFtpClient())
            {
                var dirPath = Path.GetDirectoryName(filePath);
                CreateDirectory(dirPath);
                using (var ftpStream = ftpClient.OpenWrite(filePath))
                {
                    var buffer = new byte[8 * 1024];
                    int count;
                    while ((count = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ftpStream.Write(buffer, 0, count);
                    }
                }
            }
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="ftpFilePath">ftp服务器文件相对路径</param>
        /// <param name="saveFilePath">保存文件路径</param>
        /// <returns>下载成功返回空字符串，下载失败返回对应原因</returns>
        public string DownloadFile(string ftpFilePath, string saveFilePath)
        {
            using (FtpClient ftpClient = CreateFtpClient())
            {
                if (!ftpClient.FileExists(ftpFilePath))
                {
                    return $"ftp服务器上不存在{ftpFilePath}";
                }
                FileUtil.CreateDirectory(saveFilePath);
                using (var inputStream = ftpClient.OpenRead(ftpFilePath, FtpDataType.Binary))
                {
                    using (var fileStream = File.Create(saveFilePath))
                    {
                        //inputStream.Seek(0, SeekOrigin.Begin);
                        inputStream.CopyTo(fileStream);
                    }
                }
            }
            return string.Empty;
        } 
    }
}
