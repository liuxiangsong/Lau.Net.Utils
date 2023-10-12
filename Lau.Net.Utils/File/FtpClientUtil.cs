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
        public FtpClientUtil(string ftpHost,string userId,string password)
        {
            _ftpHost = ftpHost;
            _userId = userId;
            _password = password;
        }

        /// <summary>
        /// 创建目录
        /// </summary>
        /// <param name="dirPath">目录相对路径</param>
        public void CreateDirectory(string dirPath)
        {
            using(FtpClient ftpClient = new FtpClient())
            {
                ftpClient.Host = _ftpHost;
                ftpClient.Credentials = new NetworkCredential(_userId, _password);
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
        public void UploadFile(Stream fileStream,string filePath)
        {
            using (FtpClient ftpClient = new FtpClient())
            {
                ftpClient.Host = _ftpHost;
                ftpClient.Credentials = new NetworkCredential(_userId, _password);
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
    }
}
