using Lau.Net.Utils;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class FtpUtilTest
    {
        private FtpClientUtil _ftpClientUtil;

        public FtpUtilTest()
        {
            _ftpClientUtil = new FtpClientUtil("192.168.0.1", "user01", "123456");
        }

        [Test]
        public void CreateDirectoryTest()
        {
            _ftpClientUtil.CreateDirectory("aaa/bbb/ccc");
        }

        [Test]
        public void UploadFileTest()
        {
            Stream stream = new FileStream(@"E:\test\image1.jpg", FileMode.Open, FileAccess.Read);
            _ftpClientUtil.UploadFile(stream, "0000000/1/3.jpg");
        }

        [Test]
        public void DownloadFileTest()
        {
            var message = _ftpClientUtil.DownloadFile("/0000000/1/3.jpg", @"E:\test\image1223.jpg");
            Assert.IsEmpty(message);
        }
    }
}
