using HashidsNet;
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
    public class OtherTest
    {
        [Test]
        public void ZipTest()
        {
            ZipUtil.Compress(@"E:\website\发布文件夹", @"E:\website\发布文件夹.zip");
            ZipUtil.Decompress(@"E:\website\发布文件夹.zip", @"E:\website\发布文件夹",true);
            //ZipUtil.AddFileToZip(@"E:\website\发布文件夹.zip", @"E:\website\msbuild.pubxml");
        }

        [Test]
        public void HashIdTest()
        {
            //生成时间戳
            var timeStamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds();
            var hashIds = new Hashids("this is my salt");
            var shortId = hashIds.EncodeLong(timeStamp);
            var timeStamp2 = hashIds.DecodeSingleLong(shortId);
            Assert.AreEqual(timeStamp, timeStamp2);
        }

        [Test]
        public void SavFileTest()
        {
            File.Delete(@"E:\test\1asdf.txt");
            FileUtil.SaveFile(@"E:\test\1.txt", "asdf");
        }
    }
}
