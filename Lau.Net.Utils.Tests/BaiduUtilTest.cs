using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Lau.Net.Utils.AI;
using Lau.Net.Utils.Enums;
using NUnit.Framework;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class BaiduUtilTest
    {
        [Test]
        public void OcrTest()
        {
            var clientId = "";
            var clientSecret = "";
            var baiduOcrUtil=new BaiduOcrUtil(clientId,clientSecret);
            var img = Image.FromFile(@"F:\test\1.png");
            var res= baiduOcrUtil.GetTextInImage(@"F:\test\1.png");
            Assert.Greater(res.Value<int>("words_result_num"),0);
        }
    }
}
