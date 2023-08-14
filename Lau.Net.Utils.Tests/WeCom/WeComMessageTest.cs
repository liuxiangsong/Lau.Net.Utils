using Lau.Net.Utils.WeCom;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests.WeCom
{
    [TestFixture]
    public class WeComMessageTest
    {
        WeComMessage _weComMessage = new WeComMessage("CorpId", "SECRET", "AppId");

        [Test]
        public void SendTextTest()
        {
           var res = _weComMessage.SendText("test", new string[] { "38659" });
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.NotNull(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }
    }
}
