using Lau.Net.Utils.Web;
using Lau.Net.Utils.WeCom;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests.WeCom
{
    [TestFixture]
    public class WeComMessageTest
    {
        WeComMessage _weComMessage = new WeComMessage("CorpId", "SECRET", "AppId");
        string _testAccount = "30000";

        [Test]
        public void SendTextTest()
        {
           var res = _weComMessage.SendText("test", _testAccount);
           var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void SendMarkDownTest()
        {
            var content = "您的会议室已经预定，稍后会同步到`邮箱`  \n>**事项详情**  \n>事　项：<font color=\"info\">开会</font>  \n>组织者：@miglioguan  \n>参与者：@miglioguan、@kunliu、@jamdeezhou、@kanexiong、@kisonwang  \n>  \n>会议室：<font color=\"info\">广州TIT 1楼 301</font>  \n>日　期：<font color=\"warning\">2018年5月18日</font>  \n>时　间：<font color=\"comment\">上午9:00-11:00</font>  \n>  \n>请准时参加会议。  \n>  \n>如需修改会议信息，请点击：[修改会议信息](https://work.weixin.qq.com)";
            content = "您的会议室已经预定，**稍后**会同步到`邮箱`,<font color=\"info\">开会</font>";
            var res = _weComMessage.SendMarkDown(content, _testAccount);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void SendImageTest()
        {
            var dt = HtmlTest.CreateTestTable();
            var imgBytes = HtmlUtil.ConvertTableToImageByte(dt, "测试");
            //var imgBytes = File.ReadAllBytes("E:\\test\\test.png");
            var res = _weComMessage.SendImage(imgBytes, _testAccount);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void SendFileTest()
        {
            var imgBytes = File.ReadAllBytes("E:\\test\\test.png");
            var res = _weComMessage.SendFile(imgBytes,"test.png", _testAccount);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }
    }
}
