using Lau.Net.Utils.Config;
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
        static string _WeComCorpId = AppConfigUtil.GetValue("WeComCorpId");
        WeComMessage _weComMessage = new WeComMessage(_WeComCorpId, AppConfigUtil.GetValue("WeComAppId"), AppConfigUtil.GetValue("WeComAppSecret"));
        string[] _testAccounts = new string[] { AppConfigUtil.GetValue("WeComUserId") } ;

        [Test]
        public void SendTextTest()
        {
           var res = _weComMessage.SendText("test", _testAccounts);
           var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void SendMarkDownTest()
        {
            var content = "您的会议室已经预定，稍后会同步到`邮箱`  \n>**事项详情**  \n>事　项：<font color=\"info\">开会</font>  \n>组织者：@miglioguan  \n>参与者：@miglioguan、@kunliu、@jamdeezhou、@kanexiong、@kisonwang  \n>  \n>会议室：<font color=\"info\">广州TIT 1楼 301</font>  \n>日　期：<font color=\"warning\">2018年5月18日</font>  \n>时　间：<font color=\"comment\">上午9:00-11:00</font>  \n>  \n>请准时参加会议。  \n>  \n>如需修改会议信息，请点击：[修改会议信息](https://work.weixin.qq.com)";
            content = "您的会议室已经预定，**稍后**会同步到`邮箱`,<font color=\"info\">开会</font>";
            var res = _weComMessage.SendMarkDown(content, _testAccounts);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }


        [Test]
        public void SendTextCardTest()
        {
            var desc = @"<div class='gray'>2023年9月26日</div> <div class='normal'>主题：xxxx项目紧急通知</div><div class='highlight'>请收到此信息的同事重视此通知，请认真查阅</div>";
            var state = "noticeNo:UN230822004";
            var url = $"https://open.weixin.qq.com/connect/oauth2/authorize?appid={_WeComCorpId}&redirect_uri={AppConfigUtil.GetValue("MppesWebsiteUrl")}&response_type=code&scope=snsapi_base&state={state}#wechat_redirect";
            //var url = $"https://open.weixin.qq.com/connect/oauth2/authorize?appid={_WeComCorpId}&redirect_uri=http%3A%2F%2Fls.mppes.vip%3A5000&response_type=code&scope=snsapi_base&state=#wechat_redirect";
            var res = _weComMessage.SendTextCard("紧急项目通知", desc, url, _testAccounts,btnText:"点击查看");
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void SendImageTest()
        {
            var dt = HtmlTest.CreateTestTable();
            var imgBytes = HtmlUtil.ConvertTableToImageByte(dt, "测试");
            //var imgBytes = File.ReadAllBytes("E:\\test\\test.png");
            var res = _weComMessage.SendImage(imgBytes, _testAccounts);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void SendFileTest()
        {
            var imgBytes = File.ReadAllBytes("E:\\test\\test.png");
            var res = _weComMessage.SendFile(imgBytes,"test.png", _testAccounts);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void SendMessageTest()
        {
            var json = "{\"card_type\":\"text_notice\",\"source\":{\"icon_url\":\"图片的url\",\"desc\":\"企业微信\",\"desc_color\":1},\"action_menu\":{\"desc\":\"卡片副交互辅助文本说明\",\"action_list\":[{\"text\":\"接受推送\",\"key\":\"A\"},{\"text\":\"不再推送\",\"key\":\"B\"}]},\"task_id\":\"task_id\",\"main_title\":{\"title\":\"欢迎使用企业微信\",\"desc\":\"您的好友正在邀请您加入企业微信\"},\"quote_area\":{\"type\":1,\"url\":\"https://work.weixin.qq.com\",\"title\":\"企业微信的引用样式\",\"quote_text\":\"企业微信真好用呀真好用\"},\"emphasis_content\":{\"title\":\"100\",\"desc\":\"核心数据\"},\"sub_title_text\":\"下载企业微信还能抢红包！\",\"horizontal_content_list\":[{\"keyname\":\"邀请人\",\"value\":\"张三\"},{\"type\":1,\"keyname\":\"企业微信官网\",\"value\":\"点击访问\",\"url\":\"https://work.weixin.qq.com\"},{\"type\":2,\"keyname\":\"企业微信下载\",\"value\":\"企业微信.apk\",\"media_id\":\"文件的media_id\"},{\"type\":3,\"keyname\":\"员工信息\",\"value\":\"点击查看\",\"userid\":\"zhangsan\"}],\"jump_list\":[{\"type\":1,\"title\":\"企业微信官网\",\"url\":\"https://work.weixin.qq.com\"},{\"type\":2,\"title\":\"跳转小程序\",\"appid\":\"小程序的appid\",\"pagepath\":\"/index.html\"}],\"card_action\":{\"type\":2,\"url\":\"https://work.weixin.qq.com\",\"appid\":\"小程序的appid\",\"pagepath\":\"/index.html\"}}";
            var jContent = JObject.Parse(json);
            var res = _weComMessage.SendMessage(jContent,_testAccounts,null, "template_card");
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

    }
}
