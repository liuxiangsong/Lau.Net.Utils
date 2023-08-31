using Lau.Net.Utils.Config;
using Lau.Net.Utils.Excel;
using Lau.Net.Utils.WeCom;
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
    public class WeComRobotTest
    {
        WeComRobot _wxRobot = new WeComRobot(AppConfigUtil.GetValue("WeComRobotKey"));

        [Test]
        public void SendTextTest()
        {
            var mentionedList = new List<string>() { AppConfigUtil.GetValue("WeComUserId") };
            _wxRobot.SendText("测试文本内容", mentionedList);
        }

        [Test]
        public void SendMarkDownTest()
        {
            var text = "实时新增用户反馈<font color=\"warning\">132例</font>，请相关同事注意。";
            _wxRobot.SendMarkDown(text);
        }

        [Test]
        public void SendImageTest()
        {
            var imgBytes = File.ReadAllBytes("E:\\test\\test.jpg");
            _wxRobot.SendImage(imgBytes);
        }
        [Test]
        public void SendTableImageTest()
        {
            var dt = DataTableUtil.CreateTable("列1", "列2");
            _wxRobot.SendImage(dt, "标题内容");
        }

        [Test]
        public void SendFileTest()
        {
            var fileBytes = File.ReadAllBytes("E:\\test\\test.xls");
            _wxRobot.SendFile(fileBytes, "test.xls");
        }
        [Test]
        public void SendFile2Test()
        {
            var dt = DataTableUtil.CreateTable("列1", "列2");
            var ms = NpoiStaticUtil.DataTableToStream(dt);
            _wxRobot.SendFile(ms.ToArray(), "test.xls");
        }
    }
}
