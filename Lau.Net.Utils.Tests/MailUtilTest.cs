using Lau.Net.Utils.Config;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class MailUtilTest
    {
        [Test]
        public void SendTest()
        {
            var host = AppConfigUtil.GetValue("MailHost");
            var user = AppConfigUtil.GetValue("MailUser");
            var password = AppConfigUtil.GetValue("MailPassword");
            var displayName = AppConfigUtil.GetValue("MailDisplayName");
            var reciver = AppConfigUtil.GetValue("MailReceiver");
            var mailUtil = new MailUtil(host, user, password, displayName);
            var dt = DataTableUtilTest.CreateTestTable();
            var attachments = mailUtil.DataTableToExcelAttachment(dt, "测试");
            mailUtil.Send(new List<string> { reciver }, "测试", "测试",null, attachments);
        }
    }
}
