using Lau.Net.Utils.WeCom;
using System.Linq;
using NUnit.Framework;
using Lau.Net.Utils.Config;

namespace Lau.Net.Utils.Tests.WeCom
{
    [TestFixture]
    public class WeComExternalContactTests
    {
        private WeComToken _token;
        private WeComExternalContact _externalContact;

        [SetUp]
        public void Setup()
        {
            _token = new WeComToken(AppConfigUtil.GetValue("WeComCorpId"), AppConfigUtil.GetValue("WeComAppSecret"));
            _externalContact = new WeComExternalContact(_token);
        }

        [Test]
        public void TestGetFollowUsers()
        {
            // 获取配置了客户联系功能的成员列表
            var followUsers = _externalContact.GetFollowUsers();
            
            Assert.IsNotNull(followUsers);
            Assert.IsTrue(followUsers.Count > 0);
        }

        [Test]
        public void TestGetExternalContactList()
        {
            // 先获取成员列表
            var followUsers = _externalContact.GetFollowUsers();
            Assert.IsTrue(followUsers.Count > 0);

            // 获取第一个成员的客户列表
            var userId = followUsers.First();
            var externalContacts = _externalContact.GetExternalContactList(userId);

            Assert.IsNotNull(externalContacts);
        }

        [Test]
        public void TestGetExternalContactDetail()
        {
            // 获取成员列表
            var followUsers = _externalContact.GetFollowUsers();
            Assert.IsTrue(followUsers.Count > 0);

            // 获取第一个成员的客户列表
            var userId = followUsers.First();
            var externalContacts = _externalContact.GetExternalContactList(userId);
            Assert.IsTrue(externalContacts.Count > 0);

            // 获取第一个客户的详细信息
            var externalUserId = externalContacts.Last();
            var detail = _externalContact.GetExternalContactDetail(externalUserId);

            Assert.IsNotNull(detail);
            Assert.IsNotNull(detail.external_contact);
            Assert.IsFalse(string.IsNullOrEmpty(detail.external_contact.name));
            Assert.IsFalse(string.IsNullOrEmpty(detail.external_contact.external_userid));
        }
    }
}
