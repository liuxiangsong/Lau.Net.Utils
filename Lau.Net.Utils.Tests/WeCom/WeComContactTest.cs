using Lau.Net.Utils.WeCom;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lau.Net.Utils.Excel;
using NPOI.SS.Formula.Functions;
using Lau.Net.Utils.Config;

namespace Lau.Net.Utils.Tests.WeCom
{
    [TestFixture]
    public class WeComContactTest
    {

        WeComContact _weComContact = new WeComContact(AppConfigUtil.GetValue("WeComCorpId"), AppConfigUtil.GetValue("WeComCorpSecret"));
        string _testAccount = AppConfigUtil.GetValue("WeComUserId");

        class Dept
        {
            public int id { get; set; }
            public int parentid { get; set; }
            public string name { get; set; }
            public int order { get; set; }
        }

        #region 部门管理 
        [Test]
        public void GetDeptmentsTest()
        {
            var res = _weComContact.GetDeptments("");
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            var deptList = jobject["department"].ToObject<List<Dept>>();
            var dt = DataTableUtil.ListToDataTable<Dept>(deptList);
            NpoiUtil.DataTableToExcel("D:\\1.xls", dt);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }
        #endregion

        #region 用户管理
        [Test]
        public void GetUserInfoTest()
        {
            var res = _weComContact.GetUserInfo(_testAccount);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void GetDeptmentUsersTest()
        {
            var res = _weComContact.GetDeptmentUsers("1");
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void GetUserIdsTest()
        {
            var res = _weComContact.GetUserIds(1000);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            var userIds = jobject["dept_user"].Select(t => t["userid"].As<string>()).ToList();
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void UpdateUser()
        {
            var jObj = new JObject { { "gender", "1" }, { "alias", "Sam" } };
            var res = _weComContact.UpdateUser("testUserId", false, jObj);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }
        #endregion

        #region 标签管理
        [Test]
        public void GetTagsTest()
        {
            var res = _weComContact.GetTags();
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void GetTagUsers()
        {
            var res = _weComContact.GetTagUsers("2");
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }
        #endregion
    }
}
