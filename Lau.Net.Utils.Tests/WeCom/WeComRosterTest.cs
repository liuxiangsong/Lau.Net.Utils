using Lau.Net.Utils.Config;
using Lau.Net.Utils.Excel;
using Lau.Net.Utils.WeCom;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace Lau.Net.Utils.Tests.WeCom
{
    [TestFixture]
    public class WeComRosterTest
    {
        WeComRoster _weComRoster = new WeComRoster(AppConfigUtil.GetValue("WeComCorpId"), AppConfigUtil.GetValue("WeComAppSecret"));
        string _testAccount = AppConfigUtil.GetValue("WeComUserId");

        [Test]
        public void GetFieldsTest()
        {
            var res = _weComRoster.GetFields();
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void GetStaffAllInfosTest()
        {
            var res = _weComRoster.GetStaffInfos(_testAccount);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void GetStaffInfosTest()
        {
            var param = new JObject
            {
                ["fieldids"] = new JArray
                {
                    new JObject
                    {
                        ["fieldid"] = 11004,
                        ["sub_idx"] = 0
                    },
                    new JObject
                    {
                        ["fieldid"] = 14001,
                        ["sub_idx"] = 1
                    }
                }
            };
            var res = _weComRoster.GetStaffInfos(_testAccount, param);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }

        [Test]
        public void UpdateStaffInfoTest()
        {
            var param = new JObject
            {
                ["update_items"] = new JArray
                    {
                        new JObject
                        {
                            ["fieldid"] = 12024,
                            ["value_string"] = "test"
                        }
                    }
            };
            var res = _weComRoster.UpdateStaffInfo(_testAccount, param);
            var jobject = JsonConvert.DeserializeObject<JObject>(res);
            Assert.AreEqual(jobject["errcode"].As<int>(), 0);
        }


        [Test]
        public void UpdateStaffInfoTempTest()
        {
            var table = NpoiUtil.ExcelToDataTable(@"D:\excel\1.xlsx");
            var failList = new List<string>();
            foreach (DataRow row in table.Rows)
            {
                var userId = row.GetValue<string>("账号");
                var no = row.GetValue<string>("员工工号").ToUpper();
                var param = new JObject
                {
                    ["update_items"] = new JArray
                    {
                        new JObject
                        {
                            ["fieldid"] = 12024,
                            ["value_string"] = no
                        }
                    }
                };
                var res = _weComRoster.UpdateStaffInfo(userId, param);
                var jobject = JsonConvert.DeserializeObject<JObject>(res);
                if (jobject["errcode"].As<int>() != 0)
                {
                    failList.Add(userId);
                }
            }
        }
    }
}
