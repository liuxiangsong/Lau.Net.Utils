using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class JsonTest
    {
        [Test]
        public void StringOrNullToNumberConverterTest()
        {
            string jsonString = "[{\"Prop1\":\"\",\"Prop2\":\"\"}, {\"Prop1\":\"123.45\", \"Prop2\":\"456\"}, {\"Prop1\":null, \"Prop2\":null}]";
            var settings = new JsonSerializerSettings
            {
                Converters = new List<JsonConverter>
                {
                    new StringOrNullToNumberConverter<decimal>(),
                    new StringOrNullToNumberConverter<int>()
                }
            };
            var entities = JsonConvert.DeserializeObject<List<MyEntity>>(jsonString, settings);
            Assert.IsNotNull(entities);
        }

        [Test]
        public void CreateJobjectTest()
        {
            var param = new JObject
            {
                ["userid"] = "123",
                ["update_items"] = new JArray
                    {
                        new JObject
                        {
                            ["fieldid"] = 12024,
                            ["value_string"] = "xxx"
                        },
                        new JObject
                        {
                            ["fieldid"] = 12025,
                            ["value_string"] = "yyy"
                        }
                    }
            };
            var json = JsonConvert.SerializeObject(param);
            Console.WriteLine(json);
        }

        [Test]
        public void MergeJobjectTest()
        {
            var param1 = new JObject
            {
                ["userid"] = "123",
                ["update_items"] = new JArray
                {
                    new JObject
                    {
                        ["fieldid"] = 12024,
                        ["value_string"] = "xxx"
                    }
                }
            };

            var param2 = new JObject
            {
                ["userid"] = "456",
                ["update_items"] = new JArray
                {
                    new JObject
                    {
                        ["fieldid"] = 12024,
                        ["value_string"] = "yyy"
                    }
                }
            };

            param1.Merge(param2);
            var json = JsonConvert.SerializeObject(param1);
            Console.WriteLine(json);

            Assert.AreEqual("456", param1["userid"].ToString());
            Assert.AreEqual(2, (param1["update_items"] as JArray).Count);
        }

        private class MyEntity
        {
            public decimal Prop1 { get; set; }
            public int Prop2 { get; set; }
        }
    }
}
