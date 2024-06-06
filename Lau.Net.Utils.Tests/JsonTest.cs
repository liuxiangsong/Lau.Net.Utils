using Newtonsoft.Json;
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
    }

    public class MyEntity
    {
        public decimal Prop1 { get; set; }
        public int Prop2 { get; set; }
    }
}
