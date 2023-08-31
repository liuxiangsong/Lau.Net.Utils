using Lau.Net.Utils.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class RestSharpUtilTest
    {
        string baseUrl = "https://jsonplaceholder.typicode.com";

        [Test]
        public void GetTest()
        {
            var url = baseUrl + "/posts/1";
            var result = RestSharpUtil.Get<string>(url);
            var jObj = JsonConvert.DeserializeObject<JObject>(result);
            Assert.IsNotNull(result);
            Assert.AreEqual(jObj["id"].ToString(), "1");
        }

        [Test] 
        public async Task GetAsyncTest()
        {
            var url = baseUrl + "/posts";
            var result = await RestSharpUtil.GetAsync<List<PostEntity>>(url);
            Assert.IsInstanceOf<List<PostEntity>>(result);
        }

        [Test]
        public void PostTest()
        {
            var url = baseUrl + "/posts";
            var requestBody = new
            {
                title = "foo",
                body = "bar",
                userId = 1
            };
            var result = RestSharpUtil.Post<PostEntity>(url, requestBody);
            Assert.IsInstanceOf<PostEntity>(result);
            Assert.AreEqual(requestBody.title, result.title);
        }

        [Test]
        public void PostFileTest()
        {
            var url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key=KEY&type=file";
            var fileBytes = File.ReadAllBytes("D:\\test\\test.txt");
            var res = RestSharpUtil.PostFile<string>(url, fileBytes, "test.txt");
            Assert.IsNotNull(res);
        }
    }

    class PostEntity
    {
        public int Id { get; set; }
        public int userId { get; set; }
        public string title { get; set; }
        public string body { get; set; }
    }
}
