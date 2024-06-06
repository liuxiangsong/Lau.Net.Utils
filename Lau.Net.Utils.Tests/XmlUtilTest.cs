using Newtonsoft.Json.Linq;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    internal class XmlUtilTest
    {

        #region 将幕布导出opml文件解析成json
        [Test]
       public void ConvertOpmlFileToJsonTest()
        {
            var filePath = @"D:\MP\Downloads\财务工作拆分 (1).opml";
            XDocument doc = XDocument.Load(filePath);
            XElement rootElement = doc.Root;

            if (rootElement != null && rootElement.Name.LocalName == "opml")
            {
                var bodyNode = rootElement.Elements().FirstOrDefault(node => node.Name.LocalName == "body");
                if (bodyNode == null)
                {
                    return;
                }
                JObject jObj = new JObject();
                ParseOutlineElements(bodyNode, jObj);
                AddIdForJObject(jObj,"root");
                string jsonText = jObj.ToString();
                //xunhuanJObj(jObj);
                FileUtil.SaveFile(@"D:\MP\Downloads\财务.json", jsonText);
                Console.WriteLine(jsonText);
            }
            else
            {
                Console.WriteLine("Invalid OPML file.");
            }
        }
        void ParseOutlineElements(XElement parentElement, JObject parentJson)
        {
            foreach (XElement element in parentElement.Elements())
            {
                if (element.Name.LocalName != "outline")
                {
                    continue;
                }
                JObject childJson = new JObject();
                string text = element.Attribute("text")?.Value ?? string.Empty;
                childJson["text"] = text;

                if (parentJson.ContainsKey("children"))
                {
                    JArray childrenArray = (JArray)parentJson["children"];
                    childrenArray.Add(childJson);
                }
                else
                {
                    JArray childrenArray = new JArray
                    {
                        childJson
                    };
                    parentJson["children"] = childrenArray;
                }

                ParseOutlineElements(element, childJson); // 递归解析子元素
            }
        }

        void AddIdForJObject(JObject jObj,string parentNodeId="")
        {
            var id = string.IsNullOrEmpty(parentNodeId) ? Guid.NewGuid().ToString() : parentNodeId;
            jObj["id"] =  id ; 
            if (!jObj.ContainsKey("children"))
            {
                return;
            }
            JArray childrenArray = (JArray)jObj["children"];
            foreach (JObject child in childrenArray)
            {
                child["pid"] = id;
                AddIdForJObject(child);
            }
        }

        void xunhuanJObj(JObject jObj)
        {
            if (jObj.ContainsKey("text"))
            {
                var text = jObj["text"].ToString();
                var id = jObj["id"].ToString();
                var pid = jObj["pid"].ToString();
            }
            if (jObj.ContainsKey("children"))
            {
                JArray childrenArray = (JArray)jObj["children"];
                foreach (JObject child in childrenArray)
                {
                    xunhuanJObj(child);
                }
            }
            
            
        }
        #endregion
    }
}
