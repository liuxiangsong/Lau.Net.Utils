using Lau.Net.Utils.Net;
using Lau.Net.Utils.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;

namespace Lau.Net.Utils.WeCom
{
    /// <summary>
    /// 企业微信应用消息推送
    /// 参考资料：https://developer.work.weixin.qq.com/document/path/90236
    /// </summary>
    public class WeComMessage
    {
        private WeComToken _wxToken;
        private string _applicationId;
        private string _baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=ACCESS_TOKEN";

        /// <summary>
        /// 创建企业微信文档实例
        /// </summary>
        /// <param name="corpId">企业ID</param>
        /// <param name="applicationId">企业应用ID（通过该应用发送消息）</param>
        /// <param name="applicationSecret">企业应用密钥</param>
        public WeComMessage(string corpId, string applicationId, string applicationSecret)
        {
            _wxToken = new WeComToken(corpId, applicationSecret);
            _applicationId = applicationId;
        }

        #region 发送文本消息
        /// <summary>
        /// 发送文本消息
        /// </summary>
        /// <param name="content">消息文本内容</param>
        /// <param name="userAccounts">接收消息成员Id，发送给企业应用全部成员时使用@all</param>
        /// <param name="departments">部门Id</param>
        /// <returns></returns>
        public string SendText(string content, IEnumerable<string> userAccounts, IEnumerable<string> departments = null)
        {
            var jContent = new JObject { { "content", content } };
            return SendMessage(jContent, userAccounts, departments);
        }

        /// <summary>
        /// 发送markdown格式消息
        /// </summary>
        /// <param name="content">markdown格式文本内容</param>
        /// <param name="userAccounts">接收消息成员Id，发送给企业应用全部成员时使用@all</param>
        /// <param name="departments">部门Id</param>
        /// <returns></returns>
        public string SendMarkDown(string content, IEnumerable<string> userAccounts, IEnumerable<string> departments = null)
        {
            var jContent = new JObject { { "content", content } };
            return SendMessage(jContent, userAccounts, departments, "markdown");
        }

        /// <summary>
        /// 发送文本卡片消息
        /// </summary>
        /// <param name="title">卡片标题，不超过128个字节，超过会自动截断</param>
        /// <param name="description">描述，不超过512个字节，超过会自动截断</param>
        /// <param name="url">跳转页面url</param>
        /// <param name="userAccounts">接收消息成员Id，发送给企业应用全部成员时使用@all</param>
        /// <param name="departments">部门Id</param>
        /// <param name="btnText">卡片下方按钮文案，默认为“详情”</param>
        /// <returns></returns>
        public string SendTextCard(string title, string description, string url, IEnumerable<string> userAccounts, IEnumerable<string> departments = null, string btnText="详情")
        {
            var jContent = new JObject {
                { "title", title },
                { "description",description },
                { "url", url },
                {"btntxt",btnText }
            };
            return SendMessage(jContent, userAccounts, departments, "textcard");
        }
        #endregion

        #region 发送图片
        /// <summary>
        /// 将DataTable转化为图片发送
        /// </summary>
        /// <param name="dt">Datatable数据</param>
        /// <param name="title">标题</param>
        /// <param name="userAccounts">接收消息成员Id，发送给企业应用全部成员时使用@all</param>
        /// <param name="departments">部门Id</param>
        public string SendImage(DataTable dt, string title, IEnumerable<string> userAccounts, IEnumerable<string> departments = null)
        {
            var imgBytes = HtmlUtil.ConvertTableToImageByte(dt, title);
            return SendImage(imgBytes, userAccounts, departments);
        }

        /// <summary>
        ///  将Html转化为图片发送
        /// </summary>
        /// <param name="html">html文本字符串</param>
        /// <param name="userAccounts">接收消息成员Id，发送给企业应用全部成员时使用@all</param>
        /// <param name="departments">部门Id</param>
        /// <returns></returns>
        public string SendImage(string html, IEnumerable<string> userAccounts, IEnumerable<string> departments = null)
        {
            var imgBytes = HtmlUtil.ConvertHtmlToImageByte(html);
            return SendImage(imgBytes, userAccounts, departments);
        }

        /// <summary>
        /// 发送图片
        /// </summary>
        /// <param name="imgBytes">图片文件字节</param>
        /// <param name="userAccounts">接收消息成员Id，发送给企业应用全部成员时使用@all</param>
        /// <param name="departments">部门Id</param>
        /// <returns></returns>
        public string SendImage(byte[] imgBytes, IEnumerable<string> userAccounts, IEnumerable<string> departments = null)
        {
            var mediaId = GetFileMediaId(TempFileType.IMAGE, imgBytes, "");
            var jContent = new JObject { { "media_id", mediaId } };
            return SendMessage(jContent, userAccounts, departments, "image");
        }
        #endregion

        /// <summary>
        /// 发送文件
        /// </summary>
        /// <param name="fileBytes">文件字节</param>
        /// <param name="fileName">文件名称</param>
        /// <param name="userAccounts">接收消息成员Id，发送给企业应用全部成员时使用@all</param>
        /// <param name="departments">部门Id</param>
        /// <returns></returns>
        public string SendFile(byte[] fileBytes, string fileName, IEnumerable<string> userAccounts, IEnumerable<string> departments = null)
        {
            var mediaId = GetFileMediaId(TempFileType.FILE, fileBytes, fileName);
            var jContent = new JObject { { "media_id", mediaId } };
            return SendMessage(jContent, userAccounts, departments, "file");
        }

        /// <summary>
        /// 发送应用消息
        /// </summary>
        /// <param name="content">消息内容</param>
        /// <param name="userAccounts">成员Id列表，最多1000个</param>
        /// <param name="departments">部门Id</param>
        /// <param name="msgType">消息类型：text、markdown、image、file、textcard、voice、video、news等</param>
        /// <returns></returns>
        public string SendMessage(JObject content, IEnumerable<string> userAccounts, IEnumerable<string> departments, string msgType = "text")
        {
            var url = _wxToken.ReplaceUrlToken(_baseUrl);
            var param = new JObject
            {
                ["msgtype"] = msgType,
                ["touser"] = JoinIds(userAccounts),
                ["toparty"] = JoinIds(departments),
                ["agentid"] = _applicationId,
                [msgType] = content
            };

            var res = RestSharpUtil.Post<string>(url, JsonConvert.SerializeObject(param));
            return res;
        }


        #region 私有方法
        /// <summary>
        /// 获取临时素材的media_id
        /// </summary>
        /// <param name="tempFileType">文件类型</param>
        /// <param name="fileBytes">文件字节流</param>
        /// <param name="fileName">文件名称</param>
        /// <returns></returns>
        private string GetFileMediaId(TempFileType tempFileType, byte[] fileBytes, string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = Guid.NewGuid().ToString();
            }
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/media/upload?access_token=ACCESS_TOKEN&type=TYPE";
            var url = _wxToken.ReplaceUrlToken(baseUrl).Replace("=TYPE", $"={tempFileType.ToString().ToLower()}");
            var res = RestSharpUtil.PostFile<string>(url, fileBytes, fileName);
            var jResult = JsonConvert.DeserializeObject<JObject>(res);
            var media_id = jResult["media_id"].ToString();
            return media_id;
        }

        private string JoinIds(IEnumerable<string> ids)
        {
            if (ids.HasItem())
            {
                return string.Join("|", ids.Distinct());
            }
            return "";
        }
        #endregion

        enum TempFileType
        {
            IMAGE,
            VIDEO,
            VOICE,
            FILE
        }
    }
}
