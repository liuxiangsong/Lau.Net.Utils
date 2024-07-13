using Lau.Net.Utils.Net;
using Lau.Net.Utils.Web;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.WeCom
{
    /// <summary>
    /// 企业微信聊天机器人相关方法
    /// 注：每个机器人发送的消息不能超过20条/分钟
    /// </summary>
    public class WeComRobot
    {
        #region 变量及属性
        private string _robotKey;
        /// <summary>
        /// 获取发送消息url
        /// </summary>
        /// <param name="robotKey"></param>
        /// <returns></returns>
        private string SendUrl
        {
            get
            {
                return $"https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key={_robotKey}";
            }
        }
        /// <summary>
        /// 获取上传文件url
        /// </summary>
        /// <param name="robotKey"></param>
        /// <returns></returns>
        private string UploadMediaUrl
        {
            get
            {
                return $"https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key={_robotKey}&type=file";
            }
        }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="robotKey">机器人Webhook url中的key参数</param>
        public WeComRobot(string robotKey)
        {
            _robotKey = robotKey;
        }

        #region 发送文本消息
        /// <summary>
        /// 发送文本
        /// </summary>
        /// <param name="content">文本内容，最长不超过2048个字节，必须是utf8编码</param>
        /// <param name="mentionList">userId或@all</param>
        /// <param name="mentionedMobileList">手机号码或@all</param>
        /// <returns></returns>
        public string SendText(string content, IList<string> mentionedList = null, IList<string> mentionedMobileList = null)
        {
            var textObj = new
            {
                content,
                mentioned_list = mentionedList,
                mentioned_mobile_list = mentionedMobileList
            };
            var param = new
            {
                msgtype = "text",
                text = textObj
            };
            var result = RestSharpUtil.Post<string>(SendUrl, param);
            return result;
        }

        /// <summary>
        /// 发送markdown文本消息
        /// </summary>
        /// <param name="content">markdown格式文本（最长不超过4096个字节，必须是utf8编码）</param>
        /// <returns></returns>
        public string SendMarkDown(string content)
        {
            var param = new
            {
                msgtype = "markdown",
                markdown = new
                {
                    content
                }
            };
            var result = RestSharpUtil.Post<string>(SendUrl, param);
            return result;
        }
        #endregion

        #region 发送图片
        /// <summary>
        /// 发送图片
        /// </summary>
        /// <param name="imageBytes">图片字节
        /// 注：图片（base64编码前）最大不能超过2M，支持JPG,PNG格式</param>
        /// <returns></returns>
        public string SendImage( byte[] imageBytes)
        {
            string base64String = Convert.ToBase64String(imageBytes);
            MD5CryptoServiceProvider md5Hasher = new MD5CryptoServiceProvider();
            byte[] data = md5Hasher.ComputeHash(imageBytes);
            StringBuilder sBuilder = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            var md5 = sBuilder.ToString();
            var param = string.Format("{{ \"msgtype\": \"image\",  \"image\": {{ \"base64\": \"{0}\", \"md5\": \"{1}\"  }} }} ", base64String, md5);
            var result = RestSharpUtil.Post<string>(SendUrl, param);
            return result;
        }

        /// <summary>
        /// 将DataTable转化为图片通过企业微信机器人发送消息
        /// </summary>
        /// <param name="dt">Datatable数据</param>
        /// <param name="title">标题</param>
        /// <param name="ignoreHeader">是否忽略dt表头</param>
        /// <param name="imgWidth">图片宽度：默认1024</param>
        /// <param name="imgQuality">图片的清晰度：默认为100</param>
        public string SendImage(DataTable dt,string title="", bool ignoreHeader = false, int imgWidth = 1024, int imgQuality = 100)
        {
            var imageBytes = HtmlUtil.ConvertTableToImageByte(dt, title, ignoreHeader,imgWidth,imgQuality);
            return SendImage(imageBytes);
        }

        /// <summary>
        /// 将Html转化为图片发送
        /// </summary>
        /// <param name="html">html文本字符串</param>
        /// <param name="imgWidth">图片宽度：默认1024</param>
        /// <param name="imgQuality">图片的清晰度：默认为100</param>
        /// <returns></returns>
        public string SendImage(string html, int imgWidth = 1024, int imgQuality = 100)
        {
            var imageBytes = HtmlUtil.ConvertHtmlToImageByte(html, imgWidth, imgQuality);
            return SendImage(imageBytes);
        }
        #endregion

        #region 发送文件
        /// <summary>
        /// 发送文件
        /// </summary>
        /// <param name="fileBytes">文件字节（注：文件大小在5B~20M之间）</param>
        /// <param name="fileName">文件名</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public string SendFile( byte[] fileBytes, string fileName)
        {
            var res = RestSharpUtil.PostFile<string>(UploadMediaUrl, fileBytes, fileName);
            var jResult = JsonConvert.DeserializeObject<JObject>(res);
            var media_id = jResult["media_id"].ToString();
            if (string.IsNullOrEmpty(media_id)) {
                throw new Exception(res);
            } 
            var param = new
            {
                msgtype = "file",
                file = new
                {
                    media_id
                }
            };
            res = RestSharpUtil.Post<string>(SendUrl, JsonConvert.SerializeObject(param));
            return res;
        }
        #endregion

 
    }
}
