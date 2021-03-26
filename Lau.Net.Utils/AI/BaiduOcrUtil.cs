using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace Lau.Net.Utils.AI
{
    /// <summary>
    /// 百度OCR文字识别
    /// 参考文档：https://ai.baidu.com/ai-doc/OCR/zk3h7xz52
    /// </summary>
    public class BaiduOcrUtil
    {
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _token;
        private readonly HttpRequest _httpRequest;
        private  readonly string _urlOfToken="https://aip.baidubce.com/oauth/2.0/token";
        private readonly string _urlOfGeneralBasic = "https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic";
        private readonly string _urlOfAccurateBasic = " https://aip.baidubce.com/rest/2.0/ocr/v1/accurate_basic";
        /// <summary>
        /// 高精度含位置版
        /// </summary>
        private readonly string _urlOfAccurate = "https://aip.baidubce.com/rest/2.0/ocr/v1/accurate";
        public BaiduOcrUtil(string clientId, string clientSecret)
        {
            _clientId = clientId;
            _clientSecret = clientSecret;
            _httpRequest=new HttpRequest("application/x-www-form-urlencoded");
            _token = GetToken();
        }

        private string GetToken()
        {
            var dict = new Dictionary<string, object>() {
                {"grant_type", "client_credentials"},
                {"client_id", _clientId},
                {"client_secret", _clientSecret}
            }; 
            var result = _httpRequest.RequestInForm(_urlOfToken, dict);
            return result["access_token"].As<string>();
        }

        /// <summary>
        /// 取得图片中的文字
        /// </summary>
        /// <param name="filePath">图片文件路径</param>
        /// <returns>返回识别出的文字数组</returns>
        public JObject GetTextInImage(string filePath)
        {
            string url = $"{_urlOfAccurate}?access_token={this._token}";
            var dict = new Dictionary<string, object>(){
                {"image",FileUtil.GetBase64ByFilePath(filePath) }
            }; 
            var res = _httpRequest.RequestInForm(url, dict);
            return res;
        }

        public List<string> GetTextInImage(Image image)
        {
            string url = $"{_urlOfAccurateBasic}?access_token={this._token}";
            var dict = new Dictionary<string, object>(){
                {"image",image }
            };
            var result = _httpRequest.RequestInForm(url, dict);
            return result["words_result"].Select(m => m.Value<string>("words")).ToList();
        }
    }
}
