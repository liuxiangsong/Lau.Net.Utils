using Lau.Net.Utils.Net;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.WeChat
{
    public class WxToken
    {
        private string _tokenUrl = "https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={0}&corpsecret={1}";
        private string _corpID;//企业ID
        private string _secret;//应用密钥
        private TokenGetResult _tokenGetResult;
        private DateTime _createTokenTime; //获取token的时间

        /// <summary>
        /// 创建微信Token实例
        /// </summary>
        /// <param name="corpID">企业ID</param>
        /// <param name="secret">应用密钥</param>
        public WxToken(string corpID, string secret)
        {
            _corpID = corpID;
            _secret = secret;
            RefreshToken();
        }

        /// <summary>
        /// 获取Token
        /// </summary>
        /// <returns></returns>
        public string GetToken()
        {
            if(_tokenGetResult == null || IsTimeOut())
            {
                RefreshToken();
            }
            return _tokenGetResult.access_token;
        }

        public string ReplaceUrlToken(string url)
        {
            if (url.Contains("=ACCESS_TOKEN"))
            {
                url = url.Replace("=ACCESS_TOKEN", "=" + GetToken());
            }

            return url;
        }

        private  bool IsTimeOut()
        {
            if (_tokenGetResult == null)
            {
                RefreshToken();
            }
            var expires_in = _tokenGetResult.expires_in - 30; //(防止时间差，提前半分钟到期)
            return DateTime.Now >= this._createTokenTime.AddSeconds(expires_in);
        }

        /// <summary>
        /// 刷新Token
        /// </summary>
        private void RefreshToken()
        {
            var url = string.Format(_tokenUrl, _corpID, _secret);
            _tokenGetResult = RestSharpUtil.Get<TokenGetResult>(url);
            _createTokenTime = DateTime.Now;
        }
    }

    class TokenGetResult
    {
        /// <summary>
        /// 接口调用凭证
        /// </summary>
        /// <returns></returns>
        public string access_token { get; set; }

        /// <summary>
        /// access_token接口调用凭证超时时间，单位（秒）
        /// </summary>
        /// <returns></returns>
        public int expires_in { get; set; }
    }

}
