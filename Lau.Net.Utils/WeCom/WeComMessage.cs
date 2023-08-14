using Lau.Net.Utils.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.WeCom
{
    public class WeComMessage
    {
        private WeComToken _wxToken;
        private string _applicationId;
        /// <summary>
        /// 创建企业微信文档实例
        /// </summary>
        /// <param name="corpId">企业ID</param>
        /// <param name="secret">企业应用密钥</param>
        /// <param name="applicationId">企业应用ID（通过该应用发送消息）</param>
        public WeComMessage(string corpId, string secret,string applicationId)
        {
            _wxToken = new WeComToken(corpId, secret);
            _applicationId = applicationId;
        }


        public string SendText(string content, IList<string> userAccounts)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var param = new
            {
                msgtype = "text",
                touser = string.Join("|",userAccounts),
                agentid = _applicationId,
                text = new
                {
                    content
                }
            };
            var res = RestSharpUtil.Post<string>(url, param);
            return res;
        }
    }
}
