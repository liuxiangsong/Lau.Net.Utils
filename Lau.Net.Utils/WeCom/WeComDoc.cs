using Lau.Net.Utils.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.WeCom
{
    public class WeComDoc
    {
        private WeComToken _wxToken;
        /// <summary>
        /// 创建企业微信文档实例
        /// </summary>
        /// <param name="corpID">企业ID</param>
        /// <param name="secret">应用密钥</param>
        public WeComDoc(string corpID, string secret)
        {
            _wxToken = new WeComToken(corpID, secret);
        }

        public void GetDoc()
        {
            var wxDrive = new WecomDrive(_wxToken);
            var res2 =  wxDrive.CreateSpace();
            var getDocUrl = "https://qyapi.weixin.qq.com/cgi-bin/wedoc/document/get?access_token={0}";
            CreateDoc();
            getDocUrl = string.Format(getDocUrl, _wxToken.GetToken());
            var param = new
            {
                docid = ""
            };
            var res = RestSharpUtil.Post<string>(getDocUrl,param);
        }

        public void CreateDoc()
        {
            var createDocUrl = "https://qyapi.weixin.qq.com/cgi-bin/wedoc/create_doc?access_token={0}";
            createDocUrl = string.Format(createDocUrl, _wxToken.GetToken());
            var spaceid = "";
            var param = new
            {
                spaceid ,
                fatherid = spaceid,
                doc_type = "4",
                doc_name = "测试excel2",
                admin_users = new string[] {""}
            };
            var res = RestSharpUtil.Post<string>(createDocUrl, param);
        }
    }
}
