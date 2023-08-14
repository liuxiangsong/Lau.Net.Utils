using Lau.Net.Utils.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.WeCom
{
    /// <summary>
    /// 企业微信微盘
    /// </summary>
    public class WecomDrive
    {
        private WeComToken _wxToken;
        public WecomDrive(WeComToken wxToken)
        {
            _wxToken = wxToken;
        }

        public string CreateSpace()
        {
            var url = "https://qyapi.weixin.qq.com/cgi-bin/wedrive/space_create?access_token=ACCESS_TOKEN";
            url = _wxToken.ReplaceUrlToken(url);
            var param = new
            {
                space_name = "测试文档空间",
                auth_info = new List<object>
                {
                    new
                    {
                        type = 1,
                        userid = "",
                        auth = 7
                    },
                    //new
                    //{
                    //    type = 2,
                    //    departmentid = "DEPARTMENTID",
                    //    auth = 1
                    //}
                },
                space_sub_type = 0
            };
            return RestSharpUtil.Post<string>(url, param);
        }
    }
}
