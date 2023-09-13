using Lau.Net.Utils.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.WeCom
{
    /// <summary>
    /// 企业微信通讯录管理
    /// </summary>
    public class WeComContact
    {
        private WeComToken _wxToken;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="corpID">企业ID</param>
        /// <param name="secret">应用密钥</param>
        public WeComContact(string corpID, string secret)
        {
            _wxToken = new WeComToken(corpID, secret);
        }

        #region 部门管理
        /// <summary>
        /// 获取指定部门及其下的子部门信息
        /// </summary>
        /// <param name="parentDeptId">部门Id；为空时，则获取全量组织架构</param>
        /// <returns></returns>
        public string GetDeptments(string parentDeptId)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/department/list?access_token=ACCESS_TOKEN&id=ID";
            var url = _wxToken.ReplaceUrlToken(baseUrl).Replace("=ID", $"={parentDeptId}");
            var res = RestSharpUtil.Get<string>(url);
            return res;
        }
        #endregion

        #region 成员管理
        /// <summary>
        /// 获取指定用户信息
        /// </summary>
        /// <param name="userId">用户账号</param>
        /// <returns></returns>
        public string GetUserInfo(string userId)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/user/get?access_token=ACCESS_TOKEN&userid=USERID";
            var url = _wxToken.ReplaceUrlToken(baseUrl).Replace("=USERID", $"={userId}");
            var res = RestSharpUtil.Get<string>(url);
            return res;
        }

        /// <summary>
        /// 获取部门成员
        /// </summary>
        /// <param name="deptId">部门id</param>
        /// <returns></returns>
        public string GetDeptmentUsers(string deptId)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/user/simplelist?access_token=ACCESS_TOKEN&department_id=DEPARTMENT_ID";
            var url = _wxToken.ReplaceUrlToken(baseUrl).Replace("=DEPARTMENT_ID", $"={deptId}");
            var res = RestSharpUtil.Get<string>(url);
            return res;
        }

        /// <summary>
        /// 获取成员ID列表
        /// </summary>
        /// <param name="cursor">用于分页查询的游标，字符串类型，由上一次调用返回，首次调用不填</param>
        /// <param name="limit">分页，预期请求的数据量，取值范围 1 ~ 10000</param>
        /// <returns></returns>
        public string GetUserIds(int limit,string cursor=null)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/user/list_id?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var param = new
            {
                cursor,
                limit
            };
            var res = RestSharpUtil.Post<string>(url,param);
            return res;
        }

        /// <summary>
        /// 更新成员信息
        /// </summary>
        /// <param name="userId">成员id</param>
        /// <param name="isEnable">账号是否有效</param>
        /// <param name="elseParams">其它参数</param>
        /// <returns></returns>
        public string UpdateUser(string userId,bool isEnable,JObject elseParams=null)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/user/update?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var param =new JObject
            {
                ["userid"] = userId,
                ["enable"] = isEnable?1:0,
            };
            if(elseParams != null)
            {
                param.Merge(elseParams);
            }
            var res = RestSharpUtil.Post<string>(url, JsonConvert.SerializeObject(param));
            return res;
        }

        #endregion

        #region 标签管理
        /// <summary>
        /// 获取标签列表
        /// </summary>
        /// <returns></returns>
        public string GetTags()
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/tag/list?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var res = RestSharpUtil.Get<string>(url);
            return res;
        }

        /// <summary>
        /// 获取标签成员
        /// </summary>
        /// <param name="tagId">标签ID</param>
        /// <returns></returns>
        public string GetTagUsers(string tagId)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/tag/get?access_token=ACCESS_TOKEN&tagid=TAGID";
            var url = _wxToken.ReplaceUrlToken(baseUrl).Replace("=TAGID", $"={tagId}"); ;
            var res = RestSharpUtil.Get<string>(url);
            return res;
        }
        #endregion
    }
}
