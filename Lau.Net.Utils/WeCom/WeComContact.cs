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
    /// 开发文档：https://developer.work.weixin.qq.com/document/path/90193
    /// </summary>
    public class WeComContact
    {
        #region 变量及构造函数
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
        #endregion

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
        /// 获取指定用户信息(注：新白名单IP无法使用，https://developer.work.weixin.qq.com/document/path/90196）
        /// </summary>
        /// <param name="userId">用户账号</param>
        /// <returns></returns>
        public string GetUserInfo(string userId)
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }
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
            if (string.IsNullOrWhiteSpace(deptId))
            {
                throw new ArgumentNullException(nameof(deptId));
            }
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
        public string GetUserIds(int limit, string cursor = null)
        {
            if (limit <= 0 || limit > 10000)
            {
                throw new ArgumentOutOfRangeException(nameof(limit));
            }
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/user/list_id?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var param = new
            {
                cursor,
                limit
            };
            var res = RestSharpUtil.Post<string>(url, param);
            return res;
        }

        /// <summary>
        /// 更新成员信息
        /// </summary>
        /// <param name="userId">成员id</param>
        /// <param name="isEnable">账号是否有效</param>
        /// <param name="elseParams">其它参数</param>
        /// <returns></returns>
        public string UpdateUser(string userId, bool isEnable = true, JObject elseParams = null)
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/user/update?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var param = new JObject
            {
                ["userid"] = userId,
                ["enable"] = isEnable ? 1 : 0,
            };
            if (elseParams != null)
            {
                param.Merge(elseParams);
            }
            var res = RestSharpUtil.Post<string>(url, JsonConvert.SerializeObject(param));
            return res;
        }

        #endregion

        #region 标签管理
        /// <summary>
        /// 创建标签
        /// </summary>
        /// <param name="tagName">标签名称</param>
        /// <param name="tageId">标签id，非负整型，指定此参数时新增的标签会生成对应的标签id，不指定时则以目前最大的id自增。</param>
        /// <returns></returns>
        public string CreateTag(string tagName, int tageId = -1)
        {
            if (string.IsNullOrWhiteSpace(tagName))
            {
                throw new ArgumentNullException(nameof(tagName));
            }
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/tag/create?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var param = new JObject
            {
                ["tagname"] = tagName,
                ["tagName"] = tageId
            };
            var res = RestSharpUtil.Post<string>(url, JsonConvert.SerializeObject(param));
            return res;
        }

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
        public string GetTagUsers(int tagId)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/tag/get?access_token=ACCESS_TOKEN&tagid=TAGID";
            var url = _wxToken.ReplaceUrlToken(baseUrl).Replace("=TAGID", $"={tagId}");
            var res = RestSharpUtil.Get<string>(url);
            return res;
        }

        /// <summary>
        /// 增加标签成员
        /// </summary>
        /// <param name="tagId">标签ID</param>
        /// <param name="userList">企业成员ID列表，注意：userlist、departmentList不能同时为空，单次请求个数不超过1000</param>
        /// <param name="departmentList">企业部门ID列表，注意：userlist、departmentList不能同时为空，单次请求个数不超过100</param>
        /// <returns></returns>
        public string AddTagUsers(int tagId, IEnumerable<string> userList, IEnumerable<int> departmentList = null)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/tag/addtagusers?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl).Replace("=TAGID", $"={tagId}");
            var param = new JObject
            {
                ["tagid"] = tagId
            };
            if (userList.HasItem())
            {
                param["userlist"] = new JArray(userList);
            }
            if (departmentList.HasItem<int>())
            {
                param["partylist"] = new JArray(departmentList);
            }
            var res = RestSharpUtil.Post<string>(url, JsonConvert.SerializeObject(param));
            return res;
        }
        #endregion
    }
}
