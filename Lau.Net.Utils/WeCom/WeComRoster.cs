using System;
using System.Linq;
using Lau.Net.Utils.Net;
using Newtonsoft.Json.Linq;

namespace Lau.Net.Utils.WeCom
{
    /// <summary>
    /// 人事助手->花名册
    /// 开发文档：https://developer.work.weixin.qq.com/document/path/99130
    /// </summary>
    public class WeComRoster
    {
        #region 变量及构造函数
        private WeComToken _wxToken;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="corpID">企业ID</param>
        /// <param name="secret">应用密钥</param>
        public WeComRoster(string corpID, string secret)
        {
            _wxToken = new WeComToken(corpID, secret);
        }
        #endregion

        #region 获取员工字段配置字段信息
        /// <summary>
        /// 获取员工字段配置字段信息
        /// </summary>
        /// <returns></returns>
        public string GetFields()
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/hr/get_fields?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var res = RestSharpUtil.Get<string>(url);
            return res;
        }
        #endregion

        #region 获取员工信息
        /// <summary>
        /// 获取员工信息
        /// </summary>
        /// <param name="userId">员工ID</param>
        /// <param name="fieldIds">需要获取的字段id,如果为空，则获取所有字段信息</param>
        /// <returns></returns>
        public string GetStaffInfos(string userId, params string[] fieldIds)
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }
            var param = new JObject
            {
                ["userid"] = userId,
                ["get_all"] = fieldIds.Length < 1,
                ["fieldids"] = new JArray(fieldIds.ToList().Select(id => new JObject { ["fieldid"] = id }))
            };
            return GetStaffInfos(param);
        }

        /// <summary>
        /// 获取员工信息
        /// </summary>
        /// <param name="userId">员工ID</param>
        /// <param name="param">请求参数:fieldids,示例:{"fieldids":["12024"]}</param>
        /// <returns></returns>
        public string GetStaffInfos(string userId, JObject param)
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }
            var fieldids = param["fieldids"] as JArray;
            var tempParam = new JObject
            {
                ["userid"] = userId,
                ["get_all"] = fieldids == null || fieldids.Count == 0,
                ["fieldids"] = fieldids
            };
            return GetStaffInfos(tempParam);
        }

        private string GetStaffInfos(JObject param)
        {
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/hr/get_staff_info?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var res = RestSharpUtil.Post<string>(url, param);
            return res;
        }
        #endregion

        #region 更新员工信息
        /// <summary>
        /// 更新员工信息    
        /// </summary>
        /// <param name="userId">员工ID</param>
        /// <param name="param">请求参数:update_items、remove_items、insert_items，
        /// 示例:{"update_items":[{"fieldid":12024,"value_string":"123456"}]}</param>
        /// <returns></returns>
        public string UpdateStaffInfo(string userId, JObject param)
        {
            if (string.IsNullOrWhiteSpace(userId))
            {
                throw new ArgumentNullException(nameof(userId));
            }
            var tempParam = new JObject
            {
                ["userid"] = userId,
            };
            tempParam.Merge(param);
            var baseUrl = "https://qyapi.weixin.qq.com/cgi-bin/hr/update_staff_info?access_token=ACCESS_TOKEN";
            var url = _wxToken.ReplaceUrlToken(baseUrl);
            var res = RestSharpUtil.Post<string>(url, tempParam);
            return res;
        }
        #endregion
    }
}
