using Lau.Net.Utils.Net;
using System.Collections.Generic;

namespace Lau.Net.Utils.WeCom
{
    /// <summary>
    /// 企业微信外部联系人
    /// 参考文档：https://developer.work.weixin.qq.com/document/path/96307
    /// </summary>
    public class WeComExternalContact
    {
        private readonly WeComToken _token;
        private const string GET_FOLLOW_USER_LIST = "https://qyapi.weixin.qq.com/cgi-bin/externalcontact/get_follow_user_list?access_token=ACCESS_TOKEN";
        private const string GET_EXTERNAL_CONTACT_LIST = "https://qyapi.weixin.qq.com/cgi-bin/externalcontact/list?access_token=ACCESS_TOKEN&userid={0}";
        private const string GET_EXTERNAL_CONTACT_DETAIL = "https://qyapi.weixin.qq.com/cgi-bin/externalcontact/get?access_token=ACCESS_TOKEN&external_userid={0}";

        public WeComExternalContact(WeComToken token)
        {
            _token = token;
        }

        /// <summary>
        /// 获取配置了客户联系功能的成员列表
        /// </summary>
        public List<string> GetFollowUsers()
        {
            var url = _token.ReplaceUrlToken(GET_FOLLOW_USER_LIST);
            var result = RestSharpUtil.Get<FollowUserResult>(url);
            return result.follow_user;
        }

        /// <summary>
        /// 获取指定成员的客户列表
        /// </summary>
        /// <param name="userId">企业成员的userid</param>
        public List<string> GetExternalContactList(string userId)
        {
            var url = _token.ReplaceUrlToken(string.Format(GET_EXTERNAL_CONTACT_LIST, userId));
            var result = RestSharpUtil.Get<ExternalContactListResult>(url);
            return result.external_userid;
        }

        /// <summary>
        /// 获取客户详情
        /// </summary>
        /// <param name="externalUserId">外部联系人的userid</param>
        public ExternalContactDetail GetExternalContactDetail(string externalUserId)
        {
            var url = _token.ReplaceUrlToken(string.Format(GET_EXTERNAL_CONTACT_DETAIL, externalUserId));
            return RestSharpUtil.Get<ExternalContactDetail>(url);
        }
    }

    public class FollowUserResult
    {
        public int errcode { get; set; }
        public string errmsg { get; set; }
        public List<string> follow_user { get; set; }
    }

    public class ExternalContactListResult
    {
        public int errcode { get; set; }
        public string errmsg { get; set; }
        public List<string> external_userid { get; set; }
    }

    public class ExternalContactDetail
    {
        public int errcode { get; set; }
        public string errmsg { get; set; }
        public ExternalContact external_contact { get; set; }
        public List<FollowUser> follow_user { get; set; }
    }

    public class ExternalContact
    {
        public string external_userid { get; set; }
        public string name { get; set; }
        public string avatar { get; set; }
        public int type { get; set; }
        public int gender { get; set; }
        public string unionid { get; set; }
    }

    public class FollowUser
    {
        public string userid { get; set; }
        public string remark { get; set; }
        public string description { get; set; }
        public long createtime { get; set; }
        public List<string> tags { get; set; }
    }
}