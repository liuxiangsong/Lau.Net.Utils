using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lau.Net.Utils
{
    public  class MailUtil
    {
        /// <summary>
        /// 邮件服务器地址
        /// </summary>
        private  string _smtpHost ;
        private int _smtpPort;
        /// <summary>
        /// 用户名
        /// </summary>
        private  string _smtpUser;
        /// <summary>
        /// 密码
        /// </summary>
        private  string _smtpPassword;
        /// <summary>
        /// 名称
        /// </summary>
        private  string _smtpUserDisplayName;


        public MailUtil(string smtpHost,string smtpUser,string smtpPassword,string smtpUserDisplayName,int smtpPort = 25)
        {
            _smtpHost = smtpHost;
            _smtpUser = smtpUser;
            _smtpPassword = smtpPassword;
            _smtpUserDisplayName = smtpUserDisplayName;
            _smtpPort = smtpPort;
        }

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="toList">收件人列表</param>
        /// <param name="subject">标题</param>
        /// <param name="body">邮件内容</param>
        /// <param name="ccList">抄送人列表</param>
        /// <param name="attachmentFilePaths">附件文件路径</param>
        /// <param name="isBodyHtml">body是否使用Html展示</param>
        /// <param name="enableSsl"></param>
        /// <returns></returns>
        public bool Send(IList<string> toList,string subject,string body,IList<string> ccList = null,IList<string> attachmentFilePaths = null, bool isBodyHtml = true, bool enableSsl = false)
        {
            try
            {
                var mailMessage = new MailMessage();
                toList = toList.Distinct().Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
                foreach (var to in toList)
                {
                    mailMessage.To.Add(new MailAddress(to));
                }
                if (ccList.HasItem())
                {
                    ccList = toList.Distinct().Where(t => !string.IsNullOrWhiteSpace(t)).ToList();
                    foreach (var cc in ccList)
                    {
                        mailMessage.CC.Add(new MailAddress(cc));
                    }
                }
                if(attachmentFilePaths.HasItem())
                {
                    foreach(var item in attachmentFilePaths)
                    {
                        mailMessage.Attachments.Add(new Attachment(item));
                    }                    
                }
                
                mailMessage.From = new MailAddress(_smtpUser, _smtpUserDisplayName);
                var encoding = "UTF-8";
                mailMessage.SubjectEncoding = Encoding.GetEncoding(encoding);
                mailMessage.BodyEncoding = Encoding.GetEncoding(encoding);
                mailMessage.Subject = subject;
                mailMessage.Body = body;
                mailMessage.IsBodyHtml = isBodyHtml;

                var smtpclient = new SmtpClient(_smtpHost, _smtpPort);
                smtpclient.Credentials = new System.Net.NetworkCredential(_smtpUser, _smtpPassword);
                //SSL连接
                smtpclient.EnableSsl = enableSsl;
                smtpclient.Send(mailMessage);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
