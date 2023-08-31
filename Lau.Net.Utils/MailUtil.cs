using Lau.Net.Utils.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lau.Net.Utils
{
    public class MailUtil
    {
        #region 变量定义及构造函数
        /// <summary>
        /// 邮件服务器地址
        /// </summary>
        private string _smtpHost;
        private int _smtpPort;
        /// <summary>
        /// 用户名
        /// </summary>
        private string _smtpUser;
        /// <summary>
        /// 密码
        /// </summary>
        private string _smtpPassword;
        /// <summary>
        /// 名称
        /// </summary>
        private string _smtpUserDisplayName;


        public MailUtil(string smtpHost, string smtpUser, string smtpPassword, string smtpUserDisplayName, int smtpPort = 25)
        {
            _smtpHost = smtpHost;
            _smtpUser = smtpUser;
            _smtpPassword = smtpPassword;
            _smtpUserDisplayName = smtpUserDisplayName;
            _smtpPort = smtpPort;
        }
        #endregion

        #region 发送邮件
        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="toList">收件人列表</param>
        /// <param name="subject">标题</param>
        /// <param name="body">邮件内容</param>
        /// <param name="ccList">抄送人列表</param>
        /// <param name="attachments">附件</param>
        /// <param name="isBodyHtml">body是否使用Html展示</param>
        /// <param name="enableSsl"></param>
        /// <returns></returns>
        public bool Send(IList<string> toList, string subject, string body, IList<string> ccList = null, IEnumerable<Attachment> attachments = null, bool isBodyHtml = true, bool enableSsl = false)
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
                if (attachments.HasItem())
                {
                    foreach (var item in attachments)
                    {
                        mailMessage.Attachments.Add(item);
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

        /// <summary>
        /// 发送邮件
        /// </summary>
        /// <param name="toList">收件人列表</param>
        /// <param name="subject">标题</param>
        /// <param name="body">邮件内容</param>
        /// <param name="attachmentFilePaths">附件路径</param>
        /// <param name="ccList">抄送人列表</param>
        /// <param name="isBodyHtml">body是否使用Html展示</param>
        /// <param name="enableSsl"></param>
        /// <returns></returns>
        public bool SendWithFiles(IList<string> toList, string subject, string body, IList<string> attachmentFilePaths, IList<string> ccList = null, bool isBodyHtml = true, bool enableSsl = false)
        {
            var attachments = FilesToAttachment(attachmentFilePaths);
            return Send(toList, subject, body, ccList, attachments, isBodyHtml, enableSsl);
        } 
        #endregion

        #region 转化成邮件附件

        /// <summary>
        /// 将DataTable转化为Excel邮件附件
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="excelName">生成的excel文件名，如果未传，则默认取dt.TableName</param>
        /// <returns></returns>
        public List<Attachment> DataTableToExcelAttachment(DataTable dt, string excelName)
        {
            if (dt == null)
            {
                return null;
            }
            var ds = new DataSet();
            ds.Tables.Add(dt);
            ds.DataSetName = dt.TableName;
            return DataTableToExcelAttachment(ds, excelName);
        }

        /// <summary>
        /// 将DataTable转化为Excel邮件附件
        /// </summary>
        /// <param name="ds">DataSet</param>
        /// <param name="excelName">生成的excel文件名，如果未传，则默认取ds.DataSetName</param>
        /// <returns></returns>
        public List<Attachment> DataTableToExcelAttachment(DataSet ds, string excelName)
        {
            if (ds == null)
            {
                return null;
            }
            var excelType = NpoiStaticUtil.ExcelType.Xlsx;
            if (string.IsNullOrEmpty(excelName))
            {
                excelName = string.IsNullOrEmpty(ds.DataSetName) ? "excel" : ds.DataSetName;
                excelName += ".xlsx";
            }
            var ext = Path.GetExtension(excelName).Trim().ToLower();
            if (string.IsNullOrEmpty(ext))
            {
                excelName += ".xlsx";
            }
            if (ext == ".xls")
            {
                excelType = NpoiStaticUtil.ExcelType.Xls;
            }
            var ms = NpoiStaticUtil.DataSetToStream(ds, true, excelType);
            var newms = new MemoryStream(ms.ToArray());
            var attachment = new Attachment(newms, excelName, MediaTypeNames.Application.Octet);
            return new List<Attachment> { attachment };
        }

        /// <summary>
        /// 将文件路径转化为邮件附件
        /// </summary>
        /// <param name="filePaths"></param>
        /// <returns></returns>
        public List<Attachment> FilesToAttachment(IEnumerable<string> filePaths)
        {
            var attachments = new List<Attachment>();
            if (filePaths.HasItem())
            {
                foreach (var item in filePaths)
                {
                    attachments.Add(new Attachment(item));
                }
            }
            return attachments;
        }
        #endregion
    }
}
