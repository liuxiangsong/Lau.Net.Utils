using Lau.Net.Utils.Enums;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace Lau.Net.Utils
{
   public static class HttpRequestUtil
    {

        /// <summary>
        /// 发送Http请求，模拟访问指定的Url，返回响应内容
        /// </summary>
        /// <param name="requestHeaders">请求头</param>
        /// <param name="requestData">请求体（若为GetMethod时，则该值应为空）</param>
        /// <param name="serviceUrl">要访问的Url</param>
        /// <param name="requestMethod">请求方式</param>
        /// <returns>响应内容(响应头、响应体)</returns>
        public static JObject HttpRequest(Dictionary<string, string> requestHeaders, string requestData, string serviceUrl, RequestMethod requestMethod = RequestMethod.POST)
        {
            string retString = null;
            using (var response = GetHttpWebResponse(requestHeaders, requestData, serviceUrl, requestMethod.As<string>()))
            {
                using (var myResponseStream = response.GetResponseStream())
                {
                    if (myResponseStream == null)
                    {
                        throw new InvalidOperationException();
                    }
                    var myStreamReader = new StreamReader(myResponseStream , Encoding.UTF8);
                    retString = myStreamReader.ReadToEnd();
                    myStreamReader.Close();
                    myResponseStream.Close();
                    myStreamReader = null;
                }
            }
            var jResult = JObject.Parse(retString);
            return jResult;
        }

        private static HttpWebResponse GetHttpWebResponse(Dictionary<string, string> requestHeaders, string requestData, string serviceUrl, string requestMethod)
        {
            var request = (HttpWebRequest)WebRequest.Create(serviceUrl);
            ServicePointManager.ServerCertificateValidationCallback = CheckValidationResult;
            
            request.Method = requestMethod;
            request.Headers.Set("method", serviceUrl);
            request.ContentType = "application/json";
            request.Headers.Set("Pragma", "no-cache");
            //SetHttpRequestProxy(request,null);
            if (requestHeaders != null)
            {
                foreach (var item in requestHeaders)
                {
                    request.Headers.Add(item.Key, item.Value);
                }
            }
            if (!string.IsNullOrEmpty(requestData))
            {
                //string dataStr = JsonConvert.SerializeObject(requestData);
                byte[] data = Encoding.UTF8.GetBytes(requestData);
                request.ContentLength = data.Length;
                Stream myRequestStream = null;
                try
                {
                    myRequestStream = request.GetRequestStream();
                    myRequestStream.Write(data, 0, data.Length);
                    myRequestStream.Close();
                }
                catch
                {
                    if (myRequestStream != null)
                    {
                        myRequestStream.Dispose();
                    }
                    throw;
                }
            }
            try
            {
                return (HttpWebResponse)request.GetResponse();
            }
            catch (WebException ex)
            {
                //var response = ex.Response as HttpWebResponse;
                //if (response != null)
                //{
                //    string msg = string.Format("message:{0},ResponseUri:{1},StatusCode:{2},StatusDescription{3}", ex.Message, response.ResponseUri, response.StatusCode, response.StatusDescription);
                //}
                throw ex;
            }
        }

        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true;
        }


        /// <summary>
        /// 设置请求代理
        /// </summary>
        /// <param name="request"></param>
        /// <param name="proxyInfo"></param>
        private static void SetHttpRequestProxy(HttpWebRequest request, ProxyInfo proxyInfo)
        {
            if (proxyInfo == null)
            {
                return;
            }
            NetworkCredential credentials;
            if (!string.IsNullOrEmpty(proxyInfo.ProxyDomain))
            {
                credentials = new NetworkCredential(proxyInfo.ProxyUsername, proxyInfo.ProxyPassword, proxyInfo.ProxyDomain);
            }
            else if (!string.IsNullOrEmpty(proxyInfo.ProxyUsername))
            {
                credentials = new NetworkCredential(proxyInfo.ProxyUsername, proxyInfo.ProxyPassword);
            }
            else
            {
                credentials = new NetworkCredential();
            }
            var proxy = new WebProxy(proxyInfo.ProxyUrl, Convert.ToInt32(proxyInfo.ProxyPort))
            {
                Credentials = credentials
            };
            request.UseDefaultCredentials = true;
            request.Proxy = proxy;
        }
    }

    public class ProxyInfo
    {
        public string ProxyType { get; set; }
        /// <summary>
        /// 代理地址192.168.56.110
        /// </summary>
        public string ProxyUrl { get; set; }
        /// <summary>
        /// 代理端口808
        /// </summary>
        public string ProxyPort { get; set; }
        /// <summary>
        /// 代理用户名
        /// </summary>
        public string ProxyUsername { get; set; }
        /// <summary>
        /// 代理密码
        /// </summary>
        public string ProxyPassword { get; set; }
        /// <summary>
        /// 代理域
        /// </summary>
        public string ProxyDomain { get; set; }
    }
}
