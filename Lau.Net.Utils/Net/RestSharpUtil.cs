using Lau.Net.Utils.Enums;
using RestSharp;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;

namespace Lau.Net.Utils.Net
{
    public static class RestSharpUtil
    { 
        public static T Post<T>(string url, object requestData, Dictionary<string, string> headers = null, string contentType = "application/json") where T : class
        {
            return ExecuteSync<T>(url, Method.Post, requestData, headers, contentType);
        }
        public static Task<T> PostAsync<T>(string url, object requestData, Dictionary<string, string> headers = null, string contentType = "application/json") where T : class
        {
            return ExecuteAsync<T>(url, Method.Post, requestData, headers, contentType);
        }

        public static T Get<T>(string url, Dictionary<string, string> headers = null, string contentType = "application/json") where T : class
        {
            return ExecuteSync<T>(url, Method.Get, null, headers, contentType);
        }
        public static Task<T> GetAsync<T>(string url, Dictionary<string, string> headers = null, string contentType = "application/json") where T : class
        {
            return ExecuteAsync<T>(url, Method.Get, null, headers, contentType);
        }


        public static T ExecuteSync<T>(string url, Method method, object requestBody = null, Dictionary<string, string> headers = null, string contentType = "application/json") where T : class
        {
            return Execute<T>(url, method, requestBody, contentType, headers);
        }

        public static async Task<T> ExecuteAsync<T>(string url, Method method, object requestBody = null, Dictionary<string, string> headers = null, string contentType = "application/json") where T : class
        {
            return await Task.Run(() => Execute<T>(url, method, requestBody, contentType, headers));
        }

        public static T PostFile<T>(string url, byte[] fileBytes, string fileName) where T : class
        {
            var client = new RestClient(url);
            var request = new RestRequest("", Method.Post);
            //设置请求头头部boundary的值不要用双引号括起来
            request.MultipartFormQuoteBoundary = false;
            request.AddFile("", fileBytes, fileName);
            if (typeof(T) == typeof(string))
            {
                return client.Execute(request).Content as T;
            }
            return client.Execute<T>(request).Data;
        }

        #region 私有方法

        private static T Execute<T>(string url, Method method, object requestBody = null, string contentType = "application/json", Dictionary<string, string> headers = null) where T:class 
        {
            var client = new RestClient(url);
            var request = new RestRequest("",method);

            if (headers != null)
            {
                foreach (var header in headers)
                {
                    request.AddHeader(header.Key, header.Value);
                }
            }

            if (requestBody != null)
            {
                if (contentType == "application/json")
                {
                    request.AddJsonBody(requestBody);
                }
                else if (contentType == "multipart/form-data")
                {
                    var dict = requestBody as Dictionary<string, object>;
                    if(dict != null)
                    {
                        foreach (var item in dict)
                        {
                            var bytes = item.Value as byte[];
                            if (bytes != null)
                            {
                                request.AddFile(item.Key, bytes, "file", "application/octet-stream");
                            }
                            else
                            {
                                var isString = item.Value is string;
                                var type = isString ? ParameterType.RequestBody : ParameterType.QueryString;
                                request.AddParameter(item.Key, isString ? item.Value.ToString() : item.Value, type);
                            }
                        }
                    }                    
                }
                else
                {
                    request.AddParameter(contentType, requestBody, ParameterType.RequestBody);
                    request.AddHeader("Content-Type", contentType);
                }
            }
            if (typeof(T) == typeof(string))
            {
                return client.Execute(request).Content as T;
            }
            var response = client.Execute<T>(request);
            if (response.ErrorException != null)
            {
                throw new Exception(response.ErrorException.Message);
            }
            return response.Data;
        }

        #endregion
    }
}
