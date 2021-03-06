using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;

namespace Lau.Net.Utils
{
    public class NetworkUtil
    {
        //public bool IsPing(string ip)
        //{
        //    Ping pingSender = new Ping();
        //    //第一个参数为ip地址，第二个参数为ping的时间
        //    PingReply reply = pingSender.Send(ip, 120);
        //    return (reply.Status == IPStatus.Success);
        //}

        #region 获取本机的计算机名
        /// <summary>
        /// 获取本机的计算机名
        /// </summary>
        public static string GetLocalHostName()
        {
            return Dns.GetHostName();
        }
        #endregion

        #region 获取本地机器IP地址
        /// <summary>
        /// 获取本地机器IP地址
        /// </summary>
        /// <returns></returns>
        public static string GetLocalIP()
        {
            string strHostIP = string.Empty;
            IPHostEntry ipHostEntry = Dns.GetHostEntry(Environment.MachineName);
            if (ipHostEntry.AddressList.Length > 0)
            {
                strHostIP = ipHostEntry.AddressList[0].ToString();
            }
            return strHostIP;
        }
        #endregion

        #region 检测本机是否联网（互联网）
        [DllImport("wininet")]
        private extern static bool InternetGetConnectedState(out int connectionDescription, int reservedValue);
        /// <summary>
        /// 检测本机是否联网
        /// </summary>
        /// <returns>返回True则表示已联网</returns>
        public static bool IsConnectedInternet()
        {
            int i = 0;
            return InternetGetConnectedState(out i, 0);
        }
        #endregion

        #region 获取本机的局域网IP
        /// <summary>
        /// 获取本机的局域网IP
        /// </summary>        
        public static string GetLANIP()
        {
            //获取本机的IP列表,IP列表中的第一项是局域网IP，第二项是广域网IP
            IPAddress[] addressList = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
            //如果本机IP列表为空，则返回空字符串
            if (addressList.Length < 1)
            {
                return string.Empty;
            }
            //返回本机的局域网IP
            return addressList[0].ToString();
        }
        #endregion

        #region 获取本机在Internet网络的广域网IP
        /// <summary>
        /// 获取本机在Internet网络的广域网IP
        /// </summary>        
        public static string GetWANIP()
        {
            //获取本机的IP列表,IP列表中的第一项是局域网IP，第二项是广域网IP
            IPAddress[] addressList = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
            //如果本机IP列表小于2，则返回空字符串
            if (addressList.Length < 2)
            {
                return string.Empty;
            }
            //返回本机的广域网IP
            return addressList[1].ToString();
        }
        #endregion

        #region 将字符串形式的IP地址转换成IPAddress对象
        /// <summary>
        /// 将字符串形式的IP地址转换成IPAddress对象
        /// </summary>
        /// <param name="ip">字符串形式的IP地址</param>        
        public static IPAddress IPStringToIPAddress(string ip)
        {
            return IPAddress.Parse(ip);
        }
        #endregion
    }
}
