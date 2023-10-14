using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Lau.Net.Utils
{
    public class Md5Util
    {
        #region Text MD5 Encrypt
        /// <summary>
        /// Text MD5 Encrypt
        /// </summary>
        /// <param name="text">string</param>
        /// <returns>Encrypted String</returns>
        public static string TextEncrypt(string text)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            var bytValue = Encoding.UTF8.GetBytes(text);
            var bytHash = md5.ComputeHash(bytValue);
            md5.Clear();
            var sTemp = "";
            foreach (byte t in bytHash)
            {
                sTemp += t.ToString("X").PadLeft(2, '0');
            }
            return sTemp.ToLower();
        }
        #endregion

        #region File MD5 Encrypt
        /// <summary>
        /// File MD5 Encrypt
        /// </summary>
        /// <param name="fileName">fileName</param>
        /// <returns>Encrypted String</returns>
        public static string FileEncrypt(string fileName)
        {
            try
            {
                FileStream fs = new FileStream(fileName, FileMode.Open);
                MD5 md5 = new MD5CryptoServiceProvider();
                byte[] retVal = md5.ComputeHash(fs);
                fs.Close();

                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < retVal.Length; i++)
                {
                    sb.Append(retVal[i].ToString("x2"));
                }
                return sb.ToString();
            }
            catch (Exception ex)
            {
                throw new Exception("FileEncrypt() fail,error:" + ex.Message);
            }
        }
        #endregion

        /// <summary>
        /// 
        /// </summary>
        /// <param name="encryptString"></param>
        /// <returns></returns>
        public static string Get16BitMd5(string encryptString)
        {
            byte[] result = Encoding.Default.GetBytes(encryptString);
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] output = md5.ComputeHash(result);
            string encryptResult = BitConverter.ToString(output).Replace("-", "");
            return encryptResult;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string Get32BitMd5(String input)
        {
            string cl = input;
            string pwd = "";
            MD5 md5 = MD5.Create();//实例化一个md5对像
            // 加密后是一个字节类型的数组，这里要注意编码UTF8/Unicode等的选择　
            byte[] s = md5.ComputeHash(Encoding.UTF8.GetBytes(cl));
            // 通过使用循环，将字节类型的数组转换为字符串，此字符串是常规字符格式化所得
            for (int i = 0; i < s.Length; i++)
            {
                // 将得到的字符串使用十六进制类型格式。格式后的字符是小写的字母，如果使用大写（X）则格式后的字符是大写字符 
                pwd = pwd + s[i].ToString("X2");
            }
            return pwd;
        }
    }
}
