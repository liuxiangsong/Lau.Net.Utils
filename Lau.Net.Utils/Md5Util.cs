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
            byte[] bytValue, bytHash;
            bytValue = System.Text.Encoding.UTF8.GetBytes(text);
            bytHash = md5.ComputeHash(bytValue);
            md5.Clear();
            string sTemp = "";
            for (int i = 0; i < bytHash.Length; i++)
            {
                sTemp += bytHash[i].ToString("X").PadLeft(2, '0');
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

    }
}
