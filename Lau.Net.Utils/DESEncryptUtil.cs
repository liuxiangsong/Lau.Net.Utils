using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Lau.Net.Utils
{
    public class DesEncryptUtil
    {
        private static readonly string strDesKey = "!Dc-5x@3";
        private static readonly string strDesIV = "&dDPz%@3";
        #region DES Encrypt
        /// <summary>
        ///  DES Encrypt
        /// </summary>
        /// <param name="text">string</param>
        /// <returns>encrypted string</returns>
        public static string Encrypt(string text)
        {
            try
            {

                byte[] inputByteArray = Encoding.UTF8.GetBytes(text);
                DESCryptoServiceProvider des =
                    new DESCryptoServiceProvider
                    {
                        Key = ASCIIEncoding.ASCII.GetBytes(strDesKey),
                        IV = ASCIIEncoding.ASCII.GetBytes(strDesIV)
                    };

                ICryptoTransform desencrypt = des.CreateEncryptor();
                byte[] result = desencrypt.TransformFinalBlock(inputByteArray, 0, inputByteArray.Length);
                return BitConverter.ToString(result);
            }
            catch (Exception ex)
            {
                throw new Exception("Encrypt() fail,error:" + ex.Message);
            }
        }
        #endregion

        #region DES Decrypt
        /// <summary> 
        /// DES Decrypt 
        /// </summary> 
        /// <param name="text">string</param>  
        /// <returns>Decrypted string</returns> 
        public static string Decrypt(string text)
        {
            try
            {
                string[] sInput = text.Split("-".ToCharArray());
                byte[] data = new byte[sInput.Length];
                for (int i = 0; i < sInput.Length; i++)
                {
                    data[i] = byte.Parse(sInput[i], System.Globalization.NumberStyles.HexNumber);
                }
                DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                des.Key = ASCIIEncoding.ASCII.GetBytes(strDesKey);
                des.IV = ASCIIEncoding.ASCII.GetBytes(strDesIV);
                ICryptoTransform desencrypt = des.CreateDecryptor();
                byte[] result = desencrypt.TransformFinalBlock(data, 0, data.Length);
                return Encoding.UTF8.GetString(result);
            }
            catch (Exception ex)
            {
                throw new Exception("Decrypt() fail,error:" + ex.Message);
            }
        }
        #endregion
    }
}
