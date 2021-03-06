//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
////using Microsoft.VisualBasic;

//namespace Lau.Net.Utils
//{
//    public class EncodingUtil
//    {
//        #region 简繁体转化
//        /// <summary>
//        /// 转化为简体
//        /// </summary>
//        /// <param name="text">待转化文本</param>
//        /// <returns>返回转化为简体后的文本</returns>
//        public static string LanguageToSimplified(String text)
//        {
//            return Strings.StrConv(text, VbStrConv.SimplifiedChinese, 0);
//        }

//        /// <summary>
//        /// 转化为繁体
//        /// </summary>
//        /// <param name="text">待转化文本</param>
//        /// <returns>返回转化为繁体后的文本</returns>
//        public static string LanguageToTraditional(String text)
//        {
//            return Strings.StrConv(text, VbStrConv.TraditionalChinese, 0);
//        }
//        #endregion

//        /// <summary>
//        /// 取得文件的编码
//        /// </summary>
//        /// <param name="filename">文件路径名</param>
//        /// <returns>返回文件的编码</returns>
//        public static Encoding GetFileEncode(string filename)
//        {
//            System.IO.FileStream fs = new System.IO.FileStream(filename, System.IO.FileMode.Open, System.IO.FileAccess.Read);
//            System.IO.BinaryReader br = new System.IO.BinaryReader(fs);
//            Byte[] buffer = br.ReadBytes(2);
//            if (buffer.Length > 1)
//            {
//                if (buffer[0] == 0xEF && buffer[1] == 0xBB)
//                {
//                    return System.Text.Encoding.UTF8;
//                }
//                else if (buffer[0] == 0xFE && buffer[1] == 0xFF)
//                {
//                    return System.Text.Encoding.BigEndianUnicode;
//                }
//                else if (buffer[0] == 0xFF && buffer[1] == 0xFE)
//                {
//                    return System.Text.Encoding.Unicode;
//                }
//                else if (buffer[0] == 117 && buffer[1] == 115)
//                {
//                    return Encoding.GetEncoding("big5");
//                }
//                else
//                {
//                    return System.Text.Encoding.Default;
//                }
//            }
//            else
//            {
//                return System.Text.Encoding.Default;
//            }
//        }
//    }
//}
