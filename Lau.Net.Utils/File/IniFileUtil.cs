using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Lau.Net.Utils
{ 
    public class IniFileUtil
    {
        [DllImport("kernel32", EntryPoint = "WritePrivateProfileString")]
        private static extern bool WritePrivateProfileString(string section, string key, string val, string filePath);
        //参数说明：section：INI文件中的段落；key：INI文件中的关键字；val：INI文件中关键字的数值；filePath：INI文件的完整的路径和名称。
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileString")]
        private static extern int GetPrivateProfileString(string section, string key, string defVal, byte[] retVal, int size, string filePath);
        [DllImport("kernel32", EntryPoint = "GetPrivateProfileString")]
        private static extern int GetPrivateProfileString(string section, string key, string defVal, StringBuilder retVal, int size, string filePath);
        //参数说明：section：INI文件中的段落名称；key：INI文件中的关键字；defVal：无法读取时候时候的缺省数值；retVal：读取数值；size：数值的大小；filePath：INI文件的完整路径和名称。

        private string m_FilePath;   //文件路径

        public IniFileUtil(string filePath)
        {
            if (!File.Exists(filePath))
            {
                try
                {
                    File.Create(filePath);
                }
                catch
                {
                    throw new ApplicationException("Ini文件不存在");
                }
            }
            this.m_FilePath = filePath;
        }

        /// <summary>
        /// 写Ini文件
        /// </summary>
        /// <param name="section">INI文件中的段落</param>
        /// <param name="key">INI文件中的关键字</param>
        /// <param name="value">INI文件中关键字的数值</param>
        public void WriteValue(string section, string key, string value)
        {
            if (!WritePrivateProfileString(section, key, value, this.m_FilePath))
                throw new ApplicationException("写入Ini文件出错");
        }

        /// <summary>
        /// 读Ini文件
        /// </summary>
        /// <param name="section">INI文件中的段落</param>
        /// <param name="key">INI文件中的关键字</param>
        /// <returns>返回关键字对应的值</returns>
        public string ReadValue(string section, string key)
        {
            StringBuilder sb = new StringBuilder(255);
            int i = GetPrivateProfileString(section, key, "", sb, 255, this.m_FilePath);
            return sb.ToString();
        }

        /// <summary>
        /// 读Ini文件
        /// </summary>
        /// <param name="section">INI文件中的段落</param>
        /// <param name="key">INI文件中的关键字</param>
        /// <returns>返回byte类型的值</returns>
        public byte[] ReadValues(string section, string key)
        {
            byte[] bytes = new byte[255];
            int i = GetPrivateProfileString(section, key, "", bytes, 255, this.m_FilePath);
            return bytes;
        }

        /// <summary>
        /// 删除指定Section
        /// </summary>
        /// <param name="section">INI文件中的段落</param>
        public void DeleteSection(string section)
        {
            if (!WritePrivateProfileString(section, null, null, this.m_FilePath))
                throw new ApplicationException(string.Format("无法删除Ini文件中的Section:{0}", section));
        }

        /// <summary>
        /// 删除指定Section中的Key
        /// </summary>
        /// <param name="section">INI文件中的段落</param>
        /// <param name="key">INI文件中的关键字</param>
        public void DeleteKey(string section, string key)
        {
            if (!WritePrivateProfileString(section, key, null, this.m_FilePath))
                throw new ApplicationException(string.Format("无法删除Ini文件中Section:{0},Key:{1}", section, key));
        }

        /// <summary>
        /// 取得指定Section中的所有key的名称
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        public List<string> GetKeyNames(string section)
        {
            List<string> keyList = new List<string>();
            byte[] bytes = new byte[16384];
            int i = GetPrivateProfileString(section, null, "", bytes, 16384, this.m_FilePath);
            this.GetStringsFromBuffer(bytes, i, keyList);
            return keyList;
        }

        /// <summary>
        /// 取得所有Section的名称
        /// </summary>
        /// <returns></returns>
        public List<string> GetAllSectionNames()
        {
            List<string> sectionList = new List<string>();
            byte[] bytes = new byte[65535];
            int i = GetPrivateProfileString(null, null, "", bytes, 65535, this.m_FilePath);
            this.GetStringsFromBuffer(bytes, i, sectionList);
            return sectionList;
        }

        /// <summary>
        /// 取得指定Section中的所有key和value
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        public Dictionary<string, string> GetKeyValues(string section)
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();
            List<string> keyList = this.GetKeyNames(section);
            foreach (string key in keyList)
                dict.Add(key, this.ReadValue(section, key));
            return dict;
        }

        /// <summary>
        /// 判断是否存在指定的Section
        /// </summary>
        /// <param name="section"></param>
        /// <returns></returns>
        public bool SectionExists(string section)
        {
            return this.GetAllSectionNames().Contains(section);
        }

        #region 私有方法
        private void GetStringsFromBuffer(Byte[] bytes, int bytesLen, List<string> stringList)
        {
            stringList.Clear();
            if (bytesLen != 0)
            {
                int start = 0;
                for (int i = 0; i < bytesLen; i++)
                {
                    if ((bytes[i] == 0) && ((i - start) > 0))
                    {
                        String s = Encoding.GetEncoding(0).GetString(bytes, start, i - start);
                        stringList.Add(s);
                        start = i + 1;
                    }
                }
            }
        }
        #endregion

    }
}
