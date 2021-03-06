using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Lau.Net.Utils
{
    public class RegeditUtil
    {

        #region 自动启动程序设置

        public static bool CheckStartWithWindows()
        {
            RegistryKey regkey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run");
            if (regkey != null && (string)regkey.GetValue(Application.ProductName, "null", RegistryValueOptions.None) != "null")
            {
                Registry.CurrentUser.Flush();
                return true;
            }
            Registry.CurrentUser.Flush();
            return false;
        }

        #region 设置程序是否开机启动
        /// <summary>
        /// 设置程序是否开机启动
        /// </summary>
        /// <param name="isStartWin">是否开机启动,为Ture表示开机启动</param>
        /// <param name="appName">程序名称</param>
        /// <param name="appPath">程序路径</param>
        public static void SetStartWithWindows(bool isStartWin, string appName, string appPath)
        {
            RegistryKey regkey = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true);
            if (regkey != null)
            {
                if (isStartWin)
                {
                    regkey.SetValue(appName, appPath);
                }
                else
                {
                    regkey.DeleteValue(Application.ProductName, false);
                }
                Registry.CurrentUser.Flush();
            }
        }
        #endregion

        #endregion

        #region 读取注册表
        /// <summary>
        /// 读取注册表
        /// </summary>
        /// <param name="keyPath">注册表项,不包括基项.如：Software\\MyApp</param>
        /// <param name="name">值名称</param>
        /// <param name="keyType">注册表基项枚举</param>
        /// <returns>返回字符串</returns>
        public static string GetValue(string keyPath, string name, KeyType keyType = KeyType.HKEY_CURRENT_USER)
        {
            try
            {
                RegistryKey regKey = GetRegistryKey(keyType);
                RegistryKey regKeySub = regKey.OpenSubKey(keyPath);
                if (regKeySub != null)
                    return regKeySub.GetValue(name).ToString();
                else
                    throw new Exception("无法找到指定项");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region 写入注册表
        /// <summary>
        /// 写入注册表,如果指定项已经存在,则修改指定项的值
        /// </summary>
        /// <param name="keyPath">注册表项,不包括基项.如：Software\\MyApp</param>
        /// <param name="name">值名称</param>
        /// <param name="value">值</param>
        /// <param name="keyType">注册表基项枚举</param>
        /// <returns>返回布尔值,指定操作是否成功</returns>
        public static bool SetValue(string keyPath, string name, string value, KeyType keyType = KeyType.HKEY_CURRENT_USER)
        {
            try
            {
                RegistryKey regKey = GetRegistryKey(keyType);
                RegistryKey regKeySub = regKey.CreateSubKey(keyPath);
                if (regKeySub != null)
                {
                    regKeySub.SetValue(name, value);
                    return true;
                }
                else
                {
                    throw new Exception("要写入的项不存在");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region 删除注册表中的值
        /// <summary>
        /// 删除注册表中的值
        /// </summary>
        /// <param name="keyPath">注册表项名称,不包括基项.如：Software\\MyApp</param>
        /// <param name="name">值名称</param>
        /// <param name="keyType">注册表基项枚举</param>
        /// <returns>返回布尔值,指定操作是否成功</returns>
        public static bool DeleteValue(string keyPath, string name, KeyType keyType = KeyType.HKEY_CURRENT_USER)
        {
            try
            {
                RegistryKey regKey = GetRegistryKey(keyType);
                RegistryKey regKeySub = regKey.OpenSubKey(keyPath, true);
                if (regKeySub != null)
                    regKeySub.DeleteValue(name, true);
                else
                    throw (new Exception("无法找到指定项"));
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region  删除注册表中的指定项
        /// <summary>
        /// 删除注册表中的指定项
        /// </summary>
        /// <param name="keyPath">注册表中的项,不包括基项</param>
        /// <param name="keyType">注册表基项枚举</param>
        /// <returns>返回布尔值,指定操作是否成功</returns>
        public static bool DeleteSubKey(string keyPath, KeyType keyType = KeyType.HKEY_CURRENT_USER)
        {
            try
            {
                RegistryKey regKey = (RegistryKey)GetRegistryKey(keyType);
                regKey.DeleteSubKeyTree(keyPath);
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region 判断指定项是否存在
        /// <summary>
        /// 判断指定项是否存在
        /// </summary>
        /// <param name="keyPath">指定项字符串</param>
        /// <param name="keyType">基项枚举</param>
        /// <returns>返回布尔值,说明指定项是否存在</returns>
        public static bool IsExist(string keyPath, KeyType keyType = KeyType.HKEY_CURRENT_USER)
        {
            RegistryKey regKey = (RegistryKey)GetRegistryKey(keyType);
            return (regKey.OpenSubKey(keyPath) != null);
        }

        #endregion

        #region 检索指定项关联的所有值
        /// <summary>
        /// 检索指定项关联的所有值
        /// </summary>
        /// <param name="key">指定项字符串</param>
        /// <param name="keyType">基项枚举</param>
        /// <returns>返回指定项关联的所有值的字符串数组</returns>
        public static string[] getValues(string key, KeyType keyType = KeyType.HKEY_CURRENT_USER)
        {
            RegistryKey regKey = GetRegistryKey(keyType);
            RegistryKey regKeySub = regKey.OpenSubKey(key);
            if (regKeySub != null)
            {
                string[] names = regKeySub.GetValueNames();
                if (names.Length == 0)
                    return names;
                else
                {
                    string[] values = new string[names.Length];
                    int i = 0;
                    foreach (string name in names)
                    {
                        values[i] = regKeySub.GetValue(name).ToString();
                        i++;
                    }
                    return values;
                }
            }
            else
            {
                throw (new Exception("指定项不存在"));
            }

        }

        #endregion

        #region 私有方法（取得RegistryKey对象）
        /// <summary>
        /// 取得RegistryKey对象
        /// </summary>
        /// <param name="keyType">注册表基项枚举</param>
        /// <returns></returns>
        private static RegistryKey GetRegistryKey(KeyType keyType)
        {
            RegistryKey rk = null;
            switch (keyType)
            {
                case KeyType.HKEY_CLASS_ROOT:
                    rk = Registry.ClassesRoot;
                    break;
                case KeyType.HKEY_CURRENT_USER:
                    rk = Registry.CurrentUser;
                    break;
                case KeyType.HKEY_LOCAL_MACHINE:
                    rk = Registry.LocalMachine;
                    break;
                case KeyType.HKEY_USERS:
                    rk = Registry.Users;
                    break;
                case KeyType.HKEY_CURRENT_CONFIG:
                    rk = Registry.CurrentConfig;
                    break;
            }
            return rk;
        }

        #endregion

        #region 枚举
        /// <summary>
        /// 注册表基项枚举
        /// </summary>
        public enum KeyType
        {
            /// <summary>
            /// 注册表基项 HKEY_CLASSES_ROOT
            /// </summary>
            HKEY_CLASS_ROOT,
            /// <summary>
            /// 注册表基项 HKEY_CURRENT_USER
            /// </summary>
            HKEY_CURRENT_USER,
            /// <summary>
            /// 注册表基项 HKEY_LOCAL_MACHINE
            /// </summary>
            HKEY_LOCAL_MACHINE,
            /// <summary>
            /// 注册表基项 HKEY_USERS
            /// </summary>
            HKEY_USERS,
            /// <summary>
            /// 注册表基项 HKEY_CURRENT_CONFIG
            /// </summary>
            HKEY_CURRENT_CONFIG
        }
        #endregion
    }
}
