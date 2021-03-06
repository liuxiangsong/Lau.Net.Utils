using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Lau.Net.Utils
{
    public static class AssemblyUtil
    {
        public static readonly List<Assembly> listToolStripAssembly = new List<Assembly>();

        public static void InitializeToolStripAssembly()
        {
            if (listToolStripAssembly != null && listToolStripAssembly.Count < 1)
            {
                //Assembly[] assemblyArray = AppDomain.CurrentDomain.GetAssemblies();
                //foreach (Assembly assembly in assemblyArray)
                //    if (!listToolStripAssembly.Contains(assembly))
                //        listToolStripAssembly.Add(assembly);

                string appDirectory = AppDomain.CurrentDomain.BaseDirectory;
                //string appDirectory = @"E:\file\CodeGenius\CodeGenius\bin\Debug";
                string[] dllFiles = Directory.GetFiles(appDirectory, "*.dll");
                foreach (string dllFile in dllFiles)
                {
                    try
                    {
                        Assembly assembly = Assembly.LoadFile(dllFile);
                        listToolStripAssembly.Add(assembly);
                    }
                    catch (Exception) { }
                }
            }
        }

        /// <summary>
        /// 查找所有程序集，获取 所有指定的类特性 T 的 类Type；
        /// </summary>
        public static Dictionary<Type, T> GetAttributes<T>() where T : Attribute
        {
            if (listToolStripAssembly == null || listToolStripAssembly.Count < 1)
            {
                InitializeToolStripAssembly();
            }
            return GetAttributes<T>(listToolStripAssembly);
        }


        /// <summary>
        /// 查找指定程序集，获取 所有指定的类特性 T 的 类Type；
        /// </summary>
        public static Dictionary<Type, T> GetAttributes<T>(IEnumerable<Assembly> allAssemblys) where T : Attribute
        {
            Dictionary<Type, T> list = new Dictionary<Type, T>();

            if (allAssemblys != null)
                foreach (Assembly assembly in allAssemblys/*listAssembly*/)
                {
                    try
                    {
                        Type[] types = assembly.GetTypes();
                        foreach (Type type in types)
                        {
                            if (type.IsDefined(typeof(T), false))
                            {
                                Attribute attribute = Attribute.GetCustomAttribute(type, typeof(T), false);
                                if (attribute != null)
                                {
                                    T wtattri = (T)attribute;
                                    list.Add(type, wtattri);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }

            return list;
        }

    }
}
