using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Lau.Net.Utils
{
    public class ProcessUtil
    {
        /// <summary>
        /// 启动外部程序
        /// </summary>
        /// <param name="appName">程序名称</param>
        /// <param name="isHiddenForm">为True表示隐藏窗体</param>
        public static void StartApp(string appName, bool isHiddenForm)
        {
            try
            {
                Process p = new Process();
                p.StartInfo.FileName = appName;
                p.StartInfo.WindowStyle = isHiddenForm ? ProcessWindowStyle.Hidden : ProcessWindowStyle.Normal;
                p.Start();
                p.WaitForExit();
                if (!p.HasExited)
                {
                    p.Kill();
                }
                else
                {
                    p.Close();
                }
                p.Dispose();
            }
            catch
            {
            }
        }

        /// <summary>
        /// 结束指定进程
        /// </summary>
        /// <param name="processName">进程名称(含扩展名)</param>
        public static void KillProcess(string processName)
        {
            Process[] allProcess = Process.GetProcesses();  //取得所有进程
            foreach (Process p in allProcess)
            {
                if (p.ProcessName.ToLower() == processName.ToLower())
                {
                    foreach (ProcessThread t in p.Threads)
                    {
                        t.Dispose();
                    }
                    p.Kill();
                }
            }
        }
    }
}
