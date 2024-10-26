using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;

namespace Lau.Net.Utils
{
    /// <summary>
    /// 
    /// </summary>
    public class ProcessUtil
    {
        /// <summary>
        /// 通过Vs Code 打开文件或文件夹
        /// </summary>
        /// <param name="path">文件或文件夹路径</param>
        public static void OpenPathInCode(string path)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c code \"{path}\"", // 使用 cmd 执行 code 命令
                UseShellExecute = false, // 设置为 false 以避免弹出命令窗口
                CreateNoWindow = true // 允许窗口显示
            };
            Process.Start(startInfo);
        }

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

        /// <summary>
        /// 运行命令
        /// </summary>
        /// <param name="excuteFile">命令执行文件路径或者命令名称</param>
        /// <param name="command">命令内容</param>
        /// <param name="executeLog">执行日志</param>
        /// <returns></returns>
        public static bool RunCommand(string excuteFile, string command, out string executeLog)
        {
            var processStartInfo = new ProcessStartInfo
            {
                FileName = excuteFile,
                Arguments = command,
                RedirectStandardOutput = true,
                StandardOutputEncoding = Encoding.UTF8,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            var process = new Process
            {
                StartInfo = processStartInfo
            };
            process.OutputDataReceived += Process_OutputDataReceived;
            // 启动进程并等待完成
            process.Start();
            var output = string.Empty;
            if (true)
            {
                output = process.StandardOutput.ReadToEnd(); //构建完一次性输出全部 
            }
            else
            {
                process.BeginOutputReadLine(); //实时输出 
            }
            process.WaitForExit();
            executeLog = output;
            var isRunSucess = process.ExitCode == 0;
            return isRunSucess; // 构建成功
        }

        private static void Process_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Data))
            {
                Console.WriteLine(e.Data);
            }
        }
    }
}
