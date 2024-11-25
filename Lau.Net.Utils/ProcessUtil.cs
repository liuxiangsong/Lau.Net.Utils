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
        /// 在指定的Visual Studio实例中打开文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="projectName">项目名称，如果为null则在第一个找到的VS实例中打开</param>
        public static void OpenFileInVisualStudio(string filePath, string projectName = null)
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "devenv",
                    UseShellExecute = true
                };

                if (string.IsNullOrEmpty(projectName))
                {
                    startInfo.Arguments = $"\"{filePath}\"";
                }
                else
                {
                    var targetProcess = Process.GetProcessesByName("devenv")
                        .FirstOrDefault(p => p.MainWindowTitle.Contains(projectName));

                    if (targetProcess == null)
                    {
                        throw new InvalidOperationException($"未找到包含项目 '{projectName}' 的Visual Studio实例");
                    }

                    startInfo.Arguments = $"/Edit \"{filePath}\" /ProcessID {targetProcess.Id}";
                }

                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                throw new Exception($"在Visual Studio中打开文件失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 通过Vs Code 打开文件或文件夹
        /// </summary>
        /// <param name="filePath">文件或文件夹路径</param>
        public static void OpenFileInVsCode(string filePath)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = "cmd.exe",
                Arguments = $"/c code \"{filePath}\"", // 使用 cmd 执行 code 命令
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
