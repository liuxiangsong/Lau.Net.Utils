using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.UI.WebControls;
using System.Windows.Forms;

namespace Lau.Net.Utils
{
    /// <summary>
    /// MsbuildUtil
    /// </summary>
    public class MsbuildUtil
    {
        /// <summary>
        /// 发布.net项目
        /// </summary>
        /// <param name="msbuildExePath">MSBuild.exe执行文件的路径</param>
        /// <param name="projectFilePath">.csproj项目文件路径</param>
        /// <param name="publishProfilePath">.pubxml发布配置文件路径</param>
        /// <returns></returns>
        public static bool PublishProject(string msbuildExePath, string projectFilePath,string publishProfilePath)
        {  
            //.\msbuild "D:\work\Learun.Application.Web\Learun.Application.Web.csproj" /p:Configuration=Release /p:DeployOnBuild=true /p:PublishProfile=FolderProfile /p:OutputPath=E:\website\www-oa
            var msBuildArguments = $"\"{projectFilePath}\" /p:DeployOnBuild=true /p:PublishProfile=\"{publishProfilePath}\"";
            // 创建用于执行命令行的进程
            var processStartInfo = new ProcessStartInfo
            {
                FileName = msbuildExePath, // 指定MSBuild执行文件的路径
                //FileName = GetMsBuildPath(),
                Arguments = msBuildArguments,
                RedirectStandardOutput = true,
                UseShellExecute = false,
                //CreateNoWindow = false
            }; 
            var process = new Process
            {
                StartInfo = processStartInfo
            };
            process.OutputDataReceived += Process_OutputDataReceived;
            // 启动进程并等待完成
            process.Start();
            process.BeginOutputReadLine();
            process.WaitForExit();

            var isBuildSucess = process.ExitCode == 0;
            // 检查发布是否成功
            if (isBuildSucess)
            {
                Console.WriteLine("项目发布成功！");
            }
            else
            {
                Console.WriteLine("项目发布失败！");
            }
            Console.ReadKey();
            return isBuildSucess;
        }

        private static void Process_OutputDataReceived(object sender, DataReceivedEventArgs e)
        {
            if (!string.IsNullOrEmpty(e.Data))
            {
                Console.WriteLine(e.Data);
            }
        }

        /// <summary>
        /// 获取或创建PublishProfile配置文件
        /// </summary>
        /// <param name="publishUrl">发布文件保存目录路径</param>
        /// <param name="publishProfilePath">PublishProfile配置文件路径；如果为空，则取程序启动目录下的msbulid.pubxml文件路径</param>
        /// <returns>返回PublishProfile配置文件路径</returns>
        public static string GetOrCreatePublishProfile(string publishUrl,string publishProfilePath="")
        {
            if (string.IsNullOrEmpty(publishProfilePath))
            {
                publishProfilePath = Path.Combine(Application.StartupPath, "msbuild.pubxml");
            }
            var xmlUtil = new XmlUtil("Project", "xmlns", "http://schemas.microsoft.com/developer/msbuild/2003");
            var root = xmlUtil.Document.DocumentElement;
            var groupElement = root.AppendChild("PropertyGroup");
            var dict = new Dictionary<string, string>
            {
                {"DeleteExistingFiles", "False" },
                {"ExcludeApp_Data", "False"},
                {"LaunchSiteAfterPublish","True" },
                {"LastUsedBuildConfiguration","Release" },
                {"LastUsedPlatform","Any CPU" },
                {"PublishProvider","FileSystem" },
                {"PublishUrl", publishUrl},
                {"WebPublishMethod","FileSystem" }
            };
            foreach(var kv in dict)
            {
                groupElement.AppendChild(kv.Key,kv.Value);
            }            
            xmlUtil.Save(publishProfilePath);
            return publishProfilePath;
        }
    }
}
