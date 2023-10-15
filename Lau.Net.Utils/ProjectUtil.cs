using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils
{
    /// <summary>
    /// 项目相关帮助类
    /// </summary>
    public class ProjectUtil
    {
        #region .Net 项目相关方法
        /// <summary>
        /// 构建发布.net项目
        /// </summary>
        /// <param name="msbuildExePath">MSBuild.exe执行文件的路径</param>
        /// <param name="projectFilePath">.csproj项目文件路径</param>
        /// <param name="publishProfilePath">.pubxml发布配置文件路径</param>
        /// <param name="buildLog">构建日志</param>
        public static bool BuildNetProject(string msbuildExePath, string projectFilePath, string publishProfilePath, out string buildLog)
        {
            var command = $"\"{projectFilePath}\" /p:DeployOnBuild=true /p:PublishProfile=\"{publishProfilePath}\"";
            command += " /clp:ErrorsOnly;Summary"; //仅输出错误信息;以及构建完成后的摘要信息
            return ProcessUtil.RunCommand(msbuildExePath, command, out buildLog);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="projectFilePath"></param>
        /// <param name="publishProfilePath"></param>
        /// <param name="buildLog"></param>
        /// <returns></returns>
        public static bool BuildNetCoreProject( string projectFilePath, string publishProfilePath, out string buildLog)
        {
            //var command = $"\"{projectFilePath}\" /p:DeployOnBuild=true /p:PublishProfile=\"{publishProfilePath}\"";
            //command += " /clp:ErrorsOnly;Summary"; //仅输出错误信息;以及构建完成后的摘要信息
            //
            var command = $"publish \"{projectFilePath}\" /p:PublishProfile=\"{publishProfilePath}\"";
            return ProcessUtil.RunCommand("dotnet", command, out buildLog);

        }

        /// <summary>
        /// 获取或创建PublishProfile配置文件
        /// </summary>
        /// <param name="publishUrl">发布文件保存目录路径</param>
        /// <param name="publishProfilePath">PublishProfile配置文件路径</param>
        /// <returns>返回PublishProfile配置文件路径</returns>
        public static string GetOrCreateNetPublishProfile(string publishUrl, string publishProfilePath)
        {
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
            foreach (var kv in dict)
            {
                groupElement.AppendChild(kv.Key, kv.Value);
            }
            xmlUtil.Save(publishProfilePath);
            return publishProfilePath;
        }
        #endregion

        #region 前端项目相关方法
        /// <summary>
        /// 构建前端项目
        /// </summary>
        /// <param name="projectPath">项目文件路径</param>
        /// <param name="outputPath">输出路径</param>
        /// <param name="buildLog"></param>
        /// <returns></returns>
        public static bool BuildFrontEndProject(string projectPath, string outputPath, out string buildLog)
        {
            StringBuilder outputBuilder = new StringBuilder();
            bool success = true;

            // 进入前端项目目录
            Environment.CurrentDirectory = projectPath;
            // 安装依赖
            success &= ProcessUtil.RunCommand("cmd", "/c npm install",out buildLog);
            outputBuilder.AppendLine(buildLog);
            // 构建项目
            success &= ProcessUtil.RunCommand("cmd", "/c npm run build", out buildLog);
            outputBuilder.AppendLine(buildLog);
            if (success && !string.IsNullOrEmpty(outputPath))
            {
                // 发布前端项目
                Directory.CreateDirectory(outputPath);
                Directory.Move(Path.Combine(projectPath, "dist"), outputPath);
            }

            buildLog = outputBuilder.ToString();
            return success;
        }
        #endregion
    }
}
