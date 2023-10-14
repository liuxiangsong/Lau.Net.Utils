using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class ProjectUtilTest
    {
        [Test]
        public void BuildNetProject()
        {
            var msbuildExePath = @"D:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\msbuild.exe";
            var projectFilePath = @"E:\MyCode\test\msbuild-demos\webDemo1\webDemo1.csproj";
            var publishProfile = ProjectUtil.GetOrCreateNetPublishProfile(@"E:\website\发布文件夹", @"E:\website\msbuild.pubxml");
            var buildLog = string.Empty;
            var result = ProjectUtil.BuildNetProject(msbuildExePath, projectFilePath, publishProfile,out buildLog);
            Assert.IsTrue(result);
        }

        [Test]
        public void BuildFrontEndProject()
        {
            var buildLog = string.Empty;
            var result = ProjectUtil.BuildFrontEndProject(@"D:\vueDemo", @"E:\website\发布文件夹", out buildLog);
            Assert.IsTrue(result);
        }
    }
}
