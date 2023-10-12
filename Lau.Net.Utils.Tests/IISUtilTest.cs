using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class IISUtilTest
    {
        [Test]
        public void StopSiteTest()
        {
            var result = IISUtil.StopSite("OA");
            Assert.AreEqual(result, true);
        }

        [Test]
        public void StartSiteTest()
        {
            var result = IISUtil.StartSite("OA");
            Assert.AreEqual(result, true);
        }

        [Test]
        public void ModifySitePhysicalPathTest()
        {
            IISUtil.ModifySitePhysicalPath("OA", @"D:\work\Learun.Application.Web");
        }

        [Test]
        public void StartApplicationPoolTest()
        {
            var result = IISUtil.StartApplicationPool("OA");
            Assert.AreEqual(result, true);
        }

        [Test]
        public void StopApplicationPoolTest()
        {
            var result = IISUtil.StopApplicationPool("OA");
            Assert.AreEqual(result, true);
        }
    }
}
