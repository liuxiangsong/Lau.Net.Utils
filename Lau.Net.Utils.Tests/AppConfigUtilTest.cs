using Lau.Net.Utils.Config;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class AppConfigUtilTest
    {
        [Test]
        public void SetValueTest()
        {
            AppConfigUtil.SetValue("test", "11");
        }

        [Test]
        public void GetValueTest()
        {
            AppConfigUtil.GetValue("test");
        }
    }
}
