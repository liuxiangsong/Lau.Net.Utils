using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lau.Net.Utils.Enums;
using NUnit.Framework;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class TypeUtilTest
    {
        [Test]
        public void AsTest()
        {
          var result="POST";
            Assert.AreEqual(result, RequestMethod.POST.As<string>());
        }

        [Test]
        public void AdDateTime_MutipleTest()
        {
            DateTime result = new DateTime(2000, 1, 1);
            Assert.Multiple(() =>
            {
                Assert.That("2000.1.1".As<DateTime>(), Is.EqualTo(result));
                Assert.That("2000/1/1  ".As<DateTime>(), Is.EqualTo(result));
            });
        }
    }
}
