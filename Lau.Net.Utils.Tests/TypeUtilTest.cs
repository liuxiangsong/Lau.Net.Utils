using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Lau.Net.Utils.Enums;
using NUnit.Framework;
using Lau.Net.Utils;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class TypeUtilTest
    {
        [Test]
        public void AsTest()
        {
            Assert.IsTrue(1.As<bool>());
            Assert.IsFalse(0.As<bool>());
            Assert.IsTrue("True".As<bool>());
            Assert.IsFalse("false".As<bool>());
            Assert.AreEqual("POST", RequestMethod.POST.As<string>());
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
        [Test]
        public void GetValueTest()
        {
            var dict = new Dictionary<string, string>
            {
                {"A","B" }
            };
            string value; 
            value = dict.GetValue(null);
            Assert.IsNull(value);
            value = dict.GetValue("test");
            Assert.IsNull(value);
            value = dict.GetValue("A");
            Assert.AreEqual(value, "B");
        }
        
    }
}
