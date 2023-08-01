using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    internal class DateTimeUtilTest
    {
        [Test]
        public void GetWeekdayTest()
        {
            var i = DateTimeUtil.GetWeekAmount(2023);
           var str= DateTimeUtil.GetWeekday(new DateTime(2023, 7, 25));
            Assert.AreEqual(str, "星期二");
        }
    }
}
