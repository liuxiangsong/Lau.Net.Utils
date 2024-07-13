using HashidsNet;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class OtherTest
    {
        [Test]
        public void ZipTest()
        {
            ZipUtil.Compress(@"E:\website\发布文件夹", @"E:\website\发布文件夹.zip");
            ZipUtil.Decompress(@"E:\website\发布文件夹.zip", @"E:\website\发布文件夹", true);
            //ZipUtil.AddFileToZip(@"E:\website\发布文件夹.zip", @"E:\website\msbuild.pubxml");
        }

        [Test]
        public void HashIdTest()
        {
            //生成时间戳
            var timeStamp = new DateTimeOffset(DateTime.UtcNow).ToUnixTimeMilliseconds();
            var hashIds = new Hashids("this is my salt");
            var shortId = hashIds.EncodeLong(timeStamp);
            var timeStamp2 = hashIds.DecodeSingleLong(shortId);
            Assert.AreEqual(timeStamp, timeStamp2);
        }

        [Test]
        public void SavFileTest()
        {
            File.Delete(@"E:\test\1asdf.txt");
            FileUtil.SaveFile(@"E:\test\1.txt", "asdf");
        }

        [Test]
        public void ComposeStringListTest()
        {  //如果其中某个字符串为空，则后台的字符串就不用了
            string[] items = { "AS", "B", "C", "", "E" };
            List<string> list = items
                .TakeWhile(item => !string.IsNullOrWhiteSpace(item))
                .SkipWhile(item => string.IsNullOrEmpty(item))
                .ToList();
        }


        [Test]
        public void Test()
        {

            var code2 = GenerateNextCode("AH99", "([A-HJ-NP-Z][A-HJ-NP-Z])([0-9][1-9])");
            var code1 = GenerateNextCode("ZN", "([A-HJ-NP-Z][A-HJ-NP-Z])");
            var code3 = GenerateNextCode("57B", "([0-9][1-9])([A-HJ-NP-Z])");
            var code4 = GenerateNextCode("A68", "([A-HJ-NP-Z])([0-9][1-9])");
        }

        #region 根据现在编码规则获取下一个编码
        /// <summary>
        /// 获取下一个编码
        /// </summary>
        /// <param name="code">当前编码</param>
        /// <param name="regex">编码正则</param>
        /// <returns></returns>
        string GenerateNextCode(string code, string regex)
        {
            List<string> splitRegexes = SplitRegex(regex);
            splitRegexes.Reverse();
            var vals = new List<string>();
            var isNewCycle = true;
            foreach (string r in splitRegexes)
            {
                Match match = Regex.Match(code, r);
                if (match.Success)
                {
                    var newValue = match.Value;
                    isNewCycle = isNewCycle && CheckIsNewCycle(match.Value, out newValue);
                    vals.Add(newValue);
                }
            }
            vals.Reverse();
            return string.Join("", vals);
        }

        /// <summary>
        /// 拆分正则表达式
        /// </summary>
        /// <param name="regex"></param>
        /// <returns></returns>
        List<string> SplitRegex(string regex)
        {
            List<string> splitRegexes = new List<string>();
            int startIndex = 0;
            for (int i = 0; i < regex.Length; i++)
            {
                if (regex[i] == '(')
                {
                    startIndex = i;
                }
                else if (regex[i] == ')')
                {
                    splitRegexes.Add(regex.Substring(startIndex, i - startIndex + 1));
                }
            }
            return splitRegexes;
        }

        /// <summary>
        /// 获取下一个字母
        /// </summary>
        /// <param name="columnName"></param>
        /// <returns></returns>
        string GetNextLetters(string columnName)
        {
            char[] chars = columnName.ToCharArray();
            Array.Reverse(chars);

            bool carry = true;
            for (int i = 0; i < chars.Length && carry; i++)
            {
                switch (chars[i])
                {
                    case 'N':
                        chars[i] = 'P';
                        carry = false;
                        break;
                    case 'H':
                        chars[i] = 'J';
                        carry = false;
                        break;
                    case 'Z':
                        chars[i] = 'A';
                        break;
                    default:
                        chars[i]++;
                        carry = false;
                        break;
                }
            }

            if (carry)
            {
                Array.Resize(ref chars, chars.Length + 1);
                chars[chars.Length - 1] = 'A';
            }

            Array.Reverse(chars);
            return new string(chars);
        }

        /// <summary>
        /// 获取下一个数字
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        string GetNextNumber(string number)
        {
            int value = int.Parse(number);
            value++; // Increment the number
            return value.ToString().PadLeft(number.Length, '0'); // Format the number as a two-digit string
        }
        /// <summary>
        /// 检查是否是新的循环
        /// </summary>
        /// <param name="val"></param>
        /// <param name="newVal"></param>
        /// <returns></returns>
        bool CheckIsNewCycle(string val, out string newVal)
        {
            if (int.TryParse(val, out _)) //数字
            {
                newVal = GetNextNumber(val);
                var isNewCycle = newVal.Length != val.Length;
                if (isNewCycle)
                {
                    newVal = "1".PadLeft(val.Length, '0');
                }
                return isNewCycle;
            }
            else //字母
            {
                newVal = GetNextLetters(val);
                var isNewCycle = newVal.Length != val.Length;
                if (isNewCycle)
                {
                    newVal = "A".PadLeft(val.Length, 'A');
                }
                return isNewCycle;
            }
        }

        #endregion





    }
}
