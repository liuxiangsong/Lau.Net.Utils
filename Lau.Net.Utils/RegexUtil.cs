using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Lau.Net.Utils
{
    public static class RegexUtil
    {
        /// <summary>
        /// 判断字符串中是否包含中文
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public static bool IsContainsChinese(this string text)
        {
            var containsChinese = Regex.IsMatch(text, @"[\u4e00-\u9fa5]");
            return containsChinese;
        }
    }
}
