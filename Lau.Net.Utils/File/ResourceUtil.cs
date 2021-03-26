using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Lau.Net.Utils
{
   public class ResourceUtil
    {
        /// <summary>
        /// 获取嵌入项目中的图片
        /// </summary>
        /// <param name="nameSpace">项目命名空间</param>
        /// <param name="imgPath">文件所在路径，路径以点作分隔，
        /// 如：images文件夹下有1.png，路径则为"images.1.png'</param>
        /// <returns></returns>
        public static Image GetResourceImage(string nameSpace, string imgPath)
        {
            var name = $"{nameSpace}.{imgPath}";
            var asm = Assembly.GetEntryAssembly();
            var imgStream = asm.GetManifestResourceStream(name);
            if (imgStream != null)
                return Image.FromStream(imgStream);
            return null;
        }
    }
}
