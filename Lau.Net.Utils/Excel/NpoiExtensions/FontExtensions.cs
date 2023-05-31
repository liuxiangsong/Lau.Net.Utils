using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    public static class FontExtensions
    {
        /// <summary>
        /// 设置字体样式
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <param name="bold"></param>
        /// <param name="fontColor"></param>
        /// <param name="fontName"></param>
        /// <returns></returns>
        public static IFont SetFontStyle(this IFont font, short fontSize, bool bold, short? fontColor = null, string fontName = "微软雅黑")
        {
            font.FontName = fontName;
            font.FontHeightInPoints = fontSize;
            font.Boldweight = bold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal;
            font.Color = fontColor ?? IndexedColors.Black.Index;
            return font;
        }
    }
}
