using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    public static class FontExtensions
    {
        #region 设置字体样式
        /// <summary>
        /// 设置字体样式
        /// </summary>
        /// <param name="font"></param>
        /// <param name="workbook">设置字体颜色时使用，Excel为xls格式,必须传workbook;为xlsx时，可不传workbook;为兼容xls格式，建议传</param>
        /// <param name="fontSize"></param>
        /// <param name="bold"></param>
        /// <param name="fontColor">十六进制颜色码</param>
        /// <param name="fontName"></param>
        /// <returns></returns>
        public static IFont SetFontStyle(this IFont font, IWorkbook workbook, short fontSize, bool bold, string fontColor = null, string fontName = "微软雅黑")
        {
            font.FontName = fontName;
            font.FontHeightInPoints = fontSize;
            font.Boldweight = bold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal;
            font.SetFontColor(fontColor, workbook);
            return font;
        }

        /// <summary>
        /// 设置字体样式
        /// </summary>
        /// <param name="font"></param>
        /// <param name="workbook">设置字体颜色时使用，Excel为xls格式,必须传workbook;为xlsx时，可不传workbook;为兼容xls格式，建议传</param>
        /// <param name="fontSize"></param>
        /// <param name="bold"></param>
        /// <param name="fontColor">字体颜色</param>
        /// <param name="fontName"></param>
        /// <returns></returns>
        public static IFont SetFontStyle(this IFont font, IWorkbook workbook, short fontSize, bool bold, Color fontColor, string fontName = "微软雅黑")
        {
            font.FontName = fontName;
            font.FontHeightInPoints = fontSize;
            font.Boldweight = bold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal;
            font.SetFontColor(fontColor, workbook);
            return font;
        }
        #endregion

        #region 设置字体颜色
        /// <summary>
        /// 设置字体颜色(xls格式时必须传workbook)
        /// </summary>
        /// <param name="font"></param>
        /// <param name="hexColor">十六进制颜色码</param>
        /// <param name="workbook">Excel为xls格式,必须传workbook;为xlsx时，可不传workbook;
        /// 注：xlsx是直接设置color,xls是通过颜色获得对应的颜色索引去设置，所以xls支持的颜色是有限的</param> 
        public static IFont SetFontColor(this IFont font, string hexColor, IWorkbook workbook)
        {
            if (string.IsNullOrEmpty(hexColor))
            {
                return font;
            }
            var fontColor = ColorTranslator.FromHtml(hexColor);
            font.SetFontColor(fontColor, workbook);
            return font;
        }

        /// <summary>
        /// 设置字体颜色(xls格式时必须传workbook)
        /// </summary>
        /// <param name="font"></param>
        /// <param name="fontColor">字体颜色</param>
        /// <param name="workbook">Excel为xls格式,必须传workbook;为xlsx时，可不传workbook;
        /// 注：xlsx是直接设置color,xls是通过颜色获得对应的颜色索引去设置，所以xls支持的颜色是有限的</param> 
        public static IFont SetFontColor(this IFont font, Color fontColor, IWorkbook workbook)
        {
            if (font is XSSFFont)
            {
                ((XSSFFont)font).SetColor(new XSSFColor(fontColor));
            }
            else
            {
                font.Color = workbook.ToIndexedColor(fontColor);
            }
            return font;
        }
        #endregion

    }
}
