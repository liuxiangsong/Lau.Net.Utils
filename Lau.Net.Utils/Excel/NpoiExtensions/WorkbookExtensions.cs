using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    /// <summary>
    /// Workbook扩展方法
    /// </summary>
    public static class WorkbookExtentions
    {
        /// <summary>
        /// 创建宋体
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fontName">字体名称,如果传空则默认为"微软雅黑"</param>
        /// <returns></returns>
        public static IFont CreateFont(this IWorkbook workbook, string fontName = "")
        {
            IFont font = workbook.CreateFont();
            if (string.IsNullOrWhiteSpace(fontName))
            {
                font.FontName = "微软雅黑";
            }
            return font;
        }

        /// <summary>
        /// 创建标题行单元格样式（内容默认居中）
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fontSize">字体大小：默认14</param>
        /// <param name="fontColor">字体颜色：默认黑色</param>
        /// <returns></returns>
        public static ICellStyle CreateHeaderStyle(this IWorkbook workbook, short fontSize = 14, short? fontColor = null,short? backgroundColor = null)
        {
            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont("");
            style.SetCellFontStyle(font, fontSize, true, fontColor ?? IndexedColors.Black.Index);
            style.SetCellAlignmentStyle(HorizontalAlignment.Center, VerticalAlignment.Center, false);
            if(backgroundColor != null)
            {
                style.SetCellBackgroundStyle((short)backgroundColor);
            }
            return style;
        }

        /// <summary>
        /// 将workbook转换成MemoryStream
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static MemoryStream ToMemoryStream(this IWorkbook workbook)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                //ms.Position = 0;
                return ms;
            }
        }

        /// <summary>
        /// 转化成short类型的颜色值
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="SystemColour"></param>
        /// <returns></returns>
        public static short ToIndexedColor(this IWorkbook workbook, Color SystemColour)
        {
            var hssfWorkbook = workbook as HSSFWorkbook;
            if (hssfWorkbook != null)
            {
                return ToIndexedColor(hssfWorkbook, SystemColour);
            }

            return IndexedColors.Black.Index;
        }

        /// <summary>
        /// 转化成short类型的颜色值
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="SystemColour"></param>
        /// <returns></returns>
        public static short ToIndexedColor(HSSFWorkbook workbook, Color SystemColour)
        {
            short s = 0;
            HSSFPalette XlPalette = workbook.GetCustomPalette();
            NPOI.HSSF.Util.HSSFColor XlColour = XlPalette.FindColor(SystemColour.R, SystemColour.G, SystemColour.B);
            if (XlColour == null)
            {
                if (NPOI.HSSF.Record.PaletteRecord.STANDARD_PALETTE_SIZE < 255)
                {
                    XlColour = XlPalette.FindSimilarColor(SystemColour.R, SystemColour.G, SystemColour.B);
                    s = XlColour.Indexed;
                }
            }
            else
                s = XlColour.Indexed;
            return s;
        }


        /// <summary>
        /// 创建日期类型单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="dateFormat">日期格式化样式</param>
        /// <returns></returns>
        public static ICellStyle CreateDateCellStyle(this IWorkbook workbook, string dateFormat = "yyyy-MM-dd")
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat(dateFormat);
            return style;
        }
    }
}
