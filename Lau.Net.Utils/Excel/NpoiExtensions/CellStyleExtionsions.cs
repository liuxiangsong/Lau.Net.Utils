using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    /// <summary>
    /// 设置单元格样式
    /// 颜色对照表：https://www.cnblogs.com/Brainpan/p/5804167.html
    /// </summary>
    public static class CellStyleExtionsions
    {
        #region 设置样式
        /// <summary>
        /// 设置样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="setCellStyle"></param>
        /// <returns></returns>
        public static ICellStyle SetCellStyle(this ICellStyle cellStyle, Action<ICellStyle> setCellStyle)
        {
            if (setCellStyle != null)
            {
                setCellStyle(cellStyle);
            }
            return cellStyle;
        }
        #endregion

        #region 设置字体样式
        /// <summary>
        /// 设置字体样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="workbook"></param>
        /// <param name="setFontStyle">设置字体样式方法，入参为IFont</param>
        /// <returns></returns>
        public static ICellStyle SetCellFontStyle(this ICellStyle cellStyle, IWorkbook workbook,Action<IFont> setFontStyle)
        {
            var font = cellStyle.GetFont(workbook);
            setFontStyle(font);
            return cellStyle;
        }
        /// <summary>
        /// 设置字体样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="workbook"></param>
        /// <param name="fontSize"></param>
        /// <param name="bold"></param>
        /// <param name="fontColor"></param>
        /// <param name="fontName"></param>
        /// <returns></returns>
        public static ICellStyle SetCellFontStyle(this ICellStyle cellStyle, IWorkbook workbook, short fontSize, bool bold, short? fontColor = null, string fontName = "微软雅黑")
        {
            var font = cellStyle.GetFont(workbook);
            cellStyle.SetCellFontStyle(font, fontSize, bold, fontColor, fontName);
            return cellStyle;
        }

        /// <summary>
        /// 设置字体样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="font"></param>
        /// <param name="fontSize"></param>
        /// <param name="bold"></param>
        /// <param name="fontColor"></param>
        /// <param name="fontName"></param>
        public static ICellStyle SetCellFontStyle(this ICellStyle cellStyle, IFont font, short fontSize, bool bold, short? fontColor = null, string fontName = "微软雅黑")
        {
            font.FontName = fontName;
            font.FontHeightInPoints = fontSize;
            font.Boldweight = bold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal;
            font.Color = fontColor ?? IndexedColors.Black.Index;
            cellStyle.SetFont(font);
            return cellStyle;
        }
        #endregion

        #region 设置对齐方式
        /// <summary>
        /// 设置对齐方式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="wrapText"></param>
        /// <param name="horizontalAlignment">默认居中</param>
        /// <param name="verticalAlignment">默认居中</param>
        public static ICellStyle SetCellAlignmentStyle(this ICellStyle cellStyle, bool wrapText, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center)
        {
            // 设置对齐方式和自动换行
            cellStyle.Alignment = horizontalAlignment;
            cellStyle.VerticalAlignment = verticalAlignment;
            cellStyle.WrapText = wrapText;
            return cellStyle;
        }

        /// <summary>
        /// 设置对齐方式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="horizontalAlignment"></param>
        /// <param name="verticalAlignment"></param>
        /// <param name="wrapText"></param>
        public static ICellStyle SetCellAlignmentStyle(this ICellStyle cellStyle, HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment, bool wrapText)
        {
            // 设置对齐方式和自动换行
            cellStyle.Alignment = horizontalAlignment;
            cellStyle.VerticalAlignment = verticalAlignment;
            cellStyle.WrapText = wrapText;
            return cellStyle;
        }
        #endregion

        #region 设置背景色
        /// <summary>
        /// 设置背景色
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="backgroundColor">示例值：IndexedColors.LightGreen.Index</param>
        public static ICellStyle SetCellBackgroundStyle(this ICellStyle cellStyle, short backgroundColor)
        {
            cellStyle.FillForegroundColor = backgroundColor;
            cellStyle.FillPattern = FillPattern.SolidForeground;
            return cellStyle;
        }
        #endregion

        #region 设置边框样式
        /// <summary>
        /// 设置边框样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="border"></param>
        public static ICellStyle SetCellBorderStyle(this ICellStyle cellStyle, BorderStyle border = BorderStyle.Thin)
        {
            cellStyle.BorderTop = border;
            cellStyle.BorderBottom = border;
            cellStyle.BorderLeft = border;
            cellStyle.BorderRight = border;
            return cellStyle;
        }
        #endregion

        #region 设置数据展示格式
        /// <summary>
        /// 设置数据展示格式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="workbook"></param>
        /// <param name="dataFormat">
        /// "0.0"       //小数精度
        /// "0.00%"     //百分数
        /// "#,##0.0"   //按千分位展示
        /// "[DbNum2][$-804]General"  //将数字转化为汉字大写
        /// "yyyy-MM-dd HH:mm:ss aaaa"  //日期格式  aaaa展示为星期几;aaa为星期对应中文数字
        /// </param>
        /// <returns></returns>
        public static ICellStyle SetCellDataFormat(this ICellStyle cellStyle, IWorkbook workbook, string dataFormat)
        {
            cellStyle.DataFormat = workbook.CreateDataFormat().GetFormat(dataFormat);
            return cellStyle;
        } 
        #endregion
    }
}
