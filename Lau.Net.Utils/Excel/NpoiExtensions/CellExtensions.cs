using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    public static class CellExtensions
    {
        /// <summary>
        /// 设置字体
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="fontSize"></param>
        /// <param name="bold"></param>
        /// <param name="fontColor"></param>
        /// <param name="foregroundColor"></param>
        /// <param name="wrapText"></param>
        /// <param name="horizontalAlignment"></param>
        /// <param name="verticalAlignment"></param>
        public static void SetCellStyle(this ICell cell, short fontSize, bool bold, short? fontColor = null, short? foregroundColor = null, bool wrapText = true, HorizontalAlignment horizontalAlignment = HorizontalAlignment.Center, VerticalAlignment verticalAlignment = VerticalAlignment.Center)
        {
            var workbook = cell.Sheet.Workbook;
            // 创建新的样式对象
            ICellStyle style = workbook.CreateCellStyle();
            IFont font = workbook.CreateFont("");
            // 创建新的字体对象
            style.SetCellFontStyle(font, fontSize, bold, fontColor);
            style.SetCellAlignmentStyle(horizontalAlignment, verticalAlignment, wrapText);

            // 设置背景色
            if (foregroundColor != null)
            {
                style.FillPattern = FillPattern.SolidForeground;
                //style.FillBackgroundColor = (short)backgroundColor;
                style.FillForegroundColor = (short)foregroundColor;
            }

            // 将新的样式对象应用到单元格对象中
            cell.CellStyle = style;
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="cell">单元格</param>
        /// <param name="value">单元格的值</param>
        /// <param name="columnType">DataColumn列的类型</param>
        /// <param name="dateStyle">日期类型的单元格样式</param>
        public static void SetCellValue(this ICell cell, object value, Type columnType, ICellStyle dateStyle)
        {
            switch (columnType.ToString())
            {
                case "System.String":
                    cell.SetCellValue(value.ToString());
                    break;
                case "System.DateTime":
                    DateTime time;
                    DateTime.TryParse(value.ToString(), out time);
                    cell.SetCellValue(time);
                    cell.CellStyle = dateStyle;
                    break;
                case "System.Boolean":
                    {
                        bool result = false;
                        bool.TryParse(value.ToString(), out result);
                        cell.SetCellValue(result);
                        break;
                    }
                case "System.Int16":
                case "System.Int32":
                case "System.Int64":
                case "System.Byte":
                    {
                        int num = 0;
                        int.TryParse(value.ToString(), out num);
                        cell.SetCellValue((double)num);
                        break;
                    }
                case "System.Decimal":
                case "System.Double":
                    {
                        double num2 = 0.0;
                        double.TryParse(value.ToString(), out num2);
                        cell.SetCellValue(num2);
                        break;
                    }
                case "System.DBNull":
                    cell.SetCellValue("");
                    break;
                default:
                    cell.SetCellValue(value.ToString());
                    break;
            }
        }



    }
}
