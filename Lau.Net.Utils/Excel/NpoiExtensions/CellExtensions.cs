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
        /// <param name="cellStyle">单元格样式</param>
        public static void SetCellValue(this ICell cell, object value, Type columnType, ICellStyle cellStyle)
        {
            if(cellStyle != null)
            {
                cell.CellStyle = cellStyle;
            }            
            switch (columnType.ToString())
            {
                case "System.String":
                    cell.SetCellValue(value.ToString());
                    break;
                case "System.DateTime":
                    DateTime time;
                    DateTime.TryParse(value.ToString(), out time);
                    cell.SetCellValue(time);
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

        /// <summary>
        /// 获取单元格值
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="formulaEvaluator">公式评估器，传null时，则直接当成字符串单元格处理
        /// 可通过WorkbookFactory.CreateFormulaEvaluator(workbook)或
        /// workbook.GetCreationHelper().CreateFormulaEvaluator()创建</param>
        /// <returns></returns>
        public static string GetCellValue(this ICell cell, IFormulaEvaluator formulaEvaluator = null)
        {
            switch (cell.CellType)
            {
                //case CellType.Blank:
                //    break;
                //case CellType.Boolean:
                //    break;
                //case CellType.String:
                //    break;
                //case CellType.Error:
                //    break;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString();
                    }
                    else
                    {
                        return cell.NumericCellValue.ToString();
                    }
                case CellType.Formula:
                    try
                    {   if(formulaEvaluator == null)
                        {
                            return cell.ToString();
                        }
                        CellValue evaluatedCellValue = formulaEvaluator.Evaluate(cell);
                        switch (evaluatedCellValue.CellType)
                        {
                            case CellType.Numeric:
                                if (DateUtil.IsCellDateFormatted(cell))
                                {
                                    return cell.DateCellValue.ToString();
                                }
                                else
                                {
                                    return cell.NumericCellValue.ToString();
                                }
                            case CellType.String:
                                return cell.StringCellValue;
                            case CellType.Boolean:
                                return cell.BooleanCellValue.ToString();
                            default:
                                return null;
                        }
                    }
                    catch (Exception)
                    {
                        return cell.ToString();
                    }
                default:
                    return cell.ToString();
            }
        }

        /// <summary>
        /// 为单元格设置求和公式
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="rowStart">起始行，从0开始</param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart">起始列，从0开始</param>
        /// <param name="columnEnd">如果小于0，则取columnStart</param>
        public static void SetCellFormulaForSum(this ICell cell, int rowStart, int rowEnd, int columnStart, int columnEnd = -1)
        {
            if (columnEnd < 0)
            {
                columnEnd = columnStart;
            }
            var startExcelColumn = cell.Sheet.ConvertToExcelColumn(columnStart);
            var endExcelColumn = cell.Sheet.ConvertToExcelColumn(columnEnd);
            cell.SetCellFormula($"SUM({startExcelColumn}{rowStart+1}:{endExcelColumn}{rowEnd+1})");
        }

    }
}
