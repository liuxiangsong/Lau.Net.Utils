using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    public static class SheetExtensions
    {
        /// <summary>
        /// 合并单元格(内容居中）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart">起始行，从0开始</param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd"></param>
        public static void MergeCells(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd)
        {
            CellRangeAddress region = new CellRangeAddress(rowStart, rowEnd, columnStart, columnEnd);
            sheet.AddMergedRegion(region);
            var mergedCell = sheet.GetRow(rowStart).GetCell(columnStart);

            var cellStyle = sheet.Workbook.CreateCellStyle();
            cellStyle.SetCellAlignmentStyle(false);
            mergedCell.CellStyle = cellStyle;
        }

        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <param name="columnWidth">列宽:多少个字符的宽度</param>
        public static void SetColumnWidth2(this ISheet sheet,int columnIndex, int? columnWidth = null)
        {
            //sheet.DefaultColumnWidth = 20; // 设置默认列宽为20个字符
            if (columnWidth == null)
            {
                sheet.AutoSizeColumn(columnIndex);
                return;
            }
            //sheet.SetColumnWidth(0, 20 * 256); // 设置第一个单元格列宽为20个字符
            sheet.SetColumnWidth(columnIndex,(int)columnWidth *256);
        }
    }
}
