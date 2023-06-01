using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    public static class SheetExtensions
    {
        /// <summary>
        /// 合并单元格(默认内容居中）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart">起始行，从0开始</param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd"></param>
        /// <param name="cellStyle">合并后单元格样式</param>
        public static void MergeCells(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd, ICellStyle cellStyle = null)
        {
            CellRangeAddress region = new CellRangeAddress(rowStart, rowEnd, columnStart, columnEnd);
            sheet.AddMergedRegion(region);
            var mergedCell = sheet.GetOrCreateCell(rowStart, columnStart);

            if (cellStyle == null)
            {
                cellStyle = sheet.Workbook.CreateCellStyle();
                cellStyle.SetCellAlignmentStyle(false);
            }
            mergedCell.CellStyle = cellStyle;
        }
        /// <summary>
        /// 获取表单单元格，如果单元格不存在则创建
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        public static ICell GetOrCreateCell(this ISheet sheet, int rowIndex, int columnIndex)
        {
            var row = sheet.GetOrCreateRow(rowIndex);
            var cell = row.GetCell(columnIndex);
            if (cell == null)
            {
                cell = row.CreateCell(columnIndex);
            }
            return cell;
        }
        /// <summary>
        /// 获取表单行，如果行不存在则创建
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public static IRow GetOrCreateRow(this ISheet sheet, int rowIndex)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                row = sheet.CreateRow(rowIndex);
            }
            return row;
        }

        /// <summary>
        /// 修改指定范围单元格样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="workbook"></param>
        /// <param name="rowStart">起如行（从0开始）</param>
        /// <param name="rowEnd">如果小于0则取最后一行的索引值</param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd">如果小于0时，则取当前行的最后一个单元格的索引</param>
        /// <param name="modifyCellStyle">修改单元格样式委托方法</param>
        public static void ModifyCellsStyle(this ISheet sheet,IWorkbook workbook, int rowStart, int rowEnd, int columnStart, int columnEnd,Action<ICellStyle> modifyCellStyle)
        {
            if(rowEnd < 0)
            {
                rowEnd = sheet.LastRowNum;
            }
            for(var r = rowStart; r <= rowEnd; r++)
            {
                var row = sheet.GetOrCreateRow(r);

                if (columnEnd < 0)
                {
                    columnEnd = row.LastCellNum;
                }
                for (var c = columnStart; c<= columnEnd; c++)
                {
                    ICell cell = sheet.GetOrCreateCell(r, c);
                    //ICell cell = row.GetCell(c);
                    //if (cell == null)
                    //{
                    //    cell = row.CreateCell(c);
                    //    cell.SetCellValue(string.Empty);
                    //}
                    var style = workbook.CreateCellStyle();
                    style.CloneStyleFrom(cell.CellStyle);
                    modifyCellStyle(style);
                    cell.CellStyle = style;
                }
            }
        }


        /// <summary>
        /// 设置列宽
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex">列索引，从0开始</param>
        /// <param name="columnWidth">列宽:多少个字符的宽度</param>
        public static void SetColumnWidth2(this ISheet sheet, int columnIndex, int? columnWidth = null)
        {
            //sheet.DefaultColumnWidth = 20; // 设置默认列宽为20个字符
            if (columnWidth == null)
            {
                sheet.AutoSizeColumn(columnIndex);
                return;
            }
            //sheet.SetColumnWidth(0, 20 * 256); // 设置第一个单元格列宽为20个字符
            sheet.SetColumnWidth(columnIndex, (int)columnWidth * 256);
        }

        /// <summary>
        /// 设置表单列宽自适应内容
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">指该索引行的内容进行自适应</param>
        public static void SetColumnAutoWidth(this ISheet sheet,int rowIndex = 0)
        {
            var row = sheet.GetOrCreateRow(rowIndex);
            for(var i =0; i < row.LastCellNum; i++)
            {
                sheet.AutoSizeColumn(i,true);
            }
        }

        /// <summary>
        /// 设置行样式（通过设置行内每个单元格的样式实现）
        /// 区别于IRow.RowStyle = style,这种方法会为整行中空的单元格设置样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="cellStyle"></param>
        public static void SetRowStyle(this ISheet sheet, int rowIndex, ICellStyle cellStyle)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                return;
            }
            foreach (var cell in row.Cells)
            {
                cell.CellStyle = cellStyle;
            }
        }

        /// <summary>
        /// 获取单元格字符串内容
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>如果异常则返回null</returns>
        public static string GetCellStringValue(this ISheet sheet,int rowIndex,int columnIndex)
        {
            try
            {
                return sheet.GetRow(rowIndex).GetCell(columnIndex).ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 将将DataTable添加至Sheet中
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceTable">源数据表</param>
        /// <param name="startRowIndex">起始行索引（从0开始）</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="dateCellStyle">日期单元格样式</param>
        /// <param name="headerStyle">标题行样式</param>
        public static void InsertSheetByDataTable(this ISheet sheet, DataTable sourceTable, int startRowIndex, bool isExportCaption = true, ICellStyle dateCellStyle = null, ICellStyle headerStyle = null)
        {
            var autoSizeRowIndex = startRowIndex;
            if (isExportCaption)
            {
                IRow firstRow = sheet.CreateRow(startRowIndex);
                foreach (DataColumn column in sourceTable.Columns)
                {
                    ICell cell = firstRow.CreateCell(column.Ordinal);
                    cell.SetCellValue(column.Caption);
                    if (headerStyle != null)
                    {
                        cell.CellStyle = headerStyle;
                    }
                }
                startRowIndex += 1;
            }

            foreach (DataRow dr in sourceTable.Rows)
            {
                IRow row = sheet.CreateRow(startRowIndex);
                foreach (DataColumn column in sourceTable.Columns)
                {
                    ICell cell = row.CreateCell(column.Ordinal);
                    cell.SetCellValue(dr[column], column.DataType, dateCellStyle);
                }
                startRowIndex++;
            }
            sheet.SetColumnAutoWidth(autoSizeRowIndex);
        }
    }
}
