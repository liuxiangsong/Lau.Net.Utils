using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    public static class SheetExtensions
    {
        #region 合并单元格
        /// <summary>
        /// 合并单元格(内容居中）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart">起始行，从0开始</param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd"></param>
        /// <param name="cellValue">如果cellValue不为null，则设置成合并后单元格的值</param>
        public static void MergeCellsWithCenterAlign(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd, object cellValue = null)
        {
            var cellStyle = sheet.Workbook.CreateCellStyle();
            cellStyle.SetCellAlignmentStyle(false);
            sheet.MergeCells(rowStart, rowEnd, columnStart, columnEnd, cellStyle, cellValue);
        }

        /// <summary>
        /// 合并单元格
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart">起始行，从0开始</param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd"></param>
        /// <param name="mergeCellStyle">合并后单元格样式</param>
        /// <param name="cellValue">如果cellValue不为null，则设置成合并后单元格的值</param>
        public static void MergeCells(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd, ICellStyle mergeCellStyle = null, object cellValue = null)
        {
            var range = new CellRangeAddress(rowStart, rowEnd, columnStart, columnEnd);
            sheet.AddMergedRegion(range);
            sheet.MergeCellsValue(rowStart, rowEnd, columnStart, columnEnd);
            var mergedCell = sheet.GetOrCreateCell(rowStart, columnStart);
             
            if (mergeCellStyle != null)
            {                
                mergedCell.CellStyle = sheet.Workbook.MergeStyle(mergeCellStyle, mergedCell.CellStyle);
            }
            if (cellValue != null)
            {
                if( cellValue is double ||  cellValue is decimal || cellValue is int || cellValue is float)
                {
                    mergedCell.SetCellValue(cellValue.As<double>());
                }
                else
                {
                    mergedCell.SetCellValue(cellValue.As<string>());
                }                
            }   
        }

        /// <summary>
        /// 合并单元格的值（只保留第一个单元格的值，其它单元格的值清空）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart"></param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd"></param>
        private static void MergeCellsValue(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd)
        {
            // 清空其他被合并单元格的值
            for (int row = rowStart; row <= rowEnd; row++)
            {
                var currentRow = sheet.GetRow(row);
                for (int column = columnStart; column <= columnEnd; column++)
                {
                    if (row == rowStart && column == columnStart)
                    {
                        continue; // 跳过第一个单元格
                    }

                    var cell = currentRow.GetCell(column);
                    if (cell != null)
                    {
                        cell.SetCellValue(string.Empty);
                    }
                }
            }
        }
        #endregion

        #region 获取行、单元格、单元格内容
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
        /// 获取单元格字符串内容
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>如果异常则返回null</returns>
        public static string GetCellStringValue(this ISheet sheet, int rowIndex, int columnIndex)
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
        #endregion

        #region 设置列宽、 列宽自适应
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
        /// <param name="rowIndex">以该行的单元格个数进行循环设置</param>
        public static void SetColumnAutoWidth(this ISheet sheet, int rowIndex = 0)
        {
            var row = sheet.GetOrCreateRow(rowIndex);
            for (var i = 0; i < row.LastCellNum; i++)
            {
                sheet.AutoSizeColumn(i, true);
            }
        }
        #endregion

        #region 设置行、单元格样式
        /// <summary>
        /// 设置行样式（通过设置行内每个单元格的样式实现）
        /// 区别于IRow.RowStyle = style,这种方法会为整行中空的单元格设置样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="cellStyle">非空单元格样式</param>
        /// <param name="rowStyle">整行样式，设置空的单元格样式</param>
        public static void SetRowStyle(this ISheet sheet, int rowIndex, ICellStyle cellStyle, ICellStyle rowStyle=null)
        {
            var row = sheet.GetRow(rowIndex);
            if (row == null)
            {
                return;
            }
            if (rowStyle != null)
            {
                row.RowStyle = rowStyle;
            }            
            foreach (var cell in row.Cells)
            {
                cell.CellStyle = cellStyle;
            }
        }

        /// <summary>
        /// 设置指定范围单元格样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart">起如行（从0开始）</param>
        /// <param name="rowEnd">如果小于0则取最后一行的索引值</param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd">如果小于0时，则取当前行的最后一个单元格的索引</param>
        /// <param name="setCellStyleAction">设置单元格样式委托方法(每个单元格都是独立的CellStyle)</param>
        public static void SetCellsStyle(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd, Action<ICellStyle> setCellStyleAction)
        {
            if(setCellStyleAction == null)
            {
                return;
            }
            if (rowEnd < 0)
            {
                rowEnd = sheet.LastRowNum;
            }
            for (var r = rowStart; r <= rowEnd; r++)
            {
                var row = sheet.GetOrCreateRow(r);

                if (columnEnd < 0)
                {
                    columnEnd = row.LastCellNum - 1;
                }
                for (var c = columnStart; c <= columnEnd; c++)
                {
                    var cell = sheet.GetOrCreateCell(r, c);
                    var cell = sheet.GetOrCreateCell(r, c);
                    var style = sheet.Workbook.CreateCellStyle();
                    style.CloneStyleFrom(cell.CellStyle);
                    setCellStyleAction(style);
                    cell.CellStyle = style;
                }
            }
        }

        /// <summary>
        /// 设置指定范围单元格样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart">起如行（从0开始）</param>
        /// <param name="rowEnd">如果小于0则取最后一行的索引值</param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd">如果小于0时，则取当前行的最后一个单元格的索引</param>
        /// <param name="cellStyle">单元格样式</param>
        public static void SetCellsStyle(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd, ICellStyle cellStyle)
        { 
            if (rowEnd < 0)
            {
                rowEnd = sheet.LastRowNum;
            }
            for (var r = rowStart; r <= rowEnd; r++)
            {
                var row = sheet.GetOrCreateRow(r);

                if (columnEnd < 0)
                {
                    columnEnd = row.LastCellNum - 1;
                }
                for (var c = columnStart; c <= columnEnd; c++)
                {
                    var cell = sheet.GetOrCreateCell(r, c); 
                    cell.CellStyle = cellStyle;
                }
            }
        }

        /// <summary>
        /// 设置指定范围单元格边框样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart">起如行（从0开始）</param>
        /// <param name="rowEnd">如果小于0则取最后一行的索引值</param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd">如果小于0时，则取rowStart行的最后一个单元格的索引</param>
        /// <param name="borderStyle">边框样式</param>
        public static void SetCellsBorderStyle(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd =-1,BorderStyle borderStyle = BorderStyle.Thin)
        {
            var row = sheet.GetOrCreateRow(rowStart);
            if (columnEnd < 0)
            {
                columnEnd = row.LastCellNum - 1;
            }
            sheet.SetCellsStyle(rowStart, rowEnd, columnStart, columnEnd, style =>
            {
                style.SetCellBorderStyle(borderStyle);
            });
        }
        #endregion

        #region 将DataTable插入到Sheet中
        /// <summary>
        /// 将DataTable插入到Sheet中
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceTable">源数据表</param>
        /// <param name="startRowIndex">起始行索引（从0开始）</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="headerStyle">标题行样式</param>
        public static void InsertSheetByDataTable(this ISheet sheet, DataTable sourceTable, int startRowIndex, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true,  ICellStyle headerStyle = null)
        {
            var dateCellStyle =sheet.Workbook.CreateDateCellStyle(dateFormat);
            sheet.InsertSheetByDataTable(sourceTable, startRowIndex, isExportCaption, dateCellStyle, headerStyle);
        }

        /// <summary>
        /// 将DataTable插入到Sheet中
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceTable">源数据表</param>
        /// <param name="startRowIndex">起始行索引（从0开始）</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="dateCellStyle">日期单元格样式</param>
        /// <param name="headerStyle">标题行样式</param>
        public static void InsertSheetByDataTable(this ISheet sheet, DataTable sourceTable, int startRowIndex, bool isExportCaption , ICellStyle dateCellStyle = null, ICellStyle headerStyle = null)
        {
            var autoSizeRowIndex = startRowIndex;
            if (isExportCaption)
            {
                IRow firstRow = sheet.CreateRow(startRowIndex);
                if (headerStyle == null)
                {
                    headerStyle = sheet.Workbook.CreateHeaderStyle();
                }
                headerStyle.SetCellBorderStyle();
                foreach (DataColumn column in sourceTable.Columns)
                {
                    var cell = firstRow.CreateCell(column.Ordinal);
                    cell.SetCellValue(column.Caption);
                    cell.CellStyle = headerStyle;
                }
                startRowIndex += 1;
            }
            var cellStyle = sheet.Workbook.CreateCellStyleWithBorder();
            dateCellStyle.SetCellBorderStyle();
            foreach (DataRow dr in sourceTable.Rows)
            {
                var row = sheet.CreateRow(startRowIndex);
                foreach (DataColumn column in sourceTable.Columns)
                {
                    var cell = row.CreateCell(column.Ordinal );
                    var isDateType = column.DataType == typeof(DateTime);
                    cell.SetCellValue(dr[column], column.DataType, isDateType ? dateCellStyle : cellStyle);
                }
                startRowIndex++;
            }
            sheet.SetColumnAutoWidth(autoSizeRowIndex);
        }
        #endregion

        #region 将Excel列序号转化为Excel的列名
        /// <summary>
        /// 将Excel列序号转化为Excel的列名，如第1列转化为"B"
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnNumber">列序号</param>
        /// <returns></returns>
        public static string ConvertToExcelColumn(this ISheet sheet, int columnNumber)
        {
            string columnName = "";
            while (columnNumber >= 0)
            {
                int modulo = columnNumber % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                columnNumber = (columnNumber - modulo) / 26 - 1;
            }
            return columnName;
        } 
        #endregion
    }
}
