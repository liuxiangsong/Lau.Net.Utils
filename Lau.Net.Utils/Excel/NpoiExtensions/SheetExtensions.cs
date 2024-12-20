﻿using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Lau.Net.Utils.Excel.NpoiExtensions
{
    /// <summary>
    /// Sheet扩展类
    /// </summary>
    public static class SheetExtensions
    {
        #region 合并单元格、获取多个单元格之和
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
        /// <param name="rowStart"></param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd"></param>
        /// <param name="isMergeCellsValueWithSum">为true则用加法汇总合并单元格的值，否则取第一个单元格的值作为合并后的值</param>
        /// <param name="mergeCellStyle"></param>
        public static void MergeCells(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd, bool isMergeCellsValueWithSum, ICellStyle mergeCellStyle = null)
        {
            object cellValue = null;
            if (isMergeCellsValueWithSum)
            {
                cellValue = sheet.GetCellsValueWithSum(rowStart, rowEnd, columnStart, columnEnd);
            }
            sheet.MergeCells(rowStart, rowEnd, columnStart, columnEnd, mergeCellStyle, cellValue);
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
            sheet.MergeCellsValueWithEmpty(rowStart, rowEnd, columnStart, columnEnd);

            var mergedCell = sheet.GetOrCreateCell(rowStart, columnStart);

            if (mergeCellStyle != null)
            {
                mergedCell.CellStyle = sheet.Workbook.MergeStyle(mergeCellStyle, mergedCell.CellStyle);
            }
            if (cellValue != null)
            {
                if (cellValue is double || cellValue is decimal || cellValue is int || cellValue is float)
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
        private static void MergeCellsValueWithEmpty(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd)
        {
            var boderStyle = sheet.Workbook.CreateCellStyleWithBorder();
            // 清空其他被合并单元格的值
            for (int row = rowStart; row <= rowEnd; row++)
            {
                //var currentRow = sheet.GetOrCreateRow(row);
                for (int column = columnStart; column <= columnEnd; column++)
                {
                    if (row == rowStart && column == columnStart)
                    {
                        continue; // 跳过第一个单元格
                    }
                    var cell = sheet.GetOrCreateCell(row, column, boderStyle);
                    if (cell != null)
                    {
                        cell.SetCellValue(string.Empty);
                    }
                }
            }
        }

        /// <summary>
        /// 获取指定范围内所有单元格值之和
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart"></param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd"></param>
        /// <returns></returns>
        public static decimal GetCellsValueWithSum(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd)
        {
            decimal sum = 0;
            // 清空其他被合并单元格的值
            for (int row = rowStart; row <= rowEnd; row++)
            {
                var currentRow = sheet.GetRow(row);
                for (int column = columnStart; column <= columnEnd; column++)
                {
                    var cell = currentRow.GetCell(column);
                    sum += cell.ToString().As<decimal>();
                }
            }
            return sum;
        }
        #endregion

        #region 获取行、单元格、单元格内容、指定行列数
        /// <summary>
        /// 获取表单单元格，如果单元格不存在则创建
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="cellStyle">单元格样式</param>
        /// <returns></returns>
        public static ICell GetOrCreateCell(this ISheet sheet, int rowIndex, int columnIndex, ICellStyle cellStyle = null)
        {
            var row = sheet.GetOrCreateRow(rowIndex);
            var cell = row.GetCell(columnIndex);
            if (cell == null)
            {
                cell = row.CreateCell(columnIndex);
            }
            if (cellStyle != null)
            {
                cell.CellStyle = cellStyle;
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
                return sheet.GetOrCreateCell(rowIndex, columnIndex).ToString();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 设置单元格的值
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="cellValue">单元格的值</param>
        /// <param name="columnType">单元格值类型</param>
        /// <param name="cellStyle">单元格样式</param>
        public static void SetCellValue(this ISheet sheet, int rowIndex, int columnIndex, object cellValue, Type columnType = null, ICellStyle cellStyle = null)
        {
            try
            {
                sheet.GetOrCreateCell(rowIndex, columnIndex).SetCellValue(cellValue, columnType, cellStyle);
            }
            catch
            {

            }
        }

        /// <summary>
        /// 获取指定行的列数
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <returns></returns>
        public static int GetRowColumnCount(this ISheet sheet, int rowIndex)
        {
            return sheet.GetOrCreateRow(rowIndex).LastCellNum;
        }
        #endregion

        #region 设置列宽、 列宽自适应
        ///// <summary>
        ///// 设置列宽
        ///// </summary>
        ///// <param name="sheet"></param>
        ///// <param name="columnWidth">列宽:多少个字符的宽度</param>
        ///// <param name="columnIndexs">列索引，从0开始</param>
        //public static void SetColumnWidth2(this ISheet sheet, int columnWidth, params int[] columnIndexs)
        //{
        //    foreach (var index in columnIndexs)
        //    {
        //        sheet.SetColumnWidth(index, (int)columnWidth * 256);
        //    }
        //}

        /// <summary>
        /// 通过列标识设置列宽（像素）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnWidth">列宽(像素）注：不能超过2000</param>
        /// <param name="columnLetters">列标识,多列用逗号分隔，如"A,B,C"等</param>
        public static void SetColumnWidthInPixel(this ISheet sheet, int columnWidth, string columnLetters)
        {
            var columnIndexs = columnLetters.Split(',').Select(x => CellReference.ConvertColStringToIndex(x)).ToArray();
            sheet.SetColumnWidthInPixel(columnWidth, columnIndexs);
        }


        /// <summary>
        /// 通过列索引设置列宽（像素）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnWidth">列宽(像素）注：不能超过2000</param>
        /// <param name="columnIndexs">列索引，从0开始</param>
        public static void SetColumnWidthInPixel(this ISheet sheet, int columnWidth, params int[] columnIndexs)
        {
            foreach (var index in columnIndexs)
            {
                sheet.SetColumnWidth(index, columnWidth * 256 / 8);
            }
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

        #region 设置行、单元格样式、边框样式、根据条件设置样式
        /// <summary>
        /// 设置行样式（通过设置行内每个单元格的样式实现）
        /// 区别于IRow.RowStyle = style,这种方法会为整行中空的单元格设置样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="cellStyle">非空单元格样式</param>
        /// <param name="rowStyle">整行样式，设置空的单元格样式</param>
        public static void SetRowStyle(this ISheet sheet, int rowIndex, ICellStyle cellStyle, ICellStyle rowStyle = null)
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
            if (setCellStyleAction == null)
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
            if (rowStart < 0 || columnStart < 0)
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
                    cell.CellStyle = cellStyle;
                }
            }
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="cellStyle">单元格样式</param>
        /// <returns>返回单元格</returns>
        public static ICell SetCellStyle(this ISheet sheet, int rowIndex, int columnIndex, ICellStyle cellStyle)
        {
            return sheet.GetOrCreateCell(rowIndex, columnIndex, cellStyle);
        }

        /// <summary>
        /// 设置单元格样式
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="setCellStyleAction">设置单元格样式委托方法(每个单元格都是独立的CellStyle)</param>
        public static void SetCellStyle(this ISheet sheet, int rowIndex, int columnIndex, Action<ICellStyle> setCellStyleAction)
        {
            sheet.SetCellsStyle(rowIndex, rowIndex, columnIndex, columnIndex, setCellStyleAction);
        }

        /// <summary>
        /// 根据条件设置单元格样式（通过循环sheet行来判断）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="rowStart">遍历的起始行索引</param> 
        /// <param name="conditionFunc">判断条件函数，入参为当前行索引</param>
        /// <param name="cellStyle">设置的单元格样式</param>
        /// /// <param name="columnIndexs">需要设置样式的单元格列索引,如果不指定列时,则设置一整行的样式</param>
        public static void SetCellStyleByCondition(this ISheet sheet, int rowStart, Predicate<int> conditionFunc, ICellStyle cellStyle, params int[] columnIndexs)
        {
            for (var rowIndex = rowStart; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                //var cell = sheet.GetOrCreateCell(rowIndex, columnIndex);
                //var cellValue = cell.GetCellValue();
                if (conditionFunc(rowIndex))
                {
                    if (columnIndexs == null || columnIndexs.Length < 1)
                    {
                        sheet.SetRowStyle(rowIndex, cellStyle);
                        continue;
                    }
                    foreach (var columnIndex in columnIndexs)
                    {
                        sheet.SetCellsStyle(rowIndex, rowIndex, columnIndex, columnIndex, cellStyle);
                    }
                }
            }
        }

        /// <summary>
        /// 根据条件置单元格样式（通过循环DataTable行来判断）
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dataTable">sheet对应的DataTable</param>
        /// <param name="bodyStartRowIndex">dataTable数据在Sheet中的起始行索引行</param>
        /// <param name="condition">判断条件函数，入参为DataTable中的DataRow</param>
        /// <param name="cellStyle">设置的单元格样式</param>
        /// <param name="columnIndexs">需要设置样式的单元格列索引,如果不指定列时,则设置一整行的样式</param>
        public static void SetCellStyleByCondition(this ISheet sheet, DataTable dataTable, int bodyStartRowIndex, Predicate<DataRow> condition, ICellStyle cellStyle, params int[] columnIndexs)
        {
            var dtRowIndexs = dataTable.AsEnumerable().Select((row, index) =>
            {
                return condition(row) ? index : -1;
            }).Where(i => i >= 0);
            foreach (var dtRowIndex in dtRowIndexs)
            {
                var rowIndex = bodyStartRowIndex + dtRowIndex;
                if (columnIndexs == null || columnIndexs.Length < 1)
                {
                    sheet.SetRowStyle(rowIndex, cellStyle);
                    continue;
                }
                foreach (var columnIndex in columnIndexs)
                {
                    sheet.SetCellsStyle(rowIndex, rowIndex, columnIndex, columnIndex, cellStyle);
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
        public static void SetCellsBorderStyle(this ISheet sheet, int rowStart, int rowEnd, int columnStart, int columnEnd = -1, BorderStyle borderStyle = BorderStyle.Thin)
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

        #region 设置行高、插入行

        /// <summary>
        /// 设置行高
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="height">行高，默认为300</param>
        /// <param name="startRowIndex">需要设置的起始行索引，默认为0</param>
        /// <param name="endRowIndex">需要设置的起始行索引，默认为-1；
        /// 1.如果小于0，则取sheet最后一行，
        /// 2.如果小于startRowIndex，则取startRowIndex</param> </param> 
        public static void SetRowHeight(this ISheet sheet, short height = 300, int startRowIndex = 0, int endRowIndex = -1)
        {
            if (startRowIndex < 0)
            {
                startRowIndex = 0;
            }
            var rowCount = sheet.LastRowNum;
            if (endRowIndex < 0)
            {
                endRowIndex = rowCount;
            }
            if (endRowIndex < startRowIndex)
            {
                endRowIndex = startRowIndex;
            }
            if (endRowIndex > rowCount)
            {
                endRowIndex = rowCount;
            }

            for (var i = startRowIndex; i <= endRowIndex; i++)
            {
                var row = sheet.GetRow(i);
                row.Height = height;
            }
        }

        /// <summary>
        /// 插入行
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex">插入位的起始位置</param>
        /// <param name="insertRowsCount">插入的行数</param>
        public static void InsertRows(this ISheet sheet, int startRowIndex, int insertRowsCount)
        {
            sheet.ShiftRows(startRowIndex, sheet.LastRowNum, insertRowsCount);
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
        public static void InsertSheetByDataTable(this ISheet sheet, DataTable sourceTable, int startRowIndex, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ICellStyle headerStyle = null)
        {
            var dateCellStyle = sheet.Workbook.CreateDateCellStyle(dateFormat);
            InsertSheetByDataTable(sheet, sourceTable, startRowIndex, isExportCaption, dateCellStyle, headerStyle);
        }

        /// <summary>
        /// 将DataTable插入到Sheet中
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceTable">源数据表</param>
        /// <param name="startRowIndex">起始行索引（从0开始）</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="headerStyle">标题行样式</param>
        /// <param name="setBodyCellStyle">设置单元格样式函数，第一个参数为sourceTable的行索引，第二个参数为sourceTable的列索引</param>
        public static void InsertSheetByDataTable(this ISheet sheet, DataTable sourceTable, int startRowIndex, bool isExportCaption, ICellStyle headerStyle = null, Func<int, int, ICellStyle> setBodyCellStyle = null)
        {
            InsertSheetByDataTable(sheet, sourceTable, startRowIndex, isExportCaption, null, headerStyle, setBodyCellStyle);
        }

        private static void InsertSheetByDataTable(ISheet sheet, DataTable sourceTable, int startRowIndex, bool isExportCaption, ICellStyle dateCellStyle = null, ICellStyle headerStyle = null, Func<int, int, ICellStyle> setBodyCellStyle = null)
        {
            var autoSizeRowIndex = startRowIndex;
            if (isExportCaption)
            {
                FillHeader(sheet, sourceTable, ref startRowIndex, headerStyle);
            }
            var cellStyle = sheet.Workbook.CreateCellStyleWithBorder().SetCellAlignmentStyle(false);
            if (dateCellStyle == null)
            {
                dateCellStyle = sheet.Workbook.CreateDateCellStyle("yyyy-MM-dd");
            }
            dateCellStyle.SetCellBorderStyle().SetCellAlignmentStyle(false);
            FillBodyData(sheet, sourceTable, startRowIndex, dateCellStyle, cellStyle, setBodyCellStyle);
            sheet.SetColumnAutoWidth(autoSizeRowIndex);
        }

        /// <summary>
        /// 填充表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceTable"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="headerStyle"></param>
        private static void FillHeader(ISheet sheet, DataTable sourceTable, ref int startRowIndex, ICellStyle headerStyle)
        {
            IRow firstRow = sheet.CreateRow(startRowIndex);
            if (headerStyle == null)
            {
                headerStyle = sheet.Workbook.CreateCellStyleOfHeader();
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

        /// <summary>
        /// 填充主体
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="sourceTable"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="dateCellStyle"></param>
        /// <param name="cellStyle"></param>
        /// <param name="setBodyCellStyle">设置单元格样式函数，第一个参数为sourceTable的行索引，第二个参数为sourceTable的列索引</param>
        private static void FillBodyData(ISheet sheet, DataTable sourceTable, int startRowIndex, ICellStyle dateCellStyle, ICellStyle cellStyle, Func<int, int, ICellStyle> setBodyCellStyle = null)
        {
            for (var rowIndex = 0; rowIndex < sourceTable.Rows.Count; rowIndex++)
            {
                var dr = sourceTable.Rows[rowIndex];
                var row = sheet.CreateRow(startRowIndex);
                for (var colIndex = 0; colIndex < sourceTable.Columns.Count; colIndex++)
                {
                    var column = sourceTable.Columns[colIndex];
                    ICellStyle style;
                    if (setBodyCellStyle != null)
                    {
                        style = setBodyCellStyle(rowIndex, colIndex);
                    }
                    else
                    {
                        var isDateType = column.DataType == typeof(DateTime);
                        style = isDateType ? dateCellStyle : cellStyle;
                    }
                    var cell = row.CreateCell(column.Ordinal);
                    cell.SetCellValue(dr[column], column.DataType, style);
                }
                startRowIndex++;
            }
        }
        #endregion

        #region 将Excel列索引转化列标识
        /// <summary>
        /// 将Excel列索引转化列标识，如0转化为"A"，1转化为"B"
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex">列索引（从0开始）</param>
        /// <returns>Excel列标识</returns>
        public static string ConvertToColumnLetter(this ISheet sheet, int columnIndex)
        {
            if (columnIndex < 0)
            {
                throw new ArgumentException("列索引不能小于0", nameof(columnIndex));
            }

            var columnLetter = new StringBuilder();
            while (columnIndex >= 0)
            {
                var modulo = columnIndex % 26;
                columnLetter.Insert(0, (char)(65 + modulo));
                columnIndex = (columnIndex - modulo) / 26 - 1;
            }
            return columnLetter.ToString();
        }
        #endregion

        #region  获取、插入图片
        /// <summary>
        /// 将图片插入到Sheet中
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="imageBytes">图片字节</param>
        /// <param name="rowStart">起始行索引</param>
        /// <param name="rowEnd"></param>
        /// <param name="columnStart"></param>
        /// <param name="columnEnd"></param>
        /// <param name="pictureType"></param>
        public static void InsertImage(this ISheet sheet, byte[] imageBytes, int rowStart, int rowEnd, int columnStart, int columnEnd, PictureType pictureType = PictureType.PNG)
        {
            //将图片添加到workbook中,返回图片所在workbook->Picture数组中的索引地址（从1开始）
            var pictIndex = sheet.Workbook.AddPicture(imageBytes, pictureType);
            //在sheet中创建画布
            var patriarch = sheet.CreateDrawingPatriarch();
            //设置锚点 （在起始单元格的X坐标0-1023，Y的坐标0-255，在终止单元格的X坐标0-1023，Y的坐标0-255，起始单元格行数，列数，终止单元格行数，列数）  
            var anchor = patriarch.CreateAnchor(0, 0, 0, 0, columnStart, rowStart, columnEnd, rowEnd);
            //创建图片 
            var pict = patriarch.CreatePicture(anchor, pictIndex);
        }

        /// <summary>
        /// 向单元格中插入图片
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="imageBytes">图片字节数组</param>
        /// <param name="rowIndex">行索引</param>
        /// <param name="columnIndex">列索引</param>
        /// <param name="pictureType">图片类型</param>
        public static void InsertImageToCell(this ISheet sheet, byte[] imageBytes, int rowIndex, int columnIndex, PictureType pictureType = PictureType.PNG)
        {
            sheet.InsertImage(imageBytes, rowIndex, rowIndex + 1, columnIndex, columnIndex + 1, pictureType);
        }

        /// <summary>
        /// 获取工作表中的所有图片
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <returns>图片列表,
        /// PictureData为图片数据，其中Data为图片字节数组，PictureType为图片类型
        /// ClientAnchor为图片锚点，其中Row1、Row2、Col1、Col2为图片锚点在单元格中的位置</returns>
        public static IEnumerable<IPicture> GetAllPictures(this ISheet sheet)
        {
            var patriarch = sheet.DrawingPatriarch;
            if (patriarch == null)
                yield break;

            if (patriarch is NPOI.HSSF.UserModel.HSSFPatriarch hssfPatriarch)
            {
                foreach (var shape in hssfPatriarch.Children)
                {
                    if (shape is IPicture picture)
                    {
                        yield return picture;
                    }
                }
            }
            else if (patriarch is NPOI.XSSF.UserModel.XSSFDrawing xssfDrawing)
            {
                foreach (var shape in xssfDrawing.GetShapes())
                {
                    if (shape is IPicture picture)
                    {
                        yield return picture;
                    }
                }
            }
        }

        /// <summary>
        /// 获取指定单元格中的第一张图片
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="rowIndex">行索引（从0开始）</param>
        /// <param name="columnIndex">列索引</param>
        /// <returns>单元格中的第一张图片数据，如果没有图片则返回null</returns>
        public static IPictureData GetCellPicture(this ISheet sheet, int rowIndex, int columnIndex)
        {
            var pictures = sheet.GetAllPictures();
            foreach (var picture in pictures)
            {
                var anchor = picture.ClientAnchor;
                if (anchor.Row1 <= rowIndex && anchor.Row2 >= rowIndex
                    && anchor.Col1 <= columnIndex && anchor.Col2 >= columnIndex)
                {
                    return picture.PictureData;
                }
            }
            return null;
        }
        #endregion

        #region 设置Sheet行列冻结、页签颜色

        /// <summary>
        /// 冻结行和列(默认冻结首行)
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="colSplit">要冻结的列数(从1开始)</param>
        /// <param name="rowSplit">要冻结的行数(从1开始)</param>
        /// <param name="leftmostColumn">右边区域可见的首列序号(从1开始)</param>
        /// <param name="topRow">下方区域可见的首行序号(从1开始)</param>
        public static void CreateFreezePane(this ISheet sheet, int colSplit = 0, int rowSplit = 1, int leftmostColumn = -1, int topRow = -1)
        {
            if (leftmostColumn < 0 || topRow < 0)
            {
                sheet.CreateFreezePane(colSplit, rowSplit);
            }
            else
            {
                sheet.CreateFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
            }
        }

        /// <summary>
        /// 设置页签颜色
        /// </summary>
        /// <param name="sheet">工作表</param>
        /// <param name="color">颜色</param>
        public static void SetSheetColor(this ISheet sheet, short color)
        {
            try
            {
                if (sheet is NPOI.XSSF.UserModel.XSSFSheet xssfSheet)
                {
                    // .xlsx 格式
                    xssfSheet.SetTabColor(color);
                }
                else if (sheet is NPOI.HSSF.UserModel.HSSFSheet hssfSheet)
                {
                    // .xls 格式
                    hssfSheet.TabColorIndex = color;
                }
            }
            catch (NotImplementedException)
            {
                // 如果特定格式不支持设置颜色，则静默失败
            }
        }
        #endregion
    }
}
