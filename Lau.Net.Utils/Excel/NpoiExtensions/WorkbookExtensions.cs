using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
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
        public static IFont CreateFont(this IWorkbook workbook, string fontName)
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

        /// <summary>
        /// 将workbook保存至Excel中
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="filePath">excel文件路径</param>
        /// <returns></returns>
        public static void SaveToExcel(this IWorkbook workbook,string filePath)
        {
            using (MemoryStream ms = workbook.ToMemoryStream())
            {
                File.WriteAllBytes(filePath, ms.ToArray());
            }
        }

        /// <summary>
        /// 设置导出excel属性信息
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="company"></param>
        /// <param name="author"></param>
        /// <param name="applicationName"></param>
        /// <param name="comments"></param>
        /// <param name="title"></param>
        /// <param name="subject"></param>
        public static void SetDocumentSummaryInfo(this HSSFWorkbook workbook, string company, string author, string applicationName, string comments, string title, string subject)
        {
            var dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = company;
            workbook.DocumentSummaryInformation = dsi;
            var si = PropertySetFactory.CreateSummaryInformation();
            si.Author = author; //填加xls文件作者信息
            si.ApplicationName = applicationName; //填加xls文件创建程序信息
            si.LastAuthor = author; //填加xls文件最后保存者信息
            si.Comments = comments; //填加xls文件作者信息
            si.Title = title; //填加xls文件标题信息
            si.Subject = subject;//填加文件主题信息
            si.CreateDateTime = System.DateTime.Now;
            workbook.SummaryInformation = si;
        }

        /// <summary>
        /// 将DataTable添加至Workbook中
        /// </summary>
        /// <param name="workbook">目标Workbook</param>
        /// <param name="sourceTable">源数据表</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="startRow">导出到Excel中的起始行</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="headerStyle">标题行样式</param>
        /// <returns></returns>
        public static ISheet DataTableToWorkbook(this IWorkbook workbook, DataTable sourceTable, bool isExportCaption = true, int startRow = 0, string dateFormat = "yyyy-MM-dd", ICellStyle headerStyle = null)
        {
            string sheetName = "Sheet1";
            if (!string.IsNullOrEmpty(sourceTable.TableName))
            {
                sheetName = sourceTable.TableName;
            }
            ISheet sheet = workbook.CreateSheet(sheetName);
            ICellStyle dateCellStyle = workbook.CreateDateCellStyle(dateFormat);
            if (headerStyle == null)
            {
                headerStyle = workbook.CreateHeaderStyle();
            }
            sheet.DataTableToSheet(sourceTable, startRow, isExportCaption, dateCellStyle,headerStyle);
            //for (int i = 0; i < startRow; i++)
            //{
            //    sheet.CreateRow(i);
            //}
            //if (isExportCaption)
            //{
            //    IRow firstRow = sheet.CreateRow(startRow);
            //    if (headerStyle == null)
            //    {
            //        headerStyle = workbook.CreateHeaderStyle();
            //    }
            //    foreach (DataColumn column in sourceTable.Columns)
            //    {
            //        ICell cell = firstRow.CreateCell(column.Ordinal);
            //        cell.SetCellValue(column.Caption);
            //        cell.CellStyle = headerStyle;
            //    }
            //    firstRow.RowStyle = headerStyle;
            //}

            //ICellStyle dateCellStyle = workbook.CreateDateCellStyle(dateFormat);
            //int rowNum = isExportCaption ? startRow + 1 : startRow;
            //sheet.DataTableToSheet(sourceTable, rowNum, dateCellStyle);
            return sheet;
        }
    }
}
