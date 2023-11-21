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
        #region 创建IFont
        /// <summary>
        /// 创建IFont（默认微软雅黑）
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
        /// 创建IFont
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fontSize">字体大小</param>
        /// <param name="bold">字体是否加粗（默认不加粗）</param>
        /// <param name="fontColor">字体颜色（默认黑色）</param>
        /// <param name="fontName">字体名称（默认微软雅黑）</param>
        /// <returns></returns>
        public static IFont CreateFont(this IWorkbook workbook, short fontSize, bool bold = false, short? fontColor = null, string fontName = "微软雅黑")
        {
            var font = workbook.CreateFont(fontName);
            font.FontHeightInPoints = fontSize;
            font.Boldweight = bold ? (short)FontBoldWeight.Bold : (short)FontBoldWeight.Normal;
            font.Color = fontColor ?? IndexedColors.Black.Index;
            return font;
        }
        #endregion
         
        #region 创建带边框单元格样式、居中带边框单元样式、表头样式、日期单元格样式、克隆样式
        /// <summary>
        /// 创建带边框单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        public static ICellStyle CreateCellStyleWithBorder(this IWorkbook workbook)
        {
            return workbook.CreateCellStyle().SetCellBorderStyle();
        }

        /// <summary>
        /// 创建居中显示、带边框的单元格样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="wrapText">是否自动换行，默认为不自动换行</param>
        /// <returns></returns>
        public static ICellStyle CreateCellStyleCenterWithBorder(this IWorkbook workbook, bool wrapText = false)
        {
            return workbook.CreateCellStyleWithBorder().SetCellAlignmentStyle(wrapText);
        }

        /// <summary>
        /// 创建标题行单元格样式（内容默认居中）
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="fontSize">字体大小：默认10</param>
        /// <param name="fontColor">字体颜色：默认黑色</param>
        /// <param name="bold">是否加粗</param>
        /// <param name="backgroundColor">背景色</param>
        /// <returns></returns>
        public static ICellStyle CreateCellStyleOfHeader(this IWorkbook workbook, short fontSize = 10, short? fontColor = null, bool bold = true, short? backgroundColor = 42)
        {
            ICellStyle style = workbook.CreateCellStyleCenterWithBorder(true);
            IFont font = workbook.CreateFont("");
            style.SetCellFontStyle(font, fontSize, bold, fontColor ?? IndexedColors.Black.Index);
            if (backgroundColor != null)
            {
                style.SetCellBackgroundStyle((short)backgroundColor);
            }
            return style;
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
        /// 克隆样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <returns></returns>
        public static ICellStyle CloneStyleFrom(this IWorkbook workbook, ICellStyle style)
        {
            var cloneStyle = workbook.CreateCellStyle();
            cloneStyle.CloneStyleFrom(style);
            return cloneStyle;
        }

        /// <summary>
        /// 合并样式，返回新的样式对象(不影响style和style2的样式）
        /// 注：只合并Style2中非默认的背景色样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="style"></param>
        /// <param name="style2"></param>
        /// <returns></returns>
        public static ICellStyle MergeStyle(this IWorkbook workbook, ICellStyle style, ICellStyle style2)
        {
            var cloneStyle = workbook.CreateCellStyle();
            cloneStyle.CloneStyleFrom(style);
            //// 循环遍历 ICellStyle 的属性
            //foreach (var property in typeof(ICellStyle).GetProperties())
            //{
            //    if(property.PropertyType == typeof(short))
            //    {
            //    } 
            //    var value = property.GetValue(style2);
            //    if(value != null)
            //    {
            //        property.SetValue(cloneStyle, value);
            //    }
            //    Console.WriteLine($"属性名：{property.Name}，属性值：{value}");
            //} 
            if (style2.FillForegroundColor != 64)
            {
                cloneStyle.FillForegroundColor = style2.FillForegroundColor;
            }
            if (style2.FillPattern != FillPattern.NoFill)
            {
                cloneStyle.FillPattern = style2.FillPattern;
            }
            return cloneStyle;
        }
        #endregion

        #region 将Color转化为Npoi颜色数值（只针对HSSFWorkbook有效）
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
        #endregion

        #region 将workbook转化为MemoryStream、保存成Excel
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
        /// 将workbook保存至Excel中
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="filePath">excel文件路径</param>
        /// <returns></returns>
        public static void SaveToExcel(this IWorkbook workbook, string filePath)
        {
            using (MemoryStream ms = workbook.ToMemoryStream())
            {
                File.WriteAllBytes(filePath, ms.ToArray());
            }
        }
        #endregion

        #region 设置导出excel属性信息
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
        #endregion

        #region 将DataTable转化为Sheet添加至Workbook中
        /// <summary>
        /// 将DataTable转化为Sheet添加至Workbook中
        /// </summary>
        /// <param name="workbook">目标Workbook</param>
        /// <param name="sourceTable">源数据表,如果TableName不为空，则将TableName设置为sheet的名称</param>
        /// <param name="startRow">导出到Excel中的起始行</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="headerStyle">标题行样式</param>
        /// <returns></returns>
        public static ISheet AddSheetByDataTable(this IWorkbook workbook, DataTable sourceTable, int startRow = 0, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ICellStyle headerStyle = null)
        {
            var sheet = CreateSheet(workbook, sourceTable);
            sheet.InsertSheetByDataTable(sourceTable, startRow, dateFormat, isExportCaption, headerStyle);
            return sheet;
        }

        /// <summary>
        /// 将DataTable转化为Sheet添加至Workbook中
        /// </summary>
        /// <param name="workbook">目标Workbook</param>
        /// <param name="sourceTable">源数据表,如果TableName不为空，则将TableName设置为sheet的名称</param>
        /// <param name="startRow">导出到Excel中的起始行</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="headerStyle">标题行样式</param>
        /// <param name="setBodyCellStyle">设置单元格样式函数，第一个参数为sourceTable的行索引，第二个参数为sourceTable的列索引</param>
        public static ISheet InsertSheetByDataTable(this IWorkbook workbook, DataTable sourceTable, int startRow=0, bool isExportCaption = true, ICellStyle headerStyle = null, Func<int, int, ICellStyle> setBodyCellStyle = null)
        {
            var sheet = CreateSheet(workbook, sourceTable);
            sheet.InsertSheetByDataTable(sourceTable, startRow, isExportCaption, headerStyle, setBodyCellStyle);
            return sheet;
        }

        private static ISheet CreateSheet(IWorkbook workbook, DataTable sourceTable)
        {
            string sheetName = sourceTable.TableName;
            if (string.IsNullOrEmpty(sheetName))
            {
                sheetName = $"Sheet{workbook.NumberOfSheets + 1}";
            }
            if (workbook.GetSheet(sheetName) != null)
            {
                sheetName = Guid.NewGuid().ToString();
            }
            var sheet = workbook.CreateSheet(sheetName);
            return sheet;
        }
        #endregion
    }
}
