using Lau.Net.Utils.Excel.NpoiExtensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Lau.Net.Utils.Excel
{
    /// <summary>
    /// 
    /// </summary>
    public static class NpoiUtil
    {
        /// <summary>
        /// Excel 版本
        /// </summary>
        public enum ExcelType
        {
            /// <summary>
            /// Excel2003
            /// </summary>
            Xls,
            /// <summary>
            /// Excel2007、Excel2010
            /// </summary>
            Xlsx
        }

        #region 创建workbook
        /// <summary>
        /// 创建workbook
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(ExcelType type = ExcelType.Xlsx)
        {
            IWorkbook workbook;
            if (type == ExcelType.Xlsx)
            {
                workbook = new XSSFWorkbook();
            }
            else
            {
                workbook = new HSSFWorkbook();
            }
            return workbook;
        }

        /// <summary>
        /// 创建workbook
        /// </summary>
        /// <param name="sourceTable">源数据表,将表转化为sheet,sheet名称为表的名称</param>
        /// <param name="startRowIndex">导出起始行索引</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="isExportCaption">是否导出表的列头</param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static IWorkbook CreateWorkbook(DataTable sourceTable, int startRowIndex = 0, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            var workbook = CreateWorkbook(type);
            if (sourceTable != null)
            {
                workbook.AddSheetByDataTable(sourceTable, startRowIndex, dateFormat, isExportCaption);
            }
            return workbook;
        }
        #endregion

        #region Excel导入至DataSet、DataTable
        /// <summary>
        /// ExcelToDataSet
        /// </summary>
        /// <param name="excelPath">Excel的路径</param>
        /// <param name="headerIndex">作为表头的行,表头行以下行为表数据（下标从0开始）</param>
        /// <returns>返回DataSet</returns>
        public static DataSet ExcelToDataSet(string excelPath, int headerIndex = 0)
        {
            DataSet dataSet = new DataSet();
            IWorkbook workbook = ExcelToWorkbook(excelPath);
            foreach (ISheet sheet in workbook)
            {
                DataTable dt = SheetToDataTable(sheet, headerIndex);
                if (dt != null)
                {
                    dataSet.Tables.Add(dt);
                }
                var sheetName = sheet.SheetName;
                if (!string.IsNullOrWhiteSpace(sheetName) && !dataSet.Tables.Contains(sheetName))
                {
                    dt.TableName = sheetName;
                }
            }
            return dataSet;
        }

        /// <summary>
        /// ExcelToDataTable
        /// </summary>
        /// <param name="excelPath">Excel的路径</param>
        /// <param name="sheetIndex">第几张表单（下标从0开始）</param>
        /// <param name="headerIndex">作为表头的行,表头行以下行为表数据（下标从0开始）</param>
        /// <returns>返回DataTable</returns>
        public static DataTable ExcelToDataTable(string excelPath, int sheetIndex = 0, int headerIndex = 0)
        {
            ISheet sheet = ExcelToWorkbook(excelPath).GetSheetAt(sheetIndex);
            return SheetToDataTable(sheet, headerIndex);
        }

        /// <summary>
        /// ExcelToDataTable
        /// </summary>
        /// <param name="excelPath">Excel的路径</param>
        /// <param name="sheetName">表单的名称</param>
        /// <param name="headerIndex">作为表头的行,表头行以下行为表数据（下标从0开始）</param>
        /// <returns>返回DataTable</returns>
        public static DataTable ExcelToDataTable(string excelPath, string sheetName, int headerIndex = 0)
        {
            ISheet sheet = ExcelToWorkbook(excelPath).GetSheet(sheetName);
            return SheetToDataTable(sheet, headerIndex);
        }
        #endregion

        #region 将DataSet、DataTable导入至Excel
        /// <summary>
        /// 将DataTable导入Excel中
        /// </summary>
        /// <param name="filePath">保存Excel的文件路径</param>
        /// <param name="sourceTable">源数据表</param>
        /// <param name="startRowIndex">起始行索引（从0开始）</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="type">生成Excel的类型</param>
        public static void DataTableToExcel(string filePath, DataTable sourceTable, int startRowIndex = 0, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            using (MemoryStream ms = DataTableToStream(sourceTable, startRowIndex, dateFormat, isExportCaption, type))
            {
                File.WriteAllBytes(filePath, ms.ToArray());
                //using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                //{
                //    byte[] buffer = ms.ToArray();
                //    fs.Write(buffer, 0, buffer.Length);
                //    fs.Flush();                    
                //}
            }
        }

        /// <summary>
        /// 将DataSet导入Excel中
        /// </summary>
        /// <param name="filePath">保存Excel的文件路径</param>
        /// <param name="sourceSet">源数据集</param>
        /// <param name="startRowIndex">起始行索引（从0开始）</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="type">生成Excel的类型</param>
        public static void DataSetToExcel(string filePath, DataSet sourceSet, int startRowIndex = 0, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            using (MemoryStream ms = DataSetToStream(sourceSet, startRowIndex, dateFormat, isExportCaption, type))
            {
                File.WriteAllBytes(filePath, ms.ToArray());
            }
        }

        #endregion

        #region 将DataTable、DataSet转化为MemoryStream
        /// <summary>
        /// DataTableToStream
        /// </summary>
        /// <param name="sourceTable">源数据表</param>
        /// <param name="startRowIndex">起始行索引（从0开始）</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="type">生成Excel的类型</param>
        /// <returns>MemoryStream</returns>
        public static MemoryStream DataTableToStream(DataTable sourceTable, int startRowIndex = 0, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            IWorkbook workbook = CreateWorkbook(type);
            workbook.AddSheetByDataTable(sourceTable, startRowIndex, dateFormat, isExportCaption);
            return workbook.ToMemoryStream();
        }

        /// <summary>
        /// DataSetToStream
        /// </summary>
        /// <param name="sourceSet">源数据集</param>
        /// <param name="startRowIndex">起始行索引（从0开始）</param>
        /// <param name="dateFormat">日期格式</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="type">生成Excel的类型</param>
        /// <returns>MemoryStream</returns>
        public static MemoryStream DataSetToStream(DataSet sourceSet, int startRowIndex = 0, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            IWorkbook workbook = DataSetToWorkBook(sourceSet, startRowIndex, dateFormat, isExportCaption, type);
            return workbook.ToMemoryStream();
        }

        #endregion

        #region 将DataTable、DataSet转化为IWorkbook
        public static IWorkbook DataTableToWorkBook(DataTable sourceTable, int startRowIndex = 0, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            IWorkbook workbook = CreateWorkbook(sourceTable, startRowIndex, dateFormat, isExportCaption, type);
            return workbook;
        }

        public static IWorkbook DataSetToWorkBook(DataSet sourceSet, int startRowIndex = 0, string dateFormat = "yyyy-MM-dd", bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            IWorkbook workbook = CreateWorkbook(type);
            foreach (DataTable dt in sourceSet.Tables)
            {
                workbook.AddSheetByDataTable(dt, startRowIndex, dateFormat, isExportCaption);
            }
            return workbook;
        }
        #endregion

        #region 通过HttpContext导出excel
        /// <summary>
        /// 通过HttpContext导出excel
        /// </summary>
        /// <param name="httpContext"></param>
        /// <param name="fileName">excel文件名</param>
        /// <param name="workbook"></param>
        public static void ExportByHttpContext(HttpContext httpContext, string fileName, IWorkbook workbook)
        {
            // 设置编码和附件格式
            httpContext.Response.ContentType = "application/ms-excel";
            httpContext.Response.ContentEncoding = Encoding.UTF8;
            httpContext.Response.Charset = "";
            //------------------
            //这里判断使用的浏览器是否为Firefox，Firefox导出文件时不需要对文件名显示编码，编码后文件名会乱码
            //但是IE和Google需要编码才能保持文件名正常
            var disposition = "attachment;filename=" + fileName;
            if (httpContext.Request.ServerVariables["http_user_agent"].ToString().IndexOf("Firefox") == -1)
            {
                disposition = "attachment;filename=" + HttpUtility.UrlEncode(fileName, Encoding.UTF8);
            }
            httpContext.Response.AddHeader("Content-Disposition", disposition);

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                ms.Flush();
                //ms.Position = 0;
                httpContext.Response.BinaryWrite(ms.GetBuffer());
                httpContext.Response.End();
            }
        }
        #endregion

        #region 获取设置单元格样式函数
        /// <summary>
        /// 获取设置单元格样式函数
        /// </summary>
        /// <param name="colIndexs">列索引集合</param>
        /// <param name="containsCellStyle">包含在colIndexs中的列单元格样式</param>
        /// <param name="notContainsStyle">不包含在colIndexs中的列单元格样式</param>
        /// <returns></returns>
        public static Func<int, int, ICellStyle> GetCellStyleFunc(IEnumerable<int> colIndexs, ICellStyle containsCellStyle, ICellStyle notContainsStyle)
        {
            return (rowIndex, columnIndex) =>
            {
                if (colIndexs.Contains(columnIndex))
                {
                    return containsCellStyle;
                };
                return notContainsStyle;
            };
        }
        #endregion

        #region 私有方法
        /// <summary>
        /// SheetToDataTable
        /// </summary>
        /// <param name="sheet">表单</param>
        /// <param name="headerIndex">作为表头的行(下标从0开始）</param>
        /// <returns></returns>
        private static DataTable SheetToDataTable(ISheet sheet, int headerIndex)
        {
            if (sheet.LastRowNum < headerIndex || sheet.PhysicalNumberOfRows < 1) return null;
            bool autoColumnHeader = (headerIndex < 0);
            DataTable dt = new DataTable();
            IRow headerRow = autoColumnHeader ? sheet.GetRow(0) : sheet.GetRow(headerIndex);
            int columnCount = headerRow.LastCellNum;

            var formulaEvaluator = WorkbookFactory.CreateFormulaEvaluator(sheet.Workbook);
            for (int i = 0; i < columnCount; i++)
            {
                string columnName;
                if (autoColumnHeader)
                {
                    columnName = "A" + i;
                }
                else
                {
                    ICell cell = headerRow.GetCell(i);
                    columnName = (cell == null) ? "A1" : cell.GetCellValue(formulaEvaluator);
                }

                int j = 2;
                while (dt.Columns.Contains(columnName))
                {
                    columnName = "A" + j.ToString();
                    j++;
                }
                dt.Columns.Add(columnName);
            }

            IEnumerator rowEnumerator = sheet.GetRowEnumerator();
            while (headerIndex >= 0)
            {
                rowEnumerator.MoveNext();
                headerIndex--;
            }
            while (rowEnumerator.MoveNext())
            {
                DataRow dr = dt.NewRow();
                IRow currentRow = (IRow)rowEnumerator.Current;
                for (int i = 0; i < columnCount; i++)
                {
                    ICell cell = currentRow.GetCell(i);
                    if (cell == null)
                    {
                        dr[i] = null;
                    }
                    else
                    {
                        dr[i] = cell.GetCellValue(formulaEvaluator);
                    }
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        /// <summary>
        /// 将Excel转化成Workbook
        /// </summary>
        /// <param name="excelPath">Excel的路径</param>
        /// <returns>Workbook</returns>
        private static IWorkbook ExcelToWorkbook(string excelPath)
        {
            if (!File.Exists(excelPath))
                throw new Exception("指定路径的Excel文件不存在！");
            using (FileStream fs = File.OpenRead(excelPath))
            {
                return WorkbookFactory.Create(fs);
            }
        }


        ///// <summary>
        ///// 取得Excel的所有工作表名
        ///// </summary>
        ///// <param name="excelPath">Excel文件绝对路径</param>
        ///// <returns></returns>
        //public static List<string> GetExcelTablesName(string excelPath)
        //{
        //    return ExcelToWorkbook(excelPath).Select(a => a.SheetName).ToList();
        //}

        //switch(cell.CellType)
        //{
        //    case HSSFCellType.BLANK:
        //        dr[i] = "[null]";
        //        break;
        //    case HSSFCellType.BOOLEAN:
        //        dr[i] = cell.BooleanCellValue;
        //        break;
        //    case HSSFCellType.NUMERIC:
        //        dr[i] = cell.ToString();    //This is a trick to get the correct value of the cell. NumericCellValue will return a numeric value no matter the cell value is a date or a number.
        //        break;
        //    case HSSFCellType.STRING:
        //        dr[i] = cell.StringCellValue;
        //        break;
        //    case HSSFCellType.ERROR:
        //        dr[i] = cell.ErrorCellValue;
        //        break;
        //    case HSSFCellType.FORMULA:
        //    default:
        //        dr[i] = "="+cell.CellFormula;
        //        break;
        //}


        #endregion
    }
}
