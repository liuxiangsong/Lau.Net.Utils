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

namespace Lau.Net.Utils.Excel
{
    public static class NpoiStaticUtil
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
                    dataSet.Tables.Add(dt);
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
        public static DataTable ExcelToDataTable(string excelPath, int sheetIndex, int headerIndex = 0)
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
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="type">生成Excel的类型</param>
        public static void DataTableToExcel(string filePath, DataTable sourceTable, bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            using (MemoryStream ms = DataTableToStream(sourceTable, isExportCaption, type))
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
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="type">生成Excel的类型</param>
        public static void DataSetToExcel(string filePath, DataSet sourceSet, bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            using (MemoryStream ms = DataSetToStream(sourceSet, isExportCaption, type))
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
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="type">生成Excel的类型</param>
        /// <returns>MemoryStream</returns>
        public static MemoryStream DataTableToStream(DataTable sourceTable, bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            IWorkbook workbook = CreateWorkbook(type);
            workbook.AddSheetByDataTable(sourceTable, isExportCaption);
            return workbook.ToMemoryStream();
        }

        public static IWorkbook DataSetToWorkBook(DataSet sourceSet, bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            IWorkbook workbook = CreateWorkbook(type);
            foreach (DataTable dt in sourceSet.Tables)
            {
                workbook.AddSheetByDataTable(dt, isExportCaption);
            }
            return workbook;
        }


        /// <summary>
        /// DataSetToStream
        /// </summary>
        /// <param name="sourceSet">源数据集</param>
        /// <param name="isExportCaption">是否导出表的标题</param>
        /// <param name="type">生成Excel的类型</param>
        /// <returns>MemoryStream</returns>
        public static MemoryStream DataSetToStream(DataSet sourceSet, bool isExportCaption = true, ExcelType type = ExcelType.Xlsx)
        {
            IWorkbook workbook = DataSetToWorkBook(sourceSet, isExportCaption, type);
            return workbook.ToMemoryStream();
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

            for (int i = 0; i < columnCount; i++)
            {
                string columnName = string.Empty;
                if (autoColumnHeader)
                {
                    columnName = "A" + i;
                }
                else
                {
                    ICell cell = headerRow.GetCell(i);
                    columnName = (cell == null) ? "A1" : cell.ToString();
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
                        dr[i] = cell.ToString();
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
