//using System;
//using System.IO;
//using System.Text;
//using System.Data;
//using System.Data.OleDb;
//using System.Reflection;
//using System.Collections.Generic;
//using System.Runtime.InteropServices;

//namespace Lau.Net.Utils
//{
//    /// <summary>
//    /// Excel操作辅助类（无需VBA引用）
//    /// </summary>
//    public class ExcelOleDbUtil
//    {
//        /// <summary>
//        /// Excel 版本
//        /// </summary>
//        public enum ExcelType
//        {
//            Excel2003, Excel2010
//        }

//        /// <summary>
//        /// IMEX 三种模式。
//        /// IMEX是用来告诉驱动程序使用Excel文件的模式，其值有0、1、2三种，分别代表导出、导入、混合模式。
//        /// </summary>
//        public enum IMEXType
//        {
//            ExportMode = 0, ImportMode = 1, LinkedMode = 2
//        }

//        #region 获取Excel连接字符串

//        /// <summary>
//        /// 返回Excel 连接字符串   [IMEX=1]
//        /// </summary>
//        /// <param name="excelPath">Excel文件 绝对路径</param>
//        /// <param name="header">是否把第一行作为列名</param>
//        /// <param name="eType">Excel 版本 </param>
//        /// <returns></returns>
//        public static string GetExcelConnectstring(string excelPath, bool header, ExcelType eType)
//        {
//            return GetExcelConnectstring(excelPath, header, eType, IMEXType.ImportMode);
//        }

//        /// <summary>
//        /// 返回Excel 连接字符串
//        /// </summary>
//        /// <param name="excelPath">Excel文件 绝对路径</param>
//        /// <param name="header">是否把第一行作为列名</param>
//        /// <param name="eType">Excel 版本 </param>
//        /// <param name="imex">IMEX模式</param>
//        /// <returns></returns>
//        public static string GetExcelConnectstring(string excelPath, bool header, ExcelType eType, IMEXType imex)
//        {
//            if (!File.Exists(excelPath))
//                throw new FileNotFoundException("Excel路径不存在!");

//            string connectstring = string.Empty;

//            string hdr = "NO";
//            if (header) hdr = "YES";

//            if (eType == ExcelType.Excel2003)
//                connectstring = "Provider=Microsoft.Jet.OleDb.4.0; data source=" + excelPath + ";Extended Properties='Excel 8.0; HDR=" + hdr + "; IMEX=" + imex.GetHashCode() + "'";
//            else
//                connectstring = "Provider=Microsoft.ACE.OLEDB.12.0; data source=" + excelPath + ";Extended Properties='Excel 12.0 Xml; HDR=" + hdr + "; IMEX=" + imex.GetHashCode() + "'";

//            return connectstring;
//        }

//        #endregion

//        #region 获取Excel工作表名

//        /// <summary>
//        /// 返回Excel工作表名
//        /// </summary>
//        /// <param name="excelPath">Excel文件绝对路径</param>
//        /// <param name="eType">Excel 版本 </param>
//        /// <returns></returns>
//        public static List<string> GetExcelTablesName(string excelPath, ExcelType eType)
//        {
//            string connectstring = GetExcelConnectstring(excelPath, true, eType);
//            return GetExcelTablesName(connectstring);
//        }

//        /// <summary>
//        /// 返回Excel工作表名
//        /// </summary>
//        /// <param name="connectstring">excel连接字符串</param>
//        /// <returns></returns>
//        public static List<string> GetExcelTablesName(string connectstring)
//        {
//            using (OleDbConnection conn = new OleDbConnection(connectstring))
//            {
//                return GetExcelTablesName(conn);
//            }
//        }

//        /// <summary>
//        /// 返回Excel工作表名
//        /// </summary>
//        /// <param name="connection">excel连接</param>
//        /// <returns></returns>
//        public static List<string> GetExcelTablesName(OleDbConnection connection)
//        {
//            List<string> list = new List<string>();

//            if (connection.State == ConnectionState.Closed)
//                connection.Open();

//            DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
//            if (dt != null && dt.Rows.Count > 0)
//            {
//                for (int i = 0; i < dt.Rows.Count; i++)
//                {
//                    list.Add(TypeUtil.ToString(dt.Rows[i][2]));
//                }
//            }

//            return list;
//        }

//        /// <summary>
//        /// 返回Excel第一个工作表表名
//        /// </summary>
//        /// <param name="excelPath">Excel文件 绝对路径</param>
//        /// <param name="eType">Excel 版本 </param>
//        /// <returns></returns>
//        public static string GetExcelFirstTableName(string excelPath, ExcelType eType)
//        {
//            string connectstring = GetExcelConnectstring(excelPath, true, eType);
//            return GetExcelFirstTableName(connectstring);
//        }

//        /// <summary>
//        /// 返回Excel第一个工作表表名
//        /// </summary>
//        /// <param name="connectstring">excel连接字符串</param>
//        /// <returns></returns>
//        public static string GetExcelFirstTableName(string connectstring)
//        {
//            using (OleDbConnection conn = new OleDbConnection(connectstring))
//            {
//                return GetExcelFirstTableName(conn);
//            }
//        }

//        /// <summary>
//        /// 返回Excel第一个工作表表名
//        /// </summary>
//        /// <param name="connection">excel连接</param>
//        /// <returns></returns>
//        public static string GetExcelFirstTableName(OleDbConnection connection)
//        {
//            string tableName = string.Empty;

//            if (connection.State == ConnectionState.Closed)
//                connection.Open();

//            DataTable dt = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
//            if (dt != null && dt.Rows.Count > 0)
//            {
//                tableName = TypeUtil.ToString(dt.Rows[0][2]);
//            }

//            return tableName;
//        }
//        #endregion

//        #region 取得工作表的列名

//        /// <summary>
//        /// 获取Excel文件中指定工作表的列
//        /// </summary>
//        /// <param name="excelPath">Excel文件 绝对路径</param>
//        /// <param name="sheetName">名称 excel table  例如：Sheet1$</param>
//        /// <returns></returns>
//        public static List<string> GetColumnsList(string excelPath, ExcelType eType, string sheetName)
//        {
//            List<string> list = new List<string>();
//            DataTable sheetColumns = null;
//            string connectstring = GetExcelConnectstring(excelPath, true, eType);
//            using (OleDbConnection conn = new OleDbConnection(connectstring))
//            {
//                conn.Open();
//                sheetColumns = GetSheetSchema(sheetName, conn);
//            }
//            foreach (DataRow dr in sheetColumns.Rows)
//            {
//                string columnName = dr["ColumnName"].ToString();
//                string datatype = ((OleDbType)dr["ProviderType"]).ToString();//对应数据库类型
//                string netType = dr["DataType"].ToString();//对应的.NET类型，如System.String
//                list.Add(columnName);
//            }

//            return list;
//        }
//        #endregion

//        #region EXCEL导入DataSet

//        /// <summary>
//        /// EXCEL导入DataSet
//        /// </summary>
//        /// <param name="excelPath">Excel文件 绝对路径</param>
//        /// <param name="sheetName">名称 excel table  例如：Sheet1$ </param>
//        /// <param name="header">是否把第一行作为列名</param>
//        /// <param name="eType">Excel 版本 </param>
//        /// <returns>返回Excel相应工作表中的数据 DataSet   [table不存在时返回空的DataSet]</returns>
//        public static DataSet ExcelToDataSet(string excelPath, string sheetName, bool header, ExcelType eType)
//        {
//            string connectstring = GetExcelConnectstring(excelPath, header, eType);
//            return ExcelToDataSet(connectstring, sheetName);
//        }

//        /// <summary>
//        /// EXCEL导入DataSet
//        /// </summary>
//        /// <param name="connectstring">excel连接字符串</param>
//        /// <param name="sheetName">名称 excel table  例如：Sheet1$ </param>
//        /// <returns>返回Excel相应工作表中的数据 DataSet   [table不存在时返回空的DataSet]</returns>
//        public static DataSet ExcelToDataSet(string connectstring, string sheetName)
//        {
//            using (OleDbConnection conn = new OleDbConnection(connectstring))
//            {
//                DataSet ds = new DataSet();

//                //判断该工作表在Excel中是否存在
//                if (IsExistSheetName(conn, sheetName))
//                {
//                    OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", conn);
//                    adapter.Fill(ds, sheetName);
//                }

//                return ds;
//            }
//        }

//        /// <summary>
//        /// EXCEL所有工作表导入DataSet
//        /// </summary>
//        /// <param name="excelPath">Excel文件 绝对路径</param>
//        /// <param name="header">是否把第一行作为列名</param>
//        /// <param name="eType">Excel 版本 </param>
//        /// <returns>返回Excel第一个工作表中的数据 DataSet </returns>
//        public static DataSet ExcelToDataSet(string excelPath, bool header, ExcelType eType)
//        {
//            string connectstring = GetExcelConnectstring(excelPath, header, eType);
//            return ExcelToDataSet(connectstring);
//        }

//        /// <summary>
//        /// EXCEL所有工作表导入DataSet
//        /// </summary>
//        /// <param name="connectstring">excel连接字符串</param>
//        /// <returns>返回Excel第一个工作表中的数据 DataSet </returns>
//        public static DataSet ExcelToDataSet(string connectstring)
//        {
//            using (OleDbConnection conn = new OleDbConnection(connectstring))
//            {
//                DataSet ds = new DataSet();
//                List<string> tableNames = GetExcelTablesName(conn);

//                foreach (string tableName in tableNames)
//                {
//                    OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * FROM [" + tableName + "]", conn);
//                    adapter.Fill(ds, tableName);
//                }
//                return ds;
//            }
//        }

//        #endregion

//        #region DataTable导入Excel
//        /// <summary>
//        /// 将DataTable导出为Excel(OleDb 方式操作）
//        /// </summary>
//        /// <param name="dataTable">表</param> 
//        public static void DataTableToExcel(DataTable dataTable)
//        {
//            string fileName = GetSaveFilePath();
//            if (string.IsNullOrEmpty(fileName)) return;

//            OleDbConnection oleDbConn = new OleDbConnection();
//            OleDbCommand oleDbCmd = new OleDbCommand();
//            string sqlString = "";
//            try
//            {
//                oleDbConn.ConnectionString = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" + fileName + @";Extended ProPerties=""Excel 8.0;HDR=Yes;""";
//                oleDbConn.Open();
//                oleDbCmd.CommandType = CommandType.Text;
//                oleDbCmd.Connection = oleDbConn;
//                sqlString = "CREATE TABLE sheet1 (";
//                for (int i = 0; i < dataTable.Columns.Count; i++)
//                {
//                    // 字段名称出现关键字会导致错误。
//                    if (i < dataTable.Columns.Count - 1)
//                        sqlString += "[" + dataTable.Columns[i].Caption + "] TEXT(100) ,";
//                    else
//                        sqlString += "[" + dataTable.Columns[i].Caption + "] TEXT(200) )";
//                }
//                oleDbCmd.CommandText = sqlString;
//                oleDbCmd.ExecuteNonQuery();
//                for (int j = 0; j < dataTable.Rows.Count; j++)
//                {
//                    sqlString = "INSERT INTO sheet1 VALUES('";
//                    for (int i = 0; i < dataTable.Columns.Count; i++)
//                    {
//                        if (i < dataTable.Columns.Count - 1)
//                            sqlString += dataTable.Rows[j][i].ToString() + " ','";
//                        else
//                            sqlString += dataTable.Rows[j][i].ToString() + " ')";
//                    }
//                    oleDbCmd.CommandText = sqlString;
//                    oleDbCmd.ExecuteNonQuery();
//                }
//                MsgBox.ShowInformation("导出EXCEL成功");
//            }
//            catch (Exception ex)
//            {
//                MsgBox.ShowError("导出EXCEL失败:" + ex.Message);
//            }
//            finally
//            {
//                oleDbCmd.Dispose();
//                oleDbConn.Close();
//                oleDbConn.Dispose();
//            }

//        }
//        #endregion

//        #region 辅助方法
//        private static DataTable GetSheetSchema(string sheetName, OleDbConnection connection)
//        {
//            DataTable schemaTable = null;
//            IDbCommand cmd = new OleDbCommand();
//            cmd.CommandText = string.Format("select * from [{0}]", sheetName);
//            cmd.Connection = connection;

//            using (IDataReader reader = cmd.ExecuteReader(CommandBehavior.KeyInfo | CommandBehavior.SchemaOnly))
//            {
//                schemaTable = reader.GetSchemaTable();
//            }
//            return schemaTable;
//        }

//        /// <summary>
//        /// 判断工作表名是否存在
//        /// </summary>
//        /// <param name="connection">excel连接</param>
//        /// <param name="sheetName">工作表名称  例如：Sheet1$</param>
//        /// <returns></returns>
//        private static bool IsExistSheetName(OleDbConnection connection, string sheetName)
//        {
//            List<string> list = GetExcelTablesName(connection);
//            foreach (string tName in list)
//            {
//                if (tName.ToLower() == sheetName.ToLower())
//                    return true;
//            }
//            return false;
//        }

//        private static string GetSaveFilePath()
//        {
//            string fileName = string.Empty;
//            SaveFileDialog saveFileDialog = new SaveFileDialog();
//            saveFileDialog.Filter = "xls files (*.xls)|*.xls";
//            if (saveFileDialog.ShowDialog() == DialogResult.OK)
//            {
//                fileName = saveFileDialog.FileName;
//                if (File.Exists(fileName))
//                {
//                    try
//                    {
//                        File.Delete(fileName);
//                    }
//                    catch
//                    {
//                        MessageBox.Show("该文件正在使用中,关闭文件或重新命名导出文件再试!");
//                        return string.Empty;
//                    }
//                }
//            }
//            return fileName;
//        }

//        #endregion
//    }
//}
