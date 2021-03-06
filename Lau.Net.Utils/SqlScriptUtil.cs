using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace Lau.Net.Utils
{
    public class SqlScriptUtil
    {

        #region 附加、分离、备份、恢复数据库操作

        /// <summary>
        /// 附加SqlServer数据库
        /// </summary>
        public static bool AttachDB(string connectionString, string dataBaseName, string dataBase_MDF, string dataBase_LDF)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = "sp_attach_db";
                cmd.Parameters.Add(new SqlParameter("dbname", SqlDbType.NVarChar));
                cmd.Parameters["dbname"].Value = dataBaseName;
                cmd.Parameters.Add(new SqlParameter("filename1", SqlDbType.NVarChar));
                cmd.Parameters["filename1"].Value = dataBase_MDF;
                cmd.Parameters.Add(new SqlParameter("filename2", SqlDbType.NVarChar));
                cmd.Parameters["filename2"].Value = dataBase_LDF;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.ExecuteNonQuery();
            }

            return true;
        }

        /// <summary>
        /// 分离SqlServer数据库
        /// </summary>
        public static bool DetachDB(string connectionString, string dataBaseName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = "sp_detach_db";
                cmd.Parameters.Add(new SqlParameter("dbname", SqlDbType.NVarChar));
                cmd.Parameters["dbname"].Value = dataBaseName;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.ExecuteNonQuery();
            }
            return true;
        }

        /// <summary>
        /// 还原数据库
        /// </summary>
        public static bool RestoreDataBase(string connectionString, string dataBaseName, string DataBaseOfBackupPath, string DataBaseOfBackupName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = "use master;restore database @DataBaseName From disk = @BackupFile with replace;";
                cmd.Parameters.Add(new SqlParameter("DataBaseName", SqlDbType.NVarChar));
                cmd.Parameters["DataBaseName"].Value = dataBaseName;
                cmd.Parameters.Add(new SqlParameter("BackupFile", SqlDbType.NVarChar));
                cmd.Parameters["BackupFile"].Value = Path.Combine(DataBaseOfBackupPath, DataBaseOfBackupName);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
            }
            return true;
        }

        /// <summary>
        /// 备份SqlServer数据库
        /// </summary>
        public static bool BackupDataBase(string connectionString, string dataBaseName, string DataBaseOfBackupPath, string DataBaseOfBackupName)
        {
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = conn;
                cmd.CommandText = "use master;backup database @dbname to disk = @backupname;";
                cmd.Parameters.Add(new SqlParameter("dbname", SqlDbType.NVarChar));
                cmd.Parameters["dbname"].Value = dataBaseName;
                cmd.Parameters.Add(new SqlParameter("backupname", SqlDbType.NVarChar));
                cmd.Parameters["backupname"].Value = Path.Combine(DataBaseOfBackupPath, DataBaseOfBackupName);
                cmd.CommandType = CommandType.Text;
                cmd.ExecuteNonQuery();
            }
            return true;
        }

        #endregion

        public static DataRow GetDataEngineInfo(string connectionString)
        {
            try
            {
                DataTable dt = new DataTable();
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlCommand cmd = conn.CreateCommand();
                    cmd.CommandText = @"select  ProductVersion=SERVERPROPERTY('ProductVersion'),
		                                        ProductLevel=	SERVERPROPERTY('ProductLevel'),
		                                        Edition=		SERVERPROPERTY('Edition')";
                    using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                    {
                        sda.Fill(dt);
                    }
                }
                return dt.Rows[0];
            }
            catch (Exception ex)
            {
                MsgBox.ShowError(ex.Message);
                return null;
            }

        }
    }
}
