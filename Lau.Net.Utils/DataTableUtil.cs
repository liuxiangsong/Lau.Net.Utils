using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Lau.Net.Utils
{
    /// <summary>
    /// DataTable帮助类
    /// </summary>
    public class DataTableUtil
    {
        /// <summary>
        /// 给DataTable增加自增长列。
        /// 注：如果表中存在AutoID字段则不做任何处理
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static DataTable AddIdentityColumn(DataTable dt)
        {
            if (!dt.Columns.Contains("AutoID"))
            {
                DataColumn dc = dt.Columns.Add("AutoID");
                dc.AutoIncrement = true;
                dc.AutoIncrementSeed = 1;
                dc.AutoIncrementStep = 1;
                dc.SetOrdinal(0);

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    dt.Rows[i]["AutoID"] = i + 1;
                }
            }
            return dt;
        }

        /// <summary>
        /// 检查表中是否有数据行
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>有数据行则返回True</returns>
        public static bool HasRow(DataTable dt)
        {
            if (dt != null && dt.Rows.Count > 0)
                return true;
            return false;
        }

        /// <summary>
        /// 将DataTable转化为实体集合
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static List<T> DataTableToList<T>(DataTable dt) where T : class
        {
            if (!HasRow(dt))
                return new List<T>();
            List<T> m_EntityList = new List<T>();

            T m_EntityObj = default(T);
            foreach (DataRow dr in dt.Rows)
            {
                m_EntityObj = Activator.CreateInstance<T>();
                foreach (PropertyInfo p in m_EntityObj.GetType().GetProperties())
                {
                    if (dt.Columns.Contains(p.Name) && p.CanWrite)
                    {
                        p.SetValue(m_EntityObj, TypeUtil.ToObject(dr[p.Name]), null);
                    }
                }
                m_EntityList.Add(m_EntityObj);
            }
            return m_EntityList;
        }

        /// <summary>
        /// 将DataRow转化为实体类对象
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="dr"></param>
        /// <returns></returns>
        public static T DataRowToEntity<T>(DataRow dr) where T : class
        {
            T m_EntityObj = default(T);
            m_EntityObj = Activator.CreateInstance<T>();
            foreach (PropertyInfo p in m_EntityObj.GetType().GetProperties())
            {
                if (dr.Table.Columns.Contains(p.Name) && p.CanWrite)
                {
                    p.SetValue(m_EntityObj, TypeUtil.ToObject(dr[p.Name]), null);
                }
            }
            return m_EntityObj;
        }

        /// <summary>
        /// 将实体类集合转化为DataTable
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="list">实体类集合</param>
        /// <returns></returns>
        public static DataTable ListToDataTable<T>(List<T> list) where T : class
        {
            if (list == null || list.Count <= 0)
            {
                return null;
            }
            DataTable table = new DataTable(typeof(T).Name);
            DataRow row;

            PropertyInfo[] myPropertyInfo = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo info in myPropertyInfo)
            {
                table.Columns.Add(info.Name, Nullable.GetUnderlyingType(info.PropertyType) ?? info.PropertyType);
            }
            foreach (T entity in list)
            {
                row = table.NewRow();
                for (int index = 0; index < myPropertyInfo.Length; index++)
                {
                    row[index] = myPropertyInfo[index].GetValue(entity, null) ?? DBNull.Value;
                }
                table.Rows.Add(row);
            }
            return table;
        }

        /// <summary>
        /// 通过字符串变量创建表字段，字段格式可以是：
        /// 1）列名1，列名2（如：a,b,c,d)
        /// 2）列名1|列类型，列名2|列类型(如：a|int,b|bool,c|decimal)
        /// 3）以上两种的混合格式(如：a|int,b,c|datetime)
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        public static DataTable CreateTable(params string[] columns)
        {
            DataTable dt = new DataTable();
            var cols = CreateColumns(columns);
            dt.Columns.AddRange(cols);
            return dt;
        }

        /// <summary>
        /// 通过字符串变量创建表格列，字段格式可以是：
        /// 1）列名1，列名2（如：a,b,c,d)
        /// 2）列名1|列类型，列名2|列类型(如：a|int,b|bool,c|decimal)
        /// 3）以上两种的混合格式(如：a|int,b,c|datetime)
        /// </summary>
        /// <param name="columns"></param>
        /// <returns></returns>
        public static DataColumn[] CreateColumns(params string[] columns)
        {
            var cols = columns.Select(col =>
            {
                if (string.IsNullOrEmpty(col))
                {
                    col = string.Empty;
                };
                string[] column = col.Split('|');
                DataColumn newColumn = null;
                if (column.Length == 2)
                {
                    newColumn = new DataColumn(column[0], ConvertType(column[1]));
                }
                else
                {
                    newColumn = new DataColumn(column[0]);
                }
                return newColumn;
            }).ToArray();
            return cols;
        }

        /// <summary>
        /// 在dt指定列后插入指定数量的空白列
        /// </summary>
        /// <param name="dt">需要插入列的表格</param>
        /// <param name="theColumnName">在该列后插入新的列</param>
        /// <param name="columnCount">空白列的数量</param>
        public static void InsertColumnsAfter(DataTable dt, string theColumnName,int columnCount)
        {
            var cols = new string[columnCount];
            InsertColumnsAfter(dt, theColumnName, cols);
        }
        /// <summary>
        /// 在dt指定列后添加新的列
        /// </summary>
        /// <param name="dt">需要插入列的表格</param>
        /// <param name="theColumnName">在该列后插入新的列</param>
        /// <param name="columns">同CreateColumns的参数</param>
        public static void InsertColumnsAfter(DataTable dt, string theColumnName, params string[] columns)
        {
            var theColumn = dt.Columns[theColumnName];
            var cols = CreateColumns(columns);
            dt.Columns.AddRange(cols);
            for (var i = 0; i < cols.Length; i++)
            {
                cols[i].SetOrdinal(theColumn.Ordinal + 1 + i);
            }
        }

        public static DataTable CreateTable(params KeyValuePair<string, Type>[] columns)
        {
            DataTable dt = new DataTable();
            foreach (var keyvalue in columns)
            {
                dt.Columns.Add(keyvalue.Key, keyvalue.Value);
            }
            return dt;
        }

        public static void ImportRows(DataTable sourceTable, DataRowCollection dataRows)
        {
            try
            {
                if (dataRows == null) return;
                foreach (DataRow row in dataRows)
                {
                    sourceTable.ImportRow(row);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region 私有方法
        private static Type ConvertType(string typeName)
        {
            typeName = typeName.ToLower().Replace("system.", "");
            Type newType = typeof(string);
            switch (typeName)
            {
                case "boolean":
                case "bool":
                    newType = typeof(bool);
                    break;
                case "int16":
                case "short":
                    newType = typeof(short);
                    break;
                case "int32":
                case "int":
                    newType = typeof(int);
                    break;
                case "long":
                case "int64":
                    newType = typeof(long);
                    break;
                case "uint16":
                case "ushort":
                    newType = typeof(ushort);
                    break;
                case "uint32":
                case "uint":
                    newType = typeof(uint);
                    break;
                case "uint64":
                case "ulong":
                    newType = typeof(ulong);
                    break;
                case "single":
                case "float":
                    newType = typeof(float);
                    break;
                case "string":
                    newType = typeof(string);
                    break;
                case "guid":
                    newType = typeof(Guid);
                    break;
                case "decimal":
                    newType = typeof(decimal);
                    break;
                case "double":
                    newType = typeof(double);
                    break;
                case "datetime":
                    newType = typeof(DateTime);
                    break;
                case "byte":
                    newType = typeof(byte);
                    break;
                case "char":
                    newType = typeof(char);
                    break;
            }
            return newType;
        }
        #endregion
    }
}
