﻿using NPOI.SS.Formula.Functions;
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
    public static class DataTableUtil
    {
        #region 创建表

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
        /// 创建表
        /// </summary>
        /// <param name="colDict">列名和列类型典</param>
        /// <returns></returns>
        public static DataTable CreateTable(Dictionary<string, Type> colDict)
        {
            DataTable dt = new DataTable();
            foreach (KeyValuePair<string, Type> keyvalue in colDict)
            {
                dt.Columns.Add(keyvalue.Key, keyvalue.Value);
            }
            return dt;
        }
        #endregion

        #region 创建、添加、插入、移除、调整列及添加自增长列
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
                string[] arr = col.Split('|');
                DataColumn newColumn = null;
                if (arr.Length == 2)
                {
                    newColumn = new DataColumn(arr[0], ConvertType(arr[1]));
                }
                else
                {
                    newColumn = new DataColumn(arr[0]);
                }
                return newColumn;
            }).ToArray();
            return cols;
        }

        /// <summary>
        /// 创建列
        /// </summary>
        /// <param name="columnType">列类型</param>
        /// <param name="columns">列名</param>
        /// <returns></returns>
        public static DataColumn[] CreateColumns(Type columnType, params string[] columns)
        {
            var cols = columns.Select(col =>
            {
                if (string.IsNullOrEmpty(col))
                {
                    col = string.Empty;
                };
                DataColumn newColumn;
                if (columnType != null)
                {
                    newColumn = new DataColumn(col, columnType);
                }
                else
                {
                    newColumn = new DataColumn(col);
                }
                return newColumn;
            }).ToArray();
            return cols;
        }

        /// <summary>
        /// 添加指定列数
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="columnCount"></param>
        /// <param name="columnType">添加列的类型</param> 
        public static void AddColumns(DataTable dt, int columnCount, Type columnType = null)
        {
            var colArr = new string[columnCount];
            var cols = CreateColumns(columnType, colArr);
            dt.Columns.AddRange(cols);
        }
        /// <summary>
        /// 添加指定列名
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="columnNames">列名称</param>
        /// <param name="columnType">添加列的类型</param>
        public static void AddColumns(DataTable dt, IEnumerable<string> columnNames, Type columnType = null)
        {
            var cols = CreateColumns(columnType, columnNames.ToArray());
            dt.Columns.AddRange(cols);
        }

        /// <summary>
        /// 在dt指定列后插入指定数量的空白列
        /// </summary>
        /// <param name="dt">需要插入列的表格</param>
        /// <param name="theColumnName">在该列后插入新的列</param>
        /// <param name="columnCount">空白列的数量</param>
        /// <param name="columnType">添加列的类型</param> 
        public static void InsertColumnsAfter(DataTable dt, string theColumnName, int columnCount, Type columnType = null)
        {
            var cols = new string[columnCount];
            var columns = CreateColumns(columnType, cols);
            InsertColumnsAfter(dt, theColumnName, columns);
        }
        /// <summary>
        /// 在dt指定列后添加新的列
        /// </summary>
        /// <param name="dt">需要插入列的表格</param>
        /// <param name="theColumnName">在该列后插入新的列</param>
        /// <param name="columns">同CreateColumns的参数</param>
        public static void InsertColumnsAfter(DataTable dt, string theColumnName, params string[] columns)
        {
            var cols = CreateColumns(columns);
            InsertColumnsAfter(dt, theColumnName, cols);
        }

        private static void InsertColumnsAfter(DataTable dt, string theColumnName, params DataColumn[] columns)
        {
            var theColumn = dt.Columns[theColumnName];
            dt.Columns.AddRange(columns);
            for (var i = 0; i < columns.Length; i++)
            {
                columns[i].SetOrdinal(theColumn.Ordinal + 1 + i);
            }
        }

        /// <summary>
        /// 给DataTable增加自增长列(序号列），默认设置为第一列
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="identityColumnName">自增长列名，默认为AutoID，如果dt中已存在指定自增长列名，则不做任务处理</param>
        public static void AddIdentityColumn(DataTable dt, string identityColumnName = "AutoID")
        {
            if (dt.Columns.Contains(identityColumnName))
            {
                return;
            }
            DataColumn dc = dt.Columns.Add(identityColumnName, typeof(int));
            dc.AutoIncrement = true;
            dc.AutoIncrementSeed = 1;
            dc.AutoIncrementStep = 1;
            dc.SetOrdinal(0);

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dt.Rows[i][identityColumnName] = i + 1;
            }
        }

        /// <summary>
        /// 移除指定列
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="removeColNames">需要移除的列名</param>
        public static void RemoveColumns(DataTable dt, params string[] removeColNames)
        {
            if (!removeColNames.HasItem())
            {
                return;
            }
            var colNames = removeColNames.Where(c => dt.Columns.Contains(c)).ToList();
            foreach (var colName in colNames)
            {
                dt.Columns.Remove(colName);
            }
        }
        /// <summary>
        /// 根据传入的列名称调整DataTable的列，不包含在columnNames中的列将被移除
        /// </summary>
        /// <param name="dt">要调整的DataTable</param>
        /// <param name="columnNames">需要保留的列名称列表</param>
        /// <param name="createMissingColumns">如果为true，则创建DataTable中不存在的列</param>
        public static void AdjustColumns(DataTable dt, List<string> columnNames, bool createMissingColumns = false)
        {
            if (dt == null || columnNames == null || !columnNames.Any())
            {
                return;
            }

            // 移除不在列表中的列
            var columnsToRemove = dt.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .Where(name => !columnNames.Contains(name))
                .ToList();
            RemoveColumns(dt, columnsToRemove.ToArray());

            // 如果需要，创建缺失的列 
            if (createMissingColumns)
            {
                var columnsToAdd = columnNames.Where(name => !dt.Columns.Contains(name)).ToList();
                AddColumns(dt, columnsToAdd);
            }

            // 调整列顺序
            for (int i = 0; i < columnNames.Count; i++)
            {
                if (dt.Columns.Contains(columnNames[i]))
                {
                    dt.Columns[columnNames[i]].SetOrdinal(i);
                }
            }
        }

        /// <summary>
        /// 根据传入的列名称字典调整DataTable的列
        /// </summary>
        /// <param name="dt">要调整的DataTable</param>
        /// <param name="columnNameMap">列名映射字典,key为原列名,value为新列名，不包含在key的列将被移除</param>
        /// <param name="createMissingColumns">如果为true,则创建DataTable中不存在的列</param>
        public static void AdjustColumns(DataTable dt, Dictionary<string, string> columnNameMap, bool createMissingColumns = false)
        {
            if (dt == null || columnNameMap == null || !columnNameMap.Any())
            {
                return;
            }

            // 移除不在字典中的列
            var columnsToRemove = dt.Columns.Cast<DataColumn>()
                .Select(c => c.ColumnName)
                .Where(name => !columnNameMap.ContainsKey(name))
                .ToList();
            RemoveColumns(dt, columnsToRemove.ToArray());

            // 重命名列并调整顺序
            int ordinal = 0;
            foreach (var kvp in columnNameMap)
            {
                if (dt.Columns.Contains(kvp.Key))
                {
                    dt.Columns[kvp.Key].ColumnName = kvp.Value;
                    dt.Columns[kvp.Value].SetOrdinal(ordinal);
                    ordinal++;
                }
                else if (createMissingColumns)
                {
                    var newColumn = new DataColumn(kvp.Value);
                    dt.Columns.Add(newColumn);
                    newColumn.SetOrdinal(ordinal);
                    ordinal++;
                }
            }
        }
        #endregion

        #region 表添加数据行、汇总行
        /// <summary>
        /// 将dataRows添加到sourceTable中
        /// </summary>
        /// <param name="sourceTable"></param>
        /// <param name="dataRows"></param>
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

        /// <summary>
        /// 将数据行复制到另一个表中
        /// </summary>
        /// <param name="targetTable"></param>
        /// <param name="sourceRow"></param>
        public static void CopyDataRowToTable(DataTable targetTable, DataRow sourceRow)
        {
            if (targetTable == null || sourceRow == null)
            {
                return;
            }
            var targetRow = targetTable.NewRow();
            //targetRow.ItemArray = sourceRow.ItemArray;
            foreach (DataColumn column in sourceRow.Table.Columns)
            {
                if (targetTable.Columns.Contains(column.ColumnName))
                {
                    targetRow.SetField(column.ColumnName, sourceRow[column.ColumnName]);
                }
            }
            targetTable.Rows.Add(targetRow);
        }

        /// <summary>
        /// 创建汇总行,如果汇总行每列的值都小于0，则返回null
        /// </summary>
        /// <param name="dt">原数据表</param>
        /// <param name="addToTable">是否将汇总行添加到表里</param>
        /// <param name="sumarryColNames">需要汇总的列，为空时，默认汇总列为数值类型的列</param>
        /// <param name="rowFilter">汇总过滤数据行条件</param>
        /// <returns>如果汇总行每列的值都小于0，则返回null</returns>
        public static DataRow CreateSummaryRow(DataTable dt, bool addToTable = false, List<string> sumarryColNames = null, string rowFilter = null)
        {
            var summaryRow = dt.NewRow();
            var colList = dt.Columns.Cast<DataColumn>();
            var tempTable = dt;
            if (!string.IsNullOrWhiteSpace(rowFilter))
            {
                var dataView = new DataView(dt);
                dataView.RowFilter = rowFilter;
                tempTable = dataView.ToTable();
            }

            if (sumarryColNames != null && sumarryColNames.Count > 0)
            {
                colList = colList.Where(c => sumarryColNames.Contains(c.ColumnName)).ToList();
            }
            foreach (DataColumn column in colList)
            {
                if (column.DataType == typeof(decimal) || column.DataType == typeof(double) || column.DataType == typeof(int))
                {
                    summaryRow[column] = tempTable.Compute($"SUM({column.ColumnName})", "");
                }
                else
                {
                    summaryRow[column] = DBNull.Value;
                }
            }
            if (summaryRow.ItemArray.Count(i => i.As<int>() > 0) < 1)
            {
                return null;
            }
            if (addToTable)
            {
                dt.Rows.Add(summaryRow);
            }
            return summaryRow;
        }
        #endregion

        #region 实体类集合与DataTable转化
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
        #endregion

        #region 检查表中是否有数据行
        /// <summary>
        /// 检查表中是否有数据行
        /// </summary>
        /// <param name="dt"></param>
        /// <returns>有数据行则返回True</returns>
        public static bool HasRow(this DataTable dt)
        {
            if (dt != null && dt.Rows.Count > 0)
            {
                return true;
            }
            return false;
        }
        #endregion

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

    /// <summary>
    /// DataRow扩展方法
    /// </summary>
    public static class DataRowUtil
    {
        /// <summary>
        /// 获取DataRow的值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <param name="columnName">列名</param>
        /// <param name="options">可选值，如果options不为null，而且单元格的值不包含在options中，则返回类型T的默认值</param>
        /// <returns></returns>
        public static T GetValue<T>(this DataRow row, string columnName, params T[] options)
        {
            if (row == null || !row.Table.Columns.Contains(columnName))
            {
                return default(T).As<T>(); ;
            }
            var value = row[columnName].As<T>();
            if (options != null && options.Length > 0 && !options.Contains(value))
            {
                return default(T).As<T>(); ;
            }
            return value;
        }

        /// <summary>
        /// 获取DataRow的值,并格式化
        /// </summary>
        /// <param name="row"></param>
        /// <param name="columnName"></param>
        /// <param name="format">数字类型.ToString方法中的format</param>
        /// <param name="isZoreThenEmpty">为true时，如果值等于0就返回空字段串</param>
        /// <returns></returns>
        public static string GetFormatNumberValue(this DataRow row, string columnName, string format = "0.#####", bool isZoreThenEmpty = true)
        {
            var value = row.GetValue<decimal>(columnName);
            if (isZoreThenEmpty && value == 0)
            {
                return string.Empty;
            }
            if (!string.IsNullOrWhiteSpace(format))
            {
                return value.ToString(format);
            }
            return value.ToString();
        }
    }
}
