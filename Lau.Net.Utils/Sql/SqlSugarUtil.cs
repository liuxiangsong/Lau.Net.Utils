using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq.Expressions;
using DbType = SqlSugar.DbType;

namespace Lau.Net.Utils.Sql
{
    /// <summary>
    /// SqlSugar数据库操作辅助类
    /// </summary>
    public class SqlSugarUtil
    {
        private readonly SqlSugarScope _db;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="connectionString">数据库连接字符串</param>
        /// <param name="dbType">数据库类型,默认为SqlServer</param>
        public SqlSugarUtil(string connectionString, DbType dbType = DbType.SqlServer)
        {
            _db = new SqlSugarScope(new ConnectionConfig()
            {
                ConnectionString = connectionString,
                DbType = dbType,
                IsAutoCloseConnection = true
            });
        }

        #region 查询操作

        /// <summary>
        /// 获取指定类型的所有数据
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <returns>实体列表</returns>
        public List<T> GetList<T>() where T : class, new()
        {
            return _db.Queryable<T>().ToList();
        }

        /// <summary>
        /// 根据条件获取数据
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="whereExpression">查询条件表达式</param>
        /// <returns>实体列表</returns>
        public List<T> GetList<T>(Expression<Func<T, bool>> whereExpression) where T : class, new()
        {
            return _db.Queryable<T>().Where(whereExpression).ToList();
        }

        /// <summary>
        /// 获取单个实体
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="id">主键值</param>
        /// <returns>实体</returns>
        public T GetById<T>(object id) where T : class, new()
        {
            return _db.Queryable<T>().InSingle(id);
        }

        /// <summary>
        /// 执行SQL语句并返回DataTable
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <returns>查询结果DataTable</returns>
        public DataTable GetDataTable(string sql)
        {
            return _db.Ado.GetDataTable(sql);
        }

        #endregion

        #region 插入操作

        /// <summary>
        /// 插入单个实体
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="entity">要插入的实体</param>
        /// <returns>受影响的行数</returns>
        public int Insert<T>(T entity) where T : class, new()
        {
            return _db.Insertable(entity).ExecuteCommand();
        }

        /// <summary>
        /// 批量插入数据
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="entities">实体列表</param>
        /// <returns>受影响的行数</returns>
        public int BulkInsert<T>(List<T> entities) where T : class, new()
        {
            return _db.Fastest<T>().BulkCopy(entities);
        }

        /// <summary>
        /// 批量复制DataTable到指定表
        /// </summary>
        /// <param name="dataTable">要复制的DataTable</param>
        /// <param name="tableName">目标表名</param>
        /// <returns>受影响的行数</returns>
        public int BulkCopyDataTable(DataTable dataTable, string tableName)
        {
            return _db.Fastest<object>().AS(tableName).BulkCopy(dataTable);
        }

        #endregion

        #region 更新操作

        /// <summary>
        /// 更新单个实体
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="entity">要更新的实体</param>
        /// <returns>受影响的行数</returns>
        public int Update<T>(T entity) where T : class, new()
        {
            return _db.Updateable(entity).ExecuteCommand();
        }

        /// <summary>
        /// 批量更新数据
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="entities">实体列表</param>
        /// <returns>受影响的行数</returns>
        public int BulkUpdate<T>(List<T> entities) where T : class, new()
        {
            return _db.Fastest<T>().BulkUpdate(entities);
        }

        #endregion

        #region 删除操作

        /// <summary>
        /// 删除单个实体
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="entity">要删除的实体</param>
        /// <returns>受影响的行数</returns>
        public int Delete<T>(T entity) where T : class, new()
        {
            return _db.Deleteable(entity).ExecuteCommand();
        }

        /// <summary>
        /// 根据主键删除
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        /// <param name="id">主键值</param>
        /// <returns>受影响的行数</returns>
        public int DeleteById<T>(object id) where T : class, new()
        {
            return _db.Deleteable<T>().In(id).ExecuteCommand();
        }

        #endregion

        #region 其他操作

        /// <summary>
        /// 获取指定表的架构信息
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <returns>表的列信息列表</returns>
        public List<DbColumnInfo> GetTableSchema(string tableName)
        {
            return _db.DbMaintenance.GetColumnInfosByTableName(tableName);
        }

        /// <summary>
        /// 清空指定表的所有数据
        /// </summary>
        /// <param name="tableName">要清空的表名</param>
        public void TruncateTable(string tableName)
        {
            _db.Ado.ExecuteCommand($"TRUNCATE TABLE {tableName}");
        }

        /// <summary>
        /// 执行事务
        /// </summary>
        /// <param name="action">要在事务中执行的操作</param>
        /// <returns>是否成功</returns>
        public bool ExecuteTransaction(Action action)
        {
            try
            {
                _db.Ado.BeginTran();
                action();
                _db.Ado.CommitTran();
                return true;
            }
            catch
            {
                _db.Ado.RollbackTran();
                return false;
            }
        }

        #endregion
    }
}
