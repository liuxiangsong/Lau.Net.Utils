using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils
{
    public static class SqlUtil
    {
        /// <summary>
        /// 通过字段、值字典获取插入表的sql语句
        /// </summary>
        /// <param name="tableName">数据库表名</param>
        /// <param name="fieldDict">字段、值字典</param>
        /// <param name="ignoreEmptyField">是否忽略字段值为空的字段</param>
        /// <returns></returns>
        public static string GetInsertSqlByDict(string tableName, Dictionary<string, string> fieldDict, bool ignoreEmptyField = true)
        {
            if (ignoreEmptyField)
            {
                fieldDict = fieldDict.Where(kv => !string.IsNullOrEmpty(kv.Value)).ToDictionary(kv => kv.Key, kv => kv.Value);
            }
            var fields = string.Join(",", fieldDict.Select(s => s.Key));
            var fieldsValue = string.Join(",", fieldDict.Select(s => $"'{s.Value}'"));
            var sql = $@"insert into {tableName} ({fields}) values ({fieldsValue})";
            return sql;
        }
    }
}
