using Lau.Net.Utils.Excel;
using Lau.Net.Utils.Sql;
using NPOI.SS.Formula.Functions;
using NUnit.Framework;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class SqlUtilTest
    {
        string connectionString = "Data Source=.;Initial Catalog=Test;Integrated Security=True;";
        SqlSugarUtil _sqlSugarUtil;
        public SqlUtilTest()
        {
            _sqlSugarUtil = new SqlSugarUtil(connectionString);
        }

        #region SqlSugarUtil测试用例
        [Test]
        public void CreateTestTable()
        {
            // 创建测试表
            string createTableSql = @"
                IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[SqlSugarTestTable]') AND type in (N'U'))
                BEGIN
                    CREATE TABLE [dbo].[SqlSugarTestTable](
                        [Id] [int] IDENTITY(1,1) NOT NULL,
                        [Name] [nvarchar](50) NULL,
                        [Age] [int] NULL,
                        CONSTRAINT [PK_SqlSugarTestTable] PRIMARY KEY CLUSTERED ([Id] ASC)
                    )
                END";

            int result = _sqlSugarUtil.ExecuteSql(createTableSql);

            // 验证表是否创建成功
            var tableSchema = _sqlSugarUtil.GetTableSchema("SqlSugarTestTable");
            Assert.IsNotNull(tableSchema);
            Assert.AreEqual(3, tableSchema.Count);

            // 验证列信息
            Assert.AreEqual("Id", tableSchema[0].DbColumnName);
            Assert.AreEqual("Name", tableSchema[1].DbColumnName);
            Assert.AreEqual("Age", tableSchema[2].DbColumnName);

            Console.WriteLine("测试表 'SqlSugarTestTable' 创建成功。");
        }
        [Test]
        public void TestSqlSugarUtil()
        {
            ///测试获取表架构
            //var tableSchema = _sqlSugarUtil.GetTableSchema("tb1");

            // 测试执行SQL语句
            var sql = "SELECT * FROM SqlSugarTestTable WHERE Id = @Id;";
            var dataTable = _sqlSugarUtil.GetDataTable(sql, new SugarParameter($"@Id", 1));
            Assert.IsTrue(dataTable.Rows.Count > 0);
        }

        [Test]
        public void TestBulkInsertSqlSugarUtil()
        {
            var dt = DataTableUtil.CreateTable("Name");
            //dt.Rows.Add("张三", 1);
            dt.Rows.Add("李四2");
            _sqlSugarUtil.BulkInsert(dt, "SqlSugarTestTable");
            var dataTable = _sqlSugarUtil.GetDataTable("SELECT * FROM SqlSugarTestTable;");
            Assert.IsTrue(dataTable.Rows.Count > 0);
        }
        #endregion
    }
}
