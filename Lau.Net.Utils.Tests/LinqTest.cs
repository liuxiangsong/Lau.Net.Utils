using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class LinqTest
    {
        /// <summary>
        /// 在Linq查询语法中定义变量
        /// </summary>
        [Test]
        public void DefineVar()
        {
            var dt = DataTableUtil.CreateTable("name", "num|int");
            dt.Rows.Add("张三", 12);
            dt.Rows.Add("张三", 3);
            dt.Rows.Add("李四", 12);
            var query = from row in dt.AsEnumerable()
                        group row by row.Field<string>("name")
                        into g
                        //定义变量
                        let amount = g.Sum(r => r.Field<int>("num"))
                        select new
                        {
                            g.Key,
                            amount
                        };
            var list = query.ToList();
        }
    }
}
