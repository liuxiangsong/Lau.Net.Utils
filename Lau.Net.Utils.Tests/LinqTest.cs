using NUnit.Framework;
using System.Linq;
using System.Collections.Generic;

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
            var students = new List<Student>
            {
                new Student { Name = "张三", Score = 12 },
                new Student { Name = "张三", Score = 3 },
                new Student { Name = "李四", Score = 12 }
            };

            var query = from s in students
                        group s by s.Name into g
                        let amount = g.Sum(x => x.Score) ?? 0
                        select new StudentAmount
                        {
                            Name = g.Key,
                            Amount = amount
                        };

            var list = query.ToList();
        }

        /// <summary>
        /// Linq左连接示例
        /// </summary>
        [Test]
        public void LeftJoinExample()
        {
            var students = new List<Student>
            {
                new Student { Id = 1, Name = "张三" },
                new Student { Id = 2, Name = "李四" },
                new Student { Id = 3, Name = "王五" }
            };

            var scores = new List<StudentScore>
            {
                new StudentScore { Id = 1, Score = 90 },
                new StudentScore { Id = 2, Score = 80 }
            };

            var query = from s in students
                        join sc in scores
                        on s.Id equals sc.Id into gj
                        from subScore in gj.DefaultIfEmpty()
                        select new Student
                        {
                            Id = s.Id,
                            Name = s.Name,
                            Score = subScore?.Score
                        };

            var list = query.ToList();
        }

        private class Student
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public int? Score { get; set; }
        }

        private class StudentScore
        {
            public int Id { get; set; }
            public int Score { get; set; }
        }

        private class StudentAmount
        {
            public string Name { get; set; }
            public int Amount { get; set; }
        }
    }
}
