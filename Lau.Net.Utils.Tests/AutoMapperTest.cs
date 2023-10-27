using NPOI.OpenXmlFormats.Spreadsheet;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class AutoMapperTest
    {
        [Test]
        public void MapTableTest()
        {
            var dt = DataTableUtilTest.CreateTestTable();
            var cols = dt.Columns;
            var colDict = new Dictionary<string, string>
            {
                {"月份","Month" },
                {"生产总数量","Amount" }
            }; 
            var list = AutoMapperUtil.MapToList<TestInfo>(dt.Rows, cfg =>
            {
                var map = cfg.CreateMap<DataRow, TestInfo>();
                //指定字段的映射关系
                map.ForMember(dest => dest.Amount, opt => opt.MapFrom(row => row["Amount"]));
            }); 
            Assert.Greater(list[0].Amount, 0);
        }

        [Test]
        public void MapToTest()
        {
            var userInfo = new UserInfo
            {
                UserId = "2018001",
                RealName = "大师兄",
                DepartmentId = "001",
                DepartmentName = "技术部"
            };
            var dto = AutoMapperUtil.MapTo<UserInfoDto>(userInfo);
            Assert.AreEqual(dto.departmentName, userInfo.DepartmentName);
        }

        [Test]
        public void MapToCustmoConfigTest()
        {
            var customerUserInfo = new CustomUserInfo
            {
                Pre_UserId = "2018001",
                RealName = "张三",
                departmentid = "001",
                DepartmentName = "技术部",
                Gender = 1,
                AddTime = "2023-1-1",
                Age = 3
            };
            var info = AutoMapperUtil.MapTo<UserInfo>(customerUserInfo, cfg =>
            {
                //Automapper默认只映射public成员，设置private属性也支持映射
                cfg.ShouldMapProperty = p => p.GetMethod.IsPublic || p.SetMethod.IsPrivate;
                //AutoMapper默认映射公共属性和字段，设置忽略字段映射
                cfg.ShouldMapField = p => false;
                //映射时匹配前缀、后缀（RecognizePostfixes)
                cfg.RecognizePrefixes("Pre_");

                cfg.CreateMap<CustomUserInfo, UserInfo>()
                //指定字段的映射关系
                .ForMember(dest => dest.Sex, opt => opt.MapFrom(src => src.Gender))
                //忽略映射指定字段字段
                .ForMember(dest => dest.DepartmentName, opt => opt.Ignore())
                //针对条件映射
                .ForMember(dest => dest.Age, opt => opt.Condition(src => src.Age > 5))
                //针对指定映射字段做特殊处理
                .ForMember(dest => dest.CreateTime, opt => opt.MapFrom(src => src.AddTime.As<DateTime>().AddYears(1)));
            });
            Assert.IsNull(info.RealName);
            Assert.AreEqual(customerUserInfo.departmentid, info.DepartmentId);
            Assert.AreEqual(customerUserInfo.Pre_UserId, info.UserId);
            Assert.AreEqual(customerUserInfo.Gender, info.Sex);
            Assert.IsNull(info.DepartmentName);

        }
        [Test]
        public void MergeToTest()
        {
            var userInfo = new UserInfo
            {
                UserId = "2018001",
                DepartmentId = "001",
                DepartmentName = "技术部",
                Sex = 1,
                UserOnLine = 1,
                CreateTime = DateTime.Now
            };
            var dto = new UserInfoDto
            {
                UID = 11,
                RealName = "大师兄",
                UserId = "2022",
            };
            dto = AutoMapperUtil.MergeTo<UserInfo, UserInfoDto>(userInfo, dto);
            Assert.AreEqual(dto.UID, 11);
            Assert.AreEqual(dto.departmentName, userInfo.DepartmentName);
        }

        [Test]
        public void MapToListTest()
        {
            var userInfo = new UserInfo
            {
                UserId = "2018001",
                DepartmentId = "001",
                DepartmentName = "技术部",
                Sex = 1,
                UserOnLine = 1,
                CreateTime = DateTime.Now
            };
            var list = (new int[10]).Select(r => new UserInfo
            {
                UserId = "2018001",
                DepartmentId = "001" + r,
                DepartmentName = "技术部",
                Sex = 1,
                UserOnLine = 1,
                CreateTime = DateTime.Now
            }).ToList();
            var dtoList = AutoMapperUtil.MapToList<UserInfoDto>(list);
            Assert.AreEqual(10, dtoList.Count);
        }
    }

    class UserInfo
    {
        public string UserId { get; set; }
        public string RealName { get; set; }
        public string DepartmentId { get; set; }
        public string DepartmentName { get; set; }
        public int UserOnLine { get; set; }

        public DateTime CreateTime { get; set; }

        public int Sex { get; set; }
        public int Age { get; set; }
    }

    class UserInfoDto
    {
        public int UID { get; set; }
        //public int Age { get; set; }
        //public string Ex { get; set; }
        //public DateTime AddTime { get; set; }

        public string UserId { get; set; }
        public string RealName { get; set; }
        public string DepartmentId { get; set; }
        public string departmentName { get; set; }
        public int UserOnLine { get; set; }
    }

    class CustomUserInfo
    {
        public string Pre_UserId { get; set; }
        public string DepartmentName { get; set; }
        public string RealName;
        public string departmentid { get; set; }
        public int Gender { get; set; }
        private int UserOnLine { get { return 1; } set { UserOnLine = value; } }
        public string AddTime { get; set; }
        public int Age { get; set; }

    }

    class TestInfo
    {
        public string Month { get; set; }
        public decimal Amount { get; set; }
    }
}
