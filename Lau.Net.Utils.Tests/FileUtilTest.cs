using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils.Tests
{
    [TestFixture]
    public class FileUtilTest
    {
        [Test]
        public async Task Test()
        {
            //var minioUtil = new MinioUtil("192.168.1.4:9091", "minio2", "minio@123");
            //var list = await minioUtil.GetAllBucketsAsync();

            var minioUtil = new MinioUtil("192.168.1.3:9000", "minioadmin", "minioadmin");
            var filePath = @"D:\Pictures\Screenshots\1.xlsx";
            using (var stream = File.OpenRead(filePath))
            {
                await minioUtil.UploadObjectAsync("test2","excels/"+ Path.GetFileName(filePath), stream);
            } 
            //await minioUtil.UploadObjectAsync("test2", "imya.png", filePath);
            //var objs = await minioUtil.ListObjectsAsync("test2");
            //await minioUtil.DeleteObjectAsync("test2", "imya7.png");
            //var list = await minioUtil.GetAllBucketsAsync();
            //await minioUtil.CreateBucketAsync("test2");
            //var result = await minioUtil.BucketExistsAsync("test");

        }
    }
}
