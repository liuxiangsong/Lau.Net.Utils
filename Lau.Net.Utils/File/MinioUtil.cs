using Minio;
using Minio.ApiEndpoints;
using Minio.DataModel;
using Minio.DataModel.Args;
using Minio.DataModel.Result;
using Minio.Exceptions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lau.Net.Utils
{
    /// <summary>
    /// Minio 帮助类
    /// </summary>
    public class MinioUtil
    {
        private IMinioClient _minioClient;
        /// <summary>
        /// 初始化MinioClient
        /// </summary>
        /// <param name="endpoint">Minio地址，如：192.168.1.3:9000，不能带有协议如https等,必须是api地址，而不webUI地址</param>
        /// <param name="accessKey">Minio账号</param>
        /// <param name="secretKey">Minio密码</param>
        /// <param name="secure">是否以https方式连接</param> 
        public MinioUtil(string endpoint, string accessKey, string secretKey, bool secure = false)
        {
            _minioClient = new MinioClient()
                .WithEndpoint(endpoint)
                .WithCredentials(accessKey, secretKey)
                .WithSSL(secure)
                .Build();
        }

        #region 检查桶是否存在，及创建桶等方法
        /// <summary>
        /// 检查存储桶是否存在
        /// </summary>
        /// <param name="bucketName">桶名称</param>
        /// <returns></returns>
        public async Task<bool> BucketExistsAsync(string bucketName)
        {
            try
            {
                var beArgs = new BucketExistsArgs()
                    .WithBucket(bucketName);
                var result = await _minioClient.BucketExistsAsync(beArgs).ConfigureAwait(false);
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking bucket existence: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 创建一个新的存储桶
        /// </summary>
        /// <param name="bucketName">桶名称</param>
        /// <returns></returns>
        public async Task<bool> CreateBucketAsync(string bucketName)
        {
            try
            {
                await _minioClient.MakeBucketAsync(new MakeBucketArgs().WithBucket(bucketName));
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating bucket: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 检查桶是否存在，不存在则创建
        /// </summary>
        /// <param name="bucketName"></param>
        /// <returns></returns>
        public async Task EnsureBucketExistsAsync(string bucketName)
        {
            bool bucketExists = await BucketExistsAsync(bucketName);
            if (!bucketExists)
            {
                // 如果不存在，则创建桶
                await CreateBucketAsync(bucketName);
            }
        }
        #endregion

        #region 获取所有桶名称、获取指定桶中的所有对象名称
        /// <summary>
        /// 获取所有的桶名称
        /// </summary>
        /// <returns></returns>
        public async Task<List<string>> GetAllBucketsAsync()
        {
            try
            {
                // Retrieve the list of buckets
                ListAllMyBucketsResult bucketList = await _minioClient.ListBucketsAsync().ConfigureAwait(false);
                return bucketList.Buckets.Select(b => b.Name).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取指定桶中的所有对象名称
        /// </summary>
        /// <param name="bucketName">桶名称</param>
        /// <returns>返回桶中所有对象的名称</returns>
        public async Task<List<string>> ListObjectsAsync(string bucketName)
        {
            var objects = new List<Item>();
            try
            {
                var listArgs = new ListObjectsArgs()
                                .WithBucket(bucketName);
                var objectsEnumerable = _minioClient.ListObjectsEnumAsync(listArgs);
                var objectsEnumerator = objectsEnumerable.GetAsyncEnumerator();

                while (await objectsEnumerator.MoveNextAsync())
                {
                    objects.Add(objectsEnumerator.Current);
                }
                return objects.Select(o => o.Key).ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing objects: {ex.Message}");
                return null;
            }
        }
        #endregion

        #region 上传、下载、删除桶中的对象
        /// <summary>
        /// 上传对象到指定存储桶,如果对象名称已存在，则直接覆盖文件
        /// </summary>
        /// <param name="bucketName">桶名称</param>
        /// <param name="objectName">保存的对象名称</param>
        /// <param name="filePath">对象文件路径</param>
        /// <returns></returns>
        public async Task<bool> UploadObjectAsync(string bucketName, string objectName, string filePath)
        {
            try
            {
                await EnsureBucketExistsAsync(bucketName).ConfigureAwait(false);
                var putObjectArgs = new PutObjectArgs()
                    .WithBucket(bucketName)
                    .WithObject(objectName)
                    .WithFileName(filePath);
                await _minioClient.PutObjectAsync(putObjectArgs).ConfigureAwait(false);
                Console.WriteLine($"Object '{objectName}' uploaded to bucket '{bucketName}' successfully.");
                return true;
            }
            catch (MinioException ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 上传对象到指定存储桶,如果对象名称已存在，则直接覆盖文件
        /// </summary>
        /// <param name="bucketName">桶名称</param>
        /// <param name="objectName">保存的对象名称</param>
        /// <param name="fileStream">文件流</param>
        /// <returns></returns>
        public async Task<bool> UploadObjectAsync(string bucketName, string objectName, System.IO.Stream fileStream)
        {
            try
            {
                await EnsureBucketExistsAsync(bucketName);

                // 准备上传参数
                var putObjectArgs = new PutObjectArgs()
                 .WithBucket(bucketName)
                    .WithObject(objectName)
                    .WithStreamData(fileStream)
                    .WithObjectSize(fileStream.Length);
                //.WithContentType( "application/octet-stream");

                // 使用PutObjectAsync方法上传文件流
                await _minioClient.PutObjectAsync(putObjectArgs);
                Console.WriteLine($"File uploaded successfully to bucket {bucketName} with object name {objectName}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uploading file: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 从指定存储桶下载对象
        /// </summary>
        /// <param name="bucketName">桶名称</param>
        /// <param name="objectName">下载的对象名称</param>
        /// <param name="filePath">保存文件路径</param>
        /// <returns></returns>
        public async Task DownloadObjectAsync(string bucketName, string objectName, string filePath)
        {
            try
            {
                await _minioClient.GetObjectAsync(new GetObjectArgs()
                    .WithBucket(bucketName)
                    .WithObject(objectName)
                    .WithFile(filePath)).ConfigureAwait(false);
                Console.WriteLine($"Object '{objectName}' downloaded from bucket '{bucketName}' successfully.");
            }
            catch (MinioException ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }
        }
        /// <summary>
        /// 从指定存储桶删除对象
        /// </summary>
        /// <param name="bucketName">桶名称</param>
        /// <param name="objectName">对象名称</param>
        /// <returns></returns>
        public async Task DeleteObjectAsync(string bucketName, string objectName)
        {
            try
            {
                await _minioClient.RemoveObjectAsync(new RemoveObjectArgs()
                    .WithBucket(bucketName)
                    .WithObject(objectName)).ConfigureAwait(false);
                Console.WriteLine($"Object '{objectName}' deleted from bucket '{bucketName}' successfully.");
            }
            catch (MinioException ex)
            {
                Console.WriteLine($"Error occurred: {ex.Message}");
            }
        }
        #endregion

    }
}
