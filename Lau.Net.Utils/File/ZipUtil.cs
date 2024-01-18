using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

//using ICSharpCode.SharpZipLib.Core;
//using ICSharpCode.SharpZipLib.Zip;

namespace Lau.Net.Utils
{
    /// <summary>
    /// 
    /// </summary>
    public class ZipUtil
    {
        //public void ExtractZipFile(string archiveFilenameIn, string outFolder, string password)
        //{
        //    ZipFile zf = null;
        //    try
        //    {
        //        FileStream fs = File.OpenRead(archiveFilenameIn);
        //        zf = new ZipFile(fs);
        //        if (!String.IsNullOrEmpty(password))
        //        {
        //            zf.Password = password;		// AES encrypted entries are handled automatically
        //        }
        //        foreach (ZipEntry zipEntry in zf)
        //        {
        //            if (!zipEntry.IsFile)
        //            {
        //                continue;			// Ignore directories
        //            }
        //            String entryFileName = zipEntry.Name;
        //            // to remove the folder from the entry:- entryFileName = Path.GetFileName(entryFileName);
        //            // Optionally match entrynames against a selection list here to skip as desired.
        //            // The unpacked length is available in the zipEntry.Size property.

        //            byte[] buffer = new byte[4096];		// 4K is optimum
        //            Stream zipStream = zf.GetInputStream(zipEntry);

        //            // Manipulate the output filename here as desired.
        //            String fullZipToPath = Path.Combine(outFolder, entryFileName);
        //            string directoryName = Path.GetDirectoryName(fullZipToPath);
        //            if (directoryName.Length > 0)
        //            {
        //                Directory.CreateDirectory(directoryName);
        //            }

        //            // Unzip file in buffered chunks. This is just as fast as unpacking to a buffer the full size
        //            // of the file, but does not waste memory.
        //            // The "using" will close the stream even if an exception occurs.
        //            using (FileStream streamWriter = File.Create(fullZipToPath))
        //            {
        //                StreamUtils.Copy(zipStream, streamWriter, buffer);
        //            }
        //        }
        //    }
        //    finally
        //    {
        //        if (zf != null)
        //        {
        //            zf.IsStreamOwner = true; // Makes close also shut the underlying stream
        //            zf.Close(); // Ensure we release resources
        //        }
        //    }
        //}

        #region System.IO.Compression 压缩、解压相关方法
        /// <summary>
        /// 将文件夹打包成一个压缩包
        /// </summary>
        /// <param name="sourceDirectory">目标文件夹</param>
        /// <param name="zipFilePath">压缩包保存路径</param>
        /// <param name="deleteSourceDirectory">压缩后是否删除目标文件夹</param>
        /// <returns>返回空则表示添加成功，否则返回错误提示语</returns>
        public static string Compress(string sourceDirectory, string zipFilePath, bool deleteSourceDirectory = false)
        {
            if (!Directory.Exists(sourceDirectory))
            {
                return "目标文件夹不存在";
            }
            ZipFile.CreateFromDirectory(sourceDirectory, zipFilePath);
            if (deleteSourceDirectory)
            {
                Directory.Delete(sourceDirectory, true);
            }
            return string.Empty;
        }

        /// <summary>
        /// 将多个文件进行压缩
        /// </summary>
        /// <param name="filePaths">需要压缩的文件</param>
        /// <param name="zipFilePath">压缩文件路径，如果传空，则默认保存到第一文件的文件夹下，并以第一个文件作为文件名</param>
        /// <param name="deleteSourceFiles">是否删除被压缩的源文件</param>
        /// <returns>返回压缩文件夹路径</returns>
        public static string Compress(IEnumerable<string> filePaths, string zipFilePath ="", bool deleteSourceFiles = false)
        {
            var deleteFilesFunc = new Action(() =>
            {
                if (!deleteSourceFiles)
                {
                    return;
                }
                foreach (var path in filePaths)
                {
                    File.Delete(path);
                }
            });
            if (string.IsNullOrEmpty(zipFilePath))
            {
                var firstFilePath = filePaths.FirstOrDefault();
                var zipFileName = Path.GetFileNameWithoutExtension(firstFilePath) + ".zip";
                zipFilePath = Path.Combine(Path.GetDirectoryName(firstFilePath), zipFileName); 
            }
            if (File.Exists(zipFilePath))
            {
                foreach(var path in filePaths)
                {
                    AddFileToZip(zipFilePath, path);
                }
                deleteFilesFunc();
                return zipFilePath;
            }

            // 创建一个新的压缩文件
            using (ZipArchive archive = ZipFile.Open(zipFilePath, ZipArchiveMode.Create))
            {
                foreach (string filePath in filePaths)
                {
                    if (File.Exists(filePath))
                    {
                        // 将每个文件添加到压缩文件中
                        string fileName = Path.GetFileName(filePath);
                        archive.CreateEntryFromFile(filePath, fileName);
                    } 
                }
            }
            deleteFilesFunc();
            return zipFilePath;
        }

        /// <summary>
        /// 解压
        /// </summary>
        /// <param name="zipFilePath">压缩包文件路径</param>
        /// <param name="saveDirectory">解压到文件夹路径</param>
        /// <param name="deleteZipFile">解压后是否删除压缩包</param>
        /// <returns>返回空则表示添加成功，否则返回错误提示语</returns>
        public static string Decompress(string zipFilePath, string saveDirectory, bool deleteZipFile = false)
        {
            if (!File.Exists(zipFilePath))
            {
                return "压缩包文件不存在";
            }
            ZipFile.ExtractToDirectory(zipFilePath, saveDirectory);
            if (deleteZipFile)
            {
                File.Delete(zipFilePath);
            }
            return string.Empty;
        }

        /// <summary>
        /// 添加文件至压缩包
        /// </summary>
        /// <param name="zipFilePath">压缩包路径</param>
        /// <param name="addFilePath">添加文件路径</param>
        /// <returns>返回空则表示添加成功，否则返回错误提示语</returns>
        public static string AddFileToZip(string zipFilePath, string addFilePath)
        {
            if (!File.Exists(zipFilePath))
            {
                return "压缩包文件不存在";
            }
            if (!File.Exists(addFilePath))
            {
                return "添加文件不存在";
            }
            using (FileStream fs = new FileStream(zipFilePath, FileMode.Open))
            {
                using (var archive = new ZipArchive(fs, ZipArchiveMode.Update))
                {
                    var fileName = Path.GetFileName(addFilePath);
                    archive.CreateEntryFromFile(addFilePath, fileName);
                }
            }
            return string.Empty;
        }
        #endregion


    }
}
