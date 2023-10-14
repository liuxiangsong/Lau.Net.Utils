﻿using System;
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
        /// <param name="saveFilePath">压缩包保存路径</param>
        /// <param name="deleteSourceDirectory">压缩后是否删除目标文件夹</param>
        public static void Compress(string sourceDirectory, string saveFilePath, bool deleteSourceDirectory = false)
        {
            ZipFile.CreateFromDirectory(sourceDirectory, saveFilePath);
            if (deleteSourceDirectory)
            {
                Directory.Delete(sourceDirectory, true);
            }
        }

        /// <summary>
        /// 解压
        /// </summary>
        /// <param name="zipFilePath">压缩包文件路径</param>
        /// <param name="saveDirectory">解压到文件夹路径</param>
        /// <param name="deleteZipFile">解压后是否删除压缩包</param>
        public static void Decompress(string zipFilePath, string saveDirectory, bool deleteZipFile = false)
        {
            ZipFile.ExtractToDirectory(zipFilePath, saveDirectory);
            if (deleteZipFile)
            {
                File.Delete(zipFilePath);
            }
        }

        /// <summary>
        /// 添加文件至压缩包
        /// </summary>
        /// <param name="zipFilePath">压缩包路径</param>
        /// <param name="addFilePath">添加文件路径</param>
        public static void AddFileToZip(string zipFilePath, string addFilePath)
        {
            using (FileStream fs = new FileStream(zipFilePath, FileMode.Open))
            {
                using (var archive = new ZipArchive(fs, ZipArchiveMode.Update))
                {
                    var fileName = Path.GetFileName(addFilePath);
                    archive.CreateEntryFromFile(addFilePath, fileName);
                }
            }
        } 
        #endregion


    }
}
