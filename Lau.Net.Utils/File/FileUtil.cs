using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Lau.Net.Utils
{
    public class FileUtil
    {
        #region 创建文件夹
        /// <summary>
        /// 创建文件夹
        /// D:\aaa\bbb\kk.txt 将会创建 D:\aaa\bbb\ 文件夹；
        /// D:\aaa\bbb\ 将会创建 D:\aaa\bbb\ 文件夹；
        /// D:\aaa\bbb 将会创建 D:\aaa\ 文件夹；
        /// </summary>
        /// <param name="dirPath"></param>
        public static void CreateDirectory(string dirPath)
        {
            dirPath = Path.GetDirectoryName(dirPath);
            if (!Directory.Exists(dirPath))
                Directory.CreateDirectory(dirPath);
        }
        #endregion

        /// <summary>
        /// 判断指定文件夹下是否存在文件或文件夹
        /// </summary>
        /// <param name="dirPath">文件夹路径</param>
        /// <returns>如果指定文件夹下存在文件或文件夹,则返回True;否则返回False</returns>
        public static bool ExistsFileOrDir(string dirPath)
        {
            bool result = false;
            try
            {
                result = Directory.GetFiles(dirPath).Length > 0;
                result = result || (Directory.GetDirectories(dirPath).Length > 0);
            }
            catch (Exception)
            {
                result = false;
            }
            return result;
        }

        /// <summary>
        /// 创建文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="createNew">如果文件存在是否创建新文件</param>
        public static void CreateFile(string filePath, bool createNew)
        {
            if (createNew || !File.Exists(filePath))
            {
                CreateDirectory(filePath);
                File.Create(filePath);
            }
        }

        #region 文件夹复制
        /// <summary>
        /// 文件夹复制
        /// </summary>
        /// <param name="sourceDirPath">源文件夹路径</param>
        /// <param name="targetDirPath">目标文件夹路径</param>
        /// <param name="overWrite">设置为True表示存在相同的文件名时，覆盖之前的文件</param>
        public static void CopyDirectory(string sourceDirPath, string targetDirPath, bool overWrite = true)
        {
            if (!Directory.Exists(targetDirPath))
            {
                Directory.CreateDirectory(targetDirPath);
            }
            //复制文件
            foreach (string fileName in Directory.GetFiles(sourceDirPath))
            {
                File.Copy(fileName, Path.Combine(targetDirPath, Path.GetFileName(fileName)), overWrite);
            }
            //递归复制
            foreach (string dir in Directory.GetDirectories(sourceDirPath))
            {
                CopyDirectory(dir, Path.Combine(targetDirPath, Path.GetFileName(dir)));
            }
        }
        #endregion

        /// <summary>
        /// 把字符串保存成文件
        /// </summary>
        /// <param name="filePath">文件路径（含文件名）</param>
        /// <param name="content">将保存的字符串</param>
        /// <returns>保存成功返回True,否则返回False</returns>
        public static bool SaveFile(string filePath, string content, Encoding encoding = null)
        {
            if (encoding == null) encoding = Encoding.Default;
            string message = string.Empty;
            string newFilePath = filePath;
            string newFileName, fileName = Path.GetFileName(filePath);
            int i = 2;
            while (File.Exists(newFilePath))
            {
                newFilePath = filePath.Insert(filePath.LastIndexOf('.'), "(" + i.ToString() + ")");
                i++;
            }
            newFileName = Path.GetFileName(newFilePath);

            try
            {
                CreateDirectory(filePath);
                if (string.Equals(filePath, newFilePath) == false)
                {
                    message = "文件\"" + fileName + "\"已存在，是否重命名为\"" + newFileName + "\"?";
                    DialogResult result = MsgBox.ShowYesNoCancelInformation(message);
                    if (result == DialogResult.Yes)
                    {
                        File.WriteAllText(newFilePath, content, encoding);
                    }
                    else if (result == DialogResult.No)
                    {
                        File.WriteAllText(filePath, content, encoding);
                    }
                }
                else
                {
                    File.WriteAllText(filePath, content, encoding);
                }
                return true;
            }
            catch (Exception ex)
            {
                MsgBox.ShowError(ex.Message);
                return false;
            }

        }

        /// <summary>
        /// 读取文件文本内容
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="encoding">读取编码</param>
        /// <returns>返回文件文本内容</returns>
        public static string ReadText(string filePath, Encoding encoding)
        {
            using (StreamReader reader = new StreamReader(filePath, encoding))
            {
                return reader.ReadToEnd();
            }
        }

        /// <summary>
        /// 取得文件夹下指定类型的文件
        /// </summary>
        /// <param name="searchFolder">目录文件夹</param>
        /// <param name="filters">文件扩展名数组</param>
        /// <param name="isRecursive">是否循环搜索子文件夹</param>
        /// <returns></returns>
        public static String[] GetFilesFrom(String searchFolder, String[] filters, bool isRecursive = false)
        {
            var filesFound = new List<String>();
            var searchOption = isRecursive ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
            foreach (var filter in filters)
            {
                filesFound.AddRange(Directory.GetFiles(searchFolder, $"*.{filter.Trim()}", searchOption));
            }
            return filesFound.ToArray();
        }
    }
}
