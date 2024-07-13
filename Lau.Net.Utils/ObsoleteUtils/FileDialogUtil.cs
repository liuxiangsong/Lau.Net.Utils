//using System.Drawing;

//namespace Lau.Net.Utils
//{
//    ///<summary>
//    ///打开、保存文件对话框操作辅助类
//    ///</summary>
//    public class FileDialogUtil
//    {
//        private static string ExcelFilter = string.Format("Excel{0}|{0}|All File({1})|{1}", "*.xlsx;*.xls;", "*.*");
//        private static string ImageFilter = string.Format("Image Files{0}|{0}|All File(*.*)|*.*", "*.bmp;*.jpg;*.gif;*.png;*.ico;");
//        private static string HtmlFilter = "HTML files (*.html;*.htm)|*.html;*.htm|All files (*.*)|*.*";
//        private static string AccessFilter = "Access(*.mdb)|*.mdb|All File(*.*)|*.*";
//        private static string ZipFillter = "Zip(*.zip)|*.zip|All files (*.*)|*.*";
//        private const string ConfigFilter = "配置文件(*.cfg)|*.cfg|All File(*.*)|*.*";
//        private static string TxtFilter = "(*.txt)|*.txt|All files (*.*)|*.*";

//        #region Txt相关对话框
//        /// <summary>
//        /// 打开Txt对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文本的路径</returns>
//        public static string OpenText()
//        {
//            return OpenFile("文本文件选择", TxtFilter);
//        }

//        /// <summary>
//        /// 保存Txt对话框,并返回保存文件的路径
//        /// </summary>
//        /// <returns>返回选中文本的路径</returns>
//        public static string SaveText()
//        {
//            return SaveText(string.Empty);
//        }

//        /// <summary>
//        /// 保存Txt对话框,并返回保存文件的路径
//        /// </summary>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <returns>返回选中文本的路径</returns>
//        public static string SaveText(string fileName)
//        {
//            return SaveFile("保存文本文件", TxtFilter, fileName);
//        }

//        /// <summary>
//        /// 保存Txt对话框,并返回保存文件的路径
//        /// </summary> 
//        /// <param name="fileName">保存文件的默认文件名</param>
//        /// <param name="initialDirectory">对话框显示的初始目录</param>
//        /// <returns>返回保存文件的路径</returns>
//        public static string SaveText(string fileName, string initialDirectory)
//        {
//            return SaveFile("保存文本文件", TxtFilter, fileName, initialDirectory);
//        }

//        #endregion

//        #region Excel相关对话框
//        /// <summary>
//        /// 打开Excel对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文件的路径</returns>
//        public static string OpenExcel()
//        {
//            return OpenFile("Excel选择", ExcelFilter);
//        }

//        /// <summary>
//        /// 保存Excel对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveExcel()
//        {
//            return SaveExcel(string.Empty);
//        }

//        /// <summary>
//        /// 保存Excel对话框,并返回选中文件的路径
//        /// </summary>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveExcel(string fileName)
//        {
//            return SaveFile("保存Excel", ExcelFilter, fileName);
//        }

//        /// <summary>
//        /// 保存Excel对话框,并返回选中文件的路径
//        /// </summary>
//        /// <param name="fileName">保存文件的默认文件名</param>
//        /// <param name="initialDirectory">对话框显示的初始目录</param>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveExcel(string fileName, string initialDirectory)
//        {
//            return SaveFile("保存Excel", ExcelFilter, fileName, initialDirectory);
//        }

//        #endregion

//        #region HTML相关对话框

//        /// <summary>
//        /// 打开Html对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文件的路径</returns>
//        public static string OpenHtml()
//        {
//            return OpenFile("Html选择", HtmlFilter);
//        }

//        /// <summary>
//        /// 保存Html对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveHtml()
//        {
//            return SaveHtml(string.Empty);
//        }

//        /// <summary>
//        /// 保存Html对话框,并返回选中文件的路径
//        /// </summary>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveHtml(string fileName)
//        {
//            return SaveFile("保存Html", HtmlFilter, fileName);
//        }

//        /// <summary>
//        /// 保存Html对话框,并返回选中文件的路径
//        /// </summary>
//        /// <param name="fileName">保存文件的默认文件名</param>
//        /// <param name="initialDirectory">对话框显示的初始目录</param>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveHtml(string fileName, string initialDirectory)
//        {
//            return SaveFile("保存Html", HtmlFilter, fileName, initialDirectory);
//        }

//        #endregion

//        #region 压缩文件相关
//        /// <summary>
//        /// 打开Zip对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文件的路径</returns>
//        public static string OpenZip()
//        {
//            return OpenFile("压缩文件选择", ZipFillter);
//        }

//        /// <summary>
//        /// 打开Zip对话框,并返回选中文件的路径
//        /// </summary>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <returns>返回选中文件的路径</returns>
//        public static string OpenZip(string fileName)
//        {
//            return OpenFile("压缩文件选择", ZipFillter, fileName);
//        }

//        /// <summary>
//        /// 保存Zip对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveZip()
//        {
//            return SaveZip(string.Empty);
//        }

//        /// <summary>
//        /// 保存Zip对话框,并返回选中文件的路径
//        /// </summary>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveZip(string fileName)
//        {
//            return SaveFile("压缩文件保存", ZipFillter, fileName);
//        }

//        /// <summary>
//        /// 保存Zip对话框,并返回选中文件的路径
//        /// </summary>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <param name="initialDirectory">对话框显示的初始目录</param>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveZip(string fileName, string initialDirectory)
//        {
//            return SaveFile("压缩文件保存", ZipFillter, fileName, initialDirectory);
//        }

//        #endregion

//        #region 图片相关
//        /// <summary>
//        /// 打开图片对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文件的路径</returns>
//        public static string OpenImage()
//        {
//            return OpenFile("图片选择", ImageFilter);
//        }


//        /// <summary>
//        /// 保存图片对话框,并返回选中文件的路径
//        /// </summary>
//        /// <returns>返回选中文件的路径</returns>
//        public static string SaveImage()
//        {
//            return SaveImage(string.Empty);
//        }

//        /// <summary>
//        /// 保存图片对话框并设置默认文件名,并返回保存全路径
//        /// </summary>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <returns>返回保存全路径</returns>
//        public static string SaveImage(string fileName)
//        {
//            return SaveFile("保存图片", ImageFilter, fileName);
//        }

//        /// <summary>
//        /// 保存图片对话框并设置默认文件名,并返回保存全路径
//        /// </summary>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <param name="initialDirectory">对话框显示的初始目录</param>
//        /// <returns>返回保存全路径</returns>
//        public static string SaveImage(string fileName, string initialDirectory)
//        {
//            return SaveFile("保存图片", ImageFilter, fileName, initialDirectory);
//        }

//        #endregion

//        #region 数据库备份还原

//        /// <summary>
//        /// 保存数据库备份对话框
//        /// </summary>
//        /// <returns>数据库备份路径</returns>
//        public static string SaveAccessDb()
//        {
//            return SaveFile("数据库备份", AccessFilter);
//        }

//        public static string SaveBakDb()
//        {
//            return SaveFile("数据库备份", "Access(*.bak)|*.bak");
//        }


//        public static string OpenBakDb(string file)
//        {
//            return OpenFile("数据库还原", "Access(*.bak)|*.bak", file);
//        }

//        /// <summary>
//        /// 数据库还原对话框
//        /// </summary>
//        /// <returns>数据库还原路径</returns>
//        public static string OpenAccessDb()
//        {
//            return OpenFile("数据库还原", AccessFilter);
//        }
//        #endregion

//        #region 配置文件
//        /// <summary>
//        /// 保存配置文件备份对话框
//        /// </summary>
//        /// <returns>配置文件备份路径</returns>
//        public static string SaveConfig()
//        {
//            return SaveFile("配置文件备份", ConfigFilter);
//        }

//        /// <summary>
//        /// 配置文件还原对话框
//        /// </summary>
//        /// <returns>配置文件还原路径</returns>
//        public static string OpenConfig()
//        {
//            return OpenFile("配置文件还原", ConfigFilter);
//        }
//        #endregion

//        #region 通用函数

//        #region FolderBrowserDialog
//        /// <summary>
//        /// 选择文件夹，并返回选中文件夹的路径
//        /// </summary>
//        /// <returns>返回选择文件夹的路径</returns>
//        public static string OpenDir()
//        {
//            return OpenDir(string.Empty);
//        }

//        /// <summary>
//        /// 选择文件夹，并返回选中文件夹的路径
//        /// </summary>
//        /// <param name="selectedPath">默认选择路径</param>
//        /// <returns>返回选择文件夹的路径</returns>
//        public static string OpenDir(string selectedPath)
//        {
//            FolderBrowserDialog dialog = new FolderBrowserDialog();
//            dialog.Description = "请选择路径";
//            dialog.SelectedPath = selectedPath;
//            if (dialog.ShowDialog() == DialogResult.OK)
//            {
//                return dialog.SelectedPath;
//            }
//            else
//            {
//                return string.Empty;
//            }
//        }
//        #endregion

//        #region OpenFileDialog
//        /// <summary>
//        /// 选择文件,并返回选择文件的路径
//        /// </summary>
//        /// <param name="title">对话框标题</param>
//        /// <param name="filter">文件类型和说明说明的筛选器字符串</param>
//        /// <returns>返回选择文件的路径</returns>
//        public static string OpenFile(string title, string filter)
//        {
//            return OpenFile(title, filter, string.Empty);
//        }

//        /// <summary>
//        ///  选择文件,并返回选择文件的路径
//        /// </summary>
//        /// <param name="title">对话框的标题</param>
//        /// <param name="filter">文件类型和说明说明的筛选器字符串</param>
//        /// <param name="fileName">对话框中选择的文件名</param>
//        /// <returns>返回选择文件的路径</returns>
//        public static string OpenFile(string title, string filter, string fileName)
//        {
//            OpenFileDialog dialog = new OpenFileDialog();
//            dialog.Filter = filter;
//            dialog.Title = title;
//            dialog.RestoreDirectory = true;
//            dialog.FileName = fileName;
//            if (dialog.ShowDialog() == DialogResult.OK)
//            {
//                return dialog.FileName;
//            }
//            else
//            {
//                return string.Empty;
//            }
//        }

//        #endregion

//        #region SaveFileDialog
//        /// <summary>
//        /// 保存文件,并返回文件保存的路径
//        /// </summary>
//        /// <param name="tile">对话框的标题</param>
//        /// <param name="filter">文件类型和说明说明的筛选器字符串</param>
//        /// <returns>返回文件保存的路径</returns>
//        public static string SaveFile(string title, string filter)
//        {
//            return SaveFile(title, filter, string.Empty);
//        }

//        /// <summary>
//        /// 保存文件,并返回文件保存的路径
//        /// </summary>
//        /// <param name="tile">对话框的标题</param>
//        /// <param name="filter">文件类型和说明说明的筛选器字符串</param>
//        /// <param name="fileName">保存文件的默认文件名</param>
//        /// <returns>返回文件保存的路径</returns>
//        public static string SaveFile(string title, string filter, string fileName)
//        {
//            return SaveFile(title, filter, fileName, "");
//        }

//        /// <summary>
//        /// 保存文件,并返回文件保存的路径
//        /// </summary>
//        /// <param name="tile">对话框的标题</param>
//        /// <param name="filter">文件类型和说明说明的筛选器字符串</param>
//        /// <param name="fileName">保存文件的默认文件名</param>
//        /// <param name="initialDirectory">对话框显示的初始目录</param>
//        /// <returns>返回文件保存的路径</returns>
//        public static string SaveFile(string title, string filter, string fileName, string initialDirectory)
//        {
//            SaveFileDialog dialog = new SaveFileDialog();
//            dialog.Filter = filter;
//            dialog.Title = title;
//            dialog.FileName = fileName;
//            dialog.RestoreDirectory = true;
//            if (!string.IsNullOrEmpty(initialDirectory))
//            {
//                dialog.InitialDirectory = initialDirectory;
//            }

//            if (dialog.ShowDialog() == DialogResult.OK)
//            {
//                return dialog.FileName;
//            }
//            return string.Empty;
//        }

//        #endregion
//        #endregion

//        #region 获取颜色对话框的颜色
//        /// <summary>
//        /// 选择颜色对话框，并返回选择的颜色
//        /// </summary>
//        /// <returns>返回选择的颜色</returns>
//        public static Color PickColor()
//        {
//            Color result = SystemColors.Control;
//            ColorDialog form = new ColorDialog();
//            if (DialogResult.OK == form.ShowDialog())
//            {
//                result = form.Color;
//            }
//            return result;
//        }

//        /// <summary>
//        /// 选择颜色对话框，并返回选择的颜色
//        /// </summary>
//        /// <param name="color">默认选择的颜色</param>
//        /// <returns>返回选择的颜色</returns>
//        public static Color PickColor(Color color)
//        {
//            Color result = SystemColors.Control;
//            ColorDialog form = new ColorDialog();
//            form.Color = color;
//            if (DialogResult.OK == form.ShowDialog())
//            {
//                result = form.Color;
//            }
//            return result;
//        }
//        #endregion
//    }
//}
