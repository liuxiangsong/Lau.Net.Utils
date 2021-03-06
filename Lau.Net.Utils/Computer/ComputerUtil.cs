using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using System.Management;
//using System.Runtime.InteropServices;
namespace Lau.Net.Utils
{
    /// <summary>
    /// windows api 名称
    /// </summary>
    public enum WindowsApiType
    {
        /// <summary>
        /// 内存
        /// </summary>
        Win32_PhysicalMemory,
        /// <summary>
        /// cpu
        /// </summary>
        Win32_Processor,
        /// <summary>
        /// 硬盘
        /// </summary>
        win32_DiskDrive,
        /// <summary>
        /// 电脑型号
        /// </summary>
        Win32_ComputerSystemProduct,
        /// <summary>
        /// 分辨率
        /// </summary>
        Win32_DesktopMonitor,
        /// <summary>
        /// 显卡
        /// </summary>
        Win32_VideoController,
        /// <summary>
        /// 操作系统
        /// </summary>
        Win32_OperatingSystem

    }
    public enum WindowsApiKeys
    {
        /// <summary>
        /// 名称
        /// </summary>
        Name,
        /// <summary>
        /// 显卡芯片
        /// </summary>
        VideoProcessor,
        /// <summary>
        /// 显存大小
        /// </summary>
        AdapterRAM,
        /// <summary>
        /// 分辨率宽
        /// </summary>
        ScreenWidth,
        /// <summary>
        /// 分辨率高
        /// </summary>
        ScreenHeight,
        /// <summary>
        /// 电脑型号
        /// </summary>
        Version,
        /// <summary>
        /// 硬盘容量
        /// </summary>
        Size,
        /// <summary>
        /// 内存容量
        /// </summary>
        Capacity,
        /// <summary>
        /// cpu核心数
        /// </summary>
        NumberOfCores
    }
    /// <summary>
    /// 电脑信息类 单例
    /// </summary>
    public class ComputerUtil
    {
        private static ComputerUtil _instance;
        private static readonly object Lock = new object();

        public string Defaultbrowser { private set; get; }
        public string Memory { private set; get; }
        public string Cpuinfo { private set; get; }
        public string Screenresolution { private set; get; }

        private ComputerUtil()
        {
            Defaultbrowser = GetDefaultWebBrowser();
            Memory = GetPhisicalMemory();
            Cpuinfo = GetCpu() + " " + GetCPU_Count();
            Screenresolution = GetFenbianlv();
        }
        public static ComputerUtil CreateComputer()
        {
            if (_instance == null)
            {
                lock (Lock)
                {
                    if (_instance == null)
                    {
                        _instance = new ComputerUtil();
                    }
                }
            }
            return _instance;
        }
        /// <summary>
        /// 查找cpu的名称，主频, 核心数
        /// </summary>
        /// <returns></returns>
        public string GetCpu()
        {
            string result = "";
            try
            {
                var mcCpu = new ManagementClass(WindowsApiType.Win32_Processor.ToString());
                var mocCpu = mcCpu.GetInstances();
                foreach (var m in mocCpu.Cast<ManagementObject>())
                {
                    result = m[WindowsApiKeys.Name.ToString()].ToString();
                    break;
                }
            }
            catch
            {
                // ignored
            }
            return result;
        }
        /// <summary>
        /// 获取系统位数
        /// </summary>
        /// <returns></returns>
        public string GetSystemType()
        {
            try
            {
                string st = "";
                ManagementClass mc = new ManagementClass("Win32_ComputerSystem");
                ManagementObjectCollection moc = mc.GetInstances();
                foreach (var mo in moc.Cast<ManagementObject>())
                {
                    st = mo["SystemType"].ToString();
                }
                return st;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 获取cpu核心数
        /// </summary>
        /// <returns></returns>
        public string GetCPU_Count()
        {
            var str = "";
            try
            {
                int coreCount = 0;
                foreach (var item in new ManagementObjectSearcher("Select * from " +
    WindowsApiType.Win32_Processor.ToString()).Get())
                {
                    coreCount += int.Parse(item[WindowsApiKeys.NumberOfCores.ToString()].ToString());
                }
                if (coreCount == 2)
                {
                    return "双核";
                }
                str = coreCount.ToString() + "核";
            }
            catch
            {
                // ignored
            }
            return str;
        }

        /// <summary>
        /// 获取系统内存大小
        /// </summary>
        /// <returns>内存大小（单位M）</returns>
        public string GetPhisicalMemory()
        {
            string result = "";
            try
            {
                //用于查询一些如系统信息的管理对象
                var searcher = new ManagementObjectSearcher
                {
                    Query = new SelectQuery(WindowsApiType.Win32_PhysicalMemory.ToString(), "",
                        new string[] { WindowsApiKeys.Capacity.ToString() })
                };
                //设置查询条件 
                var collection = searcher.Get();   //获取内存容量 
                var em = collection.GetEnumerator();
                long capacity = 0;
                while (em.MoveNext())
                {
                    ManagementBaseObject baseObj = em.Current;
                    if (baseObj.Properties[WindowsApiKeys.Capacity.ToString()].Value != null)
                    {
                        try
                        {
                            capacity += long.Parse(baseObj.Properties[WindowsApiKeys.Capacity.ToString()].Value.ToString());
                        }
                        catch
                        {
                            result = "";
                        }
                    }
                }
                result = ToGB((double)capacity, 1024.0);
                return result;
            }
            catch (Exception)
            {
                result = "";
                return result;
            }
        }

        /// <summary>
        /// 获取硬盘容量
        /// </summary>
        public string GetDiskSize()
        {
            string result = "";
            StringBuilder sb = new StringBuilder();
            try
            {
                var hardDisk = new ManagementClass(WindowsApiType.win32_DiskDrive.ToString());
                var hardDiskC = hardDisk.GetInstances();
                foreach (var o in hardDiskC)
                {
                    var m = (ManagementObject)o;
                    long capacity = Convert.ToInt64(m[WindowsApiKeys.Size.ToString()].ToString());
                    sb.Append(ToGB(capacity, 1000.0) + "+");
                }
                result = sb.ToString().TrimEnd('+');
            }
            catch
            {
                // ignored
            }
            return result;
        }
        /// <summary>
        /// 电脑型号
        /// </summary>
        public string GetVersion()
        {
            string str = "";
            try
            {
                var hardDisk = new ManagementClass(WindowsApiType.Win32_ComputerSystemProduct.ToString());
                var hardDiskC = hardDisk.GetInstances();
                foreach (var m in hardDiskC.Cast<ManagementObject>())
                {
                    str = m[WindowsApiKeys.Version.ToString()].ToString();
                    break;
                }
            }
            catch
            {
                // ignored
            }
            return str;
        }
        /// <summary>
        /// 获取分辨率
        /// </summary>
        public string GetFenbianlv()
        {
            string result = "";
            try
            {
                result = Desktop.Width.ToString() + "*" + Desktop.Height.ToString();
            }
            catch
            {
                // ignored
            }
            return result;
        }
        /// <summary>
        /// 显卡 芯片,显存大小
        /// </summary>
        public Tuple<string, string> GetVideoController()
        {
            Tuple<string, string> result = null;
            try
            {
                var hardDisk = new ManagementClass(WindowsApiType.Win32_VideoController.ToString());
                var hardDiskC = hardDisk.GetInstances();
                foreach (var o in hardDiskC)
                {
                    var m = (ManagementObject)o;
                    result = new Tuple<string, string>(m[WindowsApiKeys.VideoProcessor.ToString()].ToString()
.Replace("Family", ""), ToGB(Convert.ToInt64(m[WindowsApiKeys.AdapterRAM.ToString()].ToString()), 1024.0));
                    break;
                }
            }
            catch
            {
                // ignored
            }
            return result;
        }
        public string GetDefaultWebBrowser()
        {
            try
            {
                const string browserKey1 = @"Software\Clients\StartmenuInternet\";
                RegistryKey registryKey = Registry.CurrentUser.OpenSubKey(browserKey1, false) ??
                                            Registry.LocalMachine.OpenSubKey(browserKey1, false);
                if (registryKey != null)
                {
                    var result = registryKey.GetValue("") + "";
                    registryKey.Close();
                    return result;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception)
            {
                return "";
            }
        }
        /// <summary>
        /// 操作系统版本
        /// </summary>
        public string GetOS_Version()
        {
            string str = "";
            try
            {
                var hardDisk = new ManagementClass(WindowsApiType.Win32_OperatingSystem.ToString());
                var hardDiskC = hardDisk.GetInstances();
                foreach (var o in hardDiskC)
                {
                    var m = (ManagementObject)o;
                    str = m[WindowsApiKeys.Name.ToString()].ToString().Split('|')[0].Replace("Microsoft", "");
                    break;
                }
            }
            catch
            {
                // ignored
            }
            return str;
        }
        /// <summary>  
        /// 将字节转换为GB
        /// </summary>  
        /// <param name="size">字节值</param>  
        /// <param name="mod">除数，硬盘除以1000，内存除以1024</param>  
        /// <returns></returns>  
        public static string ToGB(double size, double mod)
        {
            string[] units = new string[] { "B", "KB", "MB", "GB", "TB", "PB" };
            int i = 0;
            while (size >= mod)
            {
                size /= mod;
                i++;
            }
            return Math.Round(size) + units[i];
        }

        #region Win32 API
        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr ptr);
        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        [DllImport("user32.dll", EntryPoint = "ReleaseDC")]
        static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDc);
        #endregion
        #region DeviceCaps常量
        const int HORZRES = 8;
        const int VERTRES = 10;
        const int LOGPIXELSX = 88;
        const int LOGPIXELSY = 90;
        const int DESKTOPVERTRES = 117;
        const int DESKTOPHORZRES = 118;
        #endregion

        #region 属性
        /// <summary>  
        /// 获取屏幕分辨率当前物理大小  
        /// </summary>  
        public Size WorkingArea
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                Size size = new Size();
                size.Width = GetDeviceCaps(hdc, HORZRES);
                size.Height = GetDeviceCaps(hdc, VERTRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return size;
            }
        }
        /// <summary>  
        /// 当前系统DPI_X 大小 一般为96  
        /// </summary>  
        public static int DpiX
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
                ReleaseDC(IntPtr.Zero, hdc);
                return dpiX;
            }
        }
        /// <summary>  
        /// 当前系统DPI_Y 大小 一般为96  
        /// </summary>  
        public static int DpiY
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
                ReleaseDC(IntPtr.Zero, hdc);
                return dpiY;
            }
        }
        /// <summary>  
        /// 获取真实设置的桌面分辨率大小  
        /// </summary>  
        public static Size Desktop
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                Size size = new Size();
                size.Width = GetDeviceCaps(hdc, DESKTOPHORZRES);
                size.Height = GetDeviceCaps(hdc, DESKTOPVERTRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return size;
            }
        }

        /// <summary>  
        /// 获取宽度缩放百分比  
        /// </summary>  
        public static float ScaleX
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                float scaleX = (float)(float)GetDeviceCaps(hdc, DESKTOPHORZRES) / (float)GetDeviceCaps(hdc, HORZRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return scaleX;
            }
        }
        /// <summary>  
        /// 获取高度缩放百分比  
        /// </summary>  
        public static float ScaleY
        {
            get
            {
                IntPtr hdc = GetDC(IntPtr.Zero);
                float scaleY = (float)(float)GetDeviceCaps(hdc, DESKTOPVERTRES) / (float)GetDeviceCaps(hdc, VERTRES);
                ReleaseDC(IntPtr.Zero, hdc);
                return scaleY;
            }
        }
        #endregion
    }
}
