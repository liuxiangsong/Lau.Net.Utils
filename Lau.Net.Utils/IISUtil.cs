using Microsoft.Web.Administration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.AxHost;

namespace Lau.Net.Utils
{
    /// <summary>
    /// 操作IIS帮助类
    /// </summary>
    public class IISUtil
    {
        #region 站点相关方法
        /// <summary>
        /// 停止站点
        /// </summary>
        /// <param name="siteName">站点名称</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static bool StopSite(string siteName)
        {
            using (var iisManager = new ServerManager())
            {
                var site = iisManager.Sites[siteName];
                if (site == null)
                {
                    throw new Exception($"{siteName}站点不存在");
                }
                if (IsStateStop(site.State))
                {
                    return true;
                }
                else
                {
                    var state = site.Stop();
                    var result = IsStateStop(state);
                    return result;
                }
            }                
        }

        /// <summary>
        /// 启动站点
        /// </summary>
        /// <param name="siteName">站点名称</param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
        public static bool StartSite(string siteName)
        {
            using (var iisManager = new ServerManager())
            {
                var site = iisManager.Sites[siteName];
                if (site == null)
                {
                    throw new Exception($"{siteName}站点不存在");
                }
                if (IsStateRunning(site.State))
                {
                    return true;
                }
                else
                {
                    var state = site.Start();
                    var result = IsStateRunning(state);
                    return result;
                }
            }
        }

        /// <summary>
        /// 删除站点
        /// </summary>
        /// <param name="siteName">站点名称</param>
        public static void DeleteSite(string siteName)
        {
            using (var iisManager = new ServerManager())
            {
                var site = iisManager.Sites[siteName];
                if (site == null)
                {
                    throw new Exception($"{siteName}站点不存在");
                }
                if (site != null)
                {
                    iisManager.Sites.Remove(site);
                    iisManager.CommitChanges();
                }
            }
        }

        /// <summary>
        /// 修改站名物理路径
        /// </summary>
        /// <param name="siteName">站点名称</param>
        /// <param name="newPhysicalPath">新的物理路径</param>
        /// <exception cref="Exception"></exception>
        public static void ModifySitePhysicalPath(string siteName, string newPhysicalPath)
        {
            using (var iisManager = new ServerManager())
            {
                var site = iisManager.Sites[siteName];
                if (site == null)
                {
                    throw new Exception($"{siteName}站点不存在");
                }
                site.Applications[0].VirtualDirectories[0].PhysicalPath = newPhysicalPath;
                iisManager.CommitChanges();
            }
        }
            #endregion

            #region 应用程序池相关方法
            /// <summary>
            /// 停止应用程序池
            /// </summary>
            /// <param name="appPoolName"></param>
            public static bool StopApplicationPool(string appPoolName = null)
        {
            if (string.IsNullOrWhiteSpace(appPoolName))
            {
                appPoolName = Environment.GetEnvironmentVariable("APP_POOL_ID", EnvironmentVariableTarget.Process);
            }
            using (var iisManager = new ServerManager())
            {
                var appPool = iisManager.ApplicationPools[appPoolName];
                if (IsStateStop(appPool.State))
                {
                    return true;
                }
                else
                {
                    var state = appPool.Stop();
                    return IsStateStop(state);
                }
            }                
        }

        /// <summary>
        /// 启动应用程序池
        /// </summary>
        /// <param name="appPoolName"></param>
        public static bool StartApplicationPool(string appPoolName = null)
        {
            if (string.IsNullOrWhiteSpace(appPoolName))
            {
                appPoolName = Environment.GetEnvironmentVariable("APP_POOL_ID", EnvironmentVariableTarget.Process);
            }
            using (var iisManager = new ServerManager())
            {
                var appPool = iisManager.ApplicationPools[appPoolName];
                if (IsStateRunning(appPool.State))
                {
                    return true;
                }
                else
                {
                    var state = appPool.Start();
                    return IsStateRunning(state);
                }
            }                
        }

        /// <summary>
        /// 回收应用程序池
        /// </summary>
        /// <param name="appPoolName"></param>
        /// <returns></returns>
        public static bool RecycleApplicationPool(string appPoolName = null)
        {
            if (string.IsNullOrWhiteSpace(appPoolName))
            {
                appPoolName = Environment.GetEnvironmentVariable("APP_POOL_ID", EnvironmentVariableTarget.Process);
            }
            using (var iisManager = new ServerManager())
            {
                var appPool = iisManager.ApplicationPools[appPoolName];
                var state = appPool.Recycle();
                return true;
            }
        }
        #endregion

        #region 私有方法

        private static bool IsStateStop(ObjectState state)
        {
            var result = (state == ObjectState.Stopped || state == ObjectState.Stopping);
            return result;
        }
        private static bool IsStateRunning(ObjectState state)
        {
            var result = (state == ObjectState.Started || state == ObjectState.Starting);
            return result;
        }
        #endregion
    }
}
