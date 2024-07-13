//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Runtime.InteropServices;
//using System.Text;
//using System.Windows.Forms;

//namespace Lau.Net.Utils
//{
//    public class HotKeysUtil
//    {

//        #region 引入系统API
//        //要得到扩展错误信息，需把SetLastError设为True。
//        [DllImport("user32.dll", SetLastError = true)]
//        static extern bool RegisterHotKey(IntPtr hWnd, int id, int modifiers, Keys vk);
//        [DllImport("user32.dll", SetLastError = true)]
//        static extern bool UnregisterHotKey(IntPtr hWnd, int id);
//        #endregion

//        int keyid = 10;     //区分不同的快捷键
//        public delegate void HotKeyCallBackHanlder();

//        //每一个key对于一个处理函数
//        Dictionary<int, HotKeyCallBackHanlder> keymap = new Dictionary<int, HotKeyCallBackHanlder>();

//        //组合控制键
//        public enum HotkeyModifiers
//        {
//            Alt = 1,
//            Control = 2,
//            Shift = 4,
//            Win = 8
//        }

//        #region 注册快捷键
//        /// <summary>
//        /// 注册快捷键
//        /// </summary>
//        /// <param name="hWnd">持有快捷键窗口的句柄</param>
//        /// <param name="modifiers">组合键</param>
//        /// <param name="vk">快捷键的虚拟键码</param>
//        /// <param name="callBack">回调函数</param>
//        public void Regist(IntPtr hWnd, int modifiers, Keys vk, HotKeyCallBackHanlder callBack)
//        {
//            int id = keyid++;
//            if (!RegisterHotKey(hWnd, id, modifiers, vk))
//            {
//                if (Marshal.GetLastWin32Error() == 1409)
//                {
//                    MessageBox.Show("热键被占用 ！");
//                }
//                else
//                {
//                    MessageBox.Show("注册失败！");
//                }
//            }
//            else
//            {
//                keymap[id] = callBack;
//            }

//        }
//        #endregion

//        #region 注销快捷键
//        /// <summary>
//        /// 注销快捷键
//        /// </summary>
//        /// <param name="hWnd">持有快捷键窗口的句柄</param>
//        /// <param name="callBack">回调函数</param>
//        public void UnRegist(IntPtr hWnd, HotKeyCallBackHanlder callBack)
//        {
//            foreach (KeyValuePair<int, HotKeyCallBackHanlder> var in keymap)
//            {
//                if (var.Value == callBack)
//                {
//                    UnregisterHotKey(hWnd, var.Key);
//                    return;
//                }
//            }
//        }
//        #endregion

//        #region 快捷键消息处理
//        public void ProcessHotKey(Message m)
//        {
//            if (m.Msg == 0x312)
//            {
//                int id = m.WParam.ToInt32();
//                HotKeyCallBackHanlder callback;
//                if (keymap.TryGetValue(id, out callback))
//                    callback();
//            }
//        }
//        #endregion

//    }
//}
