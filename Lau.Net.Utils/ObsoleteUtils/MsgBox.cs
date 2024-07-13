//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Windows.Forms;

//namespace Lau.Net.Utils
//{
//    public class MsgBox
//    {
//        /// <summary>
//        /// 显示一般的提示信息
//        /// </summary>
//        /// <param name="message">提示信息</param>
//        public static DialogResult ShowInformation(string message)
//        {
//            return MessageBox.Show(message, "提示信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
//        }

//        /// <summary>
//        /// 显示警告信息
//        /// </summary>
//        /// <param name="message">警告信息</param>
//        public static DialogResult ShowWarning(string message)
//        {
//            return MessageBox.Show(message, "警告信息", MessageBoxButtons.OK, MessageBoxIcon.Warning);
//        }

//        /// <summary>
//        /// 显示错误信息
//        /// </summary>
//        /// <param name="message">错误信息</param>
//        public static DialogResult ShowError(string message)
//        {
//            return MessageBox.Show(message, "错误信息", MessageBoxButtons.OK, MessageBoxIcon.Error);
//        }

//        /// <summary>
//        /// 显示询问提示信息，包含"是"和"否"
//        /// </summary>
//        /// <param name="message">提示信息</param>
//        /// <param name="defaultButton">默认按钮,默认选择第一个按钮</param>
//        /// <returns></returns>
//        public static DialogResult ShowYesNoQuestion(string message, MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
//        {
//            return MessageBox.Show(message, "提示信息", MessageBoxButtons.YesNo, MessageBoxIcon.Question, defaultButton);
//        }

//        /// <summary>
//        /// 显示询问警告信息，包含"是"和"否"
//        /// </summary>
//        /// <param name="message">警告信息</param>
//        /// <param name="defaultButton">默认按钮,默认选择第一个按钮</param>
//        /// <returns></returns>
//        public static DialogResult ShowYesNoWarning(string message, MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
//        {
//            return MessageBox.Show(message, "警告信息", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, defaultButton);
//        }

//        /// <summary>
//        /// 显示询问错误信息，包含"是"和"否"
//        /// </summary>
//        /// <param name="message">错误信息</param>
//        /// <param name="defaultButton">默认按钮,默认选择第一个按钮</param>
//        /// <returns></returns>
//        public static DialogResult ShowYesNoError(string message, MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
//        {
//            return MessageBox.Show(message, "错误信息", MessageBoxButtons.YesNo, MessageBoxIcon.Error, defaultButton);
//        }


//        /// <summary>
//        /// 显示询问用户信息，并显示提示标志
//        /// </summary>
//        /// <param name="message">错误信息</param>
//        /// <param name="defaultButton">默认按钮,默认选择第一个按钮</param>
//        /// <returns></returns>
//        public static DialogResult ShowYesNoCancelInformation(string message, MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button1)
//        {
//            return MessageBox.Show(message, "提示信息", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information, defaultButton);
//        }



//        /// <summary>
//        /// 显示YesNo选择对话框
//        /// </summary>
//        /// <param name="message">对话框的选择内容提示信息</param>
//        /// <param name="defaultButton">默认按钮,默认选择第二个按钮</param>
//        /// <returns>如果选择Yes则返回true，否则返回false</returns>
//        public static bool ConfirmYesNo(string message, MessageBoxDefaultButton defaultButton = MessageBoxDefaultButton.Button2)
//        {
//            return MessageBox.Show(message, "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, defaultButton) == DialogResult.Yes;
//        }

//    }
//}
