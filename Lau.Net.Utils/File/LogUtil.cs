using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace Lau.Net.Utils
{
    /// <summary>
    /// 用途：LOG日志帮助类
    /// 参考：http://www.cnblogs.com/springyangwc/archive/2012/03/30/2425822.html
    /// </summary>
    public class LogUtil : IDisposable
    {
        //日志队列
        private static Queue<LogMessage> _logMessages;
        //日志保存路径
        private static string _logDirectory;
        //写入状态
        private static bool _state;
        //日志频率
        private static LogType _logType;
        //日志时间戳
        private static DateTime _timeSign;

        private static StreamWriter _writer;

        //限制可同时访问某一资源或资源池的线程数
        private Semaphore _semaphore;

        //单例模式
        private static LogUtil _log;
        ////返回单列
        //public static LogUtil LogInstance
        //{
        //    get
        //    {
        //        return _log ?? (_log = new LogUtil());
        //    }
        //}

        //加锁
        private object _lockObjeck;

        private void Initialize(string logDirPath)
        {
            if (_logMessages == null)
            {
                _state = true;
                //string logPath = Application.StartupPath + "\\Log\\";
                _logDirectory = logDirPath;
                if (!Directory.Exists(_logDirectory))
                {
                    Directory.CreateDirectory(_logDirectory);
                }
                _logType = LogType.Daily;
                _lockObjeck = new object();
                _semaphore = new Semaphore(0, int.MaxValue, "LogSemaphoreName");
                _logMessages = new Queue<LogMessage>();
                var thread = new Thread(Work) { IsBackground = true };
                thread.Start();
            }
        }

        public LogUtil(string logDirPath)
        {
            if(_log == null)
            {
                Initialize(logDirPath);
                _log = this;
            }            
        }

        /// <summary>
        /// 日志频率 默认每天
        /// </summary>
        public LogType LogType { get; set; }

        public void Work()
        {
            while (true)
            {
                if (_logMessages.Count > 0)
                {
                    FileWriteMessage();
                }
                else
                {
                    if (WaitLogMessage()) break;
                }
            }
        }
        //线程等待中
        private bool WaitLogMessage()
        {
            if (_state)
            {
                WaitHandle.WaitAny(new WaitHandle[] { _semaphore }, -1, false);
                return false;
            }
            FileClose();
            return true;
        }



        //文件类型
        private string GetFilename()
        {
            DateTime now = DateTime.Now;
            string format = "";
            switch (_logType)
            {
                case LogType.Daily:
                    _timeSign = new DateTime(now.Year, now.Month, now.Day);
                    _timeSign = _timeSign.AddDays(1);
                    format = "yyyyMMdd'.log'";
                    break;
                case LogType.Weekly:
                    _timeSign = new DateTime(now.Year, now.Month, now.Day);
                    _timeSign = _timeSign.AddDays(7);
                    format = "yyyyMMdd'.log'";
                    break;
                case LogType.Monthly:
                    _timeSign = new DateTime(now.Year, now.Month, 1);
                    _timeSign = _timeSign.AddMonths(1);
                    format = "yyyyMM'.log'";
                    break;
                case LogType.Annually:
                    _timeSign = new DateTime(now.Year, 1, 1);
                    _timeSign = _timeSign.AddYears(1);
                    format = "yyyy'.log'";
                    break;
            }
            return now.ToString(format);
        }
        private void FileWriteMessage()
        {
            LogMessage logMessage = null;
            lock (_lockObjeck)
            {
                if (_logMessages.Count > 0)
                    logMessage = _logMessages.Dequeue();
            }
            if (logMessage != null)
            {
                FileWrite(logMessage);
            }
        }

        private void FileWrite(LogMessage logMessage)
        {
            try
            {
                if (_writer == null)
                {
                    FileOpen();
                }
                if (DateTime.Now >= _timeSign)
                {
                    FileClose();
                    FileOpen();
                }
                if (_writer == null) return;
                _writer.WriteLine("LogMessageTime:" + logMessage.Datetime);
                _writer.WriteLine("LogMessageType:" + logMessage.Type);
                _writer.WriteLine("LogMessageContent:" + logMessage.Text + "\r\n");
                _writer.Flush();
            }
            catch (Exception)
            {
            }
        }

        private void FileOpen()
        {
            _writer = new StreamWriter(Path.Combine(_logDirectory, GetFilename()), true, Encoding.UTF8);
        }
        private void FileClose()
        {
            if (_writer != null)
            {
                _writer.Flush();
                _writer.Close();
                _writer.Dispose();
                _writer = null;
            }
        }

        public void Write(LogMessage msg)
        {
            if (msg != null)
            {
                lock (_lockObjeck)
                {
                    _logMessages.Enqueue(msg);
                    _semaphore.Release();
                }
            }
        }
        public void Write(string text, LogMessageType type)
        {
            Write(new LogMessage(text, type));
        }
        public void Write(DateTime dateTime, string text, LogMessageType type)
        {
            Write(new LogMessage(dateTime, text, type));
        }
        public void Write(Exception e, LogMessageType type, string customInfo = "")
        {
            var msg = e.Message;
            if (!string.IsNullOrWhiteSpace(customInfo))
            {
                msg += string.Format("({0})", customInfo);
            }
            var content = string.Format("{0}\r\nSource:{1}\r\nTargetSite:{2}\r\nStackTrace:{3}\r\nInnerException:{4} \r\n",
                msg, e.Source, e.TargetSite, e.StackTrace, e.InnerException);
            Write(new LogMessage(content, type));
        }

        public void Dispose()
        {
            _state = false;
        }
    }

    /// <summary>
    /// 写入日志频率
    /// </summary>
    public enum LogType
    {
        Daily, //每天
        Weekly,//每周 
        Monthly,//每月 
        Annually,//每年
    }

    //日志类型
    public enum LogMessageType
    {
        Unknown, Information, Warning, Error, Success
    }

    /// <summary>
    /// 日志类
    /// </summary>
    public class LogMessage
    {
        public string Text { get; set; }
        public LogMessageType Type { get; set; }
        public DateTime Datetime { get; set; }
        public LogMessage()
            : this("", LogMessageType.Unknown)
        {
        }


        public LogMessage(string text, LogMessageType messageType)
            : this(DateTime.Now, text, messageType)
        {
        }

        public LogMessage(DateTime dateTime, string text, LogMessageType messageType)
        {
            Datetime = dateTime;
            Type = messageType;
            Text = text;
        }

        public new string ToString()
        {
            return Datetime.ToString(CultureInfo.InvariantCulture) + "\t" + Text + "\n";
        }

    }
}
