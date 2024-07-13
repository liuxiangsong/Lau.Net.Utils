using Microsoft.Extensions.Logging;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormApp
{
    public class LoggerTest
    {
        private readonly ILogger<LoggerTest> _logger; 
        public LoggerTest(ILogger<LoggerTest> logger )
        {
            _logger = logger; 
        }

        public void TestNlog()
        {
            _logger.LogInformation("测试Nlog,{name}","张三"); 
        }
    }
}
