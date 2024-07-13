using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormApp.Jobs
{
    internal class OaNotifyJob : IJob
    {
        public async Task Execute(IJobExecutionContext context)
        {
           await Task.Run(() =>
            {
                FrmOaTask.ShowNotifyInfo("OA任务即将开始或结束");
            });
        }
    }
}
