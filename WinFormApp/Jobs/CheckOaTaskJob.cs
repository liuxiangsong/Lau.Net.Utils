using Microsoft.Extensions.DependencyInjection;
using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormApp.OaTask;

namespace WinFormApp.Jobs
{
    internal class CheckOaTaskJob : IJob
    { 
        public async Task Execute(IJobExecutionContext context)
        {
            await Task.Run(async () =>
            {
                OaService oaService = FrmOaTask._serviceProvider.GetRequiredService<OaService>();
                var flag = await oaService.ExistsAbnormalTask();
                if (flag)
                {
                    FrmOaTask.ShowNotifyInfo("您有异常的OA任务", "警告");
                    oaService.RectifyAbnormalTask();
                    oaService.FinishTodayTask();
                } 
            });
        }
    }
}
