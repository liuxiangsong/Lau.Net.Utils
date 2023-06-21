using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;

namespace ConsoleAppDemo
{
    internal static class  ConfigureService
    {
       internal static void Configure()
        {
            HostFactory.Run(x =>
            {
                x.Service<PsTaskCheckService>();
                x.EnableServiceRecovery(r => r.RestartService(TimeSpan.FromSeconds(60)));
                x.SetServiceName("PsTaskCheckService");
                x.SetDescription("PS任务检测服务");
                x.RunAsLocalSystem();
                x.StartAutomatically();
            });
        }
    }
}
