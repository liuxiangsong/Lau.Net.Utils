using Quartz.Impl;
using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WinFormApp.Jobs;
using static Quartz.Logging.OperationName;

namespace WinFormApp.OaTask
{
    public class OaTaskScheduler
    {
        IScheduler _scheduler = null;

        public OaTaskScheduler(IScheduler scheduler)
        {
            _scheduler = scheduler;
        }
        public async Task StartTask()
        {
            //StdSchedulerFactory factory = new StdSchedulerFactory();
            //_scheduler = await factory.GetScheduler();

            // and start it off
            await _scheduler.Start();
            // 每天8:00、12:20、18:20、13:50执行
            foreach (var item in new string[] { "0 8", "20 12,18", "50 13" })
            {
                // define the job and tie it to our HelloJob class
                IJobDetail job = JobBuilder.Create<OaNotifyJob>()
                    .WithIdentity("OaNotifyJob_job", item)
                    .Build();
                ITrigger trigger = TriggerBuilder.Create()
                .WithIdentity("trigger1", item)
                 //每天8:00执行
                 .WithCronSchedule($"0 {item} ? * *")
                .Build();
                await _scheduler.ScheduleJob(job, trigger);
            }
            AddCheckJob(); 
        }

        public async void AddCheckJob()
        {
            var jobName = "ChekcOATask";
            IJobDetail job = JobBuilder.Create<CheckOaTaskJob>()
              .WithIdentity(jobName, "group2")
              .Build();

            ITrigger trigger = TriggerBuilder.Create()
                .WithIdentity(jobName, "group2")
                .WithCronSchedule("0 25 18 * * ?")
                .Build();

            await _scheduler.ScheduleJob(job, trigger);
            ////手动触发一次
            //await _scheduler.TriggerJob(job.Key);
        }

        public void AddJob<T>(DateTime startTime) where T : IJob
        {
            var jobName = typeof(T).Name + "_" + startTime.ToString("yyyyMMdd HH:mm:ss");
            IJobDetail job = JobBuilder.Create<T>()
                .WithIdentity(jobName, "group1")
                .Build();

            ITrigger trigger = TriggerBuilder.Create()
                .WithIdentity(jobName, "group1")
                .StartAt(startTime)
                .Build();

            _scheduler.ScheduleJob(job, trigger);
        }
    }
}
