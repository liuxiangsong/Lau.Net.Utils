
using FreeSql;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using NLog.Extensions.Logging;
using Quartz.Impl;
using Quartz.Spi;
using WinFormApp.Entities;
using WinFormApp.Jobs;
using WinFormApp.OaTask;

namespace WinFormApp
{
    internal static class Program
    {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            var configuration = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json")
                .Build();

            var services = new ServiceCollection();
            //注入AppSettingOptions配置
            services.Configure<AppSettingOptions>(configuration.GetSection("AppSettings"));

            services.AddSingleton<FrmOaTask>();
            services.AddTransient<LoggerTest>();

            #region 配置Nlog
            services.AddLogging(loggingBuilder =>
    {
        loggingBuilder.ClearProviders();
        loggingBuilder.SetMinimumLevel(LogLevel.Trace);
        loggingBuilder.AddNLog("XmlConfig/nlog.config");
    });
            #endregion


            #region 配置Freesql
            var fsql2 = new FreeSqlBuilder().UseConnectionString(DataType.SqlServer, configuration.GetConnectionString("OA"))
               .Build();
            services.AddSingleton<IFreeSql>(fsql2); 
            #endregion

            services.AddSingleton<OaService>();
            services.AddSingleton<OaTaskScheduler>();
            //services.AddSingleton<CheckOaTaskJob>();
            //services.AddSingleton <OaNotifyJob>();
            //services.AddTransient<IJobFactory,ServiceProviderJobFactory >
            services.AddSingleton(provider =>
            {
                // 创建 Quartz 调度器
                var schedulerFactory = new StdSchedulerFactory();
                var scheduler = schedulerFactory.GetScheduler().GetAwaiter().GetResult();
                 
                return scheduler;
            });

            using var serviceProvider = services.BuildServiceProvider();
             
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize(); 
            var frmOaTask = serviceProvider.GetRequiredService<FrmOaTask>();
            Application.Run(frmOaTask);
        }
    }
}