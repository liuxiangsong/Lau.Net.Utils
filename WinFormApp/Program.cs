
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using WinFormApp.Entities;

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
            services.Configure<AppSettingOptions>(configuration.GetSection("AppSettings"));

            services.AddSingleton<FrmOaTask>();
            using var serviceProvider = services.BuildServiceProvider();

            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize(); 
            var frmOaTask = serviceProvider.GetRequiredService<FrmOaTask>();
            Application.Run(frmOaTask);
        }
    }
}