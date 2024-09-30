using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using WinFormApp.Entities;
using WinFormApp.OaTask;
using Lau.Net.Utils;

namespace WinFormApp
{
    public partial class FrmOaTask : Form
    {
        AppSettingOptions _appSettingOptions;
        public static IServiceProvider _serviceProvider;
        readonly ILogger<FrmOaTask> _logger;
        OaService _oaService;

        public FrmOaTask(IServiceProvider serviceProvider, IOptions<AppSettingOptions> appSettingOptions, ILogger<FrmOaTask> logger, OaService oaService, OaTaskScheduler oaTaskScheduler)
        {
            InitializeComponent();
            _serviceProvider = serviceProvider;
            _appSettingOptions = appSettingOptions.Value;
            _logger = logger;
            _oaService = oaService;
            serviceProvider.GetRequiredService<LoggerTest>().TestNlog();

            oaTaskScheduler.StartTask(); 
        }


        private async void FrmOaTask_Load(object sender, EventArgs e)
        {
            _logger.LogInformation("asdf");
            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Normal;

            //await OaTaskScheduler.StartTask();
            _oaService.RectifyAbnormalTask();
        }

        public static void ShowNotifyInfo(string message, string title = "")
        {
            var frmOaTask = _serviceProvider.GetRequiredService<FrmOaTask>();
            frmOaTask.notifyIcon1.ShowBalloonTip(5, title, message, ToolTipIcon.Info);
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (this.Visible)
            {
                this.Hide();
            }
            else
            {
                this.Show();
                this.WindowState = FormWindowState.Normal;
            }
        }

        private void FrmOaTask_SizeChanged(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.Hide();
                this.notifyIcon1.Visible = true;
            }
        }

        private void tsmiQueryTodayTasks_Click(object sender, EventArgs e)
        {
            var dt = _oaService.GetTodayTask();
            var sb = new StringBuilder();
            var getTime = new Func<DataRow, string, string>((row, field) =>
            {
                var time = row.GetValue<DateTime>(field);
                if (time.Year < 2000)
                {
                    return "无";
                }
                return time.ToString("HH:mm");
            });
            foreach (DataRow row in dt.Rows)
            {
                sb.AppendLine($"{row["bm"]}、开始时间：{getTime(row, "计划开始时间")}({getTime(row, "实际开始时间")})、结束时间:{getTime(row, "计划结束时间")}({getTime(row, "实际结束时间")})");
            }
            rtxt.Text = sb.ToString();
        }

        private void tsmiRectifyTask_Click(object sender, EventArgs e)
        {
            _oaService.RectifyAbnormalTask();
        }
    }
}