using System.Threading.Tasks;
using WinFormApp.OaTask;

namespace WinFormApp
{
    public partial class FrmOaTask : Form
    {
        static FrmOaTask _Instance;
        public static FrmOaTask Instance
        {
            get
            {
                if (_Instance == null)
                {
                    _Instance = new FrmOaTask();
                }
                return _Instance;
            }
        }
        public FrmOaTask()
        {
            InitializeComponent();
            _Instance = this;
        }

        private async void FrmOaTask_Load(object sender, EventArgs e)
        {
            this.ShowInTaskbar = false;
            this.WindowState = FormWindowState.Normal;

            await OaTaskScheduler.StartTask();
        }

        public static void ShowNotifyInfo(string message, string title = "")
        {
            Instance.notifyIcon1.ShowBalloonTip(5, title, message, ToolTipIcon.Info);
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
    }
}