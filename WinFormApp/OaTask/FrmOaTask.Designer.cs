namespace WinFormApp
{
    partial class FrmOaTask
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmOaTask));
            notifyIcon1 = new NotifyIcon(components);
            menuStrip1 = new MenuStrip();
            tsmiQueryTodayTasks = new ToolStripMenuItem();
            rtxt = new RichTextBox();
            tsmiRectifyTask = new ToolStripMenuItem();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // notifyIcon1
            // 
            notifyIcon1.Icon = (Icon)resources.GetObject("notifyIcon1.Icon");
            notifyIcon1.Text = "任务提醒";
            notifyIcon1.Visible = true;
            notifyIcon1.MouseDoubleClick += notifyIcon1_MouseDoubleClick;
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new ToolStripItem[] { tsmiQueryTodayTasks, tsmiRectifyTask });
            menuStrip1.Location = new Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new Size(800, 25);
            menuStrip1.TabIndex = 0;
            menuStrip1.Text = "menuStrip1";
            // 
            // tsmiQueryTodayTasks
            // 
            tsmiQueryTodayTasks.Name = "tsmiQueryTodayTasks";
            tsmiQueryTodayTasks.Size = new Size(104, 21);
            tsmiQueryTodayTasks.Text = "查询今天的任务";
            tsmiQueryTodayTasks.Click += tsmiQueryTodayTasks_Click;
            // 
            // rtxt
            // 
            rtxt.Dock = DockStyle.Fill;
            rtxt.Location = new Point(0, 25);
            rtxt.Name = "rtxt";
            rtxt.Size = new Size(800, 425);
            rtxt.TabIndex = 1;
            rtxt.Text = "";
            // 
            // tsmiRectifyTask
            // 
            tsmiRectifyTask.Name = "tsmiRectifyTask";
            tsmiRectifyTask.Size = new Size(68, 21);
            tsmiRectifyTask.Text = "检查异常";
            tsmiRectifyTask.Click += tsmiRectifyTask_Click;
            // 
            // FrmOaTask
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(rtxt);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Name = "FrmOaTask";
            Text = "Form1";
            Load += FrmOaTask_Load;
            SizeChanged += FrmOaTask_SizeChanged;
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private NotifyIcon notifyIcon1;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem tsmiQueryTodayTasks;
        private RichTextBox rtxt;
        private ToolStripMenuItem tsmiRectifyTask;
    }
}