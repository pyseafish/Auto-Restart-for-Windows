using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace AutoRebootApp
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }

    public class Form1 : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private System.Windows.Forms.Timer checkTimer;
        private System.Windows.Forms.Timer countdownTimer;
        private System.Windows.Forms.Timer tipUpdateTimer;  // 用于更新托盘文本
        private Mutex mutex;

        private const string REG_KEY = @"Software\AutoRebootApp";
        private const string LAST_REBOOT = "LastReboot";
        private const string INTERVAL_HOURS = "IntervalHours";
        private const string TARGET_APPS = "TargetApps";
        private const string REBOOT_MODE = "RebootMode";        // 0:禁用, 1:间隔, 2:定时
        private const string SCHEDULED_TIME = "ScheduledTime";  // 定时重启的UTC ticks
        private const string LOG_DIR = "AutoRebootApp";
        private string logPath;

        private int rebootMode = 1;           // 默认间隔模式
        private int intervalHours = 72;
        private long scheduledTicks = 0;     // 定时重启的UTC ticks
        private DateTime lastRebootUtc;

        private bool isCountdownActive = false;
        private int countdownSeconds = 300;   // 5分钟倒计时
        private ToolStripMenuItem cancelRebootItem;
        private ToolStripMenuItem rebootNowItem;

        public Form1()
        {
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Load += (s, e) => this.Hide();

            // 初始化日志路径
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string logDir = Path.Combine(appData, LOG_DIR);
            Directory.CreateDirectory(logDir);
            logPath = Path.Combine(logDir, "reboot.log");

            Log("程序启动");

            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("设置", null, OpenSettings);
            trayMenu.Items.Add("立即重启", null, RebootNow);
            // 分隔线
            trayMenu.Items.Add(new ToolStripSeparator());
            // 取消/立即重启菜单项（默认隐藏）
            cancelRebootItem = new ToolStripMenuItem("取消本次重启", null, CancelReboot);
            cancelRebootItem.Visible = false;
            rebootNowItem = new ToolStripMenuItem("立即重启(倒计时)", null, RebootNow);
            rebootNowItem.Visible = false;
            trayMenu.Items.Add(cancelRebootItem);
            trayMenu.Items.Add(rebootNowItem);
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("查看日志", null, ShowLog);
            trayMenu.Items.Add("退出", null, ExitApp);

            trayIcon = new NotifyIcon();
            trayIcon.Text = "自动重启工具";
            try
            {
                // 可自行替换为自定义图标，这里使用系统图标
                trayIcon.Icon = SystemIcons.Application;
            }
            catch { }
            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.Visible = true;
            trayIcon.DoubleClick += OpenSettings;

            // 单实例检测
            bool createdNew;
            mutex = new Mutex(true, "AutoRebootApp_UniqueMutex", out createdNew);
            if (!createdNew)
            {
                MessageBox.Show("程序已在运行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Environment.Exit(0);
            }

            LoadSettingsFromRegistry();
            StartTimers();

            Log($"当前模式: {(rebootMode == 0 ? "禁用" : rebootMode == 1 ? "间隔" : "定时")}");
            UpdateTrayTip();
        }

        private void LoadSettingsFromRegistry()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY))
            {
                rebootMode = (int)(key.GetValue(REBOOT_MODE) ?? 1);
                intervalHours = (int)(key.GetValue(INTERVAL_HOURS) ?? 72);
                scheduledTicks = Convert.ToInt64(key.GetValue(SCHEDULED_TIME, 0L));

                // 读取上次重启时间（UTC ticks）
                object lastTicks = key.GetValue(LAST_REBOOT);
                if (lastTicks != null)
                    lastRebootUtc = new DateTime(long.Parse(lastTicks.ToString()), DateTimeKind.Utc);
                else
                {
                    // 首次运行，记录当前时间作为起点
                    lastRebootUtc = DateTime.UtcNow;
                    key.SetValue(LAST_REBOOT, lastRebootUtc.Ticks.ToString());
                }
            }
        }

        private void StartTimers()
        {
            // 主检查定时器：每小时检查一次
            checkTimer = new System.Windows.Forms.Timer();
            checkTimer.Interval = 3600000; // 1小时
            checkTimer.Tick += CheckRebootCondition;
            if (rebootMode != 0) checkTimer.Start();

            // 托盘文本更新定时器：每秒更新
            tipUpdateTimer = new System.Windows.Forms.Timer();
            tipUpdateTimer.Interval = 1000;
            tipUpdateTimer.Tick += (s, e) => UpdateTrayTip();
            tipUpdateTimer.Start();

            // 倒计时定时器（按需创建，初始不启动）
            countdownTimer = new System.Windows.Forms.Timer();
            countdownTimer.Interval = 1000;
            countdownTimer.Tick += CountdownTimer_Tick;
        }

        // ===================== 日志工具 =====================
        private void Log(string message)
        {
            try
            {
                string line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}";
                File.AppendAllText(logPath, line + Environment.NewLine);
            }
            catch { /* 忽略日志写入异常 */ }
        }

        // ===================== 托盘提示更新 =====================
        private void UpdateTrayTip()
        {
            string tip = "自动重启工具";
            if (isCountdownActive)
            {
                tip = $"系统将在 {countdownSeconds} 秒后自动重启";
            }
            else
            {
                switch (rebootMode)
                {
                    case 0:
                        tip += " | 已禁用";
                        break;
                    case 1:
                        DateTime nextReboot = lastRebootUtc.AddHours(intervalHours);
                        if (nextReboot < DateTime.UtcNow)
                            tip += " | 即将重启...";
                        else
                            tip += $" | 下次重启: {nextReboot.ToLocalTime():yyyy-MM-dd HH:mm:ss}";
                        break;
                    case 2:
                        if (scheduledTicks > 0)
                        {
                            DateTime scheduled = new DateTime(scheduledTicks, DateTimeKind.Utc);
                            if (scheduled < DateTime.UtcNow)
                                tip += " | 定时重启已到期，即将执行";
                            else
                                tip += $" | 定时重启: {scheduled.ToLocalTime():yyyy-MM-dd HH:mm:ss}";
                        }
                        break;
                }
            }
            // 托盘Text限制63字符，裁剪
            if (tip.Length > 63) tip = tip.Substring(0, 60) + "...";
            trayIcon.Text = tip;
        }

        // ===================== 重启条件检查 =====================
        private void CheckRebootCondition(object sender, EventArgs e)
        {
            if (isCountdownActive) return; // 倒计时中不再重复触发

            bool needReboot = false;
            switch (rebootMode)
            {
                case 1: // 间隔模式
                    if ((DateTime.UtcNow - lastRebootUtc).TotalHours >= intervalHours)
                        needReboot = true;
                    break;
                case 2: // 定时模式
                    if (scheduledTicks > 0 && DateTime.UtcNow.Ticks >= scheduledTicks)
                        needReboot = true;
                    break;
            }

            if (needReboot)
            {
                Log("检测到需要重启，启动5分钟倒计时");
                StartCountdown();
            }
        }

        // ===================== 倒计时机制 =====================
        private void StartCountdown()
        {
            isCountdownActive = true;
            countdownSeconds = 300; // 5分钟

            // 显示倒计时菜单项
            cancelRebootItem.Visible = true;
            rebootNowItem.Visible = true;

            countdownTimer.Start();
            UpdateTrayTip();

            // 托盘气泡通知
            trayIcon.ShowBalloonTip(5000, "自动重启提醒", "系统将在5分钟后重启，请及时保存工作。", ToolTipIcon.Warning);
        }

        private void CountdownTimer_Tick(object sender, EventArgs e)
        {
            countdownSeconds--;
            if (countdownSeconds <= 0)
            {
                // 倒计时结束，执行重启
                countdownTimer.Stop();
                Log("倒计时结束，执行重启");
                DoReboot();
            }
            else
            {
                // 在最后30秒再次提醒
                if (countdownSeconds == 30)
                {
                    trayIcon.ShowBalloonTip(4000, "即将重启", "系统将在30秒后自动重启！", ToolTipIcon.Warning);
                }
                UpdateTrayTip();
            }
        }

        private void CancelReboot(object sender, EventArgs e)
        {
            if (!isCountdownActive) return;

            // 停止倒计时
            countdownTimer.Stop();
            isCountdownActive = false;
            cancelRebootItem.Visible = false;
            rebootNowItem.Visible = false;

            // 重置上次重启时间，避免再次立即触发
            lastRebootUtc = DateTime.UtcNow;
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY))
                key.SetValue(LAST_REBOOT, lastRebootUtc.Ticks.ToString());

            Log("用户取消了本次自动重启");
            UpdateTrayTip();
            trayIcon.ShowBalloonTip(3000, "已取消", "本次自动重启已取消", ToolTipIcon.Info);
        }

        private void RebootNow(object sender, EventArgs e)
        {
            // 无论是手动“立即重启”还是倒计时内“立即重启”，统一处理
            // 手动立即重启：增加确认对话框
            if (!isCountdownActive) // 手动触发
            {
                DialogResult result = MessageBox.Show("确定要立即重启系统吗？", "确认重启",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result != DialogResult.Yes) return;
                Log("用户手动触发立即重启");
            }
            else
            {
                Log("倒计时内用户点击立即重启");
            }

            // 停止倒计时（如果正在运行）
            if (countdownTimer.Enabled)
                countdownTimer.Stop();

            DoReboot();
        }

        // ===================== 执行重启 =====================
        private void DoReboot()
        {
            try
            {
                // 1. 写入 RunOnce 启动项
                bool runOnceOk = WriteRunOnce();
                if (!runOnceOk)
                {
                    Log("写入 RunOnce 失败，阻止本次重启");
                    trayIcon.ShowBalloonTip(5000, "错误", "无法写入启动项，重启已取消", ToolTipIcon.Error);
                    return;
                }

                // 2. 根据模式处理状态
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY))
                {
                    if (rebootMode == 2) // 定时模式
                    {
                        // 检查设置中的“转为间隔”标记（从注册表读取一次）
                        using (RegistryKey setKey = Registry.CurrentUser.OpenSubKey(REG_KEY))
                        {
                            bool convertToInterval = (int)(setKey?.GetValue("ConvertToInterval", 0) ?? 0) == 1;
                            if (convertToInterval)
                            {
                                key.SetValue(REBOOT_MODE, 1);
                                key.SetValue(LAST_REBOOT, DateTime.UtcNow.Ticks.ToString());
                            }
                            else
                            {
                                key.SetValue(REBOOT_MODE, 0);
                            }
                            // 清除定时时间
                            key.DeleteValue(SCHEDULED_TIME, false);
                            key.DeleteValue("ConvertToInterval", false);
                        }
                    }
                    else // 间隔模式
                    {
                        key.SetValue(LAST_REBOOT, DateTime.UtcNow.Ticks.ToString());
                    }

                    key.SetValue(LAST_REBOOT, DateTime.UtcNow.Ticks.ToString());
                }

                Log("即将重启系统...");
                Process.Start("shutdown", "/r /t 0");
            }
            catch (Exception ex)
            {
                Log($"重启过程发生异常: {ex.Message}");
                trayIcon.ShowBalloonTip(5000, "错误", "无法执行重启", ToolTipIcon.Error);
            }
        }

        private bool WriteRunOnce()
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_KEY))
                {
                    string apps = key?.GetValue(TARGET_APPS)?.ToString();
                    if (!string.IsNullOrWhiteSpace(apps))
                    {
                        string[] appPaths = apps.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                        using (RegistryKey runOnce = Registry.CurrentUser.OpenSubKey(
                            @"Software\Microsoft\Windows\CurrentVersion\RunOnce", true))
                        {
                            if (runOnce != null)
                            {
                                for (int i = 0; i < appPaths.Length; i++)
                                    runOnce.SetValue($"AutoStartApp_{i}", "\"" + appPaths[i] + "\"");
                            }
                        }
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Log($"写RunOnce失败: {ex.Message}");
                return false;
            }
        }

        // ===================== 设置窗口 =====================
        private void OpenSettings(object sender, EventArgs e)
        {
            using (Form2 settingsForm = new Form2())
            {
                if (settingsForm.ShowDialog() == DialogResult.OK)
                {
                    // 重新加载设置并更新界面
                    LoadSettingsFromRegistry();
                    if (rebootMode == 0)
                        checkTimer.Stop();
                    else
                        checkTimer.Start();
                    Log("设置已更新");
                    UpdateTrayTip();
                }
            }
        }

        // ===================== 日志查看 =====================
        private void ShowLog(object sender, EventArgs e)
        {
            using (LogForm logForm = new LogForm(logPath))
                logForm.ShowDialog();
        }

        // ===================== 退出程序 =====================
        private void ExitApp(object sender, EventArgs e)
        {
            Log("程序退出");
            trayIcon.Visible = false;
            Application.Exit();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                this.Hide();
                return;
            }
            base.OnFormClosing(e);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                trayIcon?.Dispose();
                checkTimer?.Dispose();
                countdownTimer?.Dispose();
                tipUpdateTimer?.Dispose();
                mutex?.ReleaseMutex();
                mutex?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    // ===================== 设置窗体 =====================
    public class Form2 : Form
    {
        private RadioButton rdoDisabled, rdoInterval, rdoScheduled;
        private NumericUpDown numInterval;
        private DateTimePicker dtpDate, dtpTime;
        private CheckBox chkConvertToInterval;
        private ListBox lstApps;
        private Button btnAdd, btnRemove, btnUp, btnDown;
        private CheckBox chkAutoStart;
        private Button btnSave, btnCancel;

        private const string REG_KEY = @"Software\AutoRebootApp";

        public Form2()
        {
            this.Text = "设置 - 自动重启工具";
            this.Size = new System.Drawing.Size(550, 550);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            int leftMargin = 20;
            int top = 15;

            // ====== 重启模式分组 ======
            GroupBox groupMode = new GroupBox()
            {
                Text = "重启模式",
                Location = new System.Drawing.Point(leftMargin, top),
                Size = new System.Drawing.Size(500, 170)
            };
            this.Controls.Add(groupMode);

            rdoDisabled = new RadioButton() { Text = "禁用自动重启", Location = new System.Drawing.Point(15, 25), Size = new System.Drawing.Size(130, 24) };
            rdoInterval = new RadioButton() { Text = "固定间隔重启", Location = new System.Drawing.Point(15, 55), Size = new System.Drawing.Size(130, 24) };
            rdoScheduled = new RadioButton() { Text = "定时重启", Location = new System.Drawing.Point(15, 85), Size = new System.Drawing.Size(130, 24) };
            groupMode.Controls.Add(rdoDisabled);
            groupMode.Controls.Add(rdoInterval);
            groupMode.Controls.Add(rdoScheduled);

            // 间隔设置
            Label lblInterval = new Label() { Text = "间隔小时数：", Location = new System.Drawing.Point(170, 57), Size = new System.Drawing.Size(80, 23) };
            numInterval = new NumericUpDown() { Location = new System.Drawing.Point(250, 55), Size = new System.Drawing.Size(80, 23), Minimum = 1, Maximum = 720, Value = 72 };
            groupMode.Controls.Add(lblInterval);
            groupMode.Controls.Add(numInterval);

            // 定时设置
            Label lblDate = new Label() { Text = "日期：", Location = new System.Drawing.Point(170, 85), Size = new System.Drawing.Size(40, 23) };
            dtpDate = new DateTimePicker()
            {
                Format = DateTimePickerFormat.Short,
                Location = new System.Drawing.Point(210, 85),
                Size = new System.Drawing.Size(110, 23)
            };
            groupMode.Controls.Add(lblDate);
            groupMode.Controls.Add(dtpDate);

            Label lblTime = new Label() { Text = "时间：", Location = new System.Drawing.Point(330, 85), Size = new System.Drawing.Size(40, 23) };
            dtpTime = new DateTimePicker()
            {
                Format = DateTimePickerFormat.Time,
                ShowUpDown = true,
                Location = new System.Drawing.Point(370, 85),
                Size = new System.Drawing.Size(90, 23)
            };
            groupMode.Controls.Add(lblTime);
            groupMode.Controls.Add(dtpTime);

            chkConvertToInterval = new CheckBox()
            {
                Text = "定时重启后转为固定间隔重启",
                Location = new System.Drawing.Point(170, 115),
                Size = new System.Drawing.Size(250, 24)
            };
            groupMode.Controls.Add(chkConvertToInterval);

            // 事件联动：点击定时模式才启用日期时间选择器和复选框
            rdoScheduled.CheckedChanged += (s, e) =>
            {
                bool isScheduled = rdoScheduled.Checked;
                dtpDate.Enabled = isScheduled;
                dtpTime.Enabled = isScheduled;
                chkConvertToInterval.Enabled = isScheduled;
            };
            rdoInterval.CheckedChanged += (s, e) =>
            {
                numInterval.Enabled = rdoInterval.Checked;
            };

            // 默认开启间隔模式
            rdoInterval.Checked = true;

            // ====== 启动程序列表 ======
            top = 195;
            Label lblApps = new Label() { Text = "重启后自动启动的程序：", Location = new System.Drawing.Point(leftMargin, top), Size = new System.Drawing.Size(200, 23) };
            this.Controls.Add(lblApps);

            lstApps = new ListBox() { Location = new System.Drawing.Point(leftMargin, top + 25), Size = new System.Drawing.Size(340, 160), SelectionMode = SelectionMode.MultiExtended };
            this.Controls.Add(lstApps);

            btnAdd = new Button() { Text = "添加程序...", Location = new System.Drawing.Point(380, top + 25), Size = new System.Drawing.Size(120, 30) };
            btnAdd.Click += BtnAdd_Click;
            this.Controls.Add(btnAdd);

            btnRemove = new Button() { Text = "删除选中", Location = new System.Drawing.Point(380, top + 65), Size = new System.Drawing.Size(120, 30) };
            btnRemove.Click += BtnRemove_Click;
            this.Controls.Add(btnRemove);

            btnUp = new Button() { Text = "上移", Location = new System.Drawing.Point(380, top + 105), Size = new System.Drawing.Size(120, 30) };
            btnUp.Click += BtnUp_Click;
            this.Controls.Add(btnUp);

            btnDown = new Button() { Text = "下移", Location = new System.Drawing.Point(380, top + 145), Size = new System.Drawing.Size(120, 30) };
            btnDown.Click += BtnDown_Click;
            this.Controls.Add(btnDown);

            chkAutoStart = new CheckBox() { Text = "开机自动运行本软件", Location = new System.Drawing.Point(leftMargin, top + 200), Size = new System.Drawing.Size(200, 30) };
            this.Controls.Add(chkAutoStart);

            // ====== 保存/取消按钮 ======
            btnSave = new Button() { Text = "保存", Location = new System.Drawing.Point(180, top + 240), Size = new System.Drawing.Size(80, 30) };
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            btnCancel = new Button() { Text = "取消", Location = new System.Drawing.Point(280, top + 240), Size = new System.Drawing.Size(80, 30) };
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);

            LoadSettingsFromRegistry();
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "可执行文件|*.exe|所有文件|*.*";
                ofd.Title = "选择要启动的程序（可多选）";
                ofd.Multiselect = true;
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    foreach (string file in ofd.FileNames)
                    {
                        if (!lstApps.Items.Contains(file) && File.Exists(file))
                            lstApps.Items.Add(file);
                        else if (!File.Exists(file))
                            MessageBox.Show($"文件不存在，已跳过：{file}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            for (int i = lstApps.SelectedIndices.Count - 1; i >= 0; i--)
                lstApps.Items.RemoveAt(lstApps.SelectedIndices[i]);
        }

        private void BtnUp_Click(object sender, EventArgs e)
        {
            if (lstApps.SelectedIndex <= 0) return;
            int idx = lstApps.SelectedIndex;
            object item = lstApps.SelectedItem;
            lstApps.Items.RemoveAt(idx);
            lstApps.Items.Insert(idx - 1, item);
            lstApps.SelectedIndex = idx - 1;
        }

        private void BtnDown_Click(object sender, EventArgs e)
        {
            if (lstApps.SelectedIndex < 0 || lstApps.SelectedIndex >= lstApps.Items.Count - 1) return;
            int idx = lstApps.SelectedIndex;
            object item = lstApps.SelectedItem;
            lstApps.Items.RemoveAt(idx);
            lstApps.Items.Insert(idx + 1, item);
            lstApps.SelectedIndex = idx + 1;
        }

        private void LoadSettingsFromRegistry()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_KEY))
            {
                if (key != null)
                {
                    int mode = (int)(key.GetValue("RebootMode", 1));
                    rdoDisabled.Checked = mode == 0;
                    rdoInterval.Checked = mode == 1;
                    rdoScheduled.Checked = mode == 2;

                    numInterval.Value = (int)(key.GetValue("IntervalHours", 72));
                    long ticks = Convert.ToInt64(key.GetValue("ScheduledTime", 0L));
                    if (ticks > 0)
                    {
                        DateTime scheduled = new DateTime(ticks, DateTimeKind.Utc).ToLocalTime();
                        dtpDate.Value = scheduled.Date;
                        dtpTime.Value = scheduled;
                    }
                    else
                    {
                        dtpDate.Value = DateTime.Now.AddDays(1);
                        dtpTime.Value = DateTime.Now;
                    }

                    chkConvertToInterval.Checked = (int)(key.GetValue("ConvertToInterval", 0) ?? 0) == 1;

                    string apps = key.GetValue("TargetApps")?.ToString() ?? "";
                    string[] paths = apps.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    lstApps.Items.Clear();
                    foreach (string path in paths) lstApps.Items.Add(path);
                }
            }

            using (RegistryKey runKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run"))
            {
                chkAutoStart.Checked = runKey?.GetValue("AutoRebootApp") != null;
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            int mode = rdoDisabled.Checked ? 0 : (rdoInterval.Checked ? 1 : 2);
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY))
            {
                key.SetValue("RebootMode", mode);
                key.SetValue("IntervalHours", (int)numInterval.Value);

                // 定时模式：保存合并后的日期时间（UTC ticks）
                if (mode == 2)
                {
                    DateTime schedule = dtpDate.Value.Date + dtpTime.Value.TimeOfDay;
                    key.SetValue("ScheduledTime", schedule.ToUniversalTime().Ticks);
                }
                else
                {
                    key.DeleteValue("ScheduledTime", false);
                }
                key.SetValue("ConvertToInterval", chkConvertToInterval.Checked ? 1 : 0);

                // 保存程序列表
                List<string> appList = new List<string>();
                foreach (string item in lstApps.Items) appList.Add(item);
                key.SetValue("TargetApps", string.Join("|", appList));
            }

            // 开机自启
            SetAutoStart(chkAutoStart.Checked);

            MessageBox.Show("设置已保存", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void SetAutoStart(bool enable)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true))
            {
                if (enable)
                    key.SetValue("AutoRebootApp", "\"" + Application.ExecutablePath + "\"");
                else
                    key.DeleteValue("AutoRebootApp", false);
            }
        }
    }

    // ===================== 日志查看窗体 =====================
    public class LogForm : Form
    {
        private TextBox txtLog;
        private string logFilePath;

        public LogForm(string logPath)
        {
            this.logFilePath = logPath;
            this.Text = "运行日志";
            this.Size = new System.Drawing.Size(600, 450);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            txtLog = new TextBox()
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Top,
                Height = 350,
                Font = new System.Drawing.Font("Consolas", 9)
            };
            this.Controls.Add(txtLog);

            Button btnRefresh = new Button() { Text = "刷新", Location = new System.Drawing.Point(150, 365), Size = new System.Drawing.Size(80, 30) };
            btnRefresh.Click += (s, e) => LoadLog();
            this.Controls.Add(btnRefresh);

            Button btnClear = new Button() { Text = "清空日志", Location = new System.Drawing.Point(250, 365), Size = new System.Drawing.Size(80, 30) };
            btnClear.Click += (s, e) =>
            {
                try { File.WriteAllText(logFilePath, string.Empty); LoadLog(); }
                catch { MessageBox.Show("清空日志失败"); }
            };
            this.Controls.Add(btnClear);

            Button btnClose = new Button() { Text = "关闭", Location = new System.Drawing.Point(370, 365), Size = new System.Drawing.Size(80, 30) };
            btnClose.Click += (s, e) => this.Close();
            this.Controls.Add(btnClose);

            LoadLog();
        }

        private void LoadLog()
        {
            try
            {
                if (File.Exists(logFilePath))
                    txtLog.Text = File.ReadAllText(logFilePath);
                else
                    txtLog.Text = "暂无日志。";
                txtLog.SelectionStart = txtLog.Text.Length;
                txtLog.ScrollToCaret();
            }
            catch (Exception ex)
            {
                txtLog.Text = "读取日志失败: " + ex.Message;
            }
        }
    }
}