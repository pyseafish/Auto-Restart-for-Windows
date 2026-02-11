using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace AutoRebootApp
{
    // 应用程序入口
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

    // 主窗体：托盘图标、定时重启逻辑（完全隐藏）
    public class Form1 : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private System.Windows.Forms.Timer checkTimer;
        private Mutex mutex;

        private const string REG_KEY = @"Software\AutoRebootApp";
        private const string LAST_REBOOT = "LastReboot";
        private const string INTERVAL_HOURS = "IntervalHours";
        private const string TARGET_APPS = "TargetApps";

        public Form1()
        {
            // 窗体完全隐藏，永不显示
            this.WindowState = FormWindowState.Minimized;
            this.ShowInTaskbar = false;
            this.Load += (s, e) => this.Hide();

            // 托盘菜单：移除了“显示主窗口”
            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("设置", null, OpenSettings);
            trayMenu.Items.Add("立即重启", null, RebootNow);
            trayMenu.Items.Add("退出", null, ExitApp);

            // 托盘图标
            trayIcon = new NotifyIcon();
            trayIcon.Text = "自动重启工具";
            trayIcon.Icon = SystemIcons.Application; // 可替换为自定义图标
            trayIcon.ContextMenuStrip = trayMenu;
            trayIcon.Visible = true;
            trayIcon.DoubleClick += OpenSettings; // 双击直接打开设置

            // 单实例检查
            bool createdNew;
            mutex = new Mutex(true, "AutoRebootApp_UniqueMutex", out createdNew);
            if (!createdNew)
            {
                MessageBox.Show("程序已在运行！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Environment.Exit(0);
            }

            // 初始化注册表默认值（仅当键不存在时）
            EnsureInitialConfig();

            // 启动定时器（每小时检查一次）
            StartTimer();
        }

        // 初始化注册表默认值（不设置开机自启，由用户通过界面控制）
        private void EnsureInitialConfig()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY))
            {
                if (key.GetValue(INTERVAL_HOURS) == null)
                    key.SetValue(INTERVAL_HOURS, 72);
                if (key.GetValue(LAST_REBOOT) == null)
                {
                    // 第一次运行，将上次重启时间设为当前时间前71小时，1小时后第一次检查
                    DateTime firstRun = DateTime.UtcNow.AddHours(-71);
                    key.SetValue(LAST_REBOOT, firstRun.Ticks.ToString());
                }
                if (key.GetValue(TARGET_APPS) == null)
                    key.SetValue(TARGET_APPS, "");
            }
        }

        // 启动定时器（1小时间隔）
        private void StartTimer()
        {
            checkTimer = new System.Windows.Forms.Timer();
            checkTimer.Interval = 3600000; // 1小时
            checkTimer.Tick += CheckRebootCondition;
            checkTimer.Start();
        }

        // 检查是否需要重启
        private void CheckRebootCondition(object sender, EventArgs e)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_KEY))
            {
                if (key == null) return;

                string lastTicks = key.GetValue(LAST_REBOOT)?.ToString();
                if (lastTicks == null) return;
                DateTime lastReboot = new DateTime(long.Parse(lastTicks), DateTimeKind.Utc);

                int interval = (int)key.GetValue(INTERVAL_HOURS, 72);
                if ((DateTime.UtcNow - lastReboot).TotalHours >= interval)
                {
                    PerformReboot();
                }
            }
        }

        // 执行重启
        private void PerformReboot()
        {
            // 1. 更新上次重启时间
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY))
            {
                key.SetValue(LAST_REBOOT, DateTime.UtcNow.Ticks.ToString());
            }

            // 2. 从注册表读取多个程序路径，写入 RunOnce
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_KEY))
            {
                string apps = key?.GetValue(TARGET_APPS)?.ToString();
                if (!string.IsNullOrWhiteSpace(apps))
                {
                    string[] appPaths = apps.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    RegistryKey runOnce = Registry.CurrentUser.OpenSubKey(
                        @"Software\Microsoft\Windows\CurrentVersion\RunOnce", true);
                    if (runOnce != null)
                    {
                        for (int i = 0; i < appPaths.Length; i++)
                        {
                            // 每个程序使用不同的键名，避免覆盖
                            runOnce.SetValue($"AutoStartApp_{i}", "\"" + appPaths[i] + "\"");
                        }
                    }
                }
            }

            // 3. 重启系统
            Process.Start("shutdown", "/r /t 0");
        }

        // 立即重启
        private void RebootNow(object sender, EventArgs e) => PerformReboot();

        // 打开设置窗口
        private void OpenSettings(object sender, EventArgs e)
        {
            using (Form2 settingsForm = new Form2())
            {
                settingsForm.ShowDialog();
            }
        }

        // 退出程序
        private void ExitApp(object sender, EventArgs e)
        {
            trayIcon.Visible = false;
            Application.Exit();
        }

        // 禁止用户关闭主窗体（隐藏而非退出）
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

        // 释放资源
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                trayIcon?.Dispose();
                checkTimer?.Dispose();
                mutex?.ReleaseMutex();
                mutex?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    // 设置窗体：支持多程序启动 + 开机自启动开关
    public class Form2 : Form
    {
        private NumericUpDown numInterval;
        private ListBox lstApps;
        private Button btnAdd, btnRemove, btnUp, btnDown;
        private CheckBox chkAutoStart;   // 开机自启复选框
        private Button btnSave, btnCancel;

        private const string REG_KEY = @"Software\AutoRebootApp";
        private const string TARGET_APPS = "TargetApps";

        public Form2()
        {
            // 窗体属性
            this.Text = "设置 - 自动重启工具";
            this.Size = new System.Drawing.Size(500, 450);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            // 重启间隔
            Label lblInterval = new Label()
            {
                Text = "重启间隔（小时）：",
                Location = new System.Drawing.Point(20, 20),
                Size = new System.Drawing.Size(120, 23)
            };
            numInterval = new NumericUpDown()
            {
                Location = new System.Drawing.Point(150, 18),
                Size = new System.Drawing.Size(80, 23),
                Minimum = 1,
                Maximum = 720,
                Value = 72
            };
            this.Controls.Add(lblInterval);
            this.Controls.Add(numInterval);

            // 程序列表标签
            Label lblApps = new Label()
            {
                Text = "重启后自动启动的程序：",
                Location = new System.Drawing.Point(20, 60),
                Size = new System.Drawing.Size(200, 23)
            };
            this.Controls.Add(lblApps);

            // 程序列表
            lstApps = new ListBox()
            {
                Location = new System.Drawing.Point(20, 90),
                Size = new System.Drawing.Size(300, 180),
                SelectionMode = SelectionMode.MultiExtended
            };
            this.Controls.Add(lstApps);

            // 右侧操作按钮
            btnAdd = new Button()
            {
                Text = "添加程序...",
                Location = new System.Drawing.Point(340, 90),
                Size = new System.Drawing.Size(120, 30)
            };
            btnAdd.Click += BtnAdd_Click;
            this.Controls.Add(btnAdd);

            btnRemove = new Button()
            {
                Text = "删除选中",
                Location = new System.Drawing.Point(340, 130),
                Size = new System.Drawing.Size(120, 30)
            };
            btnRemove.Click += BtnRemove_Click;
            this.Controls.Add(btnRemove);

            btnUp = new Button()
            {
                Text = "上移",
                Location = new System.Drawing.Point(340, 170),
                Size = new System.Drawing.Size(120, 30)
            };
            btnUp.Click += BtnUp_Click;
            this.Controls.Add(btnUp);

            btnDown = new Button()
            {
                Text = "下移",
                Location = new System.Drawing.Point(340, 210),
                Size = new System.Drawing.Size(120, 30)
            };
            btnDown.Click += BtnDown_Click;
            this.Controls.Add(btnDown);

            // 开机自启动复选框
            chkAutoStart = new CheckBox()
            {
                Text = "开机自动运行本软件",
                Location = new System.Drawing.Point(20, 290),
                Size = new System.Drawing.Size(200, 30)
            };
            this.Controls.Add(chkAutoStart);

            // 底部按钮
            btnSave = new Button()
            {
                Text = "保存",
                Location = new System.Drawing.Point(150, 340),
                Size = new System.Drawing.Size(80, 30)
            };
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            btnCancel = new Button()
            {
                Text = "取消",
                Location = new System.Drawing.Point(250, 340),
                Size = new System.Drawing.Size(80, 30)
            };
            btnCancel.Click += (s, e) => this.Close();
            this.Controls.Add(btnCancel);

            // 加载现有设置
            LoadSettings();
        }

        // 添加程序（多选）
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
                        if (!lstApps.Items.Contains(file))
                            lstApps.Items.Add(file);
                    }
                }
            }
        }

        // 删除选中项（支持多选）
        private void BtnRemove_Click(object sender, EventArgs e)
        {
            for (int i = lstApps.SelectedIndices.Count - 1; i >= 0; i--)
            {
                lstApps.Items.RemoveAt(lstApps.SelectedIndices[i]);
            }
        }

        // 上移
        private void BtnUp_Click(object sender, EventArgs e)
        {
            if (lstApps.SelectedIndex <= 0) return;
            int idx = lstApps.SelectedIndex;
            object item = lstApps.SelectedItem;
            lstApps.Items.RemoveAt(idx);
            lstApps.Items.Insert(idx - 1, item);
            lstApps.SelectedIndex = idx - 1;
        }

        // 下移
        private void BtnDown_Click(object sender, EventArgs e)
        {
            if (lstApps.SelectedIndex < 0 || lstApps.SelectedIndex >= lstApps.Items.Count - 1) return;
            int idx = lstApps.SelectedIndex;
            object item = lstApps.SelectedItem;
            lstApps.Items.RemoveAt(idx);
            lstApps.Items.Insert(idx + 1, item);
            lstApps.SelectedIndex = idx + 1;
        }

        // 加载注册表中的设置
        private void LoadSettings()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(REG_KEY))
            {
                if (key != null)
                {
                    // 加载间隔
                    numInterval.Value = (int)key.GetValue("IntervalHours", 72);

                    // 加载多程序列表
                    string apps = key.GetValue(TARGET_APPS)?.ToString() ?? "";
                    string[] paths = apps.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    lstApps.Items.Clear();
                    foreach (string path in paths)
                    {
                        lstApps.Items.Add(path);
                    }
                }
            }

            // 检查当前开机自启状态
            using (RegistryKey runKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run"))
            {
                chkAutoStart.Checked = runKey?.GetValue("AutoRebootApp") != null;
            }
        }

        // 保存设置
        private void BtnSave_Click(object sender, EventArgs e)
        {
            // 保存间隔和程序列表到注册表
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(REG_KEY))
            {
                key.SetValue("IntervalHours", (int)numInterval.Value);

                List<string> appList = new List<string>();
                foreach (string item in lstApps.Items)
                    appList.Add(item);
                key.SetValue(TARGET_APPS, string.Join("|", appList));
            }

            // 根据复选框状态设置或取消开机自启
            SetAutoStart(chkAutoStart.Checked);

            MessageBox.Show("设置已保存", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }

        // 设置开机自启动（写入或删除 Run 键）
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
}