namespace ZOfit.Installer;

public sealed class MainForm : Form
{
    private readonly InstallService _svc = new();
    private readonly Label _lblTitle;
    private readonly Label _lblStatus;
    private readonly TextBox _txtLog;
    private readonly Button _btnInstall;
    private readonly Button _btnUninstall;
    private readonly Button _btnClose;

    public MainForm()
    {
        Text = "zOfit QS Pro v2.1 - Cài đặt 1-Click";
        Size = new Size(560, 410);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        BackColor = Color.FromArgb(246, 248, 252);
        Font = new Font("Segoe UI", 9.5f);

        var header = new Panel
        {
            Dock = DockStyle.Top,
            Height = 72,
            BackColor = Color.FromArgb(22, 54, 92)
        };
        _lblTitle = new Label
        {
            Text = "zOfit QS Pro v2.1 (Hybrid Excel Add-in + Engine)",
            ForeColor = Color.White,
            Font = new Font("Segoe UI", 13f, FontStyle.Bold),
            Location = new Point(18, 12),
            AutoSize = true
        };
        var lblSub = new Label
        {
            Text = "Nhập 1 lần QS_Data & QS_Spec  →  Tự sinh Detail, TKT Thép, Hoàn thiện, Hạ tầng & Link BOQ",
            ForeColor = Color.FromArgb(210, 226, 248),
            Font = new Font("Segoe UI", 8.8f),
            Location = new Point(18, 42),
            AutoSize = true
        };
        header.Controls.Add(_lblTitle);
        header.Controls.Add(lblSub);
        Controls.Add(header);

        _lblStatus = new Label
        {
            Location = new Point(18, 86),
            Size = new Size(506, 42),
            ForeColor = Color.FromArgb(31, 78, 120),
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
        };
        Controls.Add(_lblStatus);

        _txtLog = new TextBox
        {
            Location = new Point(18, 134),
            Size = new Size(506, 165),
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            BackColor = Color.White,
            Font = new Font("Consolas", 9f)
        };
        Controls.Add(_txtLog);

        _btnInstall = new Button
        {
            Text = "Cài đặt / Cập nhật (1-Click)",
            Location = new Point(18, 314),
            Size = new Size(210, 38),
            BackColor = Color.FromArgb(47, 85, 151),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold)
        };
        _btnInstall.Click += OnInstallClick;
        Controls.Add(_btnInstall);

        _btnUninstall = new Button
        {
            Text = "Gỡ cài đặt",
            Location = new Point(240, 314),
            Size = new Size(135, 38),
            FlatStyle = FlatStyle.Flat
        };
        _btnUninstall.Click += OnUninstallClick;
        Controls.Add(_btnUninstall);

        _btnClose = new Button
        {
            Text = "Đóng",
            Location = new Point(388, 314),
            Size = new Size(136, 38),
            FlatStyle = FlatStyle.Flat
        };
        _btnClose.Click += (_, _) => Close();
        Controls.Add(_btnClose);

        RefreshStatus();
    }

    private void RefreshStatus()
    {
        bool installed = _svc.IsAlreadyInstalled();
        bool excelOn = _svc.IsExcelRunning();
        string s1 = installed
            ? "Trạng thái: Đã cài đặt trên máy (sẵn sàng cập nhật lên v2.1)."
            : "Trạng thái: Sẵn sàng cài đặt mới v2.1 (không cần quyền Admin).";
        string s2 = excelOn
            ? "Cảnh báo: Microsoft Excel đang chạy — hãy đóng Excel trước khi bấm Cài đặt!"
            : "Excel đã đóng — có thể bấm Cài đặt ngay.";
        _lblStatus.Text = s1 + Environment.NewLine + s2;
        _lblStatus.ForeColor = excelOn ? Color.Firebrick : Color.FromArgb(31, 78, 120);
        _btnUninstall.Enabled = installed;
    }

    private void AppendLog(string line)
    {
        _txtLog.AppendText(line + Environment.NewLine);
    }

    private void OnInstallClick(object? sender, EventArgs e)
    {
        _txtLog.Clear();
        RefreshStatus();
        try
        {
            _svc.Install(AppendLog);
            RefreshStatus();
            MessageBox.Show(
                "Cài đặt zOfit QS Pro v2.1 thành công!\n\nMở Microsoft Excel và bấm tab 'zOfit QS Pro' trên thanh Ribbon để bắt đầu.",
                "zOfit QS Pro Setup", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            AppendLog("LỖI: " + ex.Message);
            MessageBox.Show(ex.Message, "zOfit QS Pro Setup", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private void OnUninstallClick(object? sender, EventArgs e)
    {
        if (MessageBox.Show("Bạn có chắc muốn gỡ bỏ Add-in zOfit QS Pro khỏi Excel?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        _txtLog.Clear();
        try
        {
            _svc.Uninstall(AppendLog);
            RefreshStatus();
            MessageBox.Show("Đã gỡ cài đặt zOfit QS Pro hoàn tất.", "zOfit QS Pro Setup", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            AppendLog("LỖI: " + ex.Message);
            MessageBox.Show(ex.Message, "zOfit QS Pro Setup", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
