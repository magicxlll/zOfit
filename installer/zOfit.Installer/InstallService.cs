using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace ZOfit.Installer;

public sealed class InstallService
{
    private const string XlamResource = "ZOfit.Installer.Payload.zOfit_QS_Pro.xlam";
    private const string EngineResource = "ZOfit.Installer.Payload.zofit-engine.exe";
    private const string XlamFileName = "zOfit_QS_Pro.xlam";

    public string AddInTargetPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "Microsoft", "AddIns", XlamFileName);

    public string EngineDirPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "zOfitQSPro");

    public string EngineTargetPath => Path.Combine(EngineDirPath, "zofit-engine.exe");

    public bool IsExcelRunning() => Process.GetProcessesByName("EXCEL").Length > 0;

    public bool IsAlreadyInstalled() => File.Exists(AddInTargetPath) && File.Exists(EngineTargetPath);

    public void Install(Action<string>? log = null)
    {
        if (IsExcelRunning())
        {
            throw new InvalidOperationException("Microsoft Excel đang mở! Vui lòng lưu file và đóng hoàn toàn tất cả cửa sổ Excel trước khi cài đặt.");
        }

        log?.Invoke("1. Chuẩn bị thư mục cài đặt...");
        Directory.CreateDirectory(Path.GetDirectoryName(AddInTargetPath)!);
        Directory.CreateDirectory(EngineDirPath);

        log?.Invoke("2. Giải nén Add-in zOfit_QS_Pro.xlam...");
        ExtractResource(XlamResource, AddInTargetPath);
        UnblockFile(AddInTargetPath);

        log?.Invoke("3. Giải nén Engine zofit-engine.exe...");
        ExtractResource(EngineResource, EngineTargetPath);
        UnblockFile(EngineTargetPath);

        log?.Invoke("4. Đăng ký tự động mở Add-in trong Microsoft Excel (HKCU)...");
        var versions = DetectInstalledOfficeVersions();
        foreach (var ver in versions)
        {
            RegisterExcelAddIn(ver);
            RegisterTrustedLocation(ver, Path.GetDirectoryName(AddInTargetPath)!);
        }

        File.WriteAllText(Path.Combine(EngineDirPath, "install_info.txt"),
            $"zOfit QS Pro v2.1.0\r\nInstalledAt={DateTime.Now:O}\r\nAddIn={AddInTargetPath}\r\nEngine={EngineTargetPath}\r\n");

        log?.Invoke("HOÀN TẤT! Mở Microsoft Excel và xem tab 'zOfit QS Pro' trên thanh Ribbon.");
    }

    public void Uninstall(Action<string>? log = null)
    {
        if (IsExcelRunning())
        {
            throw new InvalidOperationException("Microsoft Excel đang mở! Vui lòng đóng hoàn toàn Excel trước khi gỡ cài đặt.");
        }

        log?.Invoke("1. Hủy đăng ký Add-in khỏi Excel Options...");
        foreach (var ver in DetectInstalledOfficeVersions())
        {
            UnregisterExcelAddIn(ver);
        }

        log?.Invoke("2. Xóa tệp Add-in và Engine cục bộ...");
        if (File.Exists(AddInTargetPath)) File.Delete(AddInTargetPath);
        if (Directory.Exists(EngineDirPath)) Directory.Delete(EngineDirPath, true);

        log?.Invoke("Đã gỡ cài đặt zOfit QS Pro sạch sẽ khỏi máy.");
    }

    private static void ExtractResource(string resName, string destPath)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resName)
            ?? throw new FileNotFoundException($"Missing embedded payload: {resName}");
        using var fs = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None);
        stream.CopyTo(fs);
    }

    private static List<string> DetectInstalledOfficeVersions()
    {
        var list = new List<string>();
        using var officeKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office");
        if (officeKey != null)
        {
            foreach (var v in new[] { "16.0", "15.0", "14.0" })
            {
                if (officeKey.OpenSubKey(v + @"\Excel") != null)
                    list.Add(v);
            }
        }
        if (list.Count == 0) list.Add("16.0");
        return list;
    }

    private void RegisterExcelAddIn(string officeVer)
    {
        string optPath = $@"Software\Microsoft\Office\{officeVer}\Excel\Options";
        using var key = Registry.CurrentUser.CreateSubKey(optPath);
        if (key == null) return;

        // Check existing OPEN, OPEN1, OPEN2... entries
        var existingNames = key.GetValueNames()
            .Where(n => n.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var name in existingNames)
        {
            string val = key.GetValue(name)?.ToString() ?? "";
            if (val.Contains(XlamFileName, StringComparison.OrdinalIgnoreCase))
            {
                key.SetValue(name, $"\"{XlamFileName}\"", RegistryValueKind.String);
                return;
            }
        }

        // Find free slot: OPEN, OPEN1, OPEN2...
        string slot = "OPEN";
        int idx = 1;
        while (existingNames.Contains(slot, StringComparer.OrdinalIgnoreCase))
        {
            slot = $"OPEN{idx++}";
        }
        key.SetValue(slot, $"\"{XlamFileName}\"", RegistryValueKind.String);

        // Ensure not present in Add-in Manager disabled/deleted entries
        string mgrPath = $@"Software\Microsoft\Office\{officeVer}\Excel\Add-in Manager";
        using var mgrKey = Registry.CurrentUser.OpenSubKey(mgrPath, true);
        if (mgrKey != null)
        {
            foreach (var vn in mgrKey.GetValueNames())
            {
                if (vn.Contains(XlamFileName, StringComparison.OrdinalIgnoreCase))
                    mgrKey.DeleteValue(vn, false);
            }
        }
    }

    private static void RegisterTrustedLocation(string officeVer, string folderPath)
    {
        string locPath = $@"Software\Microsoft\Office\{officeVer}\Excel\Security\Trusted Locations\zOfitQSPro";
        using var key = Registry.CurrentUser.CreateSubKey(locPath);
        if (key == null) return;
        key.SetValue("Path", folderPath, RegistryValueKind.String);
        key.SetValue("AllowSubfolders", 1, RegistryValueKind.DWord);
        key.SetValue("Description", "zOfit QS Pro Add-Ins", RegistryValueKind.String);
    }

    private void UnregisterExcelAddIn(string officeVer)
    {
        string optPath = $@"Software\Microsoft\Office\{officeVer}\Excel\Options";
        using var key = Registry.CurrentUser.OpenSubKey(optPath, true);
        if (key == null) return;
        foreach (var name in key.GetValueNames().Where(n => n.StartsWith("OPEN", StringComparison.OrdinalIgnoreCase)).ToList())
        {
            string val = key.GetValue(name)?.ToString() ?? "";
            if (val.Contains(XlamFileName, StringComparison.OrdinalIgnoreCase))
                key.DeleteValue(name, false);
        }
    }

    [DllImport("kernel32", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern bool DeleteFileW(string name);

    private static void UnblockFile(string path)
    {
        try { DeleteFileW(path + ":Zone.Identifier"); } catch { }
    }
}
