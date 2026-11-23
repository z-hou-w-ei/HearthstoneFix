// 炉石一键修复 单文件启动器
// 作用：把内嵌的脚本/配置/图标释放到 %LOCALAPPDATA%\HearthstoneFix，再以普通权限启动 bat
//      （bat 内部会自己用 -Verb RunAs 请求管理员权限，已实测可用）
// 编译：见同目录 build.ps1（.NET Framework 4.x 自带 csc，无需额外安装）

using System;
using System.IO;
using System.Reflection;
using System.Diagnostics;

class HearthstoneFixLauncher
{
    const string AppName = "HearthstoneFix";

    static int Main(string[] args)
    {
        bool extractOnly = false;
        string argline = "";
        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--extract-only" || args[i] == "-ExtractOnly") { extractOnly = true; continue; }
            if (argline.Length > 0) argline += " ";
            argline += args[i];
        }

        string root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppName);
        string iconDir = Path.Combine(root, "icon");

        StreamWriter log = null;
        try
        {
            Directory.CreateDirectory(root);
            Directory.CreateDirectory(iconDir);
            log = new StreamWriter(Path.Combine(root, "launcher.log"), false, new System.Text.UTF8Encoding(false));
        }
        catch { }

        // 脚本、入口、文档、图标：每次覆盖，保证与当前 exe 一致
        SafeWrite(log, "res_ps1", Path.Combine(root, "HearthstoneFix.ps1"));
        SafeWrite(log, "res_md", Path.Combine(root, "README.md"));
        SafeWrite(log, "res_ico", Path.Combine(iconDir, "HearthstoneFix.ico"));
        // config.ini：只在不存在时写，保留用户自己的路径配置
        string cfg = Path.Combine(root, "config.ini");
        if (!File.Exists(cfg)) SafeWrite(log, "res_ini", cfg);

        // 入口最后写，避免文件占用导致前面半套更新
        bool batOk = SafeWrite(log, "res_bat", Path.Combine(root, "一键修复.bat"));

        string bat = Path.Combine(root, "一键修复.bat");
        if (!File.Exists(bat))
        {
            LogLine(log, "entry missing: " + bat);
            CloseLog(log);
            return 2;
        }
        if (!batOk) LogLine(log, "entry locked, using existing one");

        if (extractOnly)
        {
            LogLine(log, "extract-only ok: " + root);
            CloseLog(log);
            return 0;
        }

        try
        {
            ProcessStartInfo psi = new ProcessStartInfo(bat, argline);
            psi.WorkingDirectory = root;
            psi.UseShellExecute = true;
            Process.Start(psi);
            LogLine(log, "started: " + bat + " " + argline);
            CloseLog(log);
            return 0;
        }
        catch (Exception ex)
        {
            LogLine(log, "start failed: " + ex.Message);
            CloseLog(log);
            return 1;
        }
    }

    static bool SafeWrite(StreamWriter log, string resName, string destPath)
    {
        string tmp = destPath + ".new";
        try
        {
            Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(resName);
            if (s == null) { LogLine(log, "resource missing: " + resName); return false; }
            using (s)
            {
                using (FileStream fs = new FileStream(tmp, FileMode.Create, FileAccess.Write))
                {
                    byte[] buf = new byte[8192];
                    int n;
                    while ((n = s.Read(buf, 0, buf.Length)) > 0) fs.Write(buf, 0, n);
                }
            }
            if (File.Exists(destPath))
            {
                try { File.Delete(destPath); }
                catch
                {
                    LogLine(log, "locked, keep old: " + destPath);
                    try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                    return false;
                }
            }
            File.Move(tmp, destPath);
            return true;
        }
        catch (Exception ex)
        {
            LogLine(log, "write failed " + destPath + ": " + ex.Message);
            try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
            return false;
        }
    }

    static void LogLine(StreamWriter log, string msg)
    {
        if (log == null) return;
        try { log.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + msg); log.Flush(); } catch { }
    }

    static void CloseLog(StreamWriter log)
    {
        if (log == null) return;
        try { log.Close(); } catch { }
    }
}
