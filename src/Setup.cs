using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

internal static class Setup
{
    private const string Product="Code Studio for Word";
    private const string ProgAddin="CodeStudio.WordAddin";
    private const string ProgPane="CodeStudio.WordCodePane";
    private const string ClassAddin="CodeStudio.Connect";
    private const string ClassPane="CodeStudio.WordCodePane";
    private const string ClsidAddin="{9C2044B8-1E42-4A1B-BB63-A419C81BA11C}";
    private const string ClsidPane="{28BB16BD-801F-4D92-80CB-7D527E2A8680}";
    private const string ManagedCategory="{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}";
    private const string OfficeKey=@"Software\Microsoft\Office\Word\Addins\CodeStudio.WordAddin";
    private static readonly string InstallDir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"CodeStudioWord");
    private static readonly string DllPath=Path.Combine(InstallDir,"CodeStudio.Addin.dll");

    [STAThread]
    private static int Main(string[] args)
    {
        bool quiet=Array.Exists(args,a=>string.Equals(a,"/quiet",StringComparison.OrdinalIgnoreCase));
        bool uninstall=Array.Exists(args,a=>string.Equals(a,"/uninstall",StringComparison.OrdinalIgnoreCase));
        bool verify=Array.Exists(args,a=>string.Equals(a,"/verify",StringComparison.OrdinalIgnoreCase));
        try
        {
            if(verify) { bool ok=File.Exists(DllPath) && Registry.CurrentUser.OpenSubKey(OfficeKey)!=null; Report(quiet,ok?"已安装："+DllPath:"未安装。",ok); return ok?0:2; }
            if(Process.GetProcessesByName("WINWORD").Length>0) throw new InvalidOperationException("请先保存文档并关闭 Word，然后重试。安装或卸载后需重新打开 Word。");
            if(uninstall) Remove(); else Install();
            Report(quiet,uninstall?"已卸载。重新打开 Word 后按钮将消失。":"安装完成。重新打开 Word，在“代码美化”选项卡中点击“插入美化代码”。",true);
            return 0;
        }
        catch(Exception ex) { Report(quiet,(uninstall?"卸载":"安装")+"失败："+ex.Message,false); return 1; }
    }

    private static void Report(bool quiet,string message,bool success)
    {
        if(quiet) { Console.WriteLine(message); return; }
        MessageBox.Show(message,Product,MessageBoxButtons.OK,success?MessageBoxIcon.Information:MessageBoxIcon.Error);
    }

    private static void Install()
    {
        Directory.CreateDirectory(InstallDir);
        using(Stream input=Assembly.GetExecutingAssembly().GetManifestResourceStream("CodeStudio.Addin.dll"))
        {
            if(input==null) throw new InvalidOperationException("安装包缺少插件组件。");
            using(FileStream output=File.Create(DllPath)) input.CopyTo(output);
        }
        string assembly=AssemblyName.GetAssemblyName(DllPath).FullName;
        string codebase=new Uri(DllPath).AbsoluteUri;
        RegisterClass(ProgAddin,ClsidAddin,ClassAddin,assembly,codebase);
        RegisterClass(ProgPane,ClsidPane,ClassPane,assembly,codebase);
        using(var key=Registry.CurrentUser.CreateSubKey(OfficeKey))
        {
            key.SetValue("FriendlyName","代码美化插入",RegistryValueKind.String);
            key.SetValue("Description","在 Word 侧边栏输入并美化代码，插入光标位置。",RegistryValueKind.String);
            key.SetValue("LoadBehavior",3,RegistryValueKind.DWord);
        }
    }

    private static void RegisterClass(string progId,string clsid,string className,string assembly,string codebase)
    {
        const string classes=@"Software\Classes\";
        using(var p=Registry.CurrentUser.CreateSubKey(classes+progId)) p.SetValue(null,className);
        using(var p=Registry.CurrentUser.CreateSubKey(classes+progId+@"\CLSID")) p.SetValue(null,clsid);
        string baseKey=classes+"CLSID\\"+clsid;
        using(var p=Registry.CurrentUser.CreateSubKey(baseKey)) p.SetValue(null,className);
        using(var p=Registry.CurrentUser.CreateSubKey(baseKey+@"\InprocServer32"))
        {
            p.SetValue(null,"mscoree.dll");
            p.SetValue("ThreadingModel","Both");
            p.SetValue("Class",className);
            p.SetValue("Assembly",assembly);
            p.SetValue("RuntimeVersion","v4.0.30319");
            p.SetValue("CodeBase",codebase);
        }
        using(var p=Registry.CurrentUser.CreateSubKey(baseKey+@"\InprocServer32\1.0.0.0"))
        {
            p.SetValue("Class",className);
            p.SetValue("Assembly",assembly);
            p.SetValue("RuntimeVersion","v4.0.30319");
            p.SetValue("CodeBase",codebase);
        }
        using(var p=Registry.CurrentUser.CreateSubKey(baseKey+@"\ProgId")) p.SetValue(null,progId);
        Registry.CurrentUser.CreateSubKey(baseKey+@"\Implemented Categories\"+ManagedCategory).Dispose();
    }

    private static void Remove()
    {
        Registry.CurrentUser.DeleteSubKeyTree(OfficeKey,false);
        const string classes=@"Software\Classes\";
        Registry.CurrentUser.DeleteSubKeyTree(classes+"CLSID\\"+ClsidAddin,false);
        Registry.CurrentUser.DeleteSubKeyTree(classes+"CLSID\\"+ClsidPane,false);
        Registry.CurrentUser.DeleteSubKeyTree(classes+ProgAddin,false);
        Registry.CurrentUser.DeleteSubKeyTree(classes+ProgPane,false);
        if(File.Exists(DllPath)) File.Delete(DllPath);
        if(Directory.Exists(InstallDir) && Directory.GetFileSystemEntries(InstallDir).Length==0) Directory.Delete(InstallDir);
    }
}
