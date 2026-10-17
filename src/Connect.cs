using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace CodeStudio
{
    [ComVisible(true)]
    [Guid("9F07B07E-220B-4B84-AD18-53942CD8F28D")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ICodeStudioRibbonCallbacks
    {
        [DispId(1)] void ShowPane(Office.IRibbonControl control);
        [DispId(2)] object GetRibbonImage(Office.IRibbonControl control);
    }

    [ComVisible(true)]
    [Guid("9C2044B8-1E42-4A1B-BB63-A419C81BA11C")]
    [ProgId("CodeStudio.WordAddin")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(ICodeStudioRibbonCallbacks))]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer, ICodeStudioRibbonCallbacks
    {
        private Word.Application word;
        private Office.ICTPFactory factory;
        private Office.CustomTaskPane pane;

        private static void Log(string message)
        {
            try
            {
                string path=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),"CodeStudioWord","addin.log");
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.AppendAllText(path,DateTime.Now.ToString("s")+" "+message+Environment.NewLine);
            }
            catch { }
        }

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        { word = Application as Word.Application; Log("OnConnection word="+(word!=null)); }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        { try { if (pane != null) pane.Delete(); } catch { } pane=null; factory=null; word=null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        { factory=CTPFactoryInst; Log("CTPFactoryAvailable factory="+(factory!=null)); }

        public string GetCustomUI(string RibbonID)
        {
            Log("GetCustomUI ribbon="+RibbonID);
            return "<customUI xmlns=\"http://schemas.microsoft.com/office/2006/01/customui\"><ribbon><tabs><tab id=\"CodeStudioTab\" label=\"代码美化\"><group id=\"CodeStudioGroup\" label=\"代码工具\"><button id=\"CodeStudioOpen\" label=\"插入美化代码\" size=\"large\" getImage=\"GetRibbonImage\" onAction=\"ShowPane\" screentip=\"打开代码美化侧边栏\"/></group></tab></tabs></ribbon></customUI>";
        }

        public object GetRibbonImage(Office.IRibbonControl control)
        {
            try { using (var bitmap=CreateLogoBitmap(64)) { object picture=RibbonPictureConverter.Convert(bitmap); Log("GetRibbonImage success"); return picture; } }
            catch(Exception ex) { Log("GetRibbonImage error="+ex); return null; }
        }

        private static Bitmap CreateLogoBitmap(int pixels)
        {
            var bitmap=new Bitmap(pixels,pixels);
            using (var g=Graphics.FromImage(bitmap))
            {
                g.SmoothingMode=SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                g.ScaleTransform(pixels/64f,pixels/64f);
                using(var path=RoundedSquare(2,2,60,60,13))
                using(var fill=new LinearGradientBrush(new Point(0,0),new Point(64,64),Color.FromArgb(77,106,239),Color.FromArgb(38,57,150)))
                    g.FillPath(fill,path);
                using(var accent=new Pen(Color.FromArgb(90,221,240),5f))
                using(var white=new Pen(Color.White,5f))
                {
                    accent.StartCap=accent.EndCap=LineCap.Round;
                    white.StartCap=white.EndCap=LineCap.Round;
                    g.DrawLines(accent,new[]{new Point(22,22),new Point(12,32),new Point(22,42)});
                    g.DrawLines(white,new[]{new Point(42,22),new Point(52,32),new Point(42,42)});
                    g.DrawLine(white,36,19,28,45);
                }
            }
            return bitmap;
        }

        private static GraphicsPath RoundedSquare(int x,int y,int w,int h,int radius)
        {
            var path=new GraphicsPath();
            int d=radius*2;
            path.AddArc(x,y,d,d,180,90);
            path.AddArc(x+w-d,y,d,d,270,90);
            path.AddArc(x+w-d,y+h-d,d,d,0,90);
            path.AddArc(x,y+h-d,d,d,90,90);
            path.CloseFigure();
            return path;
        }

        private sealed class RibbonPictureConverter : AxHost
        {
            private RibbonPictureConverter() : base("") { }
            internal static object Convert(Image image) { return GetIPictureDispFromPicture(image); }
        }

        public void ShowPane(Office.IRibbonControl control)
        {
            Log("ShowPane word="+(word!=null)+" factory="+(factory!=null));
            try
            {
                if (word==null || factory==null) throw new InvalidOperationException("Word 尚未完成插件加载。请关闭并重新打开 Word。 ");
                bool create = pane==null;
                if (!create) { try { bool visible=pane.Visible; } catch { create=true; } }
                if (create)
                {
                    pane=factory.CreateCTP("CodeStudio.WordCodePane","代码美化插入",Type.Missing);
                    pane.DockPosition=Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                    pane.Width=430;
                    ((WordCodePane)pane.ContentControl).SetWordApplication(word);
                    Log("CreateCTP success");
                }
                pane.Visible=true;
                Log("ShowPane visible="+pane.Visible);
            }
            catch (Exception ex) { Log("ShowPane error="+ex); MessageBox.Show("无法打开代码侧边栏：\n"+ex.Message,"代码美化插入",MessageBoxButtons.OK,MessageBoxIcon.Error); }
        }
    }
}
