using System;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using Word = Microsoft.Office.Interop.Word;

namespace CodeStudio
{
    [ComVisible(true)]
    [Guid("28BB16BD-801F-4D92-80CB-7D527E2A8680")]
    [ProgId("CodeStudio.WordCodePane")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public sealed class WordCodePane : UserControl
    {
        private Word.Application word;
        private readonly TextBox code;
        private readonly ComboBox family, language, size, theme;
        private readonly CheckBox lineNumbers, skipBlankLines;
        private readonly WebBrowser preview;
        private readonly Label status;
        private readonly Timer previewTimer;
        private const string SettingsKey="Software\\CodeStudioWord\\Settings";

        public WordCodePane()
        {
            BackColor=Color.FromArgb(241,244,249);
            Font=new Font("Microsoft YaHei UI",9F);
            var scroll=new Panel { Dock=DockStyle.Fill, AutoScroll=true, BackColor=BackColor };
            Controls.Add(scroll);
            var root=new TableLayoutPanel { Dock=DockStyle.Top, AutoSize=true, AutoSizeMode=AutoSizeMode.GrowAndShrink, ColumnCount=1, Padding=new Padding(12,9,12,12), BackColor=BackColor };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,100f));
            scroll.Controls.Add(root);
            AddRow(root,MakeLabel("输入代码"));
            code=new TextBox { Multiline=true, AcceptsReturn=true, AcceptsTab=true, ScrollBars=ScrollBars.Vertical, WordWrap=true, Height=200, Font=new Font("Consolas",9.5F), BackColor=Color.White };
            AddRow(root,code);
            Button load=new Button { Text="读取 Word 选区", Height=32, FlatStyle=FlatStyle.Flat, BackColor=Color.White };
            load.Click += (s,e) => LoadSelection(); AddRow(root,load);
            AddRow(root,MakeLabel("外观系列"));
            family=MakeCombo(new[]{"CodeInWord","Easy Code Formatter"}); AddRow(root,family);
            AddRow(root,MakeLabel("编程语言"));
            language=MakeCombo(CodeRenderer.Languages); AddRow(root,language);
            AddRow(root,MakeLabel("字号"));
            size=MakeCombo(new[]{"8","9","10","11","12","14","16","18","20","24","28","36"}); AddRow(root,size);
            AddRow(root,MakeLabel("配色主题"));
            theme=MakeCombo(CodeRenderer.CodeInWordThemes); AddRow(root,theme);
            var toggles=new TableLayoutPanel { ColumnCount=2, Height=32, Margin=new Padding(0,5,0,8), Padding=new Padding(0) };
            toggles.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,45f));
            toggles.ColumnStyles.Add(new ColumnStyle(SizeType.Percent,55f));
            lineNumbers=new CheckBox { Text="显示行号", Checked=true, Dock=DockStyle.Fill, Margin=new Padding(0) }; toggles.Controls.Add(lineNumbers,0,0);
            skipBlankLines=new CheckBox { Text="去掉空白行", Checked=true, Dock=DockStyle.Fill, Margin=new Padding(0) }; toggles.Controls.Add(skipBlankLines,1,0);
            AddRow(root,toggles);
            AddRow(root,MakeLabel("插入预览"));
            preview=new WebBrowser { Height=230, ScriptErrorsSuppressed=true, IsWebBrowserContextMenuEnabled=false, WebBrowserShortcutsEnabled=false, AllowWebBrowserDrop=false };
            AddRow(root,preview);
            Button insert=new Button { Text="插入到光标位置", Height=40, BackColor=Color.FromArgb(56,86,232), ForeColor=Color.White, FlatStyle=FlatStyle.Flat, Font=new Font("Microsoft YaHei UI",9F,FontStyle.Bold) };
            insert.FlatAppearance.BorderSize=0; insert.Click += (s,e) => InsertCode(); AddRow(root,insert);
            status=new Label { Text="代码将在本机处理，插入后可继续编辑", ForeColor=Color.FromArgb(103,117,139), Height=36, AutoSize=false };
            AddRow(root,status);
            previewTimer=new Timer { Interval=180 };
            previewTimer.Tick += (s,e) => { previewTimer.Stop(); RefreshPreview(); };
            code.TextChanged += (s,e) => QueuePreview();
            family.SelectedIndexChanged += (s,e) => { ResetThemes(); QueuePreview(); };
            language.SelectedIndexChanged += (s,e) => QueuePreview();
            size.SelectedIndexChanged += (s,e) => QueuePreview();
            theme.SelectedIndexChanged += (s,e) => QueuePreview();
            lineNumbers.CheckedChanged += (s,e) => QueuePreview();
            skipBlankLines.CheckedChanged += (s,e) => QueuePreview();
            LoadSettings();
            RefreshPreview();
        }

        private static void AddRow(TableLayoutPanel root,Control control)
        {
            if(control.Margin==new Padding(3)) control.Margin=new Padding(0,0,0,6);
            int row=root.RowCount++;
            if(control is Label && control.AutoSize)
            {
                control.Dock=DockStyle.None;
                root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }
            else
            {
                control.Dock=DockStyle.Fill;
                root.RowStyles.Add(new RowStyle(SizeType.Absolute,control.Height+control.Margin.Vertical));
            }
            root.Controls.Add(control,0,row);
        }
        private static Label MakeLabel(string text) { return new Label { Text=text, Font=new Font("Microsoft YaHei UI",9F,FontStyle.Bold), Margin=new Padding(0,10,0,6), AutoSize=true }; }
        private static ComboBox MakeCombo(string[] values) { var c=new ComboBox { DropDownStyle=ComboBoxStyle.DropDownList, Height=30 }; c.Items.AddRange(values); c.SelectedIndex=0; return c; }
        private void ResetThemes()
        {
            if (theme==null || family==null) return;
            string previous=theme.SelectedItem as string;
            theme.Items.Clear(); theme.Items.AddRange(family.Text=="Easy Code Formatter"?CodeRenderer.EasyThemes:CodeRenderer.CodeInWordThemes);
            theme.SelectedItem=previous;
            if(theme.SelectedIndex<0) theme.SelectedIndex=0;
        }
        public void SetWordApplication(Word.Application app) { word=app; }
        private RenderOptions Options()
        {
            int fontSize; if(!int.TryParse(size.Text,out fontSize)) fontSize=11;
            return new RenderOptions { Family=family.Text, Theme=theme.Text, Language=language.Text, Size=fontSize, LineNumbers=lineNumbers.Checked, SkipBlankLines=skipBlankLines.Checked };
        }
        private void QueuePreview() { if(previewTimer==null) return; previewTimer.Stop(); previewTimer.Start(); }
        private void RefreshPreview()
        {
            if (preview==null || code==null || theme==null || theme.SelectedIndex<0) return;
            try { preview.DocumentText=code.TextLength==0?"<html><body style='font:12px Segoe UI;color:#8a96a7'>输入代码后可预览</body></html>":CodeRenderer.Render(code.Text,Options(),true); SaveSettings(); }
            catch (Exception ex) { status.Text="预览失败："+ex.Message; }
        }
        private void LoadSelection()
        {
            try { if(word==null) throw new InvalidOperationException("Word 尚未连接。"); string selected=word.Selection.Range.Text; if(string.IsNullOrEmpty(selected)) throw new InvalidOperationException("请先在 Word 中选中一段代码。"); code.Text=selected.TrimEnd('\r','\a').Replace("\r\n","\n").Replace('\r','\n').Replace("\n","\r\n"); status.Text="已读取选区；插入时接在选区末尾。"; }
            catch(Exception ex) { status.Text=ex.Message; }
        }
        private void InsertCode()
        {
            if (string.IsNullOrWhiteSpace(code.Text)) { status.Text="请先输入代码。"; return; }
            string tempFile=null;
            try
            {
                if(word==null || word.ActiveDocument==null) throw new InvalidOperationException("请先打开一个可编辑的 Word 文档。");
                string html=CodeRenderer.Render(code.Text,Options());
                tempFile=Path.Combine(Path.GetTempPath(),"CodeStudio-"+Guid.NewGuid().ToString("N")+".html");
                File.WriteAllText(tempFile,html,new UTF8Encoding(true));
                Word.Range range=word.Selection.Range.Duplicate;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                int insertionPoint=range.Start;
                Word.Document document=word.ActiveDocument;
                int previousTableCount=document.Tables.Count;
                object missing=Type.Missing;
                object confirm=false;
                range.InsertFile(tempFile,ref missing,ref confirm,ref missing,ref missing);
                try
                {
                    if(document.Tables.Count>previousTableCount)
                    {
                        Word.Table inserted=null;
                        int nearestStart=int.MaxValue;
                        foreach(Word.Table candidate in document.Tables)
                        {
                            int tableStart=candidate.Range.Start;
                            if(tableStart>=insertionPoint && tableStart<nearestStart) { inserted=candidate; nearestStart=tableStart; }
                        }
                        if(inserted!=null) FormatInsertedTable(inserted,Options());
                    }
                    status.Text="已在光标位置插入可编辑代码。";
                }
                catch(Exception ex) { status.Text="代码已插入；行号排版未完全应用："+ex.Message; }
            }
            catch (Exception ex) { status.Text="插入失败："+ex.Message; }
            finally { if(tempFile!=null) { try { File.Delete(tempFile); } catch { } } }
        }

        private static void FormatInsertedTable(Word.Table table, RenderOptions options)
        {
            table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed);
            table.Range.ParagraphFormat.SpaceBefore=0f;
            table.Range.ParagraphFormat.SpaceAfter=0f;
            if(!options.LineNumbers) return;
            foreach(Word.Row row in table.Rows)
            {
                if(row.Cells.Count!=2) continue;
                Word.Cell gutter=row.Cells[1];
                gutter.SetWidth(42f,Word.WdRulerStyle.wdAdjustFirstColumn);
                gutter.WordWrap=false;
                if(options.Family=="CodeInWord")
                {
                    Word.Border divider=gutter.Borders[Word.WdBorderType.wdBorderRight];
                    divider.LineStyle=Word.WdLineStyle.wdLineStyleSingle;
                    divider.LineWidth=Word.WdLineWidth.wdLineWidth300pt;
                }
            }
        }
        private void LoadSettings()
        {
            try { using(var key=Registry.CurrentUser.OpenSubKey(SettingsKey)) { if(key==null) return; SetCombo(family,key.GetValue("Family") as string); ResetThemes(); SetCombo(language,key.GetValue("Language") as string); SetCombo(size,key.GetValue("Size") as string); SetCombo(theme,key.GetValue("Theme") as string); lineNumbers.Checked=(key.GetValue("LineNumbers") as string)!="0"; skipBlankLines.Checked=(key.GetValue("SkipBlankLines") as string)!="0"; } } catch { }
        }
        private static void SetCombo(ComboBox combo,string value) { if(value!=null && combo.Items.Contains(value)) combo.SelectedItem=value; }
        private void SaveSettings()
        {
            try { using(var key=Registry.CurrentUser.CreateSubKey(SettingsKey)) { key.SetValue("Family",family.Text); key.SetValue("Language",language.Text); key.SetValue("Size",size.Text); key.SetValue("Theme",theme.Text); key.SetValue("LineNumbers",lineNumbers.Checked?"1":"0"); key.SetValue("SkipBlankLines",skipBlankLines.Checked?"1":"0"); } } catch { }
        }
    }
}
