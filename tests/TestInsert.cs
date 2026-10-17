using System;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

internal static class TestInsert
{
    [STAThread]
    private static int Main()
    {
        if (Process.GetProcessesByName("WINWORD").Length != 0) throw new Exception("Word is already running.");
        Word.Application word = (Word.Application)Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
        Word.Document doc = null;
        try
        {
            word.Visible = false;
            doc = word.Documents.Add();
            doc.Content.Text = "BEFOREAFTER";
            var pane = new CodeStudio.WordCodePane();
            pane.SetWordApplication(word);
            var flags = BindingFlags.Instance | BindingFlags.NonPublic;
            ((TextBox)typeof(CodeStudio.WordCodePane).GetField("code", flags).GetValue(pane)).Text =
                "first\n\n" + string.Join("\n", new[] { "2", "3", "4", "5", "6", "7", "8", "9", "10" });
            doc.Range(6, 6).Select();
            typeof(CodeStudio.WordCodePane).GetMethod("InsertCode", flags).Invoke(pane, null);
            string status = ((Label)typeof(CodeStudio.WordCodePane).GetField("status", flags).GetValue(pane)).Text;
            Console.WriteLine("status=" + status);
            Console.WriteLine("tables=" + doc.Tables.Count);
            Word.Table table = doc.Tables[1];
            Console.WriteLine("rows=" + table.Rows.Count);
            Console.WriteLine("gutterWidth=" + table.Columns[1].Width);
            Console.WriteLine("gutterWrap=" + table.Cell(10, 1).WordWrap);
            Console.WriteLine("dividerWidth=" + table.Cell(10, 1).Borders[Word.WdBorderType.wdBorderRight].LineWidth);
            if (table.Rows.Count != 10 || table.Columns[1].Width < 40 || table.Cell(10, 1).WordWrap || !status.StartsWith("已在光标位置插入")) return 1;
            doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            doc = word.Documents.Add();
            ((ComboBox)typeof(CodeStudio.WordCodePane).GetField("family", flags).GetValue(pane)).SelectedItem = "Easy Code Formatter";
            doc.Range(0, 0).Select();
            typeof(CodeStudio.WordCodePane).GetMethod("InsertCode", flags).Invoke(pane, null);
            status = ((Label)typeof(CodeStudio.WordCodePane).GetField("status", flags).GetValue(pane)).Text;
            Console.WriteLine("easyStatus=" + status);
            Console.WriteLine("easyRows=" + doc.Tables[1].Rows.Count);
            if (!status.StartsWith("已在光标位置插入") || doc.Tables[1].Rows.Count != 11) return 1;
            doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            doc = word.Documents.Add();
            doc.Content.Text = "alpha\rbeta\r";
            doc.Range(0, doc.Content.End - 1).Select();
            typeof(CodeStudio.WordCodePane).GetMethod("LoadSelection", flags).Invoke(pane, null);
            string selectedCode = ((TextBox)typeof(CodeStudio.WordCodePane).GetField("code", flags).GetValue(pane)).Text;
            Console.WriteLine("selectionLines=" + selectedCode.Split(new[] { "\r\n" }, StringSplitOptions.None).Length);
            if (!selectedCode.Contains("alpha\r\nbeta")) return 1;
            return 0;
        }
        finally
        {
            if (doc != null) doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            word.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
        }
    }
}
