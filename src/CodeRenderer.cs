using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace CodeStudio
{
    [ComVisible(false)]
    public sealed class RenderOptions
    {
        public string Family = "CodeInWord";
        public string Theme = "Default";
        public string Language = "自动识别";
        public int Size = 11;
        public bool LineNumbers = true;
        public bool SkipBlankLines = true;
    }

    public static class CodeRenderer
    {
        private sealed class Palette
        {
            public string Paper, Gutter, Text, Keyword, String, Comment, Number, Border, Header, HeaderText;
            public Palette(string paper, string gutter, string text, string keyword, string str, string comment, string number, string border, string header, string headerText)
            { Paper=paper; Gutter=gutter; Text=text; Keyword=keyword; String=str; Comment=comment; Number=number; Border=border; Header=header; HeaderText=headerText; }
        }

        public static readonly string[] Languages = { "自动识别", "纯文本", "Python", "JavaScript", "TypeScript", "Java", "C/C++", "C#", "CSS", "HTML/XML", "JSON", "SQL", "Bash/Shell", "PowerShell", "PHP", "Ruby", "Go", "Rust", "Swift", "Kotlin", "R", "MATLAB", "Markdown" };
        public static readonly string[] CodeInWordThemes = { "Default", "Eclipse", "Django", "Midnight", "Emacs", "FadeToGrey", "MDUltra", "RDark" };
        public static readonly string[] EasyThemes = { "Default", "Desert", "A11Y Dark", "A11Y Light" };

        private static readonly Dictionary<string,Palette> Palettes = new Dictionary<string,Palette>(StringComparer.OrdinalIgnoreCase)
        {
            {"CodeInWord/Default", new Palette("#FFFFFF","#F2F4F8","#243042","#0066B8","#A31515","#6A737D","#098658","#DBE2EA","#FFFFFF","#243042")},
            {"CodeInWord/Eclipse", new Palette("#FFFFFF","#EEF1F4","#24292F","#7F0055","#2A00FF","#3F7F5F","#1C6B25","#CBD5E1","#FFFFFF","#24292F")},
            {"CodeInWord/Django", new Palette("#112C21","#193D2D","#E5F6E9","#FFB86C","#E6DB74","#83B692","#A9DC76","#315A43","#112C21","#E5F6E9")},
            {"CodeInWord/Midnight", new Palette("#172033","#1E2B40","#E6EDF7","#82AAFF","#C3E88D","#7791AC","#F78C6C","#334155","#172033","#E6EDF7")},
            {"CodeInWord/Emacs", new Palette("#FFFFFF","#F3F3F3","#242424","#0000FF","#8B2252","#B22222","#008B8B","#D5D5D5","#FFFFFF","#242424")},
            {"CodeInWord/FadeToGrey", new Palette("#FAFAFA","#EFEFEF","#383838","#505050","#777777","#999999","#555555","#DDDDDD","#FAFAFA","#383838")},
            {"CodeInWord/MDUltra", new Palette("#FFFFFF","#EAF0F8","#203047","#3B5BDB","#D9485F","#7292A8","#0B9068","#CED9EA","#FFFFFF","#203047")},
            {"CodeInWord/RDark", new Palette("#202020","#2A2A2A","#F3F3F3","#FFD75F","#87D7FF","#A8A8A8","#FF8787","#4A4A4A","#202020","#F3F3F3")},
            {"Easy Code Formatter/Default", new Palette("#F8FAFC","#EDF2F7","#253041","#7053B7","#C2410C","#718096","#047857","#C9D5E3","#E8EEF5","#334155")},
            {"Easy Code Formatter/Desert", new Palette("#302820","#3A2F25","#F8EAD7","#FFD18C","#E6B8A2","#BDAE9A","#F1C27D","#665244","#3A2F25","#F8EAD7")},
            {"Easy Code Formatter/A11Y Dark", new Palette("#1C1C1C","#292929","#F5F5F5","#FFE66D","#7FDBFF","#B4B4B4","#FF9B85","#5C5C5C","#303030","#FFFFFF")},
            {"Easy Code Formatter/A11Y Light", new Palette("#FFFFFF","#F2F2F2","#1B1B1B","#5B2A86","#A32626","#5B5B5B","#006644","#A0A0A0","#E5E5E5","#1B1B1B")}
        };

        private static readonly HashSet<string> Keywords = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        static CodeRenderer()
        {
            foreach (string word in ("if else for while do break continue return try catch finally throw true false null undefined new class public private protected static const let var function import export from as async await yield switch case default this super extends implements interface enum void int float double bool boolean string number def lambda pass in is and or not None self print include using namespace template typename auto struct union typedef virtual override delete nullptr operator package func type map make range select where insert update create table join group by order asc desc limit distinct into values set echo fi done end fn mut impl trait pub module local foreach nil").Split(' ')) Keywords.Add(word);
        }

        private static string Esc(string value)
        {
            return value.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace("\"","&quot;").Replace("'","&#39;").Replace(" ","&nbsp;");
        }
        private static void Span(StringBuilder output, string text, string color)
        {
            if (text.Length > 0) output.Append("<span style=\"color:").Append(color).Append(";\">").Append(Esc(text)).Append("</span>");
        }
        private static bool IsOneOf(string value, params string[] choices) { foreach(string choice in choices) if(string.Equals(value,choice,StringComparison.OrdinalIgnoreCase)) return true; return false; }
        private static bool HasHashComment(string language) { return IsOneOf(language,"Python","Bash/Shell","PowerShell","Ruby","R","MATLAB"); }
        private static bool HasSlashComment(string language) { return IsOneOf(language,"JavaScript","TypeScript","Java","C/C++","C#","CSS","PHP","Go","Rust","Swift","Kotlin"); }

        public static string DetectLanguage(string code)
        {
            if (Regex.IsMatch(code, @"^\s*<\w", RegexOptions.Multiline)) return "HTML/XML";
            if (Regex.IsMatch(code, @"^\s*[\[{]", RegexOptions.Multiline) && code.Contains("\":")) return "JSON";
            if (Regex.IsMatch(code, @"\b(SELECT|INSERT INTO|CREATE TABLE|UPDATE)\b", RegexOptions.IgnoreCase)) return "SQL";
            if (Regex.IsMatch(code, @"^\s*(def |from \w+ import |class \w+:)", RegexOptions.Multiline)) return "Python";
            if (Regex.IsMatch(code, @"\b(const |let |function |console\.log|=>)", RegexOptions.Multiline)) return "JavaScript";
            if (Regex.IsMatch(code, @"^\s*(#include|int main\s*\()", RegexOptions.Multiline)) return "C/C++";
            return "纯文本";
        }

        private static string HighlightLine(string line, string language, Palette p, ref bool inBlock)
        {
            if (language == "纯文本" || language == "Markdown") return "<span style=\"color:" + p.Text + ";\">" + Esc(line) + "</span>";
            var b = new StringBuilder();
            for (int i=0; i<line.Length;)
            {
                if (inBlock)
                {
                    int end = line.IndexOf("*/", i, StringComparison.Ordinal);
                    if (end < 0) { Span(b,line.Substring(i),p.Comment); break; }
                    Span(b,line.Substring(i,end+2-i),p.Comment); i=end+2; inBlock=false; continue;
                }
                if (i+1<line.Length && line[i]=='/' && line[i+1]=='*' && HasSlashComment(language)) { inBlock=true; continue; }
                if ((line[i]=='#' && HasHashComment(language)) || (i+1<line.Length && line.Substring(i,2)=="//" && HasSlashComment(language)) || (i+1<line.Length && line.Substring(i,2)=="--" && language=="SQL")) { Span(b,line.Substring(i),p.Comment); break; }
                char ch=line[i];
                if (ch=='\"' || ch=='\'' || (ch=='`' && (language=="JavaScript" || language=="TypeScript")))
                {
                    int start=i++; while(i<line.Length) { if(line[i]=='\\') { i=Math.Min(line.Length,i+2); continue; } if(line[i++]==ch) break; }
                    Span(b,line.Substring(start,i-start),p.String); continue;
                }
                if (char.IsDigit(ch) && (i==0 || !char.IsLetterOrDigit(line[i-1])))
                {
                    int start=i++; while(i<line.Length && (char.IsDigit(line[i]) || line[i]=='.')) i++;
                    Span(b,line.Substring(start,i-start),p.Number); continue;
                }
                if (char.IsLetter(ch) || ch=='_' || ch=='$')
                {
                    int start=i++; while(i<line.Length && (char.IsLetterOrDigit(line[i]) || line[i]=='_' || line[i]=='$')) i++;
                    string word=line.Substring(start,i-start);
                    Span(b,word,Keywords.Contains(word)?p.Keyword:p.Text); continue;
                }
                Span(b,ch.ToString(),p.Text); i++;
            }
            return b.Length==0 ? "&nbsp;" : b.ToString();
        }

        public static string Render(string code, RenderOptions options, bool preview=false)
        {
            string family = options.Family == "Easy Code Formatter" ? options.Family : "CodeInWord";
            string key=family+"/"+options.Theme;
            Palette p=Palettes.ContainsKey(key)?Palettes[key]:Palettes[family+"/Default"];
            string language=options.Language=="自动识别"?DetectLanguage(code):options.Language;
            int size=Math.Max(8,Math.Min(36,options.Size));
            string font="font-family:Consolas,'Courier New',monospace;font-size:"+size+"pt;line-height:1.35;";
            string[] sourceLines=code.Replace("\r\r\n","\r\n").Replace("\r\n","\n").Replace('\r','\n').Replace("\t","    ").Split('\n');
            var visibleLines=new List<string>();
            foreach(string line in sourceLines) if(!options.SkipBlankLines || !string.IsNullOrWhiteSpace(line)) visibleLines.Add(line);
            if(!options.SkipBlankLines && visibleLines.Count>1 && visibleLines[visibleLines.Count-1].Length==0) visibleLines.RemoveAt(visibleLines.Count-1);
            if(visibleLines.Count==0) visibleLines.Add("");
            string[] lines=visibleLines.ToArray();
            int cols=options.LineNumbers?2:1;
            var b=new StringBuilder("<html><head><meta charset=\"utf-8\">");
            if(preview) b.Append("<style>html,body{margin:0;padding:0;overflow-x:hidden;}table{table-layout:fixed;width:100%;}td.gutter{width:38px;min-width:38px;max-width:38px;}td.code{white-space:pre-wrap!important;word-wrap:break-word;overflow-wrap:break-word;}</style>");
            b.Append("</head><body>");
            b.Append("<table cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse:collapse;width:100%;border:1px solid ").Append(p.Border).Append(";margin:5px 0 7px 0;\">");
            if(family=="Easy Code Formatter") b.Append("<tr><td colspan=\"").Append(cols).Append("\" style=\"background-color:").Append(p.Header).Append(";color:").Append(p.HeaderText).Append(";font-family:Arial,sans-serif;font-size:9pt;font-weight:bold;padding:7px 10px;border-bottom:1px solid ").Append(p.Border).Append(";\">CODE&nbsp;&nbsp;·&nbsp;&nbsp;").Append(Esc(language)).Append("</td></tr>");
            bool inBlock=false;
            for(int i=0;i<lines.Length;i++)
            {
                b.Append("<tr>");
                int gutterWidth=preview?38:42;
                if(options.LineNumbers) b.Append("<td class=\"gutter\" width=\"").Append(gutterWidth).Append("\" nowrap=\"nowrap\" style=\"width:").Append(gutterWidth).Append("px;min-width:").Append(gutterWidth).Append("px;").Append(font).Append("background-color:").Append(p.Gutter).Append(";color:").Append(p.Comment).Append(";text-align:right;vertical-align:top;padding:").Append(preview?"4px 5px;":"2px 9px;").Append("border-right:").Append(family=="CodeInWord"?"4px solid ":"1px solid ").Append(family=="CodeInWord"?p.Comment:p.Border).Append(";white-space:nowrap;\">").Append(i+1).Append("</td>");
                string highlighted=HighlightLine(lines[i],language,p,ref inBlock);
                if(preview) highlighted=highlighted.Replace("&nbsp;"," ");
                b.Append("<td class=\"code\" style=\"").Append(font).Append("background-color:").Append(p.Paper).Append(";color:").Append(p.Text).Append(";vertical-align:top;padding:").Append(preview?"4px 8px;white-space:pre-wrap;word-wrap:break-word;":"2px 10px;white-space:normal;").Append("\">").Append(highlighted).Append("</td></tr>");
            }
            return b.Append("</table></body></html>").ToString();
        }
    }
}
