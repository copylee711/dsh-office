// dsh-office helper: drives the running Word / Excel / PowerPoint through COM.
// One JSON command per stdin line, one JSON reply per stdout line.
// Must stay C# 5 (compiled at runtime by the in-box .NET Framework csc.exe).
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

class Fail : Exception
{
    public string Code;
    public Fail(string code, string message) : base(message) { Code = code; }
}

/// Arguments of a command or of one edit operation.
class Bag
{
    readonly Dictionary<string, object> d;
    readonly HashSet<string> read = new HashSet<string>();
    /// Fields that were given but that nothing looked at: a misnamed field must not pass silently.
    public string Unread()
    {
        List<string> names = new List<string>();
        foreach (string key in d.Keys) if (!read.Contains(key)) names.Add(key);
        return string.Join(", ", names.ToArray());
    }
    public Bag(object source)
    {
        d = source as Dictionary<string, object>;
        if (d == null) throw new Fail("BAD_ARGS", "Expected an object.");
    }
    public bool Has(string key) { return Raw(key) != null; }
    public object Raw(string key) { object v; read.Add(key); return d.TryGetValue(key, out v) ? v : null; }
    public string Str(string key, string fallback)
    {
        object v = Raw(key);
        return v == null ? fallback : Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture);
    }
    public string Need(string key)
    {
        string v = Str(key, null);
        if (v == null) throw new Fail("BAD_ARGS", "\"" + key + "\" is required.");
        return v;
    }
    public int Int(string key, int fallback)
    {
        object v = Raw(key);
        if (v == null) return fallback;
        try { return Convert.ToInt32(v, System.Globalization.CultureInfo.InvariantCulture); }
        catch { throw new Fail("BAD_ARGS", "\"" + key + "\" must be a whole number."); }
    }
    public double Num(string key, double fallback)
    {
        object v = Raw(key);
        if (v == null) return fallback;
        try { return Convert.ToDouble(v, System.Globalization.CultureInfo.InvariantCulture); }
        catch { throw new Fail("BAD_ARGS", "\"" + key + "\" must be a number."); }
    }
    public bool Flag(string key, bool fallback)
    {
        object v = Raw(key);
        return v is bool ? (bool)v : fallback;
    }
    public IList List(string key)
    {
        object v = Raw(key);
        if (v == null) return null;
        IList list = v as IList;
        if (list == null || v is string) { ArrayList one = new ArrayList(); one.Add(v); return one; }
        return list;
    }
}

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
}

/// Office rejects calls while the user has a dialog open or is typing: wait for it a few seconds.
class MessageFilter : IOleMessageFilter
{
    [DllImport("ole32.dll")]
    static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);
    public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new MessageFilter(), out old); }
    public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) { return 0; }
    public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
    {
        return dwRejectType == 2 && dwTickCount < 6000 ? 150 : -1;
    }
    public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) { return 2; }
}

static class Program
{
    static readonly JavaScriptSerializer Json = new JavaScriptSerializer();
    static TextWriter Out;

    [STAThread]
    static void Main()
    {
        Json.MaxJsonLength = int.MaxValue;
        Json.RecursionLimit = 64;
        Out = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false));
        TextReader input = new StreamReader(Console.OpenStandardInput(), new UTF8Encoding(false));
        MessageFilter.Register();
        Send(new Dictionary<string, object> { { "event", "ready" } });
        string line;
        while ((line = input.ReadLine()) != null)
        {
            if (line.Trim().Length == 0) continue;
            object id = null;
            Dictionary<string, object> reply = new Dictionary<string, object>();
            try
            {
                Bag args = new Bag(Json.DeserializeObject(line));
                id = args.Raw("id");
                reply["result"] = Run(args.Need("cmd"), args);
                reply["ok"] = true;
            }
            catch (Exception error)
            {
                string code;
                reply["ok"] = false;
                reply["error"] = Describe(error, out code);
                reply["code"] = code;
            }
            reply["id"] = id;
            Send(reply);
            // Let go of every COM object: a reference kept here would keep a closed Office app alive, hidden.
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
        }
    }

    static void Send(object message) { Out.WriteLine(Json.Serialize(message)); Out.Flush(); }

    static string Describe(Exception error, out string code)
    {
        while (error is System.Reflection.TargetInvocationException && error.InnerException != null) error = error.InnerException;
        Fail fail = error as Fail;
        if (fail != null) { code = fail.Code; return fail.Message; }
        COMException com = error as COMException;
        if (com != null)
        {
            uint hr = unchecked((uint)com.ErrorCode);
            if (hr == 0x80010001 || hr == 0x8001010A || hr == 0x800AC472)
            {
                code = "BUSY";
                return "The Office app is busy: the user is typing in a cell or has a dialog open. Nothing was changed by this call; try again in a moment.";
            }
            code = "OFFICE_ERROR";
            return com.Message.Trim() + " (0x" + hr.ToString("X8") + ")";
        }
        code = "ERROR";
        return error.Message;
    }

    // ───────────────────────── applications and documents ─────────────────────────

    static string ProgId(string kind)
    {
        if (kind == "word") return "Word.Application";
        if (kind == "excel") return "Excel.Application";
        if (kind == "ppt") return "PowerPoint.Application";
        throw new Fail("BAD_ARGS", "Unknown app \"" + kind + "\": use word, excel or ppt.");
    }

    static string AppName(string kind) { return kind == "word" ? "Word" : kind == "excel" ? "Excel" : "PowerPoint"; }

    static string KindOfPath(string path)
    {
        string ext = Path.GetExtension(path).ToLowerInvariant();
        if (ext == ".docx" || ext == ".doc" || ext == ".docm" || ext == ".rtf" || ext == ".dotx") return "word";
        if (ext == ".xlsx" || ext == ".xls" || ext == ".xlsm" || ext == ".csv" || ext == ".xlsb") return "excel";
        if (ext == ".pptx" || ext == ".ppt" || ext == ".pptm") return "ppt";
        return null;
    }

    static string Kind(Bag a)
    {
        string kind = a.Str("app", null);
        if (kind != null) { ProgId(kind); return kind; }
        string doc = a.Str("doc", a.Str("path", null));
        string guess = doc == null ? null : KindOfPath(doc);
        if (guess == null) throw new Fail("BAD_ARGS", "Say which app: \"app\" must be word, excel or ppt (it cannot be told from the file name).");
        return guess;
    }

    static dynamic Running(string kind)
    {
        try { return Marshal.GetActiveObject(ProgId(kind)); }
        catch (COMException) { return null; }
    }

    static dynamic App(string kind, bool start)
    {
        dynamic app = Running(kind);
        if (app != null) return app;
        if (!start) throw new Fail("NOT_RUNNING", AppName(kind) + " is not running. Open the file with office_open first.");
        Type type = Type.GetTypeFromProgID(ProgId(kind));
        if (type == null) throw new Fail("NOT_INSTALLED", AppName(kind) + " is not installed on this computer.");
        app = Activator.CreateInstance(type);
        // Started by automation, Excel would quit once we let go of it; hand it to the user.
        if (kind == "excel") { try { app.UserControl = true; } catch (Exception) { } }
        return app;
    }

    static void Show(string kind, dynamic app)
    {
        if (kind == "ppt") app.Visible = -1; else app.Visible = true;
    }

    static dynamic Docs(string kind, dynamic app)
    {
        return kind == "word" ? app.Documents : kind == "excel" ? app.Workbooks : app.Presentations;
    }

    static dynamic Active(string kind, dynamic app)
    {
        try
        {
            if ((int)Docs(kind, app).Count == 0) return null;
            return kind == "word" ? app.ActiveDocument : kind == "excel" ? app.ActiveWorkbook : app.ActivePresentation;
        }
        catch (COMException) { return null; }
    }

    static bool SamePath(string a, string b)
    {
        try { return string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase); }
        catch (Exception) { return false; }
    }

    static dynamic Find(string kind, dynamic app, string doc)
    {
        foreach (dynamic item in Docs(kind, app))
        {
            string name = (string)item.Name, full = (string)item.FullName;
            if (string.Equals(name, doc, StringComparison.OrdinalIgnoreCase)) return item;
            if (Path.IsPathRooted(doc) && full.IndexOf("://", StringComparison.Ordinal) < 0 && Path.IsPathRooted(full) && SamePath(full, doc)) return item;
        }
        return null;
    }

    /// The document a command is about: named by "doc" (path or window name), else the active one.
    static dynamic Doc(string kind, dynamic app, Bag a)
    {
        string doc = a.Str("doc", null);
        dynamic found = doc == null ? Active(kind, app) : Find(kind, app, doc);
        if (found == null)
        {
            throw new Fail("NO_DOCUMENT", doc == null
                ? "No document is open in " + AppName(kind) + ". Open one with office_open."
                : "\"" + doc + "\" is not open in " + AppName(kind) + ". Open it with office_open, or call office_status to see what is open.");
        }
        return found;
    }

    static bool Truthy(object value)
    {
        if (value is bool) return (bool)value;
        try { return Convert.ToInt32(value) != 0; } catch (Exception) { return false; }
    }

    static Dictionary<string, object> Info(string kind, dynamic doc, dynamic active)
    {
        Dictionary<string, object> info = new Dictionary<string, object>();
        string full = (string)doc.FullName;
        info["app"] = kind;
        info["name"] = (string)doc.Name;
        info["path"] = ((string)doc.Path).Length == 0 ? null : full;
        info["saved"] = Truthy(doc.Saved);
        info["active"] = active != null && (string)active.FullName == full;
        try { info["readOnly"] = Truthy(doc.ReadOnly); } catch (Exception) { }
        return info;
    }

    /// Excel refuses every call while a cell is being edited: wait for it briefly rather than fail half-way through a batch.
    static void WaitReady(string kind, dynamic app)
    {
        if (kind != "excel") return;
        for (int i = 0; i < 25; i++)
        {
            try { if ((bool)app.Ready) return; } catch (COMException) { }
            Thread.Sleep(200);
        }
        throw new Fail("BUSY", "Excel is busy: the user is typing in a cell or has a dialog open. Nothing was changed; try again in a moment.");
    }

    // ───────────────────────── commands ─────────────────────────

    static object Run(string cmd, Bag a)
    {
        if (cmd == "ping") return "pong";
        if (cmd == "status") return Status();
        string kind = Kind(a);
        if (cmd == "open") return Open(kind, a);
        if (cmd == "quit")
        {
            dynamic idle = Running(kind);
            if (idle == null) return "not running";
            if ((int)Docs(kind, idle).Count > 0) throw new Fail("BAD_ARGS", AppName(kind) + " still has documents open; it is left running.");
            idle.Quit();
            return "quit";
        }
        dynamic app = App(kind, false);
        WaitReady(kind, app);
        dynamic doc = Doc(kind, app, a);
        if (cmd == "read") return kind == "word" ? WordRead(doc, a) : kind == "excel" ? ExcelRead(doc, a) : PptRead(doc, a);
        if (cmd == "edit") return Edit(kind, app, doc, a);
        if (cmd == "render") return Render(kind, app, doc, a);
        if (cmd == "save") return Save(kind, app, doc, a);
        if (cmd == "close") return Close(kind, doc, a);
        if (cmd == "undo")
        {
            if (kind != "word") throw new Fail("BAD_ARGS", "Only Word keeps the agent's edits on its undo stack.");
            return (bool)doc.Undo(Math.Max(1, a.Int("times", 1))) ? "undone" : "nothing to undo";
        }
        throw new Fail("BAD_ARGS", "Unknown command \"" + cmd + "\".");
    }

    static object Status()
    {
        List<object> apps = new List<object>();
        foreach (string kind in new string[] { "word", "excel", "ppt" })
        {
            Dictionary<string, object> entry = new Dictionary<string, object>();
            entry["app"] = kind;
            entry["installed"] = Type.GetTypeFromProgID(ProgId(kind)) != null;
            List<object> open = new List<object>();
            dynamic app = Running(kind);
            entry["running"] = app != null;
            if (app != null)
            {
                try
                {
                    dynamic active = Active(kind, app);
                    foreach (dynamic doc in Docs(kind, app)) open.Add(Info(kind, doc, active));
                    entry["version"] = Convert.ToString(app.Version);
                }
                catch (COMException error) { string code; entry["error"] = Describe(error, out code); }
            }
            entry["documents"] = open;
            apps.Add(entry);
        }
        return apps;
    }

    static object Open(string kind, Bag a)
    {
        string path = a.Str("path", null);
        if (path != null)
        {
            if (!Path.IsPathRooted(path)) throw new Fail("BAD_ARGS", "\"path\" must be an absolute path.");
            path = Path.GetFullPath(path);
        }
        dynamic app = App(kind, true);
        Show(kind, app);
        WaitReady(kind, app);
        dynamic doc = path == null ? null : Find(kind, app, path);
        string how = "attached";
        if (doc == null)
        {
            dynamic docs = Docs(kind, app);
            if (path != null && File.Exists(path))
            {
                doc = kind == "ppt" ? docs.Open(path, 0, 0, -1) : docs.Open(path);
                how = "opened";
            }
            else
            {
                if (path != null && !Directory.Exists(Path.GetDirectoryName(path))) throw new Fail("BAD_ARGS", "The folder of \"" + path + "\" does not exist.");
                doc = kind == "ppt" ? docs.Add(-1) : docs.Add();
                how = "created";
                if (path != null) SaveAs(kind, app, doc, path);
            }
        }
        if (a.Flag("show", true))
        {
            try
            {
                if (kind == "ppt") doc.Windows[1].Activate(); else doc.Activate();
                if (kind == "word") app.Activate();
            }
            catch (COMException) { }
        }
        Dictionary<string, object> info = Info(kind, doc, Active(kind, app));
        info["how"] = how;
        return info;
    }

    static void SaveAs(string kind, dynamic app, dynamic doc, string path)
    {
        string ext = Path.GetExtension(path).ToLowerInvariant();
        if (kind == "word")
        {
            if (ext == ".pdf") doc.ExportAsFixedFormat(path, 17);
            else doc.SaveAs2(path, ext == ".doc" ? 0 : ext == ".rtf" ? 6 : ext == ".docm" ? 13 : 12);
        }
        else if (kind == "excel")
        {
            if (ext == ".pdf") { doc.ExportAsFixedFormat(0, path); return; }
            bool alerts = (bool)app.DisplayAlerts;
            app.DisplayAlerts = false;
            try { doc.SaveAs(path, ext == ".xls" ? 56 : ext == ".xlsm" ? 52 : ext == ".csv" ? 6 : ext == ".xlsb" ? 50 : 51); }
            finally { app.DisplayAlerts = alerts; }
        }
        else
        {
            if (ext == ".pdf") doc.SaveAs(path, 32); else doc.SaveAs(path);
        }
    }

    static object Save(string kind, dynamic app, dynamic doc, Bag a)
    {
        string path = a.Str("path", null);
        Dictionary<string, object> result = new Dictionary<string, object>();
        if (path == null)
        {
            if (((string)doc.Path).Length == 0) throw new Fail("NEEDS_PATH", "This document has never been saved: give \"path\" to say where.");
            doc.Save();
            result["path"] = (string)doc.FullName;
        }
        else
        {
            if (!Path.IsPathRooted(path)) throw new Fail("BAD_ARGS", "\"path\" must be an absolute path.");
            path = Path.GetFullPath(path);
            if (!Directory.Exists(Path.GetDirectoryName(path))) throw new Fail("BAD_ARGS", "The folder of \"" + path + "\" does not exist.");
            SaveAs(kind, app, doc, path);
            result["path"] = path;
        }
        result["bytes"] = new FileInfo((string)result["path"]).Length;
        return result;
    }

    static object Close(string kind, dynamic doc, Bag a)
    {
        bool save = a.Flag("save", false);
        if (kind == "word") doc.Close(save ? -1 : 0);
        else if (kind == "excel") doc.Close(save);
        else { if (save) doc.Save(); else doc.Saved = -1; doc.Close(); }
        return "closed";
    }

    // ───────────────────────── shared helpers ─────────────────────────

    static string Clean(string text)
    {
        if (text == null) return "";
        StringBuilder sb = new StringBuilder(text.Length);
        foreach (char c in text)
        {
            if (c == '\r' || c == '\a' || c == '\f') continue;
            if (c == '\v' || c == '\n') { sb.Append('\n'); continue; }
            if (c < ' ' && c != '\t') continue;
            sb.Append(c);
        }
        return sb.ToString();
    }

    static string Clip(string text, int max)
    {
        return text.Length <= max ? text : text.Substring(0, max) + "…(" + text.Length + " chars)";
    }

    static string Squash(string text)
    {
        StringBuilder sb = new StringBuilder();
        foreach (char c in text) if (!char.IsWhiteSpace(c)) sb.Append(c);
        return sb.ToString();
    }

    /// "#RRGGBB" to the BGR integer Office uses.
    static int Bgr(string color)
    {
        string hex = color.Trim().TrimStart('#');
        if (hex.Length == 3) hex = "" + hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
        int rgb;
        if (hex.Length != 6 || !int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, null, out rgb)) throw new Fail("BAD_ARGS", "Colour \"" + color + "\" must look like #RRGGBB.");
        return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF);
    }

    static string Lines(object value)
    {
        IList list = value as IList;
        string text;
        if (list != null && !(value is string))
        {
            List<string> parts = new List<string>();
            foreach (object item in list) parts.Add(Convert.ToString(item));
            text = string.Join("\r", parts.ToArray());
        }
        else text = Convert.ToString(value);
        return text.Replace("\r\n", "\r").Replace("\n", "\r");
    }

    // ───────────────────────── Word ─────────────────────────

    static readonly Dictionary<string, int> WordStyles = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        { "Normal", -1 }, { "Heading 1", -2 }, { "Heading 2", -3 }, { "Heading 3", -4 }, { "Heading 4", -5 }, { "Heading 5", -6 },
        { "Heading 6", -7 }, { "Title", -63 }, { "Subtitle", -75 }, { "List Bullet", -49 }, { "List Number", -50 },
        { "Quote", -181 }, { "Caption", -35 }, { "Body Text", -67 },
    };

    static void SetStyle(dynamic range, string style)
    {
        int builtin;
        try
        {
            if (WordStyles.TryGetValue(style, out builtin)) range.Style = builtin; else range.Style = style;
        }
        catch (COMException)
        {
            throw new Fail("STYLE_MISSING", "This document has no style \"" + style + "\". Built-in names that always work: " + string.Join(", ", new List<string>(WordStyles.Keys).ToArray()) + ".");
        }
    }

    static dynamic Para(dynamic doc, int index, string expect)
    {
        int total = (int)doc.Paragraphs.Count;
        if (index < 1 || index > total) throw new Fail("ANCHOR_MISSING", "There is no paragraph " + index + " (the document has " + total + "). Call office_read again.");
        dynamic p = doc.Paragraphs[index];
        if (expect != null)
        {
            string text = Squash(Clean((string)p.Range.Text)), want = Squash(expect);
            if (want.Length == 0 ? text.Length > 0 : !text.StartsWith(want, StringComparison.Ordinal))
            {
                throw new Fail("ANCHOR_MISMATCH", "Paragraph " + index + " now starts with \"" + Clip(Clean((string)p.Range.Text), 40) + "\", not " + (want.Length == 0 ? "empty" : "\"" + Clip(expect, 40) + "\"") + ". The document was edited since you read it: call office_read again and retry from this operation.");
            }
        }
        return p;
    }

    static object WordRead(dynamic doc, Bag a)
    {
        Dictionary<string, object> result = new Dictionary<string, object>();
        int total = (int)doc.Paragraphs.Count, tables = (int)doc.Tables.Count;
        result["paragraphs"] = total;
        result["tables"] = tables;
        try { result["pages"] = (int)doc.ComputeStatistics(2); } catch (COMException) { }
        if (a.Has("table"))
        {
            int n = a.Int("table", 1);
            if (n < 1 || n > tables) throw new Fail("ANCHOR_MISSING", "There is no table " + n + " (the document has " + tables + ").");
            dynamic table = doc.Tables[n];
            int rows = (int)table.Rows.Count, cols = (int)table.Columns.Count;
            List<object> grid = new List<object>();
            for (int r = 1; r <= rows; r++)
            {
                List<object> row = new List<object>();
                for (int c = 1; c <= cols; c++)
                {
                    try { row.Add(Clean((string)table.Cell(r, c).Range.Text)); }
                    catch (COMException) { row.Add(null); }
                }
                grid.Add(row);
            }
            result["table"] = n;
            result["cells"] = grid;
            return result;
        }
        int from = Math.Max(1, a.Int("from", 1)), count = Math.Max(1, Math.Min(400, a.Int("count", 120)));
        bool full = a.Flag("full", false);
        List<object> items = new List<object>();
        int shown = from - 1;
        string titleStyle = null;
        try { titleStyle = (string)doc.Styles[-63].NameLocal; } catch (Exception) { }
        if (from <= total)
        {
            dynamic p = doc.Paragraphs[from];
            for (int i = from; i < from + count && i <= total && p != null; i++)
            {
                Dictionary<string, object> item = new Dictionary<string, object>();
                string text = Clean((string)p.Range.Text);
                item["i"] = i;
                item["text"] = full ? text : Clip(text, 100);
                string style = null;
                try { style = (string)p.Style.NameLocal; } catch (Exception) { }
                if (style != null && style == titleStyle) item["style"] = "Title";
                else
                {
                    if (style != null) item["style"] = style;
                    try { int level = (int)p.OutlineLevel; if (level < 10) item["level"] = level; } catch (Exception) { }
                }
                bool inTable = false;
                try { inTable = Truthy(p.Range.Information[12]); } catch (Exception) { }
                if (inTable)
                {
                    // A table is one entry, not one per cell: its cells are read with "table".
                    dynamic table = p.Range.Tables[1];
                    int end = (int)table.Range.End, inside = (int)table.Range.Paragraphs.Count;
                    item.Clear();
                    item["i"] = i;
                    item["to"] = i + inside - 1;
                    item["table"] = (int)doc.Range(0, end).Tables.Count;
                    item["size"] = (int)table.Rows.Count + "x" + (int)table.Columns.Count;
                    items.Add(item);
                    i += inside - 1;
                    p = i < total ? doc.Range(end, end).Paragraphs[1] : null;
                    shown = i;
                    continue;
                }
                items.Add(item);
                shown = i;
                p = p.Next();
            }
        }
        result["from"] = from;
        result["items"] = items;
        if (shown < total) result["more"] = "Paragraphs " + (shown + 1) + "–" + total + " not shown; read them with \"from\".";
        return result;
    }

    static dynamic WordRange(dynamic doc, Bag op)
    {
        dynamic scope;
        if (op.Has("para"))
        {
            dynamic first = Para(doc, op.Int("para", 1), op.Str("expect", null));
            dynamic last = op.Has("to") ? Para(doc, op.Int("to", 1), null) : first;
            scope = doc.Range(first.Range.Start, last.Range.End);
        }
        else scope = doc.Content;
        if (!op.Has("find")) return scope;
        string find = op.Need("find");
        dynamic match = scope.Duplicate;
        if (!(bool)match.Find.Execute(FindText: find, MatchCase: true, Forward: true, Wrap: 0)) throw new Fail("NOT_FOUND", "The text \"" + Clip(find, 60) + "\" was not found there.");
        return match;
    }

    static void WordFormat(dynamic range, Bag op)
    {
        if (op.Has("style")) SetStyle(range, op.Need("style"));
        if (op.Has("font")) { string font = op.Need("font"); range.Font.Name = font; try { range.Font.NameFarEast = font; } catch (COMException) { } }
        if (op.Has("size")) range.Font.Size = (float)op.Num("size", 12);
        if (op.Has("bold")) range.Font.Bold = op.Flag("bold", false) ? 1 : 0;
        if (op.Has("italic")) range.Font.Italic = op.Flag("italic", false) ? 1 : 0;
        if (op.Has("underline")) range.Font.Underline = op.Flag("underline", false) ? 1 : 0;
        if (op.Has("color")) range.Font.Color = Bgr(op.Need("color"));
        if (op.Has("align"))
        {
            string align = op.Need("align");
            range.ParagraphFormat.Alignment = align == "center" ? 1 : align == "right" ? 2 : align == "justify" ? 3 : 0;
        }
        if (op.Has("firstLineIndent")) range.ParagraphFormat.FirstLineIndent = (float)op.Num("firstLineIndent", 0);
        if (op.Has("spaceBefore")) range.ParagraphFormat.SpaceBefore = (float)op.Num("spaceBefore", 0);
        if (op.Has("spaceAfter")) range.ParagraphFormat.SpaceAfter = (float)op.Num("spaceAfter", 0);
        if (op.Has("lineSpacing")) { range.ParagraphFormat.LineSpacingRule = 5; range.ParagraphFormat.LineSpacing = (float)(op.Num("lineSpacing", 1) * 12); }
    }

    /// A new empty paragraph where an insert operation points: after / before paragraph "para", or at the start / end.
    static dynamic NewParagraph(dynamic doc, Bag op)
    {
        string where = op.Str("where", op.Has("para") ? "after" : "end");
        // A new, empty document: its one empty paragraph is where the text goes, wherever the operation points.
        if ((int)doc.Paragraphs.Count == 1 && (string)doc.Paragraphs[1].Range.Text == "\r") return doc.Paragraphs[1];
        if (where == "end")
        {
            dynamic last = doc.Paragraphs.Last;
            if ((string)last.Range.Text == "\r" && !Truthy(last.Range.Information[12])) return last;
            return After(doc, last);
        }
        dynamic anchor = where == "start" ? doc.Paragraphs[1] : Para(doc, op.Int("para", 1), op.Str("expect", null));
        if (where == "after") return After(doc, anchor);
        if (where != "before" && where != "start") throw new Fail("BAD_ARGS", "\"where\" must be after, before, start or end.");
        int start = (int)anchor.Range.Start;
        anchor.Range.InsertParagraphBefore();
        return doc.Range(start, start).Paragraphs[1];
    }

    /// A new empty paragraph right after the given one.
    static dynamic After(dynamic doc, dynamic paragraph)
    {
        int end = (int)paragraph.Range.End;
        paragraph.Range.InsertParagraphAfter();
        return doc.Range(end, end).Paragraphs[1];
    }

    /// The number of a paragraph, counted from the start of the document.
    static int Index(dynamic doc, dynamic paragraph)
    {
        return (int)doc.Range(0, paragraph.Range.End).Paragraphs.Count;
    }

    static void Plain(dynamic paragraph, string style)
    {
        SetStyle(paragraph.Range, style);
        try { paragraph.Range.Font.Reset(); paragraph.Range.ParagraphFormat.Reset(); } catch (COMException) { }
    }

    static string WordOp(dynamic doc, string type, Bag op)
    {
        if (type == "replace_text")
        {
            string find = op.Need("find"), replace = Lines(op.Raw("replace") ?? "");
            bool all = op.Flag("all", true), matchCase = op.Flag("matchCase", true);
            dynamic range = doc.Content;
            int count = 0;
            while (count < 5000 && (bool)range.Find.Execute(FindText: find, MatchCase: matchCase, Forward: true, Wrap: 0))
            {
                range.Text = replace;
                range.Collapse(0);
                count++;
                if (!all) break;
            }
            if (count == 0) throw new Fail("NOT_FOUND", "The text \"" + Clip(find, 60) + "\" does not occur in the document.");
            return "replaced " + count;
        }
        if (type == "insert_paragraphs")
        {
            IList items = op.List("items");
            if (items == null) { items = new ArrayList(); items.Add(new Dictionary<string, object> { { "text", op.Need("text") } }); }
            dynamic current = null;
            int made = 0;
            foreach (object raw in items)
            {
                Bag item = raw is string ? new Bag(new Dictionary<string, object> { { "text", raw } }) : new Bag(raw);
                foreach (string text in Lines(item.Raw("text") ?? "").Split('\r'))
                {
                    dynamic p;
                    if (current == null) p = NewParagraph(doc, op);
                    else p = After(doc, current);
                    Plain(p, item.Str("style", op.Str("style", "Normal")));
                    if (text.Length > 0) p.Range.InsertBefore(text);
                    WordFormat(p.Range, StyleLess(op));
                    WordFormat(p.Range, StyleLess(item));
                    current = p;
                    made++;
                }
            }
            int lastIndex = Index(doc, current);
            return "inserted " + made + " paragraph(s), now paragraphs " + (lastIndex - made + 1) + "–" + lastIndex + " of " + (int)doc.Paragraphs.Count;
        }
        if ((type == "set_text" || (type == "delete_range" && op.Has("para"))) && op.Raw("expect") == null)
        {
            throw new Fail("BAD_ARGS", type + " needs \"expect\": the first words of paragraph " + op.Int("para", 0) + " as you last read it (\"\" for an empty paragraph), so that the wrong paragraph is never changed.");
        }
        if (type == "set_text")
        {
            dynamic p = Para(doc, op.Int("para", 0), op.Str("expect", null));
            dynamic range = p.Range;
            range.MoveEnd(1, -1);
            range.Text = Lines(op.Raw("text") ?? "");
            return "paragraph " + op.Int("para", 0) + " rewritten";
        }
        if (type == "format_text" || type == "set_style")
        {
            if (type == "set_style") op.Need("style");
            WordFormat(WordRange(doc, op), op);
            return "formatted";
        }
        if (type == "delete_range")
        {
            if (!op.Has("para") && !op.Has("find")) throw new Fail("BAD_ARGS", "Say what to delete with \"para\" (and \"to\") or \"find\".");
            WordRange(doc, op).Delete();
            return "deleted, " + (int)doc.Paragraphs.Count + " paragraphs now";
        }
        if (type == "insert_table")
        {
            IList data = op.List("data");
            int rows = op.Int("rows", data == null ? 2 : data.Count), cols = op.Int("cols", 0);
            if (data != null) foreach (object row in data) { IList cells = row as IList; if (cells != null) cols = Math.Max(cols, cells.Count); }
            if (rows < 1 || cols < 1) throw new Fail("BAD_ARGS", "Give \"data\" (rows of cells) or \"rows\" and \"cols\".");
            dynamic p = NewParagraph(doc, op);
            Plain(p, "Normal");
            dynamic table = doc.Tables.Add(p.Range, rows, cols);
            // The cells take the formatting of the paragraph that follows (a heading, say): make them plain.
            SetStyle(table.Range, "Normal");
            try
            {
                table.Range.Font.Reset(); table.Range.ParagraphFormat.Reset();
                table.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0; table.Range.ParagraphFormat.FirstLineIndent = 0;
            }
            catch (COMException) { }
            table.Borders.Enable = 1;
            if (data != null)
            {
                for (int r = 0; r < data.Count && r < rows; r++)
                {
                    IList cells = data[r] as IList;
                    if (cells == null) continue;
                    for (int c = 0; c < cells.Count; c++) if (cells[c] != null) table.Cell(r + 1, c + 1).Range.Text = Convert.ToString(cells[c]);
                }
            }
            if (op.Flag("header", true)) table.Rows[1].Range.Font.Bold = 1;
            int tableEnd = (int)doc.Range(0, table.Range.End).Paragraphs.Count;
            return "table " + (int)doc.Range(0, table.Range.End).Tables.Count + " inserted (" + rows + "×" + cols + "), now paragraphs " + (tableEnd - (int)table.Range.Paragraphs.Count + 1) + "–" + tableEnd + " of " + (int)doc.Paragraphs.Count;
        }
        if (type == "set_cell")
        {
            int n = op.Int("table", 1);
            if (n < 1 || n > (int)doc.Tables.Count) throw new Fail("ANCHOR_MISSING", "There is no table " + n + ".");
            doc.Tables[n].Cell(op.Int("row", 1), op.Int("col", 1)).Range.Text = Lines(op.Raw("text") ?? "");
            return "cell set";
        }
        if (type == "insert_image")
        {
            string path = op.Need("path");
            if (!File.Exists(path)) throw new Fail("BAD_ARGS", "Image \"" + path + "\" does not exist.");
            dynamic p = NewParagraph(doc, op);
            Plain(p, "Normal");
            dynamic picture = p.Range.InlineShapes.AddPicture(FileName: path, LinkToFile: false, SaveWithDocument: true);
            if (op.Has("width")) { picture.LockAspectRatio = -1; picture.Width = (float)op.Num("width", 300); }
            WordFormat(p.Range, StyleLess(op));
            return "image inserted";
        }
        throw new Fail("BAD_ARGS", "Unknown Word operation \"" + type + "\".");
    }

    static Bag StyleLess(Bag source)
    {
        Dictionary<string, object> copy = new Dictionary<string, object>();
        foreach (string key in new string[] { "font", "size", "bold", "italic", "underline", "color", "align", "firstLineIndent", "spaceBefore", "spaceAfter", "lineSpacing" })
            if (source.Has(key)) copy[key] = source.Raw(key);
        return new Bag(copy);
    }

    // ───────────────────────── Excel ─────────────────────────

    static dynamic Sheet(dynamic book, Bag op)
    {
        string name = op.Str("sheet", null);
        try
        {
            if (name == null) return book.ActiveSheet;
            int index;
            return int.TryParse(name, out index) ? book.Worksheets[index] : book.Worksheets[name];
        }
        catch (COMException)
        {
            List<string> names = new List<string>();
            foreach (dynamic sheet in book.Worksheets) names.Add((string)sheet.Name);
            throw new Fail("ANCHOR_MISSING", "There is no sheet \"" + name + "\". Sheets: " + string.Join(", ", names.ToArray()) + ".");
        }
    }

    static dynamic Cells(dynamic sheet, string address)
    {
        try { return sheet.Range[address]; }
        catch (COMException) { throw new Fail("BAD_ARGS", "\"" + address + "\" is not a valid range (use A1 style, e.g. B2:D10)."); }
    }

    static string Column(int column)
    {
        string name = "";
        while (column > 0) { int rem = (column - 1) % 26; name = (char)('A' + rem) + name; column = (column - 1) / 26; }
        return name;
    }

    static object CellValue(object value)
    {
        if (value is int)
        {
            switch ((int)value)
            {
                case -2146826281: return "#DIV/0!";
                case -2146826246: return "#N/A";
                case -2146826259: return "#NAME?";
                case -2146826288: return "#NULL!";
                case -2146826252: return "#NUM!";
                case -2146826265: return "#REF!";
                case -2146826273: return "#VALUE!";
            }
        }
        return value;
    }

    static object ExcelRead(dynamic book, Bag a)
    {
        Dictionary<string, object> result = new Dictionary<string, object>();
        List<object> sheets = new List<object>();
        foreach (dynamic ws in book.Worksheets)
        {
            Dictionary<string, object> entry = new Dictionary<string, object>();
            entry["name"] = (string)ws.Name;
            dynamic used = ws.UsedRange;
            entry["used"] = (string)used.Address[false, false];
            int charts = (int)ws.ChartObjects().Count;
            if (charts > 0) entry["charts"] = charts;
            sheets.Add(entry);
        }
        result["sheets"] = sheets;
        dynamic sheet = Sheet(book, a);
        result["sheet"] = (string)sheet.Name;
        dynamic range = a.Has("range") ? Cells(sheet, a.Need("range")) : sheet.UsedRange;
        int rows = (int)range.Rows.Count, cols = (int)range.Columns.Count;
        int maxRows = a.Has("range") ? 2000 : 60, maxCols = a.Has("range") ? 100 : 26;
        if (rows > maxRows || cols > maxCols)
        {
            result["clipped"] = "Showing the first " + Math.Min(rows, maxRows) + " rows × " + Math.Min(cols, maxCols) + " columns of " + rows + " × " + cols + "; pass \"range\" for more.";
            rows = Math.Min(rows, maxRows); cols = Math.Min(cols, maxCols);
            range = range.Resize[rows, cols];
        }
        result["range"] = (string)range.Address[false, false];
        object values = range.Value2, formulas = range.Formula;
        int row0 = (int)range.Row, col0 = (int)range.Column;
        List<object> grid = new List<object>();
        List<object> formulaList = new List<object>();
        object[,] v = values as object[,], f = formulas as object[,];
        for (int r = 1; r <= rows; r++)
        {
            List<object> line = new List<object>();
            for (int c = 1; c <= cols; c++)
            {
                line.Add(CellValue(v == null ? values : v[r, c]));
                string formula = Convert.ToString(f == null ? formulas : f[r, c]);
                if (formula.StartsWith("=", StringComparison.Ordinal))
                {
                    formulaList.Add(new Dictionary<string, object> { { "cell", Column(col0 + c - 1) + (row0 + r - 1) }, { "formula", formula } });
                }
            }
            grid.Add(line);
        }
        result["values"] = grid;
        if (formulaList.Count > 0) result["formulas"] = formulaList;
        return result;
    }

    static object Scalar(object value)
    {
        if (value is decimal) return Convert.ToDouble(value);
        return value;
    }

    static void ExcelFormat(dynamic range, Bag op)
    {
        if (op.Has("font")) range.Font.Name = op.Need("font");
        if (op.Has("size")) range.Font.Size = op.Num("size", 11);
        if (op.Has("bold")) range.Font.Bold = op.Flag("bold", false);
        if (op.Has("italic")) range.Font.Italic = op.Flag("italic", false);
        if (op.Has("color")) range.Font.Color = Bgr(op.Need("color"));
        if (op.Has("fill")) range.Interior.Color = Bgr(op.Need("fill"));
        if (op.Has("numberFormat")) range.NumberFormat = op.Need("numberFormat");
        if (op.Has("align"))
        {
            string align = op.Need("align");
            range.HorizontalAlignment = align == "center" ? -4108 : align == "right" ? -4152 : -4131;
        }
        if (op.Has("wrap")) range.WrapText = op.Flag("wrap", false);
        if (op.Flag("border", false)) range.Borders.LineStyle = 1;
        if (op.Flag("merge", false)) range.Merge();
        if (op.Has("columnWidth")) range.ColumnWidth = op.Num("columnWidth", 10);
        if (op.Has("rowHeight")) range.RowHeight = op.Num("rowHeight", 15);
    }

    static string ExcelOp(dynamic app, dynamic book, string type, Bag op)
    {
        if (type == "add_sheet")
        {
            dynamic sheets = book.Worksheets;
            dynamic sheet = sheets.Add(After: sheets[(int)sheets.Count]);
            if (op.Has("name")) sheet.Name = op.Need("name");
            return "sheet \"" + (string)sheet.Name + "\" added";
        }
        dynamic ws = Sheet(book, op);
        if (type == "write_range")
        {
            IList rows = op.List("values");
            if (rows == null || rows.Count == 0) throw new Fail("BAD_ARGS", "\"values\" must be rows of cells, e.g. [[\"Name\",\"Qty\"],[\"A\",3]].");
            int width = 1;
            foreach (object row in rows) { IList cells = row as IList; if (cells != null && !(row is string)) width = Math.Max(width, cells.Count); }
            object[,] grid = new object[rows.Count, width];
            for (int r = 0; r < rows.Count; r++)
            {
                IList cells = rows[r] as IList;
                if (cells == null || rows[r] is string) { grid[r, 0] = Scalar(rows[r]); continue; }
                for (int c = 0; c < cells.Count; c++) grid[r, c] = Scalar(cells[c]);
            }
            dynamic target = Cells(ws, op.Str("range", "A1")).Cells[1, 1].Resize[rows.Count, width];
            target.Formula = grid;
            ExcelFormat(target, op);
            return "wrote " + (string)target.Address[false, false];
        }
        if (type == "format_range") { ExcelFormat(Cells(ws, op.Need("range")), op); return "formatted"; }
        if (type == "autofit")
        {
            dynamic range = op.Has("range") ? Cells(ws, op.Need("range")) : ws.UsedRange;
            range.EntireColumn.AutoFit();
            return "columns fitted";
        }
        if (type == "rename_sheet") { ws.Name = op.Need("name"); return "renamed"; }
        if (type == "delete_sheet")
        {
            bool alerts = (bool)app.DisplayAlerts;
            app.DisplayAlerts = false;
            try { ws.Delete(); } finally { app.DisplayAlerts = alerts; }
            return "sheet deleted";
        }
        if (type == "insert_rows" || type == "delete_rows")
        {
            int row = op.Int("row", 0), count = Math.Max(1, op.Int("count", 1));
            if (row < 1) throw new Fail("BAD_ARGS", "\"row\" is required (1-based).");
            dynamic rows = ws.Range[row + ":" + (row + count - 1)];
            if (type == "insert_rows") rows.Insert(); else rows.Delete();
            return (type == "insert_rows" ? "inserted " : "deleted ") + count + " row(s)";
        }
        if (type == "add_chart")
        {
            dynamic source = Cells(ws, op.Need("range"));
            dynamic anchor = Cells(ws, op.Str("at", "H2"));
            dynamic holder = ws.ChartObjects().Add((double)anchor.Left, (double)anchor.Top, op.Num("width", 420), op.Num("height", 260));
            dynamic chart = holder.Chart;
            chart.SetSourceData(source);
            string kind = op.Str("chart", "column");
            chart.ChartType = kind == "bar" ? 57 : kind == "line" ? 65 : kind == "pie" ? 5 : kind == "scatter" ? -4169 : kind == "area" ? 1 : 51;
            if (op.Has("title")) { chart.HasTitle = true; chart.ChartTitle.Text = op.Need("title"); }
            if (op.Has("name")) holder.Name = op.Need("name");
            return "chart \"" + (string)holder.Name + "\" added";
        }
        throw new Fail("BAD_ARGS", "Unknown Excel operation \"" + type + "\".");
    }

    // ───────────────────────── PowerPoint ─────────────────────────

    static dynamic Slide(dynamic deck, Bag op)
    {
        int index = op.Int("slide", 0), total = (int)deck.Slides.Count;
        if (index < 1 || index > total) throw new Fail("ANCHOR_MISSING", "There is no slide " + index + " (the deck has " + total + ").");
        return deck.Slides[index];
    }

    static dynamic Shape(dynamic slide, string key)
    {
        try
        {
            if (key == "title") return slide.Shapes.Title;
            if (key == "body") return slide.Shapes.Placeholders[2];
            int index;
            return int.TryParse(key, out index) ? slide.Shapes[index] : slide.Shapes[key];
        }
        catch (COMException)
        {
            List<string> names = new List<string>();
            foreach (dynamic shape in slide.Shapes) names.Add((string)shape.Name);
            throw new Fail("ANCHOR_MISSING", "Slide " + (int)slide.SlideIndex + " has no shape \"" + key + "\". Shapes: " + (names.Count == 0 ? "(none)" : string.Join(", ", names.ToArray())) + ".");
        }
    }

    static string ShapeType(int type)
    {
        switch (type)
        {
            case 1: return "shape";
            case 3: return "chart";
            case 6: return "group";
            case 9: return "line";
            case 13: return "picture";
            case 14: return "placeholder";
            case 17: return "textbox";
            case 19: return "table";
            default: return "type" + type;
        }
    }

    static object PptRead(dynamic deck, Bag a)
    {
        Dictionary<string, object> result = new Dictionary<string, object>();
        result["slides"] = (int)deck.Slides.Count;
        result["width"] = Math.Round((double)deck.PageSetup.SlideWidth, 1);
        result["height"] = Math.Round((double)deck.PageSetup.SlideHeight, 1);
        bool full = a.Flag("full", false);
        int only = a.Int("slide", 0);
        List<object> slides = new List<object>();
        foreach (dynamic slide in deck.Slides)
        {
            int index = (int)slide.SlideIndex;
            if (only > 0 && index != only) continue;
            Dictionary<string, object> entry = new Dictionary<string, object>();
            entry["slide"] = index;
            try { entry["layout"] = (string)slide.CustomLayout.Name; } catch (Exception) { }
            List<object> shapes = new List<object>();
            foreach (dynamic shape in slide.Shapes)
            {
                Dictionary<string, object> item = new Dictionary<string, object>();
                item["name"] = (string)shape.Name;
                item["type"] = ShapeType((int)shape.Type);
                item["box"] = new double[] { Math.Round((double)shape.Left), Math.Round((double)shape.Top), Math.Round((double)shape.Width), Math.Round((double)shape.Height) };
                try
                {
                    if (Truthy(shape.HasTextFrame) && Truthy(shape.TextFrame.HasText))
                    {
                        string text = ((string)shape.TextFrame.TextRange.Text).Replace('\r', '\n').Replace('\v', '\n');
                        item["text"] = full ? text : Clip(text, 200);
                    }
                }
                catch (Exception) { }
                shapes.Add(item);
            }
            entry["shapes"] = shapes;
            slides.Add(entry);
        }
        result["items"] = slides;
        return result;
    }

    static void PptText(dynamic shape, Bag op)
    {
        if (op.Has("text")) shape.TextFrame.TextRange.Text = Lines(op.Raw("text"));
        if (!(op.Has("font") || op.Has("size") || op.Has("bold") || op.Has("italic") || op.Has("color") || op.Has("align"))) return;
        dynamic range = shape.TextFrame.TextRange;
        if (op.Has("font")) { string font = op.Need("font"); range.Font.Name = font; try { range.Font.NameFarEast = font; } catch (COMException) { } }
        if (op.Has("size")) range.Font.Size = (float)op.Num("size", 18);
        if (op.Has("bold")) range.Font.Bold = op.Flag("bold", false) ? -1 : 0;
        if (op.Has("italic")) range.Font.Italic = op.Flag("italic", false) ? -1 : 0;
        if (op.Has("color")) range.Font.Color.RGB = Bgr(op.Need("color"));
        if (op.Has("align"))
        {
            string align = op.Need("align");
            range.ParagraphFormat.Alignment = align == "center" ? 2 : align == "right" ? 3 : 1;
        }
    }

    static void PptBox(dynamic shape, Bag op)
    {
        if (op.Has("left")) shape.Left = (float)op.Num("left", 0);
        if (op.Has("top")) shape.Top = (float)op.Num("top", 0);
        if (op.Has("width")) shape.Width = (float)op.Num("width", 100);
        if (op.Has("height")) shape.Height = (float)op.Num("height", 100);
        if (op.Has("fill")) { shape.Fill.Visible = -1; shape.Fill.Solid(); shape.Fill.ForeColor.RGB = Bgr(op.Need("fill")); }
        if (op.Has("name")) shape.Name = op.Need("name");
    }

    static readonly Dictionary<string, int> Layouts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        { "title", 1 }, { "title_content", 2 }, { "two_content", 29 }, { "title_only", 11 }, { "blank", 12 }, { "section", 33 }, { "comparison", 34 },
    };

    static string PptOp(dynamic deck, string type, Bag op)
    {
        if (type == "add_slide")
        {
            int total = (int)deck.Slides.Count, index = op.Int("at", total + 1), layout;
            if (index < 1 || index > total + 1) throw new Fail("BAD_ARGS", "\"at\" must be between 1 and " + (total + 1) + ".");
            string name = op.Str("layout", "title_content");
            if (!Layouts.TryGetValue(name, out layout)) throw new Fail("BAD_ARGS", "Unknown layout \"" + name + "\": use " + string.Join(", ", new List<string>(Layouts.Keys).ToArray()) + ".");
            dynamic slide = deck.Slides.Add(index, layout);
            if (op.Has("title")) slide.Shapes.Title.TextFrame.TextRange.Text = Lines(op.Raw("title"));
            if (op.Has("body")) slide.Shapes.Placeholders[2].TextFrame.TextRange.Text = Lines(op.Raw("body"));
            if (op.Has("notes")) slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = Lines(op.Raw("notes"));
            return "slide " + index + " added";
        }
        dynamic target = Slide(deck, op);
        if (type == "delete_slide") { target.Delete(); return "slide deleted"; }
        if (type == "move_slide") { target.MoveTo(op.Int("to", 1)); return "slide moved to " + op.Int("to", 1); }
        if (type == "set_notes") { target.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = Lines(op.Raw("text") ?? ""); return "notes set"; }
        if (type == "add_textbox")
        {
            dynamic box = target.Shapes.AddTextbox(1, (float)op.Num("left", 60), (float)op.Num("top", 60), (float)op.Num("width", 400), (float)op.Num("height", 60));
            PptText(box, op);
            if (op.Has("fill") || op.Has("name")) PptBox(box, new Bag(new Dictionary<string, object> { { "fill", op.Raw("fill") }, { "name", op.Raw("name") } }));
            return "textbox \"" + (string)box.Name + "\" added";
        }
        if (type == "add_image")
        {
            string path = op.Need("path");
            if (!File.Exists(path)) throw new Fail("BAD_ARGS", "Image \"" + path + "\" does not exist.");
            dynamic picture = target.Shapes.AddPicture(path, 0, -1, (float)op.Num("left", 60), (float)op.Num("top", 60), -1, -1);
            if (op.Has("width") || op.Has("height"))
            {
                picture.LockAspectRatio = op.Has("width") && op.Has("height") ? 0 : -1;
                if (op.Has("width")) picture.Width = (float)op.Num("width", 100);
                if (op.Has("height")) picture.Height = (float)op.Num("height", 100);
            }
            if (op.Has("name")) picture.Name = op.Need("name");
            return "picture \"" + (string)picture.Name + "\" added";
        }
        dynamic shape = Shape(target, op.Need("shape"));
        if (type == "set_text") { op.Need("text"); PptText(shape, op); return "text set"; }
        if (type == "set_shape") { PptBox(shape, op); PptText(shape, op); return "shape updated"; }
        if (type == "delete_shape") { shape.Delete(); return "shape deleted"; }
        throw new Fail("BAD_ARGS", "Unknown PowerPoint operation \"" + type + "\".");
    }

    // ───────────────────────── edit batches ─────────────────────────

    static object Edit(string kind, dynamic app, dynamic doc, Bag a)
    {
        IList ops = a.List("ops");
        if (ops == null || ops.Count == 0) throw new Fail("BAD_ARGS", "\"ops\" must list at least one operation.");
        List<object> done = new List<object>();
        Dictionary<string, object> result = new Dictionary<string, object>();
        dynamic undo = null;
        // One undo step for the whole batch: the user takes it back with a single Ctrl+Z.
        if (kind == "word") { try { undo = app.UndoRecord; undo.StartCustomRecord("AI edit"); } catch (COMException) { undo = null; } }
        try
        {
            for (int i = 0; i < ops.Count; i++)
            {
                string type = "?";
                try
                {
                    Bag op = new Bag(ops[i]);
                    type = op.Need("op");
                    string line = kind == "word" ? WordOp(doc, type, op) : kind == "excel" ? ExcelOp(app, doc, type, op) : PptOp(doc, type, op);
                    string ignored = op.Unread();
                    done.Add(ignored.Length == 0 ? line : line + " — IGNORED field(s) " + ignored + ": " + type + " does not take them, check the operation list");
                }
                catch (Exception error)
                {
                    string code;
                    Dictionary<string, object> failed = new Dictionary<string, object>();
                    failed["index"] = i;
                    failed["op"] = type;
                    failed["error"] = Describe(error, out code);
                    failed["code"] = code;
                    result["failed"] = failed;
                    break;
                }
            }
        }
        finally
        {
            if (undo != null) { try { undo.EndCustomRecord(); } catch (COMException) { } }
        }
        result["done"] = done;
        result["total"] = ops.Count;
        return result;
    }

    // ───────────────────────── rendering ─────────────────────────

    static object Render(string kind, dynamic app, dynamic doc, Bag a)
    {
        string path = a.Need("out");
        int width = Math.Max(320, Math.Min(2400, a.Int("width", 1100)));
        Dictionary<string, object> result = new Dictionary<string, object>();
        if (kind == "ppt")
        {
            dynamic slide = Slide(doc, a);
            double ratio = (double)doc.PageSetup.SlideHeight / (double)doc.PageSetup.SlideWidth;
            slide.Export(path, "PNG", width, (int)Math.Round(width * ratio));
            result["what"] = "slide " + a.Int("slide", 0);
        }
        else if (kind == "word")
        {
            int page = Math.Max(1, a.Int("page", 1));
            dynamic pages = doc.Windows[1].Panes[1].Pages;
            int total = (int)pages.Count;
            if (page > total) throw new Fail("ANCHOR_MISSING", "There is no page " + page + " (the document has " + total + ").");
            byte[] bits = (byte[])pages[page].EnhMetaFileBits;
            using (MemoryStream stream = new MemoryStream(bits))
            using (Metafile meta = new Metafile(stream))
            {
                int height = (int)Math.Round(width * (double)meta.Height / meta.Width);
                using (Bitmap bitmap = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.Clear(Color.White);
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(meta, 0, 0, width, height);
                    }
                    bitmap.Save(path, ImageFormat.Png);
                }
            }
            result["what"] = "page " + page + " of " + total;
        }
        else
        {
            dynamic sheet = Sheet(doc, a);
            dynamic range = a.Has("range") ? Cells(sheet, a.Need("range")) : sheet.UsedRange;
            // Excel can only hand a picture of cells over through the clipboard: put the user's text back afterwards.
            string text = null;
            try { if (Clipboard.ContainsText()) text = Clipboard.GetText(); } catch (Exception) { }
            Image image = null;
            try
            {
                range.CopyPicture(1, 2);
                for (int i = 0; i < 10 && image == null; i++) { image = Clipboard.GetImage(); if (image == null) Thread.Sleep(100); }
            }
            finally
            {
                try { if (text != null) Clipboard.SetText(text); else Clipboard.Clear(); } catch (Exception) { }
            }
            if (image == null) throw new Fail("OFFICE_ERROR", "Excel did not produce a picture of that range.");
            using (image) image.Save(path, ImageFormat.Png);
            result["what"] = (string)sheet.Name + "!" + (string)range.Address[false, false];
        }
        result["path"] = path;
        using (Image saved = Image.FromFile(path)) { result["width"] = saved.Width; result["height"] = saved.Height; }
        return result;
    }
}
