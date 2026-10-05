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
    /// The fields as given, all counted as looked at (they are handed on to another operation).
    public Dictionary<string, object> Copy()
    {
        foreach (string key in d.Keys) read.Add(key);
        return new Dictionary<string, object>(d);
    }
    /// A length in points: a plain number is points, "12cm", "120mm", "5in" and "12pt" / "12磅" say their unit.
    public double Points(string key, double fallback)
    {
        object v = Raw(key);
        if (v == null) return fallback;
        string text = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture).Trim().ToLowerInvariant();
        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(text, @"^(-?\d+(\.\d+)?)\s*(pt|磅|cm|厘米|mm|毫米|in|英寸)?\z");
        if (!m.Success) throw new Fail("BAD_ARGS", "\"" + key + "\" must be a length: a number of points, or \"12cm\".");
        double value = double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
        string unit = m.Groups[3].Value;
        return unit == "cm" || unit == "厘米" ? value * 28.3465 : unit == "mm" || unit == "毫米" ? value * 2.83465 : unit == "in" || unit == "英寸" ? value * 72 : value;
    }
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
        catch
        {
            // "6pt" where points are meant.
            System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(Convert.ToString(v).Trim(), @"^(-?\d+(\.\d+)?)\s*(pt|磅)\z");
            if (m.Success) return double.Parse(m.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture);
            throw new Fail("BAD_ARGS", "\"" + key + "\" must be a number.");
        }
    }
    public bool Flag(string key, bool fallback)
    {
        object v = Raw(key);
        return v is bool ? (bool)v : fallback;
    }
    /// A switch given loosely: true, "thin", "all", 1 are on; false, "none", "off", 0 and absent are off.
    public bool On(string key)
    {
        object v = Raw(key);
        if (v == null) return false;
        if (v is bool) return (bool)v;
        string text = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture).Trim().ToLowerInvariant();
        return text.Length > 0 && text != "false" && text != "none" && text != "off" && text != "no" && text != "0";
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


/// A small always-on-top card shown while the agent edits: what it is doing, and two switches the user
/// can flip at any moment (follow the edits / write like typing). It never takes the focus.
class Card : Form
{
    static Card instance;
    static readonly ManualResetEvent Ready = new ManualResetEvent(false);
    /// Set once the user has flipped a switch here: from then on the card, not the settings page, decides.
    public static volatile bool FollowChosen, TypingChosen;
    /// Set for as long as an edit is running: the card stays up through it, however long one operation takes, and goes when it ends.
    public static volatile bool Hold;

    string line = "";
    DateTime until = DateTime.MinValue;
    readonly float scale;
    Rectangle followBox, typingBox;
    Point grip;
    bool pressed, dragged;
    readonly System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();

    [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr handle, int command);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr handle);
    [DllImport("user32.dll")] static extern bool IsIconic(IntPtr handle);

    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr handle, IntPtr none);
    [DllImport("user32.dll")] static extern bool AttachThreadInput(uint from, uint to, bool attach);
    [DllImport("user32.dll")] static extern bool BringWindowToTop(IntPtr handle);
    [DllImport("kernel32.dll")] static extern uint GetCurrentThreadId();

    /// Bring a window to the front. Windows lets a program do that only in some situations; sharing the input
    /// queue of the window now in front for the moment of the call is the way that always holds.
    static void Front(IntPtr window)
    {
        IntPtr now = GetForegroundWindow();
        if (now == window) return;
        uint mine = GetCurrentThreadId(), theirs = now != IntPtr.Zero ? GetWindowThreadProcessId(now, IntPtr.Zero) : 0;
        bool joined = theirs != 0 && theirs != mine && AttachThreadInput(mine, theirs, true);
        try
        {
            BringWindowToTop(window);
            SetForegroundWindow(window);
        }
        finally { if (joined) AttachThreadInput(mine, theirs, false); }
    }

    /// The user is looking at the document: its window is the foreground one.
    static bool Watching()
    {
        IntPtr doc = Program.DocWindow;
        return doc != IntPtr.Zero && GetForegroundWindow() == doc;
    }

    // Following means "keep the document in front of me and scroll to the edits". It holds while the user
    // stays in that window; once they go elsewhere it switches itself off, and the button brings them back.
    bool wasWatching, lit;

    protected override bool ShowWithoutActivation { get { return true; } }

    protected override CreateParams CreateParams
    {
        get
        {
            CreateParams p = base.CreateParams;
            p.ExStyle |= 0x08000000 | 0x00000080 | 0x00000008; // no-activate, tool window, topmost
            return p;
        }
    }

    Card()
    {
        FormBorderStyle = FormBorderStyle.None;
        ShowInTaskbar = false;
        StartPosition = FormStartPosition.Manual;
        TopMost = true;
        DoubleBuffered = true;
        BackColor = Color.FromArgb(34, 34, 38);
        Opacity = 0.94;
        using (Graphics g = CreateGraphics()) scale = g.DpiX / 96f;
        Size = new Size(S(330), S(38));
        Rectangle area = Screen.PrimaryScreen.WorkingArea;
        Location = new Point(area.Right - Width - S(18), area.Bottom - Height - S(18));
        using (System.Drawing.Drawing2D.GraphicsPath path = Rounded(new Rectangle(0, 0, Width, Height), S(10))) Region = new Region(path);
        timer.Interval = 500;
        timer.Tick += delegate
        {
            bool watching = Watching();
            if (watching) wasWatching = true;
            else if (wasWatching && Program.Following)
            {
                // The user left the document for another window.
                wasWatching = false;
                Program.Following = false;
                FollowChosen = true;
            }
            bool nowLit = Program.Following && watching;
            if (nowLit != lit) { lit = nowLit; Invalidate(); }
            if (Visible && !Hold && DateTime.UtcNow > until) Hide();
            Program.Trace("card visible " + Visible + " hold " + Hold + " following " + Program.Following + " watching " + watching);
        };
        timer.Start();
    }

    int S(int value) { return (int)Math.Round(value * scale); }

    static System.Drawing.Drawing2D.GraphicsPath Rounded(Rectangle r, int radius)
    {
        System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
        int d = radius * 2;
        path.AddArc(r.X, r.Y, d, d, 180, 90);
        path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    void Chip(Graphics g, Font font, string label, bool on, ref int right, out Rectangle box)
    {
        Size size = TextRenderer.MeasureText(label, font);
        box = new Rectangle(right - size.Width - S(14), (Height - S(24)) / 2, size.Width + S(14), S(24));
        right = box.Left - S(6);
        Color accent = Color.FromArgb(217, 119, 87);
        using (System.Drawing.Drawing2D.GraphicsPath path = Rounded(box, S(12)))
        {
            if (on) using (SolidBrush fill = new SolidBrush(accent)) g.FillPath(fill, path);
            else using (Pen pen = new Pen(Color.FromArgb(110, 110, 118))) g.DrawPath(pen, path);
        }
        TextRenderer.DrawText(g, label, font, box, on ? Color.White : Color.FromArgb(170, 170, 178), TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        Graphics g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using (Font font = new Font("Microsoft YaHei UI", 9f))
        {
            int right = Width - S(10);
            // "Fast" is lit when the text is written at once; unlit, it is written the way a person types.
            Chip(g, font, "快速", !Program.Typing, ref right, out typingBox);
            bool following = Program.Following && Watching();
            Chip(g, font, following ? "查看中" : "查看", following, ref right, out followBox);
            using (SolidBrush dot = new SolidBrush(Color.FromArgb(217, 119, 87))) g.FillEllipse(dot, S(12), (Height - S(8)) / 2, S(8), S(8));
            Rectangle text = new Rectangle(S(26), 0, right - S(30), Height);
            TextRenderer.DrawText(g, line, font, text, Color.FromArgb(235, 235, 240), TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }
    }

    protected override void OnMouseDown(MouseEventArgs e) { pressed = true; dragged = false; grip = e.Location; }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        if (!pressed) return;
        if (!dragged && Math.Abs(e.X - grip.X) + Math.Abs(e.Y - grip.Y) < S(4)) return;
        dragged = true;
        Location = new Point(Location.X + e.X - grip.X, Location.Y + e.Y - grip.Y);
    }

    protected override void OnMouseUp(MouseEventArgs e)
    {
        pressed = false;
        if (dragged) return;
        if (followBox.Contains(e.Location)) ToggleFollow();
        else if (typingBox.Contains(e.Location)) { Program.Typing = !Program.Typing; TypingChosen = true; }
        if (until < DateTime.UtcNow.AddSeconds(4)) until = DateTime.UtcNow.AddSeconds(4);
        Invalidate();
    }

    void ToggleFollow()
    {
        {
            FollowChosen = true;
            if (Program.Following && Watching()) Program.Following = false;
            else
            {
                // Bring the document to the front and follow from here on.
                Program.Following = true;
                wasWatching = true;
                IntPtr doc = Program.DocWindow;
                if (doc != IntPtr.Zero)
                {
                    if (IsIconic(doc)) ShowWindow(doc, 9);
                    Front(doc);
                }
            }
        }
    }

    /// For checks: what the card is showing, and optionally a press of the follow switch.
    public static string Probe(bool press)
    {
        Card card = instance;
        if (card == null) return "no card";
        string state = "";
        ManualResetEvent done = new ManualResetEvent(false);
        card.BeginInvoke(new MethodInvoker(delegate
        {
            if (press) { card.ToggleFollow(); card.Invalidate(); }
            state = "visible " + card.Visible + " at " + card.Bounds + " line \"" + card.line + "\" following " + Program.Following + " window " + Program.DocWindow;
            done.Set();
        }));
        done.WaitOne(3000);
        Thread.Sleep(400);
        return state + " foreground " + GetForegroundWindow() + " watching " + Watching();
    }

    public static void Start()
    {
        Thread thread = new Thread(delegate()
        {
            try { instance = new Card(); IntPtr handle = instance.Handle; }
            catch (Exception) { instance = null; }
            Ready.Set();
            if (instance != null) Application.Run();
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        Ready.WaitOne(5000);
    }

    /// Show the card (if it is not up yet) with this line; it goes away by itself a few seconds after the last report.
    public static void Report(string text) { Report(text, 1); }

    /// seconds: how long the card stays after this, when no edit is holding it up.
    public static void Report(string text, double seconds)
    {
        Card card = instance;
        if (card == null) return;
        try
        {
            card.BeginInvoke(new MethodInvoker(delegate
            {
                card.line = text;
                card.until = DateTime.UtcNow.AddSeconds(seconds);
                if (!card.Visible) ShowWindow(card.Handle, 4);
                card.Invalidate();
            }));
        }
        catch (Exception) { }
    }
}

static class Program
{
    [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

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
        try { SetProcessDPIAware(); } catch (Exception) { }
        Send(new Dictionary<string, object> { { "event", "ready" } });
        AppDomain.CurrentDomain.ProcessExit += delegate { MathQuit(); };
        Sweep();
        SweepOwn();
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

    /// The apps this helper started because they were not running (as opposed to the ones it found open).
    static readonly HashSet<string> Started = new HashSet<string>();

    // Silent mode. Word and Excel get an instance of their own that is never shown, apart from whatever the user has
    // open: nothing flashes on their screen and their windows are not touched. PowerPoint runs once per session, so a
    // deck is opened there without a window instead. What was opened this way is saved and closed when the agent stops.

    /// word / excel → the hidden instance of our own.
    static readonly Dictionary<string, object> Own = new Dictionary<string, object>();
    /// Full paths of the documents opened in the background.
    static readonly HashSet<string> Hidden = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    /// Documents brought into view on request while silent mode is on: these are worked on visibly from then on.
    static readonly HashSet<string> Revealed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

    static string OwnNotePath()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "dsh-office", "background.pid");
    }

    /// Write down the process behind a hidden window, so that a later helper can end it should this one be cut off.
    static void NoteOwn(long window)
    {
        try
        {
            uint pid;
            GetWindowThreadProcessId(new IntPtr(window), out pid);
            if (pid == 0) return;
            Directory.CreateDirectory(Path.GetDirectoryName(OwnNotePath()));
            File.AppendAllText(OwnNotePath(), pid.ToString(System.Globalization.CultureInfo.InvariantCulture) + "\n");
        }
        catch (Exception) { }
    }

    /// Hidden instances a helper before this one left behind: ended, if they still show no window.
    static void SweepOwn()
    {
        try
        {
            string note = OwnNotePath();
            if (!File.Exists(note)) return;
            foreach (string line in File.ReadAllLines(note))
            {
                int pid;
                if (!int.TryParse(line.Trim(), out pid) || pid <= 0) continue;
                try
                {
                    using (System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(pid))
                    {
                        string name = process.ProcessName.ToUpperInvariant();
                        if ((name == "WINWORD" || name == "EXCEL") && process.MainWindowHandle == IntPtr.Zero) process.Kill();
                    }
                }
                catch (Exception) { }
            }
            File.Delete(note);
        }
        catch (Exception) { }
    }

    static dynamic OwnApp(string kind, bool start)
    {
        object held;
        if (Own.TryGetValue(kind, out held))
        {
            try { string alive = Convert.ToString(((dynamic)held).Version); return held; }
            catch (Exception) { Own.Remove(kind); }
        }
        if (!start) return null;
        if (kind == "ppt")
        {
            // PowerPoint runs once per session: the user's if it is open, else one we start — which stays alive only
            // while it is held, having no window of its own.
            dynamic deckApp = App(kind, true);
            Own[kind] = deckApp;
            return deckApp;
        }
        Type type = Type.GetTypeFromProgID(ProgId(kind));
        if (type == null) throw new Fail("NOT_INSTALLED", AppName(kind) + " is not installed on this computer.");
        dynamic app = Activator.CreateInstance(type);
        try { app.Visible = false; } catch (Exception) { }
        try { if (kind == "word") app.DisplayAlerts = 0; else app.DisplayAlerts = false; } catch (Exception) { }
        if (kind == "excel") { try { NoteOwn(Convert.ToInt64(app.Hwnd)); } catch (Exception) { } }
        Own[kind] = app;
        return app;
    }

    /// The app that holds the document a command is about: our hidden one when the document is there, else the user's.
    static dynamic AppFor(string kind, Bag a)
    {
        dynamic own = OwnApp(kind, false);
        if (own != null)
        {
            string doc = a.Str("doc", null);
            try { if (doc == null ? (int)Docs(kind, own).Count > 0 : Find(kind, own, doc) != null) return own; }
            catch (Exception) { }
        }
        return App(kind, false);
    }

    /// Open (or make) a document where no one sees it.
    static object OpenQuietly(string kind, string path)
    {
        dynamic theirs = Running(kind), own = OwnApp(kind, false);
        dynamic doc = null, app = null;
        string how = "attached";
        // Open already — in the user's window or in the background: it is worked on where it is.
        if (path != null)
        {
            if (own != null) { doc = Find(kind, own, path); if (doc != null) app = own; }
            if (doc == null && theirs != null) { try { doc = Find(kind, theirs, path); if (doc != null) app = theirs; } catch (Exception) { } }
        }
        if (doc == null)
        {
            if (path != null && !File.Exists(path) && !Directory.Exists(Path.GetDirectoryName(path))) throw new Fail("BAD_ARGS", "The folder of \"" + path + "\" does not exist.");
            bool exists = path != null && File.Exists(path);
            if (kind == "ppt")
            {
                app = OwnApp(kind, true);
                // The last argument is the window: none.
                doc = exists ? app.Presentations.Open(path, 0, 0, 0) : app.Presentations.Add(0);
            }
            else
            {
                app = OwnApp(kind, true);
                WaitReady(kind, app);
                dynamic docs = Docs(kind, app);
                doc = exists ? docs.Open(path) : docs.Add();
                if (kind == "word") { try { NoteOwn(Convert.ToInt64(doc.ActiveWindow.Hwnd)); } catch (Exception) { } }
            }
            how = exists ? "opened" : "created";
            if (!exists && path != null) SaveAs(kind, app, doc, path);
            Hidden.Add((string)doc.FullName);
        }
        Dictionary<string, object> info = Info(kind, doc, null);
        info["how"] = how;
        info["background"] = Hidden.Contains((string)doc.FullName);
        DocWindow = IntPtr.Zero;
        return info;
    }

    /// Bring a document that was opened in the background into view, for work on its window.
    static void Reveal(string kind, dynamic app, dynamic doc)
    {
        string full = (string)doc.FullName;
        if (!Hidden.Contains(full)) return;
        if (kind == "ppt") { try { doc.NewWindow(); } catch (Exception) { } app.Visible = -1; }
        else
        {
            // The instance stays the one we reach this document through; it is only no longer hidden, and it is
            // ended when its last document is closed, not when the agent stops.
            app.Visible = true;
            if (kind == "excel") { try { app.UserControl = true; } catch (Exception) { } }
            try { if (kind == "word") { doc.Activate(); app.Activate(); } else doc.Activate(); } catch (Exception) { }
        }
        Hidden.Remove(full);
        Revealed.Add(full);
    }

    /// The agent has stopped: what it opened in the background is saved and closed, and what we started for it ended.
    static object Settle()
    {
        List<object> said = new List<object>();
        foreach (string kind in new string[] { "word", "excel", "ppt" })
        {
            dynamic app = OwnApp(kind, false);
            if (app == null) continue;
            List<object> mine = new List<object>();
            try { foreach (dynamic doc in Docs(kind, app)) if (Hidden.Contains((string)doc.FullName)) mine.Add(doc); }
            catch (Exception) { }
            foreach (dynamic doc in mine)
            {
                string full = "";
                try
                {
                    full = (string)doc.FullName;
                    bool unsaved = !Truthy(doc.Saved) && ((string)doc.Path).Length > 0;
                    // No one can see this document, so work left unsaved would simply be lost: it is kept.
                    if (unsaved) { if (kind == "word") FinishMath(doc); doc.Save(); }
                    if (kind == "word") doc.Close(0); else if (kind == "excel") doc.Close(false); else { doc.Saved = -1; doc.Close(); }
                    said.Add((unsaved ? "saved and closed " : "closed ") + full);
                }
                catch (Exception error) { said.Add("could not close " + full + ": " + error.Message.Trim()); }
                Hidden.Remove(full);
            }
            try
            {
                bool seen = false;
                try { seen = Truthy(app.Visible); } catch (Exception) { }
                if (kind != "ppt" && (int)Docs(kind, app).Count == 0) { app.Quit(); Own.Remove(kind); }
                else if (kind != "ppt" && !seen) { }   // still holds something of ours: stays for the next turn
                else if (kind == "ppt")
                {
                    // Ours to end only if we started it, it holds nothing and shows nothing; else it is just let go of.
                    bool shown = false;
                    try { shown = Truthy(app.Visible); } catch (Exception) { }
                    if (Started.Contains(kind) && !shown && (int)app.Presentations.Count == 0) { app.Quit(); Started.Remove(kind); }
                    Own.Remove(kind);
                }
            }
            catch (Exception) { }
            app = null;
        }
        // The apps go only once nothing here refers to them any more.
        GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); GC.WaitForPendingFinalizers();
        if (Own.Count == 0) { try { File.Delete(OwnNotePath()); } catch (Exception) { } }
        return said;
    }

    static dynamic App(string kind, bool start)
    {
        dynamic app = Running(kind);
        if (app != null) return app;
        if (!start) throw new Fail("NOT_RUNNING", AppName(kind) + " is not running. Open the file with office_open first.");
        Type type = Type.GetTypeFromProgID(ProgId(kind));
        if (type == null) throw new Fail("NOT_INSTALLED", AppName(kind) + " is not installed on this computer.");
        app = Activator.CreateInstance(type);
        Started.Add(kind);
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
        if (cmd == "card" && a.Has("probe")) return Card.Probe(a.Flag("press", false));
        if (cmd == "card")
        {
            // The agent's turn ended: say so and let the card go.
            Card.Hold = a.Flag("hold", false);
            if (cardStarted) Card.Report(a.Str("text", "AI 已完成"), a.Num("seconds", 2.5));
            return "ok";
        }
        if (cmd == "status") return Status();
        if (cmd == "settle") return Settle();
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
        dynamic app = AppFor(kind, a);
        WaitReady(kind, app);
        dynamic doc = Doc(kind, app, a);
        if (cardStarted && CardOn && !Hidden.Contains((string)doc.FullName) && (cmd == "read" || cmd == "render" || cmd == "save"))
        {
            // Between edits the agent reads and looks: the card stays, and its switch still leads to the document.
            Remember(doc);
            string name = "";
            try { name = (string)doc.Name; } catch (Exception) { }
            Card.Report((cmd == "save" ? "AI 已保存 " : cmd == "render" ? "AI 正在检查 " : "AI 正在读取 ") + name, Linger);
        }
        if (cmd == "read") return kind == "word" ? WordRead(doc, a) : kind == "excel" ? ExcelRead(doc, a) : PptRead(doc, a);
        if (cmd == "edit") return Edit(kind, app, doc, a);
        if (cmd == "render") return Render(kind, app, doc, a);
        if (cmd == "save") return Save(kind, app, doc, a);
        if (cmd == "close") return Close(kind, doc, a);
        if (cmd == "view")
        {
            // Where the window is looking (for checking the follow mode).
            Dictionary<string, object> view = new Dictionary<string, object>();
            if (kind == "word")
            {
                dynamic window = doc.Windows[1];
                view["scrolled"] = (int)window.VerticalPercentScrolled;
                view["windows"] = (int)doc.Windows.Count;
                view["viewType"] = (int)window.View.Type;
                view["selection"] = (int)window.Selection.Start;
            }
            return view;
        }
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
            try
            {
                dynamic own = OwnApp(kind, false);
                if (own != null)
                {
                    HashSet<string> listed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (object item in open) { Dictionary<string, object> one = item as Dictionary<string, object>; if (one != null) listed.Add(Convert.ToString(one["path"] ?? one["name"])); }
                    foreach (dynamic doc in Docs(kind, own))
                    {
                        Dictionary<string, object> one = Info(kind, doc, null);
                        if (listed.Contains(Convert.ToString(one["path"] ?? one["name"]))) continue;
                        one["background"] = true; open.Add(one);
                    }
                }
                foreach (object item in open) { Dictionary<string, object> one = item as Dictionary<string, object>; if (one != null && one["path"] != null && Hidden.Contains((string)one["path"])) one["background"] = true; }
            }
            catch (Exception) { }
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
        if (a.Flag("silent", false)) return OpenQuietly(kind, path);
        // A document that is open in the background is brought out rather than opened a second time.
        dynamic hiding = OwnApp(kind, false);
        dynamic app = null, doc = null;
        if (hiding != null && path != null)
        {
            dynamic there = Find(kind, hiding, path);
            if (there != null) { Reveal(kind, hiding, there); app = hiding; doc = there; }
        }
        if (app == null)
        {
            app = App(kind, true);
            Show(kind, app);
            WaitReady(kind, app);
            doc = path == null ? null : Find(kind, app, path);
        }
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
        Remember(doc);
        if (a.Flag("card", false))
        {
            if (!cardStarted) { cardStarted = true; Card.Start(); }
            Following = Card.FollowChosen ? Following : a.Flag("follow", false);
            Typing = Card.TypingChosen ? Typing : a.Flag("typing", false);
            CardOn = true;
            string opened = "";
            try { opened = (string)doc.Name; } catch (Exception) { }
            Card.Report("AI 已打开 " + opened, Linger);
        }
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
        if (kind == "word") Refresh(doc);
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
        string closing = "";
        try { closing = (string)doc.FullName; } catch (Exception) { }
        bool ours = Hidden.Remove(closing);
        Revealed.Remove(closing);
        dynamic mine = kind == "ppt" ? null : OwnApp(kind, false);
        bool inOwn = false;
        try { inOwn = mine != null && Find(kind, mine, closing) != null; } catch (Exception) { }
        if (inOwn)
        {
            if (kind == "word") doc.Close(save ? -1 : 0); else doc.Close(save);
            try { if ((int)Docs(kind, mine).Count == 0) { mine.Quit(); Own.Remove(kind); } } catch (Exception) { }
            return "closed";
        }
        if (kind == "word") doc.Close(save ? -1 : 0);
        else if (kind == "excel") doc.Close(save);
        else { if (save) doc.Save(); else doc.Saved = -1; doc.Close(); }
        // An app this helper started itself, now holding nothing, would be left as an empty window: it goes too.
        if (Started.Contains(kind))
        {
            try
            {
                dynamic app = Running(kind);
                if (app != null && (int)(kind == "word" ? app.Documents.Count : kind == "excel" ? app.Workbooks.Count : app.Presentations.Count) == 0)
                {
                    app.Quit();
                    Started.Remove(kind);
                    return "closed (and " + AppName(kind) + ", which was started for it, was closed too)";
                }
            }
            catch (Exception) { }
        }
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


    // ───────────────────────── watching the work ─────────────────────────

    /// Scroll the window to where the agent is working (off = silent mode: the user's view is left alone).
    public static volatile bool Following;
    /// The window of the document being worked on (0 when unknown), for the card to bring forward.
    public static IntPtr DocWindow = IntPtr.Zero;

    public static void Trace(string text)
    {
        if (Environment.GetEnvironmentVariable("DSH_OFFICE_DEBUG") == "1") Console.Error.WriteLine("TRACE " + text);
    }

    delegate bool EachWindow(IntPtr window, IntPtr state);
    [DllImport("user32.dll")] static extern bool EnumWindows(EachWindow each, IntPtr state);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr window, StringBuilder text, int max);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr window, StringBuilder text, int max);

    /// The top-level window that shows a document, found by its caption among the windows of the app's process.
    /// PowerPoint does not always hand out the window of a presentation; its frame carries the file name.
    static IntPtr WindowOf(dynamic doc)
    {
        string name = "";
        try { name = (string)doc.Name; } catch (Exception) { }
        if (name.Length == 0) return IntPtr.Zero;
        string bare = Path.GetFileNameWithoutExtension(name);
        uint owner = 0;
        try { GetWindowThreadProcessId(new IntPtr(Convert.ToInt64(doc.Application.HWND)), out owner); } catch (Exception) { }
        IntPtr best = IntPtr.Zero;
        int bestScore = 0;
        EnumWindows(delegate(IntPtr window, IntPtr state)
        {
            if (!IsWindowVisible(window)) return true;
            StringBuilder caption = new StringBuilder(512), kind = new StringBuilder(128);
            GetWindowText(window, caption, 512);
            string title = caption.ToString();
            if (title.IndexOf(bare, StringComparison.OrdinalIgnoreCase) < 0) return true;
            GetClassName(window, kind, 128);
            uint pid;
            GetWindowThreadProcessId(window, out pid);
            string type = kind.ToString();
            bool frame = type == "PPTFrameClass" || type == "OpusApp" || type == "XLMAIN";
            if (!frame && (owner == 0 || pid != owner)) return true;
            int score = (frame ? 4 : 0) + (owner != 0 && pid == owner ? 2 : 0) + (title.StartsWith(name, StringComparison.OrdinalIgnoreCase) || title.StartsWith(bare, StringComparison.OrdinalIgnoreCase) ? 1 : 0);
            if (score > bestScore) { bestScore = score; best = window; }
            return true;
        }, IntPtr.Zero);
        return best;
    }

    static void Remember(dynamic doc)
    {
        IntPtr window = IntPtr.Zero;
        try { window = new IntPtr(Convert.ToInt64(doc.Windows[1].HWND)); }
        catch (Exception error) { Trace("no window handle from the document: " + error.GetType().Name); }
        // PowerPoint hands out the pane inside its frame; the frame is what comes to the front and what
        // "the user is looking at it" is checked against.
        IntPtr root = window != IntPtr.Zero ? GetAncestor(window, 2) : IntPtr.Zero;
        if (root == IntPtr.Zero) root = window;
        if (root == IntPtr.Zero) { try { root = WindowOf(doc); } catch (Exception) { } }
        DocWindow = root;
        Trace("document window " + window + " frame " + root);
    }
    /// Write text a few characters at a time, the way a person types, instead of all at once.
    public static volatile bool Typing;

    static bool cardStarted;
    /// Whether the user wants the card (the setting, as of the last edit).
    static bool CardOn;
    /// How long the card stays after a step, in seconds: about the pause between two steps of the agent.
    const double Linger = 25;

    static void Follow(dynamic doc, dynamic range)
    {
        if (!Following) return;
        try { doc.Windows[1].ScrollIntoView(range, true); } catch (Exception) { }
    }

    /// Pieces to write a text in: up to 14 of them, so that a long paragraph takes about a third of a second.
    static List<string> Pieces(string text)
    {
        List<string> pieces = new List<string>();
        int size = Math.Max(3, (int)Math.Ceiling(text.Length / 14.0));
        for (int i = 0; i < text.Length; i += size) pieces.Add(text.Substring(i, Math.Min(size, text.Length - i)));
        return pieces;
    }

    /// Put text at a position of a Word document; returns the position after it.
    static int Write(dynamic doc, int at, string text)
    {
        if (text.Length == 0) return at;
        if (!Typing || text.Length < 4)
        {
            doc.Range(at, at).InsertAfter(text);
            return at + text.Length;
        }
        foreach (string piece in Pieces(text))
        {
            doc.Range(at, at).InsertAfter(piece);
            at += piece.Length;
            Thread.Sleep(14);
        }
        return at;
    }

    // ───────────────────────── equations ─────────────────────────

    static readonly Dictionary<string, string> TexSymbols = new Dictionary<string, string>
    {
        { "alpha", "α" }, { "beta", "β" }, { "gamma", "γ" }, { "delta", "δ" }, { "epsilon", "ε" }, { "varepsilon", "ε" }, { "zeta", "ζ" }, { "eta", "η" },
        { "theta", "θ" }, { "vartheta", "ϑ" }, { "iota", "ι" }, { "kappa", "κ" }, { "lambda", "λ" }, { "mu", "μ" }, { "nu", "ν" }, { "xi", "ξ" },
        { "pi", "π" }, { "rho", "ρ" }, { "sigma", "σ" }, { "tau", "τ" }, { "upsilon", "υ" }, { "phi", "φ" }, { "varphi", "φ" }, { "chi", "χ" },
        { "psi", "ψ" }, { "omega", "ω" }, { "Gamma", "Γ" }, { "Delta", "Δ" }, { "Theta", "Θ" }, { "Lambda", "Λ" }, { "Xi", "Ξ" }, { "Pi", "Π" },
        { "Sigma", "Σ" }, { "Phi", "Φ" }, { "Psi", "Ψ" }, { "Omega", "Ω" },
        { "cdot", "⋅" }, { "times", "×" }, { "div", "÷" }, { "pm", "±" }, { "mp", "∓" }, { "approx", "≈" }, { "neq", "≠" }, { "ne", "≠" },
        { "leq", "≤" }, { "le", "≤" }, { "geq", "≥" }, { "ge", "≥" }, { "ll", "≪" }, { "gg", "≫" }, { "equiv", "≡" }, { "propto", "∝" }, { "sim", "∼" },
        { "infty", "∞" }, { "partial", "∂" }, { "nabla", "∇" }, { "sum", "∑" }, { "prod", "∏" }, { "int", "∫" }, { "oint", "∮" }, { "iint", "∬" }, { "iiint", "∭" }, { "oiint", "∯" },
        { "to", "→" }, { "rightarrow", "→" }, { "leftarrow", "←" }, { "Rightarrow", "⇒" }, { "Leftrightarrow", "⇔" }, { "in", "∈" }, { "notin", "∉" },
        { "subset", "⊂" }, { "cup", "∪" }, { "cap", "∩" }, { "forall", "∀" }, { "exists", "∃" }, { "angle", "∠" }, { "perp", "⊥" }, { "parallel", "∥" },
        { "circ", "∘" }, { "degree", "°" }, { "ldots", "…" }, { "cdots", "⋯" }, { "dots", "…" }, { "prime", "′" }, { "hbar", "ℏ" }, { "ell", "ℓ" },
        { "Vert", "‖" }, { "lVert", "‖" }, { "rVert", "‖" }, { "vert", "|" }, { "lvert", "|" }, { "rvert", "|" }, { "mid", "|" },
        { "top", "⊤" }, { "langle", "⟨" }, { "rangle", "⟩" }, { "otimes", "⊗" }, { "oplus", "⊕" }, { "neg", "¬" }, { "ast", "∗" },
        { "det", "det" }, { "rank", "rank" }, { "dim", "dim" }, { "sup", "sup" }, { "inf", "inf" }, { "arg", "arg" }, { "cond", "cond" },
        { "quad", " " }, { "qquad", "  " }, { "left", "" }, { "right", "" }, { "displaystyle", "" }, { "limits", "" }, { "lim", "lim" },
        { "sin", "sin" }, { "cos", "cos" }, { "tan", "tan" }, { "cot", "cot" }, { "ln", "ln" }, { "log", "log" }, { "exp", "exp" }, { "max", "max" }, { "min", "min" },
        { "arcsin", "arcsin" }, { "arccos", "arccos" }, { "arctan", "arctan" }, { "sinh", "sinh" }, { "cosh", "cosh" }, { "tanh", "tanh" },
    };

    /// One TeX argument starting at i: a {group}, a \command, or a single character. Returns it converted.
    static string TexArg(string s, ref int i)
    {
        while (i < s.Length && s[i] == ' ') i++;
        if (i >= s.Length) return "";
        if (s[i] == '{')
        {
            int depth = 0, start = i + 1;
            for (; i < s.Length; i++)
            {
                if (s[i] == '{') depth++;
                else if (s[i] == '}' && --depth == 0) break;
            }
            string inner = s.Substring(start, Math.Max(0, Math.Min(i, s.Length) - start));
            i++;
            return Tex(inner);
        }
        if (s[i] == '\\')
        {
            int start = i;
            i++;
            while (i < s.Length && char.IsLetter(s[i])) i++;
            if (i == start + 1 && i < s.Length) i++;
            return Tex(s.Substring(start, i - start));
        }
        return s[i++].ToString();
    }

    /// A space ends the fraction or root just written; without it Word pulls what follows into the denominator.
    static void Gap(StringBuilder o, string s, int i)
    {
        while (i < s.Length && s[i] == ' ') i++;
        if (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '\\' || s[i] == '(' || s[i] == '{')) o.Append(' ');
    }

    /// The common part of LaTeX math to Word's linear format (UnicodeMath), which Word then builds into a native equation.
    static string Tex(string s)
    {
        StringBuilder o = new StringBuilder();
        int i = 0;
        // Integral signs written one after another are one multiple integral.
        if (s.IndexOf("\\int", StringComparison.Ordinal) >= 0)
        {
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\\int(?:\s|\\!)*\\int(?:\s|\\!)*\\int(?![a-zA-Z])", "\\iiint");
            s = System.Text.RegularExpressions.Regex.Replace(s, @"\\int(?:\s|\\!)*\\int(?![a-zA-Z])", "\\iint");
        }
        // A sum, product or integral was written and its limits may still follow; what comes after them is its
        // operand, which Word's linear format introduces with "▒" (left out, Word draws an empty box there).
        bool nary = false;
        while (i < s.Length)
        {
            char c = s[i];
            // Nothing may stand between the limits and "▒": a space there ends the operator with an empty operand.
            if (nary && c == ' ') { i++; continue; }
            if (nary && c == '\\' && i + 1 < s.Length && ",;:! ".IndexOf(s[i + 1]) >= 0) { i += 2; continue; }
            if (nary && c != '^' && c != '_')
            {
                if (c == '\\' && string.CompareOrdinal(s, i, "\\limits", 0, 7) == 0) { i += 7; continue; }
                if (c == '\\' && string.CompareOrdinal(s, i, "\\nolimits", 0, 9) == 0) { i += 9; continue; }
                o.Append('▒');
                nary = false;
            }
            if (c == '\\')
            {
                int start = ++i;
                while (i < s.Length && char.IsLetter(s[i])) i++;
                if (i == start)
                {
                    // \, \; \! \  are spacing; \{ \} \% and the like are the character itself.
                    char next = i < s.Length ? s[i++] : ' ';
                    if (next == '|') o.Append('‖');
                    else if (next == ',' || next == ';' || next == ':' || next == ' ') o.Append(' ');
                    else if (next == '\\') o.Append(' ');
                    else if (next != '!') o.Append(next);
                    continue;
                }
                string name = s.Substring(start, i - start);
                if (name == "begin")
                {
                    // \begin{pmatrix} a & b \\ c & d \end{pmatrix}: Word's linear form is (■(a&b@c&d)).
                    int close = s.IndexOf('}', i);
                    string env = close > i ? s.Substring(i + 1, close - i - 1).Trim() : "";
                    string ending = "\\end{" + env + "}";
                    int stop = close > i ? s.IndexOf(ending, close, StringComparison.Ordinal) : -1;
                    if (stop > 0)
                    {
                        string body = s.Substring(close + 1, stop - close - 1);
                        i = stop + ending.Length;
                        if (env == "array" && body.TrimStart().StartsWith("{", StringComparison.Ordinal)) body = body.Substring(body.IndexOf('}') + 1);
                        List<string> rows = new List<string>();
                        foreach (string row in body.Replace("\\\\", "\u0001").Split('\u0001'))
                        {
                            if (row.Trim().Length == 0) continue;
                            List<string> cells = new List<string>();
                            foreach (string cell in row.Split('&')) cells.Add(Tex(cell.Trim()));
                            rows.Add(string.Join("&", cells.ToArray()));
                        }
                        string grid = string.Join("@", rows.ToArray());
                        if (env == "pmatrix") o.Append("(■(" + grid + "))");
                        else if (env == "bmatrix") o.Append("[■(" + grid + ")]");
                        else if (env == "vmatrix") o.Append("|■(" + grid + ")|");
                        else if (env == "Vmatrix") o.Append("‖■(" + grid + ")‖");
                        else if (env == "cases") o.Append("{█(" + grid + ")┤");
                        else if (env == "aligned" || env == "align" || env == "split" || env == "gathered") o.Append("█(" + grid + ")");
                        else o.Append("■(" + grid + ")");
                        continue;
                    }
                }
                if (name == "mathbb")
                {
                    string letter = TexArg(s, ref i);
                    string doubled = letter == "R" ? "ℝ" : letter == "N" ? "ℕ" : letter == "Z" ? "ℤ" : letter == "Q" ? "ℚ" : letter == "C" ? "ℂ" : letter;
                    o.Append(doubled);
                    continue;
                }
                if (name == "frac" || name == "dfrac" || name == "tfrac") {
                    string top = TexArg(s, ref i), bottom = TexArg(s, ref i);
                    // A factor written right before the fraction (2rac{a}{b}) must not be read into its numerator.
                    // After a closing bracket a space is not enough for Word: the fraction is fenced off instead.
                    bool fenced = o.Length > 0 && ")]〗".IndexOf(o[o.Length - 1]) >= 0;
                    if (o.Length > 0 && char.IsLetterOrDigit(o[o.Length - 1])) o.Append(' ');
                    // A function applied to a fraction (\sin\frac{a}{b}) takes the whole fraction, not just its numerator.
                    bool applied = System.Text.RegularExpressions.Regex.IsMatch(o.ToString(), @"(sin|cos|tan|cot|sinh|cosh|tanh|ln|log|exp)([\^_](\([^()]*\)|[^\s()]+))*\s\z");
                    o.Append(applied || fenced ? "〖(" + top + ")/(" + bottom + ")〗" : "(" + top + ")/(" + bottom + ")");
                    Gap(o, s, i);
                }
                else if (name == "sqrt")
                {
                    string degree = null;
                    if (i < s.Length && s[i] == '[') { int end = s.IndexOf(']', i); if (end > i) { degree = Tex(s.Substring(i + 1, end - i - 1)); i = end + 1; } }
                    string body = TexArg(s, ref i);
                    o.Append(degree == null ? "√(" + body + ")" : "√(" + degree + "&" + body + ")");
                    Gap(o, s, i);
                }
                else if (name == "bar" || name == "overline") o.Append("¯(" + TexArg(s, ref i) + ")");
                else if (name == "hat") o.Append(TexArg(s, ref i) + "\u0302");
                else if (name == "vec") o.Append(TexArg(s, ref i) + "\u20D7");
                else if (name == "dot") o.Append(TexArg(s, ref i) + "\u0307");
                else if (name == "ddot") o.Append(TexArg(s, ref i) + "\u0308");
                else if (name == "tilde") o.Append(TexArg(s, ref i) + "\u0303");
                else if (name == "text" || name == "mathrm" || name == "textrm" || name == "operatorname")
                {
                    while (i < s.Length && s[i] == ' ') i++;
                    if (i < s.Length && s[i] == '{') { int end = s.IndexOf('}', i); if (end < 0) end = s.Length; o.Append("\"" + s.Substring(i + 1, end - i - 1) + "\""); i = Math.Min(s.Length, end + 1); }
                }
                else if (name == "mathbf" || name == "boldsymbol" || name == "mathit" || name == "mathbb" || name == "mathcal") o.Append(TexArg(s, ref i));
                else if (name == "not")
                {
                    // \not\equiv, \not=, \not\in: the negated sign itself.
                    while (i < s.Length && s[i] == ' ') i++;
                    string rest = s.Substring(i);
                    string[,] negated = { { "\\equiv", "≢" }, { "\\in", "∉" }, { "\\subset", "⊄" }, { "\\approx", "≉" }, { "\\sim", "≁" }, { "\\leq", "≰" }, { "\\geq", "≱" }, { "\\le", "≰" }, { "\\ge", "≱" }, { "\\parallel", "∦" }, { "\\mid", "∤" }, { "\\exists", "∄" }, { "=", "≠" }, { "<", "≮" }, { ">", "≯" } };
                    bool done = false;
                    for (int k = 0; k < negated.GetLength(0) && !done; k++)
                    {
                        string from = negated[k, 0];
                        if (!rest.StartsWith(from, StringComparison.Ordinal)) continue;
                        if (from[0] == '\\' && rest.Length > from.Length && char.IsLetter(rest[from.Length])) continue;
                        o.Append(negated[k, 1]);
                        i += from.Length;
                        done = true;
                    }
                    if (!done) o.Append('¬');
                }
                else if (name == "left" || name == "right" || name == "bigl" || name == "bigr" || name == "big" || name == "Big" || name == "Bigl" || name == "Bigr")
                {
                    // "\left." and "\right." are the invisible partner of a one-sided bracket.
                    if (i < s.Length && s[i] == '.') i++;
                }
                else
                {
                    string symbol;
                    if (TexSymbols.TryGetValue(name, out symbol))
                    {
                        // A function name right after a letter (r\cos u) would be read as one word with it.
                        if (symbol.Length > 1 && symbol == name && o.Length > 0 && char.IsLetter(o[o.Length - 1])) o.Append(' ');
                        o.Append(symbol);
                        if (symbol == "∑" || symbol == "∏" || symbol == "∫" || symbol == "∮" || symbol == "∬" || symbol == "∭" || symbol == "∯") nary = true;
                        // A function name is applied to what follows: Word needs a space after it.
                        if (symbol.Length > 1 && symbol == name && i < s.Length && s[i] != ' ' && s[i] != '^' && s[i] != '_') o.Append(' ');
                    }
                    else o.Append('\\').Append(name);
                }
                continue;
            }
            if (c == '^' || c == '_')
            {
                i++;
                bool braced = i < s.Length && s[i] == '{';
                string arg = TexArg(s, ref i);
                o.Append(c).Append(braced && arg.Length > 1 ? "(" + arg + ")" : arg);
                // A space ends the script; otherwise Word reads what follows into it: x^2y, λ_max(A).
                if (!nary && i < s.Length && "^_ )]},.;:，。、".IndexOf(s[i]) < 0) o.Append(' ');
                continue;
            }
            if (c == '{') { string group = TexArg(s, ref i); o.Append("〖" + group + "〗"); continue; }
            if (c == '}') { i++; continue; }
            if (c == '~') { o.Append(' '); i++; continue; }
            if (c == '\'')
            {
                // X'' is one double prime, not two marks set apart.
                int marks = 0;
                while (i < s.Length && s[i] == '\'') { marks++; i++; }
                o.Append(marks == 1 ? "′" : marks == 2 ? "″" : marks == 3 ? "‴" : new string('′', marks));
                continue;
            }
            o.Append(c);
            i++;
        }
        return o.ToString();
    }

    static readonly System.Text.RegularExpressions.Regex Dollars = new System.Text.RegularExpressions.Regex(@"\$\$([^$\r]+?)\$\$|\$([^$\r]+?)\$(?!\$)");

    static readonly System.Text.RegularExpressions.Regex Tag = new System.Text.RegularExpressions.Regex(@"\\tag\s*\{([^}]*)\}");

    /// Turn every $...$ (inline) and $$...$$ (on a line of its own) inside a range into a native Word equation.
    static int Mathify(dynamic doc, dynamic range)
    {
        string text = (string)range.Text;
        if (text == null || text.IndexOf('$') < 0) return 0;
        int start = (int)range.Start, made = 0;
        System.Text.RegularExpressions.MatchCollection found = Dollars.Matches(text);
        if (found.Count == 0) return 0;
        // All the formulas of this piece of text are read by Word in one round.
        List<string> sources = new List<string>();
        foreach (System.Text.RegularExpressions.Match m in found)
        {
            string source = (m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).Trim();
            System.Text.RegularExpressions.Match tagged = Tag.Match(source);
            if (tagged.Success) source = source.Remove(tagged.Index, tagged.Length).Trim();
            sources.Add(source);
        }
        Dictionary<string, string> linear = ToLinear(sources);
        UserLinear(doc.Application);
        for (int k = found.Count - 1; k >= 0; k--)
        {
            System.Text.RegularExpressions.Match m = found[k];
            bool display = m.Groups[1].Success;
            string source = (display ? m.Groups[1].Value : m.Groups[2].Value).Trim();
            try
            {
                if (BuildOne(doc, range, start, m, display, source, linear)) made++;
                else MathFailed++;
            }
            catch (Exception) { MathFailed++; }
        }
        return made;
    }

    /// Equations on a line of their own built during the current operation (their ranges: these keep pointing at the
    /// equation while the document grows). They are set off from the text once the operation is done: a paragraph
    /// inserted after one takes over its look, and must not take over this.
    static readonly List<object> Displays = new List<object>();

    /// As typeset papers do: some space above and below the equation, and a line that grows with the formula
    /// (an exact line height would cut a tall one off).
    static void SetOff()
    {
        foreach (dynamic equation in Displays)
        {
            try
            {
                dynamic paragraph = equation.Paragraphs[1];
                dynamic range = paragraph.Range;
                dynamic math = range.OMaths[1].Range;
                if ((int)range.OMaths.Count != 1) continue;
                if ((int)range.End - (int)range.Start > (int)math.End - (int)math.Start + 3) continue;
                dynamic format = paragraph.Format;
                if ((float)format.SpaceBefore < 6f) format.SpaceBefore = 6f;
                if ((float)format.SpaceAfter < 6f) format.SpaceAfter = 6f;
                // Centred on the line, not on what the first-line indent leaves of it.
                format.CharacterUnitFirstLineIndent = 0;
                format.FirstLineIndent = 0;
            }
            catch (Exception) { }
        }
        Displays.Clear();
    }

    /// Set when an equation was built or an exact line height was set: the document is then checked by Roomy().
    static bool Cramped;

    /// An exact line height cuts off whatever is taller than it: fractions, sums, integrals. Every paragraph that
    /// holds an equation and has an exact height gets the same height as a minimum instead, so its lines stay as they
    /// are and only the ones with a tall formula grow. Paragraphs without equations keep the exact height.
    static int Roomy(dynamic doc)
    {
        Cramped = false;
        int changed = 0, last = -1;
        try
        {
            int maths = (int)doc.OMaths.Count, count = maths + (int)doc.InlineShapes.Count;
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    dynamic paragraph = (i <= maths ? doc.OMaths[i].Range : doc.InlineShapes[i - maths].Range).Paragraphs[1];
                    if (i == maths + 1) last = -1;
                    int start = (int)paragraph.Range.Start;
                    if (start == last) continue;
                    last = start;
                    dynamic format = paragraph.Format;
                    if ((int)format.LineSpacingRule != 4) continue;
                    float height = (float)format.LineSpacing;
                    format.LineSpacingRule = 3;
                    format.LineSpacing = height;
                    changed++;
                }
                catch (COMException) { }
            }
        }
        catch (COMException) { }
        return changed;
    }

    /// Formulas that could not be turned into equations during the current operation (they stay as text).
    static int MathFailed;

    static string MathNote()
    {
        SetOff();
        if (MathFailed == 0) return "";
        string note = " — NOTE " + MathFailed + " formula(s) could not be built and were left as text: check them with office_render and rewrite them more simply";
        MathFailed = 0;
        return note;
    }

    // Word reads LaTeX itself once its equation input is switched to LaTeX (the "LaTeX" button on the Equation
    // tab). The helper does that in a Word of its own (see MathSource), never in the one the user works in: there,
    // each formula is read as LaTeX, and Word's own linear form of the result is what gets written into the document.

    /// What Word's LaTeX reader does not take: dropped or rewritten before it sees the formula.
    static string ForWord(string source)
    {
        string s = source.Replace("\\displaystyle", "").Replace("\\textstyle", "").Replace("\\nolimits", "").Replace("\\limits", "");
        s = s.Replace("\\dfrac", "\\frac").Replace("\\tfrac", "\\frac").Replace("\\lVert", "\\|").Replace("\\rVert", "\\|").Replace("\\lvert", "|").Replace("\\rvert", "|");
        s = s.Replace("\\ne ", "\\neq ").Replace("\\le ", "\\leq ").Replace("\\ge ", "\\geq ").Replace("~", "\\ ");
        // One double prime instead of two marks set apart.
        s = s.Replace("\'\'\'", "‴").Replace("\'\'", "″");
        return s.Trim();
    }

    static readonly System.Text.RegularExpressions.Regex Environments = new System.Text.RegularExpressions.Regex(@"\\begin\{(\w+\*?)\}");

    /// Environments Word's LaTeX reader builds; the others (aligned, array ..) go through the helper's own conversion.
    static bool WordReads(string source)
    {
        foreach (System.Text.RegularExpressions.Match m in Environments.Matches(source))
        {
            string name = m.Groups[1].Value;
            if (name != "matrix" && name != "pmatrix" && name != "bmatrix" && name != "vmatrix" && name != "Vmatrix" && name != "cases") return false;
        }
        return true;
    }

    /// Left as it was typed: a command Word did not know, or a brace or ampersand that should have been consumed.
    static bool Unbuilt(string built, string source)
    {
        if (built == null) return true;
        if (built.IndexOf('\\') >= 0) return true;
        if (built.IndexOf('&') >= 0 && source.IndexOf("\\&", StringComparison.Ordinal) < 0) return true;
        if ((built.IndexOf('{') >= 0 || built.IndexOf('}') >= 0) && source.IndexOf("\\{", StringComparison.Ordinal) < 0 && source.IndexOf("\\}", StringComparison.Ordinal) < 0 && source.IndexOf("cases", StringComparison.Ordinal) < 0) return true;
        return false;
    }

    /// LaTeX to Word's linear form, as read by Word; formulas met before are remembered. A formula Word cannot read
    /// is missing from the result.
    static readonly Dictionary<string, string> Read = new Dictionary<string, string>();

    static Dictionary<string, string> ToLinear(List<string> sources)
    {
        Dictionary<string, string> result = new Dictionary<string, string>();
        List<string> fresh = new List<string>();
        foreach (string source in sources)
        {
            string known;
            if (Read.TryGetValue(source, out known)) { if (known != null) result[source] = known; }
            else if (!fresh.Contains(source) && WordReads(source) && source.Length > 0) fresh.Add(source);
        }
        if (fresh.Count == 0 || !MathSource()) return result;
        try
        {
            dynamic bars = MathWord.CommandBars;
            MathDoc.Content.Delete();
            for (int i = 0; i < fresh.Count; i++)
            {
                if (i > 0) MathDoc.Content.InsertParagraphAfter();
                MathDoc.Paragraphs.Last.Range.InsertBefore(ForWord(fresh[i]));
            }
            bool[] built = new bool[fresh.Count];
            for (int i = 0; i < fresh.Count; i++)
            {
                try
                {
                    dynamic line = MathDoc.Paragraphs[i + 1].Range;
                    line.MoveEnd(1, -1);
                    dynamic math = MathDoc.OMaths.Add(line).OMaths[1];
                    math.BuildUp();
                    built[i] = !Unbuilt((string)MathDoc.Paragraphs[i + 1].Range.OMaths[1].Range.Text, fresh[i]);
                }
                catch (COMException) { built[i] = false; }
            }
            // Word's own linear form of what it built: that needs the linear input for a moment.
            bars.ExecuteMso("EquationUnicodeFormat");
            try
            {
                for (int i = 0; i < fresh.Count; i++)
                {
                    string value = null;
                    if (built[i])
                    {
                        try
                        {
                            dynamic line = MathDoc.Paragraphs[i + 1].Range;
                            line.OMaths[1].Linearize();
                            value = ((string)MathDoc.Paragraphs[i + 1].Range.OMaths[1].Range.Text ?? "").TrimEnd('\r');
                            if (value.Length == 0) value = null;
                            // Word writes the hat and the tilde of LaTeX as characters that stand beside the letter;
                            // read back, they would not go over it. The combining ones do.
                            if (value != null) value = value.Replace(" ˆ", " ̂").Replace(" ˜", " ̃").Replace(" ˇ", " ̌").Replace(" ˘", " ̆").Replace(" ´", " ́");
                        }
                        catch (COMException) { value = null; }
                    }
                    Read[fresh[i]] = value;
                    if (value != null) result[fresh[i]] = value;
                }
            }
            finally { bars.ExecuteMso("EquationLaTexFormat"); }
        }
        catch (Exception) { }
        return result;
    }

    /// The document's own Word must be reading the linear form while equations are built in it. It is, unless the
    /// user switched theirs to LaTeX: then it is switched for the batch and put back afterwards.
    static int UserMode;   // 0 = not looked at, 1 = linear already, 2 = switched by us, 3 = could not tell

    static void UserLinear(dynamic app)
    {
        if (UserMode != 0) return;
        try
        {
            dynamic bars = app.CommandBars;
            if (!(bool)bars.GetPressedMso("EquationLaTexFormat")) { UserMode = 1; return; }
            bars.ExecuteMso("EquationUnicodeFormat");
            UserMode = 2;
        }
        catch (Exception) { UserMode = 3; }
    }

    static bool BuildOne(dynamic doc, dynamic range, int start, System.Text.RegularExpressions.Match m, bool display, string source, Dictionary<string, string> linear)
    {
        // \tag{1} numbers a display equation: Word sets "#(1)" flush right on the equation's line.
        string tag = null;
        System.Text.RegularExpressions.Match tagged = Tag.Match(source);
        if (tagged.Success) { tag = tagged.Groups[1].Value.Trim(); source = source.Remove(tagged.Index, tagged.Length).Trim(); }
        if (!display) tag = null;
        dynamic spot = null;
        if (m.Length <= 250)
        {
            // Found by its text: positions counted in the text are off wherever the range holds a field (a caption number).
            dynamic search = range.Duplicate;
            if ((bool)search.Find.Execute(FindText: m.Value.Replace("^", "^^"), MatchCase: true, MatchWildcards: false, Forward: true, Wrap: 0)) spot = search;
        }
        if (spot == null) spot = doc.Range(start + m.Index, start + m.Index + m.Length);
        string read;
        bool byWord = linear.TryGetValue(source, out read);
        // Not read by Word (an environment it lacks, a command it does not know): the helper's own conversion.
        spot.Text = (byWord ? read : Tex(source)) + (string.IsNullOrEmpty(tag) ? "" : "#(" + tag + ")");
        try
        {
            dynamic math = doc.OMaths.Add(spot);
            dynamic built = math.OMaths[1];
            built.BuildUp();
            bool good = true;
            if (!byWord) { try { good = ((string)math.OMaths[1].Range.Text ?? "").IndexOf('\\') < 0; } catch (COMException) { } }
            if (display) { try { built.Type = 0; built.Justification = 1; } catch (COMException) { } }
            if (display) { try { Displays.Add(built.Range); } catch (Exception) { } }
            Cramped = true;
            return good;
        }
        catch (COMException) { return false; }
    }

    /// End of a batch in Word: the user's equation input is put back if it had to be switched.
    static void FinishMath(dynamic doc)
    {
        if (UserMode == 2)
        {
            try { dynamic bars = doc.Application.CommandBars; bars.ExecuteMso("EquationLaTexFormat"); }
            catch (Exception) { }
        }
        UserMode = 0;
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
        Follow(doc, p.Range);
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
        int shown = from - 1, openTable = -1;
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
                    dynamic table = p.Range.Tables[1];
                    int tableStart = (int)table.Range.Start, end = (int)table.Range.End;
                    if (tableStart != openTable)
                    {
                        int inside = (int)table.Range.Paragraphs.Count, cells = (int)table.Range.Cells.Count, rows = (int)table.Rows.Count;
                        bool nested = false;
                        try { nested = (int)table.NestingLevel > 1; } catch (Exception) { }
                        Dictionary<string, object> head = new Dictionary<string, object>();
                        head["i"] = i;
                        head["to"] = i + inside - 1;
                        if (!nested) head["table"] = (int)doc.Range(0, end).Tables.Count;
                        string size = "?";
                        try { size = rows + "x" + (int)table.Columns.Count; } catch (COMException) { size = cells + " cells"; }
                        head["size"] = size;
                        // A grid of short cells is one entry (read its cells with "table"). A table whose cells hold the
                        // document's text, as in a form or a report template, is listed paragraph by paragraph.
                        bool open = inside > cells + rows + 2;
                        if (open) head["open"] = true;
                        items.Add(head);
                        if (!open)
                        {
                            i += inside - 1;
                            p = i < total ? doc.Range(end, end).Paragraphs[1] : null;
                            shown = i;
                            continue;
                        }
                        openTable = tableStart;
                    }
                    bool rowEnd = false;
                    try { rowEnd = Truthy(p.Range.Information[31]); } catch (Exception) { }
                    if (rowEnd) { shown = i; p = p.Next(); continue; }
                    item["inTable"] = true;
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
        if (op.Has("latinFont")) { string latin = op.Need("latinFont"); try { range.Font.NameAscii = latin; range.Font.NameOther = latin; } catch (COMException) { } }
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
        if (op.Has("firstLineIndent")) Indent(range.ParagraphFormat, op.Raw("firstLineIndent"));
        if (op.Has("spaceBefore")) range.ParagraphFormat.SpaceBefore = (float)op.Num("spaceBefore", 0);
        if (op.Has("spaceAfter")) range.ParagraphFormat.SpaceAfter = (float)op.Num("spaceAfter", 0);
        if (op.Has("lineSpacing")) LineSpacing(range.ParagraphFormat, op.Raw("lineSpacing"));
        if (op.Has("indentChars")) { range.ParagraphFormat.FirstLineIndent = 0; range.ParagraphFormat.CharacterUnitFirstLineIndent = (float)op.Num("indentChars", 2); }
        if (op.Has("superscript")) range.Font.Superscript = op.Flag("superscript", false) ? 1 : 0;
        if (op.Has("subscript")) range.Font.Subscript = op.Flag("subscript", false) ? 1 : 0;
    }

    /// First-line indent as people state it: a small plain number is characters (2 = 首行缩进 2 字符; nobody means an
    /// indent of 2 points), "2字符" / "2 chars" too; "0.74cm" and "24pt" are lengths, and a larger plain number is points.
    static void Indent(dynamic format, object raw)
    {
        string text = Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture).Trim();
        System.Text.RegularExpressions.Match number = System.Text.RegularExpressions.Regex.Match(text, @"\d+(\.\d+)?");
        if (!number.Success) throw new Fail("BAD_ARGS", "firstLineIndent \"" + text + "\" is not understood: give characters (2), or a length (\"0.74cm\", \"24pt\").");
        float value = float.Parse(number.Value, System.Globalization.CultureInfo.InvariantCulture);
        string lower = text.ToLowerInvariant();
        bool cm = lower.Contains("cm") || text.Contains("厘米"), pt = lower.Contains("pt") || text.Contains("磅");
        bool chars = !cm && !pt && (lower.Contains("char") || text.Contains("字") || value <= 4);
        // Twice: on a paragraph that had an indent as a length, the first pass only removes that length and the
        // paragraph falls back to part of its style's indent; the second pass writes the explicit zero.
        for (int pass = 0; pass < 2; pass++)
        {
            format.CharacterUnitFirstLineIndent = 0;
            format.FirstLineIndent = 0;
        }
        if (value == 0) return;
        if (chars) format.CharacterUnitFirstLineIndent = value;
        else format.FirstLineIndent = cm ? value * 28.35f : value;
    }

    /// Line spacing as people state it: a plain number is a multiple of single spacing (1.5), a length is an exact
    /// height ("20pt", "固定值20磅"), and "at least 20pt" / "最小值20磅" a minimum. A plain number too large to be a
    /// multiple (5 and up) is taken as points.
    static void LineSpacing(dynamic format, object raw)
    {
        string text = Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture).Trim();
        System.Text.RegularExpressions.Match number = System.Text.RegularExpressions.Regex.Match(text, @"\d+(\.\d+)?");
        if (!number.Success) throw new Fail("BAD_ARGS", "lineSpacing \"" + text + "\" is not understood: give a multiple (1.5), an exact height (\"20pt\") or a minimum (\"at least 20pt\").");
        float value = float.Parse(number.Value, System.Globalization.CultureInfo.InvariantCulture);
        string lower = text.ToLowerInvariant();
        bool least = lower.Contains("least") || lower.Contains("min") || text.Contains("最小");
        bool points = least || lower.Contains("pt") || lower.Contains("exact") || lower.Contains("fixed") || text.Contains("磅") || text.Contains("固定") || value >= 5;
        if (!points) { format.LineSpacingRule = 5; format.LineSpacing = value * 12; }
        else { format.LineSpacingRule = least ? 3 : 4; format.LineSpacing = value; if (!least) Cramped = true; }
    }

    /// A new empty paragraph where an insert operation points: after / before paragraph "para", or at the start / end.
    static dynamic NewParagraph(dynamic doc, Bag op)
    {
        string where = op.Str("where", op.Has("para") ? "after" : "end");
        // A new, empty document: its one empty paragraph is where the text goes, wherever the operation points.
        if ((int)doc.Paragraphs.Count == 1 && (string)doc.Paragraphs[1].Range.Text == "\r")
        {
            op.Raw("para"); op.Raw("expect");
            return doc.Paragraphs[1];
        }
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

    static readonly System.Text.RegularExpressions.Regex Numbered = new System.Text.RegularExpressions.Regex(@"^\s*(第?[一二三四五六七八九十百]+[、．.章节]|[（(][一二三四五六七八九十\d]+[)）]|\d+(\.\d+)*[\s、．.])");

    static readonly System.Text.RegularExpressions.Regex PlainMath = new System.Text.RegularExpressions.Regex(@"[A-Za-z0-9\)\}]_(\{[^}]+\}|[A-Za-z0-9])|[A-Za-z0-9\)]\^(\{[^}]+\}|\(?-?\d)|\\(frac|sqrt|sum|int|alpha|beta|theta|pi)\b");

    /// Subscripts and powers written as plain text (d_k, 10^18) outside $...$: say so, with the first one found.
    static string Lint(string text)
    {
        if (text == null) return "";
        string outside = Dollars.Replace(text, " ");
        System.Text.RegularExpressions.Match m = PlainMath.Match(outside);
        if (!m.Success) return "";
        int start = Math.Max(0, m.Index - 6);
        return " — NOTE plain-text formula \"" + outside.Substring(start, Math.Min(outside.Length - start, m.Length + 12)).Trim() + "\": write formulas, subscripts and powers between dollar signs ($d_k$, $10^{18}$) so they are typeset";
    }

    /// Keep Latin words and numbers whole at the end of a line. Some templates carry "allow Latin text to wrap in
    /// the middle of a word" in a style, and every paragraph written after it inherits that: 37 / 000, mode / l.
    /// (Word's object model names this backwards: WordWrap = True is the whole-word setting.)
    static void Whole(dynamic range)
    {
        try { range.ParagraphFormat.WordWrap = -1; } catch (Exception) { }
    }

    /// The text already starts with its number ("一、", "2.1 ") and the style would add another: drop the automatic one.
    static void Unnumber(dynamic paragraph, string text)
    {
        if (!Numbered.IsMatch(text)) return;
        try { if ((int)paragraph.Range.ListFormat.ListType != 0) paragraph.Range.ListFormat.RemoveNumbers(); }
        catch (Exception) { }
    }

    /// A new empty paragraph right after the given one.
    static dynamic After(dynamic doc, dynamic paragraph)
    {
        int end = (int)paragraph.Range.End;
        paragraph.Range.InsertParagraphAfter();
        dynamic made = doc.Range(end, end).Paragraphs[1];
        // Word gives the new paragraph the look of whatever FOLLOWS (a caption, a heading); it should
        // continue the paragraph it was added after.
        try
        {
            made.Style = paragraph.Style;
            dynamic from = paragraph.Range.ParagraphFormat, to = made.Range.ParagraphFormat;
            to.Alignment = from.Alignment;
            to.LeftIndent = from.LeftIndent; to.RightIndent = from.RightIndent;
            to.CharacterUnitFirstLineIndent = from.CharacterUnitFirstLineIndent; to.FirstLineIndent = from.FirstLineIndent;
            to.SpaceBefore = from.SpaceBefore; to.SpaceAfter = from.SpaceAfter;
            // Only where it differs: written when equal, it stays behind as the paragraph's own setting and
            // no longer follows the style when that is changed later.
            if ((int)to.LineSpacingRule != (int)from.LineSpacingRule) to.LineSpacingRule = from.LineSpacingRule;
            if ((int)from.LineSpacingRule >= 3 && Math.Abs((float)to.LineSpacing - (float)from.LineSpacing) > 0.01f) to.LineSpacing = from.LineSpacing;
            dynamic font = paragraph.Range.Font, target = made.Range.Font;
            // 9999999 / empty mean "mixed" in the source paragraph: leave those to the style.
            string name = (string)font.Name, farEast = (string)font.NameFarEast;
            if (!string.IsNullOrEmpty(name)) target.Name = name;
            if (!string.IsNullOrEmpty(farEast)) target.NameFarEast = farEast;
            float size = (float)font.Size;
            if (size > 0 && size < 1000) target.Size = size;
            foreach (string flag in new string[] { "Bold", "Italic" })
            {
                int value = flag == "Bold" ? (int)font.Bold : (int)font.Italic;
                if (value == 0 || value == -1) { if (flag == "Bold") target.Bold = value; else target.Italic = value; }
            }
            int color = (int)font.Color;
            if (color != 9999999) target.Color = color;
        }
        catch (COMException) { }
        return made;
    }

    /// Breathing room around a table or figure: at least this much space before / after a paragraph.
    /// Paragraph spacing, not blank lines, so nothing has to be cleaned up later.
    const float Air = 8f;

    static void Space(dynamic paragraph, bool before)
    {
        if (paragraph == null) return;
        try
        {
            dynamic format = paragraph.Range.ParagraphFormat;
            if (before) { if ((float)format.SpaceBefore < Air) format.SpaceBefore = Air; }
            else if ((float)format.SpaceAfter < Air) format.SpaceAfter = Air;
        }
        catch (Exception) { }
    }

    /// A numbered caption, the way reports set them: "表 1 ..." centred above a table, "图 1 ..." centred below
    /// a figure. The number is a Word caption field, so it renumbers itself and can be cross-referenced.
    static void Caption(dynamic doc, dynamic target, bool table, string title, string fontName, string fontFarEast)
    {
        string label = table ? "表" : "图";
        dynamic labels = doc.Application.CaptionLabels;
        bool known = false;
        foreach (dynamic item in labels) if ((string)item.Name == label) { known = true; break; }
        if (!known) labels.Add(label);
        int start = (int)target.Start, end = (int)target.End;
        target.InsertCaption(Label: label, Title: "  " + title.Trim(), Position: table ? 0 : 1);
        // Above a table the caption is the paragraph that now sits where the table began; below a figure it is the next one.
        dynamic caption = table ? doc.Range(start, start).Paragraphs[1] : doc.Range(end, end).Paragraphs[1].Next();
        if (caption == null) return;
        try
        {
            dynamic range = caption.Range;
            range.ParagraphFormat.Alignment = 1;
            Whole(range);
            range.ParagraphFormat.CharacterUnitFirstLineIndent = 0; range.ParagraphFormat.FirstLineIndent = 0;
            range.Font.Bold = 0; range.Font.Italic = 0; range.Font.Color = -16777216; range.Font.Size = 10.5f;
            if (!string.IsNullOrEmpty(fontName)) range.Font.Name = fontName;
            if (!string.IsNullOrEmpty(fontFarEast)) range.Font.NameFarEast = fontFarEast;
            // A table's caption must not be left alone at the foot of a page.
            if (table) range.ParagraphFormat.KeepWithNext = -1;
            range.ParagraphFormat.SpaceBefore = table ? Air : 3f;
            range.ParagraphFormat.SpaceAfter = table ? 3f : Air;
            Mathify(doc, caption.Range);
        }
        catch (COMException) { }
    }

    /// The number of a paragraph, counted from the start of the document.
    static int Index(dynamic doc, dynamic paragraph)
    {
        return (int)doc.Range(0, paragraph.Range.End).Paragraphs.Count;
    }

    /// A new paragraph with no style named: body text that matches its surroundings. Written next to body
    /// text it keeps that text's font, size and spacing (a template's own look) but not the emphasis of the
    /// line it was split from (a bold label, a red hint); written after a heading it becomes Normal.
    static void Body(dynamic paragraph)
    {
        int level = 10;
        try { level = (int)paragraph.OutlineLevel; } catch (Exception) { }
        if (level != 10) { Plain(paragraph, "Normal"); return; }
        try
        {
            dynamic font = paragraph.Range.Font;
            font.Bold = 0; font.Italic = 0; font.Underline = 0; font.Color = -16777216;
            paragraph.Range.ParagraphFormat.Alignment = 3;
        }
        catch (COMException) { }
    }

    static void Plain(dynamic paragraph, string style)
    {
        SetStyle(paragraph.Range, style);
        try { paragraph.Range.Font.Reset(); paragraph.Range.ParagraphFormat.Reset(); } catch (COMException) { }
    }

    static bool IsCaption(dynamic doc, dynamic para)
    {
        try
        {
            string style = (string)para.Style.NameLocal;
            if (style.IndexOf("Caption", StringComparison.OrdinalIgnoreCase) >= 0 || style.Contains("\u9898\u6ce8")) return true;
        }
        catch (Exception) { }
        return false;
    }

    /// Where a figure (index > 0, with the empty paragraphs before it and its caption after it) or a table
    /// (index < 0, with its caption before it) begins and ends in the document.
    static bool Block(dynamic doc, int index, out int start, out int end)
    {
        start = 0; end = 0;
        if (index > 0)
        {
            dynamic para = doc.InlineShapes[index].Range.Paragraphs[1];
            try { if (Truthy(para.Range.Information[12])) return false; } catch (Exception) { }
            dynamic first = para, before = para.Previous();
            while (before != null && ((string)before.Range.Text ?? "").Trim().Length == 0 && (int)before.Range.InlineShapes.Count == 0) { first = before; before = before.Previous(); }
            dynamic last = para, after = para.Next();
            if (after != null && IsCaption(doc, after)) last = after;
            start = (int)first.Range.Start; end = (int)last.Range.End;
            Widen(doc, ref start, ref end);
            return true;
        }
        dynamic table = doc.Tables[-index];
        start = (int)table.Range.Start; end = (int)table.Range.End;
        if (start > 0)
        {
            dynamic above = doc.Range(start - 1, start - 1).Paragraphs[1];
            if (IsCaption(doc, above)) start = (int)above.Range.Start;
        }
        Widen(doc, ref start, ref end);
        return true;
    }

    /// Figures and tables that follow one another directly move as one.
    static void Absorb(dynamic doc, ref int start, ref int end)
    {
        for (int guard = 0; guard < 4; guard++)
        {
            int last = (int)doc.Content.End;
            if (end >= last - 1) return;
            dynamic next = doc.Range(end, end).Paragraphs[1];
            int to = 0;
            if (Truthy(next.Range.Information[12])) to = (int)next.Range.Tables[1].Range.End;
            else if ((int)next.Range.InlineShapes.Count > 0)
            {
                to = (int)next.Range.End;
                dynamic caption = next.Next();
                if (caption != null && IsCaption(doc, caption)) to = (int)caption.Range.End;
            }
            else if (IsCaption(doc, next))
            {
                dynamic under = next.Next();
                if (under != null && Truthy(under.Range.Information[12])) to = (int)under.Range.Tables[1].Range.End;
            }
            if (to <= end) return;
            end = to;
            Widen(doc, ref start, ref end);
        }
    }

    /// Empty paragraphs around a figure or table belong to it: they are its spacing.
    static void Widen(dynamic doc, ref int start, ref int end)
    {
        int last = (int)doc.Content.End;
        for (int guard = 0; guard < 6 && start > 0; guard++)
        {
            dynamic above = doc.Range(start - 1, start - 1).Paragraphs[1];
            if (((string)above.Range.Text ?? "").Trim().Length > 0 || (int)above.Range.InlineShapes.Count > 0 || Truthy(above.Range.Information[12])) break;
            start = (int)above.Range.Start;
        }
        for (int guard = 0; guard < 6 && end < last - 1; guard++)
        {
            dynamic below = doc.Range(end, end).Paragraphs[1];
            if (((string)below.Range.Text ?? "").Trim().Length > 0 || (int)below.Range.InlineShapes.Count > 0 || Truthy(below.Range.Information[12])) break;
            end = (int)below.Range.End;
        }
    }

    /// A figure or table that does not fit under the text before it goes to the next page and leaves the rest of
    /// the page empty. As a typesetter would, the paragraphs that follow it are brought before it until the page is full.
    static string FillGaps(dynamic doc)
    {
        int moved = 0, blocks = 0, figures = (int)doc.InlineShapes.Count, tables = (int)doc.Tables.Count;
        List<int> all = new List<int>();
        for (int i = 1; i <= figures; i++) all.Add(i);
        for (int i = 1; i <= tables; i++) all.Add(-i);
        int docEnd = 0;
        foreach (int index in all)
        {
            int here = 0;
            bool more = false;   // the paragraph just moved announces what follows it (it ends with a colon)
            for (int step = 0; step < 8; step++)
            {
                int start, end;
                try { if (!Block(doc, index, out start, out end)) break; } catch (Exception) { break; }
                docEnd = (int)doc.Content.End;
                if (start <= 0 || end >= docEnd - 1) break;
                dynamic before = doc.Range(start - 1, start - 1).Paragraphs[1];
                if (!more)
                {
                    if ((int)before.OutlineLevel != 10) break;
                    dynamic tip = doc.Range(start - 1, start - 1), head = doc.Range(start, start);
                    if ((int)head.Information[3] <= (int)tip.Information[3]) break;
                    dynamic setup = before.Range.Sections[1].PageSetup;
                    double bottom = (double)setup.PageHeight - (double)setup.BottomMargin;
                    double gap = bottom - Convert.ToDouble(tip.Information[6]) - 20;
                    Trace("block " + index + " pages " + (int)tip.Information[3] + "/" + (int)head.Information[3] + " gap " + gap);
                    if (gap < Math.Max(90, (bottom - (double)setup.TopMargin) * 0.12)) break;
                }
                Absorb(doc, ref start, ref end);
                if (end >= docEnd - 1) break;
                dynamic after = doc.Range(end, end).Paragraphs[1];
                string text = ((string)after.Range.Text ?? "").Trim();
                if (text.Length == 0 || (int)after.OutlineLevel != 10 || (int)after.Range.InlineShapes.Count > 0 || IsCaption(doc, after)) break;
                try { if (Truthy(after.Range.Information[12])) break; } catch (Exception) { }
                // A paragraph that is a list item or the last one before a heading moves like any other; what must not
                // happen is text jumping over a heading, and that is excluded above.
                doc.Range(start, start).FormattedText = after.Range.FormattedText;
                // The block moved down by what was put before it: find it, and what follows it, again.
                if (!Block(doc, index, out start, out end)) break;
                Absorb(doc, ref start, ref end);
                doc.Range(end, end).Paragraphs[1].Range.Delete();
                moved++; here++;
                more = text.EndsWith(":", StringComparison.Ordinal) || text.EndsWith("\uff1a", StringComparison.Ordinal);
            }
            if (here > 0) blocks++;
        }
        if (moved == 0) return "no page is left part-empty before a figure or table that text could be moved up to fill";
        return moved + " paragraph(s) brought before " + blocks + " figure(s) / table(s) to fill the space that was left empty at the foot of the page before them; look at those pages";
    }

    static bool Nested;   // a picture or table being put in from inside insert_paragraphs

    static string WordOp(dynamic doc, string type, Bag op)
    {
        if (type == "replace_text")
        {
            string find = op.Need("find"), replace = Lines(op.Raw("replace") ?? "");
            bool all = op.Flag("all", true), matchCase = op.Flag("matchCase", true);
            dynamic range = doc.Content;
            int count = 0;
            // "^" starts a special code in Word's find box; here the text is always meant literally.
            string literal = find.Replace("^", "^^");
            bool rich = replace.IndexOf('$') >= 0 || replace.IndexOf("\\cite", StringComparison.Ordinal) >= 0;
            string extra = "";
            while (count < 5000 && (bool)range.Find.Execute(FindText: literal, MatchCase: matchCase, MatchWildcards: false, Forward: true, Wrap: 0))
            {
                range.Text = replace;
                if (rich)
                {
                    // Building formulas changes the length: keep the place by its distance from the end.
                    int tail = (int)doc.Content.End - (int)range.End;
                    extra = Enrich(doc, range);
                    int resume = (int)doc.Content.End - tail;
                    range = doc.Range(resume, resume);
                }
                else range.Collapse(0);
                count++;
                if (!all) break;
            }
            if (count == 0) throw new Fail("NOT_FOUND", "The text \"" + Clip(find, 60) + "\" does not occur in the document.");
            return "replaced " + count + extra + Lint(replace);
        }
        if (type == "insert_paragraphs")
        {
            IList items = op.List("items");
            if (items == null) { items = new ArrayList(); items.Add(new Dictionary<string, object> { { "text", op.Need("text") } }); }
            dynamic current = null;
            int made = 0, equations = 0, cites = 0;
            string lint = "";
            Bag resume = null;
            foreach (object raw in items)
            {
                Bag item = raw is string ? new Bag(new Dictionary<string, object> { { "text", raw } }) : new Bag(raw);
                // A picture or a table written among the paragraphs goes in right there, in order.
                string nested = item.Str("op", null);
                if (nested == null && !item.Has("text")) nested = item.Has("path") ? "insert_image" : item.Has("data") ? "insert_table" : null;
                if (nested == "insert_image" || nested == "insert_table")
                {
                    Dictionary<string, object> inner = item.Copy();
                    inner.Remove("op"); inner.Remove("para"); inner.Remove("where"); inner.Remove("expect");
                    Bag place = current != null ? null : resume ?? op;
                    int before = (int)doc.Paragraphs.Count, tail;
                    if (current != null) { int at = Index(doc, current); inner["para"] = at; inner["where"] = "after"; tail = before - at; }
                    else
                    {
                        string where = place.Str("where", place.Has("para") ? "after" : "end");
                        int para = place.Int("para", 1);
                        if (place.Has("para")) inner["para"] = para;
                        inner["where"] = where;
                        if (place.Has("expect")) inner["expect"] = place.Raw("expect");
                        tail = where == "end" || (before == 1 && (string)doc.Paragraphs[1].Range.Text == "\r") ? 0 : where == "start" ? before : where == "before" ? before - para + 1 : before - para;
                    }
                    Nested = true;
                    try { WordOp(doc, nested, new Bag(inner)); } finally { Nested = false; }
                    // Where the text goes on: after what was just put in.
                    int end = (int)doc.Paragraphs.Count - tail;
                    Dictionary<string, object> next = new Dictionary<string, object>();
                    if (tail > 0)
                    {
                        bool inTable = false;
                        try { inTable = Truthy(doc.Paragraphs[end].Range.Information[12]); } catch (Exception) { }
                        next["para"] = inTable ? end + 1 : end;
                        next["where"] = inTable ? "before" : "after";
                    }
                    resume = new Bag(next);
                    current = null;
                    made++;
                    continue;
                }
                foreach (string text in Lines(item.Raw("text") ?? "").Split('\r'))
                {
                    dynamic p;
                    if (current == null) p = NewParagraph(doc, resume ?? op);
                    else p = After(doc, current);
                    string style = item.Str("style", op.Str("style", null));
                    if (style != null) Plain(p, style);
                    else Body(p);
                    WordFormat(p.Range, StyleLess(op));
                    WordFormat(p.Range, StyleLess(item));
                    Follow(doc, p.Range);
                    Write(doc, (int)p.Range.Start, text);
                    if (lint.Length == 0) lint = Lint(text);
                    Unnumber(p, text);
                    try
                    {
                        if (Unnumbered.Count > 0 && Unnumbered.Contains((string)doc.FullName + "|" + (string)p.Style.NameLocal) && (int)p.Range.ListFormat.ListType != 0) p.Range.ListFormat.RemoveNumbers();
                    }
                    catch (Exception) { }
                    Whole(p.Range);
                    cites += Cite(doc, p.Range);
                    equations += Mathify(doc, p.Range);
                    current = p;
                    made++;
                }
            }
            // The batch may end on a picture or a table: then there is no last paragraph of text to count from.
            if (current == null) return "inserted " + made + " item(s), " + (int)doc.Paragraphs.Count + " paragraphs now" + (equations > 0 ? ", " + equations + " equation(s)" : "") + (cites > 0 ? ", " + cites + " citation(s)" : "") + MathNote() + lint + SmallNote();
            int lastIndex = Index(doc, current);
            return "inserted " + made + " paragraph(s), now paragraphs " + (lastIndex - made + 1) + "–" + lastIndex + " of " + (int)doc.Paragraphs.Count + (equations > 0 ? ", " + equations + " equation(s)" : "") + (cites > 0 ? ", " + cites + " citation(s)" : "") + MathNote() + lint + SmallNote();
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
            range.Text = "";
            Write(doc, (int)p.Range.Start, Lines(op.Raw("text") ?? ""));
            Whole(p.Range);
            string built = Enrich(doc, p.Range);
            return "paragraph " + op.Int("para", 0) + " rewritten" + built + Lint(Lines(op.Raw("text") ?? ""));
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
            dynamic gone = WordRange(doc, op);
            string note = "";
            try { gone.Delete(); }
            catch (COMException)
            {
                // The mark that ends a table cell cannot be deleted: take everything before it.
                int limit = (int)gone.End - 1;
                try { int from = (int)gone.Start; limit = Math.Min(limit, (int)doc.Range(from, from).Cells[1].Range.End - 1); } catch (Exception) { }
                gone.End = limit;
                if ((int)gone.End <= (int)gone.Start) throw new Fail("BAD_ARGS", "This is the last paragraph of a table cell; a cell always keeps one. Leave it empty.");
                gone.Delete();
                note = " (the last paragraph of the cell stays, a cell always keeps one)";
            }
            return "deleted, " + (int)doc.Paragraphs.Count + " paragraphs now" + note;
        }
        if (type == "insert_table")
        {
            IList data = op.List("data");
            int rows = op.Int("rows", data == null ? 2 : data.Count), cols = op.Int("cols", 0);
            if (data != null) foreach (object row in data) { IList cells = row as IList; if (cells != null) cols = Math.Max(cols, cells.Count); }
            if (rows < 1 || cols < 1) throw new Fail("BAD_ARGS", "Give \"data\" (rows of cells) or \"rows\" and \"cols\".");
            dynamic p = NewParagraph(doc, op);
            string fontName = null, fontFarEast = null;
            float fontSize = 0;
            try
            {
                if ((int)p.OutlineLevel == 10)
                {
                    fontName = (string)p.Range.Font.Name; fontFarEast = (string)p.Range.Font.NameFarEast; fontSize = (float)p.Range.Font.Size;
                }
            }
            catch (Exception) { }
            Plain(p, "Normal");
            Follow(doc, p.Range);
            string cellLint = "";
            dynamic table = doc.Tables.Add(p.Range, rows, cols);
            // The cells take the formatting of the paragraph that follows (a heading, say): make them plain.
            SetStyle(table.Range, "Normal");
            try
            {
                table.Range.Font.Reset(); table.Range.ParagraphFormat.Reset();
                table.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0; table.Range.ParagraphFormat.FirstLineIndent = 0;
            }
            catch (COMException) { }
            try
            {
                // Same typeface and size as the body text the table sits in; cells centred like a data table.
                if (!string.IsNullOrEmpty(fontName)) table.Range.Font.Name = fontName;
                if (!string.IsNullOrEmpty(fontFarEast)) table.Range.Font.NameFarEast = fontFarEast;
                if (fontSize > 4 && fontSize < 100) table.Range.Font.Size = fontSize;
                table.Range.ParagraphFormat.Alignment = 1;
                Whole(table.Range);
                table.Range.Cells.VerticalAlignment = 1;
            }
            catch (COMException) { }
            Rules(table, op.Str("borders", "grid"));
            try { table.AutoFitBehavior(2); } catch (COMException) { }
            if (data != null)
            {
                for (int r = 0; r < data.Count && r < rows; r++)
                {
                    IList cells = data[r] as IList;
                    if (cells == null) continue;
                    for (int c = 0; c < cells.Count; c++)
                    {
                        if (cells[c] == null) continue;
                        dynamic cell = table.Cell(r + 1, c + 1).Range;
                        string value = Convert.ToString(cells[c], System.Globalization.CultureInfo.InvariantCulture);
                        cell.Text = value;
                        if (value.IndexOf('$') >= 0 || value.IndexOf("\\cite", StringComparison.Ordinal) >= 0) Enrich(doc, table.Cell(r + 1, c + 1).Range);
                        else if (cellLint.Length == 0) cellLint = Lint(value);
                    }
                    if (Typing) Thread.Sleep(25);
                }
            }
            if (op.Flag("header", true))
            {
                table.Rows[1].Range.Font.Bold = 1;
                // On a page break the header row is repeated at the top of the next page.
                try { table.Rows[1].HeadingFormat = -1; } catch (COMException) { }
            }
            try { table.Rows.AllowBreakAcrossPages = 0; } catch (COMException) { }
            if (op.Has("caption")) Caption(doc, table.Range, true, op.Need("caption"), fontName, fontFarEast);
            else
            {
                int tableStart = (int)table.Range.Start;
                if (tableStart > 0) Space(doc.Range(tableStart - 1, tableStart - 1).Paragraphs[1], false);
            }
            int afterTable = (int)table.Range.End;
            if (afterTable < (int)doc.Content.End) Space(doc.Range(afterTable, afterTable).Paragraphs[1], true);
            int tableEnd = (int)doc.Range(0, table.Range.End).Paragraphs.Count;
            bool nested = false;
            try { nested = (int)table.NestingLevel > 1; } catch (Exception) { }
            return (nested ? "table inserted inside a cell of table " + (int)doc.Range(0, table.Range.End).Tables.Count : "table " + (int)doc.Range(0, table.Range.End).Tables.Count + " inserted") + " (" + rows + "×" + cols + "), now paragraphs " + (tableEnd - (int)table.Range.Paragraphs.Count + 1) + "–" + tableEnd + " of " + (int)doc.Paragraphs.Count + cellLint;
        }
        if (type == "set_cell")
        {
            int n = op.Int("table", 1);
            if (n < 1 || n > (int)doc.Tables.Count) throw new Fail("ANCHOR_MISSING", "There is no table " + n + ".");
            dynamic cell = doc.Tables[n].Cell(op.Int("row", 1), op.Int("col", 1));
            Follow(doc, cell.Range);
            cell.Range.Text = "";
            Write(doc, (int)cell.Range.Start, Lines(op.Raw("text") ?? ""));
            string built = Enrich(doc, doc.Tables[n].Cell(op.Int("row", 1), op.Int("col", 1)).Range);
            return "cell set" + built + Lint(Lines(op.Raw("text") ?? ""));
        }
        if (type == "insert_image")
        {
            string path = op.Need("path");
            if (!File.Exists(path)) throw new Fail("BAD_ARGS", "Image \"" + path + "\" does not exist.");
            dynamic p = NewParagraph(doc, op);
            string bodyFont = null, bodyFarEast = null;
            try { if ((int)p.OutlineLevel == 10) { bodyFont = (string)p.Range.Font.Name; bodyFarEast = (string)p.Range.Font.NameFarEast; } } catch (Exception) { }
            Plain(p, "Normal");
            Follow(doc, p.Range);
            dynamic picture = p.Range.InlineShapes.AddPicture(FileName: path, LinkToFile: false, SaveWithDocument: true);
            Cramped = true;
            float drawn = 0;
            try { drawn = (float)picture.Width * 100f / Math.Max(1f, (float)picture.ScaleWidth); } catch (Exception) { }
            if (op.Has("width")) { picture.LockAspectRatio = -1; picture.Width = (float)op.Points("width", 300); }
            try { Lettering(path, (float)picture.Width, drawn, false); } catch (Exception) { }
            try { p.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0; p.Range.ParagraphFormat.FirstLineIndent = 0; p.Range.ParagraphFormat.Alignment = 1; } catch (COMException) { }
            WordFormat(p.Range, StyleLess(op));
            p.Range.ParagraphFormat.SpaceBefore = Air;
            if (op.Has("caption")) Caption(doc, picture.Range, false, op.Need("caption"), bodyFont, bodyFarEast);
            else p.Range.ParagraphFormat.SpaceAfter = Air;
            return "image inserted" + (Nested ? "" : SmallNote());
        }
        if (type == "insert_references") return References(doc, op);
        if (type == "style_format") return StyleFormat(doc, op);
        if (type == "page_setup") return PageSetup(doc, op);
        if (type == "page_numbers") return PageNumbers(doc, op);
        if (type == "insert_toc") return Toc(doc, op);
        if (type == "page_break")
        {
            dynamic p = Para(doc, op.Int("para", 0), op.Str("expect", null));
            p.Range.ParagraphFormat.PageBreakBefore = op.Flag("on", true) ? -1 : 0;
            return "paragraph " + op.Int("para", 0) + " now starts a new page";
        }
        if (type == "header")
        {
            dynamic header = doc.Sections[1].Headers[1].Range;
            header.Text = Lines(op.Raw("text") ?? "");
            header.ParagraphFormat.Alignment = 1;
            if (op.Has("size")) header.Font.Size = (float)op.Num("size", 9);
            return "page header set";
        }
        if (type == "fill_gaps") return FillGaps(doc);
        if (type == "update_fields") { Refresh(doc); return "table of contents and cross-references refreshed"; }
        if (type == "format_table")
        {
            int n = op.Int("table", 1);
            if (n < 1 || n > (int)doc.Tables.Count) throw new Fail("NOT_FOUND", "There is no table " + n + " (the document has " + (int)doc.Tables.Count + ").");
            dynamic table = doc.Tables[n];
            Follow(doc, table.Range);
            if (op.Has("borders")) Rules(table, op.Need("borders"));
            if (op.Has("font")) { string font = op.Need("font"); table.Range.Font.Name = font; try { table.Range.Font.NameFarEast = font; } catch (COMException) { } }
            if (op.Has("latinFont")) { string latin = op.Need("latinFont"); try { table.Range.Font.NameAscii = latin; table.Range.Font.NameOther = latin; } catch (COMException) { } }
            if (op.Has("size")) table.Range.Font.Size = (float)op.Num("size", 10.5);
            if (op.Has("align")) { string align = op.Need("align"); table.Range.ParagraphFormat.Alignment = align == "center" ? 1 : align == "right" ? 2 : 0; }
            if (op.Has("header")) table.Rows[1].Range.Font.Bold = op.Flag("header", true) ? 1 : 0;
            if (op.Has("rowHeight")) { table.Rows.HeightRule = 1; table.Rows.Height = (float)op.Points("rowHeight", 20); }
            if (op.Has("lineSpacing")) LineSpacing(table.Range.ParagraphFormat, op.Raw("lineSpacing"));
            if (op.Has("autofit")) { string fit = op.Need("autofit"); try { table.AutoFitBehavior(fit == "content" ? 1 : fit == "fixed" ? 0 : 2); } catch (COMException) { } }
            if (op.Has("keepTogether") && op.On("keepTogether"))
            {
                // Every row but the last stays with the next one, so the table is not split across two pages.
                int rows = (int)table.Rows.Count;
                for (int r = 1; r < rows; r++) { try { table.Rows[r].Range.ParagraphFormat.KeepWithNext = -1; } catch (COMException) { } }
            }
            return "table " + n + " formatted";
        }
        if (type == "set_image")
        {
            int total = (int)doc.InlineShapes.Count;
            dynamic picture = null;
            if (op.Has("para"))
            {
                dynamic range = Para(doc, op.Int("para", 1), null).Range;
                if ((int)range.InlineShapes.Count == 0) throw new Fail("NOT_FOUND", "Paragraph " + op.Int("para", 1) + " holds no picture. Pictures show as \"/\" in office_read.");
                picture = range.InlineShapes[1];
            }
            else
            {
                int n = op.Int("image", 1);
                if (n < 1 || n > total) throw new Fail("NOT_FOUND", "There is no picture " + n + " (the document has " + total + ").");
                picture = doc.InlineShapes[n];
            }
            Follow(doc, picture.Range);
            picture.LockAspectRatio = -1;
            if (op.Has("width")) picture.Width = (float)op.Points("width", 300);
            if (op.Has("height")) picture.Height = (float)op.Points("height", 200);
            if (op.Has("align")) { string align = op.Need("align"); picture.Range.ParagraphFormat.Alignment = align == "center" ? 1 : align == "right" ? 2 : 0; }
            return "picture resized to " + Math.Round((double)picture.Width / 28.3465, 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " × " + Math.Round((double)picture.Height / 28.3465, 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " cm";
        }
        throw new Fail("BAD_ARGS", "Unknown Word operation \"" + type + "\".");
    }

    /// The lines of a table: "grid" (all of them), "none", or "three-line" (三线表: a heavy rule above and below the
    /// table, a light one under the header row, nothing else), the way tables are set in papers.
    static void Rules(dynamic table, string style)
    {
        string kind = style.Trim().ToLowerInvariant();
        if (kind == "grid" || kind == "all" || kind == "true") { table.Borders.Enable = 1; return; }
        table.Borders.Enable = 0;
        if (kind == "none" || kind == "false") return;
        if (kind != "three-line" && kind != "threeline" && kind != "three_line" && kind != "booktabs" && kind != "三线表")
            throw new Fail("BAD_ARGS", "\"borders\" must be grid, three-line or none.");
        foreach (int side in new int[] { -1, -3 })
        {
            dynamic rule = table.Borders[side];
            rule.LineStyle = 1;
            rule.LineWidth = 12;
        }
        if ((int)table.Rows.Count > 1)
        {
            dynamic under = table.Rows[1].Borders[-3];
            under.LineStyle = 1;
            under.LineWidth = 6;
        }
    }

    static Bag StyleLess(Bag source)
    {
        Dictionary<string, object> copy = new Dictionary<string, object>();
        foreach (string key in new string[] { "font", "latinFont", "size", "bold", "italic", "underline", "color", "align", "firstLineIndent", "spaceBefore", "spaceAfter", "lineSpacing", "indentChars", "superscript", "subscript" })
            if (source.Has(key)) copy[key] = source.Raw(key);
        return new Bag(copy);
    }


    // ───────────────────────── citations ─────────────────────────

    static readonly System.Text.RegularExpressions.Regex Cites = new System.Text.RegularExpressions.Regex(@"\\cite\s*\{([0-9,\-–\s]+)\}");

    /// One citation number at a position, as a superscript that jumps to the reference: a cross-reference field
    /// when the reference list is already there, otherwise a link to the bookmark the list will bring.
    static void CiteNumber(dynamic doc, int at, string number)
    {
        string mark = "cite_" + number;
        dynamic spot = doc.Range(at, at);
        if ((bool)doc.Bookmarks.Exists(mark))
        {
            // MERGEFORMAT keeps the superscript when the field is refreshed.
            dynamic field = doc.Fields.Add(spot, -1, "REF " + mark + " \\h \\* MERGEFORMAT", false);
            field.Result.Font.Superscript = 1;
        }
        else
        {
            dynamic link = doc.Hyperlinks.Add(Anchor: spot, Address: "", SubAddress: mark, TextToDisplay: number);
            link.Range.Font.Superscript = 1;
            link.Range.Font.Underline = 0;
            link.Range.Font.Color = -16777216;
        }
    }

    static readonly System.Text.RegularExpressions.Regex Typed = new System.Text.RegularExpressions.Regex(@"^\[[1-9]\d{0,2}([,，\-–][1-9]\d{0,2})*\]\z");

    /// Citations typed by hand in a range: "[2]" or "[1,3]" in running text is rewritten as \cite{..}, and "[2] ..."
    /// opening a paragraph is an entry of a hand-made reference list, whose number gets the bookmark citations jump to.
    static void Handwritten(dynamic doc, dynamic range)
    {
        bool listed = false;
        int from = (int)range.Start;
        for (int guard = 0; guard < 300; guard++)
        {
            int end = (int)range.End;
            if (from >= end) break;
            dynamic hit = doc.Range(from, end);
            if (!(bool)hit.Find.Execute(FindText: "\\[[0-9,，–\\-]@\\]", MatchWildcards: true, Forward: true, Wrap: 0)) break;
            if ((int)hit.End > end || (int)hit.Start < from) break;
            from = (int)hit.End;
            string found = (string)hit.Text;
            if (found == null || !Typed.IsMatch(found)) continue;
            try { if ((int)hit.Font.Superscript != 0 || (int)hit.Fields.Count > 0) continue; } catch (Exception) { continue; }
            dynamic paragraph = hit.Paragraphs[1].Range;
            int head = (int)paragraph.Start;
            if ((int)hit.Start == head)
            {
                string number = found.Substring(1, found.Length - 2);
                int n;
                if (!int.TryParse(number, out n) || (int)paragraph.End - (int)hit.End < 4) continue;
                string mark = "cite_" + n;
                if ((bool)doc.Bookmarks.Exists(mark)) doc.Bookmarks[mark].Delete();
                doc.Bookmarks.Add(mark, doc.Range(head + 1, head + 1 + number.Length));
                listed = true;
                Listed = true;
                continue;
            }
            // Inside a formula still written between dollar signs ($[1,2]$ is an interval) it is not a citation.
            string before = (string)doc.Range(head, hit.Start).Text ?? "";
            int dollars = 0;
            foreach (char ch in before) if (ch == '$') dollars++;
            if (dollars % 2 == 1) continue;
            string rewritten = "\\cite{" + found.Substring(1, found.Length - 2).Replace("，", ",").Replace("–", "-") + "}";
            hit.Text = rewritten;
            from = (int)hit.Start + rewritten.Length;
        }
        if (listed) Relink(doc);
    }

    /// Set when reference entries were written in this batch: the text is then checked for citations of them.
    static bool Listed;

    /// A reference list nobody cites: said once, at the end of the batch that wrote the list.
    static string Uncited(dynamic doc)
    {
        try
        {
            int entries = 0;
            foreach (dynamic mark in doc.Bookmarks) { if (((string)mark.Name).StartsWith("cite_", StringComparison.Ordinal)) entries++; }
            if (entries == 0) return "";
            foreach (dynamic field in doc.Fields) { if (((string)field.Code.Text).IndexOf("REF cite_", StringComparison.Ordinal) >= 0) return ""; }
            foreach (dynamic link in doc.Hyperlinks)
            {
                string target = "";
                try { target = (string)link.SubAddress; } catch (Exception) { }
                if (target != null && target.StartsWith("cite_", StringComparison.Ordinal)) return "";
            }
            return "(NOTE the reference list has " + entries + " entries but the text cites none of them: put \\cite{n} — or [n] — where each source is used, e.g. with replace_text)";
        }
        catch (COMException) { return ""; }
    }

    /// **text** written in a range becomes bold text (a label such as "摘要：" in front of a paragraph).
    static void Bold(dynamic doc, dynamic range)
    {
        string text = (string)range.Text;
        if (text == null || text.IndexOf("**", StringComparison.Ordinal) < 0) return;
        int from = (int)range.Start;
        for (int guard = 0; guard < 100; guard++)
        {
            int end = (int)range.End;
            if (from >= end) break;
            dynamic hit = doc.Range(from, end);
            if (!(bool)hit.Find.Execute(FindText: "\\*\\*[!\\*^13]@\\*\\*", MatchWildcards: true, Forward: true, Wrap: 0)) break;
            if ((int)hit.End > end || (int)hit.Start < from) break;
            string found = (string)hit.Text;
            string inner = found.Substring(2, found.Length - 4);
            hit.Text = inner;
            hit.Font.Bold = 1;
            from = (int)hit.Start + inner.Length;
        }
    }

    /// Citations written before their reference existed are links; once the targets exist, make them cross-references.
    static int Relink(dynamic doc)
    {
        int linked = 0;
        for (int i = (int)doc.Hyperlinks.Count; i >= 1; i--)
        {
            dynamic link = doc.Hyperlinks[i];
            string target = "";
            try { target = (string)link.SubAddress; } catch (Exception) { }
            if (target == null || !target.StartsWith("cite_", StringComparison.Ordinal) || !(bool)doc.Bookmarks.Exists(target)) continue;
            dynamic spot = link.Range;
            int at = (int)spot.Start;
            link.Delete();
            dynamic text = doc.Range(at, at + target.Length - 5);
            text.Text = "";
            dynamic field = doc.Fields.Add(doc.Range(at, at), -1, "REF " + target + " \\h \\* MERGEFORMAT", false);
            field.Result.Font.Superscript = 1;
            linked++;
        }
        return linked;
    }

    /// \cite{1}, \cite{2,5}, \cite{3-6} in a range become superscript [1], [2,5], [3-6] whose numbers jump to the references.
    static int Cite(dynamic doc, dynamic range)
    {
        string text = (string)range.Text;
        if (text == null) return 0;
        try { Bold(doc, range); } catch (COMException) { }
        if (text.IndexOf('[') >= 0)
        {
            try { Handwritten(doc, range); } catch (COMException) { }
            text = (string)range.Text ?? "";
        }
        if (text.IndexOf("\\cite", StringComparison.Ordinal) < 0) return 0;
        int made = 0;
        System.Text.RegularExpressions.MatchCollection found = Cites.Matches(text);
        for (int k = found.Count - 1; k >= 0; k--)
        {
            System.Text.RegularExpressions.Match m = found[k];
            dynamic search = range.Duplicate;
            if (!(bool)search.Find.Execute(FindText: m.Value, MatchCase: true, MatchWildcards: false, Forward: true, Wrap: 0)) continue;
            int at = (int)search.Start;
            search.Text = "";
            // Written back to front at one position, so each piece lands before the one after it.
            List<string> pieces = new List<string>();
            pieces.Add("[");
            foreach (string part in m.Groups[1].Value.Replace("–", "-").Replace(" ", "").Split(','))
            {
                if (part.Length == 0) continue;
                if (pieces.Count > 1) pieces.Add(",");
                string[] ends = part.Split('-');
                pieces.Add("#" + ends[0]);
                if (ends.Length > 1 && ends[1].Length > 0) { pieces.Add("-"); pieces.Add("#" + ends[1]); }
            }
            pieces.Add("]");
            for (int p = pieces.Count - 1; p >= 0; p--)
            {
                string piece = pieces[p];
                if (piece[0] == '#') { CiteNumber(doc, at, piece.Substring(1)); continue; }
                doc.Range(at, at).InsertAfter(piece);
                dynamic plain = doc.Range(at, at + piece.Length);
                plain.Font.Superscript = 1;
            }
            made++;
        }
        return made;
    }

    /// Formulas and citations written in the text of a range.
    static string Enrich(dynamic doc, dynamic range)
    {
        int cites = Cite(doc, range), equations = Mathify(doc, range);
        return (equations > 0 ? ", " + equations + " equation(s)" : "") + (cites > 0 ? ", " + cites + " citation(s)" : "") + MathNote();
    }

    /// The reference list: "[n] ..." paragraphs whose numbers carry the bookmarks the citations point at.
    static string References(dynamic doc, Bag op)
    {
        IList items = op.List("items");
        if (items == null || items.Count == 0) throw new Fail("BAD_ARGS", "\"items\" must list the references, one string each.");
        System.Text.RegularExpressions.Regex lead = new System.Text.RegularExpressions.Regex(@"^\s*\[(\d+)\]\s*");
        dynamic current = null;
        int number = 0;
        foreach (object raw in items)
        {
            string body = Convert.ToString(raw).Trim();
            System.Text.RegularExpressions.Match m = lead.Match(body);
            if (m.Success) { number = int.Parse(m.Groups[1].Value); body = body.Substring(m.Length); }
            else number++;
            dynamic p = current == null ? NewParagraph(doc, op) : After(doc, current);
            Plain(p, "Normal");
            Follow(doc, p.Range);
            string label = "[" + number + "] ";
            Write(doc, (int)p.Range.Start, label + body);
            Whole(p.Range);
            try
            {
                dynamic format = p.Range.ParagraphFormat;
                format.CharacterUnitFirstLineIndent = 0; format.FirstLineIndent = 0; format.Alignment = 3;
                format.CharacterUnitLeftIndent = 0; format.LeftIndent = 21f; format.FirstLineIndent = -21f;
            }
            catch (COMException) { }
            int start = (int)p.Range.Start;
            string mark = "cite_" + number;
            if ((bool)doc.Bookmarks.Exists(mark)) doc.Bookmarks[mark].Delete();
            doc.Bookmarks.Add(mark, doc.Range(start + 1, start + 1 + number.ToString().Length));
            current = p;
        }
        int linked = Relink(doc);
        return items.Count + " reference(s) listed" + (linked > 0 ? ", " + linked + " citation(s) in the text now cross-reference them" : "");
    }

    // ───────────────────────── styles and page layout ─────────────────────────

    /// Styles the agent switched automatic numbering off for (document path | style name).
    static readonly HashSet<string> Unnumbered = new HashSet<string>();

    static dynamic StyleOf(dynamic doc, string name)
    {
        int builtin;
        try { return WordStyles.TryGetValue(name, out builtin) ? doc.Styles[builtin] : doc.Styles[name]; }
        catch (COMException) { throw new Fail("STYLE_MISSING", "This document has no style \"" + name + "\". Built-in names that always work: " + string.Join(", ", new List<string>(WordStyles.Keys).ToArray()) + "."); }
    }

    /// Change what a style looks like, for every paragraph that uses it.
    static string StyleFormat(dynamic doc, Bag op)
    {
        string name = op.Need("style");
        dynamic style = StyleOf(doc, name);
        dynamic font = style.Font, format = style.ParagraphFormat;
        if (op.Has("font")) { string f = op.Need("font"); font.Name = f; try { font.NameFarEast = f; } catch (COMException) { } }
        if (op.Has("latinFont")) { string f = op.Need("latinFont"); try { font.NameAscii = f; font.NameOther = f; } catch (COMException) { } }
        if (op.Has("size")) font.Size = (float)op.Num("size", 12);
        if (op.Has("bold")) font.Bold = op.Flag("bold", false) ? 1 : 0;
        if (op.Has("italic")) font.Italic = op.Flag("italic", false) ? 1 : 0;
        if (op.Has("color")) font.Color = Bgr(op.Need("color"));
        if (op.Has("align")) { string align = op.Need("align"); format.Alignment = align == "center" ? 1 : align == "right" ? 2 : align == "justify" ? 3 : 0; }
        if (op.Has("indentChars")) { format.FirstLineIndent = 0; format.CharacterUnitFirstLineIndent = (float)op.Num("indentChars", 2); }
        if (op.Has("firstLineIndent")) Indent(format, op.Raw("firstLineIndent"));
        if (op.Has("spaceBefore")) format.SpaceBefore = (float)op.Num("spaceBefore", 0);
        if (op.Has("spaceAfter")) format.SpaceAfter = (float)op.Num("spaceAfter", 0);
        if (op.Has("lineSpacing")) LineSpacing(format, op.Raw("lineSpacing"));
        if (op.Has("pageBreakBefore")) format.PageBreakBefore = op.Flag("pageBreakBefore", false) ? -1 : 0;
        if (op.Has("numbering") && !op.On("numbering"))
        {
            Unnumbered.Add((string)doc.FullName + "|" + (string)style.NameLocal);
            // Stop the style from numbering its paragraphs by itself (headings that carry their number in the text).
            try { style.LinkToListTemplate(null); } catch (Exception) { }
            foreach (dynamic p in doc.Paragraphs)
            {
                try { if ((string)p.Style.NameLocal == (string)style.NameLocal && (int)p.Range.ListFormat.ListType != 0) p.Range.ListFormat.RemoveNumbers(); }
                catch (Exception) { }
            }
        }
        try { format.WordWrap = -1; } catch (Exception) { }
        return "style \"" + (string)style.NameLocal + "\" changed";
    }

    static string PageSetup(dynamic doc, Bag op)
    {
        dynamic setup = doc.PageSetup;
        const float cm = 28.3465f;
        if (op.Has("paper")) { string paper = op.Need("paper").ToUpperInvariant(); setup.PaperSize = paper == "A3" ? 6 : paper == "LETTER" ? 2 : paper == "B5" ? 13 : 7; }
        if (op.Has("orientation")) setup.Orientation = op.Need("orientation") == "landscape" ? 1 : 0;
        if (op.Has("top")) setup.TopMargin = (float)op.Num("top", 2.54) * cm;
        if (op.Has("bottom")) setup.BottomMargin = (float)op.Num("bottom", 2.54) * cm;
        if (op.Has("left")) setup.LeftMargin = (float)op.Num("left", 3.17) * cm;
        if (op.Has("right")) setup.RightMargin = (float)op.Num("right", 3.17) * cm;
        return "page setup changed";
    }

    static string PageNumbers(dynamic doc, Bag op)
    {
        string align = op.Str("align", "center");
        dynamic numbers = doc.Sections[1].Footers[1].PageNumbers;
        numbers.Add(align == "left" ? 0 : align == "right" ? 2 : 1, op.Flag("firstPage", true));
        if (op.Has("start")) { numbers.RestartNumberingAtSection = true; numbers.StartingNumber = op.Int("start", 1); }
        return "page numbers added to the footer";
    }

    static string Toc(dynamic doc, Bag op)
    {
        dynamic p = NewParagraph(doc, op);
        Plain(p, "Normal");
        if (op.Has("title"))
        {
            p.Range.InsertBefore(op.Need("title"));
            p.Range.ParagraphFormat.Alignment = 1; p.Range.ParagraphFormat.CharacterUnitFirstLineIndent = 0; p.Range.ParagraphFormat.FirstLineIndent = 0;
            p.Range.Font.Bold = 1; p.Range.Font.Size = 16f;
            p = After(doc, p);
            Plain(p, "Normal");
        }
        Follow(doc, p.Range);
        doc.TablesOfContents.Add(Range: p.Range, UseHeadingStyles: true, UpperHeadingLevel: 1, LowerHeadingLevel: Math.Max(1, Math.Min(9, op.Int("levels", 3))));
        return "table of contents inserted (it is refreshed on every save)";
    }

    /// Bring fields up to date: the table of contents, cross-references, caption numbers.
    static void Refresh(dynamic doc)
    {
        try { foreach (dynamic toc in doc.TablesOfContents) toc.Update(); } catch (Exception) { }
        try { doc.Fields.Update(); } catch (Exception) { }
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
        int bang = address.LastIndexOf('!');
        if (bang > 0)
        {
            // "汇总!B2:D10": the sheet named there, not the one the operation is on.
            string name = address.Substring(0, bang).Trim('\'', '=');
            dynamic other;
            try { other = sheet.Parent.Worksheets[name]; }
            catch (COMException)
            {
                List<string> names = new List<string>();
                foreach (dynamic one in sheet.Parent.Worksheets) names.Add((string)one.Name);
                throw new Fail("ANCHOR_MISSING", "There is no sheet \"" + name + "\" (in \"" + address + "\"). Sheets: " + string.Join(", ", names.ToArray()) + ".");
            }
            sheet = other;
            address = address.Substring(bang + 1);
        }
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
        bool[] dated = new bool[cols + 1];
        if (v != null) for (int c = 1; c <= cols; c++) { try { dated[c] = DateColumn(range.Columns[c], v, c); } catch (Exception) { } }
        for (int r = 1; r <= rows; r++)
        {
            List<object> line = new List<object>();
            for (int c = 1; c <= cols; c++)
            {
                object one = CellValue(v == null ? values : v[r, c]);
                if (one is double) one = double.Parse(((double)one).ToString("G12", System.Globalization.CultureInfo.InvariantCulture), System.Globalization.CultureInfo.InvariantCulture);
                if (dated[c] && IsNumber(one))
                {
                    try
                    {
                        DateTime day = DateTime.FromOADate(Convert.ToDouble(one));
                        one = day.TimeOfDay.TotalSeconds < 1 ? day.ToString("yyyy-MM-dd") : day.ToString("yyyy-MM-dd HH:mm");
                    }
                    catch (Exception) { }
                }
                line.Add(one);
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
        List<object> objects = Objects(sheet);
        if (objects.Count > 0) result["objects"] = objects;
        // A table too long to show whole is described column by column instead; so is any block when asked.
        if (a.Flag("profile", result.ContainsKey("clipped") && !a.Has("range")))
        {
            try
            {
                dynamic whole = a.Has("range") ? Cells(sheet, a.Need("range")) : sheet.UsedRange;
                if ((long)whole.Rows.Count * (long)whole.Columns.Count <= 3000000) result["profile"] = Profile(whole);
            }
            catch (Fail) { throw; }
            catch (Exception) { }
        }
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
        if (op.On("border")) range.Borders.LineStyle = 1;
        if (op.On("merge")) range.Merge();
        if (op.Has("columnWidth")) range.ColumnWidth = op.Num("columnWidth", 10);
        if (op.Has("rowHeight")) range.RowHeight = op.Num("rowHeight", 15);
    }

    /// Bring the cells being worked on into view, without touching the selection.
    static void Show(dynamic book, dynamic sheet, dynamic range)
    {
        if (!Following) return;
        try
        {
            sheet.Activate();
            dynamic window = book.Windows[1];
            dynamic visible = window.VisibleRange;
            int row = (int)range.Row, column = (int)range.Column;
            int top = (int)visible.Row, left = (int)visible.Column;
            if (row < top || row >= top + (int)visible.Rows.Count - 1) window.ScrollRow = Math.Max(1, row - 3);
            if (column < left || column >= left + (int)visible.Columns.Count - 1) window.ScrollColumn = Math.Max(1, column - 1);
        }
        catch (Exception) { }
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
        string worked = ExcelData(app, book, ws, type, op);
        if (worked != null) return worked;
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
            Show(book, ws, target);
            if (Typing && rows.Count > 1 && rows.Count <= 60)
            {
                for (int r = 0; r < rows.Count; r++)
                {
                    object[,] line = new object[1, width];
                    for (int c = 0; c < width; c++) line[0, c] = grid[r, c];
                    target.Rows[r + 1].Formula = line;
                    Thread.Sleep(35);
                }
            }
            else target.Formula = grid;
            ExcelFormat(target, op);
            return "wrote " + (string)target.Address[false, false] + Errors(target);
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
        throw new Fail("BAD_ARGS", "Unknown Excel operation \"" + type + "\".");
    }

    // ───────────────────────── Excel: data work ─────────────────────────

    /// The block of data an operation works on: the given range, or everything that is filled on the sheet.
    static dynamic Data(dynamic ws, Bag op)
    {
        return op.Has("range") ? Cells(ws, op.Need("range")) : ws.UsedRange;
    }

    /// A range that may name its sheet ("产品表!A1:C9"); without one it is on the sheet at hand.
    static dynamic Ref(dynamic book, dynamic ws, string address)
    {
        int bang = address.LastIndexOf('!');
        if (bang < 0) return Cells(ws, address);
        string sheet = address.Substring(0, bang).Trim('\'', '=');
        dynamic other;
        try { other = book.Worksheets[sheet]; }
        catch (COMException) { throw new Fail("ANCHOR_MISSING", "There is no sheet \"" + sheet + "\"."); }
        return Cells(other, address.Substring(bang + 1));
    }

    static string Text(object value)
    {
        return value == null ? "" : Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
    }

    /// The first row of a block, as texts.
    static List<string> Headers(dynamic data)
    {
        List<string> heads = new List<string>();
        int cols = (int)data.Columns.Count;
        object raw = data.Rows[1].Value2;
        object[,] row = raw as object[,];
        for (int c = 1; c <= cols; c++) heads.Add(Text(row == null ? raw : row[1, c]).Trim());
        return heads;
    }

    /// A column of a block by its header text, its letter ("C") or its position in the block; the position is returned.
    static int Col(dynamic data, object key)
    {
        int cols = (int)data.Columns.Count;
        List<string> heads = Headers(data);
        string name = Text(key).Trim();
        for (int c = 0; c < heads.Count; c++) if (string.Equals(heads[c], name, StringComparison.OrdinalIgnoreCase)) return c + 1;
        if (key is int || key is double || key is decimal || key is long)
        {
            int at = Convert.ToInt32(key);
            if (at >= 1 && at <= cols) return at;
        }
        if (name.Length > 0 && name.Length <= 3 && System.Text.RegularExpressions.Regex.IsMatch(name, "^[A-Za-z]+\\z"))
        {
            int letter = 0;
            foreach (char ch in name.ToUpperInvariant()) letter = letter * 26 + (ch - 'A' + 1);
            int at = letter - (int)data.Column + 1;
            if (at >= 1 && at <= cols) return at;
        }
        throw new Fail("ANCHOR_MISSING", "There is no column \"" + name + "\" in " + (string)data.Address[false, false] + ". Its headers: " + string.Join(", ", heads.ToArray()) + ".");
    }

    /// The columns an operation names ("columns": [..] or "column": one), as positions in the block; all of them when none is named.
    static List<int> Cols(dynamic data, Bag op, bool allWhenNone)
    {
        List<int> list = new List<int>();
        IList many = op.List("columns");
        if (many != null) foreach (object key in many) list.Add(Col(data, key));
        else if (op.Has("column")) list.Add(Col(data, op.Raw("column")));
        else if (allWhenNone) for (int c = 1; c <= (int)data.Columns.Count; c++) list.Add(c);
        return list;
    }

    static object[,] Grid(dynamic range)
    {
        object raw = range.Value2;
        object[,] grid = raw as object[,];
        if (grid != null) return grid;
        grid = (object[,])Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 });
        grid[1, 1] = raw;
        return grid;
    }

    static bool IsNumber(object value) { return value is double || value is int || value is decimal || value is float || value is long; }

    static bool IsBlank(object value) { return value == null || (value is string && ((string)value).Trim().Length == 0); }

    static string Tidy(string text)
    {
        string s = text.Replace(' ', ' ').Replace('　', ' ').Replace("\t", " ").Trim();
        while (s.Contains("  ")) s = s.Replace("  ", " ");
        return s;
    }

    /// A number out of what a person typed: "¥1,299", "89元", "10%", " 12 件".
    static bool ParseNumber(string text, out double number)
    {
        number = 0;
        string s = Tidy(text).Replace(",", "").Replace("，", "").Replace(" ", "");
        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(s, "^[¥￥$€£]?\\s*([-+]?\\d*\\.?\\d+(?:[eE][-+]?\\d+)?)\\s*(%|％)?\\s*(?:[^\\d\\s\\-.,A-Za-z%％]{1,3}|kg|g|km|cm|mm|m|pcs|k|w)?\\z");
        if (!m.Success) return false;
        if (!double.TryParse(m.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out number)) return false;
        if (m.Groups[2].Success) number /= 100;
        return true;
    }

    /// A date out of what a person typed: 2025/3/1, 2025.03.01, 20250301, 2025年3月1日, 3月1日 (the year given).
    static bool ParseDate(string text, int year, out DateTime date)
    {
        date = DateTime.MinValue;
        string s = Tidy(text);
        System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(s, "^(\\d{4})(\\d{2})(\\d{2})\\z");
        if (!m.Success) m = System.Text.RegularExpressions.Regex.Match(s, "^(\\d{4})\\s*[-/.年]\\s*(\\d{1,2})\\s*[-/.月]\\s*(\\d{1,2})\\s*日?(?:\\s.*)?\\z");
        int y = 0, mo = 0, d = 0;
        if (m.Success) { y = int.Parse(m.Groups[1].Value); mo = int.Parse(m.Groups[2].Value); d = int.Parse(m.Groups[3].Value); }
        else
        {
            m = System.Text.RegularExpressions.Regex.Match(s, "^(\\d{1,2})\\s*[-/.月]\\s*(\\d{1,2})\\s*日?\\z");
            if (m.Success) { y = year; mo = int.Parse(m.Groups[1].Value); d = int.Parse(m.Groups[2].Value); }
        }
        if (m.Success)
        {
            if (y < 1900 || y > 2200 || mo < 1 || mo > 12 || d < 1 || d > DateTime.DaysInMonth(y, mo)) return false;
            date = new DateTime(y, mo, d);
            return true;
        }
        return DateTime.TryParse(s, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out date) && date.Year >= 1900;
    }

    /// Whether the numbers of a column are shown as dates.
    static bool DateColumn(dynamic column, object[,] grid, int c)
    {
        int rows = grid.GetLength(0), seen = 0, dates = 0;
        for (int r = 2; r <= rows && seen < 12; r++)
        {
            if (!IsNumber(grid[r, c])) continue;
            seen++;
            try
            {
                string format = Convert.ToString(column.Cells[r, 1].NumberFormat);
                if (System.Text.RegularExpressions.Regex.IsMatch(format, "[ymd年月日]", System.Text.RegularExpressions.RegexOptions.IgnoreCase) && !format.Contains("0.0") && format != "General") dates++;
            }
            catch (Exception) { }
        }
        return seen > 0 && dates * 2 > seen;
    }

    static string Plain(double number)
    {
        if (Math.Abs(number - Math.Round(number)) < 1e-9 && Math.Abs(number) < 1e15) return Math.Round(number).ToString("0", System.Globalization.CultureInfo.InvariantCulture);
        return number.ToString("0.####", System.Globalization.CultureInfo.InvariantCulture);
    }

    /// What is in each column of a block: kinds of values, range, the commonest texts, and what looks wrong. This is
    /// how a long table is understood without reading every row.
    static List<object> Profile(dynamic data)
    {
        List<object> lines = new List<object>();
        object[,] grid = Grid(data);
        int rows = grid.GetLength(0), cols = grid.GetLength(1), row0 = (int)data.Row, col0 = (int)data.Column;
        if (rows < 2) return lines;
        int emptyRows = 0;
        Dictionary<string, int> whole = new Dictionary<string, int>();
        int sameRows = 0;
        for (int r = 2; r <= rows; r++)
        {
            StringBuilder key = new StringBuilder();
            bool any = false;
            for (int c = 1; c <= cols; c++) { if (!IsBlank(grid[r, c])) any = true; key.Append(Text(grid[r, c])).Append('\u0001'); }
            if (!any) { emptyRows++; continue; }
            string k = key.ToString();
            if (whole.ContainsKey(k)) sameRows++; else whole[k] = 1;
        }
        lines.Add((rows - 1) + " data row(s) under the header row " + row0 + (emptyRows > 0 ? "; " + emptyRows + " entirely empty row(s)" : "") + (sameRows > 0 ? "; " + sameRows + " row(s) repeat an earlier row exactly" : "") + ".");
        for (int c = 1; c <= cols; c++)
        {
            int blank = 0, numbers = 0, texts = 0, errors = 0, spaced = 0, numberLike = 0, dateLike = 0, negative = 0;
            double min = double.MaxValue, max = double.MinValue, sum = 0;
            Dictionary<string, int> seen = new Dictionary<string, int>();
            List<double> all = new List<double>();
            for (int r = 2; r <= rows; r++)
            {
                object v = grid[r, c];
                if (IsBlank(v)) { blank++; continue; }
                if (v is int && CellValue(v) is string) { errors++; continue; }
                string shown = Text(v);
                if (IsNumber(v))
                {
                    double n = Convert.ToDouble(v);
                    numbers++; sum += n; all.Add(n);
                    if (n < min) min = n;
                    if (n > max) max = n;
                    if (n < 0) negative++;
                }
                else
                {
                    texts++;
                    string s = (string)(v as string ?? shown);
                    if (s != Tidy(s)) spaced++;
                    double n; DateTime d;
                    bool dateShaped = System.Text.RegularExpressions.Regex.IsMatch(Tidy(s), "^(\\d{8}|\\d{4}\\s*[-/.年]\\s*\\d{1,2}\\s*[-/.月]\\s*\\d{1,2}.*|\\d{1,2}\\s*月\\s*\\d{1,2}\\s*日?)\\z");
                    if (dateShaped && ParseDate(s, DateTime.Now.Year, out d)) dateLike++;
                    else if (ParseNumber(s, out n)) numberLike++;
                }
                int count;
                seen[shown] = seen.TryGetValue(shown, out count) ? count + 1 : 1;
            }
            bool dates = numbers > 0 && DateColumn(data.Columns[c], grid, c);
            StringBuilder line = new StringBuilder();
            string head = Text(grid[1, c]).Trim();
            line.Append(Column(col0 + c - 1)).Append(' ').Append(head.Length == 0 ? "(no header)" : "\"" + head + "\"").Append(": ");
            List<string> kinds = new List<string>();
            if (numbers > 0) kinds.Add(numbers + (dates ? " date(s)" : " number(s)"));
            if (texts > 0) kinds.Add(texts + " text");
            if (blank > 0) kinds.Add(blank + " blank");
            if (errors > 0) kinds.Add(errors + " error(s)");
            line.Append(kinds.Count == 0 ? "empty" : string.Join(", ", kinds.ToArray()));
            if (numbers > 0)
            {
                if (dates) line.Append("; from " + DateTime.FromOADate(min).ToString("yyyy-MM-dd") + " to " + DateTime.FromOADate(max).ToString("yyyy-MM-dd"));
                else
                {
                    all.Sort();
                    double median = all.Count % 2 == 1 ? all[all.Count / 2] : (all[all.Count / 2 - 1] + all[all.Count / 2]) / 2;
                    line.Append("; min " + Plain(min) + ", median " + Plain(median) + ", mean " + Plain(Math.Round(sum / numbers, 2)) + ", max " + Plain(max) + ", sum " + Plain(Math.Round(sum, 2)));
                    if (negative > 0) line.Append("; " + negative + " negative");
                    // Values far beyond the bulk of the column: more than three times the spread above the upper quartile.
                    if (all.Count >= 20)
                    {
                        double q1 = all[all.Count / 4], q3 = all[all.Count * 3 / 4], fence = q3 + 3 * Math.Max(q3 - q1, 1e-9);
                        int far = 0;
                        foreach (double n in all) if (n > fence) far++;
                        if (far > 0 && far <= all.Count / 20) line.Append("; " + far + " far above the rest (over " + Plain(Math.Round(fence, 2)) + ")");
                    }
                }
            }
            int distinct = seen.Count, filled = numbers + texts;
            if (dates && texts == 0) { }
            else if (dates) line.Append("; the texts are written in other ways, e.g. " + Sample(grid, c));
            else if (texts > 0 || (numbers > 0 && distinct <= 12))
            {
                line.Append("; " + distinct + " distinct");
                if (distinct <= 40 || distinct * 2 < filled)
                {
                    List<KeyValuePair<string, int>> top = new List<KeyValuePair<string, int>>(seen);
                    top.Sort(delegate(KeyValuePair<string, int> a, KeyValuePair<string, int> b) { return b.Value != a.Value ? b.Value.CompareTo(a.Value) : string.CompareOrdinal(a.Key, b.Key); });
                    List<string> parts = new List<string>();
                    int limit = distinct <= 24 ? 24 : 10;
                    for (int i = 0; i < top.Count && i < limit; i++) parts.Add("\"" + Clip(top[i].Key, 24) + "\" " + top[i].Value);
                    line.Append(": " + string.Join(", ", parts.ToArray()) + (top.Count > limit ? ", .." : ""));
                }
                else if (filled > distinct) line.Append(" (" + (filled - distinct) + " repeat an earlier value)");
            }
            List<string> odd = new List<string>();
            if (spaced > 0) odd.Add(spaced + " with stray spaces");
            if (numberLike > 0) odd.Add(numberLike + " number(s) written as text");
            if (dateLike > 0) odd.Add(dateLike + " date(s) written as text");
            if (numbers > 0 && texts > 0) odd.Add("mixed numbers and text");
            if (odd.Count > 0) line.Append(". LOOK: " + string.Join(", ", odd.ToArray()));
            lines.Add(line.ToString());
        }
        return lines;
    }

    /// A few of the texts of a column, one of each shape (digits read as 0), to show how they are written.
    static string Sample(object[,] grid, int c)
    {
        Dictionary<string, string> shapes = new Dictionary<string, string>();
        for (int r = 2; r <= grid.GetLength(0) && shapes.Count < 5; r++)
        {
            string s = grid[r, c] as string;
            if (s == null || s.Trim().Length == 0) continue;
            string shape = System.Text.RegularExpressions.Regex.Replace(s, "\\d", "0");
            if (!shapes.ContainsKey(shape)) shapes[shape] = s;
        }
        List<string> parts = new List<string>();
        foreach (string one in shapes.Values) parts.Add("\"" + Clip(one, 20) + "\"");
        return string.Join(", ", parts.ToArray());
    }

    /// Charts, pivot tables and tables of a sheet, so that they can be named in later operations.
    static List<object> Objects(dynamic ws)
    {
        List<object> lines = new List<object>();
        try
        {
            foreach (dynamic holder in ws.ChartObjects())
            {
                string title = "";
                try { if (Truthy(holder.Chart.HasTitle)) title = " \"" + (string)holder.Chart.ChartTitle.Text + "\""; } catch (Exception) { }
                string place = "";
                try { place = " at " + (string)holder.TopLeftCell.Address[false, false] + ":" + (string)holder.BottomRightCell.Address[false, false]; } catch (Exception) { }
                int series = 0;
                try { series = (int)holder.Chart.SeriesCollection().Count; } catch (Exception) { }
                lines.Add("chart \"" + (string)holder.Name + "\"" + title + place + ", " + series + " series");
            }
        }
        catch (Exception) { }
        try { foreach (dynamic pivot in ws.PivotTables()) lines.Add("pivot table \"" + (string)pivot.Name + "\" at " + (string)pivot.TableRange1.Address[false, false]); } catch (Exception) { }
        try { foreach (dynamic table in ws.ListObjects) lines.Add("table \"" + (string)table.Name + "\" at " + (string)table.Range.Address[false, false]); } catch (Exception) { }
        try { if (Truthy(ws.AutoFilterMode)) lines.Add("a filter is on: some rows may be hidden (clear_filter shows them all)"); } catch (Exception) { }
        return lines;
    }

    /// Delete whole rows of a sheet, given by their numbers; from the bottom up, several at a time.
    static void DeleteRows(dynamic ws, List<int> rows)
    {
        rows.Sort();
        for (int i = rows.Count - 1; i >= 0; )
        {
            StringBuilder address = new StringBuilder();
            int taken = 0;
            while (i >= 0 && taken < 18)
            {
                // Runs of neighbouring rows go as one piece.
                int last = rows[i], first = last;
                while (i > 0 && rows[i - 1] == first - 1) { i--; first = rows[i]; }
                i--;
                if (address.Length > 0) address.Append(',');
                address.Append(first).Append(':').Append(last);
                taken++;
            }
            ws.Range[address.ToString()].Delete();
        }
    }

    static string RowList(List<int> rows)
    {
        List<string> parts = new List<string>();
        for (int i = 0; i < rows.Count && i < 12; i++) parts.Add(rows[i].ToString());
        return string.Join(", ", parts.ToArray()) + (rows.Count > 12 ? ", .." : "");
    }

    /// Errors among the results of formulas just written: said at once, so that they are not found later.
    static string Errors(dynamic range)
    {
        try
        {
            object[,] grid = Grid(range);
            int bad = 0;
            string first = null, kind = null;
            int row0 = (int)range.Row, col0 = (int)range.Column;
            for (int r = 1; r <= grid.GetLength(0); r++)
            {
                for (int c = 1; c <= grid.GetLength(1); c++)
                {
                    object shown = CellValue(grid[r, c]);
                    if (!(grid[r, c] is int) || !(shown is string)) continue;
                    bad++;
                    if (first == null) { first = Column(col0 + c - 1) + (row0 + r - 1); kind = (string)shown; }
                }
            }
            return bad == 0 ? "" : " — " + bad + " cell(s) show an error (first: " + first + " " + kind + "): fix the formula or the data it reads";
        }
        catch (Exception) { return ""; }
    }

    static int ChartKind(string kind)
    {
        switch (kind)
        {
            case "bar": return 57;
            case "stacked": case "stacked column": return 52;
            case "stacked bar": return 58;
            case "line": return 65;
            case "smooth": return 4;
            case "pie": return 5;
            case "doughnut": case "donut": return -4120;
            case "scatter": return -4169;
            case "scatter line": return 74;
            case "area": return 1;
            case "radar": return -4151;
            case "column": case "combo": return 51;
        }
        throw new Fail("BAD_ARGS", "Unknown chart \"" + kind + "\": use column, bar, stacked, stacked bar, line, pie, doughnut, scatter, scatter line, area, radar or combo.");
    }

    /// Fill in or change a chart from the fields of an operation.
    static string Chart(dynamic book, dynamic ws, dynamic holder, Bag op, bool fresh)
    {
        dynamic chart = holder.Chart;
        string kind = fresh ? op.Str("type", op.Str("chart", "column")) : op.Str("type", null);
        IList series = op.List("series");
        if (series != null || op.Has("range"))
        {
            // A chart made while the selection is inside data arrives with series of Excel's own guessing.
            try { while ((int)chart.SeriesCollection().Count > 0) chart.SeriesCollection(1).Delete(); } catch (Exception) { }
        }
        if (series != null)
        {
            dynamic categories = op.Has("categories") ? Ref(book, ws, op.Need("categories")) : null;
            foreach (object raw in series)
            {
                Bag item = new Bag(raw);
                dynamic one = chart.SeriesCollection().NewSeries();
                one.Values = Ref(book, ws, item.Need("values"));
                if (item.Has("x")) one.XValues = Ref(book, ws, item.Need("x"));
                else if (categories != null) one.XValues = categories;
                if (item.Has("name"))
                {
                    string name = item.Need("name");
                    // A cell reference names the series after that cell; anything else is the name itself.
                    if (System.Text.RegularExpressions.Regex.IsMatch(name, "^(?:[^!]+!)?\\$?[A-Za-z]{1,3}\\$?\\d+\\z")) one.Name = "=" + (string)Ref(book, ws, name).Address[true, true, 1, true];
                    else one.Name = name;
                }
            }
        }
        else if (op.Has("range")) chart.SetSourceData(Ref(book, ws, op.Need("range")));
        if (kind != null) chart.ChartType = ChartKind(kind);
        if (series != null)
        {
            int index = 0;
            foreach (object raw in series)
            {
                index++;
                Bag item = new Bag(raw);
                dynamic one = chart.SeriesCollection(index);
                if (item.Has("chart")) one.ChartType = ChartKind(item.Need("chart"));
                else if (kind == "combo" && index == series.Count && series.Count > 1) one.ChartType = 65;
                if (item.Str("axis", "") == "secondary" || item.Str("axis", "") == "right") one.AxisGroup = 2;
                if (item.Has("color")) { try { one.Format.Fill.ForeColor.RGB = Bgr(item.Need("color")); one.Format.Line.ForeColor.RGB = Bgr(item.Need("color")); } catch (Exception) { } }
            }
        }
        if (op.Has("title")) { chart.HasTitle = true; chart.ChartTitle.Text = op.Need("title"); try { chart.ChartTitle.Font.Size = 13; chart.ChartTitle.Font.Bold = true; } catch (Exception) { } }
        else if (fresh) { try { chart.HasTitle = false; } catch (Exception) { } }
        bool round = kind == "pie" || kind == "doughnut" || kind == "donut" || kind == "radar";
        if (op.Has("xTitle") && !round) { try { chart.Axes(1).HasTitle = true; chart.Axes(1).AxisTitle.Text = op.Need("xTitle"); } catch (Exception) { } }
        if (op.Has("yTitle") && !round) { try { chart.Axes(2).HasTitle = true; chart.Axes(2).AxisTitle.Text = op.Need("yTitle"); } catch (Exception) { } }
        if (op.Has("y2Title")) { try { chart.Axes(2, 2).HasTitle = true; chart.Axes(2, 2).AxisTitle.Text = op.Need("y2Title"); } catch (Exception) { } }
        if (op.Has("numberFormat")) { try { chart.Axes(2).TickLabels.NumberFormat = op.Need("numberFormat"); } catch (Exception) { } }
        if (op.Has("min")) { try { chart.Axes(2).MinimumScale = op.Num("min", 0); } catch (Exception) { } }
        if (op.Has("max")) { try { chart.Axes(2).MaximumScale = op.Num("max", 0); } catch (Exception) { } }
        if (op.Has("labels") || (fresh && kind != null && (kind == "pie" || kind == "doughnut" || kind == "donut")))
        {
            bool on = !op.Has("labels") || op.On("labels") || op.Str("labels", "") == "percent" || op.Str("labels", "") == "value";
            try
            {
                foreach (dynamic one in chart.SeriesCollection())
                {
                    one.HasDataLabels = on;
                    if (!on) continue;
                    string what = op.Str("labels", "");
                    if (what == "percent" || (round && what != "value")) { one.DataLabels().ShowPercentage = true; one.DataLabels().ShowValue = false; one.DataLabels().ShowCategoryName = kind != null && kind != "radar"; }
                    try { one.DataLabels().Font.Size = 9; } catch (Exception) { }
                }
            }
            catch (Exception) { }
        }
        if (op.Has("legend") || fresh)
        {
            string legend = op.Str("legend", null);
            int count = 0;
            try { count = (int)chart.SeriesCollection().Count; } catch (Exception) { }
            if (legend == null) legend = round ? "right" : count > 1 ? "bottom" : "none";
            try
            {
                chart.HasLegend = legend != "none" && legend != "false";
                if (Truthy(chart.HasLegend)) chart.Legend.Position = legend == "right" ? -4152 : legend == "top" ? -4160 : legend == "left" ? -4131 : -4107;
            }
            catch (Exception) { }
        }
        if (fresh)
        {
            // A plain, readable look: one typeface, pale grid lines, bars that are not thin.
            try { chart.ChartArea.Font.Name = "Microsoft YaHei"; chart.ChartArea.Font.Size = 10; } catch (Exception) { }
            try { chart.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = Bgr("#E4E6EA"); } catch (Exception) { }
            try { if (kind == "column" || kind == "bar" || kind == "combo" || kind == "stacked" || kind == "stacked bar") chart.ChartGroups(1).GapWidth = 80; } catch (Exception) { }
            try { chart.ChartArea.Format.Line.Visible = 0; } catch (Exception) { }
        }
        if (op.Has("at")) { dynamic anchor = Cells(ws, op.Need("at")); holder.Left = (double)anchor.Left; holder.Top = (double)anchor.Top; }
        if (op.Has("width")) holder.Width = op.Num("width", 420);
        if (op.Has("height")) holder.Height = op.Num("height", 260);
        if (op.Has("name")) holder.Name = op.Need("name");
        int total = 0, points = 0;
        try { total = (int)chart.SeriesCollection().Count; if (total > 0) points = ((Array)chart.SeriesCollection(1).Values).Length; } catch (Exception) { }
        string where = "";
        try { where = " at " + (string)holder.TopLeftCell.Address[false, false] + ":" + (string)holder.BottomRightCell.Address[false, false]; } catch (Exception) { }
        // Data taken from a pivot table makes a pivot chart: it shows every value of that table and carries field buttons.
        try
        {
            if (chart.PivotLayout != null)
            {
                try { chart.ShowAllFieldButtons = false; } catch (Exception) { }
                where += " (its data is a pivot table, so it shows all values of that table and follows its filters; to chart only some of them, write them to plain cells first)";
            }
        }
        catch (Exception) { }
        // Charts lying over one another is the commonest fault of a dashboard: say so at once.
        try
        {
            double l = (double)holder.Left, t = (double)holder.Top, r = l + (double)holder.Width, b = t + (double)holder.Height;
            foreach (dynamic other in ws.ChartObjects())
            {
                if ((string)other.Name == (string)holder.Name) continue;
                double ol = (double)other.Left, ot = (double)other.Top, or = ol + (double)other.Width, ob = ot + (double)other.Height;
                if (l < or - 2 && ol < r - 2 && t < ob - 2 && ot < b - 2) { where += " — it OVERLAPS chart \"" + (string)other.Name + "\" (" + (string)other.TopLeftCell.Address[false, false] + ":" + (string)other.BottomRightCell.Address[false, false] + "): move one with set_chart {chart, at}"; break; }
            }
        }
        catch (Exception) { }
        return "chart \"" + (string)holder.Name + "\" " + (fresh ? "added" : "changed") + " on \"" + (string)ws.Name + "\"" + where + ": " + total + " series" + (points > 0 ? " × " + points + " point(s)" : "") + (total == 0 ? " — NOTHING is plotted: check the range" : "");
    }

    static dynamic ChartByName(dynamic ws, string name)
    {
        List<string> names = new List<string>();
        foreach (dynamic holder in ws.ChartObjects())
        {
            if (string.Equals((string)holder.Name, name, StringComparison.OrdinalIgnoreCase)) return holder;
            names.Add("\"" + (string)holder.Name + "\"");
        }
        int index;
        if (int.TryParse(name, out index) && index >= 1 && index <= names.Count) return ws.ChartObjects(index);
        throw new Fail("ANCHOR_MISSING", "Sheet \"" + (string)ws.Name + "\" has no chart \"" + name + "\". Charts: " + (names.Count == 0 ? "none" : string.Join(", ", names.ToArray())) + ".");
    }

    static bool Test(object value, string how, object against, IList among)
    {
        string text = Tidy(Text(value));
        switch (how)
        {
            case "blank": return IsBlank(value);
            case "notblank": case "not blank": return !IsBlank(value);
            case "contains": return text.IndexOf(Text(against), StringComparison.OrdinalIgnoreCase) >= 0;
            case "in":
                if (among != null) foreach (object one in among) if (string.Equals(text, Tidy(Text(one)), StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            case "text": return value is string && !IsBlank(value);
            case "error": return value is int && CellValue(value) is string;
        }
        if (IsBlank(value)) return false;
        double a = 0, b = 0;
        bool left = true, right = true;
        if (IsNumber(value)) a = Convert.ToDouble(value); else left = ParseNumber(text, out a);
        if (IsNumber(against)) b = Convert.ToDouble(against); else right = ParseNumber(Text(against), out b);
        bool numbers = left && right;
        int order = numbers ? a.CompareTo(b) : string.Compare(text, Tidy(Text(against)), StringComparison.OrdinalIgnoreCase);
        switch (how)
        {
            case "=": case "==": case "equals": return order == 0;
            case "<>": case "!=": return order != 0;
            case ">": return numbers && order > 0;
            case "<": return numbers && order < 0;
            case ">=": return numbers && order >= 0;
            case "<=": return numbers && order <= 0;
        }
        throw new Fail("BAD_ARGS", "Unknown test \"" + how + "\": use blank, notblank, =, <>, >, <, >=, <=, contains, in, text or error.");
    }

    /// The operations for working with data. Returns null for an operation that is not one of them.
    static string ExcelData(dynamic app, dynamic book, dynamic ws, string type, Bag op)
    {
        if (type == "sort")
        {
            dynamic data = Data(ws, op);
            IList by = op.List("by");
            if (by == null) { by = new ArrayList(); by.Add(new Dictionary<string, object> { { "column", op.Raw("column") }, { "order", op.Str("order", "asc") } }); }
            dynamic sort = ws.Sort;
            sort.SortFields.Clear();
            List<string> said = new List<string>();
            foreach (object raw in by)
            {
                Bag key = raw is string ? new Bag(new Dictionary<string, object> { { "column", raw } }) : new Bag(raw);
                int c = Col(data, key.Raw("column"));
                bool down = key.Str("order", "asc").StartsWith("d", StringComparison.OrdinalIgnoreCase);
                dynamic column = data.Columns[c];
                sort.SortFields.Add(Key: column, SortOn: 0, Order: down ? 2 : 1, DataOption: 0);
                said.Add(Headers(data)[c - 1] + (down ? " ↓" : " ↑"));
            }
            sort.SetRange(data);
            sort.Header = op.Flag("header", true) ? 1 : 2;
            sort.Apply();
            return "sorted " + (string)data.Address[false, false] + " by " + string.Join(", ", said.ToArray());
        }
        if (type == "filter")
        {
            dynamic data = Data(ws, op);
            int c = Col(data, op.Raw("column"));
            IList values = op.List("values");
            if (values != null)
            {
                string[] wanted = new string[values.Count];
                for (int i = 0; i < values.Count; i++) wanted[i] = Text(values[i]);
                data.AutoFilter(Field: c, Criteria1: wanted, Operator: 7);
            }
            else if (op.Has("criteria2")) data.AutoFilter(Field: c, Criteria1: op.Need("criteria"), Operator: op.Str("join", "and") == "or" ? 2 : 1, Criteria2: op.Need("criteria2"));
            else data.AutoFilter(Field: c, Criteria1: op.Need("criteria"));
            int shown = 0;
            try { shown = (int)data.Columns[1].SpecialCells(12).Cells.Count - 1; } catch (Exception) { }
            return "filter on \"" + Headers(data)[c - 1] + "\": " + shown + " of " + ((int)data.Rows.Count - 1) + " row(s) shown (the others are hidden, not deleted; clear_filter shows all)";
        }
        if (type == "clear_filter")
        {
            try { if (Truthy(ws.FilterMode)) ws.ShowAllData(); } catch (Exception) { }
            if (!op.Flag("keepArrows", false)) { try { ws.AutoFilterMode = false; } catch (Exception) { } }
            return "all rows shown";
        }
        if (type == "remove_duplicates")
        {
            dynamic data = Data(ws, op);
            object[,] grid = Grid(data);
            List<int> keys = Cols(data, op, true);
            Dictionary<string, int> seen = new Dictionary<string, int>();
            List<int> gone = new List<int>();
            int row0 = (int)data.Row;
            for (int r = 2; r <= grid.GetLength(0); r++)
            {
                StringBuilder key = new StringBuilder();
                bool any = false;
                foreach (int c in keys) { if (!IsBlank(grid[r, c])) any = true; key.Append(Tidy(Text(grid[r, c])).ToUpperInvariant()).Append('\u0001'); }
                if (!any) continue;
                string k = key.ToString();
                if (seen.ContainsKey(k)) gone.Add(row0 + r - 1); else seen[k] = r;
            }
            string what = keys.Count == grid.GetLength(1) ? "whole rows" : "the same " + string.Join(" + ", keys.ConvertAll<string>(delegate(int c) { return Headers(data)[c - 1]; }).ToArray());
            if (gone.Count == 0) return "no duplicates (" + what + ")";
            string list = RowList(gone);
            DeleteRows(ws, gone);
            return "removed " + gone.Count + " duplicate row(s) (" + what + "; the first of each kept): rows " + list + ". Row numbers below them have changed";
        }
        if (type == "delete_rows" && (op.Has("where") || op.Has("blank")))
        {
            dynamic data = Data(ws, op);
            object[,] grid = Grid(data);
            int row0 = (int)data.Row, cols = grid.GetLength(1);
            List<int> gone = new List<int>();
            string said;
            if (op.Has("where"))
            {
                Bag where = new Bag(op.Raw("where"));
                int c = Col(data, where.Raw("column"));
                string how = where.Str("is", where.Str("op", where.Has("value") ? "=" : "blank"));
                object against = where.Raw("value");
                IList among = where.List("values");
                if (among != null && !where.Has("is") && !where.Has("op")) how = "in";
                for (int r = 2; r <= grid.GetLength(0); r++) if (Test(grid[r, c], how, against, among)) gone.Add(row0 + r - 1);
                said = "\"" + Headers(data)[c - 1] + "\" " + how + (against != null ? " " + Text(against) : among != null ? " [" + among.Count + " values]" : "");
            }
            else
            {
                for (int r = 2; r <= grid.GetLength(0); r++)
                {
                    bool any = false;
                    for (int c = 1; c <= cols; c++) if (!IsBlank(grid[r, c])) { any = true; break; }
                    if (!any) gone.Add(row0 + r - 1);
                }
                said = "entirely empty";
            }
            if (gone.Count == 0) return "no row is " + said + ": nothing deleted";
            string list = RowList(gone);
            DeleteRows(ws, gone);
            return "deleted " + gone.Count + " row(s) where " + said + ": rows " + list + ". Row numbers below them have changed";
        }
        if (type == "replace")
        {
            dynamic data = Data(ws, op);
            List<int> only = Cols(data, op, false);
            dynamic target = only.Count == 1 ? data.Columns[only[0]] : data;
            string find = op.Need("find"), with = op.Str("replace", op.Str("with", ""));
            bool whole = op.Flag("whole", false), matchCase = op.Flag("matchCase", false);
            int count = 0;
            object[,] grid = Grid(target);
            foreach (object v in grid)
            {
                string s = v as string;
                if (s == null) continue;
                if (whole ? string.Equals(s, find, matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) : s.IndexOf(find, matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase) >= 0) count++;
            }
            if (count == 0) return "\"" + find + "\" does not occur in " + (string)target.Address[false, false] + ": nothing replaced";
            target.Replace(What: find.Replace("~", "~~").Replace("*", "~*").Replace("?", "~?"), Replacement: with, LookAt: whole ? 1 : 2, MatchCase: matchCase);
            return "replaced in " + count + " cell(s) of " + (string)target.Address[false, false];
        }
        if (type == "clean")
        {
            dynamic data = Data(ws, op);
            List<int> columns = Cols(data, op, !op.Has("to") && !op.Has("map") && !op.Has("case") && !op.Has("remove"));
            if (columns.Count == 0) throw new Fail("BAD_ARGS", "clean needs \"column\" or \"columns\" when it converts, maps, changes case or removes text.");
            string to = op.Str("to", null), casing = op.Str("case", null);
            bool trim = op.Flag("trim", true), noSpaces = op.Flag("noSpaces", false);
            IList remove = op.List("remove");
            if (remove == null && op.Has("remove")) { remove = new ArrayList(); remove.Add(op.Need("remove")); }
            Dictionary<string, object> map = op.Raw("map") as Dictionary<string, object>;
            Dictionary<string, object> lookup = null;
            if (map != null)
            {
                lookup = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (KeyValuePair<string, object> pair in map) lookup[Tidy(pair.Key)] = pair.Value;
            }
            int year = op.Int("year", 0);
            object[,] grid = Grid(data);
            int rows = grid.GetLength(0), row0 = (int)data.Row, col0 = (int)data.Column;
            List<string> said = new List<string>();
            foreach (int c in columns)
            {
                dynamic column = data.Columns[c];
                bool formulas = true;
                try { object has = column.HasFormula; formulas = !(has is bool) || (bool)has; } catch (Exception) { }
                if (to == "date" && year == 0)
                {
                    // Dates given without a year take the year the rest of the column is in.
                    Dictionary<int, int> years = new Dictionary<int, int>();
                    for (int r = 2; r <= rows; r++)
                    {
                        DateTime known = DateTime.MinValue;
                        object v = grid[r, c];
                        if (IsNumber(v)) { try { known = DateTime.FromOADate(Convert.ToDouble(v)); } catch (Exception) { } }
                        else if (v is string && !System.Text.RegularExpressions.Regex.IsMatch(Tidy((string)v), "^\\d{1,2}\\s*[-/.月]")) ParseDate((string)v, 1, out known);
                        if (known.Year > 1900) { int n; years[known.Year] = years.TryGetValue(known.Year, out n) ? n + 1 : 1; }
                    }
                    int best = 0;
                    foreach (KeyValuePair<int, int> pair in years) if (pair.Value > best) { best = pair.Value; year = pair.Key; }
                    if (year == 0) year = DateTime.Now.Year;
                }
                int changed = 0, failed = 0;
                List<string> samples = new List<string>();
                object[,] fresh = (object[,])Array.CreateInstance(typeof(object), new int[] { rows, 1 }, new int[] { 1, 1 });
                fresh[1, 1] = grid[1, c];
                for (int r = 2; r <= rows; r++)
                {
                    object before = grid[r, c], after = before;
                    string s = before as string;
                    if (s != null)
                    {
                        if (trim) s = Tidy(s);
                        if (noSpaces) s = s.Replace(" ", "");
                        if (remove != null) foreach (object piece in remove) { string cut = Text(piece); if (cut.Length > 0) s = s.Replace(cut, ""); }
                        if (trim) s = s.Trim();
                        if (casing == "upper") s = s.ToUpperInvariant();
                        else if (casing == "lower") s = s.ToLowerInvariant();
                        else if (casing == "proper") s = System.Globalization.CultureInfo.InvariantCulture.TextInfo.ToTitleCase(s.ToLowerInvariant());
                        after = s.Length == 0 ? null : (object)s;
                        object mapped;
                        if (lookup != null && lookup.TryGetValue(s, out mapped)) { after = Scalar(mapped); s = after as string; }
                        if (s != null && s.Length > 0)
                        {
                            double number; DateTime date;
                            if (to == "number") { if (ParseNumber(s, out number)) after = number; else { failed++; if (samples.Count < 5) samples.Add(Column(col0 + c - 1) + (row0 + r - 1) + " \"" + Clip(s, 16) + "\""); } }
                            else if (to == "date") { if (ParseDate(s, year, out date)) after = date.ToOADate(); else { failed++; if (samples.Count < 5) samples.Add(Column(col0 + c - 1) + (row0 + r - 1) + " \"" + Clip(s, 16) + "\""); } }
                        }
                    }
                    else if (before != null && to == "text") after = Text(before);
                    else if (to == "date" && IsNumber(before) && Convert.ToDouble(before) >= 19000101 && Convert.ToDouble(before) <= 22001231)
                    {
                        DateTime date;
                        if (ParseDate(Convert.ToInt64(Convert.ToDouble(before)).ToString(), year, out date)) after = date.ToOADate();
                        else { failed++; if (samples.Count < 5) samples.Add(Column(col0 + c - 1) + (row0 + r - 1) + " " + Text(before)); }
                    }
                    else if (before != null && lookup != null) { object mapped; if (lookup.TryGetValue(Text(before), out mapped)) after = Scalar(mapped); }
                    fresh[r, 1] = after;
                    if (!object.Equals(before, after)) changed++;
                }
                if (changed > 0)
                {
                    if (!formulas)
                    {
                        if (to == "text") column.NumberFormat = "@";
                        column.Value2 = fresh;
                    }
                    else
                    {
                        // Formulas stand in this column: only the plain cells that changed are written.
                        for (int r = 2; r <= rows; r++)
                        {
                            if (object.Equals(grid[r, c], fresh[r, 1])) continue;
                            dynamic cell = column.Cells[r, 1];
                            if (Truthy(cell.HasFormula)) continue;
                            cell.Value2 = fresh[r, 1];
                        }
                    }
                }
                if (to == "date") { try { data.Columns[c].Offset[1, 0].Resize[Math.Max(1, rows - 1), 1].NumberFormat = op.Str("numberFormat", "yyyy-mm-dd"); } catch (Exception) { } }
                else if (to == "number" && op.Has("numberFormat")) { try { data.Columns[c].Offset[1, 0].Resize[Math.Max(1, rows - 1), 1].NumberFormat = op.Need("numberFormat"); } catch (Exception) { } }
                // What the column holds now, when it is a handful of values: the check that the cleaning worked.
                Dictionary<string, int> seen = new Dictionary<string, int>();
                bool few = true;
                for (int r = 2; r <= rows && few; r++)
                {
                    if (IsBlank(fresh[r, 1])) continue;
                    string shown = Text(fresh[r, 1]);
                    int n;
                    seen[shown] = seen.TryGetValue(shown, out n) ? n + 1 : 1;
                    if (seen.Count > 16) few = false;
                }
                string now = "";
                if (few && seen.Count > 0 && to != "number" && to != "date")
                {
                    List<string> parts = new List<string>();
                    foreach (KeyValuePair<string, int> pair in seen) parts.Add("\"" + Clip(pair.Key, 20) + "\" " + pair.Value);
                    now = "; values now: " + string.Join(", ", parts.ToArray());
                }
                said.Add("\"" + Text(grid[1, c]).Trim() + "\": " + changed + " cell(s) changed" + (failed > 0 ? ", " + failed + " could NOT be read as a " + to + " and were left (" + string.Join(", ", samples.ToArray()) + (failed > samples.Count ? ", .." : "") + ")" : "") + now);
            }
            return "cleaned " + string.Join("; ", said.ToArray());
        }
        if (type == "fill_blanks")
        {
            dynamic data = Data(ws, op);
            List<int> columns = Cols(data, op, false);
            if (columns.Count == 0) throw new Fail("BAD_ARGS", "fill_blanks needs \"column\" or \"columns\".");
            object with = op.Raw("with");
            if (with == null) throw new Fail("BAD_ARGS", "fill_blanks needs \"with\": a value, or \"above\", \"mean\", \"median\", \"mode\".");
            object[,] grid = Grid(data);
            int rows = grid.GetLength(0);
            List<string> said = new List<string>();
            foreach (int c in columns)
            {
                object fill = Scalar(with);
                string how = with as string;
                if (how == "mean" || how == "median" || how == "mode")
                {
                    List<double> numbers = new List<double>();
                    Dictionary<string, int> seen = new Dictionary<string, int>();
                    object commonest = null;
                    int most = 0;
                    for (int r = 2; r <= rows; r++)
                    {
                        if (IsBlank(grid[r, c])) continue;
                        if (IsNumber(grid[r, c])) numbers.Add(Convert.ToDouble(grid[r, c]));
                        int n;
                        string key = Text(grid[r, c]);
                        seen[key] = seen.TryGetValue(key, out n) ? n + 1 : 1;
                        if (seen[key] > most) { most = seen[key]; commonest = grid[r, c]; }
                    }
                    if (how == "mode") fill = commonest;
                    else
                    {
                        if (numbers.Count == 0) throw new Fail("BAD_ARGS", "Column \"" + Text(grid[1, c]) + "\" has no numbers to take the " + how + " of.");
                        numbers.Sort();
                        double sum = 0;
                        foreach (double n in numbers) sum += n;
                        fill = how == "mean" ? Math.Round(sum / numbers.Count, 4) : numbers.Count % 2 == 1 ? numbers[numbers.Count / 2] : (numbers[numbers.Count / 2 - 1] + numbers[numbers.Count / 2]) / 2;
                    }
                }
                int filled = 0;
                object last = null;
                dynamic column = data.Columns[c];
                for (int r = 2; r <= rows; r++)
                {
                    if (!IsBlank(grid[r, c])) { last = grid[r, c]; continue; }
                    object value = how == "above" ? last : fill;
                    if (value == null) continue;
                    column.Cells[r, 1].Value2 = value;
                    filled++;
                }
                said.Add("\"" + Text(grid[1, c]).Trim() + "\": " + filled + " blank(s) filled" + (how == "above" ? " from the cell above" : " with " + Text(fill)));
            }
            return string.Join("; ", said.ToArray());
        }
        if (type == "fill_formula")
        {
            dynamic target = Cells(ws, op.Need("range"));
            string formula = op.Need("formula");
            if (!formula.StartsWith("=", StringComparison.Ordinal)) formula = "=" + formula;
            Show(book, ws, target);
            target.Formula = formula;
            ExcelFormat(target, op);
            return "formula filled into " + (string)target.Address[false, false] + " (written for its first cell, adjusted for the others)" + Errors(target);
        }
        if (type == "add_column")
        {
            dynamic data = Data(ws, op);
            int rows = (int)data.Rows.Count, cols = (int)data.Columns.Count;
            string header = op.Need("header");
            // A column of that name is filled again, not added twice.
            int at = cols + 1;
            List<string> heads = Headers(data);
            for (int c = 0; c < heads.Count; c++) if (string.Equals(heads[c], header, StringComparison.OrdinalIgnoreCase)) at = c + 1;
            if (op.Has("after") && at == cols + 1)
            {
                at = Col(data, op.Raw("after")) + 1;
                data.Columns[at].EntireColumn.Insert();
                data = data.Resize[rows, cols + 1];
            }
            dynamic head = data.Cells[1, at];
            head.Value2 = header;
            try { head.Font.Bold = data.Cells[1, 1].Font.Bold; } catch (Exception) { }
            string where = Column((int)head.Column);
            if (rows < 2) return "column " + where + " \"" + header + "\" added (no data rows to fill)";
            dynamic body = data.Cells[2, at].Resize[rows - 1, 1];
            Show(book, ws, body);
            if (op.Has("formula"))
            {
                string formula = op.Need("formula");
                if (!formula.StartsWith("=", StringComparison.Ordinal)) formula = "=" + formula;
                body.Formula = formula;
            }
            else if (op.Has("value")) body.Value2 = Scalar(op.Raw("value"));
            ExcelFormat(body, op);
            if (op.Flag("values", false)) body.Value2 = body.Value2;
            return "column " + where + " \"" + header + "\" " + (at <= cols && !op.Has("after") ? "filled" : "added") + ": " + (string)body.Address[false, false] + Errors(body);
        }
        if (type == "insert_columns" || type == "delete_columns")
        {
            string at = op.Need("column");
            int count = Math.Max(1, op.Int("count", 1));
            dynamic first;
            try { first = ws.Range[at + "1"]; } catch (COMException) { throw new Fail("BAD_ARGS", "\"column\" must be a column letter, e.g. \"C\"."); }
            dynamic columns = first.Resize[1, count].EntireColumn;
            if (type == "insert_columns") columns.Insert(); else columns.Delete();
            return (type == "insert_columns" ? "inserted " : "deleted ") + count + " column(s) at " + at;
        }
        if (type == "copy_range")
        {
            dynamic source = Ref(book, ws, op.Need("range"));
            dynamic target = Ref(book, ws, op.Need("to"));
            if (op.Flag("values", false))
            {
                dynamic into = target.Cells[1, 1].Resize[(int)source.Rows.Count, (int)source.Columns.Count];
                // Writing the values over would make Excel read every text again ("20250301" would turn into a
                // number): its own copy keeps each cell as it is, and only the formulas are then replaced by results.
                source.Copy(into.Cells[1, 1]);
                try { app.CutCopyMode = 0; } catch (Exception) { }
                try
                {
                    dynamic formulas = into.SpecialCells(-4123);
                    foreach (dynamic area in formulas.Areas) area.Value2 = area.Value2;
                }
                catch (COMException) { }
                for (int c = 1; c <= (int)source.Columns.Count; c++) { try { into.Columns[c].ColumnWidth = source.Columns[c].ColumnWidth; } catch (Exception) { } }
                return "values of " + (string)source.Address[false, false] + " copied to " + (string)into.Worksheet.Name + "!" + (string)into.Address[false, false];
            }
            // Excel's own copy takes formulas and formats along; it goes through nothing the user owns.
            source.Copy(target.Cells[1, 1]);
            try { app.CutCopyMode = 0; } catch (Exception) { }
            return (string)source.Address[false, false] + " copied to " + (string)target.Worksheet.Name + "!" + (string)target.Cells[1, 1].Address[false, false];
        }
        if (type == "pivot")
        {
            dynamic source = op.Has("source") ? Ref(book, ws, op.Need("source")) : ws.UsedRange;
            List<string> heads = Headers(source);
            for (int c = 0; c < heads.Count; c++) if (heads[c].Length == 0) throw new Fail("BAD_ARGS", "The source " + (string)source.Address[false, false] + " has an empty header in column " + Column((int)source.Column + c) + ": a pivot table needs a name over every column.");
            dynamic target;
            if (op.Has("to")) target = Ref(book, ws, op.Need("to")).Cells[1, 1];
            else
            {
                dynamic sheets = book.Worksheets;
                dynamic made = sheets.Add(After: sheets[(int)sheets.Count]);
                string wanted = op.Str("name", "透视表");
                try { made.Name = wanted; } catch (COMException) { }
                target = made.Range["A3"];
            }
            dynamic cache = book.PivotCaches().Create(SourceType: 1, SourceData: source);
            dynamic table = cache.CreatePivotTable(TableDestination: target);
            if (op.Has("name")) { try { table.Name = op.Need("name"); } catch (Exception) { } }
            Func<object, string> field = delegate(object key)
            {
                string name = Text(key).Trim();
                foreach (string head in heads) if (string.Equals(head, name, StringComparison.OrdinalIgnoreCase)) return head;
                throw new Fail("ANCHOR_MISSING", "The source has no column \"" + name + "\". Its headers: " + string.Join(", ", heads.ToArray()) + ".");
            };
            List<string> notes = new List<string>();
            int position = 0;
            foreach (object key in op.List("rows") ?? new ArrayList()) { dynamic f = table.PivotFields(field(key)); f.Orientation = 1; f.Position = ++position; }
            position = 0;
            foreach (object key in op.List("columns") ?? new ArrayList()) { dynamic f = table.PivotFields(field(key)); f.Orientation = 2; f.Position = ++position; }
            foreach (object key in op.List("filters") ?? new ArrayList()) { dynamic f = table.PivotFields(field(key)); f.Orientation = 3; }
            IList values = op.List("values");
            if (values == null || values.Count == 0) throw new Fail("BAD_ARGS", "pivot needs \"values\": [{field, fn?: sum|count|average|max|min, name?}].");
            foreach (object raw in values)
            {
                Bag item = raw is string ? new Bag(new Dictionary<string, object> { { "field", raw } }) : new Bag(raw);
                string name = field(item.Raw("field")), fn = item.Str("fn", "sum").ToLowerInvariant();
                int code = fn == "count" ? -4112 : fn == "average" || fn == "mean" || fn == "avg" ? -4106 : fn == "max" ? -4136 : fn == "min" ? -4139 : -4157;
                string caption = item.Str("name", (fn == "count" ? "计数：" : fn.StartsWith("a", StringComparison.Ordinal) || fn == "mean" ? "平均：" : fn == "max" ? "最大：" : fn == "min" ? "最小：" : "合计：") + name);
                dynamic data = table.AddDataField(table.PivotFields(name), caption, code);
                if (item.Has("numberFormat")) { try { data.NumberFormat = item.Need("numberFormat"); } catch (Exception) { } }
                else if (fn != "count") { try { data.NumberFormat = "#,##0.00"; } catch (Exception) { } }
            }
            if (op.Has("group"))
            {
                // Dates on the rows or columns, gathered by month, quarter or year.
                Bag group = new Bag(op.Raw("group"));
                string by = group.Str("by", "month");
                try
                {
                    dynamic f = table.PivotFields(field(group.Raw("field")));
                    dynamic cell = f.DataRange.Cells[1, 1];
                    object[] periods = new object[] { false, false, false, false, by == "month", by == "quarter", by == "year" };
                    cell.Group(Start: true, End: true, Periods: periods);
                }
                catch (Exception error)
                {
                    notes.Add("the dates could not be grouped by " + by + " (" + error.Message.Trim() + "): the column must hold real dates only; or add a month column with a formula and pivot on that");
                }
            }
            if (op.Has("sort"))
            {
                Bag sort = new Bag(op.Raw("sort"));
                try
                {
                    dynamic f = table.PivotFields(field(sort.Raw("field")));
                    string byName = sort.Has("by") ? sort.Need("by") : (string)table.DataFields[1].Name;
                    f.AutoSort(sort.Str("order", "desc").StartsWith("d", StringComparison.OrdinalIgnoreCase) ? 2 : 1, byName);
                }
                catch (Exception) { notes.Add("the sort was not applied: \"by\" must be the name of a value as it is shown in the table"); }
            }
            try { table.RowAxisLayout(1); } catch (Exception) { }
            dynamic made2 = table.TableRange1;
            dynamic home = made2.Worksheet;
            try { made2.EntireColumn.AutoFit(); } catch (Exception) { }
            // The result, so that it need not be read again.
            object[,] grid = Grid(made2);
            StringBuilder shown = new StringBuilder();
            int limit = Math.Min(grid.GetLength(0), 40);
            for (int r = 1; r <= limit; r++)
            {
                List<string> cells = new List<string>();
                for (int c = 1; c <= Math.Min(grid.GetLength(1), 14); c++)
                {
                    object v = grid[r, c];
                    cells.Add(IsNumber(v) ? Plain(Math.Round(Convert.ToDouble(v), 2)) : Text(v));
                }
                shown.Append("\n").Append((int)made2.Row + r - 1).Append(": ").Append(string.Join(" | ", cells.ToArray()));
            }
            return "pivot table \"" + (string)table.Name + "\" at " + (string)home.Name + "!" + (string)made2.Address[false, false] + (notes.Count > 0 ? " — NOTE " + string.Join("; ", notes.ToArray()) : "") + ":" + shown + (grid.GetLength(0) > limit ? "\n.. (" + (grid.GetLength(0) - limit) + " more rows)" : "");
        }
        if (type == "table")
        {
            dynamic data = Data(ws, op);
            dynamic table = ws.ListObjects.Add(1, data, Type.Missing, 1);
            if (op.Has("name")) { try { table.Name = op.Need("name"); } catch (Exception) { } }
            try { table.TableStyle = op.Str("style", "TableStyleMedium2"); } catch (Exception) { }
            return "table \"" + (string)table.Name + "\" made of " + (string)data.Address[false, false] + " (banded rows, filter arrows, formulas may use its column names)";
        }
        if (type == "conditional_format")
        {
            dynamic target = Cells(ws, op.Need("range"));
            string rule = op.Str("rule", ">").ToLowerInvariant();
            if (rule == "clear" || rule == "none") { target.FormatConditions.Delete(); return "conditional formats cleared"; }
            dynamic made = null;
            Func<string, string> eq = delegate(string key) { string v = Text(op.Raw(key)); return v.StartsWith("=", StringComparison.Ordinal) ? v : IsNumber(op.Raw(key)) ? "=" + v : "=\"" + v.Replace("\"", "\"\"") + "\""; };
            if (rule == "colorscale" || rule == "color scale") { target.FormatConditions.AddColorScale(3); return "colour scale on " + (string)target.Address[false, false]; }
            if (rule == "databar" || rule == "data bar") { target.FormatConditions.AddDatabar(); return "data bars on " + (string)target.Address[false, false]; }
            if (rule == "duplicate" || rule == "unique") { made = target.FormatConditions.AddUniqueValues(); made.DupeUnique = rule == "duplicate" ? 1 : 0; }
            else if (rule == "top" || rule == "bottom") { made = target.FormatConditions.AddTop10(); made.TopBottom = rule == "top" ? 1 : 0; made.Rank = Math.Max(1, op.Int("value", 10)); }
            else if (rule == "formula") made = target.FormatConditions.Add(Type: 2, Formula1: eq("value"));
            else if (rule == "blank") made = target.FormatConditions.Add(Type: 10);
            else
            {
                int code = rule == "between" ? 1 : rule == "=" ? 3 : rule == "<>" ? 4 : rule == ">" ? 5 : rule == "<" ? 6 : rule == ">=" ? 7 : rule == "<=" ? 8 : 0;
                if (code == 0) throw new Fail("BAD_ARGS", "Unknown rule \"" + rule + "\": use >, <, >=, <=, =, <>, between, top, bottom, duplicate, unique, blank, formula, colorscale, databar or clear.");
                if (code == 1) made = target.FormatConditions.Add(Type: 1, Operator: code, Formula1: eq("value"), Formula2: eq("value2"));
                else made = target.FormatConditions.Add(Type: 1, Operator: code, Formula1: eq("value"));
            }
            made.Interior.Color = Bgr(op.Str("fill", "#FFC7CE"));
            made.Font.Color = Bgr(op.Str("color", "#9C0006"));
            if (op.Has("bold")) made.Font.Bold = op.Flag("bold", false);
            return "conditional format on " + (string)target.Address[false, false] + " (" + rule + ")";
        }
        if (type == "freeze")
        {
            dynamic cell = Cells(ws, op.Str("cell", "A2")).Cells[1, 1];
            ws = cell.Worksheet;
            dynamic window = book.Windows[1];
            // Panes belong to the window as it shows this sheet: the sheet must be the one in front, in the normal view.
            dynamic was = null;
            try { was = book.ActiveSheet; } catch (Exception) { }
            ws.Activate();
            try { if ((int)window.View != 1) window.View = 1; } catch (Exception) { }
            Exception last = null;
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    try { window.FreezePanes = false; } catch (Exception) { }
                    try { window.Split = false; } catch (Exception) { }
                    window.ScrollRow = 1; window.ScrollColumn = 1;
                    window.SplitRow = (int)cell.Row - 1;
                    window.SplitColumn = (int)cell.Column - 1;
                    window.FreezePanes = (int)cell.Row > 1 || (int)cell.Column > 1;
                    last = null;
                    break;
                }
                catch (Exception error)
                {
                    last = error;
                    // Excel refuses while the window is not the active one: bring it forward inside Excel and try again.
                    try { window.Activate(); } catch (Exception) { }
                    Thread.Sleep(150);
                }
            }
            if (!Following && was != null) { try { was.Activate(); } catch (Exception) { } }
            if (last != null) throw new Fail("OFFICE_ERROR", "Excel would not freeze the panes of \"" + (string)ws.Name + "\" just now (" + last.Message.Trim() + "). The rest of your batch can go on without it; it is a convenience for scrolling only.");
            return "panes frozen above and left of " + (string)cell.Address[false, false];
        }
        if (type == "validation")
        {
            dynamic target = Cells(ws, op.Need("range"));
            IList list = op.List("list");
            if (list == null) throw new Fail("BAD_ARGS", "validation needs \"list\": the allowed values.");
            List<string> items = new List<string>();
            foreach (object item in list) items.Add(Text(item));
            target.Validation.Delete();
            target.Validation.Add(Type: 3, AlertStyle: 1, Operator: 1, Formula1: string.Join(",", items.ToArray()));
            return "cells " + (string)target.Address[false, false] + " take only: " + string.Join(", ", items.ToArray());
        }
        if (type == "add_chart")
        {
            dynamic anchor = Cells(ws, op.Str("at", "H2"));
            op.Raw("at");
            dynamic holder = null;
            try
            {
                // The way charts are made since Excel 2013: today's colours and type instead of those of 2007.
                dynamic shape = ws.Shapes.AddChart2(-1, ChartKind(op.Str("type", op.Str("chart", "column"))), (double)anchor.Left, (double)anchor.Top, op.Num("width", 480), op.Num("height", 290));
                holder = ws.ChartObjects((string)shape.Name);
            }
            catch (Fail) { throw; }
            catch (Exception) { holder = null; }
            if (holder == null) holder = ws.ChartObjects().Add((double)anchor.Left, (double)anchor.Top, op.Num("width", 480), op.Num("height", 290));
            try { return Chart(book, ws, holder, op, true); }
            catch (Exception) { try { holder.Delete(); } catch (Exception) { } throw; }
        }
        if (type == "set_chart") return Chart(book, ws, ChartByName(ws, op.Need("chart")), op, false);
        if (type == "delete_chart") { dynamic holder = ChartByName(ws, op.Need("chart")); holder.Delete(); return "chart deleted"; }
        return null;
    }

    // ───────────────────────── PowerPoint ─────────────────────────

    static dynamic Slide(dynamic deck, Bag op)
    {
        int index = op.Int("slide", 0), total = (int)deck.Slides.Count;
        if (index < 1 || index > total) throw new Fail("ANCHOR_MISSING", "There is no slide " + index + " (the deck has " + total + ").");
        return deck.Slides[index];
    }

    /// The shapes of a slide one can write into or replace: those inside groups too (a template keeps most of its
    /// text there), the groups themselves left out.
    static List<object> Leaves(dynamic slide)
    {
        List<object> all = new List<object>();
        foreach (dynamic shape in slide.Shapes)
        {
            bool group = false;
            try { group = (int)shape.Type == 6; } catch (Exception) { }
            if (!group) { all.Add(shape); continue; }
            try { foreach (dynamic item in shape.GroupItems) all.Add(item); }
            catch (Exception) { all.Add(shape); }
        }
        return all;
    }

    /// A shape by "title" / "body", by "#id" (as office_read lists it), by name, or by position among the shapes.
    static dynamic Shape(dynamic slide, string key)
    {
        try
        {
            if (key == "title") return slide.Shapes.Title;
            if (key == "body") return slide.Shapes.Placeholders[2];
            if (key.StartsWith("#", StringComparison.Ordinal))
            {
                int id;
                if (int.TryParse(key.Substring(1), out id)) foreach (dynamic shape in Leaves(slide)) if ((int)shape.Id == id) return shape;
                throw new COMException("no such id");
            }
            int index;
            if (int.TryParse(key, out index)) return slide.Shapes[index];
            try { return slide.Shapes[key]; }
            catch (COMException) { foreach (dynamic shape in Leaves(slide)) if ((string)shape.Name == key) return shape; throw; }
        }
        catch (COMException)
        {
            List<string> names = new List<string>();
            foreach (dynamic shape in Leaves(slide)) { if (names.Count < 40) names.Add("\"" + (string)shape.Name + "\" #" + (int)shape.Id); }
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
            int plain = 0;
            foreach (dynamic shape in Leaves(slide))
            {
                Dictionary<string, object> item = new Dictionary<string, object>();
                int kind = 1;
                try { kind = (int)shape.Type; } catch (Exception) { }
                bool worded = false;
                try { worded = Truthy(shape.HasTextFrame) && Truthy(shape.TextFrame.HasText); } catch (Exception) { }
                // Decoration (lines, blank blocks) is counted, not listed: a template has hundreds of them.
                if (!worded && !full && kind != 13 && kind != 14 && kind != 3 && kind != 19 && kind != 24) { plain++; continue; }
                item["name"] = (string)shape.Name;
                try { item["id"] = (int)shape.Id; } catch (Exception) { }
                item["type"] = ShapeType(kind);
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
            if (plain > 0) entry["plain"] = plain;
            slides.Add(entry);
        }
        result["items"] = slides;
        return result;
    }

    static void PptWrite(dynamic textRange, string text)
    {
        // A line that starts with tabs (or two spaces per step) is a bullet that many levels down.
        string[] lines = text.Split('\r');
        int[] levels = new int[lines.Length];
        bool nested = false;
        for (int i = 0; i < lines.Length; i++)
        {
            string line = lines[i];
            int at = 0, level = 0;
            while (at < line.Length)
            {
                if (line[at] == '\t') { level++; at++; }
                else if (at + 1 < line.Length && line[at] == ' ' && line[at + 1] == ' ') { level++; at += 2; }
                else break;
            }
            if (level > 0 && at < line.Length) { lines[i] = line.Substring(at); levels[i] = Math.Min(level, 4); nested = true; }
        }
        if (nested) text = string.Join("\r", lines);
        PptWriteText(textRange, text);
        // Written piece by piece, the range in hand may no longer span what was written: take the frame's text anew.
        try
        {
            string now = (string)textRange.Text ?? "";
            Trace("written " + text.Length + " chars, the range holds " + now.Length);
            if (now.Length < text.Length) textRange = textRange.Parent.TextRange;
        }
        catch (Exception) { }
        if (nested)
        {
            for (int i = 0; i < lines.Length; i++)
            {
                try { textRange.Paragraphs(i + 1).IndentLevel = levels[i] + 1; } catch (Exception) { }
            }
        }
        if (text.IndexOf('$') >= 0) PptMath(textRange);
    }

    static void PptWriteText(dynamic textRange, string text)
    {
        if (!Typing || text.Length < 4) textRange.Text = text;
        else
        {
            textRange.Text = "";
            foreach (string piece in Pieces(text)) { textRange.InsertAfter(piece); Thread.Sleep(14); }
        }
    }

    // PowerPoint shows native equations but has no call to make one. Word has: the formula is built in a hidden
    // scratch document there and carried over through the clipboard, which gives PowerPoint's own editable equation.
    // The user's clipboard is put back when the batch is over.

    static dynamic MathWord, MathDoc;
    static bool MathReads, MathWasLatex;
    static int PptMathFailed;

    // That Word is nobody's to see, so it lives for one batch only: started at the first formula (about a second),
    // gone when the batch ends. Its process id is written down while it lives; if the helper is ever cut off in the
    // middle, the next helper finds the note and ends what was left behind.

    static int MathPid;

    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
    [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);
    [DllImport("user32.dll")] static extern uint GetClipboardSequenceNumber();

    static string MathNotePath()
    {
        return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "dsh-office", "math-word.pid");
    }

    /// A hidden Word left by a helper that did not get to close it: ended, if it still is that process and shows no window.
    static void Sweep()
    {
        try
        {
            string note = MathNotePath();
            if (!File.Exists(note)) return;
            int pid;
            if (int.TryParse(File.ReadAllText(note).Trim(), out pid) && pid > 0)
            {
                try
                {
                    using (System.Diagnostics.Process process = System.Diagnostics.Process.GetProcessById(pid))
                    {
                        if (process.ProcessName.Equals("WINWORD", StringComparison.OrdinalIgnoreCase) && process.MainWindowHandle == IntPtr.Zero) process.Kill();
                    }
                }
                catch (Exception) { }
            }
            File.Delete(note);
        }
        catch (Exception) { }
    }

    /// A Word of our own, never shown: the formulas are built there, so the Word the user works in is left alone.
    static bool MathSource()
    {
        if (MathDoc != null) return MathReads;
        try
        {
            MathWord = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
            MathWord.Visible = false;
            try { MathWord.DisplayAlerts = 0; } catch (Exception) { }
            MathDoc = MathWord.Documents.Add();
            try
            {
                uint pid;
                GetWindowThreadProcessId(new IntPtr((int)MathDoc.ActiveWindow.Hwnd), out pid);
                MathPid = (int)pid;
                if (MathPid != 0)
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(MathNotePath()));
                    File.WriteAllText(MathNotePath(), MathPid.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
            }
            catch (Exception) { }
            dynamic bars = MathWord.CommandBars;
            MathWasLatex = (bool)bars.GetPressedMso("EquationLaTexFormat");
            if (!MathWasLatex) bars.ExecuteMso("EquationLaTexFormat");
            MathReads = (bool)bars.GetPressedMso("EquationLaTexFormat");
        }
        catch (Exception) { MathReads = false; }
        return MathReads && MathDoc != null;
    }

    // The clipboard is the only way into a slide for an equation. The user's is treated as the computer-use plugin
    // treats it: everything on it is set aside first, and put back right after, unless the user copied something
    // in between (then theirs is the newer one and stays).

    static DataObject ClipSaved;
    static bool ClipTaken;
    static uint ClipOurs;

    static DataObject Snapshot()
    {
        for (int i = 0; i < 5; i++)
        {
            try
            {
                IDataObject now = Clipboard.GetDataObject();
                if (now == null) return null;
                string[] formats = now.GetFormats(false);
                if (formats.Length == 0) return null;
                DataObject copy = new DataObject();
                foreach (string format in formats)
                {
                    if (format == "CanIncludeInClipboardHistory" || format == "CanUploadToCloudClipboard" || format == "ExcludeClipboardContentFromMonitorProcessing") continue;
                    try { object value = now.GetData(format, false); if (value != null) copy.SetData(format, false, value); } catch (Exception) { }
                }
                return copy;
            }
            catch (ExternalException) { Thread.Sleep(50); }
        }
        return null;
    }

    static void SaveClipboard()
    {
        if (ClipTaken) return;
        ClipSaved = Snapshot();
        ClipTaken = true;
    }

    /// Back to what the user had, kept out of the clipboard history; not if they copied something meanwhile.
    static void RestoreClipboard()
    {
        if (!ClipTaken) return;
        ClipTaken = false;
        DataObject data = ClipSaved;
        ClipSaved = null;
        if (GetClipboardSequenceNumber() != ClipOurs) return;
        for (int i = 0; i < 5; i++)
        {
            try
            {
                if (data == null) Clipboard.Clear();
                else
                {
                    data.SetData("ExcludeClipboardContentFromMonitorProcessing", false, new MemoryStream(new byte[] { 1, 0, 0, 0 }));
                    data.SetData("CanIncludeInClipboardHistory", false, new MemoryStream(BitConverter.GetBytes(0)));
                    data.SetData("CanUploadToCloudClipboard", false, new MemoryStream(BitConverter.GetBytes(0)));
                    Clipboard.SetDataObject(data, true);
                }
                return;
            }
            catch (ExternalException) { Thread.Sleep(50); }
        }
    }

    /// End of a batch: the clipboard is the user's again, and the helper's Word is closed.
    static void MathDone()
    {
        RestoreClipboard();
        if (MathWord == null) return;
        try
        {
            dynamic bars = MathWord.CommandBars;
            if (!MathWasLatex && !(bool)bars.GetPressedMso("EquationUnicodeFormat")) bars.ExecuteMso("EquationUnicodeFormat");
        }
        catch (Exception) { }
        try { if (MathDoc != null) MathDoc.Close(SaveChanges: 0); } catch (Exception) { }
        try
        {
            // Should a document of the user's have landed in this Word meanwhile, it is theirs: shown, not closed.
            if ((int)MathWord.Documents.Count == 0) MathWord.Quit(SaveChanges: 0);
            else MathWord.Visible = true;
        }
        catch (Exception) { }
        MathDoc = null;
        MathWord = null;
        MathPid = 0;
        try { File.Delete(MathNotePath()); } catch (Exception) { }
    }

    static void MathQuit() { try { MathDone(); } catch (Exception) { } }

    /// Where PowerPoint keeps an equation in its typed-out form (a table cell does), a formula is written as text
    /// instead: Word's linear form of it, with the subscripts and powers set as such.
    static void PptFlatMath(dynamic textRange)
    {
        string text = (string)textRange.Text;
        if (text == null || text.IndexOf('$') < 0) return;
        List<string> sources = new List<string>();
        foreach (System.Text.RegularExpressions.Match m in Dollars.Matches(text)) sources.Add((m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).Trim());
        Dictionary<string, string> linear = ToLinear(sources);
        StringBuilder o = new StringBuilder();
        List<int[]> spans = new List<int[]>();   // start, length, 1 = subscript / 2 = superscript
        int at = 0;
        foreach (System.Text.RegularExpressions.Match m in Dollars.Matches(text))
        {
            o.Append(text, at, m.Index - at);
            at = m.Index + m.Length;
            string source = (m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).Trim(), read;
            if (!linear.TryGetValue(source, out read)) { o.Append(source); PptMathFailed++; continue; }
            int i = 0;
            while (i < read.Length)
            {
                char c = read[i];
                if ((c == '_' || c == '^') && i + 1 < read.Length)
                {
                    int from = i + 1, to;
                    if (read[from] == '(')
                    {
                        int depth = 0;
                        for (to = from; to < read.Length; to++) { if (read[to] == '(') depth++; else if (read[to] == ')' && --depth == 0) break; }
                        if (to >= read.Length) { o.Append(c); i++; continue; }
                        string inner = read.Substring(from + 1, to - from - 1);
                        spans.Add(new int[] { o.Length, inner.Length, c == '_' ? 1 : 2 });
                        o.Append(inner);
                        i = to + 1;
                    }
                    else
                    {
                        to = from;
                        while (to < read.Length && (char.IsLetterOrDigit(read[to]) || char.IsSurrogate(read[to]) || read[to] == '−' || read[to] == '-' || read[to] == '+' || read[to] == '∞') && !(to > from && (read[to] == '-' || read[to] == '+' || read[to] == '−'))) to++;
                        if (to == from) { o.Append(c); i++; continue; }
                        spans.Add(new int[] { o.Length, to - from, c == '_' ? 1 : 2 });
                        o.Append(read, from, to - from);
                        i = to;
                    }
                    // The space that ended the script in the linear form has done its work.
                    if (i < read.Length && read[i] == ' ') i++;
                    continue;
                }
                if (c == '▒' || c == '〖' || c == '〗' || c == '\u2061') { i++; continue; }
                o.Append(c);
                i++;
            }
        }
        o.Append(text, at, text.Length - at);
        textRange.Text = o.ToString();
        foreach (int[] span in spans)
        {
            try
            {
                dynamic part = textRange.Characters(span[0] + 1, span[1]);
                if (span[2] == 1) part.Font.Subscript = -1; else part.Font.Superscript = -1;
            }
            catch (Exception) { }
        }
    }

    /// $...$ in the text of a shape become native equations, in place.
    static int PptMath(dynamic textRange)
    {
        string text = (string)textRange.Text;
        if (text == null || text.IndexOf('$') < 0) return 0;
        System.Text.RegularExpressions.MatchCollection found = Dollars.Matches(text);
        if (found.Count == 0) return 0;
        Trace("text of " + text.Length + " chars with " + found.Count + " formula(s): " + Clip(text, 60));
        if (!MathSource()) { PptMathFailed += found.Count; return 0; }
        SaveClipboard();
        int made = 0;
        for (int k = found.Count - 1; k >= 0; k--)
        {
            System.Text.RegularExpressions.Match m = found[k];
            string source = (m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).Trim();
            System.Text.RegularExpressions.Match tagged = Tag.Match(source);
            if (tagged.Success) source = source.Remove(tagged.Index, tagged.Length).Trim();
            try
            {
                if (!WordReads(source)) { PptMathFailed++; continue; }
                string latex = ForWord(source);
                MathDoc.Content.Delete();
                MathDoc.Content.InsertBefore(latex);
                dynamic built = MathDoc.OMaths.Add(MathDoc.Range(0, latex.Length)).OMaths[1];
                built.BuildUp();
                dynamic equation = MathDoc.OMaths[1].Range;
                if (Unbuilt((string)equation.Text, source)) { PptMathFailed++; continue; }
                dynamic spot = textRange.Characters(m.Index + 1, m.Length);
                bool pasted = false;
                for (int attempt = 0; attempt < 8 && !pasted; attempt++)
                {
                    try
                    {
                        int ink = 0; float size = 0;
                        try { ink = (int)spot.Font.Color.RGB; size = (float)spot.Font.Size; } catch (Exception) { }
                        string was = (string)textRange.Text ?? "";
                        int before = was.Length;
                        equation.Copy();
                        spot.Paste();
                        // PowerPoint at times reports a paste that put nothing in (the clipboard was not ready yet):
                        // the text is then exactly as before, dollar signs and all. That is no paste; wait and go again.
                        if (((string)textRange.Text ?? "") == was)
                        {
                            Thread.Sleep(80 + 80 * attempt);
                            spot = textRange.Characters(m.Index + 1, m.Length);
                            continue;
                        }
                        pasted = true;
                        ClipOurs = GetClipboardSequenceNumber();
                        // The equation arrives in Word's black: it takes the colour and size of the text around it.
                        // What was pasted is what now stands where the source stood, as long as the text grew or shrank by.
                        try
                        {
                            string now = (string)textRange.Text ?? "";
                            int length = now.Length - (before - m.Length), start = m.Index + 1;
                            // PowerPoint may take the space before the formula away as it pastes: the equation then
                            // begins one place earlier, and its first letter would be left in Word's black.
                            if (m.Index > 0 && was[m.Index - 1] == ' ' && (now.Length < m.Index || now[m.Index - 1] != ' ')) { start = m.Index; length += 1; }
                            if (length > 0)
                            {
                                dynamic placed = textRange.Characters(start, length);
                                placed.Font.Color.RGB = ink;
                                if (size > 1) placed.Font.Size = size;
                            }
                        }
                        catch (Exception) { }
                    }
                    catch (COMException) { Thread.Sleep(60 + 60 * attempt); }
                }
                if (pasted) made++; else PptMathFailed++;
                Trace("formula " + k + " \"" + source + "\" pasted " + pasted + "; text now " + ((string)textRange.Text ?? "").Length);
            }
            catch (Exception error) { PptMathFailed++; Trace("formula " + k + " failed: " + error.GetType().Name + " " + error.Message); }
        }
        RestoreClipboard();
        return made;
    }

    static void GoTo(dynamic deck, int slide)
    {
        if (!Following) return;
        try { deck.Windows[1].View.GotoSlide(slide); } catch (Exception error) { Trace("goto slide failed: " + error.GetType().Name); }
    }

    static void PptText(dynamic shape, Bag op)
    {
        if (op.Has("text")) PptWrite(shape.TextFrame.TextRange, Lines(op.Raw("text")));
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

    static readonly Dictionary<string, int> Forms = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        { "rectangle", 1 }, { "rounded", 5 }, { "ellipse", 9 }, { "diamond", 4 }, { "triangle", 7 }, { "arrow", 33 }, { "arrow_left", 34 }, { "arrow_up", 35 }, { "arrow_down", 36 },
        { "chevron", 52 }, { "pentagon", 51 }, { "star", 92 }, { "callout", 108 }, { "hexagon", 10 }, { "cloud", 179 },
    };

    /// SmartArt layouts by a plain name; the value is the end of the layout's id.
    static readonly Dictionary<string, string> Diagrams = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "list", "default" }, { "bullet_list", "vList2" }, { "horizontal_list", "hList1" }, { "process", "process1" }, { "chevron", "chevron1" }, { "arrows", "hProcess9" },
        { "timeline", "hProcess11" }, { "steps", "process2" }, { "cycle", "cycle2" }, { "radial", "radial1" }, { "hierarchy", "hierarchy1" }, { "org_chart", "orgChart1" },
        { "tree", "hierarchy2" }, { "pyramid", "pyramid1" }, { "venn", "venn1" }, { "matrix", "matrix1" }, { "funnel", "funnel1" }, { "target", "target1" }, { "balance", "balance1" }, { "gear", "gear1" },
    };

    /// One item of a diagram on its node: a text, or {text, children:[..]} whose children hang below it.
    static int Branch(dynamic node, object raw)
    {
        IList children = null;
        string text;
        Dictionary<string, object> map = raw as Dictionary<string, object>;
        if (map != null)
        {
            object value;
            text = map.TryGetValue("text", out value) && value != null ? Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) : "";
            if (map.TryGetValue("children", out value)) children = value as IList;
        }
        else text = Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture);
        if (text.IndexOf('$') >= 0)
        {
            // A diagram node holds plain text only: a formula is written in Word's linear form (Φ, x^2).
            List<string> sources = new List<string>();
            foreach (System.Text.RegularExpressions.Match m in Dollars.Matches(text)) sources.Add((m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).Trim());
            Dictionary<string, string> linear = ToLinear(sources);
            text = Dollars.Replace(text, delegate(System.Text.RegularExpressions.Match m)
            {
                string source = (m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value).Trim(), read;
                return linear.TryGetValue(source, out read) ? read : source;
            });
        }
        node.TextFrame2.TextRange.Text = text;
        int made = 1;
        if (children != null)
        {
            dynamic previous = null;
            foreach (object child in children)
            {
                dynamic below = previous == null ? node.AddNode(5) : previous.AddNode(2);
                made += Branch(below, child);
                previous = below;
            }
        }
        return made;
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
            GoTo(deck, index);
            if (op.Has("title")) PptWrite(slide.Shapes.Title.TextFrame.TextRange, Lines(op.Raw("title")));
            if (op.Has("body")) PptWrite(slide.Shapes.Placeholders[2].TextFrame.TextRange, Lines(op.Raw("body")));
            if (op.Has("notes")) slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = Lines(op.Raw("notes"));
            return "slide " + index + " added";
        }
        if (type == "theme") return SetTheme(deck, op);
        if (type == "slide") return DesignSlide(deck, op);
        if (type == "reuse_slide")
        {
            int count = (int)deck.Slides.Count, from = op.Int("from", 0), at = op.Int("at", count + 1);
            if (from < 1 || from > count) throw new Fail("ANCHOR_MISSING", "\"from\" is slide " + from + ", but the deck has " + count + " slides.");
            if (at < 1 || at > count + 1) throw new Fail("BAD_ARGS", "\"at\" must be between 1 and " + (count + 1) + ".");
            deck.Slides[from].Duplicate();
            deck.Slides[from + 1].MoveTo(at);
            dynamic made = deck.Slides[at];
            GoTo(deck, at);
            int written = 0, swapped = 0;
            Dictionary<string, object> texts = op.Raw("texts") as Dictionary<string, object>;
            if (texts != null)
            {
                foreach (KeyValuePair<string, object> pair in texts)
                {
                    dynamic slot = Shape(made, pair.Key);
                    // The words change, the look stays: the new text takes the type of the old, and shrinks to fit its box.
                    try { slot.TextFrame2.AutoSize = 2; } catch (Exception) { }
                    PptWrite(slot.TextFrame.TextRange, Lines(pair.Value ?? ""));
                    written++;
                }
            }
            Dictionary<string, object> images = op.Raw("images") as Dictionary<string, object>;
            if (images != null)
            {
                foreach (KeyValuePair<string, object> pair in images)
                {
                    dynamic old = Shape(made, pair.Key);
                    string path = Convert.ToString(pair.Value);
                    if (!File.Exists(path)) throw new Fail("BAD_ARGS", "Image \"" + path + "\" does not exist.");
                    float left = (float)old.Left, top = (float)old.Top, width = (float)old.Width, height = (float)old.Height;
                    string name = (string)old.Name;
                    Page page = new Page();
                    page.Slide = made;
                    dynamic picture = Photo(page, left, top, width, height, path);
                    try { old.Delete(); } catch (Exception) { }
                    try { picture.Name = name; } catch (Exception) { }
                    swapped++;
                }
            }
            IList drop = op.List("delete");
            if (drop != null) foreach (object key in drop) { try { Shape(made, Convert.ToString(key)).Delete(); } catch (Fail) { } }
            if (op.Has("notes")) { try { made.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = Lines(op.Raw("notes")); } catch (Exception) { } }
            Renumber = true;
            return "slide " + from + " reused as slide " + at + " (" + written + " text(s), " + swapped + " picture(s) replaced)";
        }
        if (type == "delete_slides")
        {
            int count = (int)deck.Slides.Count, first = op.Int("from", 1), last = op.Int("to", first);
            if (first < 1 || last > count || first > last) throw new Fail("BAD_ARGS", "\"from\"..\"to\" must lie within 1.." + count + ".");
            if (last - first + 1 >= count) throw new Fail("BAD_ARGS", "That would delete every slide of the deck.");
            for (int i = last; i >= first; i--) deck.Slides[i].Delete();
            Renumber = true;
            return "slides " + first + "–" + last + " deleted, " + (int)deck.Slides.Count + " left";
        }
        dynamic target = Slide(deck, op);
        GoTo(deck, op.Int("slide", 0));
        if (type == "delete_slide") { target.Delete(); Renumber = true; return "slide deleted"; }
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
            dynamic picture = target.Shapes.AddPicture(Path.GetFullPath(path), 0, -1, (float)op.Num("left", 60), (float)op.Num("top", 60), -1, -1);
            float drawn = (float)picture.Width;
            if (op.Has("width") || op.Has("height"))
            {
                picture.LockAspectRatio = op.Has("width") && op.Has("height") ? 0 : -1;
                if (op.Has("width")) picture.Width = (float)op.Num("width", 100);
                if (op.Has("height")) picture.Height = (float)op.Num("height", 100);
            }
            if (op.Has("name")) picture.Name = op.Need("name");
            Lettering(path, (float)picture.Width, drawn, true);
            float sw = (float)deck.PageSetup.SlideWidth, sh = (float)deck.PageSetup.SlideHeight;
            return "picture \"" + (string)picture.Name + "\" added at [" + Math.Round((float)picture.Left) + ", " + Math.Round((float)picture.Top) + ", " + Math.Round((float)picture.Width) + ", " + Math.Round((float)picture.Height) + "], " + Math.Round(100 * (float)picture.Width * (float)picture.Height / (sw * sh)) + "% of the slide" + SmallNote();
        }
        if (type == "duplicate_slide")
        {
            dynamic copy = target.Duplicate();
            int at = (int)copy.SlideIndex;
            if (op.Has("to")) { copy.MoveTo(op.Int("to", at)); at = op.Int("to", at); }
            GoTo(deck, at);
            return "slide duplicated as slide " + at;
        }
        if (type == "set_layout")
        {
            int layout;
            string name = op.Need("layout");
            if (!Layouts.TryGetValue(name, out layout)) throw new Fail("BAD_ARGS", "Unknown layout \"" + name + "\": use " + string.Join(", ", new List<string>(Layouts.Keys).ToArray()) + ".");
            target.Layout = layout;
            return "layout set";
        }
        if (type == "set_background")
        {
            target.FollowMasterBackground = 0;
            target.Background.Fill.Visible = -1;
            target.Background.Fill.Solid();
            target.Background.Fill.ForeColor.RGB = Bgr(op.Need("color"));
            return "background set";
        }
        if (type == "add_shape")
        {
            string kind = op.Str("kind", "rectangle").ToLowerInvariant();
            float left = (float)op.Num("left", 60), top = (float)op.Num("top", 60), width = (float)op.Num("width", 200), height = (float)op.Num("height", 80);
            dynamic made;
            if (kind == "line" || kind == "arrow_line")
            {
                made = target.Shapes.AddLine(left, top, left + width, top + height);
                if (kind == "arrow_line") made.Line.EndArrowheadStyle = 2;
            }
            else
            {
                int form;
                if (!Forms.TryGetValue(kind, out form)) throw new Fail("BAD_ARGS", "Unknown shape kind \"" + kind + "\": use " + string.Join(", ", new List<string>(Forms.Keys).ToArray()) + ", line, arrow_line.");
                made = target.Shapes.AddShape(form, left, top, width, height);
                if (op.Has("fill")) { made.Fill.Visible = -1; made.Fill.Solid(); made.Fill.ForeColor.RGB = Bgr(op.Need("fill")); }
                if (op.Has("text")) PptText(made, op);
            }
            if (op.Has("line"))
            {
                string line = op.Need("line");
                if (line == "none") made.Line.Visible = 0; else { made.Line.Visible = -1; made.Line.ForeColor.RGB = Bgr(line); }
            }
            if (op.Has("lineWidth")) made.Line.Weight = (float)op.Num("lineWidth", 1);
            if (op.Has("name")) made.Name = op.Need("name");
            return "shape \"" + (string)made.Name + "\" added";
        }
        if (type == "add_table")
        {
            IList data = op.List("data");
            if (data == null || data.Count == 0) throw new Fail("BAD_ARGS", "\"data\" must list the rows of the table.");
            int rows = data.Count, cols = 1;
            foreach (object row in data) { IList cells = row as IList; if (cells != null) cols = Math.Max(cols, cells.Count); }
            dynamic made = target.Shapes.AddTable(rows, cols, (float)op.Num("left", 60), (float)op.Num("top", 120), (float)op.Num("width", 600), (float)op.Num("height", 28 * rows));
            dynamic grid = made.Table;
            for (int r = 0; r < rows; r++)
            {
                IList cells = data[r] as IList;
                if (cells == null) continue;
                for (int c = 0; c < cells.Count; c++)
                {
                    if (cells[c] == null) continue;
                    dynamic range = grid.Cell(r + 1, c + 1).Shape.TextFrame.TextRange;
                    string value = Convert.ToString(cells[c], System.Globalization.CultureInfo.InvariantCulture);
                    range.Text = value;
                    if (op.Has("size")) range.Font.Size = (float)op.Num("size", 16);
                    if (op.Has("font")) { string font = op.Need("font"); range.Font.Name = font; try { range.Font.NameFarEast = font; } catch (COMException) { } }
                    if (op.Has("align")) { string align = op.Need("align"); range.ParagraphFormat.Alignment = align == "center" ? 2 : align == "right" ? 3 : 1; }
                    if (value.IndexOf('$') >= 0) PptFlatMath(range);
                }
                if (Typing) Thread.Sleep(25);
            }
            if (op.Has("header") && !op.Flag("header", true)) grid.FirstRow = 0;
            if (op.Has("name")) made.Name = op.Need("name");
            return "table \"" + (string)made.Name + "\" added (" + rows + "×" + cols + ")";
        }
        if (type == "add_smartart")
        {
            IList items = op.List("items");
            if (items == null || items.Count == 0) throw new Fail("BAD_ARGS", "\"items\" must list the texts of the diagram.");
            string wanted = op.Str("layout", "process");
            string id;
            if (!Diagrams.TryGetValue(wanted, out id)) id = wanted;
            dynamic layouts = deck.Application.SmartArtLayouts, layout = null;
            int count = (int)layouts.Count;
            for (int i = 1; i <= count && layout == null; i++)
            {
                dynamic candidate = layouts[i];
                string full = (string)candidate.Id;
                if (full.EndsWith("/" + id, StringComparison.OrdinalIgnoreCase) || string.Equals((string)candidate.Name, wanted, StringComparison.OrdinalIgnoreCase)) layout = candidate;
            }
            if (layout == null) throw new Fail("BAD_ARGS", "Unknown SmartArt layout \"" + wanted + "\": use " + string.Join(", ", new List<string>(Diagrams.Keys).ToArray()) + ", or the name the layout has in PowerPoint.");
            dynamic made = target.Shapes.AddSmartArt(layout, (float)op.Num("left", 60), (float)op.Num("top", 130), (float)op.Num("width", 840), (float)op.Num("height", 330));
            dynamic art = made.SmartArt;
            // The layout comes with sample nodes: all but one go, and the items are built from that one.
            for (int guard = 0; guard < 60 && (int)art.AllNodes.Count > 1; guard++) art.AllNodes[(int)art.AllNodes.Count].Delete();
            dynamic first = art.AllNodes[1], previous = null;
            int nodes = 0;
            foreach (object raw in items)
            {
                dynamic node = previous == null ? first : previous.AddNode(2);
                nodes += Branch(node, raw);
                previous = node;
            }
            if (op.Has("colors")) { try { art.Color = deck.Application.SmartArtColors[op.Int("colors", 1)]; } catch (Exception) { } }
            if (op.Has("look")) { try { art.QuickStyle = deck.Application.SmartArtQuickStyles[op.Int("look", 1)]; } catch (Exception) { } }
            if (op.Has("name")) made.Name = op.Need("name");
            return "SmartArt \"" + (string)made.Name + "\" added (" + (string)layout.Name + ", " + nodes + " items)";
        }
        dynamic shape = Shape(target, op.Need("shape"));
        if (type == "format_text")
        {
            dynamic all = shape.TextFrame.TextRange, part = all;
            if (op.Has("find"))
            {
                part = all.Find(op.Need("find"));
                if (part == null) throw new Fail("NOT_FOUND", "The text \"" + op.Need("find") + "\" does not occur in that shape.");
            }
            if (op.Has("font")) { string font = op.Need("font"); part.Font.Name = font; try { part.Font.NameFarEast = font; } catch (COMException) { } }
            if (op.Has("size")) part.Font.Size = (float)op.Num("size", 18);
            if (op.Has("bold")) part.Font.Bold = op.Flag("bold", false) ? -1 : 0;
            if (op.Has("italic")) part.Font.Italic = op.Flag("italic", false) ? -1 : 0;
            if (op.Has("underline")) part.Font.Underline = op.Flag("underline", false) ? -1 : 0;
            if (op.Has("color")) part.Font.Color.RGB = Bgr(op.Need("color"));
            if (op.Has("align")) { string align = op.Need("align"); part.ParagraphFormat.Alignment = align == "center" ? 2 : align == "right" ? 3 : 1; }
            if (op.Has("bullets")) part.ParagraphFormat.Bullet.Visible = op.Flag("bullets", true) ? -1 : 0;
            if (op.Has("lineSpacing")) part.ParagraphFormat.SpaceWithin = (float)op.Num("lineSpacing", 1);
            if (op.Has("fit")) { try { shape.TextFrame2.AutoSize = op.Need("fit") == "shrink" ? 2 : op.Need("fit") == "grow" ? 1 : 0; } catch (Exception) { } }
            return "text formatted";
        }
        if (type == "set_text") { op.Need("text"); PptText(shape, op); return "text set"; }
        if (type == "set_shape") { PptBox(shape, op); PptText(shape, op); return "shape updated"; }
        if (type == "delete_shape") { shape.Delete(); return "shape deleted"; }
        throw new Fail("BAD_ARGS", "Unknown PowerPoint operation \"" + type + "\".");
    }

    // ───────────────────────── designed slides ─────────────────────────
    //
    // A slide laid out by the helper from its content: the agent says what the slide is (a cover, three figures, a
    // chart, a formula ..) and gives the words; position, size, colour, type and motion come from the deck's theme.
    // Everything is measured on a 960 × 540 canvas and scaled to the deck's real slide size.

    class Theme
    {
        public string Name = "ink";
        public string Bg = "#FFFFFF", Surface = "#F3F5F9", Text = "#1A2233", Muted = "#6B7280", Primary = "#0F3D75", Accent = "#C8102E", Line = "#D9DEE7";
        public string TitleFont = "微软雅黑", BodyFont = "微软雅黑";
        public bool Dark;
        public string Transition = "fade";
        public bool Animate = true;
        /// What lies behind the content: "" (the flat colour), paper, gradient, grid, dots, or "image:" and a path.
        public string Texture = "";
        /// How the section slides of the deck look: solid, side, band or number.
        public string Section = "";
        /// A slide of the deck every designed slide is drawn on (a template's background page), 0 = a blank slide.
        public int Base;
        /// The part of the slide the content may use: left, top, width, height in points (empty = all of it).
        public double[] Area;
    }

    static Theme Named(string name)
    {
        Theme t = new Theme();
        t.Name = name;
        switch (name)
        {
            case "ink": break;
            case "paper": t.Bg = "#FCFBF8"; t.Surface = "#EEF1F5"; t.Text = "#1F2937"; t.Muted = "#6B7280"; t.Primary = "#1F4E79"; t.Accent = "#C2452D"; t.Line = "#D6DAE0"; t.TitleFont = "Noto Serif SC"; break;
            case "night": t.Bg = "#0E0E10"; t.Surface = "#1B1B1F"; t.Text = "#F2EDE4"; t.Muted = "#9A948A"; t.Primary = "#E7C873"; t.Accent = "#C8553D"; t.Line = "#35322E"; t.TitleFont = "Noto Serif SC"; t.Dark = true; break;
            case "chalk": t.Bg = "#1F3A32"; t.Surface = "#274A40"; t.Text = "#F3EFE0"; t.Muted = "#A9BDB3"; t.Primary = "#F2D16B"; t.Accent = "#F2A7A0"; t.Line = "#56776A"; t.TitleFont = "楷体"; t.Dark = true; break;
            case "ocean": t.Bg = "#F7FAFC"; t.Surface = "#E4EEF5"; t.Text = "#102A43"; t.Muted = "#627D98"; t.Primary = "#0B7285"; t.Accent = "#E8590C"; t.Line = "#CBD9E4"; break;
            case "graphite": t.Bg = "#FFFFFF"; t.Surface = "#F2F2F2"; t.Text = "#222222"; t.Muted = "#777777"; t.Primary = "#2B2B2B"; t.Accent = "#E8590C"; t.Line = "#DDDDDD"; break;
            case "forest": t.Bg = "#FAFBF7"; t.Surface = "#EAF0E4"; t.Text = "#1F2A1F"; t.Muted = "#6B7A66"; t.Primary = "#2F6B3A"; t.Accent = "#B7791F"; t.Line = "#D5DDCC"; break;
            case "sepia": t.Bg = "#F5EEDC"; t.Surface = "#EAE0C8"; t.Text = "#3B2F23"; t.Muted = "#7A6A55"; t.Primary = "#7A1F1F"; t.Accent = "#A8782A"; t.Line = "#D9CCAF"; t.TitleFont = "Noto Serif SC"; t.Texture = "paper"; break;
            case "crimson": t.Bg = "#FFFDF8"; t.Surface = "#F7EDE4"; t.Text = "#2B1D1A"; t.Muted = "#8A7268"; t.Primary = "#B01E23"; t.Accent = "#C89B3C"; t.Line = "#EAD9CC"; t.Texture = "gradient"; break;
            case "slate": t.Bg = "#F4F6F8"; t.Surface = "#E6EAEE"; t.Text = "#1F2A37"; t.Muted = "#66727F"; t.Primary = "#2F4A63"; t.Accent = "#0FA3B1"; t.Line = "#D3DAE1"; t.Texture = "grid"; break;
            case "plum": t.Bg = "#17131F"; t.Surface = "#231D30"; t.Text = "#F1ECF8"; t.Muted = "#A79FB8"; t.Primary = "#C9A7FF"; t.Accent = "#FF8FA3"; t.Line = "#3A3150"; t.Dark = true; break;
            default: throw new Fail("BAD_ARGS", "Unknown theme \"" + name + "\": use ink, paper, ocean, forest, graphite, sepia, crimson, slate (light) or night, chalk, plum (dark), and override single colours if you like.");
        }
        return t;
    }

    static Theme ThemeOf(dynamic deck)
    {
        try
        {
            string stored = (string)deck.Tags["DSHTHEME"];
            if (!string.IsNullOrEmpty(stored)) return Json.Deserialize<Theme>(stored);
        }
        catch (Exception) { }
        return Named("ink");
    }

    static string SetTheme(dynamic deck, Bag op)
    {
        Theme t = op.Has("name") ? Named(op.Need("name").ToLowerInvariant()) : ThemeOf(deck);
        if (op.Has("bg")) t.Bg = op.Need("bg");
        if (op.Has("surface")) t.Surface = op.Need("surface");
        if (op.Has("text")) t.Text = op.Need("text");
        if (op.Has("muted")) t.Muted = op.Need("muted");
        if (op.Has("primary")) t.Primary = op.Need("primary");
        if (op.Has("accent")) t.Accent = op.Need("accent");
        if (op.Has("line")) t.Line = op.Need("line");
        if (op.Has("titleFont")) t.TitleFont = op.Need("titleFont");
        if (op.Has("bodyFont")) t.BodyFont = op.Need("bodyFont");
        if (op.Has("dark")) t.Dark = op.Flag("dark", false);
        if (op.Has("transition")) t.Transition = op.Need("transition").ToLowerInvariant();
        if (op.Has("animate")) t.Animate = op.Flag("animate", true);
        if (op.Has("background"))
        {
            string ground = op.Need("background").Trim();
            string kind = ground.ToLowerInvariant();
            if (kind == "plain" || kind == "none" || kind == "flat") t.Texture = "";
            else if (kind == "paper" || kind == "gradient" || kind == "grid" || kind == "dots") t.Texture = kind;
            else if (File.Exists(ground)) t.Texture = "image:" + Path.GetFullPath(ground);
            else throw new Fail("BAD_ARGS", "\"background\" is plain, paper, gradient, grid or dots, or the path of a picture to lay faintly behind every slide; \"" + ground + "\" is neither.");
        }
        if (op.Has("section")) t.Section = op.Need("section").ToLowerInvariant();
        // Decks should not all open their parts the same way: unless told, one of the looks is taken by chance.
        if (string.IsNullOrEmpty(t.Section)) t.Section = new string[] { "solid", "side", "band", "number" }[new Random().Next(4)];
        if (t.Section != "solid" && t.Section != "side" && t.Section != "band" && t.Section != "number") throw new Fail("BAD_ARGS", "\"section\" is how section slides look: solid, side, band or number.");
        if (op.Has("base")) t.Base = op.Int("base", 0);
        if (op.Has("area")) t.Area = Box(op.List("area"));
        foreach (string colour in new string[] { t.Bg, t.Surface, t.Text, t.Muted, t.Primary, t.Accent, t.Line }) Bgr(colour);
        try { deck.Tags.Add("DSHTHEME", Json.Serialize(t)); } catch (Exception) { }
        return "theme \"" + t.Name + "\" set (background: " + (t.Texture.Length == 0 ? "plain" : t.Texture.StartsWith("image:", StringComparison.Ordinal) ? "a picture" : t.Texture) + "; section slides: " + t.Section + "): it applies to the slides made with the slide operation from here on";
    }

    /// The picture that lies behind the content of a theme with a textured ground, made once and kept in the temp folder.
    static string Ground(Theme t)
    {
        if (string.IsNullOrEmpty(t.Texture)) return null;
        string key = t.Texture + "|" + t.Bg + "|" + t.Primary + "|" + t.Line + "|" + t.Dark + "|3";
        if (t.Texture.StartsWith("image:", StringComparison.Ordinal)) { try { key += File.GetLastWriteTimeUtc(t.Texture.Substring(6)).Ticks; } catch (Exception) { } }
        uint hash = 2166136261;
        foreach (char ch in key) { hash ^= ch; hash *= 16777619; }
        bool photo = t.Texture == "paper" || t.Texture.StartsWith("image:", StringComparison.Ordinal);
        string file = Path.Combine(Path.GetTempPath(), "dsh-office-ground-" + hash.ToString("x8") + (photo ? ".jpg" : ".png"));
        if (File.Exists(file)) return file;
        const int W = 1600, H = 900;
        Color bg = ColorTranslator.FromHtml(t.Bg);
        using (Bitmap bitmap = new Bitmap(W, H, PixelFormat.Format24bppRgb))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(bg);
                if (t.Texture == "gradient")
                {
                    // The ground deepens a little towards the lower right, in the theme's own colour.
                    using (System.Drawing.Drawing2D.LinearGradientBrush brush = new System.Drawing.Drawing2D.LinearGradientBrush(new Rectangle(0, 0, W, H), bg, ColorTranslator.FromHtml(Mix(t.Bg, t.Primary, t.Dark ? 0.22 : 0.11)), 35f))
                        g.FillRectangle(brush, 0, 0, W, H);
                    using (SolidBrush glow = new SolidBrush(Color.FromArgb(t.Dark ? 16 : 70, t.Dark ? Color.White : ColorTranslator.FromHtml(Mix(t.Bg, "#FFFFFF", 0.7)))))
                        g.FillEllipse(glow, -W / 4, -H / 2, W, H);
                }
                else if (t.Texture == "grid")
                {
                    using (Pen pen = new Pen(ColorTranslator.FromHtml(Mix(t.Bg, t.Line, 0.55)), 1f))
                    {
                        for (int x = 0; x < W; x += 50) g.DrawLine(pen, x, 0, x, H);
                        for (int y = 0; y < H; y += 50) g.DrawLine(pen, 0, y, W, y);
                    }
                }
                else if (t.Texture == "dots")
                {
                    using (SolidBrush dot = new SolidBrush(ColorTranslator.FromHtml(Mix(t.Bg, t.Line, 0.9))))
                        for (int y = 20; y < H; y += 36) for (int x = 20; x < W; x += 36) g.FillEllipse(dot, x - 2, y - 2, 4, 4);
                }
                else if (t.Texture.StartsWith("image:", StringComparison.Ordinal))
                {
                    // The picture, scaled to cover, under a veil of the ground colour: felt rather than seen.
                    using (Image picture = Image.FromFile(t.Texture.Substring(6)))
                    {
                        double scale = Math.Max((double)W / picture.Width, (double)H / picture.Height);
                        int pw = (int)Math.Ceiling(picture.Width * scale), ph = (int)Math.Ceiling(picture.Height * scale);
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(picture, (W - pw) / 2, (H - ph) / 2, pw, ph);
                    }
                    using (SolidBrush veil = new SolidBrush(Color.FromArgb(t.Dark ? 205 : 226, bg))) g.FillRectangle(veil, 0, 0, W, H);
                }
                if (t.Texture == "paper")
                {
                    // Aged paper: darker towards the edges.
                    using (System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath())
                    {
                        path.AddEllipse(-W * 0.25f, -H * 0.35f, W * 1.5f, H * 1.7f);
                        using (System.Drawing.Drawing2D.PathGradientBrush brush = new System.Drawing.Drawing2D.PathGradientBrush(path))
                        {
                            brush.CenterColor = Color.FromArgb(0, bg);
                            brush.SurroundColors = new Color[] { Color.FromArgb(t.Dark ? 120 : 70, ColorTranslator.FromHtml(t.Dark ? "#000000" : Mix(t.Bg, "#6B4E1E", 0.5))) };
                            brush.FocusScales = new PointF(0.55f, 0.5f);
                            g.FillRectangle(brush, 0, 0, W, H);
                        }
                    }
                }
            }
            if (t.Texture == "paper")
            {
                // And its grain: every dot a touch lighter or darker.
                BitmapData data = bitmap.LockBits(new Rectangle(0, 0, W, H), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
                byte[] bytes = new byte[data.Stride * H];
                Marshal.Copy(data.Scan0, bytes, 0, bytes.Length);
                Random chance = new Random(7);
                for (int y = 0; y < H; y++)
                {
                    int row = y * data.Stride;
                    for (int x = 0; x < W; x++)
                    {
                        int delta = chance.Next(-6, 7);
                        for (int c = 0; c < 3; c++) { int v = bytes[row + x * 3 + c] + delta; bytes[row + x * 3 + c] = (byte)(v < 0 ? 0 : v > 255 ? 255 : v); }
                    }
                }
                Marshal.Copy(bytes, 0, data.Scan0, bytes.Length);
                bitmap.UnlockBits(data);
            }
            if (photo)
            {
                ImageCodecInfo jpeg = null;
                foreach (ImageCodecInfo codec in ImageCodecInfo.GetImageEncoders()) if (codec.MimeType == "image/jpeg") jpeg = codec;
                EncoderParameters quality = new EncoderParameters(1);
                quality.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 88L);
                bitmap.Save(file, jpeg, quality);
            }
            else bitmap.Save(file, ImageFormat.Png);
        }
        return file;
    }

    static double[] Box(IList raw)
    {
        if (raw == null || raw.Count != 4) throw new Fail("BAD_ARGS", "\"area\" must be [left, top, width, height] in points.");
        double[] box = new double[4];
        for (int i = 0; i < 4; i++) box[i] = Number(raw[i]);
        if (box[2] < 100 || box[3] < 80) throw new Fail("BAD_ARGS", "\"area\" is too small: it is [left, top, width, height] in points, the size of the free part of the slide.");
        return box;
    }

    /// What is being drawn on: the slide, its theme, and the scale from the 960 × 540 canvas.
    class Page
    {
        public dynamic Slide;
        public Theme T;
        public float Sx = 1, Sy = 1, S = 1, Ox, Oy;
        public bool Based;   // drawn on a copy of a template slide: its background is left as it is
        public List<object> Motion = new List<object>();   // (markers left by the layouts; the units below are what moves)
        public List<List<object>> Units = new List<List<object>>();
        public List<object> Current;   // the unit being drawn; null while the fixed parts (the head) are drawn
        public int Texts;
    }

    /// From here on, what is drawn belongs to a new item: its shapes will come in together, after the item before.
    static void Unit(Page p)
    {
        p.Current = new List<object>();
        p.Units.Add(p.Current);
    }

    static dynamic Track(Page p, dynamic shape)
    {
        if (p.Current != null && shape != null) p.Current.Add(shape);
        return shape;
    }

    static float PX(Page p, double v) { return p.Ox + (float)(v * p.Sx); }
    static float PY(Page p, double v) { return p.Oy + (float)(v * p.Sy); }
    static float SX(Page p, double v) { return (float)(v * p.Sx); }
    static float SY(Page p, double v) { return (float)(v * p.Sy); }

    static dynamic Block(Page p, double x, double y, double w, double h, string fill, bool rounded)
    {
        dynamic shape = p.Slide.Shapes.AddShape(rounded ? 5 : 1, PX(p, x), PY(p, y), Math.Max(0.5f, SX(p, w)), Math.Max(0.5f, SY(p, h)));
        shape.Line.Visible = 0;
        if (fill == null) shape.Fill.Visible = 0;
        else
        {
            shape.Fill.Visible = -1; shape.Fill.Solid(); shape.Fill.ForeColor.RGB = Bgr(fill);
            // On a template's picture the grounds of cards and bands are white and let a little of it through.
            if (p.Based && fill == p.T.Surface && w > 60 && h > 30) { shape.Fill.ForeColor.RGB = Bgr(p.T.Dark ? "#000000" : "#FFFFFF"); try { shape.Fill.Transparency = p.T.Dark ? 0.45f : 0.14f; } catch (Exception) { } }
        }
        if (rounded) { try { shape.Adjustments[1] = (float)Math.Min(0.5, 8.0 / Math.Max(1, Math.Min(w, h))); } catch (Exception) { } }
        try { shape.Shadow.Visible = 0; } catch (Exception) { }
        return Track(p, shape);
    }

    static dynamic Outline(Page p, double x, double y, double w, double h, string colour, double weight, bool rounded)
    {
        dynamic shape = Block(p, x, y, w, h, null, rounded);
        shape.Line.Visible = -1;
        shape.Line.ForeColor.RGB = Bgr(colour);
        shape.Line.Weight = (float)(weight * p.S);
        return shape;
    }

    static dynamic Rule(Page p, double x1, double y1, double x2, double y2, string colour, double weight)
    {
        dynamic line = p.Slide.Shapes.AddLine(PX(p, x1), PY(p, y1), PX(p, x2), PY(p, y2));
        line.Line.ForeColor.RGB = Bgr(colour);
        line.Line.Weight = (float)(weight * p.S);
        return Track(p, line);
    }

    static dynamic Dot(Page p, double cx, double cy, double r, string fill)
    {
        dynamic shape = p.Slide.Shapes.AddShape(9, PX(p, cx) - SX(p, r), PY(p, cy) - SX(p, r), SX(p, 2 * r), SX(p, 2 * r));
        shape.Line.Visible = 0;
        shape.Fill.Solid(); shape.Fill.ForeColor.RGB = Bgr(fill);
        try { shape.Shadow.Visible = 0; } catch (Exception) { }
        return Track(p, shape);
    }

    /// A piece of text. align: 1 left, 2 centre, 3 right; anchor: 1 top, 3 middle, 4 bottom. **x** inside is set
    /// bold in the emphasis colour, and formulas between dollar signs become equations.
    static dynamic Label(Page p, double x, double y, double w, double h, string text, double size, string colour, bool bold, string font, int align, int anchor, string emphasis)
    {
        dynamic box = p.Slide.Shapes.AddTextbox(1, PX(p, x), PY(p, y), SX(p, w), SY(p, h));
        dynamic frame = box.TextFrame;
        frame.WordWrap = -1;
        try { frame.AutoSize = 0; } catch (Exception) { }
        frame.MarginLeft = 0; frame.MarginRight = 0; frame.MarginTop = 0; frame.MarginBottom = 0;
        frame.VerticalAnchor = anchor;
        box.Height = SY(p, h);
        // Emphasis marks out, their places kept.
        StringBuilder plain = new StringBuilder();
        List<int[]> strong = new List<int[]>();
        string source = (text ?? "").Replace("\r\n", "\r").Replace("\n", "\r");
        int at = 0;
        while (at < source.Length)
        {
            int open = source.IndexOf("**", at, StringComparison.Ordinal);
            int close = open < 0 ? -1 : source.IndexOf("**", open + 2, StringComparison.Ordinal);
            if (open < 0 || close < 0) { plain.Append(source, at, source.Length - at); break; }
            plain.Append(source, at, open - at);
            strong.Add(new int[] { plain.Length, close - open - 2 });
            plain.Append(source, open + 2, close - open - 2);
            at = close + 2;
        }
        dynamic range = frame.TextRange;
        range.Text = plain.ToString();
        range.Font.Size = (float)Math.Max(6, (size < 20 ? size * 1.2 : size) * Math.Max(p.S, 0.84f));
        range.Font.Bold = bold ? -1 : 0;
        range.Font.Color.RGB = Bgr(colour);
        range.Font.Name = font;
        try { range.Font.NameFarEast = font; } catch (COMException) { }
        range.ParagraphFormat.Alignment = align;
        try { range.ParagraphFormat.SpaceWithin = 1.08f; } catch (Exception) { }
        foreach (int[] span in strong)
        {
            if (span[1] <= 0) continue;
            try { dynamic part = range.Characters(span[0] + 1, span[1]); part.Font.Bold = -1; if (emphasis != null) part.Font.Color.RGB = Bgr(emphasis); }
            catch (Exception) { }
        }
        if (plain.ToString().IndexOf('$') >= 0) { try { PptMath(range); } catch (Exception error) { PptMathFailed++; Trace("formulas of a label: " + error.GetType().Name + " " + error.Message); } }
        p.Texts++;
        return Track(p, box);
    }

    static dynamic Words(Page p, double x, double y, double w, double h, string text, double size, string colour, bool bold)
    {
        return Label(p, x, y, w, h, text, size, colour, bold, p.T.BodyFont, 1, 1, p.T.Accent);
    }

    // Icons are small line drawings, written out as SVG in the colour wanted and placed as vector pictures.
    static readonly Dictionary<string, string> Glyphs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        { "check", "<polyline points='20 6 9 17 4 12'/>" },
        { "clock", "<circle cx='12' cy='12' r='10'/><polyline points='12 6 12 12 16 14'/>" },
        { "search", "<circle cx='11' cy='11' r='8'/><line x1='21' y1='21' x2='16.65' y2='16.65'/>" },
        { "star", "<polygon points='12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2'/>" },
        { "chart", "<line x1='18' y1='20' x2='18' y2='10'/><line x1='12' y1='20' x2='12' y2='4'/><line x1='6' y1='20' x2='6' y2='14'/>" },
        { "list", "<line x1='8' y1='6' x2='21' y2='6'/><line x1='8' y1='12' x2='21' y2='12'/><line x1='8' y1='18' x2='21' y2='18'/><line x1='3' y1='6' x2='3.01' y2='6'/><line x1='3' y1='12' x2='3.01' y2='12'/><line x1='3' y1='18' x2='3.01' y2='18'/>" },
        { "people", "<path d='M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2'/><circle cx='9' cy='7' r='4'/><path d='M23 21v-2a4 4 0 0 0-3-3.87'/><path d='M16 3.13a4 4 0 0 1 0 7.75'/>" },
        { "target", "<circle cx='12' cy='12' r='10'/><circle cx='12' cy='12' r='6'/><circle cx='12' cy='12' r='2'/>" },
        { "bolt", "<polygon points='13 2 3 14 12 14 11 22 21 10 12 10 13 2'/>" },
        { "flag", "<path d='M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z'/><line x1='4' y1='22' x2='4' y2='15'/>" },
        { "globe", "<circle cx='12' cy='12' r='10'/><line x1='2' y1='12' x2='22' y2='12'/><path d='M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z'/>" },
        { "lock", "<rect x='3' y='11' width='18' height='11' rx='2' ry='2'/><path d='M7 11V7a5 5 0 0 1 10 0v4'/>" },
        { "book", "<path d='M4 19.5A2.5 2.5 0 0 1 6.5 17H20'/><path d='M6.5 2H20v20H6.5A2.5 2.5 0 0 1 4 19.5v-15A2.5 2.5 0 0 1 6.5 2z'/>" },
        { "calendar", "<rect x='3' y='4' width='18' height='18' rx='2' ry='2'/><line x1='16' y1='2' x2='16' y2='6'/><line x1='8' y1='2' x2='8' y2='6'/><line x1='3' y1='10' x2='21' y2='10'/>" },
        { "document", "<path d='M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z'/><polyline points='14 2 14 8 20 8'/><line x1='16' y1='13' x2='8' y2='13'/><line x1='16' y1='17' x2='8' y2='17'/>" },
        { "layers", "<polygon points='12 2 2 7 12 12 22 7 12 2'/><polyline points='2 17 12 22 22 17'/><polyline points='2 12 12 17 22 12'/>" },
        { "trend", "<polyline points='23 6 13.5 15.5 8.5 10.5 1 18'/><polyline points='17 6 23 6 23 12'/>" },
        { "link", "<path d='M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71'/><path d='M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71'/>" },
        { "code", "<polyline points='16 18 22 12 16 6'/><polyline points='8 6 2 12 8 18'/>" },
        { "heart", "<path d='M20.84 4.61a5.5 5.5 0 0 0-7.78 0L12 5.67l-1.06-1.06a5.5 5.5 0 0 0-7.78 7.78l1.06 1.06L12 21.23l7.78-7.78 1.06-1.06a5.5 5.5 0 0 0 0-7.78z'/>" },
        { "warning", "<path d='M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z'/><line x1='12' y1='9' x2='12' y2='13'/><line x1='12' y1='17' x2='12.01' y2='17'/>" },
        { "home", "<path d='M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z'/><polyline points='9 22 9 12 15 12 15 22'/>" },
        { "idea", "<path d='M9 18h6'/><path d='M10 22h4'/><path d='M12 2a7 7 0 0 0-4 12.7V16h8v-1.3A7 7 0 0 0 12 2z'/>" },
        { "pin", "<path d='M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z'/><circle cx='12' cy='10' r='3'/>" },
        { "money", "<line x1='12' y1='1' x2='12' y2='23'/><path d='M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6'/>" },
        { "arrow", "<line x1='5' y1='12' x2='19' y2='12'/><polyline points='12 5 19 12 12 19'/>" },
        { "cloud", "<path d='M18 10h-1.26A8 8 0 1 0 9 20h9a5 5 0 0 0 0-10z'/>" },
        { "mail", "<path d='M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z'/><polyline points='22,6 12,13 2,6'/>" },
        { "shield", "<path d='M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z'/>" },
        { "eye", "<path d='M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z'/><circle cx='12' cy='12' r='3'/>" },
        { "pie", "<path d='M21.21 15.89A10 10 0 1 1 8 2.83'/><path d='M22 12A10 10 0 0 0 12 2v10z'/>" },
        { "wave", "<polyline points='22 12 18 12 15 21 9 3 6 12 2 12'/>" },
        { "question", "<circle cx='12' cy='12' r='10'/><path d='M9.09 9a3 3 0 0 1 5.83 1c0 2-3 3-3 3'/><line x1='12' y1='17' x2='12.01' y2='17'/>" },
        { "info", "<circle cx='12' cy='12' r='10'/><line x1='12' y1='16' x2='12' y2='12'/><line x1='12' y1='8' x2='12.01' y2='8'/>" },
        { "grid", "<rect x='3' y='3' width='7' height='7'/><rect x='14' y='3' width='7' height='7'/><rect x='14' y='14' width='7' height='7'/><rect x='3' y='14' width='7' height='7'/>" },
        { "edit", "<path d='M12 20h9'/><path d='M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z'/>" },
        { "plus", "<line x1='12' y1='5' x2='12' y2='19'/><line x1='5' y1='12' x2='19' y2='12'/>" },
        { "play", "<polygon points='5 3 19 12 5 21 5 3'/>" },
    };

    static dynamic Icon(Page p, double x, double y, double size, string name, string colour)
    {
        string drawing;
        if (string.IsNullOrEmpty(name) || !Glyphs.TryGetValue(name, out drawing)) return null;
        try
        {
            string folder = Path.Combine(Path.GetTempPath(), "dsh-office", "icons");
            Directory.CreateDirectory(folder);
            string file = Path.Combine(folder, name.ToLowerInvariant() + "-" + colour.Trim('#') + ".svg");
            if (!File.Exists(file))
                File.WriteAllText(file, "<svg xmlns='http://www.w3.org/2000/svg' width='96' height='96' viewBox='0 0 24 24' fill='none' stroke='" + colour + "' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round'>" + drawing + "</svg>", new UTF8Encoding(false));
            dynamic picture = p.Slide.Shapes.AddPicture(file, 0, -1, PX(p, x), PY(p, y), SX(p, size), SX(p, size));
            picture.Name = "Icon " + name;
            return Track(p, picture);
        }
        catch (Exception) { return null; }
    }

    /// A picture that fills the box exactly: scaled to cover it, the overflow cropped away evenly.
    /// Where the part of a cut picture is taken from when nothing is said: "top", "bottom", "left", "right", "center".
    static string Focus = null;
    /// Pictures of this batch that lost much to the cut, for the author to look at.
    static List<string> Cut = new List<string>();

    static string CutNote()
    {
        if (Cut.Count == 0) return "";
        string note = "\nCut to fit: " + string.Join("; ", Cut.ToArray()) + ". Look at these (office_render): if a head or the subject is cut off, give the picture \"focus\": top | center | bottom | left | right, or \"fit\": \"contain\" to show it whole.";
        Cut.Clear();
        return note;
    }

    static dynamic Photo(Page p, double x, double y, double w, double h, string path)
    {
        if (!File.Exists(path)) throw new Fail("BAD_ARGS", "Image \"" + path + "\" does not exist.");
        path = Path.GetFullPath(path);
        float bx = PX(p, x), by = PY(p, y), bw = SX(p, w), bh = SY(p, h);
        dynamic picture = p.Slide.Shapes.AddPicture(path, 0, -1, bx, by, -1, -1);
        float pw = (float)picture.Width, ph = (float)picture.Height;
        float scale = Math.Max(bw / pw, bh / ph);
        try
        {
            dynamic crop = picture.PictureFormat.Crop;
            crop.ShapeWidth = bw; crop.ShapeHeight = bh; crop.ShapeLeft = bx; crop.ShapeTop = by;
            crop.PictureWidth = pw * scale; crop.PictureHeight = ph * scale;
            // What does not fit is cut away. Of a picture taller than its place the upper part is kept, a fifth down
            // from the top: that is where heads are. The author may say otherwise.
            float overX = pw * scale - bw, overY = ph * scale - bh;
            string focus = (Focus ?? "").ToLowerInvariant();
            float shareX = focus == "left" ? 0f : focus == "right" ? 1f : 0.5f;
            float shareY = focus == "top" ? 0f : focus == "bottom" ? 1f : focus == "center" || focus == "centre" || focus == "middle" ? 0.5f : 0.2f;
            crop.PictureOffsetX = overX * (0.5f - shareX);
            crop.PictureOffsetY = overY * (0.5f - shareY);
            double lost = 1 - (bw * bh) / (pw * scale * ph * scale);
            if (lost > 0.4 && Cut.Count < 8) Cut.Add("\"" + Path.GetFileName(path) + "\" shows " + Math.Round((1 - lost) * 100) + "% of itself (" + (overY > overX ? "the " + (shareY < 0.4 ? "upper" : shareY > 0.6 ? "lower" : "middle") + " part of its height" : "the " + (shareX < 0.4 ? "left" : shareX > 0.6 ? "right" : "middle") + " part of its width") + ")");
        }
        catch (Exception)
        {
            picture.LockAspectRatio = 0; picture.Left = bx; picture.Top = by; picture.Width = bw; picture.Height = bh;
        }
        return Track(p, picture);
    }

    /// A drawing or chart on a plain light ground, as opposed to a photograph: its four corners are the same light colour.
    static bool IsSvg(string path) { return path.EndsWith(".svg", StringComparison.OrdinalIgnoreCase); }

    /// Width over height of a picture file; 0 when it cannot be told.
    static double Aspect(string path)
    {
        try
        {
            if (IsSvg(path))
            {
                string text = File.ReadAllText(path);
                int at = text.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
                if (at < 0) return 0;
                string tag = text.Substring(at, Math.Max(0, text.IndexOf('>', at) - at));
                System.Text.RegularExpressions.Match box = System.Text.RegularExpressions.Regex.Match(tag, "viewBox\\s*=\\s*[\"']\\s*[-\\d.]+[\\s,]+[-\\d.]+[\\s,]+([\\d.]+)[\\s,]+([\\d.]+)");
                if (!box.Success) box = System.Text.RegularExpressions.Regex.Match(tag, "width\\s*=\\s*[\"']([\\d.]+)[a-z]*[\"'][^>]*?height\\s*=\\s*[\"']([\\d.]+)");
                if (!box.Success) return 0;
                double bw = double.Parse(box.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture), bh = double.Parse(box.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture);
                return bh > 0 ? bw / bh : 0;
            }
            using (Bitmap bitmap = new Bitmap(path)) return bitmap.Height > 0 ? (double)bitmap.Width / bitmap.Height : 0;
        }
        catch (Exception) { return 0; }
    }

    /// Said once per batch about drawings whose lettering ends up too small to read where they were put.
    static List<string> Small = new List<string>();
    static bool Advised, SmallOnSlide;   // the advice is given once per batch

    /// shown / drawn = how much a drawing was scaled; much under 1, a chart drawn with 10 pt type is no longer readable.
    static void Lettering(string path, double shown, double drawn, bool slide)
    {
        try
        {
            string kindOf = Path.GetExtension(path).ToLowerInvariant();
            if (kindOf == ".jpg" || kindOf == ".jpeg" || kindOf == ".webp") return;   // photographs and scans come as these; drawings as png
            if (drawn <= 0 || shown <= 0 || IsSvg(path) || !IsFigure(path)) return;
            double scale = shown / drawn;
            if (scale >= (slide ? 0.68 : 0.62)) return;
            if (slide) SmallOnSlide = true;
            string name = Path.GetFileName(path);
            foreach (string said in Small) if (said.StartsWith("\"" + name + "\"", StringComparison.Ordinal)) return;
            Small.Add("\"" + name + "\" is shown at " + Math.Round(scale * 100) + "% of the size it was drawn at (10 pt lettering in it reads as " + Math.Round(10 * scale, 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " pt; it is " + Math.Round(shown / 72, 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + " in wide here)");
        }
        catch (Exception) { }
    }

    static string SmallNote()
    {
        if (Small.Count == 0) return "";
        string note = "\nFigures: " + string.Join("; ", Small.ToArray()) + ".";
        if (!Advised)
        {
            note += " Lettering that small is hard to read" + (SmallOnSlide ? " (a slide wants 12 pt or more)" : "") + ": a drawing you made yourself is better drawn again at the width it is shown at, with type sized for that width, and a drawing with several panels side by side split into one per panel" + (SmallOnSlide ? "; on a slide, kind image gives a drawing the whole width" : "") + ".";
            Advised = true;
        }
        Small.Clear();
        return note;
    }

    static bool IsFigure(string path)
    {
        if (IsSvg(path)) return true;
        try
        {
            using (Bitmap bitmap = new Bitmap(path))
            {
                int w = bitmap.Width, h = bitmap.Height;
                Color[] at = { bitmap.GetPixel(2, 2), bitmap.GetPixel(w - 3, 2), bitmap.GetPixel(2, h - 3), bitmap.GetPixel(w - 3, h - 3), bitmap.GetPixel(w / 2, 2), bitmap.GetPixel(2, h / 2) };
                foreach (Color c in at)
                {
                    if (c.A < 250) return true;
                    if (c.R < 232 || c.G < 232 || c.B < 232) return false;
                    if (Math.Abs(c.R - at[0].R) > 6 || Math.Abs(c.G - at[0].G) > 6 || Math.Abs(c.B - at[0].B) > 6) return false;
                }
                return true;
            }
        }
        catch (Exception) { return false; }
    }

    /// A picture shown whole inside the box, on a white card when it is a figure and the deck is not white itself.
    static dynamic Figure(Page p, double x, double y, double w, double h, string path)
    {
        if (!File.Exists(path)) throw new Fail("BAD_ARGS", "Image \"" + path + "\" does not exist.");
        path = Path.GetFullPath(path);
        double pad = 10;
        dynamic card = Block(p, x, y, w, h, "#FFFFFF", true);
        float bx = PX(p, x + pad), by = PY(p, y + pad), bw = SX(p, w - 2 * pad), bh = SY(p, h - 2 * pad);
        dynamic picture = p.Slide.Shapes.AddPicture(path, 0, -1, bx, by, -1, -1);
        picture.LockAspectRatio = -1;
        float pw = (float)picture.Width, ph = (float)picture.Height, scale = Math.Min(bw / pw, bh / ph);
        Lettering(path, pw * scale, pw, true);
        picture.Width = pw * scale;
        picture.Left = bx + (bw - pw * scale) / 2;
        picture.Top = by + (bh - ph * scale) / 2;
        return Track(p, picture);
    }

    /// A veil over a picture so that text on it reads: the colour, strong at one side and fading to the other.
    /// direction 1 fades left to right, 2 top to bottom. It is a picture a few dots wide, stretched over the place.
    static void Veil(Page p, double x, double y, double w, double h, string colour, int direction, double strong, double weak)
    {
        if (Math.Abs(strong - weak) < 0.02)
        {
            dynamic flat = Block(p, x, y, w, h, colour, false);
            try { flat.Fill.Transparency = (float)(1 - strong); } catch (Exception) { }
            return;
        }
        string file = Path.Combine(Path.GetTempPath(), "dsh-office-veil-" + Guid.NewGuid().ToString("N") + ".png");
        try
        {
            Color tone = ColorTranslator.FromHtml(colour);
            const int steps = 256;
            using (Bitmap bitmap = new Bitmap(direction == 1 ? steps : 2, direction == 1 ? 2 : steps, PixelFormat.Format32bppArgb))
            {
                for (int i = 0; i < steps; i++)
                {
                    double share = i / (double)(steps - 1), eased = share * share * (3 - 2 * share);
                    int alpha = (int)Math.Round(255 * Math.Max(0, Math.Min(1, strong + (weak - strong) * eased)));
                    Color dot = Color.FromArgb(alpha, tone.R, tone.G, tone.B);
                    if (direction == 1) { bitmap.SetPixel(i, 0, dot); bitmap.SetPixel(i, 1, dot); }
                    else { bitmap.SetPixel(0, i, dot); bitmap.SetPixel(1, i, dot); }
                }
                bitmap.Save(file, ImageFormat.Png);
            }
            dynamic picture = p.Slide.Shapes.AddPicture(file, 0, -1, PX(p, x), PY(p, y), SX(p, w), SY(p, h));
            picture.Name = "Veil";
            Track(p, picture);
        }
        finally { try { File.Delete(file); } catch (Exception) { } }
    }

    static string Field(object raw, string key)
    {
        Dictionary<string, object> map = raw as Dictionary<string, object>;
        if (map == null) return key == "text" ? Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture) : null;
        object value;
        return map.TryGetValue(key, out value) && value != null ? Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) : null;
    }

    static IList Items(Bag op, params string[] keys)
    {
        foreach (string key in keys) { IList list = op.List(key); if (list != null && list.Count > 0) return list; }
        return new ArrayList();
    }

    static string Pick(Page p, int index) { return index % 2 == 0 ? p.T.Primary : p.T.Accent; }

    /// The top of an ordinary content slide: a small dash, the section line, the page number, the headline and what
    /// stands under it. Returns where the body may begin.
    static double Head(Page p, Bag op)
    {
        Theme t = p.T;
        if (p.Based)
        {
            // A template page has its own headline, in its own place and type: the title goes there.
            dynamic own = null;
            try { if (Truthy(p.Slide.Shapes.HasTitle)) own = p.Slide.Shapes.Title; } catch (Exception) { }
            if (own != null)
            {
                try { own.TextFrame.TextRange.Text = op.Str("title", ""); } catch (Exception) { }
                op.Str("kicker", null);
                double start = 4;
                string under = op.Str("subtitle", null);
                if (!string.IsNullOrEmpty(under)) { Label(p, 48, 6, 864, 30, under, 16, t.Text, false, t.BodyFont, 1, 1, t.Accent).Name = "Subtitle"; start = 48; }
                string source = op.Str("note", null);
                if (!string.IsNullOrEmpty(source)) Label(p, 48, 512, 864, 16, source, 10, t.Muted, false, t.BodyFont, 1, 1, null).Name = "Note";
                return start;
            }
        }
        Block(p, 48, 40, 26, 3, t.Accent, false);
        string kicker = op.Str("kicker", null);
        if (!string.IsNullOrEmpty(kicker)) Label(p, 48, 50, 640, 16, kicker, 10.5, t.Muted, false, t.BodyFont, 1, 1, null).Name = "Kicker";
        dynamic number = Label(p, 792, 50, 120, 16, "", 10, t.Muted, false, t.BodyFont, 3, 1, null);
        number.Name = "PageNumber";
        string title = op.Str("title", "");
        if (title.Trim().Length == 0 && string.IsNullOrEmpty(op.Str("subtitle", null)))
        {
            // No headline: the content begins right under the dash, instead of under an empty band.
            string foot = op.Str("note", null);
            if (!string.IsNullOrEmpty(foot)) Label(p, 48, 506, 864, 14, foot, 9, t.Muted, false, t.BodyFont, 1, 1, null).Name = "Note";
            return 72;
        }
        dynamic head = Label(p, 48, 70, 864, 40, title, title.Length > 26 ? 23 : 27, t.Primary, true, t.TitleFont, 1, 3, t.Accent);
        head.Name = "Title";
        double y = 114;
        string subtitle = op.Str("subtitle", null);
        if (!string.IsNullOrEmpty(subtitle)) { Label(p, 48, 113, 864, 26, subtitle, 14.5, t.Text, false, t.BodyFont, 1, 1, t.Accent).Name = "Subtitle"; y = 146; }
        Rule(p, 48, y + 4, 912, y + 4, t.Line, 0.75);
        string note = op.Str("note", null);
        if (!string.IsNullOrEmpty(note)) Label(p, 48, 506, 864, 14, note, 9, t.Muted, false, t.BodyFont, 1, 1, null).Name = "Note";
        return y + 18;
    }

    /// The sample content of a template page that lies in the area about to be drawn on: texts and pictures whose
    /// middle is inside it are removed (the page's title is kept, and so is everything outside: logo, navigation).
    static void Clear(Page p, double[] area)
    {
        dynamic title = null;
        try { if (Truthy(p.Slide.Shapes.HasTitle)) title = p.Slide.Shapes.Title; } catch (Exception) { }
        int titleId = 0;
        try { if (title != null) titleId = (int)title.Id; } catch (Exception) { }
        List<object> doomed = new List<object>();
        foreach (dynamic shape in Leaves(p.Slide))
        {
            try
            {
                if ((int)shape.Id == titleId) continue;
                double cx = (double)shape.Left + (double)shape.Width / 2, cy = (double)shape.Top + (double)shape.Height / 2;
                if (cx < area[0] || cx > area[0] + area[2] || cy < area[1] || cy > area[1] + area[3]) continue;
                // A shape as large as the area itself is the page's ground, not its sample content.
                if ((double)shape.Width > area[2] * 0.95 && (double)shape.Height > area[3] * 0.95) continue;
                doomed.Add(shape);
            }
            catch (Exception) { }
        }
        foreach (dynamic shape in doomed) { try { shape.Delete(); } catch (Exception) { } }
    }

    const double BodyBottom = 494;

    /// Where content of a given height starts inside a taller body: a third of the spare room goes above it.
    static double Settle(double y, double h, double used)
    {
        double spare = h - used;
        return spare > 36 ? y + Math.Min(spare / 3, 70) : y;
    }

    static string DesignSlide(dynamic deck, Bag op)
    {
        string kind = op.Str("kind", "bullets").ToLowerInvariant();
        int total = (int)deck.Slides.Count, index = op.Int("at", total + 1);
        if (index < 1 || index > total + 1) throw new Fail("BAD_ARGS", "\"at\" must be between 1 and " + (total + 1) + ".");
        Page p = new Page();
        p.T = ThemeOf(deck);
        Theme t = p.T;
        int from = op.Int("base", t.Base);
        if (from > 0)
        {
            // Drawn on a copy of a slide of the deck: a template's background page keeps its pictures and motion.
            if (from > total) throw new Fail("ANCHOR_MISSING", "\"base\" is slide " + from + ", but the deck has " + total + " slides.");
            deck.Slides[from].Duplicate();
            deck.Slides[from + 1].MoveTo(index);
            p.Slide = deck.Slides[index];
            p.Based = true;
        }
        else p.Slide = deck.Slides.Add(index, 12);
        GoTo(deck, index);
        double[] area = op.Has("area") ? Box(op.List("area")) : t.Area;
        float canvasW = (float)deck.PageSetup.SlideWidth, canvasH = (float)deck.PageSetup.SlideHeight;
        if (area != null && area.Length == 4) { p.Ox = (float)area[0]; p.Oy = (float)area[1]; canvasW = (float)area[2]; canvasH = (float)area[3]; }
        p.Sx = canvasW / 960f;
        p.Sy = canvasH / 540f;
        p.S = Math.Min(p.Sx, p.Sy);
        if (p.Based && area != null && area.Length == 4) Clear(p, area);
        if (!p.Based)
        {
            p.Slide.FollowMasterBackground = 0;
            p.Slide.Background.Fill.Visible = -1;
            p.Slide.Background.Fill.Solid();
            p.Slide.Background.Fill.ForeColor.RGB = Bgr(t.Bg);
            try { string ground = Ground(t); if (ground != null) p.Slide.Background.Fill.UserPicture(ground); }
            catch (Exception error) { Trace("ground: " + error.Message); }
        }
        try { p.Slide.Tags.Add("DSHKIND", kind); } catch (Exception) { }

        string image = op.Str("image", null);
        Focus = op.Str("focus", null);
        if (kind == "cover" || kind == "closing")
        {
            bool cover = kind == "cover";
            bool full = image != null && !IsFigure(image) && op.Str("style", "") == "full";
            string titleInk = full ? "#FFFFFF" : t.Primary, bodyInk = full ? "#FFFFFF" : t.Text, softInk = full ? "#E6E6E6" : t.Muted;
            if (full)
            {
                // The picture is the slide; a dark veil, heavier at the left where the words stand, lets them read.
                Photo(p, 0, 0, 960, 540, image);
                Veil(p, 0, 0, 960, 540, "#000000", 1, 0.62, 0.05);
            }
            else if (image != null && IsFigure(image)) Figure(p, 500, 70, 412, 400, image);
            else if (image != null)
            {
                Photo(p, 440, 0, 520, 540, image);
                Veil(p, 440, 0, 200, 540, t.Bg, 1, 1, 0);
            }
            else if (!p.Based)
            {
                dynamic big = Dot(p, 800, 130, 250, t.Surface);
                dynamic small = Dot(p, 905, 430, 120, t.Primary);
                try { small.Fill.Transparency = t.Dark ? 0.75f : 0.88f; } catch (Exception) { }
                dynamic ring = p.Slide.Shapes.AddShape(9, PX(p, 610), PY(p, 300), SX(p, 150), SX(p, 150));
                ring.Fill.Visible = 0; ring.Line.ForeColor.RGB = Bgr(t.Accent); ring.Line.Weight = 1.5f * p.S;
            }
            string kicker = op.Str("kicker", null);
            Block(p, 56, cover ? 132 : 190, 64, 4, t.Accent, false);
            if (!string.IsNullOrEmpty(kicker)) Label(p, 56, cover ? 104 : 162, 420, 18, kicker, 12, softInk, false, t.BodyFont, 1, 1, null).Name = "Kicker";
            string title = op.Str("title", "");
            Unit(p);
            dynamic head = Label(p, 56, cover ? 152 : 210, image != null && !full ? 400 : 560, 130, title, title.Length > 18 ? 30 : title.Length > 9 ? 36 : 44, titleInk, true, t.TitleFont, 1, 1, full ? null : t.Accent);
            head.Name = "Title";
            p.Motion.Add(head);
            string subtitle = op.Str("subtitle", null);
            Unit(p);
            if (!string.IsNullOrEmpty(subtitle)) { dynamic sub = Label(p, 56, cover ? 290 : 350, image != null && !full ? 372 : 560, 70, subtitle, 15, bodyInk, false, t.BodyFont, 1, 1, full ? null : t.Accent); sub.Name = "Subtitle"; p.Motion.Add(sub); }
            Dictionary<string, object> figure = op.Raw("stat") as Dictionary<string, object>;
            if (figure != null)
            {
                string value = Field(figure, "value") ?? "", unit = Field(figure, "unit");
                dynamic big = Label(p, 56, 352, 400, 56, value + (unit == null ? "" : " " + unit), 40, t.Accent, true, t.TitleFont, 1, 3, null);
                if (unit != null) { try { dynamic tail = big.TextFrame.TextRange.Characters(value.Length + 1, unit.Length + 1); tail.Font.Size = (float)(16 * p.S); tail.Font.Bold = 0; tail.Font.Color.RGB = Bgr(t.Text); } catch (Exception) { } }
                Label(p, 56, 408, 400, 18, Field(figure, "label") ?? "", 11, t.Muted, false, t.BodyFont, 1, 1, null);
                p.Motion.Add(big);
            }
            string meta = op.Str("meta", null);
            if (!string.IsNullOrEmpty(meta)) { Rule(p, 56, 446, 300, 446, full ? "#FFFFFF" : t.Line, 0.75); Label(p, 56, 456, 420, 56, meta, 11.5, softInk, false, t.BodyFont, 1, 1, null).Name = "Meta"; }
        }
        else if (kind == "section")
        {
            string ground = t.Dark ? t.Surface : t.Primary, ink = t.Dark ? t.Text : "#FFFFFF", soft = t.Dark ? t.Primary : "#FFFFFF";
            string look = op.Str("style", string.IsNullOrEmpty(t.Section) ? "solid" : t.Section).ToLowerInvariant();
            if (!p.Based && image == null && look != "solid")
            {
                string digits = op.Str("number", null), name = op.Str("title", ""), under = op.Str("subtitle", null);
                string strong = t.Dark ? t.Primary : t.Primary, onGround = t.Dark ? t.Text : "#FFFFFF";
                dynamic head2;
                if (look == "side")
                {
                    // A panel at the left carries the number; the words stand on the plain ground beside it.
                    Block(p, 0, 0, 320, 540, ground, false);
                    if (!string.IsNullOrEmpty(digits)) Label(p, 0, 170, 320, 200, digits, 130, onGround, true, t.TitleFont, 2, 3, null).Name = "Number";
                    Block(p, 384, 214, 56, 4, t.Accent, false);
                    Unit(p);
                    head2 = Label(p, 384, 232, 520, 64, name, name.Length > 12 ? 32 : 40, strong, true, t.TitleFont, 1, 1, null);
                    Unit(p);
                    if (!string.IsNullOrEmpty(under)) { dynamic sub2 = Label(p, 384, 308, 520, 80, under, 15, t.Text, false, t.BodyFont, 1, 1, t.Accent); sub2.Name = "Subtitle"; p.Motion.Add(sub2); }
                }
                else if (look == "band")
                {
                    // A band across the slide holds the title; the number stands above it, the line under it.
                    if (!string.IsNullOrEmpty(digits)) Label(p, 56, 92, 400, 100, "PART " + digits, 20, t.Accent, true, t.TitleFont, 1, 4, null).Name = "Number";
                    Block(p, 0, 204, 960, 132, ground, false);
                    Block(p, 0, 204, 12, 132, t.Accent, false);
                    Unit(p);
                    head2 = Label(p, 56, 204, 848, 132, name, name.Length > 14 ? 34 : 42, onGround, true, t.TitleFont, 1, 3, null);
                    Unit(p);
                    if (!string.IsNullOrEmpty(under)) { dynamic sub2 = Label(p, 56, 356, 848, 60, under, 15.5, t.Text, false, t.BodyFont, 1, 1, t.Accent); sub2.Name = "Subtitle"; p.Motion.Add(sub2); }
                }
                else
                {
                    // The number, very large and faint, fills the right; the words stand at the left over a thin line.
                    if (!string.IsNullOrEmpty(digits))
                    {
                        // The line of a text reaches a little past its last letter; at this size that is 80 points,
                        // so the box ends well inside the slide. Anything past the edge makes PowerPoint show the slide off centre.
                        dynamic huge = Label(p, 380, 60, 480, 420, digits, 280, strong, true, t.TitleFont, 3, 3, null);
                        huge.Name = "Number";
                        try { huge.TextFrame2.TextRange.Font.Fill.Transparency = t.Dark ? 0.8f : 0.88f; } catch (Exception) { }
                    }
                    Block(p, 56, 226, 56, 4, t.Accent, false);
                    Unit(p);
                    head2 = Label(p, 56, 244, 620, 64, name, name.Length > 12 ? 34 : 42, strong, true, t.TitleFont, 1, 1, null);
                    Rule(p, 56, 330, 520, 330, t.Line, 0.75);
                    Unit(p);
                    if (!string.IsNullOrEmpty(under)) { dynamic sub2 = Label(p, 56, 344, 560, 80, under, 15, t.Muted, false, t.BodyFont, 1, 1, t.Accent); sub2.Name = "Subtitle"; p.Motion.Add(sub2); }
                }
                head2.Name = "Title";
                p.Motion.Add(head2);
                if (op.Has("notes")) { try { p.Slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = Lines(op.Raw("notes")); } catch (Exception) { } }
                Move(p);
                Renumber = true;
                return "slide " + index + " added (section, theme " + t.Name + ")";
            }
            if (p.Based) { ink = t.Primary; soft = t.Accent; }
            else { p.Slide.Background.Fill.Solid(); p.Slide.Background.Fill.ForeColor.RGB = Bgr(ground); }
            if (image != null) { Photo(p, 0, 0, 960, 540, image); Veil(p, 0, 0, 960, 540, ground, 1, 0.92, 0.55); }
            string number = op.Str("number", null);
            if (!string.IsNullOrEmpty(number))
            {
                dynamic big = Label(p, 56, 120, 360, 170, number, 120, soft, true, t.TitleFont, 1, 1, null);
                try { big.TextFrame.TextRange.Font.Color.RGB = Bgr(soft); big.TextFrame2.TextRange.Font.Fill.Transparency = 0.78f; } catch (Exception) { }
            }
            Block(p, 56, 300, 56, 4, t.Dark || p.Based ? t.Accent : "#FFFFFF", false);
            Unit(p);
            dynamic head = Label(p, 56, 318, 820, 60, op.Str("title", ""), 36, ink, true, t.TitleFont, 1, 1, null);
            head.Name = "Title";
            p.Motion.Add(head);
            string subtitle = op.Str("subtitle", null);
            Unit(p);
            if (!string.IsNullOrEmpty(subtitle)) { dynamic sub = Label(p, 56, 384, 820, 60, subtitle, 16, ink, false, t.BodyFont, 1, 1, null); sub.Name = "Subtitle"; p.Motion.Add(sub); }
        }
        else if (kind == "image")
        {
            if (image == null) throw new Fail("BAD_ARGS", "An image slide needs \"image\": the path of a picture.");
            if (IsFigure(image))
            {
                // A chart or diagram is not cropped to fill the slide: it is shown whole, as large as the slide lets it be.
                string words = op.Str("text", null), under = op.Str("caption", null);
                double top = Head(p, op), foot = (string.IsNullOrEmpty(words) ? 0 : 46) + (string.IsNullOrEmpty(under) ? 0 : 20);
                p.Current = null;
                Unit(p);
                p.Motion.Add(Figure(p, 48, top, 864, BodyBottom - top - foot, image));
                double fy = BodyBottom - foot;
                if (!string.IsNullOrEmpty(under)) { Label(p, 48, fy + 2, 864, 16, under, 10.5, t.Muted, false, t.BodyFont, 2, 1, null).Name = "Caption"; fy += 20; }
                if (!string.IsNullOrEmpty(words)) { Unit(p); Label(p, 48, fy + 4, 864, 40, words, 15, t.Text, false, t.BodyFont, 2, 3, t.Accent).Name = "Text"; }
                if (op.Has("notes")) { try { p.Slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = Lines(op.Raw("notes")); } catch (Exception) { } }
                Move(p);
                Renumber = true;
                return "slide " + index + " added (image, theme " + t.Name + ")" + SmallNote();
            }
            Photo(p, 0, 0, 960, 540, image);
            Veil(p, 0, 230, 960, 310, "#000000", 2, 0, 0.82);
            dynamic head = Label(p, 56, 392, 848, 48, op.Str("title", ""), 30, "#FFFFFF", true, t.TitleFont, 1, 4, null);
            head.Name = "Title";
            string text = op.Str("text", op.Str("subtitle", null));
            if (!string.IsNullOrEmpty(text)) Label(p, 56, 446, 848, 44, text, 14, "#FFFFFF", false, t.BodyFont, 1, 1, null).Name = "Subtitle";
            string caption = op.Str("caption", op.Str("note", null));
            if (!string.IsNullOrEmpty(caption)) Label(p, 480, 512, 424, 14, caption, 8.5, "#DDDDDD", false, t.BodyFont, 3, 1, null).Name = "Note";
            string kicker = op.Str("kicker", null);
            if (!string.IsNullOrEmpty(kicker)) Label(p, 56, 40, 600, 16, kicker, 10.5, "#FFFFFF", false, t.BodyFont, 1, 1, null).Name = "Kicker";
        }
        else if (kind == "split")
        {
            bool right = op.Str("side", "right") != "left";
            double ix = right ? 480 : 0, tx = right ? 56 : 536, tw = 368;
            if (image != null && IsFigure(image)) { Block(p, ix, 0, 480, 540, "#FFFFFF", false); Figure(p, ix + 20, 50, 440, 440, image); }
            else if (image != null) Photo(p, ix, 0, 480, 540, image);
            else Block(p, ix, 0, 480, 540, t.Dark ? t.Surface : t.Primary, false);
            string kicker = op.Str("kicker", null);
            if (!string.IsNullOrEmpty(kicker)) Label(p, tx, 76, tw, 18, kicker, 11, t.Muted, false, t.BodyFont, 1, 1, null).Name = "Kicker";
            Block(p, tx, 104, 56, 4, t.Accent, false);
            string title = op.Str("title", "");
            Unit(p);
            dynamic head = Label(p, tx, 120, tw, 112, title, title.Length > 14 ? 26 : 32, t.Primary, true, t.TitleFont, 1, 1, t.Accent);
            head.Name = "Title";
            p.Motion.Add(head);
            double by = title.Length > 22 ? 256 : 222;
            string words = op.Str("text", op.Str("subtitle", null));
            if (!string.IsNullOrEmpty(words)) { Unit(p); dynamic body = Label(p, tx, by, tw, 96, words, 13, t.Text, false, t.BodyFont, 1, 1, t.Accent); body.Name = "Text"; p.Motion.Add(body); by += 104; }
            IList points = Items(op, "points", "items");
            int count = Math.Min(points.Count, 4);
            double row = count == 0 ? 0 : Math.Min(60, (500 - by) / count);
            for (int i = 0; i < count; i++)
            {
                Unit(p);
                string step = Field(points[i], "head"), text = Field(points[i], "text") ?? "";
                Block(p, tx, by + 9, 6, 6, Pick(p, i), false);
                p.Motion.Add(Words(p, tx + 18, by, tw - 18, row - 6, (step != null ? "**" + step + "**　" : "") + text, 12.5, t.Text, false));
                by += row;
            }
            string note = op.Str("note", op.Str("caption", null));
            if (!string.IsNullOrEmpty(note)) Label(p, tx, 508, tw, 14, note, 8.5, t.Muted, false, t.BodyFont, 1, 1, null).Name = "Note";
        }
        else if (kind == "quote")
        {
            if (image != null) { Photo(p, 0, 0, 960, 540, image); Veil(p, 0, 0, 960, 540, t.Bg, 1, 0.9, 0.8); }
            Label(p, 96, 108, 120, 110, "“", 110, t.Accent, true, t.TitleFont, 1, 1, null);
            Unit(p);
            dynamic words = Label(p, 120, 190, 720, 190, op.Str("text", op.Str("title", "")), 28, t.Text, false, t.TitleFont, 1, 3, t.Primary);
            words.Name = "Title";
            p.Motion.Add(words);
            string by = op.Str("by", null);
            if (!string.IsNullOrEmpty(by)) { Block(p, 120, 400, 40, 3, t.Accent, false); Label(p, 120, 412, 720, 24, by, 13, t.Muted, false, t.BodyFont, 1, 1, null).Name = "By"; }
        }
        else
        {
            double top = Head(p, op), bottom = BodyBottom, left = 48, width = 864;
            p.Current = null;
            // A picture beside the content: the right third of the body.
            if (image != null && kind != "chart" && kind != "canvas" && kind != "custom" && kind != "diagram" && kind != "code" && kind != "gallery" && kind != "figures")
            {
                Unit(p);
                bool figure = IsFigure(image), list = kind == "bullets" || kind == "agenda";
                double aspect = Aspect(image), body = bottom - top;
                if (figure && list && aspect >= 1.9)
                {
                    // A wide drawing would be a thin strip beside the text: it goes across the slide, the points in a row under it.
                    IList listed = Items(op, "points", "items", "bullets");
                    bool headed = false;
                    foreach (object raw in listed) if (Field(raw, "head") != null) headed = true;
                    double under = (listed.Count == 0 ? 0 : headed ? 104 : 66) + (string.IsNullOrEmpty(op.Str("callout", null)) ? 0 : 54);
                    double fh = Math.Min(864 / aspect + 20, body - under - 12);
                    // What is left over goes above and below, so the slide does not end in an empty band.
                    double spare = Math.Max(0, body - under - 14 - fh);
                    top += Math.Min(spare * 0.3, 30);
                    p.Motion.Add(Figure(p, left, top, 864, fh, image));
                    top += fh + 14 + Math.Min(spare * 0.2, 20);
                }
                else if (figure)
                {
                    double fw = list ? Math.Max(352, Math.Min(480, body * (aspect > 0 ? aspect : 1) + 20)) : 352;
                    p.Motion.Add(Figure(p, 912 - fw, top, fw, body, image));
                    width = 864 - fw - 22;
                }
                else
                {
                    p.Motion.Add(Photo(p, 612, top, 300, body, image));
                    width = 540;
                }
            }
            if (kind == "bullets" || kind == "agenda") Bullets(p, op, left, top, width, bottom - top, kind == "agenda");
            else if (kind == "cards") Cards(p, op, left, top, width, bottom - top);
            else if (kind == "stats") Stats(p, op, left, top, width, bottom - top);
            else if (kind == "chart") ChartSlide(p, op, left, top, width, bottom - top);
            else if (kind == "table") TableSlide(p, op, left, top, width, bottom - top);
            else if (kind == "formula") Formulas(p, op, left, top, width, bottom - top);
            else if (kind == "timeline") Timeline(p, op, left, top, width, bottom - top);
            else if (kind == "compare") Compare(p, op, left, top, width, bottom - top);
            else if (kind == "process") Process(p, op, left, top, width, bottom - top);
            else if (kind == "theorem" || kind == "definition" || kind == "lemma") Theorem(p, op, left, top, width, bottom - top);
            else if (kind == "gallery" || kind == "figures") Gallery(p, op, left, top, width, bottom - top);
            else if (kind == "summary" || kind == "takeaways") Summary(p, op, left, top, width, bottom - top);
            else if (kind == "canvas" || kind == "custom" || kind == "diagram") Canvas(p, op);
            else if (kind == "code") Code(p, op, left, top, width, bottom - top);
            else throw new Fail("BAD_ARGS", "Unknown slide kind \"" + kind + "\": use cover, section, agenda, bullets, cards, stats, chart, table, formula, theorem, gallery, summary, code, timeline, compare, process, split, image, quote, canvas or closing.");
        }
        if (op.Has("notes")) { try { p.Slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = Lines(op.Raw("notes")); } catch (Exception) { } }
        Move(p);
        Renumber = true;
        return "slide " + index + " added (" + kind + ", theme " + t.Name + ")" + SmallNote() + CutNote();
    }

    /// The slide comes in with the deck's transition, and its parts follow one another in.
    static void Move(Page p)
    {
        try
        {
            string way = p.T.Transition;
            if (way != "none")
            {
                dynamic change = p.Slide.SlideShowTransition;
                change.EntryEffect = way == "push" ? 3852 : way == "wipe" ? 2817 : way == "split" ? 3848 : 3849;
                try { change.Duration = 0.6f; } catch (Exception) { }
            }
        }
        catch (Exception) { }
        if (!p.T.Animate) return;
        try
        {
            dynamic sequence = p.Slide.TimeLine.MainSequence;
            int items = 0, effects = 0;
            foreach (List<object> unit in p.Units)
            {
                if (unit.Count == 0) continue;
                bool first = true;
                foreach (dynamic shape in unit)
                {
                    // The first shape of an item follows the item before it; the rest of the item comes with it.
                    dynamic effect = sequence.AddEffect(shape, 10, 0, first ? (items == 0 ? 2 : 3) : 2);
                    effect.Timing.Duration = 0.35f;
                    first = false;
                    if (++effects >= 160) return;
                }
                items++;
            }
        }
        catch (Exception) { }
    }

    /// Set when designed slides were added or removed: the page numbers are written at the end of the batch.
    static bool Renumber;

    static void Numbers(dynamic deck)
    {
        Renumber = false;
        try
        {
            int total = (int)deck.Slides.Count;
            foreach (dynamic slide in deck.Slides)
            {
                foreach (dynamic shape in slide.Shapes)
                {
                    if ((string)shape.Name != "PageNumber") continue;
                    int n = (int)slide.SlideIndex;
                    shape.TextFrame.TextRange.Text = n.ToString("00") + " / " + total.ToString("00");
                }
            }
        }
        catch (Exception) { }
    }

    static void Bullets(Page p, Bag op, double x, double y, double w, double h, bool numbered)
    {
        Theme t = p.T;
        IList points = Items(op, "points", "items", "bullets");
        int n = Math.Max(1, Math.Min(points.Count, 8));
        bool heads = false;
        foreach (object raw in points) if (Field(raw, "head") != null) heads = true;
        bool columns = numbered && n > 4;
        int perColumn = columns ? (n + 1) / 2 : n;
        bool foot = !string.IsNullOrEmpty(op.Str("callout", null));
        double row = Math.Min(heads ? 108 : 78, (h - (foot ? 54 : 0)) / perColumn), colWidth = columns ? (w - 32) / 2 : w;
        double start = Settle(y, h - (foot ? 54 : 0), row * perColumn);
        if (!numbered && n <= 5 && (h - (foot ? 54 : 0)) / n < (heads ? 66 : 40))
        {
            // Too low for a column of points (a wide picture stands above): they go side by side.
            double gap = 22, cw = (w - gap * (n - 1)) / n, tall = h - (foot ? 54 : 0);
            for (int i = 0; i < n; i++)
            {
                Unit(p);
                object raw = points[i];
                string head = Field(raw, "head"), text = Field(raw, "text") ?? "";
                double cx = x + i * (cw + gap);
                Block(p, cx, y + 2, 28, 3, Pick(p, i), false);
                if (head != null)
                {
                    // One box for both, so that a head that takes two lines pushes its text down.
                    dynamic both = Words(p, cx, y + 10, cw, tall - 12, head + "\r" + text, 12, t.Muted, false);
                    try
                    {
                        dynamic first = both.TextFrame.TextRange.Paragraphs(1);
                        first.Font.Size = (float)(17 * p.S); first.Font.Bold = -1; first.Font.Color.RGB = Bgr(t.Text);
                        first.ParagraphFormat.SpaceAfter = 4;
                    }
                    catch (Exception) { }
                    p.Motion.Add(both);
                }
                else p.Motion.Add(Words(p, cx, y + 10, cw, tall - 12, text, 13, t.Text, false));
            }
            Callout(p, op, x, y + h - 46, w);
            return;
        }
        for (int i = 0; i < n; i++)
        {
            Unit(p);

            object raw = points[i];
            string head = Field(raw, "head"), text = Field(raw, "text") ?? "";
            double cx = x + (columns && i >= perColumn ? colWidth + 32 : 0), cy = start + (columns ? i % perColumn : i) * row;
            double inset = numbered ? 52 : 22;
            dynamic mark;
            if (numbered) mark = Label(p, cx, cy, 44, row - 10, (i + 1).ToString("00"), 26, Pick(p, i), true, t.TitleFont, 1, 3, null);
            else mark = Block(p, cx, cy + (head != null ? 10 : 9), 8, 8, Pick(p, i), false);
            dynamic body;
            if (head != null)
            {
                dynamic title = Words(p, cx + inset, cy, colWidth - inset, 28, head, 16.5, t.Text, true);
                body = Words(p, cx + inset, cy + 32, colWidth - inset, row - 38, text, 13.5, t.Muted, false);
                p.Motion.Add(title);
            }
            else
            {
                body = Label(p, cx + inset, cy, colWidth - inset, row - 10, text, numbered ? 16.5 : 15.5, t.Text, false, t.BodyFont, 1, numbered ? 3 : 1, t.Accent);
                p.Motion.Add(body);
            }
            if (i < n - 1 && !columns) Rule(p, cx + inset, cy + row - 6, cx + colWidth, cy + row - 6, t.Line, 0.5);
        }
        Callout(p, op, x, y + h - 46, w);
    }

    /// A remark set apart at the foot of the body, when the slide has one.
    static void Callout(Page p, Bag op, double x, double y, double w)
    {
        string text = op.Str("callout", null);
        if (string.IsNullOrEmpty(text)) return;
        Theme t = p.T;
        Unit(p);
        dynamic back = Block(p, x, y, w, 40, t.Surface, false);
        Block(p, x, y, 4, 40, t.Accent, false);
        dynamic words = Label(p, x + 18, y, w - 30, 40, text, 13.5, t.Text, false, t.BodyFont, 1, 3, t.Accent);
        words.Name = "Callout";
        p.Motion.Add(words);
    }

    static void Cards(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        IList cards = Items(op, "cards", "items", "points");
        int n = Math.Max(1, Math.Min(cards.Count, 6));
        int perRow = n <= 4 ? n : 3, rows = (n + perRow - 1) / perRow;
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        double usable = h - (callout ? 54 : 0), gap = 18;
        double cw = (w - gap * (perRow - 1)) / perRow, ch = Math.Min(rows == 1 ? (perRow >= 4 ? 250 : 230) : 160, (usable - gap * (rows - 1)) / rows);
        y = Settle(y, usable, rows * ch + gap * (rows - 1));
        for (int i = 0; i < n; i++)
        {
            Unit(p);

            object raw = cards[i];
            double cx = x + (i % perRow) * (cw + gap), cy = y + (i / perRow) * (ch + gap);
            string tone = Pick(p, i);
            dynamic back = Block(p, cx, cy, cw, ch, t.Surface, true);
            Block(p, cx + 16, cy, 36, 4, tone, false);
            double ty = cy + 20;
            string icon = Field(raw, "icon"), value = Field(raw, "value");
            ty += rows == 1 ? 8 : 0;
            if (Icon(p, cx + 20, ty, rows == 1 ? 34 : 26, icon, tone) != null) ty += rows == 1 ? 52 : 38;
            if (value != null) { Label(p, cx + 20, ty, cw - 40, 48, value, rows == 1 ? 36 : 28, tone, true, t.TitleFont, 1, 1, null); ty += rows == 1 ? 56 : 42; }
            dynamic head = Words(p, cx + 20, ty, cw - 40, 30, Field(raw, "head") ?? "", rows == 1 ? 17.5 : 15.5, t.Text, true);
            Words(p, cx + 20, ty + (rows == 1 ? 38 : 30), cw - 40, cy + ch - ty - 46, Field(raw, "text") ?? "", rows == 1 ? 13.5 : 12.5, t.Muted, false);
            p.Motion.Add(back);
        }
        Callout(p, op, x, y + h - 42, w);
    }

    static void Stats(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        IList stats = Items(op, "stats", "items");
        int n = Math.Max(1, Math.Min(stats.Count, 6));
        int perRow = n <= 4 ? n : 3, rows = (n + perRow - 1) / perRow;
        IList points = Items(op, "points");
        double usable = h - (points.Count > 0 ? 34 * Math.Min(points.Count, 3) + 12 : 0) - (string.IsNullOrEmpty(op.Str("callout", null)) ? 0 : 54);
        double gap = 28, cw = (w - gap * (perRow - 1)) / perRow, ch = Math.Min(rows == 1 ? 190 : 150, usable / rows);
        if (rows == 1 && points.Count == 0)
        {
            // Nothing else on the slide: each figure on a card of its own, tall, with the number as the main thing.
            gap = 20; cw = (w - gap * (n - 1)) / n;
            double tall = Math.Min(usable - 20, 270), top = y + (usable - tall) / 2 - 6;
            for (int i = 0; i < n; i++)
            {
                Unit(p);
                object one = stats[i];
                double cx = x + i * (cw + gap);
                string tone = Pick(p, i), value = Field(one, "value") ?? "", unit = Field(one, "unit"), note = Field(one, "delta") ?? Field(one, "text");
                dynamic back = Block(p, cx, top, cw, tall, t.Surface, true);
                Block(p, cx + 22, top, 40, 4, tone, false);
                double ty = top + 26;
                if (Icon(p, cx + 22, ty, 28, Field(one, "icon"), tone) != null) ty += 44;
                double size = value.Length > 7 ? 38 : value.Length > 4 ? 50 : n >= 4 ? 60 : 68;
                dynamic figure = Label(p, cx + 22, ty, cw - 34, 84, value + (unit == null ? "" : " " + unit), size, tone, true, t.TitleFont, 1, 3, null);
                if (unit != null)
                {
                    try { dynamic tail = figure.TextFrame.TextRange.Characters(value.Length + 1, unit.Length + 1); tail.Font.Size = (float)(18 * p.S); tail.Font.Bold = 0; tail.Font.Color.RGB = Bgr(t.Text); }
                    catch (Exception) { }
                }
                ty += 92;
                Block(p, cx + 22, ty, 24, 2, t.Line, false);
                Label(p, cx + 22, ty + 12, cw - 44, note != null ? 44 : top + tall - ty - 26, Field(one, "label") ?? "", 13, t.Text, false, t.BodyFont, 1, 1, t.Accent);
                if (note != null) Label(p, cx + 22, ty + 60, cw - 44, top + tall - ty - 72, note, 11, t.Muted, false, t.BodyFont, 1, 1, null);
                p.Motion.Add(back);
            }
            Callout(p, op, x, y + h - 42, w);
            return;
        }
        y = Settle(y, usable, rows * ch);
        for (int i = 0; i < n; i++)
        {
            Unit(p);

            object raw = stats[i];
            double cx = x + (i % perRow) * (cw + gap), cy = y + (i / perRow) * ch;
            string tone = Pick(p, i);
            Rule(p, cx, cy + 4, cx + cw, cy + 4, tone, 1.5);
            double ty = cy + 14;
            if (Icon(p, cx, ty, 15, Field(raw, "icon"), tone) != null) Label(p, cx + 24, ty + 1, cw - 24, 18, Field(raw, "label") ?? "", 11.5, t.Muted, false, t.BodyFont, 1, 1, null);
            else Label(p, cx, ty + 1, cw, 18, Field(raw, "label") ?? "", 11.5, t.Muted, false, t.BodyFont, 1, 1, null);
            string value = Field(raw, "value") ?? "", unit = Field(raw, "unit");
            double size = value.Length > 8 ? 32 : perRow >= 4 ? 42 : 50;
            dynamic figure = Label(p, cx, ty + 26, cw, 66, value + (unit == null ? "" : " " + unit), size, tone, true, t.TitleFont, 1, 3, null);
            if (unit != null)
            {
                try { dynamic tail = figure.TextFrame.TextRange.Characters(value.Length + 1, unit.Length + 1); tail.Font.Size = (float)(17 * p.S); tail.Font.Bold = 0; tail.Font.Color.RGB = Bgr(t.Text); }
                catch (Exception) { }
            }
            string delta = Field(raw, "delta") ?? Field(raw, "text");
            if (delta != null) Label(p, cx, ty + 98, cw, 40, delta, 11.5, delta.TrimStart().StartsWith("-", StringComparison.Ordinal) || delta.Contains("下降") || delta.Contains("减少") ? t.Accent : t.Muted, false, t.BodyFont, 1, 1, null);
            p.Motion.Add(figure);
        }
        double py = y + rows * ch + 14;
        for (int i = 0; i < Math.Min(points.Count, 3); i++)
        {
            Unit(p);
            Block(p, x, py + 9, 6, 6, t.Accent, false);
            p.Motion.Add(Words(p, x + 18, py, w - 18, 30, Field(points[i], "text") ?? "", 14.5, t.Text, false));
            py += 38;
        }
        Callout(p, op, x, y + h - 42, w);
    }

    /// How high a formula stands, in lines of its own type size.
    static double Tall(string tex)
    {
        System.Text.RegularExpressions.RegexOptions none = System.Text.RegularExpressions.RegexOptions.None;
        double lines = 1.5;
        if (System.Text.RegularExpressions.Regex.IsMatch(tex, @"\\(lim|max|min|sup|inf|arg\s*max|arg\s*min|argmax|argmin|operatorname\*?\{[^}]*\})\s*(\\limits)?\s*_", none)) lines = Math.Max(lines, 2.2);
        if (System.Text.RegularExpressions.Regex.IsMatch(tex, @"\\(d|t|c)?frac|\\binom|\\over\b", none)) lines = Math.Max(lines, 2.5);
        if (System.Text.RegularExpressions.Regex.IsMatch(tex, @"\\(sum|prod|coprod|bigcup|bigcap|bigoplus|bigotimes|int|iint|iiint|oint)\s*(\\limits)?\s*[_^]", none)) lines = Math.Max(lines, 3.1);
        if (System.Text.RegularExpressions.Regex.IsMatch(tex, @"\\(under|over)brace", none)) lines += 1.2;
        if (tex.Contains("\\begin{"))
        {
            int rows = System.Text.RegularExpressions.Regex.Matches(tex, @"\\\\").Count + 1;
            lines = Math.Max(lines, 1.5 * rows + 0.6);
        }
        return lines;
    }

    static void Formulas(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        IList formulas = Items(op, "formulas", "formula");
        IList points = Items(op, "points");
        int n = Math.Max(1, Math.Min(formulas.Count, 3));
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        double textHeight = Math.Min(points.Count, 4) * 36 + (callout ? 54 : 0) + 8;
        // Each card is as high as its formula stands: a sum with limits or a fraction needs more than a plain line.
        double size = n == 1 ? 26 : 22, room = h - textHeight - 8 * (n - 1) - 6, sum;
        double[] need = new double[n];
        while (true)
        {
            sum = 0;
            for (int i = 0; i < n; i++)
            {
                string tex = Field(formulas[i], "latex") ?? Field(formulas[i], "text") ?? "";
                need[i] = (Field(formulas[i], "label") != null ? 20 : 0) + Tall(tex) * size + 12;
                sum += need[i];
            }
            if (sum <= room || size <= 14) break;
            size -= 1;
        }
        double spare = Math.Max(0, Math.Min((room - sum) / n, n == 1 ? 40 : 14));
        double cy = y;
        for (int i = 0; i < n; i++)
        {
            Unit(p);

            object raw = formulas[i];
            string latex = Field(raw, "latex") ?? Field(raw, "text") ?? "", label = Field(raw, "label");
            string tone = Pick(p, i);
            double each = need[i] + spare;
            dynamic back = Block(p, x, cy, w, each, t.Surface, false);
            Block(p, x, cy, 4, each, tone, false);
            if (label != null) Label(p, x + 18, cy + 8, w - 36, 16, label, 11, t.Muted, false, t.BodyFont, 1, 1, null);
            string source = latex.Trim().Trim('$');
            dynamic formula = Label(p, x + 18, cy + (label != null ? 20 : 0), w - 36, each - (label != null ? 20 : 0), '$' + source + '$', size, t.Text, false, "Cambria Math", 2, 3, null);
            formula.Name = "Formula " + (i + 1);
            p.Motion.Add(formula);
            cy += each + 8;
        }
        cy += 6;
        for (int i = 0; i < Math.Min(points.Count, 4); i++)
        {
            Unit(p);
            Block(p, x, cy + 8, 6, 6, Pick(p, i), false);
            p.Motion.Add(Words(p, x + 18, cy, w - 18, 30, Field(points[i], "text") ?? "", 14.5, t.Text, false));
            cy += 36;
        }
        Callout(p, op, x, y + h - 42, w);
    }

    static void Timeline(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        IList steps = Items(op, "steps", "items", "points");
        int n = Math.Max(1, Math.Min(steps.Count, 6));
        y = Settle(y, h - (string.IsNullOrEmpty(op.Str("callout", null)) ? 0 : 54), 250);
        double cw = w / n, line = y + 96;
        Rule(p, x, line, x + w, line, t.Line, 1.25);
        for (int i = 0; i < n; i++)
        {
            Unit(p);

            object raw = steps[i];
            double cx = x + i * cw;
            string tone = i == n - 1 ? t.Accent : t.Primary;
            string date = Field(raw, "when") ?? (i + 1).ToString("00");
            double units = 0;
            foreach (char ch in date) units += ch > 255 ? 1.0 : ch == '.' || ch == ' ' || ch == ',' ? 0.32 : 0.6;
            double fits = Math.Max(13, Math.Min(30, (cw - 16) / Math.Max(1, units * 1.16)));
            dynamic when = Label(p, cx, y + 24, cw - 12, 52, date, fits < 20 ? fits / 1.2 : fits, tone, true, t.TitleFont, 1, 4, null);
            Dot(p, cx + 8, line, 8, tone);
            Words(p, cx, line + 28, cw - 16, 30, Field(raw, "head") ?? "", 16.5, t.Text, true);
            Words(p, cx, line + 64, cw - 16, h - 170, Field(raw, "text") ?? "", 13.5, t.Muted, false);
            p.Motion.Add(when);
        }
        Callout(p, op, x, y + h - 42, w);
    }

    static void Process(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        IList steps = Items(op, "steps", "items", "points");
        int n = Math.Max(1, Math.Min(steps.Count, 5));
        double gap = 26, cw = (w - gap * (n - 1)) / n, ch = Math.Min(300, h - (string.IsNullOrEmpty(op.Str("callout", null)) ? 10 : 60));
        for (int i = 0; i < n; i++)
        {
            Unit(p);

            object raw = steps[i];
            double cx = x + i * (cw + gap);
            string tone = Pick(p, i);
            dynamic back = Outline(p, cx, y + 20, cw, ch - 20, t.Line, 1, true);
            dynamic disc = Dot(p, cx + 30, y + 20, 18, tone);
            Label(p, cx + 12, y + 2, 36, 36, (i + 1).ToString(), 16, t.Dark ? t.Bg : "#FFFFFF", true, t.TitleFont, 2, 3, null);
            Words(p, cx + 14, y + 56, cw - 28, 52, Field(raw, "head") ?? "", 16, t.Text, true);
            Words(p, cx + 14, y + 112, cw - 28, ch - 126, Field(raw, "text") ?? "", 13.5, t.Muted, false);
            if (i < n - 1)
            {
                dynamic arrow = p.Slide.Shapes.AddShape(33, PX(p, cx + cw + 5), PY(p, y + ch / 2 + 2), SX(p, gap - 10), SY(p, 14));
                arrow.Line.Visible = 0; arrow.Fill.Solid(); arrow.Fill.ForeColor.RGB = Bgr(t.Line);
                Track(p, arrow);
            }
            p.Motion.Add(back);
        }
        Callout(p, op, x, y + h - 42, w);
    }

    static void Compare(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        double gap = 24, cw = (w - gap) / 2, ch = h - (callout ? 54 : 4);
        string[] sides = { "left", "right" };
        for (int s = 0; s < 2; s++)
        {
            Dictionary<string, object> side = op.Raw(sides[s]) as Dictionary<string, object>;
            if (side == null) continue;
            double cx = x + s * (cw + gap);
            string tone = s == 0 ? t.Primary : t.Accent;
            Unit(p);
            dynamic back = Block(p, cx, y, cw, ch, t.Surface, true);
            Block(p, cx, y, cw, 40, tone, false);
            Label(p, cx + 16, y, cw - 32, 40, Field(side, "head") ?? "", 16, t.Dark ? t.Bg : "#FFFFFF", true, t.BodyFont, 1, 3, null);
            object raw;
            IList points = side.TryGetValue("points", out raw) ? raw as IList : null;
            double py = y + 56;
            if (points != null)
            {
                double row = Math.Min(64, (ch - 66) / Math.Max(1, points.Count));
                foreach (object point in points)
                {
                    if (py + row > y + ch + 2) break;
                    Block(p, cx + 18, py + 9, 6, 6, tone, false);
                    Words(p, cx + 34, py, cw - 52, row - 4, Field(point, "text") ?? "", 14.5, t.Text, false);
                    py += row;
                }
            }
            p.Motion.Add(back);
        }
        Callout(p, op, x, y + h - 42, w);
    }

    // ───────────────────────── more layouts ─────────────────────────

    /// A colour by its role in the theme (primary, accent, text, muted, surface, line, bg, white) or as #RRGGBB.
    static string Tone(Page p, string name, string fallback)
    {
        Theme t = p.T;
        switch ((name ?? "").Trim().ToLowerInvariant())
        {
            case "": return fallback;
            case "primary": return t.Primary;
            case "accent": return t.Accent;
            case "text": return t.Text;
            case "muted": return t.Muted;
            case "surface": return t.Surface;
            case "line": return t.Line;
            case "bg": case "background": return t.Bg;
            case "white": return "#FFFFFF";
            case "black": return "#000000";
        }
        return name.StartsWith("#", StringComparison.Ordinal) ? name : fallback;
    }

    /// Two colours mixed: share 0 gives the first, 1 the second.
    static string Mix(string a, string b, double share)
    {
        Color ca = ColorTranslator.FromHtml(a), cb = ColorTranslator.FromHtml(b);
        return "#" + ((int)Math.Round(ca.R + (cb.R - ca.R) * share)).ToString("X2") + ((int)Math.Round(ca.G + (cb.G - ca.G) * share)).ToString("X2") + ((int)Math.Round(ca.B + (cb.B - ca.B) * share)).ToString("X2");
    }

    /// The colours a chart of many parts goes through: the theme's two, then lighter steps of them.
    static string[] Palette(Page p)
    {
        Theme t = p.T;
        return new string[] { t.Primary, t.Accent, Mix(t.Primary, t.Bg, 0.45), Mix(t.Accent, t.Bg, 0.45), t.Muted, Mix(t.Primary, t.Bg, 0.7), Mix(t.Accent, t.Bg, 0.7), t.Line };
    }

    /// A definition, theorem or lemma: the statement set apart on a card, what leads to it listed underneath.
    static void Theorem(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        string label = op.Str("label", null), name = op.Str("name", null), statement = op.Str("statement", op.Str("text", ""));
        IList proof = Items(op, "proof", "steps", "points");
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        double usable = h - (callout ? 54 : 0);
        bool tall = System.Text.RegularExpressions.Regex.IsMatch(statement, @"\\(d|t)?frac|\\sum|\\int|\\prod|\\begin");
        double head = label != null || name != null ? 36 : 0;
        double sh = head + 34 + Math.Ceiling(statement.Length / 40.0) * 30 + (tall ? 26 : 0);
        sh = Math.Max(96, Math.Min(sh, usable * (proof.Count > 0 ? 0.52 : 0.9)));
        Unit(p);
        Block(p, x, y, w, sh, t.Surface, false);
        Block(p, x, y, 6, sh, t.Accent, false);
        double ty = y + 16;
        if (head > 0)
        {
            double cw = string.IsNullOrEmpty(label) ? 0 : label.Length * 15 + 26;
            if (cw > 0)
            {
                Block(p, x + 26, ty, cw, 24, t.Accent, true);
                Label(p, x + 26, ty, cw, 24, label, 10.5, "#FFFFFF", true, t.BodyFont, 2, 3, null).Name = "Label";
            }
            if (!string.IsNullOrEmpty(name)) Label(p, x + 26 + (cw > 0 ? cw + 12 : 0), ty, w - 64 - cw, 24, name, 12.5, t.Primary, true, t.BodyFont, 1, 3, null).Name = "Name";
            ty += 34;
        }
        dynamic said = Label(p, x + 26, ty, w - 52, y + sh - ty - 12, statement, 15.5, t.Text, false, t.TitleFont, 1, 3, t.Accent);
        said.Name = "Statement";
        p.Motion.Add(said);
        double py = y + sh + 16;
        int n = Math.Min(proof.Count, 5);
        if (n > 0)
        {
            Label(p, x, py, 300, 18, op.Str("proofTitle", "证明思路"), 10.5, t.Muted, true, t.BodyFont, 1, 1, null);
            py += 26;
            double row = Math.Min(60, (y + usable - py) / n);
            for (int i = 0; i < n; i++)
            {
                Unit(p);
                object raw = proof[i];
                string step = Field(raw, "head"), text = Field(raw, "text") ?? "";
                Dot(p, x + 13, py + 13, 13, Pick(p, i));
                Label(p, x, py, 26, 26, (i + 1).ToString(), 10.5, "#FFFFFF", true, t.BodyFont, 2, 3, null);
                p.Motion.Add(Words(p, x + 40, py + 1, w - 40, row - 6, (step != null ? "**" + step + "**　" : "") + text, 13.5, t.Text, false));
                py += row;
            }
        }
        Callout(p, op, x, y + h - 42, w);
    }

    /// Two to six pictures side by side, each with its caption: results to compare, the panels of an experiment.
    static void Gallery(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        IList images = Items(op, "images", "figures", "items");
        if (images.Count == 0) throw new Fail("BAD_ARGS", "A gallery slide needs \"images\": [{image, caption?}, ..].");
        int n = Math.Min(images.Count, 6);
        string text = op.Str("text", null);
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        double foot = (string.IsNullOrEmpty(text) ? 0 : 46) + (callout ? 54 : 0);
        // Upright pictures (portraits of people) stand in one row of upright places, so that little of them is cut;
        // others go in rows of up to three.
        double shape = 0;
        int counted = 0;
        for (int i = 0; i < n; i++)
        {
            string file = Field(images[i], "image") ?? Field(images[i], "path") ?? Field(images[i], "text");
            double one = string.IsNullOrEmpty(file) || !File.Exists(file) ? 0 : Aspect(file);
            if (one > 0) { shape += one; counted++; }
        }
        shape = counted > 0 ? shape / counted : 1.5;
        int perRow = shape < 0.95 ? n : n <= 3 ? n : n == 4 ? 2 : 3, rows = (n + perRow - 1) / perRow;
        double gap = 16, cw = (w - gap * (perRow - 1)) / perRow, ch = (h - foot - gap * (rows - 1)) / rows;
        // A single row of upright places need not be wider than its pictures ask for.
        if (rows == 1 && shape < 0.95) cw = Math.Min(cw, (ch - 22) * Math.Max(shape * 1.15, 0.62));
        string wanted = op.Str("focus", null);
        // Captions too long for one line of their column get two, under every picture alike.
        bool twoLines = false;
        for (int i = 0; i < n; i++) { string said = Field(images[i], "caption") ?? Field(images[i], "head"); if (said != null && said.Length * 11.5 > cw + 12) twoLines = true; }
        for (int i = 0; i < n; i++)
        {
            Unit(p);
            object raw = images[i];
            string path = Field(raw, "image") ?? Field(raw, "path") ?? Field(raw, "text"), caption = Field(raw, "caption") ?? Field(raw, "head");
            if (string.IsNullOrEmpty(path)) throw new Fail("BAD_ARGS", "Picture " + (i + 1) + " of the gallery has no \"image\".");
            // The last row may hold fewer: it stands in the middle, not at the left with a hole beside it.
            int row = i / perRow, inRow = Math.Min(perRow, n - row * perRow);
            double rowWidth = inRow * cw + (inRow - 1) * gap;
            double cx = x + (w - rowWidth) / 2 + (i % perRow) * (cw + gap), cy = y + row * (ch + gap), under = caption != null ? (twoLines ? 36 : 22) : 0;
            Focus = Field(raw, "focus") ?? wanted;
            bool whole = Field(raw, "fit") == "contain" || (Field(raw, "fit") != "cover" && IsFigure(path));
            p.Motion.Add(whole ? Figure(p, cx, cy, cw, ch - under, path) : Photo(p, cx, cy, cw, ch - under, path));
            if (caption != null) Label(p, cx - 6, cy + ch - under + 4, cw + 12, twoLines ? 30 : 16, caption, 9.5, t.Muted, false, t.BodyFont, 2, 1, t.Accent).Name = "Caption " + (i + 1);
        }
        if (!string.IsNullOrEmpty(text)) { Unit(p); Label(p, x, y + h - foot + 6, w, 38, text, 13.5, t.Text, false, t.BodyFont, 2, 3, t.Accent).Name = "Text"; }
        Callout(p, op, x, y + h - 42, w);
    }

    /// The two to four things to take away, numbered large; the first stands out.
    static void Summary(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        IList points = Items(op, "points", "items");
        int n = Math.Max(1, Math.Min(points.Count, 4));
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        double usable = h - (callout ? 54 : 0), gap = 18, cw = (w - gap * (n - 1)) / n, ch = Math.Min(usable, 290);
        double top = Settle(y, usable, ch);
        for (int i = 0; i < n; i++)
        {
            Unit(p);
            object raw = points[i];
            double cx = x + i * (cw + gap);
            bool lead = i == 0;
            string ground = lead ? (t.Dark ? t.Accent : t.Primary) : t.Surface, ink = lead ? "#FFFFFF" : t.Text, soft = lead ? "#FFFFFF" : t.Muted, tone = lead ? "#FFFFFF" : Pick(p, i);
            dynamic back = Block(p, cx, top, cw, ch, ground, false);
            dynamic number = Label(p, cx + 22, top + 16, cw - 44, 62, (i + 1).ToString("00"), 40, tone, true, t.TitleFont, 1, 1, null);
            if (lead) { try { number.TextFrame2.TextRange.Font.Fill.Transparency = 0.45f; } catch (Exception) { } }
            Block(p, cx + 22, top + 88, 30, 3, lead ? "#FFFFFF" : tone, false);
            string head = Field(raw, "head"), text = Field(raw, "text") ?? "";
            double ty = top + 104;
            if (head != null) { Label(p, cx + 22, ty, cw - 44, 52, head, 15, ink, true, t.BodyFont, 1, 1, null); ty += 58; }
            Label(p, cx + 22, ty, cw - 44, top + ch - ty - 16, text, head != null ? 11.5 : 13.5, head != null ? soft : ink, false, t.BodyFont, 1, 1, lead ? null : t.Accent);
            p.Motion.Add(back);
        }
        Callout(p, op, x, y + h - 42, w);
    }

    /// Shapes placed by the author on the 960 x 540 canvas, in the colours and type of the theme: for a layout that
    /// none of the ready kinds gives.
    static void Canvas(Page p, Bag op)
    {
        Theme t = p.T;
        IList items = Items(op, "items", "elements");
        if (items.Count == 0) throw new Fail("BAD_ARGS", "A canvas slide needs \"items\": [{type, box: [x, y, w, h], ..}, ..] on the 960 x 540 canvas.");
        int index = 0;
        Dictionary<string, object> known = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
        foreach (object raw in items)
        {
            index++;
            Bag it = new Bag(raw);
            string type = it.Str("type", "text").ToLowerInvariant();
            if (type == "edge" || type == "connect" || type == "link")
            {
                // A connector from one named item to another: it finds the nearest sides by itself and stays attached.
                object a, b2;
                string from = it.Need("from"), to = it.Need("to");
                if (!known.TryGetValue(from, out a) || !known.TryGetValue(to, out b2)) throw new Fail("BAD_ARGS", "Item " + index + " connects \"" + from + "\" to \"" + to + "\", but " + (known.ContainsKey(from) ? "\"" + to + "\"" : "\"" + from + "\"") + " is not the id of an item before it. Ids so far: " + string.Join(", ", new List<string>(known.Keys).ToArray()) + ".");
                string way = it.Str("kind", it.Str("shape", "elbow")).ToLowerInvariant();
                dynamic link = p.Slide.Shapes.AddConnector(way == "straight" ? 1 : way == "curve" || way == "curved" ? 3 : 2, 0, 0, 10, 10);
                Dictionary<string, int> sides = new Dictionary<string, int> { { "top", 1 }, { "left", 2 }, { "bottom", 3 }, { "right", 4 } };
                link.ConnectorFormat.BeginConnect((dynamic)a, 1);
                link.ConnectorFormat.EndConnect((dynamic)b2, 1);
                try { link.RerouteConnections(); } catch (Exception) { }
                int side;
                if (it.Has("fromSide") && sides.TryGetValue(it.Need("fromSide").ToLowerInvariant(), out side)) { try { link.ConnectorFormat.BeginConnect((dynamic)a, side); } catch (Exception) { } }
                if (it.Has("toSide") && sides.TryGetValue(it.Need("toSide").ToLowerInvariant(), out side)) { try { link.ConnectorFormat.EndConnect((dynamic)b2, side); } catch (Exception) { } }
                link.Line.ForeColor.RGB = Bgr(Tone(p, it.Str("color", null), t.Muted));
                link.Line.Weight = (float)(it.Num("weight", 1.5) * p.S);
                if (!it.Has("arrow") || it.On("arrow")) { try { link.Line.EndArrowheadStyle = 2; } catch (Exception) { } }
                if (it.Flag("dashed", false)) { try { link.Line.DashStyle = 4; } catch (Exception) { } }
                if (p.Current == null) Unit(p);
                Track(p, link);
                if (it.Has("label"))
                {
                    double mx = ((double)link.Left + (double)link.Width / 2 - p.Ox) / p.Sx, my = ((double)link.Top + (double)link.Height / 2 - p.Oy) / p.Sy;
                    Label(p, mx - 70, my - 20, 140, 16, it.Need("label"), 10, Tone(p, it.Str("color", null), t.Muted), false, t.BodyFont, 2, 4, null);
                }
                continue;
            }
            if (type == "node") type = "panel";
            double[] b = null;
            IList given = it.List("box");
            if (given != null && given.Count == 4) { b = new double[4]; for (int k = 0; k < 4; k++) b[k] = Number(given[k]); }
            if (b == null || b.Length != 4) throw new Fail("BAD_ARGS", "Item " + index + " of the canvas (" + type + ") needs \"box\": [x, y, w, h] on the 960 x 540 canvas.");
            if (!it.Flag("with", false) || p.Current == null) Unit(p);
            string color = it.Str("color", null), fill = it.Str("fill", null);
            int align = it.Str("align", "left") == "center" ? 2 : it.Str("align", "left") == "right" ? 3 : 1;
            int anchor = it.Str("valign", "top") == "middle" ? 3 : it.Str("valign", "top") == "bottom" ? 4 : 1;
            dynamic made = null;
            if (type == "text" || type == "title")
            {
                made = Label(p, b[0], b[1], b[2], b[3], Lines(it.Raw("text") ?? ""), it.Num("size", type == "title" ? 24 : 14), Tone(p, color, type == "title" ? t.Primary : t.Text), it.Flag("bold", type == "title"), type == "title" || it.Str("font", "") == "title" ? t.TitleFont : t.BodyFont, align, anchor, t.Accent);
            }
            else if (type == "card")
            {
                string tone = Tone(p, color, Pick(p, index - 1));
                made = Block(p, b[0], b[1], b[2], b[3], Tone(p, fill, t.Surface), true);
                Block(p, b[0] + 16, b[1], 36, 4, tone, false);
                double ty = b[1] + 18;
                if (Icon(p, b[0] + 18, ty, 26, it.Str("icon", null), tone) != null) ty += 36;
                string value = it.Str("value", null), head = it.Str("head", null), text = it.Str("text", "");
                if (value != null) { Label(p, b[0] + 18, ty, b[2] - 36, 40, value, 26, tone, true, t.TitleFont, 1, 1, null); ty += 42; }
                if (head != null) { Label(p, b[0] + 18, ty, b[2] - 36, 26, head, it.Num("size", 14) + 1, t.Text, true, t.BodyFont, 1, 1, null); ty += 30; }
                Label(p, b[0] + 18, ty, b[2] - 36, Math.Max(16, b[1] + b[3] - ty - 12), text, it.Num("size", 14) - 2, t.Muted, false, t.BodyFont, 1, 1, t.Accent);
            }
            else if (type == "panel" || type == "shape" || type == "box" || type == "circle")
            {
                string shape = type == "circle" ? "circle" : it.Str("shape", it.Str("type", "") == "node" ? "round" : "rect");
                string ground = Tone(p, fill ?? color, t.Surface);
                if (shape == "circle") made = Dot(p, b[0] + b[2] / 2, b[1] + b[3] / 2, Math.Min(b[2], b[3]) / 2, ground);
                else made = Block(p, b[0], b[1], b[2], b[3], ground, shape == "round" || shape == "rounded");
                if (it.Has("transparency")) { try { made.Fill.Transparency = (float)it.Num("transparency", 0); } catch (Exception) { } }
                if (it.Has("text"))
                {
                    // Words on a filled shape read in white when the ground is dark.
                    Color g = ColorTranslator.FromHtml(ground);
                    bool dark = g.R * 0.299 + g.G * 0.587 + g.B * 0.114 < 150;
                    Label(p, b[0] + 8, b[1], b[2] - 16, b[3], Lines(it.Raw("text")), it.Num("size", 13), Tone(p, it.Str("ink", null), dark ? "#FFFFFF" : t.Text), it.Flag("bold", false), t.BodyFont, it.Has("align") ? align : 2, it.Has("valign") ? anchor : 3, dark ? null : t.Accent);
                }
            }
            else if (type == "line" || type == "arrow")
            {
                made = Rule(p, b[0], b[1], b[0] + b[2], b[1] + b[3], Tone(p, color, type == "arrow" ? t.Muted : t.Line), it.Num("weight", type == "arrow" ? 1.5 : 0.75));
                if (type == "arrow") { try { made.Line.EndArrowheadStyle = 2; } catch (Exception) { } }
                if (it.Flag("dashed", false)) { try { made.Line.DashStyle = 4; } catch (Exception) { } }
            }
            else if (type == "image" || type == "figure" || type == "photo")
            {
                string path = it.Str("image", it.Str("path", null));
                if (path == null) throw new Fail("BAD_ARGS", "Item " + index + " of the canvas (image) needs \"image\": a path.");
                Focus = it.Str("focus", null);
                bool whole = type == "figure" || it.Str("fit", "") == "contain" || (type == "image" && it.Str("fit", "") != "cover" && IsFigure(path));
                made = whole ? Figure(p, b[0], b[1], b[2], b[3], path) : Photo(p, b[0], b[1], b[2], b[3], path);
            }
            else if (type == "formula")
            {
                string tone = Tone(p, color, Pick(p, index - 1)), label = it.Str("label", null);
                Block(p, b[0], b[1], b[2], b[3], Tone(p, fill, t.Surface), false);
                Block(p, b[0], b[1], 4, b[3], tone, false);
                if (label != null) Label(p, b[0] + 18, b[1] + 8, b[2] - 36, 16, label, 11, t.Muted, false, t.BodyFont, 1, 1, null);
                string source = it.Str("latex", it.Str("text", "")).Trim().Trim('$');
                made = Label(p, b[0] + 18, b[1] + (label != null ? 20 : 0), b[2] - 36, b[3] - (label != null ? 20 : 0), '$' + source + '$', it.Num("size", 22), t.Text, false, "Cambria Math", 2, 3, null);
            }
            else if (type == "stat")
            {
                string tone = Tone(p, color, Pick(p, index - 1)), value = it.Str("value", ""), unit = it.Str("unit", null);
                Rule(p, b[0], b[1] + 4, b[0] + b[2], b[1] + 4, tone, 1.5);
                Label(p, b[0], b[1] + 14, b[2], 18, it.Str("label", ""), 11.5, t.Muted, false, t.BodyFont, 1, 1, null);
                made = Label(p, b[0], b[1] + 38, b[2], Math.Max(40, b[3] - 70), value + (unit == null ? "" : " " + unit), it.Num("size", 44), tone, true, t.TitleFont, 1, 3, null);
                if (unit != null) { try { dynamic tail = made.TextFrame.TextRange.Characters(value.Length + 1, unit.Length + 1); tail.Font.Size = (float)(17 * p.S); tail.Font.Bold = 0; tail.Font.Color.RGB = Bgr(t.Text); } catch (Exception) { } }
                if (it.Has("text")) Label(p, b[0], b[1] + b[3] - 30, b[2], 30, it.Need("text"), 11.5, t.Muted, false, t.BodyFont, 1, 1, null);
            }
            else if (type == "chart")
            {
                Dictionary<string, object> chart = (it.Raw("chart") as Dictionary<string, object>) ?? (raw as Dictionary<string, object>);
                Draw(p, chart, b[0], b[1], b[2], b[3]);
            }
            else if (type == "icon") made = Icon(p, b[0], b[1], Math.Min(b[2], b[3]), it.Str("icon", it.Str("name", null)), Tone(p, color, t.Primary));
            else if (type == "bullets" || type == "points") Bullets(p, it, b[0], b[1], b[2], b[3], false);
            else throw new Fail("BAD_ARGS", "Item " + index + " of the canvas has the unknown type \"" + type + "\": use text, title, card, panel, node, circle, line, arrow, edge, image, figure, formula, stat, chart, icon or bullets.");
            if (made != null && it.Has("name")) { try { made.Name = it.Need("name"); } catch (Exception) { } }
            if (made != null && it.Has("id")) known[it.Need("id")] = made;
        }
    }

    static readonly System.Text.RegularExpressions.Regex CodeParts = new System.Text.RegularExpressions.Regex(
        "(?<comment>#.*|//.*)|(?<text>\"(?:\\\\.|[^\"\\\\])*\"|'(?:\\\\.|[^'\\\\])*')|(?<number>\\b\\d+(?:\\.\\d+)?\\b)|(?<word>\\b(?:def|class|return|if|elif|else|for|while|in|import|from|as|with|try|except|finally|raise|lambda|yield|pass|break|continue|and|or|not|is|None|True|False|self|function|const|let|var|new|async|await|public|private|static|void|int|float|double|bool|string|struct|fn|func|package|namespace|using|this|null|true|false|switch|case|do|typedef|template|typename|auto|export|default|extends|implements|interface|type|enum|match|where|select|end|then|begin)\\b)",
        System.Text.RegularExpressions.RegexOptions.Compiled);

    /// A listing of code: a dark panel, a fixed-width face, line numbers, comments and keywords in their own colours.
    static void Code(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        string code = (op.Str("code", op.Str("text", "")) ?? "").Replace("\r\n", "\n").Replace("\r", "\n").Replace("\t", "    ").TrimEnd();
        if (code.Length == 0) throw new Fail("BAD_ARGS", "A code slide needs \"code\": the listing.");
        string[] lines = code.Split('\n');
        if (lines.Length > 26) throw new Fail("BAD_ARGS", "The listing has " + lines.Length + " lines; a slide holds 26 at most (and reads best with 8 to 16). Show the part that matters, or split it over two slides.");
        IList points = Items(op, "points", "side");
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        double usable = h - (callout ? 54 : 0), cw = points.Count > 0 ? w * 0.62 : w;
        int longest = 1;
        foreach (string line in lines) { int length = 0; foreach (char ch in line) length += ch > 255 ? 2 : 1; longest = Math.Max(longest, length); }
        // As large as the listing allows, by its number of lines and its longest line.
        double size = Math.Min(15, Math.Min((usable - 52) / (lines.Length * 1.38), (cw - 78) / (longest * 0.61)));
        size = Math.Max(7.5, size);
        double pitch = size * 1.38, ph = Math.Min(usable, lines.Length * pitch + 52);
        string ground = t.Dark ? t.Surface : "#1E2430", ink = "#E6EAF2", dim = "#7C8798";
        Unit(p);
        dynamic back = Block(p, x, y, cw, ph, ground, true);
        back.Name = "Code panel";
        Dot(p, x + 18, y + 16, 4.5, "#FF5F57"); Dot(p, x + 34, y + 16, 4.5, "#FEBC2E"); Dot(p, x + 50, y + 16, 4.5, "#28C840");
        string language = op.Str("language", op.Str("lang", null));
        if (!string.IsNullOrEmpty(language)) Label(p, x + cw - 170, y + 8, 154, 16, language, 9, dim, false, "Consolas", 3, 1, null);
        StringBuilder numbers = new StringBuilder();
        for (int i = 0; i < lines.Length; i++) numbers.Append(i > 0 ? "\r" : "").Append(i + 1);
        string[] faces = { "Consolas", "Consolas" };
        for (int part = 0; part < 2; part++)
        {
            dynamic box = p.Slide.Shapes.AddTextbox(1, PX(p, x + (part == 0 ? 10 : 50)), PY(p, y + 34), SX(p, part == 0 ? 30 : cw - 64), SY(p, ph - 44));
            dynamic frame = box.TextFrame;
            frame.WordWrap = 0;
            try { frame.AutoSize = 0; } catch (Exception) { }
            frame.MarginLeft = 0; frame.MarginRight = 0; frame.MarginTop = 0; frame.MarginBottom = 0;
            dynamic range = frame.TextRange;
            range.Text = part == 0 ? numbers.ToString() : string.Join("\r", lines);
            range.Font.Name = faces[part];
            try { range.Font.NameFarEast = "Microsoft YaHei"; } catch (Exception) { }
            range.Font.Size = (float)(size * p.S);
            range.Font.Bold = 0;
            range.Font.Color.RGB = Bgr(part == 0 ? dim : ink);
            range.ParagraphFormat.Alignment = part == 0 ? 3 : 1;
            try { range.ParagraphFormat.SpaceWithin = 1.15f; } catch (Exception) { }
            box.Name = part == 0 ? "Line numbers" : "Code";
            Track(p, box);
            if (part == 0) continue;
            int offset = 0;
            foreach (string line in lines)
            {
                foreach (System.Text.RegularExpressions.Match m in CodeParts.Matches(line))
                {
                    string tone = m.Groups["comment"].Success ? "#7F8C98" : m.Groups["text"].Success ? "#E6B673" : m.Groups["number"].Success ? "#D19A66" : "#7FB4FF";
                    try { dynamic piece = range.Characters(offset + m.Index + 1, m.Length); piece.Font.Color.RGB = Bgr(tone); if (m.Groups["comment"].Success) piece.Font.Italic = -1; }
                    catch (Exception) { }
                    if (m.Groups["comment"].Success) break;
                }
                offset += line.Length + 1;
            }
            p.Motion.Add(box);
        }
        double sx = x + cw + 28, sy = y + 4, sw = w - cw - 28;
        for (int i = 0; i < Math.Min(points.Count, 4); i++)
        {
            Unit(p);
            string head = Field(points[i], "head"), text = Field(points[i], "text") ?? "";
            Block(p, sx, sy, 34, 3, Pick(p, i), false);
            if (head != null) { Words(p, sx, sy + 10, sw, 22, head, 14, t.Text, true); sy += 26; }
            p.Motion.Add(Words(p, sx, sy + 10, sw, 62, text, 12.5, head != null ? t.Muted : t.Text, false));
            sy += head != null ? 62 : 76;
        }
        Callout(p, op, x, y + h - 42, w);
    }

    /// A table drawn from shapes, cell by cell: unlike a table of PowerPoint's own, its cells can hold equations.
    static void DrawnTable(Page p, IList data, int rows, int cols, double x, double y, double w, double rowHeight, double size)
    {
        Theme t = p.T;
        // The columns share the width by how much they hold, within bounds.
        double[] weight = new double[cols];
        for (int c = 0; c < cols; c++)
        {
            double longest = 2;
            for (int r = 0; r < rows; r++)
            {
                IList cells = data[r] as IList;
                string value = cells != null && c < cells.Count && cells[c] != null ? Convert.ToString(cells[c], System.Globalization.CultureInfo.InvariantCulture) : "";
                double length = 0;
                foreach (char ch in System.Text.RegularExpressions.Regex.Replace(value, @"\\[a-zA-Z]+|[{}$^_\\]", "")) length += ch > 255 ? 2 : 1;
                longest = Math.Max(longest, length);
            }
            weight[c] = Math.Max(8, Math.Min(longest, 40));
        }
        double total = 0;
        foreach (double one in weight) total += one;
        double cy = y;
        for (int r = 0; r < rows; r++)
        {
            Unit(p);
            IList cells = data[r] as IList;
            dynamic band = Block(p, x, cy, w, rowHeight, r == 0 ? t.Primary : r % 2 == 0 ? t.Surface : t.Bg, false);
            band.Name = "Row " + (r + 1);
            double cx = x;
            for (int c = 0; c < cols; c++)
            {
                double cw = w * weight[c] / total;
                string value = cells != null && c < cells.Count && cells[c] != null ? Convert.ToString(cells[c], System.Globalization.CultureInfo.InvariantCulture) : "";
                if (value.Length > 0) Label(p, cx + 12, cy, cw - 24, rowHeight, value, size / 1.2, r == 0 ? (t.Dark ? t.Bg : "#FFFFFF") : t.Text, r == 0, t.BodyFont, c == 0 ? 1 : 2, 3, r == 0 ? null : t.Accent);
                cx += cw;
            }
            if (r > 0) Rule(p, x, cy + rowHeight, x + w, cy + rowHeight, t.Line, 0.5);
            p.Motion.Add(band);
            cy += rowHeight;
        }
    }

    /// Round steps for an axis that must cover lo to hi (either may be negative).
    static void Axis(ref double lo, ref double hi, out double step)
    {
        if (hi <= lo) { hi = lo + 1; }
        double span = hi - lo, raw = span / 5, power = Math.Pow(10, Math.Floor(Math.Log10(raw)));
        step = 10 * power;
        foreach (double m in new double[] { 1, 2, 2.5, 5, 10 }) if (m * power >= raw) { step = m * power; break; }
        // An axis of positive values that starts near zero starts at zero.
        if (lo >= 0 && lo < span * 0.6) lo = 0;
        lo = Math.Floor(lo / step + 1e-9) * step;
        hi = Math.Ceiling(hi / step - 1e-9) * step;
    }

    /// An axis from a start the author fixed: finer steps, and no reaching back to zero.
    static void Pinned(double from, double to, out double lo, out double hi, out double step)
    {
        if (to <= from) to = from + 1;
        double raw = (to - from) / 7, power = Math.Pow(10, Math.Floor(Math.Log10(raw)));
        step = 10 * power;
        foreach (double m in new double[] { 1, 2, 2.5, 5, 10 }) if (m * power >= raw) { step = m * power; break; }
        lo = Math.Floor(from / step + 1e-9) * step;
        hi = Math.Ceiling(to / step - 1e-9) * step;
    }

    static string Tick(double value, double step)
    {
        int decimals = step >= 1 ? 0 : (int)Math.Min(4, Math.Ceiling(-Math.Log10(step) - 1e-9));
        if (Math.Abs(step * Math.Pow(10, decimals) - Math.Round(step * Math.Pow(10, decimals))) > 1e-6) decimals = Math.Min(4, decimals + 1);
        if (Math.Abs(value) < step * 1e-6) value = 0;
        return value.ToString("F" + decimals, System.Globalization.CultureInfo.InvariantCulture);
    }

    /// Points and curves over two number axes, drawn from shapes: a scatter of measurements, a fitted line, the graph of a function.
    static void DrawXY(Page p, Dictionary<string, object> chart, string type, double x, double y, double w, double h)
    {
        Theme t = p.T;
        object raw;
        IList series = chart.TryGetValue("series", out raw) ? raw as IList : null;
        if (series == null || series.Count == 0) throw new Fail("BAD_ARGS", "A " + type + " chart needs \"series\": [{name?, points: [[x, y], ..]}] (or x: [..] and y: [..]).");
        List<List<double[]>> all = new List<List<double[]>>();
        List<string> names = new List<string>();
        List<bool> joined = new List<bool>();
        double x0 = double.MaxValue, x1 = double.MinValue, y0 = double.MaxValue, y1 = double.MinValue;
        int m = Math.Min(series.Count, 6);
        for (int s = 0; s < m; s++)
        {
            Dictionary<string, object> one = series[s] as Dictionary<string, object>;
            List<double[]> points = new List<double[]>();
            IList pairs = one != null && one.TryGetValue("points", out raw) ? raw as IList : one == null ? series[s] as IList : null;
            if (pairs != null)
            {
                foreach (object pair in pairs)
                {
                    IList xy = pair as IList;
                    if (xy != null && xy.Count >= 2) points.Add(new double[] { Number(xy[0]), Number(xy[1]) });
                }
            }
            else if (one != null)
            {
                IList xs = one.TryGetValue("x", out raw) ? raw as IList : null, ys = one.TryGetValue("y", out raw) ? raw as IList : one.TryGetValue("values", out raw) ? raw as IList : null;
                for (int i = 0; xs != null && ys != null && i < xs.Count && i < ys.Count; i++) points.Add(new double[] { Number(xs[i]), Number(ys[i]) });
            }
            if (points.Count == 0) throw new Fail("BAD_ARGS", "Series " + (s + 1) + " of the chart has no points: give points: [[x, y], ..] or x: [..] and y: [..].");
            if (points.Count > 600) points = points.GetRange(0, 600);
            foreach (double[] point in points) { x0 = Math.Min(x0, point[0]); x1 = Math.Max(x1, point[0]); y0 = Math.Min(y0, point[1]); y1 = Math.Max(y1, point[1]); }
            all.Add(points);
            names.Add(one != null && one.TryGetValue("name", out raw) && raw != null ? Convert.ToString(raw) : "");
            bool line = type != "scatter";
            if (one != null && one.TryGetValue("line", out raw) && raw != null) line = Truthy(raw) || Convert.ToString(raw) == "True" || Convert.ToString(raw) == "true";
            joined.Add(line);
        }
        if (chart.TryGetValue("xMin", out raw) && raw != null) x0 = Number(raw);
        if (chart.TryGetValue("xMax", out raw) && raw != null) x1 = Number(raw);
        if (chart.TryGetValue("yMin", out raw) && raw != null) y0 = Number(raw);
        if (chart.TryGetValue("yMax", out raw) && raw != null) y1 = Number(raw);
        double xs1, ys1;
        // An end of an axis the author fixed stays where it was put.
        double fx0 = x0, fx1 = x1, fy0 = y0, fy1 = y1;
        bool pinX = chart.TryGetValue("xMin", out raw) && raw != null, pinY = chart.TryGetValue("yMin", out raw) && raw != null;
        Axis(ref x0, ref x1, out xs1);
        Axis(ref y0, ref y1, out ys1);
        if (pinX && x0 < fx0) Pinned(fx0, fx1, out x0, out x1, out xs1);
        if (pinY && y0 < fy0) Pinned(fy0, fy1, out y0, out y1, out ys1);
        string xTitle = chart.TryGetValue("xTitle", out raw) && raw != null ? Convert.ToString(raw) : null;
        string yTitle = chart.TryGetValue("yTitle", out raw) && raw != null ? Convert.ToString(raw) : chart.TryGetValue("unit", out raw) && raw != null ? Convert.ToString(raw) : null;
        bool fit = chart.TryGetValue("fit", out raw) && raw != null && (Truthy(raw) || Convert.ToString(raw).ToLowerInvariant() == "true");
        string[] tones = Palette(p);
        for (int s = 0; s < m; s++)
        {
            Dictionary<string, object> one = series[s] as Dictionary<string, object>;
            if (one != null && one.TryGetValue("color", out raw) && raw != null) tones[s] = Tone(p, Convert.ToString(raw), tones[s]);
        }
        string where = chart.TryGetValue("legend", out raw) && raw != null ? Convert.ToString(raw).ToLowerInvariant() : "";
        bool named = false;
        foreach (string one in names) if (one.Length > 0) named = true;
        // Lines are named where they end, which reads faster than a legend; a scatter keeps the legend.
        bool ends = named && where != "top" && where != "none" && (where == "end" || type != "scatter");
        bool legend = m > 1 && !ends && where != "none";
        double margin = 14;
        if (ends) { foreach (string one in names) { double wide = 0; foreach (char ch in one) wide += ch > 255 ? 11.5 : 6.5; margin = Math.Max(margin, Math.Min(170, wide + 60)); } }
        if (yTitle != null) Label(p, x, y, 300, 14, yTitle, 9.5, t.Muted, false, t.BodyFont, 1, 1, null);
        if (legend)
        {
            double lx = x + w;
            for (int s = m - 1; s >= 0; s--)
            {
                double width = 16 + names[s].Length * 11 + 14;
                lx -= width;
                Block(p, lx, y + 3, 9, 9, tones[s], false);
                Label(p, lx + 14, y, width - 14, 14, names[s], 10, t.Muted, false, t.BodyFont, 1, 1, null);
            }
        }
        double head = yTitle != null || legend ? 22 : 6, axis = 46, foot = xTitle != null ? 42 : 24;
        double px = x + axis, pw = w - axis - margin, py = y + head + 10, ph = h - head - 10 - foot;
        List<double[]> tips = new List<double[]>();   // where each named line ends: series, height on the slide, last value
        for (double gy = y0; gy <= y1 + ys1 * 1e-6; gy += ys1)
        {
            double sy = py + ph - ph * (gy - y0) / (y1 - y0);
            bool zero = Math.Abs(gy) < ys1 * 1e-6;
            Rule(p, px, sy, px + pw, sy, zero ? t.Muted : t.Line, zero || gy <= y0 + ys1 * 1e-6 ? 1 : 0.5);
            Label(p, x, sy - 8, axis - 6, 16, Tick(gy, ys1), 9.5, t.Muted, false, t.BodyFont, 3, 3, null);
        }
        for (double gx = x0; gx <= x1 + xs1 * 1e-6; gx += xs1)
        {
            double sx = px + pw * (gx - x0) / (x1 - x0);
            bool zero = Math.Abs(gx) < xs1 * 1e-6;
            Rule(p, sx, py, sx, py + ph, zero ? t.Muted : t.Line, zero || gx <= x0 + xs1 * 1e-6 ? 1 : 0.5);
            Label(p, sx - 30, py + ph + 5, 60, 16, Tick(gx, xs1), 9.5, t.Muted, false, t.BodyFont, 2, 1, null);
        }
        if (xTitle != null) Label(p, px, py + ph + 24, pw, 16, xTitle, 10.5, t.Muted, false, t.BodyFont, 2, 1, null);
        for (int s = 0; s < m; s++)
        {
            List<double[]> points = all[s];
            if (joined[s])
            {
                dynamic builder = null;
                foreach (double[] point in points)
                {
                    // What leaves the plot is held at its edge.
                    double cy = Math.Max(y0, Math.Min(y1, point[1])), cx = Math.Max(x0, Math.Min(x1, point[0]));
                    float fx = PX(p, px + pw * (cx - x0) / (x1 - x0)), fy = PY(p, py + ph - ph * (cy - y0) / (y1 - y0));
                    if (builder == null) builder = p.Slide.Shapes.BuildFreeform(0, fx, fy);
                    else builder.AddNodes(0, 0, fx, fy);
                }
                if (builder != null && points.Count > 1)
                {
                    dynamic path = Track(p, builder.ConvertToShape());
                    path.Fill.Visible = 0;
                    path.Line.ForeColor.RGB = Bgr(tones[s]);
                    path.Line.Weight = 2.25f * p.S;
                    Dictionary<string, object> meant = series[s] as Dictionary<string, object>;
                    object dash;
                    if (meant != null && meant.TryGetValue("dashed", out dash) && dash != null && (Truthy(dash) || Convert.ToString(dash).ToLowerInvariant() == "true")) { try { path.Line.DashStyle = 4; } catch (Exception) { } }
                    p.Motion.Add(path);
                }
                if (ends && names[s].Length > 0)
                {
                    double[] last = points[points.Count - 1];
                    tips.Add(new double[] { s, py + ph - ph * (Math.Max(y0, Math.Min(y1, last[1])) - y0) / (y1 - y0), last[1] });
                }
            }
            if (!joined[s] || points.Count <= 14)
            {
                double radius = points.Count > 120 ? 2.2 : points.Count > 40 ? 3 : 4;
                foreach (double[] point in points)
                {
                    if (point[0] < x0 || point[0] > x1 || point[1] < y0 || point[1] > y1) continue;
                    dynamic dot = Dot(p, px + pw * (point[0] - x0) / (x1 - x0), py + ph - ph * (point[1] - y0) / (y1 - y0), radius, tones[s]);
                    if (!joined[s]) { try { dot.Fill.Transparency = points.Count > 60 ? 0.35f : 0.15f; } catch (Exception) { } }
                }
            }
            if (fit && !joined[s] && points.Count >= 3)
            {
                // The least-squares line through the points, and how well it holds.
                double n = points.Count, sx = 0, sy = 0, sxx = 0, sxy = 0, syy = 0;
                foreach (double[] point in points) { sx += point[0]; sy += point[1]; sxx += point[0] * point[0]; sxy += point[0] * point[1]; syy += point[1] * point[1]; }
                double den = n * sxx - sx * sx;
                if (Math.Abs(den) > 1e-12)
                {
                    double slope = (n * sxy - sx * sy) / den, cut = (sy - slope * sx) / n;
                    double varY = n * syy - sy * sy, r2 = varY > 1e-12 ? Math.Pow(n * sxy - sx * sy, 2) / (den * varY) : 1;
                    double ax = x0, bx = x1, ay = slope * ax + cut, by = slope * bx + cut;
                    // Kept inside the plot.
                    if (Math.Abs(slope) > 1e-12)
                    {
                        if (ay < y0) { ax = (y0 - cut) / slope; ay = y0; } else if (ay > y1) { ax = (y1 - cut) / slope; ay = y1; }
                        if (by < y0) { bx = (y0 - cut) / slope; by = y0; } else if (by > y1) { bx = (y1 - cut) / slope; by = y1; }
                    }
                    dynamic trend = Rule(p, px + pw * (ax - x0) / (x1 - x0), py + ph - ph * (ay - y0) / (y1 - y0), px + pw * (bx - x0) / (x1 - x0), py + ph - ph * (by - y0) / (y1 - y0), s == 0 ? t.Accent : tones[s], 1.75);
                    try { trend.Line.DashStyle = 4; } catch (Exception) { }
                    string sign = cut < 0 ? " − " : " + ";
                    Label(p, px + pw - 250, py + 4 + s * 16, 244, 14, "y = " + slope.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "x" + sign + Math.Abs(cut).ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) + "，R² = " + r2.ToString("0.000", System.Globalization.CultureInfo.InvariantCulture), 10, s == 0 ? t.Accent : tones[s], false, t.BodyFont, 3, 1, null);
                }
            }
        }
        bool withValues = chart.TryGetValue("endValues", out raw) && raw != null && (Truthy(raw) || Convert.ToString(raw).ToLowerInvariant() == "true");
        Tips(p, tips, names, tones, px + pw + 8, py + 8, py + ph - 8, margin - 10, withValues);
        // Lines to read against: a threshold, a moment in time.
        IList marks = chart.TryGetValue("lines", out raw) ? raw as IList : null;
        if (marks != null)
        {
            foreach (object mark in marks)
            {
                Dictionary<string, object> at = mark as Dictionary<string, object>;
                if (at == null) continue;
                object value;
                string text = at.TryGetValue("text", out value) && value != null ? Convert.ToString(value) : null;
                if (at.TryGetValue("y", out value) && value != null)
                {
                    double gy = Number(value);
                    if (gy < y0 || gy > y1) continue;
                    double sy = py + ph - ph * (gy - y0) / (y1 - y0);
                    dynamic rule = Rule(p, px, sy, px + pw, sy, t.Accent, 1);
                    try { rule.Line.DashStyle = 4; } catch (Exception) { }
                    if (text != null) Label(p, px + 6, sy - 16, pw - 12, 14, text, 9.5, t.Accent, false, t.BodyFont, 1, 1, null);
                }
                else if (at.TryGetValue("x", out value) && value != null)
                {
                    double gx = Number(value);
                    if (gx < x0 || gx > x1) continue;
                    double sx = px + pw * (gx - x0) / (x1 - x0);
                    dynamic rule = Rule(p, sx, py, sx, py + ph, t.Accent, 1);
                    try { rule.Line.DashStyle = 4; } catch (Exception) { }
                    if (text != null) Label(p, sx + 5, py + 2, 200, 14, text, 9.5, t.Accent, false, t.BodyFont, 1, 1, null);
                }
            }
        }
        // Words set at a point of the plot, with a mark on the point.
        IList said = chart.TryGetValue("annotations", out raw) ? raw as IList : null;
        if (said != null)
        {
            foreach (object one in said)
            {
                Dictionary<string, object> at = one as Dictionary<string, object>;
                object value;
                if (at == null || !at.TryGetValue("text", out value) || value == null) continue;
                string text = Convert.ToString(value);
                double gx = at.TryGetValue("x", out value) ? Number(value) : (x0 + x1) / 2, gy = at.TryGetValue("y", out value) ? Number(value) : (y0 + y1) / 2;
                double sx = px + pw * (Math.Max(x0, Math.Min(x1, gx)) - x0) / (x1 - x0), sy = py + ph - ph * (Math.Max(y0, Math.Min(y1, gy)) - y0) / (y1 - y0);
                Dot(p, sx, sy, 4, t.Accent);
                // The words go to the side that has room.
                bool left = sx > px + pw * 0.62, below = sy < py + 44;
                double tw = Math.Min(240, pw * 0.45), tx = left ? sx - tw - 10 : sx + 10, ty = below ? sy + 10 : sy - 34;
                Rule(p, sx, sy, left ? tx + tw : tx, ty + (below ? 0 : 26), t.Accent, 0.75);
                Label(p, tx, ty, tw, 28, text, 10.5, t.Accent, true, t.BodyFont, left ? 3 : 1, below ? 1 : 4, null).Name = "Annotation";
            }
        }
    }

    /// The names of lines at their ends, moved apart where two would lie on one another.
    static void Tips(Page p, List<double[]> tips, List<string> names, string[] tones, double x, double top, double bottom, double width, bool values)
    {
        if (tips.Count == 0) return;
        Theme t = p.T;
        tips.Sort(delegate(double[] a, double[] b) { return a[1].CompareTo(b[1]); });
        double gap = values ? 30 : 17;
        // Downwards so that none overlaps the one above; then upwards, should the last have left the plot.
        for (int i = 1; i < tips.Count; i++) if (tips[i][1] < tips[i - 1][1] + gap) tips[i][1] = tips[i - 1][1] + gap;
        if (tips[tips.Count - 1][1] > bottom) { tips[tips.Count - 1][1] = bottom; for (int i = tips.Count - 2; i >= 0; i--) if (tips[i][1] > tips[i + 1][1] - gap) tips[i][1] = tips[i + 1][1] - gap; }
        foreach (double[] tip in tips)
        {
            int s = (int)tip[0];
            string text = names[s] + (values ? "\r" + Short(tip[2]) : "");
            dynamic label = Label(p, x, tip[1] - (values ? 15 : 8), width, values ? 30 : 16, text, 10, tones[s] == t.Line ? t.Text : tones[s], true, t.BodyFont, 1, 3, null);
            label.Name = "Series " + (s + 1);
        }
    }

    /// Shares of a whole as a pie or a ring, drawn from shapes, with the parts listed beside it.
    static void DrawPie(Page p, Dictionary<string, object> chart, bool ring, double x, double y, double w, double h)
    {
        Theme t = p.T;
        object raw;
        IList categories = chart.TryGetValue("categories", out raw) ? raw as IList : null;
        IList series = chart.TryGetValue("series", out raw) ? raw as IList : null;
        IList list = chart.TryGetValue("values", out raw) ? raw as IList : null;
        if (list == null && series != null && series.Count > 0)
        {
            Dictionary<string, object> one = series[0] as Dictionary<string, object>;
            list = one != null && one.TryGetValue("values", out raw) ? raw as IList : series[0] as IList;
        }
        if (categories == null || list == null || categories.Count == 0) throw new Fail("BAD_ARGS", "A pie chart needs \"categories\": [..] and \"values\": [..] (or series: [{values}]).");
        int n = Math.Min(Math.Min(categories.Count, list.Count), 8);
        double[] values = new double[n];
        double total = 0;
        for (int i = 0; i < n; i++) { values[i] = Math.Max(0, Number(list[i])); total += values[i]; }
        if (total <= 0) throw new Fail("BAD_ARGS", "The values of the pie chart add up to nothing.");
        string unit = chart.TryGetValue("unit", out raw) && raw != null ? Convert.ToString(raw) : null;
        Decimals = chart.TryGetValue("decimals", out raw) && raw != null ? (int)Number(raw) : -1;
        string[] tones = Palette(p);
        double d = Math.Min(h - 8, w * 0.5), cx = x + d / 2 + 8, cy = y + h / 2;
        double angle = -90;
        for (int i = 0; i < n; i++)
        {
            double sweep = 360 * values[i] / total;
            if (sweep <= 0.05) continue;
            if (sweep >= 359.9) { Dot(p, cx, cy, d / 2, tones[i]); break; }
            dynamic slice = p.Slide.Shapes.AddShape(142, PX(p, cx - d / 2), PY(p, cy - d / 2), SX(p, d), SX(p, d));
            double from = angle, to = angle + sweep;
            // The shape counts its angles from −180 to 180.
            while (from > 180) from -= 360;
            while (to > 180) to -= 360;
            try { slice.Adjustments[1] = (float)from; slice.Adjustments[2] = (float)to; } catch (Exception) { }
            slice.Fill.Solid(); slice.Fill.ForeColor.RGB = Bgr(tones[i]);
            slice.Line.Visible = -1; slice.Line.ForeColor.RGB = Bgr(t.Bg); slice.Line.Weight = 1.5f * p.S;
            try { slice.Shadow.Visible = 0; } catch (Exception) { }
            slice.Name = "Slice " + (i + 1);
            Track(p, slice);
            angle += sweep;
        }
        if (ring)
        {
            Dot(p, cx, cy, d * 0.29, t.Bg);
            string centre = chart.TryGetValue("center", out raw) && raw != null ? Convert.ToString(raw) : Short(total) + (unit == null ? "" : " " + unit);
            Label(p, cx - d * 0.27, cy - 22, d * 0.54, 44, centre, 20, t.Primary, true, t.TitleFont, 2, 3, null);
        }
        // The parts, largest first as given: name, value, share.
        double lx = x + d + 44, lw = x + w - lx, row = Math.Min(40, (h - 8) / n), ly = y + (h - row * n) / 2;
        for (int i = 0; i < n; i++)
        {
            Block(p, lx, ly + row / 2 - 6, 12, 12, tones[i], false);
            Label(p, lx + 22, ly, lw - 150, row, Convert.ToString(categories[i]), 12, t.Text, false, t.BodyFont, 1, 3, null);
            Label(p, lx + lw - 128, ly, 70, row, Short(values[i]) + (unit == null ? "" : " " + unit), 11, t.Muted, false, t.BodyFont, 3, 3, null);
            Label(p, lx + lw - 54, ly, 54, row, (100 * values[i] / total).ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + "%", 12, tones[i] == t.Line ? t.Text : tones[i], true, t.BodyFont, 3, 3, null);
            if (i < n - 1) Rule(p, lx, ly + row, lx + lw, ly + row, t.Line, 0.5);
            ly += row;
        }
    }

    static void TableSlide(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        IList data = op.List("data");
        if (data == null || data.Count == 0) throw new Fail("BAD_ARGS", "A table slide needs \"data\": the rows of the table, the header row first.");
        int rows = Math.Min(data.Count, 12), cols = 1;
        foreach (object row in data) { IList cells = row as IList; if (cells != null) cols = Math.Max(cols, cells.Count); }
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        bool math = false;
        foreach (object row in data) { IList cells = row as IList; if (cells != null) foreach (object cell in cells) if (cell != null && Convert.ToString(cell).IndexOf('$') >= 0) math = true; }
        double usable = h - (callout ? 54 : 0), rowHeight = Math.Min(math ? 56 : 46, usable / rows);
        double tableTop = Settle(y, usable, rowHeight * rows);
        if (math || op.Flag("drawn", false))
        {
            DrawnTable(p, data, rows, cols, x, tableTop, w, rowHeight, rows > 9 ? 12 : rows > 6 ? 14 : 16);
            Callout(p, op, x, y + h - 42, w);
            return;
        }
        dynamic made = p.Slide.Shapes.AddTable(rows, cols, PX(p, x), PY(p, tableTop), SX(p, w), SY(p, rowHeight * rows));
        made.Name = "Table";
        Unit(p);
        Track(p, made);
        dynamic grid = made.Table;
        double size = rows > 9 ? 12 : rows > 6 ? 14 : 16;
        for (int r = 0; r < rows; r++)
        {
            IList cells = data[r] as IList;
            try { grid.Rows[r + 1].Height = SY(p, rowHeight); } catch (Exception) { }
            for (int c = 0; c < cols; c++)
            {
                dynamic cell = grid.Cell(r + 1, c + 1);
                dynamic range = cell.Shape.TextFrame.TextRange;
                string value = cells != null && c < cells.Count && cells[c] != null ? Convert.ToString(cells[c], System.Globalization.CultureInfo.InvariantCulture) : "";
                range.Text = value;
                range.Font.Size = (float)(size * p.S);
                range.Font.Name = t.BodyFont;
                try { range.Font.NameFarEast = t.BodyFont; } catch (COMException) { }
                range.Font.Bold = r == 0 ? -1 : 0;
                range.Font.Color.RGB = Bgr(r == 0 ? (t.Dark ? t.Bg : "#FFFFFF") : t.Text);
                range.ParagraphFormat.Alignment = c == 0 ? 1 : 2;
                cell.Shape.TextFrame.VerticalAnchor = 3;
                cell.Shape.Fill.Visible = -1;
                cell.Shape.Fill.Solid();
                cell.Shape.Fill.ForeColor.RGB = Bgr(r == 0 ? t.Primary : r % 2 == 0 ? t.Surface : t.Bg);
                for (int side = 1; side <= 4; side++)
                {
                    try { dynamic edge = cell.Borders[side]; edge.Visible = -1; edge.ForeColor.RGB = Bgr(side == 3 ? t.Line : r == 0 ? t.Primary : r % 2 == 0 ? t.Surface : t.Bg); edge.Weight = side == 3 ? 0.75f : 0.25f; }
                    catch (Exception) { }
                }
                if (value.IndexOf('$') >= 0) { try { PptFlatMath(range); } catch (Exception) { } }
            }
        }
        p.Motion.Add(made);
        Callout(p, op, x, y + h - 42, w);
    }

    static double Number(object raw)
    {
        if (raw == null) return 0;
        double value;
        return double.TryParse(Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out value) ? value : 0;
    }

    /// Decimals the figures of the chart being drawn show (-1 = as many as the value needs).
    static int Decimals = -1;

    static string Short(double value)
    {
        if (Decimals >= 0) return value.ToString("F" + Math.Min(Decimals, 4), System.Globalization.CultureInfo.InvariantCulture);
        double abs = Math.Abs(value);
        if (abs >= 100 || value == Math.Round(value)) return Math.Round(value).ToString("0", System.Globalization.CultureInfo.InvariantCulture);
        return value.ToString(abs >= 10 ? "0.#" : "0.##", System.Globalization.CultureInfo.InvariantCulture);
    }

    /// A round number at or above the largest value, and how many steps the scale takes to reach it.
    static double Ceiling(double max, out int steps)
    {
        steps = 4;
        if (max <= 0) return 1;
        double power = Math.Pow(10, Math.Floor(Math.Log10(max)));
        foreach (double m in new double[] { 1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10 })
        {
            if (m * power >= max) { steps = m == 1.2 || m == 3 || m == 6 ? 3 : m == 1.5 || m == 2.5 || m == 5 || m == 10 ? 5 : 4; return m * power; }
        }
        return 10 * power;
    }

    /// A chart drawn from shapes: column, bar (lying) or line. One series may point at one category.
    static void ChartSlide(Page p, Bag op, double x, double y, double w, double h)
    {
        Theme t = p.T;
        Dictionary<string, object> chart = op.Raw("chart") as Dictionary<string, object>;
        if (chart == null) throw new Fail("BAD_ARGS", "A chart slide needs \"chart\": {type, categories:[..], series:[{name, values:[..]}]}.");
        IList side = Items(op, "side", "points");
        bool callout = !string.IsNullOrEmpty(op.Str("callout", null));
        double ch = h - (callout ? 54 : 0), cw = side.Count > 0 ? w * 0.6 : w;
        Unit(p);
        Draw(p, chart, x, y, cw, ch);
        double sx = x + cw + 30, sy = y + 6, sw = w - cw - 30;
        for (int i = 0; i < Math.Min(side.Count, 4); i++)
        {
            Unit(p);

            object raw = side[i];
            string head = Field(raw, "head"), text = Field(raw, "text") ?? "";
            Block(p, sx, sy, 34, 3, Pick(p, i), false);
            string value = Field(raw, "value");
            if (value != null)
            {
                string unit = Field(raw, "unit");
                if (head != null) { Label(p, sx, sy + 10, sw, 18, head, 11.5, t.Muted, false, t.BodyFont, 1, 1, null); sy += 20; }
                dynamic big = Label(p, sx, sy + 8, sw, 46, value + (unit == null ? "" : " " + unit), 34, Pick(p, i), true, t.TitleFont, 1, 3, null);
                if (unit != null) { try { dynamic tail = big.TextFrame.TextRange.Characters(value.Length + 1, unit.Length + 1); tail.Font.Size = (float)(15 * p.S); tail.Font.Bold = 0; tail.Font.Color.RGB = Bgr(t.Text); } catch (Exception) { } }
                Words(p, sx, sy + 56, sw, 40, text, 12, t.Muted, false);
                p.Motion.Add(big);
                sy += 104;
                continue;
            }
            if (head != null) { Words(p, sx, sy + 10, sw, 22, head, 14.5, t.Text, true); sy += 26; }
            dynamic words = Words(p, sx, sy + 10, sw, 62, text, 13, head != null ? t.Muted : t.Text, false);
            p.Motion.Add(words);
            sy += head != null ? 62 : 74;
        }
        Callout(p, op, x, y + h - 42, w);
    }

    static void Draw(Page p, Dictionary<string, object> chart, double x, double y, double w, double h)
    {
        Theme t = p.T;
        object raw;
        string type = chart.TryGetValue("type", out raw) && raw != null ? Convert.ToString(raw).ToLowerInvariant() : "column";
        if (type == "scatter" || type == "curve" || type == "xy" || type == "function") { DrawXY(p, chart, type, x, y, w, h); return; }
        if (type == "pie" || type == "donut" || type == "doughnut" || type == "ring") { DrawPie(p, chart, type != "pie", x, y, w, h); return; }
        IList categories = chart.TryGetValue("categories", out raw) ? raw as IList : null;
        IList series = chart.TryGetValue("series", out raw) ? raw as IList : null;
        if (categories == null || series == null || categories.Count == 0 || series.Count == 0) throw new Fail("BAD_ARGS", "The chart needs \"categories\" and \"series\": [{name, values}].");
        string unit = chart.TryGetValue("unit", out raw) && raw != null ? Convert.ToString(raw) : null;
        Decimals = chart.TryGetValue("decimals", out raw) && raw != null ? (int)Number(raw) : -1;
        int highlight = -1;
        if (chart.TryGetValue("highlight", out raw) && raw != null)
        {
            string wanted = Convert.ToString(raw, System.Globalization.CultureInfo.InvariantCulture);
            int number;
            if (int.TryParse(wanted, out number)) highlight = number - 1;
            for (int i = 0; i < categories.Count; i++) if (Convert.ToString(categories[i]) == wanted) highlight = i;
        }
        int n = categories.Count, m = Math.Min(series.Count, 4);
        List<double[]> values = new List<double[]>();
        List<string> names = new List<string>();
        double max = 0;
        for (int s = 0; s < m; s++)
        {
            Dictionary<string, object> one = series[s] as Dictionary<string, object>;
            IList list = one != null && one.TryGetValue("values", out raw) ? raw as IList : series[s] as IList;
            double[] row = new double[n];
            for (int i = 0; i < n && list != null && i < list.Count; i++) { row[i] = Number(list[i]); max = Math.Max(max, row[i]); }
            values.Add(row);
            names.Add(one != null && one.TryGetValue("name", out raw) && raw != null ? Convert.ToString(raw) : "");
        }
        string[] tones = { t.Primary, t.Accent, t.Muted, t.Line };
        int steps;
        double top = Ceiling(max, out steps), floor = 0;
        if (type == "line")
        {
            double min = max;
            foreach (double[] row in values) foreach (double v in row) min = Math.Min(min, v);
            // Values that sit close together would draw as a flat line from zero: the scale starts just below them.
            if (min > top * 0.45)
            {
                double step = top / steps;
                floor = Math.Floor(min / step) * step;
                if (floor >= min) floor -= step;
                floor = Math.Max(0, floor);
                steps = Math.Max(2, (int)Math.Round((top - floor) / step));
            }
        }
        bool legend = m > 1;
        if (unit != null) Label(p, x, y, 200, 14, unit, 9.5, t.Muted, false, t.BodyFont, 1, 1, null);
        if (legend)
        {
            double lx = x + w;
            for (int s = m - 1; s >= 0; s--)
            {
                double width = 16 + names[s].Length * 11 + 14;
                lx -= width;
                Block(p, lx, y + 3, 9, 9, tones[s], false);
                Label(p, lx + 14, y, width - 14, 14, names[s], 10, t.Muted, false, t.BodyFont, 1, 1, null);
            }
        }
        double head = unit != null || legend ? 22 : 6;
        if (type == "bar")
        {
            double labelWidth = 0;
            foreach (object c in categories) labelWidth = Math.Max(labelWidth, Convert.ToString(c).Length * 13.0);
            labelWidth = Math.Min(Math.Max(labelWidth, 40), w * 0.3);
            double plotX = x + labelWidth + 10, plotW = w - labelWidth - 10 - 56, row = (h - head) / n, bar = Math.Min(26, row * 0.62 / m);
            for (int i = 0; i < n; i++)
            {
                double cy = y + head + i * row + (row - bar * m) / 2;
                Label(p, x, y + head + i * row, labelWidth, row, Convert.ToString(categories[i]), 11.5, t.Text, false, t.BodyFont, 3, 3, null);
                for (int s = 0; s < m; s++)
                {
                    double length = Math.Max(1, plotW * values[s][i] / top);
                    string tone = m == 1 && i == highlight ? t.Accent : tones[s];
                    dynamic shape = Block(p, plotX, cy + s * bar, length, bar - 2, tone, false);
                    Label(p, plotX + length + 6, cy + s * bar - 2, 60, bar, Short(values[s][i]), 10.5, tone == t.Accent ? t.Accent : t.Text, m == 1 && i == highlight, t.BodyFont, 1, 3, null);
                    if (s == 0) p.Motion.Add(shape);
                }
            }
            Rule(p, plotX, y + head, plotX, y + h, t.Line, 0.75);
            return;
        }
        double axis = 40, foot = 24;
        double px = x + axis, pw = w - axis - 8, py = y + head + 12, ph = h - head - 12 - foot;
        for (int k = 0; k <= steps; k++)
        {
            double gy = py + ph - ph * k / steps;
            Rule(p, px, gy, px + pw, gy, t.Line, k == 0 ? 1 : 0.5);
            Label(p, x, gy - 8, axis - 6, 16, Short(floor + (top - floor) * k / steps), 9.5, t.Muted, false, t.BodyFont, 3, 3, null);
        }
        double slot = pw / n;
        for (int i = 0; i < n; i++) Label(p, px + i * slot, py + ph + 5, slot, 16, Convert.ToString(categories[i]), 10.5, t.Muted, false, t.BodyFont, 2, 1, null);
        if (type == "line")
        {
            for (int s = 0; s < m; s++)
            {
                dynamic builder = null;
                for (int i = 0; i < n; i++)
                {
                    float fx = PX(p, px + slot * (i + 0.5)), fy = PY(p, py + ph - ph * (values[s][i] - floor) / (top - floor));
                    if (builder == null) builder = p.Slide.Shapes.BuildFreeform(0, fx, fy);
                    else builder.AddNodes(0, 0, fx, fy);
                }
                if (builder != null && n > 1)
                {
                    dynamic path = Track(p, builder.ConvertToShape());
                    path.Fill.Visible = 0;
                    path.Line.ForeColor.RGB = Bgr(tones[s]);
                    path.Line.Weight = 2.25f * p.S;
                    p.Motion.Add(path);
                }
                for (int i = 0; i < n; i++)
                {
                    double cx = px + slot * (i + 0.5), cy = py + ph - ph * (values[s][i] - floor) / (top - floor);
                    Dot(p, cx, cy, 3.5, m == 1 && i == highlight ? t.Accent : tones[s]);
                    if (m == 1 || i == n - 1) Label(p, cx - 34, cy - 20, 68, 14, Short(values[s][i]), 10, m == 1 && i == highlight ? t.Accent : t.Text, false, t.BodyFont, 2, 1, null);
                }
            }
            return;
        }
        double group = Math.Min(slot * 0.62, 84 * m), barWidth = group / m;
        for (int i = 0; i < n; i++)
        {
            for (int s = 0; s < m; s++)
            {
                double bh = Math.Max(1, ph * values[s][i] / top), bx = px + i * slot + (slot - group) / 2 + s * barWidth;
                string tone = m == 1 && i == highlight ? t.Accent : tones[s];
                dynamic shape = Block(p, bx, py + ph - bh, barWidth - (m > 1 ? 2 : 0), bh, tone, false);
                if (m <= 2) Label(p, bx - 14, py + ph - bh - 16, barWidth + 28, 14, Short(values[s][i]), 10, tone == t.Accent ? t.Accent : t.Text, m == 1 && i == highlight, t.BodyFont, 2, 1, null);
                if (s == 0 && i < 10) p.Motion.Add(shape);
            }
        }
    }

    // ───────────────────────── edit batches ─────────────────────────

    static object Edit(string kind, dynamic app, dynamic doc, Bag a)
    {
        IList ops = a.List("ops");
        if (ops == null || ops.Count == 0) throw new Fail("BAD_ARGS", "\"ops\" must list at least one operation.");
        List<object> done = new List<object>();
        Dictionary<string, object> result = new Dictionary<string, object>();
        // The settings give the two modes, until the user flips one on the card.
        if (!Card.FollowChosen) Following = a.Flag("follow", false);
        if (!Card.TypingChosen) Typing = a.Flag("typing", false);
        bool card = a.Flag("card", false);
        bool quiet = false;
        try { string whole = (string)doc.FullName; quiet = Hidden.Contains(whole) || (a.Flag("silent", false) && !Revealed.Contains(whole)); } catch (Exception) { }
        if (quiet) { Following = false; Typing = false; card = false; }
        CardOn = card;
        if (!quiet) Remember(doc); else DocWindow = IntPtr.Zero;
        if (card && !cardStarted) { cardStarted = true; Card.Start(); }
        string docName = (string)doc.Name;
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
                    if (card) Card.Hold = true;
                    if (card) Card.Report("AI 正在编辑 " + docName + (ops.Count > 1 ? " · " + (i + 1) + "/" + ops.Count : ""));
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
            if (kind == "ppt")
            {
                Advised = false; SmallOnSlide = false; Small.Clear(); Cut.Clear(); Focus = null;
                if (Renumber) Numbers(doc);
                try { MathDone(); } catch (Exception) { }
                if (PptMathFailed > 0) done.Add("(NOTE " + PptMathFailed + " formula(s) could not be built and were left as text between dollar signs: rewrite them more simply)");
                PptMathFailed = 0;
            }
            if (kind == "word")
            {
                Advised = false; SmallOnSlide = false; Small.Clear();
                try { FinishMath(doc); } catch (Exception) { }
                try { MathDone(); } catch (Exception) { }
            }
            if (kind == "word" && Cramped)
            {
                int roomy = Roomy(doc);
                if (roomy > 0) done.Add("(" + roomy + " paragraph(s) with formulas or pictures: exact line height turned into a minimum of the same height, so nothing tall is cut off)");
            }
            if (kind == "word" && Listed)
            {
                Listed = false;
                string uncited = Uncited(doc);
                if (uncited.Length > 0) done.Add(uncited);
            }
            if (undo != null) { try { undo.EndCustomRecord(); } catch (COMException) { } }
        }
        if (card) { Card.Hold = false; Card.Report("AI 已编辑 " + docName, Linger); }
        result["done"] = done;
        result["total"] = ops.Count;
        return result;
    }

    // ───────────────────────── rendering ─────────────────────────

    /// One Word page as a picture of the given width.
    static Bitmap WordPage(dynamic pages, int page, int width)
    {
        byte[] bits = (byte[])pages[page].EnhMetaFileBits;
        using (MemoryStream stream = new MemoryStream(bits))
        using (Metafile meta = new Metafile(stream))
        {
            int height = (int)Math.Round(width * (double)meta.Height / meta.Width);
            // Drawn straight at the final size, hairlines (table borders) fall between pixels and vanish:
            // draw the page large first, then scale the picture down.
            int factor = Math.Max(2, (int)Math.Ceiling(3000.0 / width));
            Bitmap bitmap = new Bitmap(width, height);
            using (Bitmap large = new Bitmap(width * factor, height * factor))
            {
                using (Graphics g = Graphics.FromImage(large))
                {
                    g.Clear(Color.White);
                    g.DrawImage(meta, 0, 0, large.Width, large.Height);
                }
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.DrawImage(large, 0, 0, width, height);
                }
            }
            return bitmap;
        }
    }

    /// Several pages or slides side by side in one picture, each with its number: the whole document at a glance.
    static object Sheet(string kind, dynamic doc, Bag a)
    {
        string path = Path.GetFullPath(a.Need("out"));
        int cell = Math.Max(300, Math.Min(1000, a.Int("width", 640)));
        dynamic pages = kind == "word" ? doc.Windows[1].Panes[1].Pages : null;
        int total = kind == "word" ? (int)pages.Count : (int)doc.Slides.Count;
        if (total == 0) throw new Fail("ANCHOR_MISSING", "There is nothing to picture yet.");
        int first = Math.Max(1, a.Int("from", 1)), last = Math.Min(total, a.Int("to", total));
        if (first > total) throw new Fail("ANCHOR_MISSING", "There is no " + (kind == "word" ? "page " : "slide ") + first + " (there are " + total + ").");
        last = Math.Min(last, first + 11);
        int count = last - first + 1, columns = Math.Min(count, kind == "word" ? 3 : 2), rows = (count + columns - 1) / columns;
        List<Bitmap> pictures = new List<Bitmap>();
        string temp = path + ".slide.png";
        try
        {
            for (int n = first; n <= last; n++)
            {
                if (kind == "word") pictures.Add(WordPage(pages, n, cell));
                else
                {
                    double ratio = (double)doc.PageSetup.SlideHeight / (double)doc.PageSetup.SlideWidth;
                    doc.Slides[n].Export(temp, "PNG", cell, (int)Math.Round(cell * ratio));
                    using (Image loaded = Image.FromFile(temp)) pictures.Add(new Bitmap(loaded));
                }
            }
            int gap = 14, label = 30, height = 0;
            foreach (Bitmap picture in pictures) height = Math.Max(height, picture.Height);
            using (Bitmap sheet = new Bitmap(columns * cell + (columns + 1) * gap, rows * (height + label + gap) + gap))
            {
                using (Graphics g = Graphics.FromImage(sheet))
                using (Font font = new Font("Microsoft YaHei UI", 14f, FontStyle.Bold, GraphicsUnit.Pixel))
                using (Pen border = new Pen(Color.FromArgb(150, 150, 150)))
                {
                    g.Clear(Color.FromArgb(232, 232, 236));
                    for (int k = 0; k < pictures.Count; k++)
                    {
                        int x = gap + (k % columns) * (cell + gap), y = gap + (k / columns) * (height + label + gap);
                        g.DrawString((kind == "word" ? "page " : "slide ") + (first + k), font, Brushes.Black, x, y + 4);
                        g.DrawImage(pictures[k], x, y + label);
                        g.DrawRectangle(border, x, y + label, pictures[k].Width - 1, pictures[k].Height - 1);
                    }
                }
                sheet.Save(path, ImageFormat.Png);
            }
        }
        finally
        {
            foreach (Bitmap picture in pictures) picture.Dispose();
            try { File.Delete(temp); } catch (Exception) { }
        }
        Dictionary<string, object> result = new Dictionary<string, object>();
        result["what"] = (kind == "word" ? "pages " : "slides ") + first + "–" + last + " of " + total;
        result["first"] = first; result["last"] = last; result["total"] = total;
        result["path"] = path;
        using (Image saved = Image.FromFile(path)) { result["width"] = saved.Width; result["height"] = saved.Height; }
        return result;
    }

    static object Render(string kind, dynamic app, dynamic doc, Bag a)
    {
        if (a.Flag("sheet", false) && kind != "excel") return Sheet(kind, doc, a);
        string path = Path.GetFullPath(a.Need("out"));
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
                // Drawn straight at the final size, hairlines (table borders) fall between pixels and vanish:
                // draw the page large first, then scale the picture down.
                int factor = Math.Max(2, (int)Math.Ceiling(3000.0 / width));
                using (Bitmap large = new Bitmap(width * factor, height * factor))
                using (Bitmap bitmap = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(large))
                    {
                        g.Clear(Color.White);
                        g.DrawImage(meta, 0, 0, large.Width, large.Height);
                    }
                    using (Graphics g = Graphics.FromImage(bitmap))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(large, 0, 0, width, height);
                    }
                    bitmap.Save(path, ImageFormat.Png);
                }
            }
            result["what"] = "page " + page + " of " + total;
        }
        else
        {
            dynamic sheet = Sheet(doc, a);
            if (a.Has("chart"))
            {
                dynamic holder = ChartByName(sheet, a.Need("chart"));
                holder.Chart.Export(path, "PNG");
                result["what"] = "chart \"" + (string)holder.Name + "\" on " + (string)sheet.Name;
                result["path"] = path;
                using (Image made = Image.FromFile(path)) { result["width"] = made.Width; result["height"] = made.Height; }
                return result;
            }
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
