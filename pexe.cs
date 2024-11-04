// csc /target:winexe /win32icon:pexe.ico /out:pexe.exe pexe.cs

using System;
using System.Linq;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

using System.Reflection;
using System.Runtime.CompilerServices;

using System.Collections.Generic;

//-----------------------------------------------------------------------------
public class Window {
    [DllImport("user32.dll")]
    private static extern IntPtr GetDC(IntPtr hwnd);

    [DllImport("gdi32.dll")]
    private static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

    [DllImport("user32.dll")]
    private static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

    [DllImport("user32.dll")]
    extern static bool SetProcessDPIAware();

    private const int LOGPIXELSX = 88;
    private const int LOGPIXELSY = 90;

    static public double DPI_SCALING;

    static Window() {
        SetProcessDPIAware();
        IntPtr hdc = GetDC(IntPtr.Zero);
        int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
        DPI_SCALING = dpiX / 96.0;
        ReleaseDC(IntPtr.Zero, hdc);
    }

    static public int DpiScaledInt(int value) {
        return (int)(value * DPI_SCALING);
    }
    static public int DpiScaledCeil(int value) {
        return (int)Math.Ceiling(value * DPI_SCALING);
    }

    public class TaskBar {
        [DllImport("shell32.dll", SetLastError = true)]
        private static extern IntPtr SHAppBarMessage(uint dwMessage, ref APPBARDATA pData);
        private const uint ABM_GETTASKBARPOS = 0x00000005;

        [StructLayout(LayoutKind.Sequential)]
        private struct APPBARDATA {
            public int cbSize;
            public IntPtr hWnd;
            public uint uCallbackMessage;
            public uint uEdge;
            public RECT rc;
            public int lParam;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        static public int LEFT;
        static public int TOP;
        static public int RIGHT;
        static public int BOTTOM;
        static public int HEIGHT;

        static TaskBar() {
            APPBARDATA appBarData = new APPBARDATA();
            appBarData.cbSize = Marshal.SizeOf(appBarData);

            IntPtr result = SHAppBarMessage(ABM_GETTASKBARPOS, ref appBarData);

            if (result != IntPtr.Zero) {
                LEFT   = appBarData.rc.left;
                TOP    = appBarData.rc.top;
                RIGHT  = appBarData.rc.right;
                BOTTOM = appBarData.rc.bottom;
                HEIGHT = BOTTOM - TOP;
            }
        }
    }
}


public class IniFile
{
    static private string s_path;

    [DllImport("kernel32")]
    private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);

    [DllImport("kernel32")]
    private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

    static public void Open(string p) {
        s_path = p;
    }

    static public void Write(string key, string value) {
        WritePrivateProfileString("Settings", key, value, s_path);
    }

    static public string Read(string key, string def = "") {
        StringBuilder retVal = new StringBuilder(255);
        GetPrivateProfileString("Settings", key, def, retVal, 255, s_path);
        return retVal.ToString();
    }

    static public int ReadInt(string key, int def = 0) {
        try {
            string val = IniFile.Read(key, def.ToString());
            return int.Parse(val);
        }
        catch {
            return def;
        }
    }

    static public double ReadDouble(string key, double def = 0.0) {
        try {
            string val = IniFile.Read(key, def.ToString());
            return double.Parse(val);
        }
        catch {
            return def;
        }
    }
}

//-----------------------------------------------------------------------------

public class History {
    static private readonly string   HIST_PATH = IniFile.Read("HistPath", AppDomain.CurrentDomain.BaseDirectory + "\\.pexe_history.txt");
    static private readonly Encoding ENCODING  = Encoding.GetEncoding("Shift_JIS");

    private readonly string[] hists;
    private int head;
    private int tail;
    private int cnt;
    private int idx;
    private int size;

    public History(int sz) {
        if (sz < 0) {
            throw new InvalidOperationException("hists empty");
        }
        size = sz;
        hists = new string[size];
        head = 0;
        tail = 0;
        cnt  = 0;
        idx  = -1;

        if (File.Exists(HIST_PATH)) {
            StreamReader reader = new StreamReader(HIST_PATH, ENCODING);
            for(int i1 = 0; !reader.EndOfStream; i1++) {
                string  line = reader.ReadLine();
                this.Add(line);
            }
            reader.Close();
        }
    }

    public void Close() {
        if ((0 < cnt) && (0 < HIST_PATH.Length)) {
            StreamWriter writer = new StreamWriter(HIST_PATH, false, ENCODING);
            for (int i1 = tail; ; i1 = nextIdx(i1)) {
                writer.WriteLine(hists[i1]);
                if (i1 == lastIdx) {
                    break;
                }
            }
            writer.Close();
        }
    }

    private int nextIdx(int idx) {
        return (idx + 1) % size;
    }
    private int prevIdx(int idx) {
        return (idx - 1 + size) % size;
    }
    private int lastIdx { get { return prevIdx(head);}}

    public void Add(string cmd) {
        if (cmd.Length < 1) {
            return;
        }

        if (cnt < size) {
            cnt++;
        }
        else {
            tail = nextIdx(tail);
        }

        if (cmd.Equals(hists[lastIdx])) {
            return;
        }

        hists[head] = cmd;
        head = nextIdx(head);
    }

    public string Prev() {
        if (cnt < 1) {
            throw new InvalidOperationException("hists empty");
        }

        if (idx == -1) {
            idx = lastIdx;
        }
        else if (idx == tail) {
            throw new InvalidOperationException("no more hist");
        }
        else {
            idx = prevIdx(idx);
        }

        return hists[idx];
    }

    public string Next() {
        if (cnt < 1) {
            throw new InvalidOperationException("hists empty");
        }

        if (idx == -1) {
            throw new InvalidOperationException("no more hist");
        }
        else if (idx == lastIdx) {
            idx = -1;
            return "";
        }
        else {
            idx = nextIdx(idx);
        }

        return hists[idx];
    }

    public void ResetPos() {
        idx = -1;
    }

    public string LastMatch(string prefix) {
        if (prefix.Length < 1) {
            throw new InvalidOperationException("prefix empty");
        }
        if (cnt < 1) {
            throw new InvalidOperationException("hists empty");
        }

        for (int i1 = lastIdx; ; i1 = prevIdx(i1)) {
            if (hists[i1].StartsWith(prefix)) {
                return hists[i1];
            }

            if (i1 == tail) {
                break;
            }
        }

        throw new InvalidOperationException("no match");
    }
}

//-----------------------------------------------------------------------------
public class ColorHandle {
    private struct COLOR {
        public Color normal;
        public Color light;
    }

    static private readonly COLOR[] COLOR_TABLE = {
        new COLOR { normal = ColorTranslator.FromHtml("#0C0C0C"), light = ColorTranslator.FromHtml("#767676") }, // black
        new COLOR { normal = ColorTranslator.FromHtml("#C50F1F"), light = ColorTranslator.FromHtml("#E74856") }, // red
        new COLOR { normal = ColorTranslator.FromHtml("#13A10E"), light = ColorTranslator.FromHtml("#16C60C") }, // green
        new COLOR { normal = ColorTranslator.FromHtml("#C19C00"), light = ColorTranslator.FromHtml("#F9F1A5") }, // yellow
        new COLOR { normal = ColorTranslator.FromHtml("#0037DA"), light = ColorTranslator.FromHtml("#3B78FF") }, // blue
        new COLOR { normal = ColorTranslator.FromHtml("#881798"), light = ColorTranslator.FromHtml("#B4009E") }, // magenta
        new COLOR { normal = ColorTranslator.FromHtml("#3A96DD"), light = ColorTranslator.FromHtml("#61D6D6") }, // cyan
        new COLOR { normal = ColorTranslator.FromHtml("#CCCCCC"), light = ColorTranslator.FromHtml("#F2F2F2") }, // white
    };

    static public Color ReverseColor(Color color) {
        return Color.FromArgb(color.ToArgb() ^ 0xffffff);
    }

    static public bool IsAnsiColorCode(string[] codes) {
        if (0 < codes.Length) {
            int colorCode;
            if (int.TryParse(codes[0], out colorCode)) {
                switch ((colorCode / 10) * 10) {
                    case  30:
                    case  40:
                    case  90:
                    case 100:
                        return true;
                    default:
                        break;
                }

            }
        }

        return false;
    }

    static private string ansi8BitToRgb(int c) {
        c = c % 256;
        return "#" + (
            (c < 16)  ? new[]{"000000","800000","008000","808000","000080","800080","008080","C0C0C0","808080","FF0000","00FF00","FFFF00","0000FF","FF00FF","00FFFF","FFFFFF"}[c] :
            (c < 232) ? string.Format("{0:X2}{1:X2}{2:X2}", ((c - 16) / 36 % 6) * 51, ((c - 16) / 6 % 6) * 51, (c - 16) % 6 * 51) :
            string.Format("{0:X2}{0:X2}{0:X2}", (c - 232) * 10 + 8)
        );
    }

    static public bool ParseColorCode(string[] codes, ref Color color, ref Color backColor) {
        bool  colorSpecify = false;
        bool  isBack       = false;
        Color res          = Color.Empty;

        foreach (string code in codes) {
            int colorCode;
            if (int.TryParse(code, out colorCode)) {
                if (colorSpecify) {
                    if ((colorCode == 5) && (3 <= codes.Length)) {
                        if (int.TryParse(codes[2], out colorCode)) {
                            res = ColorTranslator.FromHtml(ColorHandle.ansi8BitToRgb(colorCode));
                            break;
                        }
                    }
                    else if ((colorCode == 2) && (5 <= codes.Length)) {
                        int r,g,b;
                        if (int.TryParse(codes[2], out r)) {
                            if (int.TryParse(codes[3], out g)) {
                                if (int.TryParse(codes[4], out b)) {
                                    res = ColorTranslator.FromHtml(string.Format("#{0:X2}{1:X2}{2:X2}", r, g, b));
                                    break;
                                }
                            }
                        }
                    }
                    else {
                        Console.WriteLine("1other: " + colorCode);
                        return false;
                    }
                }

                bool light = false;
                switch ((colorCode / 10) * 10) {
                    case 30:
                        break;
                    case 90:
                        light = true;
                        break;
                    case 40:
                        isBack = true;
                        break;
                    case 100:
                        isBack = true;
                        light = true;
                        break;
                    default:
                        Console.WriteLine("other: " + colorCode);
                        return false;
                }

                int idx = colorCode % 10;
                if (idx < COLOR_TABLE.Length) {
                    res = light ? COLOR_TABLE[idx].light : COLOR_TABLE[idx].normal;
                    break;
                }
                else if (idx == 8) {
                    colorSpecify = true;
                }
                else {
                    Console.WriteLine("other: " + colorCode);
                    return false;
                }
            }
        }

        if (res != Color.Empty) {
            if (isBack) {
                backColor = res;
            }
            else {
                color = res;
            }
            return true;
        }

        return false;
    }
}

//-----------------------------------------------------------------------------
public class PwshProcess: Process {
    public  List<string>  Result       = new List<string>();

    private StringBuilder outputString = new StringBuilder();
    private bool          executing    = false;

    public PwshProcess() : base() {
        this.StartInfo = new ProcessStartInfo {
            FileName               = "pwsh.exe",
            RedirectStandardInput  = true,
            RedirectStandardOutput = true,
            RedirectStandardError  = true,
            UseShellExecute        = false,
            CreateNoWindow         = true,
            WorkingDirectory       = IniFile.Read("WorkingDirectory", "."),
        };
    }

    public async void Run() {
        Start();

        await StandardInput.WriteLineAsync(
            "$newsize = $(get-host).ui.rawui.buffersize; "
            + "$newsize.width =" + IniFile.Read("OutputColums", "80") + ";"
            + "$(get-host).ui.rawui.buffersize = $newsize"
        );

        BeginOutputReadLine();
        BeginErrorReadLine();
        OutputDataReceived += (sender, e) => { if (e.Data != null) outputString.AppendLine(e.Data); };
        ErrorDataReceived += (sender, e) => { if (e.Data != null) outputString.AppendLine(e.Data); };
    }


    public async Task<string[]> ExecuteCmdAsync(string cmd) {

        if (executing) {
            return null;
        }
        executing = true;

        outputString.Clear();
        Result.Clear();

        string endMarker = Guid.NewGuid().ToString();
        await StandardInput.WriteLineAsync(string.Format("{0}; Write-Output {1}", cmd, endMarker));
        await StandardInput.FlushAsync();

        string[] lines;

        while (true) {
            int idx;
            lines = outputString.ToString().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            if (-1 < (idx = Array.IndexOf(lines, endMarker))) {
                lines[idx] = lines[idx].Replace(endMarker, "");

                Result = new List<string>(new ArraySegment<string>(lines, 1, lines.Length - 2));
                executing = false;

                return Result.ToArray();
            }
            else {
                if (2 < lines.Length) {
                    Result = new List<string>(new ArraySegment<string>(lines, 1, lines.Length - 2));
                }
            }

            await Task.Delay(100);
        }
    }

    public async Task<string[]> TabCompleteAsync(string cmd) {
        string[] lines = await ExecuteCmdAsync(string.Format("$(TabExpansion2 \"{0}\").CompletionMatches.CompletionText", cmd));

        return lines.Where(x => !string.IsNullOrWhiteSpace(x)).ToArray();
    }


    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool GenerateConsoleCtrlEvent(uint dwCtrlEvent, uint dwProcessGroupId);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern bool AttachConsole(uint dwProcessId);

    [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
    private static extern bool FreeConsole();

    private const uint CTRL_C_EVENT = 0;

    public void SendTerm() {
        AttachConsole((uint)Id);
        GenerateConsoleCtrlEvent(CTRL_C_EVENT, 0);
        FreeConsole();
    }
}

//-----------------------------------------------------------------------------
public class CmdInputPanel : Panel {
    private History     history;
    private PopupForm   compForm;
    private PwshProcess pwshProcess;
    private RichTextBox inputBox;
    private Pen         borderPen;
    private Color       bordercolor;

    public delegate void OnKeyDownDeligate(string cmd);
    private Dictionary<Keys, OnKeyDownDeligate> onKeyDown = new Dictionary<Keys, OnKeyDownDeligate>();


    private bool   isProcessing = false;
    private string appendText   = "";
    private string inputText {
        get {
            return appendText != ""
                ? inputBox.Text.Substring(0, Math.Max(0, inputBox.Text.Length - appendText.Length))
                : inputBox.Text;
        }
        set {
            appendText = "";
            inputBox.Text = value;
        }
    }

    public CmdInputPanel(int height, PwshProcess process) : base() {
        BackColor   = Color.White;
        Padding     = new Padding(Window.DpiScaledInt(4), Window.DpiScaledCeil(6),1,1);
        BorderStyle = BorderStyle.None;
        bordercolor = Color.LightGray;
        Height      = height;

        pwshProcess = process;

        inputBox = new RichTextBox() {
            // Font           = new Font("PlemolJP Console", 12),
            Font           = new Font(IniFile.Read("Font", "Arial"), (int)Math.Ceiling(12 / Window.DPI_SCALING)),
            Dock           = DockStyle.Fill,
            BorderStyle    = BorderStyle.None,
            ScrollBars     = RichTextBoxScrollBars.None,
            LanguageOption = RichTextBoxLanguageOptions.UIFonts,
            Multiline      = false,
            WordWrap       = false,
        };
        Controls.Add(inputBox);

        history   = new History(100);
        compForm  = new PopupForm(foreColor:"#111", backColor:"#eee", borderColor:"#ddd");
        borderPen = new Pen(bordercolor, 1);

        inputBox.PreviewKeyDown += (s, e) => { if (e.KeyCode == Keys.Tab) {e.IsInputKey = true;} };
        inputBox.KeyDown += OnKeyDown;

        inputBox.TextChanged += OnTextChanged;
        inputBox.SelectionChanged += OnSelectionChanged;
        inputBox.MouseUp += OnMouseUp;
    }


    public void Quit() {
        history.Close();
        borderPen.Dispose();
    }

    private void OnSelectionChanged(object sender, EventArgs e) {
        if (isProcessing)
            return;

        if (inputText.Length < inputBox.SelectionStart) {
            inputBox.SelectionStart = inputText.Length;
        }
    }

    private void OnTextChanged(object sender, EventArgs e)
    {
        if (isProcessing)
            return;

        isProcessing = true;

        string actualText = inputText;
        int caretPosition = inputBox.SelectionStart;

        inputBox.Clear();
        inputBox.SelectionStart = 0;
        inputBox.SelectionColor = Color.Black;
        inputBox.AppendText(actualText);

        try {
            string hint = history.LastMatch(actualText);
            appendText = hint.Substring(actualText.Length, hint.Length - actualText.Length);
        }
        catch {
            appendText = "";
        }

        inputBox.SelectionStart = actualText.Length;
        inputBox.SelectionColor = Color.Silver;
        inputBox.AppendText(appendText);

        inputBox.SelectionStart = caretPosition;
        inputBox.SelectionLength = 0;
        inputBox.SelectionColor = Color.Black;

        isProcessing = false;
    }

    protected override void OnPaint(PaintEventArgs e) {
        base.OnPaint(e);

        Rectangle rect = new Rectangle(0, 0, Width - 1, Height - 1);
        e.Graphics.DrawRectangle(borderPen, rect);
    }

    private void OnMouseUp(object sender, MouseEventArgs e) {
        if (e.Button == MouseButtons.Right) {
            inputText += Clipboard.GetText();
            CursorEnd();
        }
    }

    public void OnKey(Keys key, OnKeyDownDeligate func) {
        onKeyDown[key] = func;
    }

    public void CursorEnd() {
        inputBox.Select(inputText.Length, 0);
    }

    public void ClearBox() {
        appendText = "";
        inputBox.Clear();
    }
    public void FocusBox() {
        inputBox.Focus();
    }

    private async void OnKeyDown(object sender, KeyEventArgs e) {
        switch (e.KeyCode) {
            case Keys.Up: {
                e.Handled = true;
                if (compForm.Visible) {
                    compForm.SelectIndexDec();
                    string line = compForm.SelectedLine();
                    if (0 < line.Length) {
                        replaceLastArg(line);
                    }
                }
                else {
                    try { inputText = history.Prev(); CursorEnd(); } catch {}
                }
                break;
            }
            case Keys.Down: {
                e.Handled = true;
                if (compForm.Visible) {
                    compForm.SelectIndexInc();
                    string line = compForm.SelectedLine();
                    if (0 < line.Length) {
                        replaceLastArg(line);
                    }
                }
                else {
                    try { inputText = history.Next(); CursorEnd(); } catch {}
                }
                break;
            }
            case Keys.Right: {
                if (inputBox.SelectionStart == inputText.Length) {
                    e.Handled = true;
                    try { inputText = history.LastMatch(inputText); CursorEnd(); } catch {}
                }
                break;
            }
            case Keys.Enter: {
                e.Handled = true;
                if (inputText.Length == 0) {
                    return;
                }

                compForm.Hide();
                history.Add(inputText);
                history.ResetPos();
                break;
            }
            case Keys.Tab: {
                e.Handled = true;
                e.SuppressKeyPress = true;
                if (inputText.Length == 0) {
                    return;
                }

                string[] res = await pwshProcess.TabCompleteAsync(inputText);

                if (res.Length == 1) {
                    compForm.Clear();
                    compForm.Hide();

                    replaceLastArg(res[0]);
                }
                else if (0 < res.Length) {
                    compForm.Clear();

                    foreach (string s in res) {
                        compForm.AppendLine(s);
                    }

                    Point pos = PointToScreen(Point.Empty);

                    compForm.Width = Width;
                    compForm.SetLB(pos.X, pos.Y, true);
                    compForm.ScrollTop();

                    inputBox.Focus();
                }
                break;
            }
        }

        if (onKeyDown.ContainsKey(e.KeyCode)) {
            e.Handled = true;
            onKeyDown[e.KeyCode](inputText);
        }
    }

    private void replaceLastArg(string rep) {
        var matches = Regex.Matches(inputText, @"('[^']*'|\S+)");

        if (0 < matches.Count) {
            if (inputText[inputText.Length - 1] == ' ') {
                inputText += rep;
            }
            else {
                string lastArg = matches[matches.Count - 1].Value;
                inputText = inputText.Substring(0, inputText.Length - lastArg.Length) + rep;
            }
            CursorEnd();
        }
    }
}


//-----------------------------------------------------------------------------
public class PopupForm : Form
{
    static private readonly int OUT_FORM_MAX_HEIGHT = Window.DpiScaledInt(500);
    private const int BORDER_WIDTH        = 1;

    static private Button    s_expandBtn;
    static private Form      s_expandForm;
    static private PopupForm s_activeForm;
    static private readonly List<PopupForm> instances = new List<PopupForm>();

    static PopupForm() {
        bool dark = (IniFile.ReadInt("ExpandButtonDark", 0) != 0);
        s_expandBtn = new Button {
            Text = "∧",
            FlatStyle = FlatStyle.Flat,

            BackColor = ColorTranslator.FromHtml(dark ? "#485269" : "#EFF6FC"),
            ForeColor = ColorTranslator.FromHtml(dark ? "#ccc" : "#777"),
        };

        s_expandBtn.FlatAppearance.BorderSize = 1;
        s_expandBtn.FlatAppearance.BorderColor = ColorTranslator.FromHtml(dark ? "#555" : "#ddd");
        s_expandBtn.Click += (s, e) => {
            if (s_expandForm != null) {
                if (s_expandBtn.Text.Equals("∧")) {
                    s_expandForm.Show();
                }
                else {
                    s_expandForm.Hide();
                }
            }
        };
    }

    public static void HideAll() {
        foreach (PopupForm form in instances ) {
            form.Hide();
        }
    }
    public static bool HasActive() {
        return (s_activeForm != null) ? true : false;
    }


    private RichTextBox textBox;
    private Font        font;
    private Pen         borderPen;

    private int    selectIdx = -1;
    private double lineHeight;

    public PopupForm(
        int    colNum      = 80,
        string foreColor   = "#CCCCCC",
        string backColor   = "#00001E",
        string borderColor = "#334"
    ): base()
    {
        Text            = "Result";
        BackColor       = ColorTranslator.FromHtml(backColor);
        TopMost         = true;
        StartPosition   = FormStartPosition.Manual;
        FormBorderStyle = FormBorderStyle.None;

        Padding   = new Padding(BORDER_WIDTH);
        borderPen = new Pen(ColorTranslator.FromHtml(borderColor), BORDER_WIDTH);
        Cursor    = Cursors.Arrow;

        font = new Font(IniFile.Read("Font", "Arial"), (int)(10 / Window.DPI_SCALING));
        Icon = new Icon(AppDomain.CurrentDomain.BaseDirectory + "\\pexe.ico");


        using (Graphics g = CreateGraphics()) {
            string testString = new string('W', colNum);
            SizeF size = g.MeasureString(testString, font);
            Width = (int)size.Width;
            lineHeight = font.Height;
        }

        textBox = new RichTextBox {
            Multiline   = true,
            Dock        = DockStyle.Fill,
            ReadOnly    = true,
            BackColor   = ColorTranslator.FromHtml(backColor),
            ForeColor   = ColorTranslator.FromHtml(foreColor),
            Font        = font,
            BorderStyle = BorderStyle.None,
            Cursor      = Cursors.Arrow,
        };

        Controls.Add(textBox);

        Paint += OnPaint;
        VisibleChanged += OnVisibleChanged;
        textBox.MouseUp += OnMouseUp;

        instances.Add(this);
    }

    public void Quit() {
        textBox.Dispose();
        borderPen.Dispose();
    }

    private void OnPaint(object sender, PaintEventArgs e) {
        e.Graphics.DrawRectangle(borderPen, 0, 0, this.ClientSize.Width - 1, this.ClientSize.Height - 1);
    }

    private void OnVisibleChanged(object sender, EventArgs e) {
        if (this.Visible) {
            if ((s_activeForm != null) && (s_activeForm != this)) {
                s_activeForm.Hide();
            }
            if (s_expandForm == this) {
                s_expandBtn.Text = "∨";
            }
            s_activeForm = this;
        }
        else {
            if ((s_activeForm != null) && (s_activeForm == this)) {
                s_activeForm = null;
            }
            if (s_expandForm == this) {
                s_expandBtn.Text = "∧";
            }
        }
    }

    private void OnMouseUp(object sender, MouseEventArgs e) {
        if (e.Button == MouseButtons.Left) {
            if (0 < textBox.SelectedText.Length) {
                Clipboard.SetText(textBox.SelectedText);
            }
        }
    }

    public Button ExpandBtn() {
        s_expandForm = this;
        return s_expandBtn;
    }

    public void ScrollTop() {
        textBox.SelectionStart = 0;
        textBox.ScrollToCaret();
    }

    public void AppendLine(string line) {
        textBox.AppendText(line + Environment.NewLine);
    }

    public void AppendColoredLine(string line) {
        line = line + Environment.NewLine;
        string[] parts = Regex.Split(line, @"(\x1b\[[0-9;]*m)");

        Color color     = textBox.ForeColor;
        Color backColor = textBox.BackColor;

        bool bold      = false;
        bool italic    = false;
        bool underline = false;
        bool negative  = false;
        int  start     = -1;

        foreach (string part in parts) {
            if (Regex.IsMatch(part, @"\x1b\[[0-9;]*m")) {
                Match match = Regex.Match(part, @"\x1b\[(?<codes>[0-9;]*)m");
                if (match.Success) {
                    string[] codes = match.Groups["codes"].Value.Split(';');
                    if (ColorHandle.IsAnsiColorCode(codes)) {
                        ColorHandle.ParseColorCode(codes, ref color, ref backColor);
                    }
                    else {
                        foreach (string code in codes) {
                            int colorCode;
                            if (int.TryParse(code, out colorCode)) {
                                switch (colorCode) {
                                    case 0:
                                        color     = textBox.ForeColor;
                                        backColor = textBox.BackColor;
                                        bold      = false;
                                        italic    = false;
                                        underline = false;
                                        negative  = false;
                                        break;
                                    case 1:
                                        bold = true;
                                        break;
                                    case 3:
                                        italic = true;
                                        break;
                                    case 4:
                                        underline = true;
                                        break;
                                    case 7:
                                        negative = true;
                                        break;
                                    default:
                                        Console.WriteLine("other: " + colorCode);
                                        break;
                                }
                                break;
                            }
                        }
                    }

                }
            }
            else {
                start = textBox.TextLength;
                textBox.AppendText(part);
                if (0 <= start) {
                    textBox.SelectionStart     = start;
                    textBox.SelectionLength    = textBox.TextLength - start;
                    textBox.SelectionColor     = color;
                    textBox.SelectionBackColor = backColor;

                    FontStyle style = FontStyle.Regular;
                    if (bold)      style |= FontStyle.Bold;
                    if (italic)    style |= FontStyle.Italic;
                    if (underline) style |= FontStyle.Underline;

                    textBox.SelectionFont = new Font(textBox.Font, style);

                    if (negative) {
                        textBox.SelectionBackColor = color;
                        textBox.SelectionColor = backColor;
                    }
                }

                start = -1;
            }
        }
    }

    public void Clear() {
        textBox.Clear();
        selectIdx = -1;
    }

    public void SetLB(int left, int bottom, bool show) {
        int lastCharIndex = textBox.Text.Length - 1;
        Point lastCharPosition = textBox.GetPositionFromCharIndex(lastCharIndex);

        Height = Math.Min((int)lineHeight * 2 + 6 + lastCharPosition.Y + textBox.Margin.Vertical, OUT_FORM_MAX_HEIGHT);

        Location = new Point(left, bottom - Height);

        if (show) {
            Visible = true;
        }
    }

    public void SelectIndexDec() {
        selectLine();

        if (0 < selectIdx) {
            selectIdx--;
        }
        selectLine();
    }
    public void SelectIndexInc() {
        selectLine();
        if (selectIdx + 1 < textBox.Lines.Length) {
            selectIdx++;
        }
        selectLine();
    }

    private void selectLine() {
        if (selectIdx < 0) {
            return;
        }
        if (textBox.Lines.Length <= selectIdx) {
            return;
        }

        int start = GetUnwrappedStartIndex(selectIdx);
        int length = textBox.Lines[selectIdx].Length;

        textBox.Select(start, length);
        textBox.SelectionColor = ColorHandle.ReverseColor(textBox.SelectionColor);
        textBox.SelectionBackColor = ColorHandle.ReverseColor(textBox.SelectionBackColor);
    }

    public string SelectedLine() {
        if (selectIdx < 0) {
            return "";
        }
        if (textBox.Lines.Length <= selectIdx) {
            return "";
        }
        return textBox.Lines[selectIdx];
    }

    private int GetUnwrappedStartIndex(int selectIdx) {
        int start = 0;
        for (int i1 = 0; i1 < selectIdx; i1++) {
            start += textBox.Lines[i1].Length + 1;
        }
        return start;
    }
}

//-----------------------------------------------------------------------------
public class PsForm : Form
{
    private CmdInputPanel inputPanel;
    private PopupForm     outputForm;
    private PwshProcess   pwshProcess;

    private Timer         focusCheckTimer;

    private Bitmap        bmp;
    private Graphics      graphics;

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    static private readonly int   WIDTH            = Window.DpiScaledInt(300);
    static private readonly int   HEIGHT           = Window.DpiScaledInt(40);
    static private readonly int   INPUT_BOX_HEIGHT = Window.DpiScaledInt(30);

    static private readonly int   TASK_BAR_OFFSET_IN  = (Window.TaskBar.HEIGHT - INPUT_BOX_HEIGHT) / 2;
    static private readonly int   TASK_BAR_OFFSET_OUT = HEIGHT - INPUT_BOX_HEIGHT - TASK_BAR_OFFSET_IN;

    static private readonly int    POS_LEFT     = IniFile.ReadInt("PosLeft", 10);
    static private readonly int    POS_TOP      = Window.TaskBar.TOP - TASK_BAR_OFFSET_OUT;

    public PsForm()
    {
        Text            = "pexe";
        Width           = WIDTH;
        Height          = HEIGHT;
        FormBorderStyle = FormBorderStyle.None;
        StartPosition   = FormStartPosition.Manual;
        Location        = new Point(POS_LEFT, POS_TOP);
        TopMost         = true;
        BackColor       = Color.LightBlue;
        TransparencyKey = Color.LightBlue;
        Icon            = new Icon(AppDomain.CurrentDomain.BaseDirectory + "\\pexe.ico");


        pwshProcess = new PwshProcess();

        outputForm = new PopupForm(IniFile.ReadInt("OutputColums", 80));

        inputPanel = new CmdInputPanel(INPUT_BOX_HEIGHT, pwshProcess) {
            Dock   = DockStyle.Bottom,
        };

        Button expandBtn = outputForm.ExpandBtn();
        expandBtn.Width  = INPUT_BOX_HEIGHT;
        expandBtn.Height = inputPanel.Height;
        expandBtn.Dock   = DockStyle.Right;

        var panel = new Panel { Dock = DockStyle.Bottom, Height = INPUT_BOX_HEIGHT };
        panel.Controls.Add(inputPanel);
        panel.Controls.Add(expandBtn);

        Controls.Add(panel);

        focusCheckTimer = new Timer();
        focusCheckTimer.Interval = 500;
        focusCheckTimer.Tick += FocusCheckTimer_Tick;
        focusCheckTimer.Start();

        bmp = new Bitmap(1,1);
        graphics = Graphics.FromImage(bmp);

        inputPanel.OnKey(Keys.Enter, OnEneter);
        inputPanel.OnKey(Keys.Escape, (cmd) => {
            if (PopupForm.HasActive()) {
                PopupForm.HideAll();
            }
            else {
                inputPanel.ClearBox();
            }
        });

        pwshProcess.Run();
    }

    private async void OnEneter(string cmd) {
        inputPanel.ClearBox();
        outputForm.Clear();

        if (string.Compare(cmd, "exit", true) == 0) {
            Close();
            return;
        }
        if (string.Compare(cmd, "term", true) == 0) {
            pwshProcess.SendTerm();
            return;
        }

        Task task = pwshProcess.ExecuteCmdAsync(cmd);

        bool first = true;

        outputForm.AppendLine("PS > " + cmd);
        if (outputForm.Visible) {
            outputForm.SetLB(this.Location.X, this.Location.Y + TASK_BAR_OFFSET_OUT, first);
            first = false;
        }

        int idx = 0;
        bool fin = false;
        do {
            Task tout = Task.Delay(TimeSpan.FromSeconds(1));
            fin = (await Task.WhenAny(task, tout) == task);

            if (idx < pwshProcess.Result.Count) {
                for (; idx < pwshProcess.Result.Count; idx++) {
                    outputForm.AppendColoredLine(pwshProcess.Result[idx]);
                }
                outputForm.SetLB(this.Location.X, this.Location.Y + TASK_BAR_OFFSET_OUT, first);
                first = false;
                outputForm.ScrollTop();
                inputPanel.FocusBox();
            }
        }
        while (!fin);

        return;
    }

    private int retryDelay = 0;
    private void FocusCheckTimer_Tick(object sender, EventArgs e) {
        if (0 < retryDelay) {
            retryDelay--;
        }
        else {
            try {
                IntPtr activeWindow = GetForegroundWindow();
                if (activeWindow != this.Handle && activeWindow != outputForm.Handle) {
                    PopupForm.HideAll();
                }

                graphics.CopyFromScreen(POS_LEFT + 20, POS_TOP + 20, 0, 0, bmp.Size);
                Color pixel = bmp.GetPixel(0, 0);

                if ((pixel.R < 255) && (pixel.G < 255) && (pixel.B < 255)) {
                    if ((Location.X != POS_LEFT) || (Location.Y != POS_TOP)) {
                        Location = new Point(POS_LEFT, POS_TOP);
                    }

                    TopMost = true;
                }
            }
            catch {
                retryDelay = 10;
            }
        }
    }

    protected override void OnFormClosing(FormClosingEventArgs e) {
        base.OnFormClosing(e);
        pwshProcess.Close();
        inputPanel.Quit();
        outputForm.Quit();
        bmp.Dispose();
    }

//-----------------------------------------------------------------------------
    [STAThread]
    static void Main() {
        IniFile.Open(AppDomain.CurrentDomain.BaseDirectory + "\\pexe.ini");

        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        Application.Run(new PsForm());
    }
}

