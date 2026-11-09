using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyTitle("המילון התורני 3.1")]
[assembly: AssemblyProduct("המילון התורני 3.1")]
[assembly: AssemblyDescription("מתקין מילון תורני (תנ״ך, משנה, בבלי, ירושלמי ועוד) ל-Microsoft Word")]
[assembly: AssemblyVersion("3.1.0.0")]
[assembly: AssemblyFileVersion("3.1.0.0")]
[assembly: AssemblyInformationalVersion("3.1.0.0")]

namespace OtzariaDict
{
    // ------------------------------------------------------------------------------------------ screen scale
    // Everything in the windows is measured in "design units" = pixels of a normal screen (100 %, 96 dpi) and multiplied by Dpi.Scale when the
    // window is made or drawn. The program tells Windows (SetProcessDPIAware, called in code before the first window; mcs ignores a manifest file)
    // that it handles a magnified screen by itself. Without that, Windows draws the finished window small and stretches the picture: that made it blurry at 125 % / 150 %.
    static class Dpi
    {
        public const int WindowHeight = 674;       // the main window, in design units (its width is 720)
        public static float Scale = 1f;
        public static bool Forced;                 // --dpi N : for the developer's screenshots on a machine without Windows

        [DllImport("user32.dll")] static extern bool SetProcessDPIAware();

        // percent 50..400 = use exactly this scale (screenshots); 0 = ask Windows. Must be called before any window exists.
        public static void Init(int percent)
        {
            if (percent >= 50 && percent <= 400) { Scale = percent / 100f; Forced = true; return; }
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) return;
            try
            {
                try { SetProcessDPIAware(); } catch (Exception) { }
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) Scale = Math.Max(1f, g.DpiX / 96f);
            }
            catch (Exception) { Scale = 1f; }
            try
            {
                // a screen too low for the window (for example 1920x1080 at 150 %): a little smaller, instead of cut off at the bottom
                int frame = SystemInformation.CaptionHeight + 2 * SystemInformation.FixedFrameBorderSize.Height + 8;
                float fit = (Screen.PrimaryScreen.WorkingArea.Height - frame) / (float)WindowHeight;
                if (fit < Scale) Scale = Math.Max(1f, fit);
            }
            catch (Exception) { }
        }

        public static int S(int v) { return (int)Math.Floor(v * Scale + 0.5f); }
        // a thin line: 1 pixel up to 175 %, 2 pixels from 200 %
        public static int Hair { get { return Math.Max(1, (int)Math.Floor(Scale + 0.01f)); } }
        // neighbours that touch in design units still touch on the screen (the edges are rounded, not the sizes)
        public static Rectangle R(int x, int y, int w, int h) { int a = S(x), b = S(y); return new Rectangle(a, b, S(x + w) - a, S(y + h) - b); }
        public static Size Sz(int w, int h) { return new Size(S(w), S(h)); }
    }

    // Look (3.1, "cleaner and more modern"): the light grey window of Windows 11 with white cards, rounded corners, the blue of Word as the
    // accent, light borders, semibold titles. The colours keep their old names, so every window of the program takes the new look.
    static class Theme
    {
        static Color C(string hex) { return ColorTranslator.FromHtml(hex); }
        public static readonly Color Primary = C("#185ABD");        // the blue of Word (the accent)
        public static readonly Color PrimaryHover = C("#1550A8");
        public static readonly Color PrimaryDark = C("#0F4290");    // pressed button, word count of the chosen card
        public static readonly Color OnPrimary = C("#FFFFFF");
        public static readonly Color OnPrimarySoft = C("#D6E4F7");  // small text on the blue (second line of the install button)
        public static readonly Color Soft = C("#EEF4FD");           // light blue: the chosen card, a quiet blue button
        public static readonly Color SoftHover = C("#E1ECFB");
        public static readonly Color SoftPill = C("#DCE8FA");       // the word count of the chosen card
        public static readonly Color Window = C("#F4F5F7");         // the main window: light grey, the cards on it are white
        public static readonly Color Surface = C("#FFFFFF");        // dialogs
        public static readonly Color Panel = C("#F4F5F7");          // the button strip at the bottom of a dialog
        public static readonly Color Card = C("#FFFFFF");           // cards, list rows, buttons
        public static readonly Color RowHover = C("#F8F9FB");
        public static readonly Color Line = C("#E6E8EB");           // light dividing lines
        public static readonly Color Border = C("#DFE2E6");         // the border of a card
        public static readonly Color BorderHover = C("#C3C8CF");
        public static readonly Color Stroke = C("#C9CDD3");         // the border of a button, of a list, of a text box
        public static readonly Color Mark = C("#8A9099");           // an empty round mark / switch
        public static readonly Color Track = C("#E3E6EA");          // empty progress line
        public static readonly Color Pill = C("#EEF0F3");           // a quiet label (word count, status)
        public static readonly Color BtnHover = C("#F4F5F7");
        public static readonly Color BtnPress = C("#EAECEF");
        public static readonly Color DisabledBack = C("#F1F2F4");
        public static readonly Color DisabledText = C("#A3A8AF");
        public static readonly Color DisabledLine = C("#D8DCE0");
        public static readonly Color Text = C("#1B1D21");
        public static readonly Color TextSoft = C("#4F5560");       // descriptions
        public static readonly Color TextQuiet = C("#6B717B");      // footer, version, small notes (4.9:1 on the window grey)
        public static readonly Color Ok = C("#0F7B3F");             // "installed"
        public static readonly Color OkSoft = C("#E6F4EC");
        public static readonly Color OnOk = C("#0B5A2E");
        public static readonly Color Warn = C("#B07A00");           // "can be updated"
        public static readonly Color WarnSoft = C("#FFF5D6");
        public static readonly Color OnWarn = C("#6B4A00");
        public static readonly Color Error = C("#C42B1C");
        public static readonly Color ErrorSoft = C("#FDECEA");
        public static readonly Color OnError = C("#8A1C12");

        // fonts are created once and reused. size = height in design units (pixels at 100 %); the font is made in whole real pixels.
        // "bold" is Segoe UI Semibold where Windows has it (the weight of Windows 11 titles): lighter and more modern than Bold.
        static readonly Dictionary<string, Font> fonts = new Dictionary<string, Font>();
        static string semibold;
        static bool semiChecked;
        static string Semibold()
        {
            if (semiChecked) return semibold;
            semiChecked = true;
            try
            {
                using (System.Drawing.Text.InstalledFontCollection all = new System.Drawing.Text.InstalledFontCollection())
                    foreach (FontFamily f in all.Families)
                        if (string.Equals(f.Name, "Segoe UI Semibold", StringComparison.OrdinalIgnoreCase)) { semibold = f.Name; break; }
            }
            catch (Exception) { }
            return semibold;
        }
        public static Font F(float size, bool bold)
        {
            string key = size.ToString(CultureInfo.InvariantCulture) + (bold ? "b" : "r") + "@" + Dpi.Scale.ToString(CultureInfo.InvariantCulture);
            Font f;
            lock (fonts)
            {
                if (!fonts.TryGetValue(key, out f))
                {
                    float px = (float)Math.Floor(size * Dpi.Scale + 0.5);
                    string semi = bold ? Semibold() : null;
                    f = semi != null ? new Font(semi, px, FontStyle.Regular, GraphicsUnit.Pixel)
                                     : new Font("Segoe UI", px, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel);
                    fonts[key] = f;
                }
            }
            return f;
        }

        public static Color Blend(Color fg, double a, Color bg)
        {
            return Color.FromArgb((int)(fg.R * a + bg.R * (1 - a)), (int)(fg.G * a + bg.G * (1 - a)), (int)(fg.B * a + bg.B * (1 - a)));
        }
    }

    enum Align { Right, Center, Left }

    static class Gfx
    {
        // rounded rectangle; rad = corner radius in pixels (0 = square corners)
        public static GraphicsPath Round(RectangleF r, float rad)
        {
            GraphicsPath p = new GraphicsPath();
            float d = Math.Min(rad * 2f, Math.Min(r.Width, r.Height));
            if (d < 1f) { p.AddRectangle(r); return p; }
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }
        // The graphics library of Linux (libgdiplus, used only for the developer's screenshots) fills a path one pixel too wide and too high.
        // Windows fills exactly (as long as PixelOffsetMode.Half is on), so this correction is only made when the program runs under Mono.
        static readonly bool onMono = Type.GetType("Mono.Runtime") != null;
        static Rectangle Exact(Rectangle r) { return onMono ? new Rectangle(r.X, r.Y, r.Width - 1, r.Height - 1) : r; }

        public static void Fill(Graphics g, Rectangle r, int rad, Color c)
        { using (GraphicsPath p = Round(Exact(r), rad)) using (SolidBrush b = new SolidBrush(c)) g.FillPath(b, p); }

        // A rounded rectangle with a border of w whole pixels: first the whole shape in the border colour, then the inside (smaller by w on every
        // side) on top of it. Two fills, not a pen: a filled edge on a whole-pixel position is sharp, while a pen line that is centred on the
        // edge between two pixels is smeared over both of them (that - and the soft shadows - is what made the old borders look grainy).
        public static void Box(Graphics g, Rectangle r, int rad, Color back, Color border, int w)
        {
            Fill(g, r, rad, border);
            Rectangle inner = Rectangle.Inflate(r, -w, -w);
            if (inner.Width > 0 && inner.Height > 0) Fill(g, inner, Math.Max(0, rad - w), back);
        }
        public static void Dot(Graphics g, Rectangle r, Color c)
        { using (SolidBrush b = new SolidBrush(c)) g.FillEllipse(b, Exact(r)); }
        // a round mark with a ring of w pixels (the inside in the colour of what lies behind it)
        public static void RingDot(Graphics g, Rectangle r, Color back, Color ring, int w)
        {
            Dot(g, r, ring);
            Dot(g, Rectangle.Inflate(r, -w, -w), back);
        }
        // a plain filled rectangle (always whole pixels, so its edges are sharp)
        public static void Rect(Graphics g, Rectangle r, Color c)
        { using (SolidBrush b = new SolidBrush(c)) g.FillRectangle(b, r); }
        // a thin straight line (a filled strip, always whole pixels)
        public static void HLine(Graphics g, int x, int y, int w, Color c)
        { Rect(g, new Rectangle(x, y, w, Dpi.Hair), c); }
        // an outline of w whole pixels (four strips)
        public static void Frame(Graphics g, Rectangle r, Color c, int w)
        {
            Rect(g, new Rectangle(r.X, r.Y, r.Width, w), c);
            Rect(g, new Rectangle(r.X, r.Bottom - w, r.Width, w), c);
            Rect(g, new Rectangle(r.X, r.Y + w, w, r.Height - 2 * w), c);
            Rect(g, new Rectangle(r.Right - w, r.Y + w, w, r.Height - 2 * w), c);
        }

        // Hebrew text: right-to-left reading order, aligned to the right unless centred
        public static void Text(Graphics g, string s, Font f, Color c, Rectangle r, bool center, bool vcenter)
        { TextAt(g, s, f, c, r, center ? Align.Center : Align.Right, vcenter, true); }

        // wrap = false: one line, never wrapped. Left = the visual left edge (the text itself still reads right to left)
        public static void TextAt(Graphics g, string s, Font f, Color c, Rectangle r, Align a, bool vcenter, bool wrap)
        {
            TextFormatFlags fl = TextFormatFlags.RightToLeft | TextFormatFlags.NoPrefix;
            // one line that is too long ends with "…" (instead of being cut off - with Windows' own fonts a text may be a little wider)
            fl |= wrap ? (TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl) : (TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis);
            fl |= a == Align.Center ? TextFormatFlags.HorizontalCenter : a == Align.Right ? TextFormatFlags.Right : TextFormatFlags.Left;
            if (vcenter) fl |= TextFormatFlags.VerticalCenter;
            TextRenderer.DrawText(g, s, f, r, c, fl);
        }

        // one line on the right, cut with "…" at its end when it does not fit
        public static void Line(Graphics g, string s, Font f, Color c, Rectangle r)
        {
            TextRenderer.DrawText(g, s, f, r, c, TextFormatFlags.RightToLeft | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine
                                                 | TextFormatFlags.Right | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
        }

        public static Color MarkColor(MsgKind k)
        { return k == MsgKind.Ok ? Theme.Ok : k == MsgKind.Warn ? Theme.Warn : k == MsgKind.Error ? Theme.Error : Theme.Primary; }

        // a round mark: a tick (done), "!" (attention, or it did not work), "?" (a question) or "i" (good to know).
        // glyph = the size of the "!" / "?" / "i" in design units
        public static void Mark(Graphics g, Rectangle m, MsgKind k, float glyph)
        {
            Dot(g, m, MarkColor(k));
            if (k == MsgKind.Ok)
            {
                float s = m.Width / 22f;
                using (Pen tick = new Pen(Theme.OnPrimary, 2.4f * s))
                {
                    tick.StartCap = LineCap.Round; tick.EndCap = LineCap.Round; tick.LineJoin = LineJoin.Round;
                    g.DrawLines(tick, new PointF[] { new PointF(m.X + 6.0f * s, m.Y + 11.4f * s), new PointF(m.X + 9.5f * s, m.Y + 14.8f * s), new PointF(m.X + 16.0f * s, m.Y + 7.6f * s) });
                }
            }
            else Text(g, k == MsgKind.Question ? "?" : k == MsgKind.Info ? "i" : "!", Theme.F(glyph, true), Theme.OnPrimary, m, true, true);
        }
    }

    class BaseControl : Control
    {
        public BaseControl()
        {
            TabStop = false;   // controls that can be used (buttons, rows, lists, links) turn it on
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
        }
    }

    // the top of the main window, on the grey of the window: the logo (the blue square with alef), the name and one line under it.
    // The state of the computer ("installed: ...") is a pill on the other side of it (a child control, placed by MainForm).
    class Header : BaseControl
    {
        public Header() { BackColor = Theme.Window; }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            g.Clear(BackColor);
            Rectangle logo = new Rectangle(Width - Dpi.S(32) - Dpi.S(48), Dpi.S(24), Dpi.S(48), Dpi.S(48));
            Gfx.Fill(g, logo, Dpi.S(12), Theme.Primary);
            Gfx.Text(g, "א", Theme.F(28, true), Theme.OnPrimary, logo, true, true);
            int right = logo.X - Dpi.S(16), left = Dpi.S(340);
            Gfx.TextAt(g, MainForm.Title, Theme.F(24, true), Theme.Text, new Rectangle(left, Dpi.S(22), right - left, Dpi.S(32)), Align.Right, true, false);
            Gfx.TextAt(g, "כדי שוורד יפסיק לסמן מילים תורניות כשגיאות", Theme.F(14, false), Theme.TextSoft, new Rectangle(left, Dpi.S(54), right - left, Dpi.S(20)), Align.Right, true, false);
        }
    }

    // the light strip at the bottom of a dialog: a thin line on top, the buttons on it
    class LowerPanel : BaseControl
    {
        public LowerPanel() { BackColor = Theme.Panel; }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(BackColor);
            Gfx.HLine(e.Graphics, 0, 0, Width, Theme.Line);
        }
    }

    // a white card with rounded corners and a light border, on the grey of the main window (its controls sit inside it)
    class CardPanel : BaseControl
    {
        public CardPanel() { BackColor = Theme.Card; }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Window);
            Gfx.Box(g, ClientRectangle, Dpi.S(8), Theme.Card, Theme.Border, Dpi.Hair);
        }
    }

    class TextBlock : BaseControl
    {
        public Color Color = Theme.Text;
        public bool Center, VCenter;
        public bool Wrap = true;
        string text = "";
        public override string Text { get { return text; } set { text = value ?? ""; Invalidate(); } }
        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(Parent != null ? Parent.BackColor : Theme.Surface);
            Gfx.TextAt(e.Graphics, text, Font, Color, new Rectangle(0, 0, Width, Height), Center ? Align.Center : Align.Right, VCenter, Wrap);
        }
    }

    enum LabelKind { None, Ok, Warn }

    // "installed: ..." - a pill: a coloured dot and the text, on a light background of the same colour (green = up to date,
    // yellow = can be updated, grey = not installed). It hugs the left edge of its control.
    class StatusLabel : BaseControl
    {
        public LabelKind Kind = LabelKind.None;
        string text = "";
        public override string Text { get { return text; } set { text = value ?? ""; Invalidate(); } }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Surface);
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            Font f = Theme.F(13, true);
            Color back = Kind == LabelKind.Ok ? Theme.OkSoft : Kind == LabelKind.Warn ? Theme.WarnSoft : Theme.Pill;
            Color fore = Kind == LabelKind.Ok ? Theme.OnOk : Kind == LabelKind.Warn ? Theme.OnWarn : Theme.TextSoft;
            Color dotColor = Kind == LabelKind.Ok ? Theme.Ok : Kind == LabelKind.Warn ? Theme.Warn : Theme.Mark;
            int dot = Dpi.S(8), gap = Dpi.S(8), padL = Dpi.S(12), padR = Dpi.S(12);
            int tw;
            try { tw = TextRenderer.MeasureText(g, text, f, new Size(Width, Height), TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width; }
            catch (Exception) { tw = Width - dot - gap - padL - padR; }
            int w = Math.Min(Width, padL + tw + gap + dot + padR);
            Rectangle pill = new Rectangle(0, 0, w, Height);
            Gfx.Fill(g, pill, Height / 2, back);
            // right to left: the dot first (on the right), then the text
            Rectangle d = new Rectangle(pill.Right - padR - dot, (Height - dot) / 2, dot, dot);
            Gfx.Dot(g, d, dotColor);
            Gfx.TextAt(g, text, f, fore, new Rectangle(padL, 0, d.X - gap - padL + Dpi.S(2), Height), Align.Right, true, false);
        }
    }

    // the result of the last action, in a light coloured bar with a mark (a tick when it worked, "!" when it did not): its title in bold and
    // under it the first sentence (the whole text is in the message window that opened with it). Title and text empty = nothing shown.
    class InfoBar : BaseControl
    {
        public bool Ok = true;
        string title = "", text = "";
        public string Title { get { return title; } set { title = value ?? ""; Invalidate(); } }
        public override string Text { get { return text; } set { text = value ?? ""; Invalidate(); } }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Window);
            if (title.Length == 0 && text.Length == 0) return;
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            Gfx.Fill(g, ClientRectangle, Dpi.S(6), Ok ? Theme.OkSoft : Theme.ErrorSoft);
            int m = Dpi.S(18);
            Rectangle mark = new Rectangle(Width - Dpi.S(14) - m, Dpi.S(title.Length > 0 ? 8 : 14), m, m);
            Gfx.Mark(g, mark, Ok ? MsgKind.Ok : MsgKind.Error, 13);
            Color fore = Ok ? Theme.OnOk : Theme.OnError;
            int left = Dpi.S(14), w = mark.X - Dpi.S(10) - left;
            if (title.Length == 0)
                Gfx.TextAt(g, text, Theme.F(13.5f, false), fore, new Rectangle(left, Dpi.S(8), w, Height - Dpi.S(14)), Align.Right, false, true);
            else
            {
                // two lines, each cut with "…" if it is longer than the bar (the mark is centred on the first line)
                int top = mark.Y + m / 2 - Dpi.S(10);
                Gfx.Line(g, title, Theme.F(13.5f, true), fore, new Rectangle(left, top, w, Dpi.S(20)));
                Gfx.Line(g, text, Theme.F(13.5f, false), fore, new Rectangle(left, top + Dpi.S(20), w, Dpi.S(20)));
            }
        }
    }

    class ProgressLine : BaseControl
    {
        int val;
        public int Value { get { return val; } set { val = Math.Max(0, Math.Min(100, value)); Invalidate(); } }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Surface);
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            Gfx.Fill(g, ClientRectangle, Height / 2, Theme.Track);
            int w = (int)(Width * (val / 100.0));
            if (w > 0) Gfx.Fill(g, new Rectangle(Width - w, 0, w, Height), Height / 2, Theme.Primary);   // fills from the right (RTL)
        }
    }

    enum BtnStyle { Filled, Outlined, Text, Tonal }

    // Filled = the blue button · Outlined = white with a light border · Tonal = light blue, no border (a quiet blue button) · Text = no frame.
    // Rounded corners (6), semibold text.
    class FlatBtn : BaseControl
    {
        public BtnStyle Style = BtnStyle.Filled;
        public bool Light;   // quiet secondary button: regular weight text
        public bool Small;   // small text (the buttons inside a row of the word list)
        bool hover, down, viaMouse;
        string subText = "";
        // a small second line under the text of the button (the install button uses it to say exactly what will be installed)
        public string SubText { get { return subText; } set { subText = value ?? ""; Invalidate(); } }
        public FlatBtn() { Cursor = Cursors.Hand; TabStop = true; AccessibleRole = AccessibleRole.PushButton; }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { down = true; viaMouse = true; Focus(); Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Cursor = Enabled ? Cursors.Hand : Cursors.Default; Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { viaMouse = false; Invalidate(); base.OnLostFocus(e); }   // the focus frame is for keyboard use only
        protected override bool IsInputKey(Keys k) { return k == Keys.Enter || k == Keys.Space || base.IsInputKey(k); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            viaMouse = false;
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; }
            base.OnKeyDown(e);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Color page = Parent != null ? Parent.BackColor : Theme.Surface;
            g.Clear(page);
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;   // Half: whole numbers are pixel EDGES, so shapes and lines land exactly on pixels
            Rectangle r = ClientRectangle;
            int rad = Dpi.S(Small ? 4 : 6);
            Color bg = Color.Transparent, fg = Theme.Primary, border = Color.Transparent;
            if (!Enabled)
            {
                fg = Theme.DisabledText;
                if (Style == BtnStyle.Filled || Style == BtnStyle.Tonal) bg = Theme.DisabledBack;
                else if (Style == BtnStyle.Outlined) { bg = Theme.Card; border = Theme.DisabledLine; }
            }
            else if (Style == BtnStyle.Filled)
            {
                fg = Theme.OnPrimary;
                bg = down ? Theme.PrimaryDark : hover ? Theme.PrimaryHover : Theme.Primary;
            }
            else if (Style == BtnStyle.Tonal)
            {
                fg = Theme.Primary;
                bg = down ? Theme.SoftPill : hover ? Theme.SoftHover : Theme.Soft;
            }
            else if (Style == BtnStyle.Outlined)
            {
                fg = Theme.Text;
                bg = down ? Theme.BtnPress : hover ? Theme.BtnHover : Theme.Card;
                border = hover ? Theme.BorderHover : Theme.Stroke;
            }
            else
            {
                bg = down ? Theme.SoftHover : hover ? Theme.Soft : Color.Transparent;
            }
            Color inside = bg == Color.Transparent ? page : bg;   // the colour inside the button (what a focus frame is drawn on)
            if (border != Color.Transparent) Gfx.Box(g, r, rad, inside, border, Dpi.Hair);
            else if (bg != Color.Transparent) Gfx.Fill(g, r, rad, bg);
            if (Focused && Enabled && !viaMouse && ShowFocusCues)   // ShowFocusCues: like every Windows button, the frame appears only once the keyboard (Tab) is used
                Gfx.Frame(g, Rectangle.Inflate(r, -Dpi.S(3), -Dpi.S(3)), Style == BtnStyle.Filled ? Theme.OnPrimary : Theme.Primary, Dpi.S(2));

            Font f = Theme.F(Small ? 13 : (Style == BtnStyle.Filled ? 15 : 14), !Light);
            if (subText.Length == 0)
                Gfx.TextAt(g, Text, f, fg, new Rectangle(Dpi.S(8), 0, Width - Dpi.S(16), Height), Align.Center, true, false);
            else
            {
                Color sub = Enabled ? (Style == BtnStyle.Filled ? Theme.OnPrimarySoft : Theme.TextSoft) : fg;
                int top = (Height - Dpi.S(40)) / 2;
                Gfx.TextAt(g, Text, f, fg, new Rectangle(Dpi.S(8), top, Width - Dpi.S(16), Dpi.S(22)), Align.Center, true, false);
                Gfx.TextAt(g, subText, Theme.F(12.5f, false), sub, new Rectangle(Dpi.S(8), top + Dpi.S(22), Width - Dpi.S(16), Dpi.S(18)), Align.Center, true, false);
            }
        }
    }

    enum Glyph { Pencil, Folder, Double, Trash }

    // a quiet button of the tools row: a small drawn icon and the text next to it, no frame; light grey under the mouse
    class IconBtn : BaseControl
    {
        public Glyph Icon;
        bool hover, down, viaMouse;
        public IconBtn() { Cursor = Cursors.Hand; TabStop = true; AccessibleRole = AccessibleRole.PushButton; }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; down = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { down = true; viaMouse = true; Focus(); Invalidate(); base.OnMouseDown(e); }
        protected override void OnMouseUp(MouseEventArgs e) { down = false; Invalidate(); base.OnMouseUp(e); }
        protected override void OnEnabledChanged(EventArgs e) { Cursor = Enabled ? Cursors.Hand : Cursors.Default; Invalidate(); base.OnEnabledChanged(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { viaMouse = false; Invalidate(); base.OnLostFocus(e); }
        protected override bool IsInputKey(Keys k) { return k == Keys.Enter || k == Keys.Space || base.IsInputKey(k); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            viaMouse = false;
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; }
            base.OnKeyDown(e);
        }
        public override string Text { get { return base.Text; } set { base.Text = value; Invalidate(); } }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Color page = Parent != null ? Parent.BackColor : Theme.Card;
            g.Clear(page);
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            if (Enabled && (hover || down)) Gfx.Fill(g, ClientRectangle, Dpi.S(6), down ? Theme.BtnPress : Theme.BtnHover);
            if (Focused && Enabled && !viaMouse && ShowFocusCues) Gfx.Frame(g, Rectangle.Inflate(ClientRectangle, -Dpi.S(2), -Dpi.S(2)), Theme.Primary, Dpi.S(2));
            Font f = Theme.F(14, false);
            Color fg = Enabled ? Theme.Text : Theme.DisabledText, ic = Enabled ? Theme.Primary : Theme.DisabledText;
            int icon = Dpi.S(18), gap = Dpi.S(8);
            int tw;
            try { tw = TextRenderer.MeasureText(g, Text, f, new Size(Width, Height), TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width; }
            catch (Exception) { tw = Width / 2; }
            tw = Math.Min(tw, Width - icon - gap - Dpi.S(12));
            int total = icon + gap + tw, x0 = (Width - total) / 2;
            // right to left: the icon on the right, the text on its left
            Rectangle ir = new Rectangle(x0 + tw + gap, (Height - icon) / 2, icon, icon);
            DrawGlyph(g, Icon, ir, ic);
            Gfx.TextAt(g, Text, f, fg, new Rectangle(x0 - Dpi.S(2), 0, tw + Dpi.S(4), Height), Align.Center, true, false);
        }

        // simple line icons, drawn (no icon font is needed): a pencil, a folder, two lines of the same word, a bin
        public static void DrawGlyph(Graphics g, Glyph kind, Rectangle r, Color c)
        {
            float k = r.Width / 18f, x = r.X, y = r.Y;
            using (Pen p = new Pen(c, 1.6f * k))
            {
                p.StartCap = LineCap.Round; p.EndCap = LineCap.Round; p.LineJoin = LineJoin.Round;
                if (kind == Glyph.Pencil)
                {
                    g.DrawPolygon(p, new PointF[] { new PointF(x + 12.5f * k, y + 2.5f * k), new PointF(x + 15.5f * k, y + 5.5f * k), new PointF(x + 6f * k, y + 15f * k),
                                                    new PointF(x + 2.5f * k, y + 15.5f * k), new PointF(x + 3f * k, y + 12f * k) });
                    g.DrawLine(p, x + 10.5f * k, y + 4.5f * k, x + 13.5f * k, y + 7.5f * k);
                }
                else if (kind == Glyph.Folder)
                {
                    g.DrawPolygon(p, new PointF[] { new PointF(x + 2f * k, y + 4.5f * k), new PointF(x + 7f * k, y + 4.5f * k), new PointF(x + 8.5f * k, y + 6.5f * k),
                                                    new PointF(x + 16f * k, y + 6.5f * k), new PointF(x + 16f * k, y + 14.5f * k), new PointF(x + 2f * k, y + 14.5f * k) });
                    g.DrawLine(p, x + 2f * k, y + 8.5f * k, x + 16f * k, y + 8.5f * k);
                }
                else if (kind == Glyph.Double)
                {
                    g.DrawLine(p, x + 2.5f * k, y + 6f * k, x + 8f * k, y + 6f * k);
                    g.DrawLine(p, x + 10f * k, y + 6f * k, x + 15.5f * k, y + 6f * k);
                    g.DrawLine(p, x + 2.5f * k, y + 12f * k, x + 8f * k, y + 12f * k);
                    g.DrawLine(p, x + 10f * k, y + 12f * k, x + 15.5f * k, y + 12f * k);
                }
                else
                {
                    g.DrawLine(p, x + 3f * k, y + 5f * k, x + 15f * k, y + 5f * k);
                    g.DrawLine(p, x + 7f * k, y + 5f * k, x + 7.5f * k, y + 3f * k);
                    g.DrawLine(p, x + 7.5f * k, y + 3f * k, x + 10.5f * k, y + 3f * k);
                    g.DrawLine(p, x + 10.5f * k, y + 3f * k, x + 11f * k, y + 5f * k);
                    g.DrawPolygon(p, new PointF[] { new PointF(x + 4.5f * k, y + 5f * k), new PointF(x + 13.5f * k, y + 5f * k), new PointF(x + 12.5f * k, y + 15.5f * k), new PointF(x + 5.5f * k, y + 15.5f * k) });
                    g.DrawLine(p, x + 7.8f * k, y + 8f * k, x + 8f * k, y + 13f * k);
                    g.DrawLine(p, x + 10.2f * k, y + 8f * k, x + 10f * k, y + 13f * k);
                }
            }
        }
    }

    enum RowPos { First, Middle, Last }

    // One card of the choice (three of them, one under the other, with a small gap): a round mark (pick one) or a switch (CheckStyle, on/off).
    // A white card with rounded corners and a light border; the chosen one is light blue with a blue border. The word count is a pill on the left.
    class OptionRow : BaseControl
    {
        public string Title = "", Desc = "", Count = "";
        public bool CheckStyle;
        public RowPos Pos = RowPos.Middle;   // (kept for the callers; every card draws its whole frame now)
        bool selected, hover, viaMouse;
        public event EventHandler Chosen;
        public bool Selected { get { return selected; } set { selected = value; Invalidate(); } }
        public OptionRow() { Cursor = Cursors.Hand; TabStop = true; }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e) { viaMouse = true; base.OnMouseDown(e); }
        protected override void OnClick(EventArgs e) { Focus(); if (Chosen != null) Chosen(this, EventArgs.Empty); base.OnClick(e); }
        protected override bool IsInputKey(Keys k) { return k == Keys.Enter || k == Keys.Space || base.IsInputKey(k); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            viaMouse = false;
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; }
            base.OnKeyDown(e);
        }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { viaMouse = false; Invalidate(); base.OnLostFocus(e); }   // the focus frame is for keyboard use only
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;   // Half: whole numbers are pixel EDGES, so shapes and lines land exactly on pixels
            g.Clear(Parent != null ? Parent.BackColor : Theme.Window);
            Rectangle r = ClientRectangle;
            int rad = Dpi.S(8);
            Color back = selected ? Theme.Soft : (hover ? Theme.RowHover : Theme.Card);
            if (selected) Gfx.Box(g, r, rad, back, Theme.Primary, Math.Max(2, Dpi.S(2)));
            else Gfx.Box(g, r, rad, back, hover ? Theme.BorderHover : Theme.Border, Dpi.Hair);
            if (Focused && !viaMouse && ShowFocusCues) Gfx.Frame(g, Rectangle.Inflate(r, -Dpi.S(5), -Dpi.S(5)), Theme.Primary, Dpi.S(2));

            // the mark on the right: a round mark, or a switch
            int markRight = r.Right - Dpi.S(20);
            int markLeft;
            if (CheckStyle)
            {
                Rectangle sw = new Rectangle(markRight - Dpi.S(40), r.Y + (r.Height - Dpi.S(20)) / 2, Dpi.S(40), Dpi.S(20));
                int knob = Dpi.S(12), pad = Dpi.S(4);
                if (selected)
                {
                    Gfx.Fill(g, sw, sw.Height / 2, Theme.Primary);
                    Gfx.Dot(g, new Rectangle(sw.X + pad, sw.Y + pad, knob, knob), Theme.OnPrimary);   // on: the knob on the left (right to left)
                }
                else
                {
                    Gfx.Box(g, sw, sw.Height / 2, back, Theme.Mark, Dpi.Hair);
                    Gfx.Dot(g, new Rectangle(sw.Right - pad - knob, sw.Y + pad, knob, knob), Theme.Mark);   // off: on the right
                }
                markLeft = sw.X;
            }
            else
            {
                int m = Dpi.S(20);
                Rectangle dot = new Rectangle(markRight - m, r.Y + (r.Height - m) / 2, m, m);
                if (selected)
                {
                    Gfx.Dot(g, dot, Theme.Primary);
                    Gfx.Dot(g, Rectangle.Inflate(dot, -Dpi.S(6), -Dpi.S(6)), Theme.OnPrimary);
                }
                else Gfx.RingDot(g, dot, back, Theme.Mark, Math.Max(1, (int)Math.Round(1.5f * Dpi.Scale)));
                markLeft = dot.X;
            }

            // the number of words: a pill on the left
            Font cf = Theme.F(13, true);
            int pillW = 0;
            if (Count.Length > 0)
            {
                int tw;
                try { tw = TextRenderer.MeasureText(g, Count, cf, new Size(Width, Height), TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine).Width; }
                catch (Exception) { tw = Dpi.S(90); }
                pillW = tw + Dpi.S(24);
                Rectangle pill = new Rectangle(r.X + Dpi.S(20), r.Y + (r.Height - Dpi.S(28)) / 2, pillW, Dpi.S(28));
                Gfx.Fill(g, pill, pill.Height / 2, selected ? Theme.SoftPill : Theme.Pill);
                Gfx.TextAt(g, Count, cf, selected ? Theme.PrimaryDark : Theme.TextSoft, pill, Align.Center, true, false);
            }

            int textRight = markLeft - Dpi.S(16);
            int textLeft = r.X + Dpi.S(20) + Math.Max(pillW, Dpi.S(110)) + Dpi.S(20);
            int top = r.Y + (r.Height - Dpi.S(62)) / 2;   // title (22) + two lines of description (2 x 20) = 62, in the middle of the card
            Gfx.TextAt(g, Title, Theme.F(16, true), Theme.Text, new Rectangle(textLeft, top, textRight - textLeft, Dpi.S(22)), Align.Right, true, false);
            Gfx.TextAt(g, Desc, Theme.F(13.5f, false), Theme.TextSoft, new Rectangle(textLeft, top + Dpi.S(23), textRight - textLeft, Dpi.S(40)), Align.Right, false, true);
        }
    }

    // ------------------------------------------------------------------------------------------ "Word is open" window
    // Shown before an installation or a removal while Word is running. Word reads its dictionaries only when it starts, so a change made
    // while it is open is not in effect until it is closed and opened again: the person is asked to close it first.
    // Choice: Cancel (also the X and Esc) = do nothing; Again = look again (Word was closed meanwhile); GoOn = go on although Word is open.
    class WordOpenDialog : Form
    {
        public const int Cancel = 0, Again = 1, GoOn = 2;
        public int Choice = Cancel;

        public WordOpenDialog(bool installing)
        {
            Text = MainForm.Title;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Theme.Card;
            Font = Theme.F(14, false);
            ClientSize = Dpi.Sz(560, 360);
            Icon = MainForm.AppIcon();

            string job = installing ? "ההתקנה" : "ההסרה";
            string goOn = installing
                ? "המילון יותקן, אבל וורד יכיר אותו רק אחרי שתסגרו אותו ותפתחו אותו מחדש."
                : "המילון יוסר, אבל וורד יפסיק להשתמש בו רק אחרי שתסגרו אותו ותפתחו אותו מחדש.";
            TextBlock head = new TextBlock { Bounds = Dpi.R(28, 20, 504, 34), Font = Theme.F(22, true), Color = Theme.Text, VCenter = true, Wrap = false, Text = "וורד פתוח עכשיו" };
            TextBlock body = new TextBlock
            {
                Bounds = Dpi.R(28, 64, 504, 214), Font = Theme.F(14, false), Color = Theme.Text,
                Text = "וורד טוען את המילונים רק כשהוא נפתח, ולכן כדאי לסגור אותו לפני " + job + ".\r\n\r\n"
                     + "שמרו את המסמכים, סגרו את כל חלונות וורד ולחצו על \"בדוק שוב\".\r\n\r\n"
                     + "אין חלון וורד פתוח? ייתכן שוורד רץ ברקע, למשל בגלל תוכנה אחרת. חכו רגע ולחצו שוב, ואם זה לא עוזר, הפעילו את המחשב מחדש.\r\n\r\n"
                     + "אפשר גם ללחוץ על \"המשך בכל זאת\": " + goOn
            };
            LowerPanel strip = new LowerPanel { Bounds = Dpi.R(0, 280, 560, 80) };
            FlatBtn again = new FlatBtn { Bounds = Dpi.R(372, 18, 160, 44), Text = "בדוק שוב", Style = BtnStyle.Filled };
            FlatBtn goOnBtn = new FlatBtn { Bounds = Dpi.R(180, 18, 180, 44), Text = "המשך בכל זאת", Style = BtnStyle.Outlined };
            FlatBtn cancel = new FlatBtn { Bounds = Dpi.R(58, 18, 110, 44), Text = "ביטול", Style = BtnStyle.Outlined };
            again.Click += delegate { Choice = Again; Close(); };
            goOnBtn.Click += delegate { Choice = GoOn; Close(); };
            cancel.Click += delegate { Choice = Cancel; Close(); };
            strip.Controls.AddRange(new Control[] { again, goOnBtn, cancel });   // the tab order follows this order (right to left)
            Controls.AddRange(new Control[] { head, body, strip });
            ActiveControl = again;
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData == Keys.Escape) { Choice = Cancel; Close(); return true; }
            return base.ProcessDialogKey(keyData);
        }
    }

    // ------------------------------------------------------------------------------------------ a message or a question
    enum MsgKind { Ok, Info, Warn, Error, Question }

    // the title of a message window: a round mark at its start (on the right), then the text (it may take two lines)
    class MarkTitle : BaseControl
    {
        public MsgKind Kind = MsgKind.Info;
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Card);
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            int mark = Dpi.S(24), gap = Dpi.S(12), line = Dpi.S(32);
            Rectangle m = new Rectangle(Width - mark, (line - mark) / 2, mark, mark);
            Gfx.Mark(g, m, Kind, 15);
            Gfx.TextAt(g, Text, Font, Theme.Text, new Rectangle(0, 0, m.X - gap, Height), Align.Right, Height <= line, true);
        }
    }

    // 3.1: a message or a question in the look of the program, instead of the plain message box of Windows - a title with a round mark,
    // the text, and the buttons (in Hebrew on every Windows) on the light strip at the bottom. The window is as high as its text;
    // a very long text scrolls. Escape and the close box give the safe answer.
    class MsgDialog : Form
    {
        DialogResult answer;
        readonly DialogResult escape;
        static readonly TextFormatFlags Wrap = TextFormatFlags.RightToLeft | TextFormatFlags.NoPrefix | TextFormatFlags.WordBreak | TextFormatFlags.TextBoxControl;

        // labels / answers: the buttons from right to left - the first one is the blue one; focus = the button that has the focus at the start
        // (Enter presses it); escapeAnswer = the answer of Escape and of the close box
        MsgDialog(string text, string title, MsgKind kind, string[] labels, DialogResult[] answers, int focus, DialogResult escapeAnswer)
        {
            answer = escape = escapeAnswer;
            Text = MainForm.Title;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            if (Dpi.Forced) { StartPosition = FormStartPosition.Manual; Location = new Point(0, 0); }   // the developer's screenshots
            else StartPosition = FormStartPosition.CenterParent;
            BackColor = Theme.Card;
            Font = Theme.F(14, false);
            Icon = MainForm.AppIcon();
            // a screen reader announces a dialog, its title and its text (the text is painted, so it is given here too)
            AccessibleRole = AccessibleRole.Dialog;
            AccessibleName = title;
            AccessibleDescription = text;
            copyText = title + "\r\n\r\n" + text;

            int width = Dpi.S(520), pad = Dpi.S(28), inner = width - 2 * pad;
            Font titleFont = Theme.F(20, true), bodyFont = Theme.F(14, false);
            int titleW = inner - Dpi.S(24) - Dpi.S(12);
            int titleH = Math.Max(Dpi.S(32), TextRenderer.MeasureText(title, titleFont, new Size(titleW, int.MaxValue), Wrap).Height + Dpi.S(2));
            MarkTitle head = new MarkTitle { Bounds = new Rectangle(pad, Dpi.S(22), inner, titleH), Kind = kind, Font = titleFont, Text = title };
            head.AccessibleName = title; head.AccessibleRole = AccessibleRole.StaticText;

            // the text: as high as it needs, up to what the screen has room for; beyond that it scrolls
            int top = head.Bottom + Dpi.S(10);
            int need = TextRenderer.MeasureText(text, bodyFont, new Size(inner, int.MaxValue), Wrap).Height + Dpi.S(4);
            int room = Dpi.S(440);
            try
            {
                int frame = SystemInformation.CaptionHeight + 2 * SystemInformation.FixedFrameBorderSize.Height + Dpi.S(16);
                room = Math.Min(room, Screen.PrimaryScreen.WorkingArea.Height - frame - top - Dpi.S(24 + 76));
            }
            catch (Exception) { }
            room = Math.Max(room, Dpi.S(120));
            Control body;
            if (need <= room)
                body = new TextBlock { Bounds = new Rectangle(pad, top, inner, need), Font = bodyFont, Color = Theme.Text, Text = text };
            else
                body = new TextBox
                {
                    Bounds = new Rectangle(pad, top, inner, room), Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, WordWrap = true,
                    BorderStyle = BorderStyle.None, BackColor = Theme.Card, ForeColor = Theme.Text, Font = bodyFont, RightToLeft = RightToLeft.Yes,
                    TabStop = true, Text = text   // reachable with Tab, to scroll and copy with the keyboard
                };
            body.AccessibleName = text; body.AccessibleRole = AccessibleRole.StaticText;

            LowerPanel strip = new LowerPanel { Bounds = new Rectangle(0, body.Bottom + Dpi.S(24), width, Dpi.S(76)) };
            List<FlatBtn> buttons = new List<FlatBtn>();
            int x = width - pad;
            for (int i = 0; i < labels.Length; i++)
            {
                Font f = Theme.F(i == 0 ? 15 : 14, true);   // the fonts FlatBtn draws them with
                int w = Math.Max(Dpi.S(110), TextRenderer.MeasureText(labels[i], f).Width + Dpi.S(40));
                x -= w;
                FlatBtn b = new FlatBtn { Bounds = new Rectangle(x, Dpi.S(16), w, Dpi.S(44)), Text = labels[i], Style = i == 0 ? BtnStyle.Filled : BtnStyle.Outlined };
                b.AccessibleName = labels[i];
                b.AccessibleRole = AccessibleRole.PushButton;
                DialogResult a = answers[i];
                b.Click += delegate { answer = a; Close(); };
                buttons.Add(b);
                x -= Dpi.S(12);
            }
            strip.Controls.AddRange(buttons.ToArray());   // the tab order follows this order (right to left)
            Controls.AddRange(new Control[] { head, body, strip });
            ClientSize = new Size(width, strip.Bottom);
            FlatBtn first = buttons[Math.Max(0, Math.Min(focus, buttons.Count - 1))];
            first.AccessibleDescription = title + ". " + text;   // read out with the button that has the focus when the window opens
            ActiveControl = first;
        }

        readonly string copyText;

        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData == Keys.Escape) { answer = escape; Close(); return true; }
            return base.ProcessDialogKey(keyData);
        }

        // Ctrl+C copies the title and the text, as the Windows message box does (in the scrolling text box it copies what is selected there)
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.C) && !(ActiveControl is TextBox))
            {
                try { Clipboard.SetText(copyText); } catch (Exception) { }
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // the sound of the Windows message box of the same kind
        static void Sound(MsgKind kind)
        {
            if (Environment.OSVersion.Platform != PlatformID.Win32NT) return;
            try
            {
                if (kind == MsgKind.Warn || kind == MsgKind.Error) System.Media.SystemSounds.Exclamation.Play();
                else if (kind == MsgKind.Ok || kind == MsgKind.Info) System.Media.SystemSounds.Asterisk.Play();
            }
            catch (Exception) { }
        }

        public static DialogResult Show(IWin32Window owner, string text, string title, MsgKind kind, string[] labels, DialogResult[] answers, int focus, DialogResult escapeAnswer)
        {
            Sound(kind);
            using (MsgDialog d = new MsgDialog(text ?? "", title ?? "", kind, labels, answers, focus, escapeAnswer))
            {
                if (owner != null) d.ShowDialog(owner); else d.ShowDialog();
                return d.answer;
            }
        }

        // a message with one button
        public static void Tell(IWin32Window owner, string text, string title, MsgKind kind)
        {
            Show(owner, text, title, kind, new string[] { "אישור" }, new DialogResult[] { DialogResult.OK }, 0, DialogResult.OK);
        }

        // a yes / no question about something that cannot easily be undone: the focus (Enter) starts on "no"
        public static DialogResult YesNo(IWin32Window owner, string text, string title)
        {
            return Show(owner, text, title, MsgKind.Warn, new string[] { "כן", "לא" }, new DialogResult[] { DialogResult.Yes, DialogResult.No }, 1, DialogResult.No);
        }
    }

    // ------------------------------------------------------------------------------------------ main window
    class MainForm : Form
    {
        Header header; StatusLabel chip; TextBlock secTitle, stage, footer; InfoBar status; CardPanel tools;
        OptionRow cardBasic, cardExt, cardComp;
        FlatBtn btnInstall, btnRemove;
        IconBtn btnManual, btnWords, btnRepeat, btnAdvanced;
        ToolTip tips;
        ProgressLine prog;
        bool busy;
        bool asking;   // the "Word is open" window (or the short wait after it) is on: a second click must not start anything
        InstalledInfo lastInfo;   // what the last look at this computer found (used for the line "what will be installed")

        public const string Title = "המילון התורני 3.1";

        // The window icon, taken from the icon file inside the program WITH ALL ITS SIZES, so that Windows can pick the sharp one for the title bar,
        // the taskbar and Alt+Tab. (Icon.ExtractAssociatedIcon gives one size only, which then has to be shrunk and looks grainy.)
        public static Icon AppIcon()
        {
            try
            {
                using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("Otzaria.app.ico"))
                    if (s != null) return new Icon(s);
            }
            catch (Exception) { }
            try { return Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch (Exception) { return null; }
        }

        public MainForm()
        {
            Text = Title;
            AutoScaleMode = AutoScaleMode.None;   // every size is scaled by Dpi (see above), never by Windows Forms itself
            ClientSize = Dpi.Sz(720, Dpi.WindowHeight);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            if (Dpi.Forced) { StartPosition = FormStartPosition.Manual; Location = new Point(0, 0); }   // the developer's screenshots (--dpi): top left corner
            else StartPosition = FormStartPosition.CenterScreen;
            BackColor = Theme.Window;
            Font = Theme.F(14, false);
            DoubleBuffered = true;
            Icon = AppIcon();

            // 3.1: the light grey window with white cards. Top: the name (right) and the state of this computer as a pill (left).
            header = new Header { Bounds = Dpi.R(0, 0, 720, 96) };
            chip = new StatusLabel { Bounds = Dpi.R(32, 35, 300, 28) };
            header.Controls.Add(chip);
            secTitle = new TextBlock { Bounds = Dpi.R(440, 104, 248, 24), Font = Theme.F(16, true), Color = Theme.Text, VCenter = true, Wrap = false, Text = "בחרו מילון" };

            // the choice: three cards, one under the other, 8 apart
            cardBasic = new OptionRow
            {
                Bounds = Dpi.R(32, 136, 656, 84), Pos = RowPos.First, Title = "בסיסי", Selected = true,
                Desc = "מילים מהתנ״ך, מהמשנה, מהתלמוד הבבלי ומהתלמוד הירושלמי, וכל ראשי התיבות של 2-3 אותיות (כמו רש\"י, חז\"ל).",
                Count = CountText(false, false)
            };
            cardExt = new OptionRow
            {
                Bounds = Dpi.R(32, 228, 656, 84), Pos = RowPos.Middle, Title = "מורחב",
                Desc = "כולל את הבסיסי, ובנוסף מילים משאר הספרים: ראשונים, אחרונים, הלכה, מדרש ועוד, וראשי התיבות שמופיעים בהם.",
                Count = CountText(true, false)
            };
            cardComp = new OptionRow
            {
                Bounds = Dpi.R(32, 320, 656, 84), Pos = RowPos.Last, Title = "חיבור תורני", CheckStyle = true,
                Desc = "תוספת לבחירה: מילים בעברית עם אותיות חיבור כמו בארמית (דהעניין, ודלא, כדמצוה...), מילים שאינן במילון עצמו.",
                Count = ""   // filled in by UpdateSummary: how many words it adds depends on the edition chosen above it
            };
            // nothing changes while an action runs, or while the "Word is open" window (or the short wait after it) is on
            cardBasic.Chosen += delegate { if (!busy && !asking) { cardBasic.Selected = true; cardExt.Selected = false; UpdateSummary(); } };
            cardExt.Chosen += delegate { if (!busy && !asking) { cardExt.Selected = true; cardBasic.Selected = false; UpdateSummary(); } };
            cardComp.Chosen += delegate { if (!busy && !asking) { cardComp.Selected = !cardComp.Selected; UpdateSummary(); } };

            btnRemove = new FlatBtn { Bounds = Dpi.R(32, 424, 240, 48), Text = "הסר מוורד", Style = BtnStyle.Outlined };
            btnInstall = new FlatBtn { Bounds = Dpi.R(288, 424, 400, 48), Text = "התקן", Style = BtnStyle.Filled };
            btnInstall.Click += delegate { StartInstall(); };
            btnRemove.Click += delegate { StartRemove(); };

            prog = new ProgressLine { Bounds = Dpi.R(32, 486, 656, 4) };
            // the line under the progress bar: while the program works it shows what it is doing (stage); when it is done, the result (status).
            // They take the same place, and only one of them is visible at a time (SetStage).
            stage = new TextBlock { Bounds = Dpi.R(32, 498, 656, 22), Font = Theme.F(14, false), Color = Theme.TextSoft, Text = "", Visible = false };
            status = new InfoBar { Bounds = Dpi.R(32, 494, 656, 54) };
            prog.Visible = false;   // the progress line is shown only while an action runs; its result is shown in the bar under it

            // the tools: one quiet row in a white card - the person's own words (3.0.0.9), the manual route, repeated words, advanced removal
            tools = new CardPanel { Bounds = Dpi.R(32, 556, 656, 60) };
            btnWords = new IconBtn { Bounds = Dpi.R(494, 8, 154, 44), Text = "עריכת מילים", Icon = Glyph.Pencil };
            btnManual = new IconBtn { Bounds = Dpi.R(332, 8, 154, 44), Text = "התקנה ידנית", Icon = Glyph.Folder };
            btnRepeat = new IconBtn { Bounds = Dpi.R(170, 8, 154, 44), Text = "מילים כפולות", Icon = Glyph.Double };
            btnAdvanced = new IconBtn { Bounds = Dpi.R(8, 8, 154, 44), Text = "הסרה מתקדמת", Icon = Glyph.Trash };
            btnWords.Click += delegate { StartWords(); };
            btnManual.Click += delegate { StartManual(); };
            btnRepeat.Click += delegate { StartRepeated(); };
            btnAdvanced.Click += delegate { StartAdvanced(); };
            tools.Controls.AddRange(new Control[] { btnWords, btnManual, btnRepeat, btnAdvanced });
            tips = new ToolTip();
            tips.SetToolTip(btnWords, "הוספה, תיקון והסרה של מילים במילון, ושליחת הצעה לבעל המילון");
            tips.SetToolTip(btnManual, "אם ההתקנה האוטומטית לא עובדת: הקובץ נשמר, ומוסיפים אותו בוורד בעצמכם");
            tips.SetToolTip(btnAdvanced, "מוחקת מוורד את כל המילונים המותאמים אישית (נשמר גיבוי)");

            string wv = Core.WordVersion();
            footer = new TextBlock { Bounds = Dpi.R(32, 630, 656, 18), Font = Theme.F(12.5f, false), Color = Theme.TextQuiet, Center = true, Wrap = false,
                Text = "גרסה " + Core.AppVersion + "   ·   וורד שזוהה במחשב: " + (wv ?? "לא זוהה") };

            // the order of the Tab key follows this order: right to left, so "install" comes before "remove"
            Controls.AddRange(new Control[] { header, secTitle, cardBasic, cardExt, cardComp, btnInstall, btnRemove, prog, stage, status, tools, footer });
            RefreshState(true);
        }

        // shows what the program is doing (stage), in the place where the result of the last action is shown; "" = back to the result
        void SetStage(string s)
        {
            stage.Text = s ?? "";
            bool on = stage.Text.Length > 0;
            stage.Visible = on;
            status.Visible = !on;
        }

        public void Preview()
        {
            prog.Value = 100; SetStage("");
            Result r = new Result();
            Installer.Describe(r, "המורחב + חיבור תורני", Core.EditionCount(true, true).ToString("N0"), null, 1);
            status.Ok = r.Ok; status.Title = r.Title; status.Text = Headline(r.Text);
            cardExt.Selected = true; cardBasic.Selected = false; cardComp.Selected = true;
            chip.Text = "מותקן: מורחב + חיבור תורני · " + Core.AppVersion; chip.Kind = LabelKind.Ok;
            UpdateSummary();
        }

        bool closeWhenDone;                         // the window was closed during an action: it is hidden, and closes when the action ends
        System.Windows.Forms.Timer closeTimer;

        // Stopping an action in the middle could leave Word without its dictionary, a half-written file or a hidden Word running.
        // So a close during an action hides the window and lets the action finish; if it hangs (Word does not answer), the program closes after two minutes.
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (busy && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                if (closeWhenDone) return;
                if (Ask("אם תסגרו עכשיו, החלון ייסגר והפעולה תסתיים ברקע (לכל היותר שתי דקות).\r\n\r\nלסגור?", "הפעולה עדיין רצה") != DialogResult.Yes) return;
                if (!busy) { e.Cancel = false; base.OnFormClosing(e); return; }   // the action ended while the question was open: close now
                closeWhenDone = true;
                Hide();
                closeTimer = new System.Windows.Forms.Timer();
                closeTimer.Interval = 120000;
                closeTimer.Tick += delegate { closeTimer.Stop(); busy = false; Close(); };
                closeTimer.Start();
                return;
            }
            base.OnFormClosing(e);
        }

        // says on the install button exactly what a click will install, and how many words "chibur torani" adds to the chosen edition
        void UpdateSummary()
        {
            try
            {
                bool ext = cardExt.Selected, comp = cardComp.Selected;
                cardComp.Count = "+" + (Core.EditionCount(ext, true) - Core.EditionCount(ext, false)).ToString("N0") + " מילים";
                cardComp.Invalidate();
                string what = (ext ? "מורחב" : "בסיסי") + (comp ? " + חיבור תורני" : "");
                bool same = lastInfo != null && lastInfo.Edition == (ext ? 2 : 1) && lastInfo.Compose == comp
                            && lastInfo.Version != null && Core.CompareToCurrent(lastInfo.Version) == 0;
                btnInstall.SubText = what + " · " + Core.EditionCount(ext, comp).ToString("N0") + " מילים" + (same ? " · כבר מותקן" : "");
            }
            catch (Exception ex)
            {
                // a number that cannot be shown must never keep the window from opening
                Core.Log("summary failed: " + Core.Msg(ex));
                cardComp.Count = ""; cardComp.Invalidate();
                btnInstall.SubText = "";
            }
        }

        // the number of words of an edition for a card; if it cannot be found the window still opens, only without the number
        static string CountText(bool extended, bool compose)
        {
            try { return Core.EditionCount(extended, compose).ToString("N0") + " מילים"; }
            catch (Exception ex) { Core.Log("count failed: " + Core.Msg(ex)); return ""; }
        }

        // what the program found on this computer
        void RefreshState(bool selectInstalled)
        {
            InstalledInfo info = Core.DetectInstalled();
            lastInfo = info;
            if (info.Edition == 0 && !info.Compose)
            {
                chip.Text = "עדיין לא הותקן"; chip.Kind = LabelKind.None;
            }
            else
            {
                string name = (info.Edition == 2 ? "מורחב" : info.Edition == 1 ? "בסיסי" : "חיבור תורני בלבד") + (info.Edition > 0 && info.Compose ? " + חיבור תורני" : "");
                // short, so that it fits in the pill with Windows' fonts too (3.1 review): the installed version, or "an update is available"
                string ver;
                bool current = false;
                if (info.Version == null) ver = "יש עדכון";
                else
                {
                    int cmp = Core.CompareToCurrent(info.Version);
                    current = cmp == 0 || cmp == 1;
                    ver = cmp == -1 ? "יש עדכון" : info.Version;
                }
                chip.Text = "מותקן: " + name + " · " + ver;
                chip.Kind = current ? LabelKind.Ok : LabelKind.Warn;
            }
            chip.Invalidate();
            if (selectInstalled)
            {
                if (info.Edition == 2) { cardExt.Selected = true; cardBasic.Selected = false; }
                if (info.Compose) cardComp.Selected = true;
            }
            bool repeatOff = Stamp.ReadRepeatOff();
            btnRepeat.Text = repeatOff ? "החזר סימון כפולות" : "מילים כפולות";   // the button says which way it works now (as in 3.0)
            if (tips != null) tips.SetToolTip(btnRepeat, repeatOff ? "וורד לא מסמן עכשיו מילים כפולות. לחיצה מחזירה את הסימון."
                                                                   : "שוורד יפסיק לסמן כשגיאה מילה שחוזרת פעמיים ברצף (\"לא לא\")");
            UpdateSummary();
        }

        void UI(Action a)
        {
            try
            {
                if (IsDisposed || !IsHandleCreated) return;   // the window is gone: nothing to show (and never touch it from the worker thread)
                if (InvokeRequired) BeginInvoke((MethodInvoker)delegate { if (!IsDisposed) a(); });
                else a();
            }
            catch (InvalidOperationException) { }   // the window was closed in the meantime (ObjectDisposedException is one of these too)
        }

        void SetBusy(bool b)
        {
            busy = b;
            btnInstall.Enabled = !b; btnRemove.Enabled = !b; btnManual.Enabled = !b; btnWords.Enabled = !b; btnRepeat.Enabled = !b; btnAdvanced.Enabled = !b;
            UseWaitCursor = b;
        }

        DialogResult Ask(string text, string title)
        {
            return MsgDialog.YesNo(this, text, title);
        }

        // the result of an action: in a message window, and its title and first sentence stay in the bar of the main window
        void Say(Result r)
        {
            status.Ok = r.Ok;
            status.Title = r.Title;
            status.Text = Headline(r.Text);
            prog.Visible = false;
            MsgDialog.Tell(this, r.Text, r.Title, r.Ok ? MsgKind.Ok : MsgKind.Error);
        }

        // the first paragraph of a text, in one line
        internal static string Headline(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            int cut = text.IndexOf("\r\n\r\n", StringComparison.Ordinal);
            return (cut >= 0 ? text.Substring(0, cut) : text).Replace("\r\n", " ").Trim();
        }

        // for the developer: the message windows with real texts (--preview-msg done | manual | ask | fail | long)
        public static void PreviewMessage(string which)
        {
            Result r = new Result();
            if (which == "ask")
            {
                MsgDialog.YesNo(null, "לא רק את המילון התורני הזה, אלא גם מילונים ומילים שהוספתם בעצמכם (\"הוסף למילון\") או שנוספו על ידי תוכנות אחרות.\r\n\r\n"
                                      + "יישאר רק המילון המובנה של וורד. לפני המחיקה נשמר עותק גיבוי של הכול.\r\n\r\nלהמשיך?", "ההסרה המתקדמת מוחקת את כל המילונים המותאמים אישית");
                return;
            }
            if (which == "manual")
            {
                r.Ok = true; r.Title = "התקנה ידנית";
                r.Text = "הקבצים נשמרו בתיקייה שנפתחה עכשיו:\r\nC:\\Users\\User\\AppData\\Roaming\\Microsoft\\UProof\r\n\r\n"
                       + "עכשיו בוורד:\r\n1. קובץ ← אפשרויות ← הגהה\r\n2. לחצו על \"מילונים מותאמים אישית…\"\r\n3. לחצו \"הוספה…\" ובחרו את הקובץ:\r\n   Milon_Torani.dic\r\n"
                       + "4. אישור, ואם צריך לסגור ולפתוח את וורד מחדש.";
            }
            else if (which == "fail")
            {
                r.Ok = false; r.Title = "צריך להשלים ידנית";
                r.Text = "קובץ המילון נשמר, אבל לא הצלחתי לרשום אותו בוורד אוטומטית.\r\n\r\nלחצו על \"התקנה ידנית\" ופעלו לפי ההוראות (לוקח חצי דקה).\r\n\r\n"
                       + "פרטי השגיאה: Error HRESULT E_FAIL has been returned from a call to a COM component.";
            }
            else if (which == "long")
            {
                r.Ok = false; r.Title = "המילון הותקן, אבל וורד לא מזהה את כל המילים";
                System.Text.StringBuilder t = new System.Text.StringBuilder("המילון הבסיסי (69,498 מילים, בקובץ אחד) נוסף לוורד.\r\n\r\n");
                for (int i = 1; i <= 40; i++) t.Append("שורה " + i + " של טקסט ארוך מאוד, כדי לראות שהחלון לא יוצא מהמסך ושהטקסט נגלל.\r\n");
                r.Text = t.ToString();
            }
            else
                Installer.Describe(r, "המורחב + חיבור תורני", Core.EditionCount(true, true).ToString("N0"),
                    "בדיקה אוטומטית בוורד חדש שנפתח לבדיקה: כל 16 מילות הבדיקה מזוהות ✓\r\nמילונים בוורד: CUSTOM.DIC, Milon_Torani.dic, RoamingCustom.dic", 1);
            MsgDialog.Tell(null, r.Text, r.Title, r.Ok ? MsgKind.Ok : MsgKind.Error);
        }

        void Work(Func<Result> job)
        {
            if (busy) return;
            SetBusy(true);
            status.Title = ""; status.Text = "";
            prog.Visible = true;
            prog.Value = 2; SetStage("מתחיל…");
            Thread t = new Thread(delegate ()
            {
                Result r;
                try { r = job(); }
                catch (Exception ex)
                {
                    Core.Log("ERROR " + ex);
                    r = new Result { Ok = false, Title = "אירעה שגיאה", Text = "משהו השתבש:\r\n" + Core.Msg(ex) + "\r\n\r\nקובץ יומן לשליחה: " + Core.LogPath };
                }
                UI(delegate
                {
                    if (closeWhenDone) { if (closeTimer != null) closeTimer.Stop(); busy = false; Close(); return; }   // closed during the action
                    prog.Value = r.Ok ? 100 : 0; SetStage(""); SetBusy(false); RefreshState(false); Say(r);
                });
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);   // Word automation prefers an STA thread
            t.Start();
        }

        // Before an installation or a removal: is Word closed? A Word that is running does not take a change in its dictionaries until it is
        // closed and opened again (and an install that finds Word open cannot check the result), so the person is asked to close it first.
        // true = go on (Word is closed, or its state cannot be told, or the person chose to go on anyway); false = do not start.
        bool WordIsClosedOrIgnored(bool installing)
        {
            asking = true;
            try
            {
                while (true)
                {
                    int n = WordCom.RunningWordCount();
                    if (n <= 0) return true;   // 0 = closed; -1 = cannot be told, and the person is not held up for that
                    Core.Log("Word is running (" + n + " process(es)) before " + (installing ? "the installation" : "the removal") + ": the person is asked to close it");
                    int choice;
                    using (WordOpenDialog dlg = new WordOpenDialog(installing)) { dlg.ShowDialog(this); choice = dlg.Choice; }
                    if (choice == WordOpenDialog.GoOn) { Core.Log("The person chose to go on while Word is open"); return true; }
                    if (choice != WordOpenDialog.Again || IsDisposed) return false;

                    // "check again": Word needs a moment to really disappear after its last window was closed
                    SetStage("בודק שוורד נסגר…");
                    UseWaitCursor = true;
                    try
                    {
                        for (int i = 0; i < 20 && !IsDisposed && WordCom.RunningWordCount() > 0; i++)
                        {
                            Thread.Sleep(250);
                            Application.DoEvents();
                        }
                    }
                    finally { if (!IsDisposed) { UseWaitCursor = false; SetStage(""); } }
                    if (IsDisposed) return false;
                }
            }
            finally { asking = false; }
        }

        void StartInstall()
        {
            if (busy || asking) return;
            if (!WordIsClosedOrIgnored(true)) return;
            bool ext = cardExt.Selected, comp = cardComp.Selected;   // read after the "Word is open" window: what is chosen now is what is installed
            Work(delegate { return Installer.Install(ext, comp, delegate (int p, string s) { UI(delegate { prog.Value = p; SetStage(s); }); }, true); });
        }

        void StartRemove()
        {
            if (busy || asking) return;
            if (!WordIsClosedOrIgnored(false)) return;
            Work(delegate { return Installer.Remove(true); });
        }

        // the word-editing window (Words.cs). It shows the dictionary that is installed (or, if none is, the one chosen above);
        // after "save" the same dictionary is installed again, with the person's changes, so that Word has them.
        void StartWords()
        {
            if (busy || asking) return;
            InstalledInfo info = Core.DetectInstalled();
            bool installed = info.Edition > 0;
            bool ext = installed ? info.Edition == 2 : cardExt.Selected;
            bool comp = installed ? info.Compose : cardComp.Selected;
            WordsForm wf;
            UseWaitCursor = true;
            Cursor.Current = Cursors.WaitCursor;
            try { wf = new WordsForm(ext, comp, installed, false); }
            catch (Exception ex)
            {
                Core.Log("Opening the word editing failed: " + ex);
                Say(new Result { Ok = false, Title = "עריכת מילים", Text = "לא הצלחתי לפתוח את עריכת המילים:\r\n" + Core.Msg(ex) + "\r\n\r\nקובץ יומן לשליחה: " + Core.LogPath });
                return;
            }
            finally { UseWaitCursor = false; }
            using (wf)
            {
                wf.ShowDialog(this);
                if (!wf.Saved) return;
                string head = "השינויים שלכם במילון נשמרו"
                            + (wf.Sent ? ", וההצעה נשלחה לבעל המילון. תודה!" : wf.ViaBrowser ? ", ודף ההצעה נפתח בגיטהב." : ".");
                if (!installed)
                {
                    Say(new Result { Ok = true, Title = "השינויים נשמרו", Text = head + "\r\n\r\nהם יחולו כשתתקינו את המילון." });
                    return;
                }
                if (!WordIsClosedOrIgnored(true))
                {
                    Say(new Result { Ok = true, Title = "השינויים נשמרו", Text = head + "\r\n\r\nוורד עוד לא עודכן. כדי שהשינויים יחולו, סגרו את וורד ולחצו על \"התקן\"." });
                    return;
                }
                Work(delegate
                {
                    Result r = Installer.Install(ext, comp, delegate (int p, string s) { UI(delegate { prog.Value = p; SetStage(s); }); }, true);
                    r.Text = head + "\r\n\r\n" + r.Text;
                    return r;
                });
            }
        }

        // switch Word's "flag repeated words" off (for ALL repeated words) - or back on
        void StartRepeated()
        {
            if (busy || asking) return;
            if (Stamp.ReadRepeatOff())
            {
                if (Ask("וורד יחזור לסמן כשגיאה מילים שחוזרות פעמיים ברצף.", "להחזיר את סימון המילים הכפולות בוורד?") != DialogResult.Yes) return;
                Work(delegate { return Installer.SetRepeatedWords(true); });
                return;
            }
            string warn = "שימו לב: האפשרות הזאת חלה על כלל המילים הכפולות.\r\n\r\n"
                        + "וורד יפסיק לסמן כשגיאה כל מילה שחוזרת פעמיים ברצף (למשל \"לא לא\" או \"הדר הדר\"). "
                        + "זה לא רק מילים תורניות, וזה חל על כל המסמכים.\r\n\r\n"
                        + "אפשר להחזיר את הסימון בכל זמן, באותו כפתור.\r\n\r\n"
                        + "להמשיך?";
            if (Ask(warn, "הסרת שגיאות ממילים כפולות") != DialogResult.Yes) return;
            Work(delegate { return Installer.SetRepeatedWords(false); });
        }

        // remove every custom dictionary from Word, whoever added it
        void StartAdvanced()
        {
            if (busy || asking) return;
            string warn ="זהירות: ההסרה המתקדמת מוחקת מוורד את כל המילונים המותאמים אישית.\r\n\r\n"
                        + "לא רק את המילון התורני הזה, אלא גם מילונים ומילים שהוספתם בעצמכם (\"הוסף למילון\") או שנוספו על ידי תוכנות אחרות.\r\n\r\n"
                        + "יישאר רק המילון המובנה של וורד.\r\n\r\n"
                        + "לפני המחיקה נשמר עותק גיבוי של הכול.\r\n\r\n"
                        + "להמשיך?";
            if (Ask(warn, "הסרה מתקדמת") != DialogResult.Yes) return;
            if (!WordIsClosedOrIgnored(false)) return;   // asked only after the person agreed to the removal itself
            Work(delegate { return Installer.RemoveAdvanced(true); });
        }

        void StartManual()
        {
            if (busy || asking) return;
            bool ext = cardExt.Selected, comp = cardComp.Selected;
            Work(delegate
            {
                List<string> names = Installer.WriteForManual(ext, comp);
                string dir = Core.UProofDir();
                try { Process.Start("explorer.exe", "\"" + dir + "\""); } catch (Exception) { }
                Result r = new Result();
                r.Ok = true; r.Title = "התקנה ידנית";
                r.Text = "הקבצים נשמרו בתיקייה שנפתחה עכשיו:\r\n" + dir + "\r\n\r\n"
                       + "עכשיו בוורד:\r\n"
                       + "1. קובץ ← אפשרויות ← הגהה\r\n"
                       + "2. לחצו על \"מילונים מותאמים אישית…\"\r\n"
                       + "3. לחצו \"הוספה…\" ובחרו את הקובץ:\r\n   " + string.Join("\r\n   ", names.ToArray()) + "\r\n"
                       + "4. אישור, ואם צריך לסגור ולפתוח את וורד מחדש.";
                return r;
            });
        }
    }

    static class Program
    {
        [DllImport("user32.dll")] static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool ShowWindow(IntPtr hWnd, int cmd);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);

        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length >= 2 && args[0] == "--selftest")
            {
                // the self-test writes this user's real Word settings in the registry (the install record, the dictionary list): developer machine only
                if (Type.GetType("Mono.Runtime") == null) { Console.WriteLine("The self-test runs only in the developer's environment (Mono)."); return 2; }
                return SelfTest(args[1]);
            }

            // how big is the screen's magnification (125 %, 150 %...)? Must come before the first window. "--dpi 150" = pretend 150 % (developer's screenshots)
            int forcedDpi = 0;
            for (int i = 0; i + 1 < args.Length; i++) if (args[i] == "--dpi") int.TryParse(args[i + 1], out forcedDpi);
            Dpi.Init(forcedDpi);

            if (args.Length >= 1 && (args[0] == "--preview-words" || args[0] == "--preview-suggest"))   // for the developer: the word-editing windows, with sample changes that are never saved
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                if (args[0] == "--preview-words") using (WordsForm f = new WordsForm(true, false, true, true)) { f.Shown += delegate { f.Demo(); }; f.ShowDialog(); }
                else using (SuggestDialog d = new SuggestDialog(WordsForm.DemoChanges(), "מורחב")) d.ShowDialog();
                return 0;
            }

            if (args.Length >= 1 && args[0] == "--preview-msg")   // for the developer: the message windows (done | manual | ask | fail | long)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                MainForm.PreviewMessage(args.Length >= 2 && !args[1].StartsWith("--", StringComparison.Ordinal) ? args[1] : "done");
                return 0;
            }

            if (args.Length >= 1 && args[0] == "--preview-wordopen")   // for the developer: look at the "Word is open" window, without Word ("remove" = the removal wording)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                using (WordOpenDialog d = new WordOpenDialog(!(args.Length >= 2 && args[1] == "remove"))) d.ShowDialog();
                return 0;
            }

            // one window at a time: two copies working together would fight over the same files
            bool first;
            Mutex single = null;
            try { single = new Mutex(true, @"Local\TorahDictionaryForWord.SingleInstance", out first); }
            catch (Exception) { first = true; }
            if (!first)
            {
                bool shown = false;
                try
                {
                    IntPtr h = OtherWindow();
                    if (h != IntPtr.Zero) { if (IsIconic(h)) ShowWindow(h, 9); shown = SetForegroundWindow(h); }   // 9 = restore
                }
                catch (Exception) { }
                if (!shown)
                    MessageBox.Show("המילון התורני כבר פתוח.\r\n\r\nחפשו את החלון שלו בשורת המשימות.", MainForm.Title, MessageBoxButtons.OK, MessageBoxIcon.Information,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                return 0;
            }

            try
            {
                Core.TrimLog();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.ThreadException += delegate (object s, ThreadExceptionEventArgs e) { Fatal(e.Exception); };
                AppDomain.CurrentDomain.UnhandledException += delegate (object s, UnhandledExceptionEventArgs e) { Fatal(e.ExceptionObject as Exception); };
                try { MainForm mf = new MainForm(); if (args.Length >= 1 && args[0] == "--preview") mf.Preview(); Application.Run(mf); }
                catch (Exception ex) { Fatal(ex); }
            }
            finally
            {
                if (single != null)
                {
                    try { single.ReleaseMutex(); } catch (Exception) { }
                    single.Close();
                }
            }
            return 0;
        }

        // The window of the copy that is already running, found through its process. (Looking for the title also found a folder with the
        // same name open in Explorer.) Zero = not found: the person then gets a short message instead.
        static IntPtr OtherWindow()
        {
            try
            {
                using (Process me = Process.GetCurrentProcess())
                {
                    foreach (Process p in Process.GetProcessesByName(me.ProcessName))
                    {
                        try { if (p.Id != me.Id && p.MainWindowHandle != IntPtr.Zero) return p.MainWindowHandle; }
                        catch (Exception) { }
                        finally { p.Dispose(); }
                    }
                }
            }
            catch (Exception) { }
            return IntPtr.Zero;
        }

        static void Fatal(Exception ex)
        {
            Core.Log("FATAL " + ex);
            try
            {
                MessageBox.Show("אירעה שגיאה בתוכנה:\r\n" + (ex == null ? "" : Core.Msg(ex)) + "\r\n\r\nקובץ יומן: " + Core.LogPath,
                    MainForm.Title, MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            catch (Exception) { }
        }

        // ---------------------------------------------------------------- self test: runs the install logic without Word, against a temporary folder
        static int fails;
        static void Check(bool ok, string what)
        {
            Console.WriteLine((ok ? "  ok    " : "  FAIL  ") + what);
            if (!ok) fails++;
        }

        static bool SameBytes(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
            return true;
        }

        static int WordsIn(string path, out bool bom, out string first)
        {
            byte[] raw = File.ReadAllBytes(path);
            bom = raw.Length > 1 && raw[0] == 0xFF && raw[1] == 0xFE;
            string txt = System.Text.Encoding.Unicode.GetString(raw, 2, raw.Length - 2);
            string[] lines = txt.Split(new string[] { "\r\n" }, StringSplitOptions.RemoveEmptyEntries);
            first = lines.Length > 0 ? lines[0] : "";
            return lines.Length;
        }

        // ---------------------------------------------------------------- a pretend Word for the self-test
        // WordCom talks to Word by member names (late binding), so a pretend Word only needs members with the same names.
        public class FakeDict
        {
            readonly string name, folder;
            readonly FakeDictList owner;
            bool languageSpecific = true;
            public FakeDict(FakeDictList owner, string name, string folder) { this.owner = owner; this.name = name; this.folder = folder; }
            public string Name { get { return name; } }
            public string Path { get { return folder; } }   // like the real one: the folder, not the file
            public bool LanguageSpecific { get { return languageSpecific; } set { languageSpecific = value; } }
            public void Delete()
            {
                if (owner.App.Refuse.Contains(name)) throw new InvalidOperationException("simulated: Word refused to remove " + name);
                owner.Items.Remove(this); owner.App.Deleted++;
                if (owner.Active == this) owner.Active = null;
            }
        }

        public class FakeDictList
        {
            public FakeApp App;
            public List<FakeDict> Items = new List<FakeDict>();
            public FakeDict Active;      // the default dictionary, where Word puts the words the person adds
            public bool FailAdd;         // true = Word refuses to add a dictionary
            public Func<bool> FailAddWhen;   // Word refuses while this says so (3.0.0.10: its own list in the registry is full)
            public int Count { get { return Items.Count; } }
            public FakeDict Item(int i) { return Items[i - 1]; }   // Word counts from 1
            // like Word: when none was chosen, the first dictionary of the list is the default
            public FakeDict ActiveCustomDictionary
            {
                get { return Active ?? (Items.Count > 0 ? Items[0] : null); }
                set { Active = value; }
            }
            public FakeDict Add(string path)
            {
                if (FailAdd || (FailAddWhen != null && FailAddWhen())) throw new InvalidOperationException("simulated: Word could not add the dictionary");
                App.Added.Add(path);
                FakeDict d = new FakeDict(this, System.IO.Path.GetFileName(path), System.IO.Path.GetDirectoryName(path));
                Items.Add(d);
                return d;
            }
        }

        public class FakeApp
        {
            public FakeDictList Dicts;
            public List<string> Added = new List<string>();
            public int Deleted, DocumentsTouched;
            public bool Quitted;
            public bool InEffect;                                   // true = this Word knows the words of the dictionaries it lists (a Word that has just started)
            public HashSet<string> Known = new HashSet<string>();   // the words it knows when InEffect
            public HashSet<string> Refuse = new HashSet<string>(StringComparer.OrdinalIgnoreCase);   // dictionaries this Word will not remove
            public Func<string> Probe;   // asked when this Word quits; the answer is kept in AtQuit
            public string AtQuit;
            public FakeApp() { Dicts = new FakeDictList(); Dicts.App = this; }
            public FakeDictList CustomDictionaries { get { return Dicts; } }
            public FakeDocs Documents { get { DocumentsTouched++; FakeDocs d = new FakeDocs(); d.App = this; return d; } }
            public void Quit(params object[] args) { Quitted = true; if (Probe != null) AtQuit = Probe(); }
            public bool Marks(string word)
            {
                if (word == "קזקזקזק") return true;    // nonsense is always marked
                if (word == "שלום") return false;      // an ordinary word never is
                return !(InEffect && Known.Contains(word));
            }
        }

        public class FakeDocs { public FakeApp App; public FakeDoc Add(params object[] args) { FakeDoc d = new FakeDoc(); d.App = App; return d; } }
        public class FakeDoc
        {
            public FakeApp App;
            public FakeRange Content { get { FakeRange r = new FakeRange(); r.App = App; return r; } }
            public void Close(params object[] args) { }
        }
        public class FakeRange
        {
            public FakeApp App;
            string text = "";
            public string Text { get { return text; } set { text = value; } }
            public int LanguageID { get; set; }
            public bool NoProofing { get; set; }
            public FakeErrors SpellingErrors { get { FakeErrors e = new FakeErrors(); e.Count = App.Marks(text) ? 1 : 0; return e; } }
        }
        public class FakeErrors { public int Count { get; set; } }

        static int fakeNewWords;   // how many times the installer asked for a NEW Word in the last InstallWithFakeWord

        // runs the installer against pretend Words: 'first' is the Word it finds (already open) or starts, 'fresh' the Word it starts afterwards for the check
        static Result InstallWithFakeWord(FakeApp first, bool firstWasOpen, FakeApp fresh)
        {
            fakeNewWords = 0;
            WordCom.OpenHook = delegate { WordSession s = new WordSession(); s.App = first; s.Created = !firstWasOpen; return s; };
            WordCom.OpenNewHook = delegate
            {
                fakeNewWords++;
                if (fresh == null) throw new InvalidOperationException("no new Word was expected");
                WordSession s = new WordSession(); s.App = fresh; s.Created = true; return s;
            };
            try { return Installer.Install(false, false, delegate (int p, string t) { }, true); }
            finally { WordCom.OpenHook = null; WordCom.OpenNewHook = null; }
        }

        static int SelfTest(string dir)
        {
            Core.UProofOverride = dir;
            MyWords.FolderOverride = Path.Combine(dir, "settings");   // the person's own changes (stage 4): never the real ones
            Directory.CreateDirectory(dir);
            fails = 0;
            ProgressFn quiet = delegate (int p, string t) { };
            Console.WriteLine("version " + Core.AppVersion);

            Console.WriteLine("-- acronyms");
            List<string> a2 = Core.MakeAcronyms(2, 2), a3 = Core.MakeAcronyms(3, 3), a4 = Core.MakeAcronyms(4, 4);
            Check(a2.Count == 484 && a3.Count == 10648 && a4.Count == 234256, "counts 2/3/4 letters = " + a2.Count + "/" + a3.Count + "/" + a4.Count);
            HashSet<string> s2 = new HashSet<string>(a2), s3 = new HashSet<string>(a3), s4 = new HashSet<string>(a4);
            Check(s2.Contains("ר\"ת") && s2.Contains("ר\"ח"), "2 letters: ר\"ת, ר\"ח");
            Check(s3.Contains("רש\"י") && s3.Contains("חז\"ל") && s3.Contains("שו\"ת"), "3 letters: רש\"י, חז\"ל, שו\"ת");
            Check(s4.Contains("רמב\"ם") && s4.Contains("רשב\"א") && s4.Contains("מהר\"ם"), "4 letters: רמב\"ם, רשב\"א, מהר\"ם");
            bool badFinal = false;
            foreach (string w in a4) { if (w.IndexOf('ך') >= 0 && w.IndexOf('ך') != w.Length - 1) badFinal = true; }
            Check(!badFinal, "final-form letters appear only at the end");

            Console.WriteLine("-- composition (chibur torani)");
            List<string> comp = Core.MakeComposition();
            int baseCount = Core.CountWords(Core.BaseRes);
            Check(comp.Count == baseCount * Core.Connectors.Length, "count = " + comp.Count + " (" + baseCount + " words x " + Core.Connectors.Length + " connectors)");
            Check(comp.Contains("דהוא") && comp.Contains("ודהוא") && comp.Contains("כדהוא"), "connectors ד, וד, כד in front of a base word");

            Console.WriteLine("-- install: extended + composition (over an old 3.0.0.1 install with five files)");
            foreach (string f in Core.LegacyFiles) Core.WriteDic(Path.Combine(dir, f), new List<string> { "ישן", "קובץ" });
            Result r = Installer.Install(true, true, delegate (int p, string t) { Console.WriteLine("  " + p + "% " + t); }, false);
            Console.WriteLine("  -> ok=" + r.Ok + " | " + r.Title);
            string main = Path.Combine(dir, Core.MainFile);
            bool bom = false; string first = "";
            int n = File.Exists(main) ? WordsIn(main, out bom, out first) : 0;
            Console.WriteLine("  " + Core.MainFile + ": " + n + " words, first=" + first);
            List<string> expected = Core.BuildWords(true, true);
            Check(bom && n == expected.Count && n > 150000, "one file with BOM and all " + expected.Count + " words");
            Check(new HashSet<string>(expected).Count == expected.Count, "no duplicate words");
            Check(Core.CountDicWords(main) == n, "word counter agrees");
            string refFile = Path.Combine(dir, "reference.tmp");
            Core.WriteDic(refFile, expected);
            Check(SameBytes(File.ReadAllBytes(main), File.ReadAllBytes(refFile)), "the installed file is byte for byte what WriteDic(BuildWords) writes");
            File.Delete(refFile);
            Check(expected.IndexOf("מיבעי") < expected.IndexOf("דמסתברא") && expected.IndexOf("דמסתברא") < expected.IndexOf("ודהנושא"), "order: basic, then extended, then composition");
            bool anyOld = false;
            foreach (string f in Core.LegacyFiles) if (File.Exists(Path.Combine(dir, f))) anyOld = true;
            Check(!anyOld, "files of older versions removed");
            InstalledInfo info = Core.DetectInstalled();
            Check(info.Edition == 2 && info.Compose && info.Version == Core.AppVersion, "detected: extended + composition, version " + info.Version);
            Check(Core.CompareToCurrent(info.Version) == 0, "installed version equals program version");

            Console.WriteLine("-- install: basic only");
            Installer.Install(false, false, quiet, false);
            info = Core.DetectInstalled();
            Check(info.Edition == 1 && !info.Compose, "detected: basic, no composition");
            int nb = WordsIn(main, out bom, out first);
            Check(nb == Core.BuildWords(false, false).Count && nb < Core.BuildWords(true, false).Count, "basic file has " + nb + " words (fewer than the extended one)");

            Console.WriteLine("-- advanced removal (foreign dictionaries too)");
            Core.WriteDic(Path.Combine(dir, "CUSTOM.DIC"), new List<string> { "מילהפרטית", "עודאחת" });
            Core.WriteDic(Path.Combine(dir, "OtherProgram.dic"), new List<string> { "דוגמה" });
            Result ra = Installer.RemoveAdvanced(false);
            Console.WriteLine("  -> ok=" + ra.Ok + " | " + ra.Title);
            int left = 0;
            foreach (string f in Directory.GetFiles(dir)) if (f.EndsWith(".dic", StringComparison.OrdinalIgnoreCase)) left++;
            Check(left == 0, "no .dic file left in the folder");
            string[] backups = Directory.GetDirectories(dir, "Backup_*");
            Check(backups.Length == 1 && File.Exists(Path.Combine(backups[0], "CUSTOM.DIC")) && File.Exists(Path.Combine(backups[0], Core.MainFile)), "backup folder holds CUSTOM.DIC and our files");
            Check(Core.DetectInstalled().Edition == 0, "detected: nothing installed");

            Console.WriteLine("-- plain remove");
            Installer.Install(true, true, quiet, false);
            Result rr = Installer.Remove(false);
            Check(Core.DetectInstalled().Edition == 0 && !Core.DetectInstalled().Compose, "remove clears everything of ours | " + rr.Text);

            Console.WriteLine("-- exact word counts of the editions (shown on the cards without opening the word lists)");
            bool[] eFlags = { false, false, true, true };
            bool[] cFlags = { false, true, false, true };
            for (int i = 0; i < 4; i++)
            {
                int built = Core.BuildWords(eFlags[i], cFlags[i]).Count;
                int listed = Core.EditionCount(eFlags[i], cFlags[i]);
                Check(built == listed, (eFlags[i] ? "extended" : "basic") + (cFlags[i] ? " + composition" : "") + ": built " + built + ", counts file says " + listed);
            }

            Console.WriteLine("-- Word is busy: wait and ask again (no Word needed)");
            WordCom.RetryMs = 3000;
            int calls = 0;
            object got = WordCom.Retry<object>(delegate
            {
                calls++;
                if (calls < 3) throw new TargetInvocationException(new COMException("busy", unchecked((int)0x80010001)));
                return "answer";
            });
            Check("answer".Equals(got) && calls == 3, "two 'call rejected' answers, then success (" + calls + " calls)");
            calls = 0;
            try { WordCom.Retry<object>(delegate { calls++; throw new COMException("server gone", unchecked((int)0x800706BA)); }); }
            catch (COMException) { }
            Check(calls == 1, "a real failure is not retried (" + calls + " call)");
            WordCom.RetryMs = 600; calls = 0;
            try { WordCom.Retry<object>(delegate { calls++; throw new COMException("busy", unchecked((int)0x8001010A)); }); }
            catch (COMException) { }
            Check(calls >= 2 && calls <= 6, "gives up after the time limit (" + calls + " calls)");
            WordCom.RetryMs = 15000;
            Check(WordCom.IsBusy(new TargetInvocationException(new COMException("x", unchecked((int)0x8001010A)))) && !WordCom.IsBusy(new InvalidOperationException()), "busy detection");
            Check(Core.Msg(new COMException("boom", unchecked((int)0x80010001))).Contains("0x80010001"), "error text shows the Windows error number");
            Check(!WordCom.WaitForOurWordsToExit(100), "no hidden Word of ours to wait for -> reports it did not wait");

            Console.WriteLine("-- finding a dictionary file when Word gives no usable folder");
            string probe = Path.Combine(dir, "RoamingCustom.dic");
            Core.WriteDic(probe, new List<string> { "אאא", "בבב", "גגג" });
            Check(Core.ResolveDicPath(null, "RoamingCustom.dic") == probe, "no folder from Word -> found in the UProof folder");
            Check(Core.ResolveDicPath("", "RoamingCustom.dic") == probe, "empty folder from Word -> found in the UProof folder");
            Check(Core.ResolveDicPath(Path.Combine(dir, "no-such-folder"), "RoamingCustom.dic") == probe, "wrong folder from Word -> found in the UProof folder");
            Check(Core.ResolveDicPath(dir, "RoamingCustom.dic") == probe, "right folder from Word");
            Check(Core.ResolveDicPath(null, "Missing.dic") == null, "file that does not exist -> null");
            Check(Core.CountDicWords(probe) == 3, "word count of that dictionary: 3");
            File.Delete(probe);

            Console.WriteLine("-- Word version label");
            Check(Core.WordVersionLabel("16", "O365ProPlusRetail") == "Microsoft 365", "O365ProPlusRetail -> Microsoft 365");
            Check(Core.WordVersionLabel("16", "ProPlus2021Retail") == "2021", "ProPlus2021Retail -> 2021");
            Check(Core.WordVersionLabel("16", "ProPlus2024Volume") == "2024", "ProPlus2024Volume -> 2024");
            Check(Core.WordVersionLabel("16", "ProPlus2019Retail") == "2019", "ProPlus2019Retail -> 2019");
            Check(Core.WordVersionLabel("16", null) == "2016 / 2019 / 2021 / 365", "no install info -> all the 16.x names");
            Check(Core.WordVersionLabel("15", null) == "2013" && Core.WordVersionLabel("14", null) == "2010", "older Word numbers");

            Console.WriteLine("-- the text after registering (3.0.0.5): never a wrong 'Word does not recognize' when Word was already open");
            Check(WordCom.RunningWordCount() >= -1, "counting the running Word processes does not fail (" + WordCom.RunningWordCount() + ")");
            Result d1 = new Result(); d1.WordWasOpen = true;
            Installer.Describe(d1, "הבסיסי", "63,074", null, -1);
            Check(d1.Ok && d1.Title.Contains("נשאר לסגור") && d1.Text.Contains("מיבעי") && !d1.Text.Contains("לא מזהה") && !d1.Text.Contains("מתוך"), "Word was open, no check: neutral text | " + d1.Title);
            Result d2 = new Result(); d2.WordWasOpen = true;
            Installer.Describe(d2, "הבסיסי", "63,074", "בדיקה אוטומטית בוורד שהיה פתוח: 0 מתוך 14 מילות בדיקה מזוהות.", 0);
            Check(d2.Ok && d2.Title.Contains("נשאר לסגור") && !d2.Text.Contains("מתוך") && !d2.Text.Contains("לא מזהה"), "Word was open: even a stray 'not recognized' answer is never shown");
            Result d3 = new Result();
            Installer.Describe(d3, "הבסיסי", "63,074", "בדיקה אוטומטית בוורד חדש שנפתח לבדיקה: כל 15 מילות הבדיקה מזוהות ✓", 1);
            Check(d3.Ok && d3.Title == "ההתקנה הושלמה" && d3.Text.Contains("מזוהות ✓"), "fresh Word, all words known: done");
            Result d4 = new Result();
            Installer.Describe(d4, "הבסיסי", "63,074", "בדיקה אוטומטית בוורד חדש שנפתח לבדיקה: 3 מתוך 15 מילות בדיקה מזוהות.", 0);
            Check(!d4.Ok && d4.Title.Contains("לא מזהה") && d4.Text.Contains("מתוך"), "fresh Word, words not known: says so");
            Result d5 = new Result();
            Installer.Describe(d5, "הבסיסי", "63,074", "הבדיקה האוטומטית לא הצליחה (x).", -1);
            Check(d5.Ok && d5.Title == "ההתקנה הושלמה", "the check could not run: the installation is still reported as done");

            Console.WriteLine("-- installing while Word is already open (the 13:53 case), against a pretend Word");
            string mainPath = Path.Combine(dir, Core.MainFile);
            FakeApp open = new FakeApp();
            open.Dicts.Items.Add(new FakeDict(open.Dicts, Core.MainFile, dir));          // the old entry of ours
            open.Dicts.Items.Add(new FakeDict(open.Dicts, "RoamingCustom.dic", dir));    // the person's own dictionary
            Result ro = InstallWithFakeWord(open, true, null);
            Check(ro.WordWasOpen, "it attached to the Word that was already open");
            Check(open.DocumentsTouched == 0 && fakeNewWords == 0, "nothing was opened in that Word and no new Word was started: the check is skipped");
            Check(open.Deleted == 1 && open.Added.Count == 1 && open.Added[0] == mainPath, "the old entry of ours was replaced by one new entry (deleted " + open.Deleted + ", added " + open.Added.Count + ")");
            FakeDict mine = open.Dicts.Items.Find(delegate (FakeDict x) { return x.Name == "RoamingCustom.dic"; });
            FakeDict ours = open.Dicts.Items.Find(delegate (FakeDict x) { return x.Name == Core.MainFile; });
            Check(open.Dicts.Items.Count == 2 && mine != null && ours != null, "the person's own dictionary is still there, ours is there once");
            Check(ours != null && !ours.LanguageSpecific, "ours applies to every language");
            Check(ro.Ok && ro.Title.Contains("נשאר לסגור") && !ro.Text.Contains("לא מזהה") && !ro.Text.Contains("מתוך"), "the answer is neutral: " + ro.Title);
            Check(!open.Quitted, "that Word was left running (it is not ours to close)");
            Check(Core.DetectInstalled().Edition == 1, "the installation is recorded (basic)");

            Console.WriteLine("-- installing while Word is closed: the next, new Word makes the check");
            FakeApp hidden = new FakeApp();
            FakeApp fresh = new FakeApp();
            fresh.InEffect = true;
            fresh.Known = new HashSet<string>(Core.BuildWords(false, false));
            fresh.Dicts.Items.Add(new FakeDict(fresh.Dicts, Core.MainFile, dir));
            fresh.Dicts.Items.Add(new FakeDict(fresh.Dicts, "RoamingCustom.dic", dir));
            Result rf = InstallWithFakeWord(hidden, false, fresh);
            Check(!rf.WordWasOpen && hidden.Added.FindAll(delegate (string p) { return p == mainPath; }).Count == 1 && hidden.DocumentsTouched == 0, "the dictionary went into the hidden Word, nothing was checked inside it");
            Check(fakeNewWords == 1 && fresh.DocumentsTouched > 0, "the check was made in a new Word");
            Check(hidden.Quitted && fresh.Quitted, "both hidden Words were closed");
            Check(rf.Ok && rf.Title == "ההתקנה הושלמה" && rf.Text.Contains("מזוהות ✓"), "all sample words known: " + rf.Title);

            FakeApp hidden2 = new FakeApp();
            FakeApp blind = new FakeApp();   // a new Word that does NOT know the dictionary
            Result rb = InstallWithFakeWord(hidden2, false, blind);
            Check(!rb.Ok && rb.Title.Contains("לא מזהה") && rb.Text.Contains("מתוך"), "a new Word that does not know the words: says so (" + rb.Title + ")");
            Check(hidden2.Quitted && blind.Quitted, "both hidden Words were closed");

            Console.WriteLine("-- the dictionary file cannot be written (locked / no permission): a clear message, not a crash");
            if (File.Exists(main)) File.Delete(main);
            Directory.CreateDirectory(main);   // a folder where the file should be, so writing there must fail
            Result rl = Installer.Install(false, false, quiet, false);
            Check(!rl.Ok && rl.Title.Contains("לא הצלחתי לכתוב"), "install answers: " + rl.Title);
            Directory.Delete(main);
            Check(Core.DetectInstalled().Edition == 0, "nothing is recorded as installed after that failure");

            SelfTestReviewFixes(dir);
            SelfTestWordEditing(dir);

            Console.WriteLine(fails == 0 ? "SELFTEST PASSED" : "SELFTEST FAILED (" + fails + ")");
            return fails;
        }

        static FakeApp FakeWith(string dir, params string[] names)
        {
            FakeApp a = new FakeApp();
            foreach (string nm in names) a.Dicts.Items.Add(new FakeDict(a.Dicts, nm, dir));
            return a;
        }

        static string ActiveName(FakeApp a) { FakeDict d = a.Dicts.ActiveCustomDictionary; return d == null ? null : d.Name; }

        static string RegNumberOf(string path)
        {
            using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath))
            {
                if (k == null) return null;
                foreach (string vn in k.GetValueNames())
                {
                    string v = k.GetValue(vn) as string;
                    if (v != null && string.Equals(v, path, StringComparison.OrdinalIgnoreCase)) return vn;
                }
            }
            return null;
        }

        static bool RegHasValue(string name)
        {
            using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath))
                return k != null && k.GetValue(name) != null;
        }

        static void RegDelete(params string[] names)
        {
            using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath, true))
                if (k != null) foreach (string n in names) k.DeleteValue(n, false);
        }

        // 3.0.0.9: the person's own changes to the dictionary (add / remove / correct) and the suggestion sent to the dictionary's owner
        static void SelfTestWordEditing(string dir)
        {
            ProgressFn quiet = delegate (int p, string t) { };
            string main = Path.Combine(dir, Core.MainFile);
            string custom = Path.Combine(dir, "CUSTOM.DIC");
            if (Directory.Exists(MyWords.FolderOverride)) Directory.Delete(MyWords.FolderOverride, true);
            if (File.Exists(custom)) File.Delete(custom);

            Console.WriteLine("-- 3.0.0.9: what the person types is cleaned first");
            Check(WordEdits.Clean("  מיבעיא ") == "מיבעיא", "spaces around a word are dropped");
            Check(WordEdits.Clean("חז\"ל") == "חז\"ל" && WordEdits.Clean("ר׳") == "ר׳" && WordEdits.Clean("רמב״ם") == "רמב״ם", "gershayim and geresh stay");
            Check(WordEdits.Clean("בְּרֵאשִׁית") == "בראשית", "vowel points are dropped");
            Check(WordEdits.Clean("\u200Fשלום\u200E") == "שלום", "invisible direction marks are dropped");
            Check(WordEdits.Clean("שתי מילים") == null && WordEdits.Clean("") == null && WordEdits.Clean("א<ב") == null && WordEdits.Clean(new string('א', 80)) == null,
                  "two words, nothing, a sign or a very long one: not a word");

            Console.WriteLine("-- 3.0.0.9: adding, removing, correcting, undoing");
            List<string> edition = Core.BuildWords(false, false);
            WordEdits e = new WordEdits(edition, new List<WordChange>());
            int n0 = e.Count;
            Check(n0 == edition.Count, "the list starts as the edition (" + n0 + " words)");
            Check(e.Add("תיבהחדשה") == EditResult.Done && e.Contains("תיבהחדשה") && e.Count == n0 + 1, "a new word is added");
            Check(e.Add("תיבהחדשה") == EditResult.AlreadyThere && e.Add("מיבעי") == EditResult.AlreadyThere, "a word that is there already: said so");
            Check(e.Remove("מתרצינן") == EditResult.Done && !e.Contains("מתרצינן"), "a word of the dictionary is removed");
            Check(e.Remove("מתרצינן") == EditResult.NotThere, "removing it again: it is not there");
            Check(e.Fix("מיבעי", "מיבעיי") == EditResult.Done && !e.Contains("מיבעי") && e.Contains("מיבעיי"), "a correction: the old spelling goes, the new one comes");
            Check(e.Fix("מיבעיי", "מיבעיי") == EditResult.Same, "a 'correction' to the same spelling: nothing to do");
            Check(e.Changes.Count == 3 && e.Changes[0].Kind == '+' && e.Changes[1].Kind == '-' && e.Changes[2].Kind == '~', "three changes are listed, in order");
            Check(e.Remove("תיבהחדשה") == EditResult.Done && e.Changes.Count == 2 && !e.Contains("תיבהחדשה"), "removing a word of one's own takes back its addition");
            Check(e.Add("מתרצינן") == EditResult.Done && e.Changes.Count == 1 && e.Contains("מתרצינן"), "adding back a removed word takes back its removal");
            e.Undo(0);
            Check(e.Changes.Count == 0 && e.Contains("מיבעי") && !e.Contains("מיבעיי") && e.Count == n0, "undoing the correction: the dictionary is as it was");
            List<string> sorted = e.Sorted();
            bool inOrder = sorted.Count == n0;
            for (int i = 1; i < sorted.Count && inOrder; i++) if (string.CompareOrdinal(sorted[i - 1], sorted[i]) >= 0) inOrder = false;
            Check(inOrder, "the list of all the words is sorted, without repeats");
            e.Add("תיבהחדשה");
            Check(e.Sorted().Contains("תיבהחדשה") && e.Find("יבהחדש").Count >= 1, "the list and the search follow the changes");

            Console.WriteLine("-- 3.0.0.9: the changes are kept in a file, and every installation applies them");
            List<WordChange> mine = new List<WordChange>();
            mine.Add(new WordChange('+', "תיבהחדשה", null));
            mine.Add(new WordChange('-', "מתרצינן", null));
            mine.Add(new WordChange('~', "מיבעי", "מיבעיי"));
            mine[0].Sent = true;
            MyWords.Save(mine);
            List<WordChange> back = MyWords.Load();
            Check(back.Count == 3 && back[0].Sent && !back[1].Sent && back[2].Kind == '~' && back[2].Word == "מיבעי" && back[2].NewWord == "מיבעיי", "saved and read back, also what was already sent");
            Installer.Install(false, false, quiet, false);
            List<string> inFile = Core.ReadDicWords(main);
            Check(inFile.Contains("תיבהחדשה") && inFile.Contains("מיבעיי") && !inFile.Contains("מתרצינן") && !inFile.Contains("מיבעי"), "the installed file has the person's changes");
            Check(inFile.Count == edition.Count, "and nothing else (" + inFile.Count + " words: 2 added, 2 gone)");

            List<string> withWordsOwn = Core.ReadDicWords(main);
            withWordsOwn.Add("מילהשוורדהוסיף");   // a word that Word itself put into our file ("Add to Dictionary")
            Core.WriteDic(main, withWordsOwn);
            mine.RemoveAt(0);                     // the person takes back an addition
            MyWords.Save(mine);
            Installer.Install(false, false, quiet, false);
            inFile = Core.ReadDicWords(main);
            Check(!inFile.Contains("תיבהחדשה") && inFile.Contains("מילהשוורדהוסיף"), "an addition taken back goes; a word that Word added by itself stays");

            FakeApp hidden = new FakeApp();
            FakeApp fresh = new FakeApp(); fresh.InEffect = true;
            fresh.Known = new HashSet<string>(MyWords.Apply(Core.BuildWords(false, false), MyWords.Load()));
            Result r = InstallWithFakeWord(hidden, false, fresh);
            Check(r.Ok && r.Text.Contains("מזוהות ✓"), "the check after installing uses only words of the person's dictionary (" + r.Title + ")");

            Installer.Remove(false);
            List<string> saved = File.Exists(custom) ? Core.ReadDicWords(custom) : new List<string>();
            Check(MyWords.Load().Count == 2 && saved.Contains("מילהשוורדהוסיף") && !saved.Contains("מיבעיי"),
                  "removing the dictionary keeps the person's list; only words that Word added go to CUSTOM.DIC");
            if (File.Exists(custom)) File.Delete(custom);
            if (File.Exists(custom + ".bak")) File.Delete(custom + ".bak");
            MyWords.Save(new List<WordChange>());
            Check(MyWords.Load().Count == 0, "an empty list is saved as empty");
            File.WriteAllText(Path.Combine(MyWords.Folder(), MyWords.FileName),
                "# by hand\r\n+שתי מילים\r\n~מיבעי\r\n-<x>\r\n+  תיבהחדשה \tsent\r\n~מיבעי\tמיבעי\r\n", System.Text.Encoding.UTF8);
            List<WordChange> byHand = MyWords.Load();
            Check(byHand.Count == 1 && byHand[0].Word == "תיבהחדשה" && byHand[0].Sent, "a list changed by hand: only real words are read (" + byHand.Count + ")");
            MyWords.Save(new List<WordChange>());

            Console.WriteLine("-- 3.0.0.9: the suggestion that goes to the dictionary's owner");
            List<WordChange> s = new List<WordChange>();
            s.Add(new WordChange('+', "ומתרצינן", null));
            s.Add(new WordChange('-', "בראשיתת", null));
            s.Add(new WordChange('~', "חזל", "חז\"ל"));
            string title = Suggest.Title(s);
            string body = Suggest.Body(s, "ברכות דף ב", "מורחב");
            Check(title.Contains("1 להוסיף") && title.Contains("1 להסיר") && title.Contains("1 לתקן"), "the title counts the changes");
            Check(body.Contains("ומתרצינן") && body.Contains("בראשיתת") && body.Contains("חזל ← חז\"ל") && body.Contains("ברכות דף ב") && body.Contains(Core.AppVersion) && body.Contains("מורחב"),
                  "the text has every change, the note, the version and the edition");
            Check(Suggest.Json("t", "a \"b\" \\ c\nd") == "{\"title\":\"t\",\"body\":\"a \\\"b\\\" \\\\ c\\nd\"}", "JSON: quotes, backslash and new line are escaped");
            string url = Suggest.BrowserUrl(title, body);
            Check(url.StartsWith("https://github.com/elicohen8565-hue/milon-torani-3.0/issues/new?title=") && url.Contains("&body=") && url.IndexOf(' ') < 0, "the page to open in the browser");
            Check(Suggest.BrowserUrl(title, new string('א', 20000)).Length <= 8000, "a very long suggestion still fits the address of the page");
            byte[] hiddenKey = Suggest.Encode("github_pat_TEST_123");
            Check(Suggest.Decode(hiddenKey) == "github_pat_TEST_123" && System.Text.Encoding.UTF8.GetString(hiddenKey).IndexOf("github_pat") < 0, "the key is kept hidden inside the program, and read back");
            string tk = Suggest.Token();
            Check(tk == null || (tk.Length > 20 && tk.IndexOf(' ') < 0), "the key inside this program (" + (tk == null ? "none" : "there") + ") reads back whole");
        }

        // 3.0.0.8: the findings of the code review (3-11-2026), each one first shown failing on 3.0.0.7
        static void SelfTestReviewFixes(string dir)
        {
            ProgressFn quiet = delegate (int p, string t) { };
            string main = Path.Combine(dir, Core.MainFile);
            string custom = Path.Combine(dir, "CUSTOM.DIC");
            HashSet<string> basicWords = new HashSet<string>(Core.BuildWords(false, false));

            Console.WriteLine("-- 3.0.0.8: our file must never be Word's default dictionary (the person's own words would go into it)");
            if (File.Exists(custom)) File.Delete(custom);
            FakeApp empty = new FakeApp();               // Word lists no custom dictionary at all
            FakeApp check1 = new FakeApp(); check1.InEffect = true; check1.Known = basicWords;
            InstallWithFakeWord(empty, false, check1);
            Check(ActiveName(empty) != null && !Core.IsOurs(ActiveName(empty)), "only ours in Word -> another dictionary is made the default (" + ActiveName(empty) + ")");
            Check(File.Exists(custom) && empty.Dicts.Items.Exists(delegate (FakeDict x) { return string.Equals(x.Name, "CUSTOM.DIC", StringComparison.OrdinalIgnoreCase); }), "Word's usual CUSTOM.DIC was created and added");
            FakeApp roam = FakeWith(dir, "RoamingCustom.dic");
            roam.Dicts.Active = roam.Dicts.Items[0];
            FakeApp check2 = new FakeApp(); check2.InEffect = true; check2.Known = basicWords;
            InstallWithFakeWord(roam, false, check2);
            Check(ActiveName(roam) == "RoamingCustom.dic" && roam.Dicts.Count == 2, "the person's default stays the default, nothing else is added");
            FakeApp oursFirst = FakeWith(dir, Core.MainFile, "RoamingCustom.dic");   // ours was the default (the first of the list)
            oursFirst.Dicts.Active = oursFirst.Dicts.Items[0];
            InstallWithFakeWord(oursFirst, true, null);
            Check(ActiveName(oursFirst) == "RoamingCustom.dic", "ours was the default -> the person's dictionary becomes the default (" + ActiveName(oursFirst) + ")");

            Console.WriteLine("-- 3.0.0.8: words the person added into our file are kept (update) and moved to CUSTOM.DIC (removal)");
            List<string> withTheirs = Core.BuildWords(false, false);
            withTheirs.Add("מילהשלי"); withTheirs.Add("עודמילהשלי");
            Core.WriteDic(main, withTheirs);
            Installer.Install(true, false, quiet, false);
            List<string> after = Core.ReadDicWords(main);
            Check(after.Contains("מילהשלי") && after.Contains("עודמילהשלי") && after.Count == Core.BuildWords(true, false).Count + 2, "after an update the two words are still in the file (" + after.Count + " words)");
            if (File.Exists(custom)) File.Delete(custom);
            Result rm = Installer.Remove(false);
            List<string> saved = File.Exists(custom) ? Core.ReadDicWords(custom) : new List<string>();
            Check(!File.Exists(main) && saved.Contains("מילהשלי") && saved.Contains("עודמילהשלי") && saved.Count == 2, "after a removal they are in CUSTOM.DIC (" + saved.Count + " words)");
            Check(rm.Ok && rm.Text.Contains("CUSTOM.DIC"), "the removal says where they went");
            if (File.Exists(custom)) File.Delete(custom);

            Console.WriteLine("-- 3.0.0.8: the new file cannot be written: Word keeps the dictionary it had");
            Installer.Install(false, false, quiet, false);              // a working basic installation
            byte[] before = File.ReadAllBytes(main);
            FakeApp w1 = FakeWith(dir, Core.MainFile, "RoamingCustom.dic");
            Core.FailWriteHook = delegate (string p) { return string.Equals(p, main, StringComparison.OrdinalIgnoreCase); };
            Core.FailCopyHook = Core.FailWriteHook;   // held by another program: neither writing nor copying over it works
            Result f1;
            try { f1 = InstallWithFakeWord(w1, true, null); } finally { Core.FailWriteHook = null; Core.FailCopyHook = null; }
            Check(!f1.Ok && f1.Title.Contains("לא הצלחתי לכתוב"), "says the file could not be written (" + f1.Title + ")");
            Check(w1.Dicts.Items.FindAll(delegate (FakeDict x) { return x.Name == Core.MainFile; }).Count == 1, "our entry is in Word again (given back)");
            Check(SameBytes(File.ReadAllBytes(main), before) && Core.DetectInstalled().Edition == 1, "the old file and the record are untouched");
            Installer.Install(false, false, quiet, false);
            Check(!File.Exists(main + ".tmp"), "the next installation leaves no temporary file");
            FakeApp w2 = FakeWith(dir, Core.MainFile, "RoamingCustom.dic");
            Core.FailWriteHook = delegate (string p) { return p.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase); };
            Result f2;
            try { f2 = InstallWithFakeWord(w2, true, null); } finally { Core.FailWriteHook = null; }
            Check(!f2.Ok && w2.Deleted == 0 && w2.Added.Count == 0, "the folder cannot be written at all: Word was not touched");

            Console.WriteLine("-- 3.0.0.10: Word's own list of dictionaries (registry, one entry per file and account) is full");
            {
                string acct = "16_7bd116fa7f21d40b_LiveId";
                string roamPath = Path.Combine(Path.Combine(dir, "Proofing"), "RoamingCustom.dic");
                Directory.CreateDirectory(Path.GetDirectoryName(roamPath));
                Core.WriteDic(roamPath, new List<string> { "מילהשלי" });
                Action<string[]> setList = delegate (string[] files)
                {
                    Microsoft.Win32.Registry.CurrentUser.DeleteSubKeyTree(RegFallback.KeyPath, false);
                    using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegFallback.KeyPath))
                    {
                        k.SetValue("UpdateComplete", 1, Microsoft.Win32.RegistryValueKind.DWord);
                        for (int i = 0; i < files.Length; i++)
                        {
                            string h = (i + 1).ToString(CultureInfo.InvariantCulture) + "_" + acct;
                            k.SetValue(h, files[i], Microsoft.Win32.RegistryValueKind.String);
                            k.SetValue(h + "_state", new byte[] { 0, 0, 0, 0 }, Microsoft.Win32.RegistryValueKind.Binary);
                            k.SetValue(h + "_roamed", new byte[] { 0, 0, 0, 0 }, Microsoft.Win32.RegistryValueKind.Binary);
                            k.SetValue(h + "_external", new byte[] { 0, 0, 0, 0 }, Microsoft.Win32.RegistryValueKind.Binary);
                        }
                    }
                };
                Func<List<string>> wordList = delegate
                {
                    List<string> l = new List<string>();
                    using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath))
                        foreach (DictEntry e in DictList.Read(k)) if (e.Main != null && e.Identity == acct) l.Add(e.Number + ":" + e.Leaf);
                    return l;
                };
                Func<bool> full = delegate { return wordList().Count >= 10; };
                int n0; string id0; bool m0;
                Check(DictList.Parse("4_" + acct, out n0, out id0, out m0) && n0 == 4 && id0 == acct && m0
                      && DictList.Parse("10_" + acct + "_roamed", out n0, out id0, out m0) && n0 == 10 && !m0
                      && DictList.Parse("1_state", out n0, out id0, out m0) && id0 == "" && !m0
                      && !DictList.Parse("UpdateComplete", out n0, out id0, out m0), "the names of Word's list are read");

                // a. as on the user's computer (9-11): files of 3.0.0.1, a WordCheck trial and five size-check files fill the list
                setList(new string[] { roamPath, "Otzaria_Torah_Basic.dic", "Otzaria_Torah_Acronyms.dic", Core.MainFile, Path.Combine(dir, "wordcheck_trial.dic"),
                    "Milon_SizeTest_175000.dic", "Milon_SizeTest_200000.dic", "Milon_SizeTest_250000.dic", "Milon_SizeTest_300000.dic", "Milon_SizeTest_400000.dic" });
                FakeApp fullWord = new FakeApp(); fullWord.Dicts.FailAddWhen = full;
                FakeApp checkF = new FakeApp(); checkF.InEffect = true; checkF.Known = basicWords;
                Result fa = InstallWithFakeWord(fullWord, false, checkF);
                Check(fa.Ok && fullWord.Added.Contains(main), "before Word starts the old entries go, and Word takes the dictionary (" + fa.Title + ")");
                Check(string.Join(" ", wordList().ToArray()) == "1:RoamingCustom.dic 2:Milon_Torani.dic", "the list: the person's dictionary, ours moved up to 2 (" + string.Join(" ", wordList().ToArray()) + ")");
                Check(fullWord.Added.Contains(roamPath), "the person's Office dictionary, not in use, is put back into use");
                // (3.1 review) it is put back before the default dictionary is chosen: it is the default, not a new empty CUSTOM.DIC
                Check(fullWord.Dicts.ActiveCustomDictionary != null && fullWord.Dicts.ActiveCustomDictionary.Name == "RoamingCustom.dic"
                      && !fullWord.Added.Exists(delegate (string p) { return string.Equals(Path.GetFileName(p), "CUSTOM.DIC", StringComparison.OrdinalIgnoreCase); }),
                      "Word's default dictionary is the person's Office dictionary (" + (fullWord.Dicts.ActiveCustomDictionary == null ? "none" : fullWord.Dicts.ActiveCustomDictionary.Name) + ")");

                // b. the list is full of the person's own dictionaries and ours: ours goes too, and a new hidden Word takes the dictionary
                setList(new string[] { roamPath, "Own1.dic", "Own2.dic", "Own3.dic", "Own4.dic", "Own5.dic", "Own6.dic", "Own7.dic", "Own8.dic", Core.MainFile });
                FakeApp fullWord2 = new FakeApp(); fullWord2.Dicts.FailAddWhen = full;
                FakeApp freshF = new FakeApp(); freshF.Dicts.FailAddWhen = full; freshF.InEffect = true; freshF.Known = basicWords;
                Result fb = InstallWithFakeWord(fullWord2, false, freshF);
                Check(fb.Ok && freshF.Added.Contains(main) && fb.Text.Contains("מזוהות ✓"), "Word refused, the list was cleaned, a new Word took it (" + fb.Title + ")");
                List<string> left = wordList();
                Check(left.Count == 9 && left.Contains("2:Own1.dic") && left.Contains("9:Own8.dic"), "the person's eight dictionaries stay (" + string.Join(" ", left.ToArray()) + ")");

                // c. the person's own Word is open and refuses: our entry is not taken out of the list (that Word would write its old copy
                //    back). (Whether a Word runs is asked of Windows; here none does, so the old files' entries still go before the start.)
                setList(new string[] { roamPath, "Otzaria_Torah_Basic.dic", Core.MainFile });
                FakeApp open = new FakeApp(); open.Dicts.FailAdd = true;
                InstallWithFakeWord(open, true, null);
                Check(string.Join(" ", wordList().ToArray()) == "1:RoamingCustom.dic 2:Milon_Torani.dic", "Word was open: our entry stays in the list (" + string.Join(" ", wordList().ToArray()) + ")");

                // d. (3.1 review) the list exactly as on the user's computer, with what is outside the account too: Office's "1_state" and
                //    "1_external", and "11" = our file (an older version's registry route). The clean must not move "11" to "1", and
                //    "remove from Word" must not take the person's "1_<account>" (RoamingCustom.dic) or Office's settings with it.
                string[] nineEleven = { roamPath, "Otzaria_Torah_Basic.dic", "Otzaria_Torah_Acronyms.dic", Core.MainFile, Path.Combine(dir, "wordcheck_trial.dic"),
                    "Milon_SizeTest_175000.dic", "Milon_SizeTest_200000.dic", "Milon_SizeTest_250000.dic", "Milon_SizeTest_300000.dic", "Milon_SizeTest_400000.dic" };
                setList(nineEleven);
                using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath, true))
                {
                    k.SetValue("1_state", new byte[] { 1, 0, 0, 0 }, Microsoft.Win32.RegistryValueKind.Binary);
                    k.SetValue("1_external", new byte[] { 0, 0, 0, 0 }, Microsoft.Win32.RegistryValueKind.Binary);
                    k.SetValue("11", main, Microsoft.Win32.RegistryValueKind.String);
                }
                Func<string, object> val = delegate (string name)
                {
                    using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath)) return k == null ? null : k.GetValue(name);
                };
                FakeApp fullWord4 = new FakeApp(); fullWord4.Dicts.FailAddWhen = full;
                FakeApp checkD = new FakeApp(); checkD.InEffect = true; checkD.Known = basicWords;
                Result fd = InstallWithFakeWord(fullWord4, false, checkD);
                Check(fd.Ok && val("11") != null && val("1") == null, "the entry without an account is not moved (" + fd.Title + ")");
                Installer.Remove(false);
                Check(roamPath.Equals(val("1_" + acct)) && val("1_" + acct + "_state") != null && val("1_" + acct + "_roamed") != null && val("1_" + acct + "_external") != null,
                      "remove from Word: the person's Office dictionary stays in Word's list, with its settings");
                Check(val("1_state") != null && val("1_external") != null && val("11") == null, "remove from Word: our old entry goes, Office's own settings stay");

                // e. (3.1 review) a Word still runs (also our own one that is still closing): Word's list is not changed at all - that Word
                //    would write its old copy back
                setList(nineEleven);
                int keepWait = WordCom.CloseWaitMs;
                WordCom.AnyWordHook = delegate { return true; };
                WordCom.CloseWaitMs = 300;
                try
                {
                    FakeApp fullWord5 = new FakeApp(); fullWord5.Dicts.FailAddWhen = full;
                    FakeApp checkE = new FakeApp(); checkE.Dicts.FailAddWhen = full; checkE.InEffect = true; checkE.Known = basicWords;
                    InstallWithFakeWord(fullWord5, false, checkE);
                }
                finally { WordCom.AnyWordHook = null; WordCom.CloseWaitMs = keepWait; }
                Check(wordList().Count == 10 && wordList().Contains("6:Milon_SizeTest_175000.dic"), "a Word still runs: Word's list was not changed (" + wordList().Count + " entries)");

                // f. (3.1 review) closing up the numbers never overwrites: settings that Office left without a file keep their number
                setList(new string[] { roamPath, "Otzaria_Torah_Basic.dic", "Own1.dic" });
                using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath, true))
                {
                    k.DeleteValue("2_" + acct, false);                                                  // entry 2: only its settings are left
                    k.SetValue("4_" + acct, "Otzaria_Torah_Basic.dic", Microsoft.Win32.RegistryValueKind.String);   // a stale file at 4
                    List<string> doneF = DictList.Clean(k, false);
                    Check(doneF.Count == 1 && doneF[0].StartsWith("deleted 4"), "only the stale entry goes; nothing moves onto the settings left at 2 (" + string.Join("; ", doneF.ToArray()) + ")");
                    Check(k.GetValue("3_" + acct) as string == "Own1.dic" && k.GetValue("2_" + acct + "_state") != null, "the person's entry 3 and the settings at 2 are where they were");
                }
                Microsoft.Win32.Registry.CurrentUser.DeleteSubKeyTree(RegFallback.KeyPath, false);
            }

            Console.WriteLine("-- 3.0.0.8: Word refuses to add the dictionary: the registry route is checked, not just assumed");
            FakeApp refuseAdd = new FakeApp(); refuseAdd.Dicts.FailAdd = true;
            FakeApp check3 = new FakeApp(); check3.InEffect = true; check3.Known = basicWords;
            check3.Dicts.FailAdd = true;   // 3.0.0.10: a new hidden Word is asked again after the list is cleaned - this one refuses too
            Result g1 = InstallWithFakeWord(refuseAdd, false, check3);
            Check(RegNumberOf(main) != null, "our file is in Word's registry list");
            Check(g1.Ok && g1.Title == "ההתקנה הושלמה" && g1.Text.Contains("מזוהות ✓"), "a new Word knows the words: done (" + g1.Title + ")");
            FakeApp refuseAdd2 = new FakeApp(); refuseAdd2.Dicts.FailAdd = true;
            Result g2 = InstallWithFakeWord(refuseAdd2, true, null);   // Word was open: it cannot be checked
            Check(!g2.Ok && g2.Text.Contains("התקנה ידנית"), "cannot be checked: not reported as done (" + g2.Title + ")");
            RegFallback.RemoveOurs();
            using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegFallback.KeyPath))
            {
                k.SetValue("7_state", 0, Microsoft.Win32.RegistryValueKind.DWord);            // settings left over from an entry that is gone
                k.SetValue("7_culturetag", "en-US", Microsoft.Win32.RegistryValueKind.String);
            }
            RegFallback.Add(main);
            string num = RegNumberOf(main);
            int numValue;
            Check(num != null && int.TryParse(num, out numValue) && numValue > 7, "a new registry entry never takes the number of left-over settings (" + num + ")");
            if (num != null) using (Microsoft.Win32.RegistryKey k = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RegFallback.KeyPath)) k.SetValue(num + "_state", 1, Microsoft.Win32.RegistryValueKind.DWord);
            RegFallback.RemoveOurs();
            Check(num != null && !RegHasValue(num) && !RegHasValue(num + "_state"), "removing our registry entry removes its settings too");
            RegDelete("7_state", "7_culturetag");

            Console.WriteLine("-- 3.0.0.8: a file that cannot be deleted: the removal says so");
            Installer.Install(false, false, quiet, false);
            Core.FailDeleteHook = delegate (string p) { return string.Equals(Path.GetFileName(p), Core.MainFile, StringComparison.OrdinalIgnoreCase); };
            Result h1;
            try { h1 = Installer.Remove(false); } finally { Core.FailDeleteHook = null; }
            Check(!h1.Ok && File.Exists(main), "not reported as removed (" + h1.Title + ")");
            InstalledInfo hi = Core.DetectInstalled();
            Check(hi.Edition == 1 && hi.Version == Core.AppVersion, "still shown as installed (basic, " + hi.Version + "), not as 'an old version'");
            Result h2 = Installer.Remove(false);
            Check(h2.Ok && !File.Exists(main), "the second try removes it");

            Console.WriteLine("-- 3.0.0.8: advanced removal tells the truth");
            Installer.Install(false, false, quiet, false);
            FakeApp adv = FakeWith(dir, Core.MainFile, "RoamingCustom.dic");
            adv.Refuse.Add("RoamingCustom.dic");
            WordCom.OpenHook = delegate { WordSession s = new WordSession(); s.App = adv; s.Created = true; return s; };
            Result ar;
            try { ar = Installer.RemoveAdvanced(true); } finally { WordCom.OpenHook = null; }
            Check(!ar.Ok && ar.Title.Contains("חלקית") && ar.Text.Contains("RoamingCustom.dic"), "Word kept a dictionary: 'partly done', and it is named (" + ar.Title + ")");
            bool listed = false;
            foreach (string b in Directory.GetDirectories(dir, "Backup_*"))
            {
                string note = Path.Combine(b, "dictionaries.txt");
                if (File.Exists(note) && File.ReadAllText(note).Contains("RoamingCustom.dic")) listed = true;
            }
            Check(listed, "the backup says where every dictionary came from (dictionaries.txt)");
            foreach (string f in Directory.GetFiles(dir)) if (f.EndsWith(".dic", StringComparison.OrdinalIgnoreCase)) File.Delete(f);
            RegFallback.Add(Path.Combine(dir, "Other.dic"));   // only an entry in the registry list is left
            Result ar2 = Installer.RemoveAdvanced(false);
            Check(ar2.Ok && !ar2.Text.Contains("לא נמצאו") && ar2.Text.Contains("Backup_"), "a registry entry alone also counts as removed, and the backup is named");

            Console.WriteLine("-- 3.0.0.8: Word's other 'busy' answer");
            Check(WordCom.IsBusy(new COMException("busy", unchecked((int)0x800AC472))), "0x800AC472 is 'busy': wait and ask again");

            Console.WriteLine("-- 3.0.0.8, second look (the check of the fixes)");
            // b. writing over the file fails, but copying the complete new file over it works
            Installer.Install(false, false, quiet, false);
            byte[] good = File.ReadAllBytes(main);
            FakeApp w3 = FakeWith(dir, Core.MainFile, "RoamingCustom.dic");
            Core.FailWriteHook = delegate (string p) { return string.Equals(p, main, StringComparison.OrdinalIgnoreCase); };
            Result f3;
            try { f3 = InstallWithFakeWord(w3, true, null); } finally { Core.FailWriteHook = null; }
            Check(f3.Ok && SameBytes(File.ReadAllBytes(main), good) && !File.Exists(main + ".tmp"), "the first write fails, the second try (a copy of the new file) works: installed (" + f3.Title + ")");

            // c. the person's words cannot be saved anywhere else: nothing at all is removed, not even from Word
            List<string> mixed = Core.BuildWords(false, false);
            mixed.Add("מילהשלי");
            Core.WriteDic(main, mixed);
            FakeApp w4 = FakeWith(dir, Core.MainFile, "RoamingCustom.dic");
            WordCom.OpenHook = delegate { WordSession s = new WordSession(); s.App = w4; s.Created = true; return s; };
            Core.FailWriteHook = delegate (string p) { return string.Equals(Path.GetFileName(p), "CUSTOM.DIC", StringComparison.OrdinalIgnoreCase); };
            Result h3;
            try { h3 = Installer.Remove(true); } finally { Core.FailWriteHook = null; WordCom.OpenHook = null; }
            Check(!h3.Ok && File.Exists(main) && w4.Deleted == 0, "their words cannot be saved elsewhere: nothing is removed, not even from Word (" + h3.Title + ")");

            // d. the person's CUSTOM.DIC is copied before their words are added to it
            Core.WriteDic(custom, new List<string> { "שלהםקודם" });
            Installer.Remove(false);
            Check(File.Exists(custom + ".bak") && Core.ReadDicWords(custom + ".bak").Count == 1 && Core.ReadDicWords(custom).Contains("מילהשלי"), "CUSTOM.DIC is copied (CUSTOM.DIC.bak) before it is changed");
            if (File.Exists(custom)) File.Delete(custom);
            if (File.Exists(custom + ".bak")) File.Delete(custom + ".bak");

            // e. a CUSTOM.DIC from another folder in Word's list is not the one the words were moved to
            Installer.Install(false, false, quiet, false);
            List<string> mixed2 = Core.ReadDicWords(main);
            mixed2.Add("מילהשלי");
            Core.WriteDic(main, mixed2);
            FakeApp w5 = new FakeApp();
            w5.Dicts.Items.Add(new FakeDict(w5.Dicts, "CUSTOM.DIC", Path.Combine(dir, "elsewhere")));
            w5.Dicts.Items.Add(new FakeDict(w5.Dicts, Core.MainFile, dir));
            WordCom.OpenHook = delegate { WordSession s = new WordSession(); s.App = w5; s.Created = true; return s; };
            try { Installer.Remove(true); } finally { WordCom.OpenHook = null; }
            Check(w5.Added.Contains(custom), "a CUSTOM.DIC from another folder is in Word's list: the one in the UProof folder is still added");
            if (File.Exists(custom)) File.Delete(custom);
            if (File.Exists(custom + ".bak")) File.Delete(custom + ".bak");

            // f. registry route: our hidden Word is closed BEFORE the registry entry is written (Word may write its own list when it closes)
            RegFallback.RemoveOurs();
            FakeApp refuseAdd3 = new FakeApp(); refuseAdd3.Dicts.FailAdd = true;
            refuseAdd3.Probe = delegate { return RegNumberOf(main) == null ? "no entry yet" : "entry already written"; };
            FakeApp check4 = new FakeApp(); check4.InEffect = true; check4.Known = basicWords;
            InstallWithFakeWord(refuseAdd3, false, check4);
            Check(refuseAdd3.AtQuit == "no entry yet", "the hidden Word was closed before the registry entry was written (" + refuseAdd3.AtQuit + ")");
            RegFallback.RemoveOurs();

            // g. advanced removal: Word refuses everything and there is nothing else: no 'nothing found', and the backup is named
            RegFallback.RemoveAllDictionaries();
            foreach (string f in Directory.GetFiles(dir)) if (f.EndsWith(".dic", StringComparison.OrdinalIgnoreCase)) File.Delete(f);
            string elsewhere = Path.Combine(dir, "elsewhere");
            Directory.CreateDirectory(elsewhere);
            Core.WriteDic(Path.Combine(elsewhere, "RoamingCustom.dic"), new List<string> { "אחת" });
            FakeApp w6 = new FakeApp();
            w6.Dicts.Items.Add(new FakeDict(w6.Dicts, "RoamingCustom.dic", elsewhere));
            w6.Refuse.Add("RoamingCustom.dic");
            WordCom.OpenHook = delegate { WordSession s = new WordSession(); s.App = w6; s.Created = true; return s; };
            Result ar3;
            try { ar3 = Installer.RemoveAdvanced(true); } finally { WordCom.OpenHook = null; }
            Check(!ar3.Ok && !ar3.Text.Contains("לא נמצאו") && ar3.Text.Contains("Backup_") && ar3.Text.Contains("RoamingCustom.dic"), "Word refused everything: the answer says so and names the backup");
            Directory.Delete(elsewhere, true);
        }
    }
}
