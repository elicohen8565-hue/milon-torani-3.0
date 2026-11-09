using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Threading;
using System.Windows.Forms;

// The word-editing window (stage 4, version 3.0.0.9): the person searches the dictionary, corrects the spelling of a word, removes a word or
// adds one. The changes are the person's own (Core.cs: MyWords, WordEdits): every installation applies them, also of a new version.
// After saving they can be sent to the dictionary's owner as a suggestion (Core.cs: Suggest), which becomes an issue on GitHub.
namespace OtzariaDict
{
    public delegate void IndexFn(int index);

    // ------------------------------------------------------------------------------------------ a text box in the look of the window
    // A plain Windows text box without a border of its own, inside a frame drawn like the rest of the window (blue while it has the focus).
    // Cue = grey text shown while the box is empty and the focus is elsewhere. Side = small grey text on the left, inside the frame.
    class FlatTextBox : BaseControl
    {
        public readonly TextBox Box = new TextBox();
        string cue = "", side = "";
        bool showingCue, settingText;
        public int SideUnits;                // room for Side on the left, in design units
        public event EventHandler Changed;   // the person changed the text (not when the cue comes or goes)
        public event EventHandler Submit;    // Enter

        public FlatTextBox(float fontSize)
        {
            BackColor = Theme.Card;
            Cursor = Cursors.IBeam;
            Box.BorderStyle = BorderStyle.None;
            Box.RightToLeft = RightToLeft.Yes;
            Box.Font = Theme.F(fontSize, false);
            Box.BackColor = Theme.Card;
            Box.ForeColor = Theme.Text;
            Box.MaxLength = 200;
            Box.TextChanged += delegate { if (!settingText && Changed != null) Changed(this, EventArgs.Empty); };
            Box.GotFocus += delegate { if (showingCue) Put("", false); Invalidate(); };
            Box.LostFocus += delegate { if (Box.Text.Length == 0) Put(cue, true); Invalidate(); };
            Box.KeyDown += delegate (object s, KeyEventArgs e)
            {
                if (e.Control && e.KeyCode == Keys.A) { Box.SelectAll(); e.SuppressKeyPress = true; e.Handled = true; return; }   // (not built into every Windows text box)
                if (e.KeyCode != Keys.Enter) return;
                e.SuppressKeyPress = true;   // no "ding"
                e.Handled = true;
                if (Submit != null) Submit(this, EventArgs.Empty);
            };
            Controls.Add(Box);
        }

        void Put(string text, bool isCue)
        {
            settingText = true;
            try
            {
                showingCue = isCue && text.Length > 0;
                Box.ForeColor = showingCue ? Theme.TextQuiet : Theme.Text;
                Box.Text = text;
            }
            finally { settingText = false; }
        }

        public string Cue
        {
            get { return cue; }
            set { cue = value ?? ""; if ((showingCue || Box.Text.Length == 0) && !Box.Focused) Put(cue, true); }
        }

        // the text the person typed ("" while the cue is shown)
        public string Value
        {
            get { return showingCue ? "" : Box.Text; }
            set
            {
                string v = value ?? "";
                if (v.Length == 0 && !Box.Focused) Put(cue, true);
                else Put(v, false);
                if (Changed != null) Changed(this, EventArgs.Empty);
            }
        }

        public string Side { get { return side; } set { side = value ?? ""; Invalidate(); } }

        public void FocusBox() { Box.Focus(); }
        protected override void OnMouseDown(MouseEventArgs e) { Box.Focus(); base.OnMouseDown(e); }
        protected override void OnEnabledChanged(EventArgs e) { Box.Enabled = Enabled; Invalidate(); base.OnEnabledChanged(e); }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            int pad = Dpi.S(10);
            int left = pad + (SideUnits > 0 ? Dpi.S(SideUnits) : 0);
            Box.Bounds = new Rectangle(left, (Height - Box.Height) / 2, Math.Max(10, Width - left - pad), Box.Height);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Gfx.Rect(g, ClientRectangle, Enabled ? Theme.Card : Theme.DisabledBack);
            bool focused = Box.Focused;
            Gfx.Frame(g, ClientRectangle, focused ? Theme.Primary : (Enabled ? Theme.Stroke : Theme.DisabledLine), focused ? Math.Max(2, Dpi.S(2)) : Dpi.Hair);
            if (side.Length > 0 && SideUnits > 0)
                Gfx.TextAt(g, side, Theme.F(13, false), Theme.TextQuiet, new Rectangle(Dpi.S(10), 0, Dpi.S(SideUnits), Height), Align.Left, true, false);
        }
    }

    // ------------------------------------------------------------------------------------------ a list that scrolls (drawn by hand)
    // Rows of one height, a frame, a scroll bar on the left (the reading side ends there), the mouse wheel and the keyboard.
    abstract class ScrollList : BaseControl
    {
        protected readonly int RowUnits;
        protected int top, sel = -1, hover = -1;
        protected readonly VScrollBar bar = new VScrollBar();
        public string EmptyText = "";
        public event EventHandler SelectionChanged;

        protected ScrollList(int rowUnits)
        {
            RowUnits = rowUnits;
            TabStop = true;
            BackColor = Theme.Card;
            bar.Visible = false;
            bar.TabStop = false;
            bar.Scroll += delegate { top = bar.Value; hover = -1; Invalidate(); };
            Controls.Add(bar);
        }

        protected abstract int Count { get; }
        protected abstract void DrawRow(Graphics g, int index, Rectangle r, bool selected, bool hovered);

        protected virtual bool Locked { get { return false; } }   // (while a row is being edited: nothing moves)
        public int SelectedIndex { get { return sel; } }
        protected int RowPx { get { return Math.Max(1, Dpi.S(RowUnits)); } }
        protected Rectangle Inner
        {
            get
            {
                int h = Dpi.Hair, bw = bar.Visible ? bar.Width : 0;
                return new Rectangle(h + bw, h, Math.Max(0, Width - 2 * h - bw), Math.Max(0, Height - 2 * h));
            }
        }
        protected int FullRows { get { return Math.Max(1, Inner.Height / RowPx); } }
        int MaxTop { get { return Math.Max(0, Count - FullRows); } }

        protected Rectangle RowRect(int i)
        {
            Rectangle inn = Inner;
            return new Rectangle(inn.X, inn.Y + (i - top) * RowPx, inn.Width, RowPx);
        }

        protected int RowAt(Point p)
        {
            Rectangle inn = Inner;
            if (!inn.Contains(p)) return -1;
            int i = top + (p.Y - inn.Y) / RowPx;
            return i < Count ? i : -1;
        }

        // after the rows changed: keepSel = the row to select (-1 = none, and back to the top)
        protected void Reset(int keepSel)
        {
            sel = keepSel >= 0 && keepSel < Count ? keepSel : -1;
            hover = -1;
            if (sel < 0) top = 0;
            UpdateBar();
            if (sel >= 0) EnsureVisible(sel);
            Invalidate();
            if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty);
        }

        void UpdateBar()
        {
            bool need = Count > FullRows;
            int hair = Dpi.Hair;
            bar.Bounds = new Rectangle(hair, hair, SystemInformation.VerticalScrollBarWidth, Math.Max(0, Height - 2 * hair));
            bar.Visible = need;
            top = Math.Max(0, Math.Min(top, MaxTop));
            if (need)
            {
                bar.Value = 0;
                bar.Minimum = 0;
                bar.Maximum = Count - 1;
                bar.LargeChange = FullRows;
                bar.SmallChange = 1;
                bar.Value = top;
            }
        }

        protected override void OnResize(EventArgs e) { base.OnResize(e); UpdateBar(); }

        public void EnsureVisible(int i)
        {
            if (i < 0 || i >= Count) return;
            if (i < top) top = i;
            else if (i >= top + FullRows) top = i - FullRows + 1;
            top = Math.Max(0, Math.Min(top, MaxTop));
            if (bar.Visible) bar.Value = top;
            Invalidate();
        }

        public void Select(int i)
        {
            if (Count == 0) return;
            i = Math.Max(0, Math.Min(Count - 1, i));
            if (i != sel) { sel = i; if (SelectionChanged != null) SelectionChanged(this, EventArgs.Empty); }
            EnsureVisible(i);
            Invalidate();
        }

        void ScrollBy(int rows)
        {
            if (!bar.Visible || Locked) return;
            top = Math.Max(0, Math.Min(MaxTop, top + rows));
            bar.Value = top;
            hover = -1;
            Invalidate();
        }

        protected override void OnMouseWheel(MouseEventArgs e) { ScrollBy(-e.Delta / 120 * 3); base.OnMouseWheel(e); }
        protected override void OnMouseMove(MouseEventArgs e)
        {
            int h = Locked ? -1 : RowAt(e.Location);
            if (h != hover) { hover = h; Invalidate(); }
            base.OnMouseMove(e);
        }
        protected override void OnMouseLeave(EventArgs e) { if (hover != -1) { hover = -1; Invalidate(); } base.OnMouseLeave(e); }
        protected override void OnMouseDown(MouseEventArgs e)
        {
            if (Locked) return;
            Focus();
            int i = RowAt(e.Location);
            if (i >= 0) Select(i);
            base.OnMouseDown(e);
        }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }

        protected override bool IsInputKey(Keys k)
        {
            Keys key = k & Keys.KeyCode;
            if (key == Keys.Up || key == Keys.Down || key == Keys.PageUp || key == Keys.PageDown || key == Keys.Home || key == Keys.End
                || key == Keys.Enter || key == Keys.Delete || key == Keys.F2) return true;
            return base.IsInputKey(k);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (!Locked && Count > 0)
            {
                int page = Math.Max(1, FullRows - 1);
                switch (e.KeyCode)
                {
                    case Keys.Up: Select(sel < 0 ? 0 : sel - 1); e.Handled = true; break;
                    case Keys.Down: Select(sel < 0 ? 0 : sel + 1); e.Handled = true; break;
                    case Keys.PageUp: Select(sel < 0 ? 0 : sel - page); e.Handled = true; break;
                    case Keys.PageDown: Select(sel < 0 ? 0 : sel + page); e.Handled = true; break;
                    case Keys.Home: Select(0); e.Handled = true; break;
                    case Keys.End: Select(Count - 1); e.Handled = true; break;
                }
            }
            base.OnKeyDown(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            Gfx.Rect(g, ClientRectangle, Theme.Card);
            Rectangle inn = Inner;
            if (Count == 0)
            {
                if (EmptyText.Length > 0)
                    Gfx.TextAt(g, EmptyText, Theme.F(14, false), Theme.TextQuiet, Rectangle.Inflate(inn, -Dpi.S(20), -Dpi.S(10)), Align.Center, true, true);
            }
            else
            {
                GraphicsState st = g.Save();
                g.SetClip(inn);
                int rh = RowPx;
                for (int i = top, y = inn.Y; i < Count && y < inn.Bottom; i++, y += rh)
                    DrawRow(g, i, new Rectangle(inn.X, y, inn.Width, rh), i == sel, i == hover);
                g.Restore(st);
            }
            Gfx.Frame(g, ClientRectangle, Focused ? Theme.Primary : Theme.Stroke, Dpi.Hair);
        }
    }

    // ------------------------------------------------------------------------------------------ the list of all the words, with editing in place
    class WordList : ScrollList
    {
        List<string> items = new List<string>();
        RowEditor editor;
        string editWord;
        public event EventHandler EditRequested;    // Enter, F2, a double click
        public event EventHandler RemoveRequested;  // Delete
        public event EventHandler EditCommit;       // "אישור" or Enter while editing (the window reads EditWord and EditText)

        public WordList() : base(30) { EmptyText = ""; }

        protected override int Count { get { return items.Count; } }
        protected override bool Locked { get { return Editing; } }
        public string SelectedWord { get { return sel >= 0 && sel < items.Count ? items[sel] : null; } }
        public bool Editing { get { return editor != null && editor.Visible; } }
        public string EditWord { get { return editWord; } }
        public string EditText { get { return editor == null ? "" : editor.Box.Text; } set { if (editor != null) editor.Box.Text = value ?? ""; } }

        // a new list of words (sorted); keepWord = the word to select if it is in it, else the row at fallback (-1 = none)
        public void SetItems(List<string> list, string keepWord, int fallback)
        {
            if (Editing) StopEdit();
            items = list;
            int i = keepWord == null ? -1 : list.BinarySearch(keepWord, StringComparer.Ordinal);
            if (i < 0 && fallback >= 0 && list.Count > 0) i = Math.Min(fallback, list.Count - 1);
            Reset(i);
        }

        public void BeginEdit()
        {
            if (sel < 0 || Editing) return;
            EnsureVisible(sel);
            editWord = items[sel];
            if (editor == null)
            {
                editor = new RowEditor();
                editor.Ok.Click += delegate { if (EditCommit != null) EditCommit(this, EventArgs.Empty); };
                editor.Cancel.Click += delegate { StopEdit(); };
                editor.Box.KeyDown += delegate (object s, KeyEventArgs e)
                {
                    if (e.Control && e.KeyCode == Keys.A) { editor.Box.SelectAll(); e.SuppressKeyPress = true; e.Handled = true; return; }
                    if (e.KeyCode != Keys.Enter) return;
                    e.SuppressKeyPress = true; e.Handled = true;
                    if (EditCommit != null) EditCommit(this, EventArgs.Empty);
                };
                Controls.Add(editor);
            }
            editor.Bounds = RowRect(sel);
            editor.Box.Text = editWord;
            editor.Visible = true;
            editor.BringToFront();
            bar.Enabled = false;
            editor.Box.Focus();
            editor.Box.SelectAll();
            Invalidate();
        }

        public void StopEdit()
        {
            if (editor == null || !editor.Visible) return;
            editor.Visible = false;
            bar.Enabled = true;
            Focus();
            Invalidate();
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            if (!Editing && RowAt(e.Location) >= 0 && EditRequested != null) EditRequested(this, EventArgs.Empty);
            base.OnMouseDoubleClick(e);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (!Editing && sel >= 0)
            {
                if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.F2) && EditRequested != null) { EditRequested(this, EventArgs.Empty); e.Handled = true; e.SuppressKeyPress = true; }
                else if (e.KeyCode == Keys.Delete && RemoveRequested != null) { RemoveRequested(this, EventArgs.Empty); e.Handled = true; }
            }
            base.OnKeyDown(e);
        }

        protected override void DrawRow(Graphics g, int i, Rectangle r, bool selected, bool hovered)
        {
            Gfx.Rect(g, r, selected ? Theme.Soft : (hovered ? Theme.RowHover : Theme.Card));
            if (selected) Gfx.Rect(g, new Rectangle(r.Right - Dpi.S(4), r.Y, Dpi.S(4), r.Height), Theme.Primary);
            Gfx.HLine(g, r.X, r.Bottom - Dpi.Hair, r.Width, Theme.Line);
            Gfx.TextAt(g, items[i], Theme.F(15, false), Theme.Text, new Rectangle(r.X + Dpi.S(14), r.Y, r.Width - Dpi.S(28), r.Height), Align.Right, true, false);
        }
    }

    // the row while it is being edited: the word in a text box, "אישור" and "ביטול"
    class RowEditor : BaseControl
    {
        public readonly TextBox Box = new TextBox();
        public readonly FlatBtn Ok, Cancel;
        Rectangle frame;

        public RowEditor()
        {
            BackColor = Theme.Soft;
            Visible = false;
            Box.BorderStyle = BorderStyle.None;
            Box.RightToLeft = RightToLeft.Yes;
            Box.Font = Theme.F(15, false);
            Box.BackColor = Theme.Card;
            Box.ForeColor = Theme.Text;
            Box.MaxLength = 200;
            Ok = new FlatBtn { Text = "אישור", Style = BtnStyle.Filled, Small = true };
            Cancel = new FlatBtn { Text = "ביטול", Style = BtnStyle.Outlined, Light = true, Small = true };
            Controls.AddRange(new Control[] { Box, Ok, Cancel });
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            int h = Math.Min(Height - 2, Dpi.S(24)), y = (Height - h) / 2, pad = Dpi.S(10), gap = Dpi.S(6), bw = Dpi.S(56);
            Cancel.Bounds = new Rectangle(pad, y, bw, h);
            Ok.Bounds = new Rectangle(pad + bw + gap, y, bw, h);
            int fx = Ok.Right + gap, fr = Width - Dpi.S(14);
            frame = new Rectangle(fx, y, Math.Max(20, fr - fx), h);
            int inset = Dpi.Hair + Dpi.S(5);
            Box.Bounds = new Rectangle(frame.X + inset, frame.Y + (frame.Height - Box.Height) / 2, Math.Max(10, frame.Width - 2 * inset), Box.Height);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            Gfx.Rect(g, ClientRectangle, Theme.Soft);
            Gfx.Rect(g, new Rectangle(Width - Dpi.S(4), 0, Dpi.S(4), Height), Theme.Primary);
            Gfx.Rect(g, frame, Theme.Card);
            Gfx.Frame(g, frame, Theme.Primary, Dpi.Hair);
        }
    }

    // ------------------------------------------------------------------------------------------ the list of the person's changes
    // A coloured mark (+ added, − removed, a pencil = corrected), the word ("wrong ← right" for a correction), and "בטל" on the left.
    class ChangeList : ScrollList
    {
        static readonly Color Green = ColorTranslator.FromHtml("#107C10");
        List<WordChange> items = new List<WordChange>();
        public bool AllowUndo = true;
        public event IndexFn UndoClicked;

        public ChangeList(int rowUnits) : base(rowUnits) { EmptyText = "עוד אין שינויים."; }

        protected override int Count { get { return items.Count; } }

        public void SetItems(List<WordChange> list)
        {
            int keep = sel;
            items = list;
            Reset(keep >= 0 && keep < list.Count ? keep : -1);
        }

        Rectangle UndoRect(Rectangle row) { return new Rectangle(row.X + Dpi.S(8), row.Y, Dpi.S(44), row.Height); }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            int i = RowAt(e.Location);
            Cursor = AllowUndo && i >= 0 && UndoRect(RowRect(i)).Contains(e.Location) ? Cursors.Hand : Cursors.Default;
            base.OnMouseMove(e);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            int i = RowAt(e.Location);
            if (AllowUndo && i >= 0 && UndoRect(RowRect(i)).Contains(e.Location) && UndoClicked != null) UndoClicked(i);
            base.OnMouseClick(e);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (AllowUndo && sel >= 0 && e.KeyCode == Keys.Delete && UndoClicked != null) { UndoClicked(sel); e.Handled = true; }
            base.OnKeyDown(e);
        }

        public static string Describe(WordChange c) { return c.Kind == '~' ? c.Word + " ← " + c.NewWord : c.Word; }

        protected override void DrawRow(Graphics g, int i, Rectangle r, bool selected, bool hovered)
        {
            WordChange c = items[i];
            Gfx.Rect(g, r, selected ? Theme.Soft : (hovered ? Theme.RowHover : Theme.Card));
            if (i < items.Count - 1) Gfx.HLine(g, r.X, r.Bottom - Dpi.Hair, r.Width, Theme.Line);

            int m = Dpi.S(22);
            Rectangle mark = new Rectangle(r.Right - Dpi.S(14) - m, r.Y + (r.Height - m) / 2, m, m);
            DrawMark(g, mark, c.Kind);

            int left = r.X + Dpi.S(14);
            if (AllowUndo)
            {
                Rectangle u = UndoRect(r);
                Gfx.TextAt(g, "בטל", Theme.F(13, false), Theme.Primary, u, Align.Center, true, false);
                left = u.Right + Dpi.S(6);
            }
            if (c.Sent)
            {
                Gfx.TextAt(g, "נשלח", Theme.F(12, false), Theme.TextQuiet, new Rectangle(left, r.Y, Dpi.S(40), r.Height), Align.Left, true, false);
                left += Dpi.S(44);
            }
            Gfx.TextAt(g, Describe(c), Theme.F(15, true), Theme.Text, new Rectangle(left, r.Y, Math.Max(0, mark.X - Dpi.S(10) - left), r.Height), Align.Right, true, false);
        }

        // + on green, − on red, a pencil on blue: drawn, so that no special font is needed
        public static void DrawMark(Graphics g, Rectangle m, char kind)
        {
            Gfx.Fill(g, m, Dpi.S(3), kind == '+' ? Green : kind == '-' ? Theme.Error : Theme.Primary);
            float k = m.Width / 22f;
            using (Pen p = new Pen(Theme.OnPrimary, 2.2f * k))
            {
                p.StartCap = LineCap.Round; p.EndCap = LineCap.Round;
                if (kind == '+')
                {
                    g.DrawLine(p, m.X + 11f * k, m.Y + 6f * k, m.X + 11f * k, m.Y + 16f * k);
                    g.DrawLine(p, m.X + 6f * k, m.Y + 11f * k, m.X + 16f * k, m.Y + 11f * k);
                }
                else if (kind == '-')
                {
                    g.DrawLine(p, m.X + 6f * k, m.Y + 11f * k, m.X + 16f * k, m.Y + 11f * k);
                }
                else
                {
                    using (Pen body = new Pen(Theme.OnPrimary, 3.4f * k))
                    {
                        body.StartCap = LineCap.Flat; body.EndCap = LineCap.Round;
                        g.DrawLine(body, m.X + 8.6f * k, m.Y + 13.4f * k, m.X + 15.4f * k, m.Y + 6.6f * k);
                    }
                    using (SolidBrush b = new SolidBrush(Theme.OnPrimary))
                        g.FillPolygon(b, new PointF[] { new PointF(m.X + 5.4f * k, m.Y + 16.6f * k), new PointF(m.X + 6.6f * k, m.Y + 12.2f * k), new PointF(m.X + 9.8f * k, m.Y + 15.4f * k) });
                }
            }
        }
    }

    // ------------------------------------------------------------------------------------------ small pieces
    // a link: blue underlined text on the left side of its box
    class LinkText : BaseControl
    {
        bool hover;
        Font under;
        public LinkText() { Cursor = Cursors.Hand; TabStop = true; }
        protected override void OnMouseEnter(EventArgs e) { hover = true; Invalidate(); base.OnMouseEnter(e); }
        protected override void OnMouseLeave(EventArgs e) { hover = false; Invalidate(); base.OnMouseLeave(e); }
        protected override void OnGotFocus(EventArgs e) { Invalidate(); base.OnGotFocus(e); }
        protected override void OnLostFocus(EventArgs e) { Invalidate(); base.OnLostFocus(e); }
        protected override bool IsInputKey(Keys k) { return k == Keys.Enter || k == Keys.Space || base.IsInputKey(k); }
        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; }
            base.OnKeyDown(e);
        }
        protected override void Dispose(bool disposing) { if (disposing && under != null) { under.Dispose(); under = null; } base.Dispose(disposing); }
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Surface);
            if (under == null) { Font f = Theme.F(14, false); under = new Font(f, FontStyle.Underline); }
            Color c = !Enabled ? Theme.DisabledText : hover ? Theme.PrimaryDark : Theme.Primary;
            TextRenderer.DrawText(g, Text, under, ClientRectangle, c, TextFormatFlags.RightToLeft | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine | TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
            if (Focused && ShowFocusCues) Gfx.Frame(g, ClientRectangle, Theme.Primary, Dpi.Hair);
        }
    }

    // a big title with a round blue tick after it
    class TitleTick : BaseControl
    {
        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Parent != null ? Parent.BackColor : Theme.Card);
            g.SmoothingMode = SmoothingMode.AntiAlias; g.PixelOffsetMode = PixelOffsetMode.Half;
            Font f = Theme.F(22, true);
            TextFormatFlags fl = TextFormatFlags.RightToLeft | TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding;
            int mark = Dpi.S(22), gap = Dpi.S(10);
            int tw;
            try { tw = TextRenderer.MeasureText(g, Text, f, new Size(Width, Height), fl).Width; }
            catch (Exception) { tw = Width - mark - gap; }
            tw = Math.Max(0, Math.Min(tw, Width - mark - gap));
            TextRenderer.DrawText(g, Text, f, new Rectangle(Width - tw, 0, tw, Height), Theme.Text, fl | TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
            Rectangle m = new Rectangle(Width - tw - gap - mark, (Height - mark) / 2, mark, mark);
            Gfx.Dot(g, m, Theme.Primary);
            float k = mark / 22f;
            using (Pen tick = new Pen(Theme.OnPrimary, 2.6f * k))
            {
                tick.StartCap = LineCap.Round; tick.EndCap = LineCap.Round; tick.LineJoin = LineJoin.Round;
                g.DrawLines(tick, new PointF[] { new PointF(m.X + 6.0f * k, m.Y + 11.4f * k), new PointF(m.X + 9.5f * k, m.Y + 14.8f * k), new PointF(m.X + 16.0f * k, m.Y + 7.6f * k) });
            }
        }
    }

    // ------------------------------------------------------------------------------------------ the word-editing window
    class WordsForm : Form
    {
        readonly bool installed, demo;
        readonly WordEdits edits;
        List<WordChange> saved;   // the person's changes as they are on this computer
        readonly string editionName;
        readonly string allWordsHead;
        TextBlock listLabel, listNote, addHint, chLabel;
        FlatTextBox search, addBox;
        WordList list;
        ChangeList changes;
        FlatBtn btnFix, btnRemove, btnAdd, btnSave;
        System.Windows.Forms.Timer searchTimer;
        string lastQuery;  // what the list shows now (the search waits until the person stops typing for a moment)
        bool closingOk;   // closing after "save", or the person chose not to keep the changes

        public bool Saved;        // the changes were saved
        public bool Sent;         // ... and sent to the dictionary's owner by the program
        public bool ViaBrowser;   // ... or the GitHub page was opened for the person to send them

        public WordsForm(bool extended, bool compose, bool isInstalled, bool demoOnly)
        {
            installed = isInstalled;
            demo = demoOnly;
            editionName = (extended ? "מורחב" : "בסיסי") + (compose ? " + חיבור תורני" : "");
            Stopwatch sw = Stopwatch.StartNew();
            List<string> words = Core.BuildWords(extended, compose);
            saved = demo ? new List<WordChange>() : MyWords.Load();
            edits = new WordEdits(words, saved);
            Core.Log("Word editing (" + editionName + "): " + edits.Count + " words, " + saved.Count + " change(s) of the person, ready after " + sw.ElapsedMilliseconds + " ms");
            allWordsHead = installed ? "כל המילים במילון שלכם" : "כל המילים במילון ה" + editionName;

            Text = MainForm.Title;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            if (Dpi.Forced) { StartPosition = FormStartPosition.Manual; Location = new Point(0, 0); }
            else StartPosition = FormStartPosition.CenterParent;
            BackColor = Theme.Card;
            Font = Theme.F(14, false);
            ClientSize = Dpi.Sz(780, 640);
            Icon = MainForm.AppIcon();

            TextBlock title = new TextBlock { Bounds = Dpi.R(28, 20, 724, 34), Font = Theme.F(22, true), Color = Theme.Text, VCenter = true, Wrap = false, Text = "עריכת מילים" };
            TextBlock sub = new TextBlock
            {
                Bounds = Dpi.R(28, 58, 724, 44), Font = Theme.F(14, false), Color = Theme.TextSoft,
                Text = "חפשו מילה כדי לתקן את האיות שלה או להסיר אותה מהמילון, או הוסיפו מילה חדשה. השינויים "
                     + (installed ? "חלים אצלכם אחרי \"שמור\", ונשארים גם אחרי עדכונים של המילון" : "יחולו אצלכם כשתתקינו את המילון")
                     + ", ואפשר לשלוח אותם כהצעה לשיפור המילון לכולם."
            };

            // right side: every word of the dictionary
            listLabel = new TextBlock { Bounds = Dpi.R(322, 108, 430, 22), Font = Theme.F(14, true), Color = Theme.Text, VCenter = true, Wrap = false };
            search = new FlatTextBox(15) { Bounds = Dpi.R(322, 134, 430, 36), SideUnits = 110 };
            search.Cue = "חיפוש מילה במילון";
            list = new WordList { Bounds = Dpi.R(322, 178, 430, 302) };
            listNote = new TextBlock { Bounds = Dpi.R(322, 484, 430, 22), Font = Theme.F(13, false), Color = Theme.TextQuiet, VCenter = true, Wrap = false };
            btnFix = new FlatBtn { Bounds = Dpi.R(612, 512, 140, 36), Text = "תקן איות", Style = BtnStyle.Outlined, Light = true };
            btnRemove = new FlatBtn { Bounds = Dpi.R(460, 512, 140, 36), Text = "הסר מהמילון", Style = BtnStyle.Outlined, Light = true };

            // left side: a new word, and the person's changes
            TextBlock addLabel = new TextBlock { Bounds = Dpi.R(28, 108, 280, 22), Font = Theme.F(14, true), Color = Theme.Text, VCenter = true, Wrap = false, Text = "הוספת מילה חדשה" };
            addBox = new FlatTextBox(15) { Bounds = Dpi.R(118, 134, 190, 36) };
            addBox.Cue = "מילה חדשה";
            btnAdd = new FlatBtn { Bounds = Dpi.R(28, 134, 80, 36), Text = "הוסף", Style = BtnStyle.Filled };
            addHint = new TextBlock { Bounds = Dpi.R(28, 175, 280, 38), Font = Theme.F(13, false), Color = Theme.Primary };
            chLabel = new TextBlock { Bounds = Dpi.R(28, 214, 280, 22), Font = Theme.F(14, true), Color = Theme.Text, VCenter = true, Wrap = false };
            changes = new ChangeList(40) { Bounds = Dpi.R(28, 240, 280, 242) };
            TextBlock chNote = new TextBlock { Bounds = Dpi.R(28, 488, 280, 40), Font = Theme.F(13, false), Color = Theme.TextQuiet,
                Text = "השינויים נשמרים במחשב שלכם, וחלים גם אחרי כל עדכון של המילון." };

            LowerPanel lower = new LowerPanel { Bounds = Dpi.R(0, 568, 780, 72) };
            btnSave = new FlatBtn { Bounds = Dpi.R(532, 16, 220, 40), Text = installed ? "שמור ועדכן את וורד" : "שמור", Style = BtnStyle.Filled };
            FlatBtn btnClose = new FlatBtn { Bounds = Dpi.R(28, 16, 110, 40), Text = "סגור", Style = BtnStyle.Outlined, Light = true };
            lower.Controls.AddRange(new Control[] { btnSave, btnClose });

            // the order of the Tab key: right to left, top to bottom
            Controls.AddRange(new Control[] { title, sub, listLabel, search, list, listNote, btnFix, btnRemove, addLabel, addBox, btnAdd, addHint, chLabel, changes, chNote, lower });

            searchTimer = new System.Windows.Forms.Timer { Interval = 150 };
            searchTimer.Tick += delegate { searchTimer.Stop(); if (search.Value.Trim() != lastQuery) ApplySearch(list.SelectedWord, -1); };
            search.Changed += delegate
            {
                if (search.Value.Trim() == lastQuery) return;
                if (list.Editing) { list.StopEdit(); Note("", false); }   // a new search: the row being edited is let go
                searchTimer.Stop();
                searchTimer.Start();
            };
            search.Submit += delegate { searchTimer.Stop(); ApplySearch(list.SelectedWord, -1); if (list.SelectedIndex < 0) list.Select(0); list.Focus(); };
            list.SelectionChanged += delegate { UpdateButtons(); };
            list.EditRequested += delegate { StartFix(); };
            list.RemoveRequested += delegate { DoRemove(); };
            list.EditCommit += delegate { CommitFix(); };
            btnFix.Click += delegate { StartFix(); };
            btnRemove.Click += delegate { DoRemove(); };
            addBox.Changed += delegate { UpdateAddHint(); };
            addBox.Submit += delegate { DoAdd(); };
            btnAdd.Click += delegate { DoAdd(); };
            changes.UndoClicked += delegate (int i) { DoUndo(i); };
            btnSave.Click += delegate { DoSave(); };
            btnClose.Click += delegate { Close(); };

            RefreshAll(null, -1);
            ActiveControl = search.Box;
        }

        bool Dirty { get { return !SameChanges(edits.Changes, saved); } }

        static bool SameChanges(List<WordChange> a, List<WordChange> b)
        {
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
                if (a[i].Kind != b[i].Kind || a[i].Word != b[i].Word || a[i].NewWord != b[i].NewWord) return false;
            return true;
        }

        static List<WordChange> Copy(List<WordChange> list)
        {
            List<WordChange> c = new List<WordChange>(list.Count);
            foreach (WordChange x in list) c.Add(x.Copy());
            return c;
        }

        void ApplySearch(string keepWord, int fallback)
        {
            string q = search.Value.Trim();
            lastQuery = q;
            List<string> found = edits.Find(q);
            list.EmptyText = q.Length > 0 ? "המילה לא נמצאה במילון. אפשר להוסיף אותה משמאל." : "";
            list.SetItems(found, keepWord, fallback);
            search.Side = q.Length > 0 ? "נמצאו " + found.Count.ToString("N0") : "";
            UpdateButtons();
        }

        // after every change: the list, the search, the person's changes, the counts, the buttons
        void RefreshAll(string keepWord, int fallback)
        {
            listLabel.Text = allWordsHead + " (" + edits.Count.ToString("N0") + ")";
            ApplySearch(keepWord, fallback);
            changes.SetItems(edits.Changes);
            chLabel.Text = "השינויים שלי (" + edits.Changes.Count.ToString("N0") + ")";
            UpdateAddHint();
            UpdateButtons();
        }

        void UpdateButtons()
        {
            bool row = list.SelectedWord != null && !list.Editing;
            btnFix.Enabled = row;
            btnRemove.Enabled = row;
            btnSave.Enabled = Dirty;
        }

        void Note(string text, bool error)
        {
            listNote.Color = error ? Theme.Error : Theme.TextQuiet;
            listNote.Text = text;
        }

        static string Q(string w) { return "\"" + w + "\""; }

        void UpdateAddHint()
        {
            string raw = addBox.Value;
            string w = WordEdits.Clean(raw);
            addHint.Color = Theme.Primary;
            if (raw.Trim().Length == 0) addHint.Text = "";
            else if (w == null) { addHint.Color = Theme.Error; addHint.Text = "אפשר להוסיף מילה אחת בלבד, באותיות עבריות (מותר גם ' ו-\")."; }
            else if (edits.Contains(w)) addHint.Text = Q(w) + " כבר נמצאת במילון.";
            else addHint.Text = Q(w) + " לא נמצאת במילון – אפשר להוסיף.";
        }

        void DoAdd()
        {
            string raw = addBox.Value;
            EditResult r = edits.Add(raw);
            if (r == EditResult.Done)
            {
                string w = WordEdits.Clean(raw);
                addBox.Value = "";
                RefreshAll(list.SelectedWord, -1);
                addHint.Color = Theme.Primary;
                addHint.Text = Q(w) + " נוספה למילון.";
                addBox.FocusBox();
            }
            else UpdateAddHint();
        }

        void StartFix()
        {
            if (list.SelectedWord == null || list.Editing) return;
            list.BeginEdit();
            UpdateButtons();
            Note("כתבו את האיות הנכון ולחצו \"אישור\" (או Enter). Esc – ביטול.", false);
        }

        void CommitFix()
        {
            if (!list.Editing) return;
            string wrong = list.EditWord, right = WordEdits.Clean(list.EditText);
            EditResult r = edits.Fix(wrong, list.EditText);
            if (r == EditResult.Invalid) { Note("האיות הנכון צריך להיות מילה אחת, באותיות עבריות.", true); return; }
            list.StopEdit();
            if (r == EditResult.Same) { Note("", false); UpdateButtons(); return; }
            RefreshAll(right, list.SelectedIndex);
            Note(r == EditResult.Done ? "תוקן: " + wrong + " ← " + right : "", false);
        }

        void DoRemove()
        {
            string w = list.SelectedWord;
            if (w == null || list.Editing) return;
            int at = list.SelectedIndex;
            if (edits.Remove(w) != EditResult.Done) return;
            RefreshAll(null, at);
            Note("הוסרה: " + w + ". אפשר לבטל ב\"השינויים שלי\".", false);
            list.Focus();
        }

        void DoUndo(int i)
        {
            if (i < 0 || i >= edits.Changes.Count) return;
            if (list.Editing) list.StopEdit();
            string what = ChangeList.Describe(edits.Changes[i]);
            edits.Undo(i);
            RefreshAll(list.SelectedWord, -1);
            Note("בוטל: " + what, false);
        }

        // saves the changes, offers to send them (the ones not sent yet), and closes; the main window then updates Word
        void DoSave()
        {
            if (list.Editing)
            {
                CommitFix();
                if (list.Editing) return;   // not a valid word: the person sees why, nothing is saved yet
            }
            if (!Dirty) return;
            if (!demo)
            {
                try { MyWords.Save(edits.Changes); }
                catch (Exception ex)
                {
                    Core.Log("Saving the person's changes failed: " + Core.Msg(ex));
                    MsgDialog.Tell(this, "לא הצלחתי לשמור את השינויים:\r\n" + Core.Msg(ex), "השמירה לא הצליחה", MsgKind.Error);
                    return;
                }
            }
            Core.Log("The person's changes were saved: " + edits.Changes.Count);
            Saved = true;
            saved = Copy(edits.Changes);

            List<WordChange> unsent = edits.Changes.FindAll(delegate (WordChange c) { return !c.Sent; });
            if (unsent.Count > 0)
            {
                using (SuggestDialog d = new SuggestDialog(unsent, editionName))
                {
                    d.ShowDialog(this);
                    if (d.Sent || d.ViaBrowser)
                    {
                        Sent = d.Sent; ViaBrowser = d.ViaBrowser;
                        foreach (WordChange c in unsent) c.Sent = true;   // (the same objects as in edits.Changes)
                        if (!demo)
                        {
                            try { MyWords.Save(edits.Changes); }
                            catch (Exception ex) { Core.Log("Saving the 'sent' marks failed: " + Core.Msg(ex)); }
                        }
                    }
                }
            }
            closingOk = true;
            DialogResult = DialogResult.OK;
            Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!closingOk && !demo && Dirty && e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult a = MsgDialog.Show(this, "יש שינויים שעוד לא נשמרו. לשמור אותם לפני הסגירה?", "לשמור את השינויים?", MsgKind.Question,
                    new string[] { "שמור", "אל תשמור", "ביטול" }, new DialogResult[] { DialogResult.Yes, DialogResult.No, DialogResult.Cancel }, 0, DialogResult.Cancel);
                if (a == DialogResult.Cancel) { e.Cancel = true; return; }
                if (a == DialogResult.Yes) { e.Cancel = true; BeginInvoke((MethodInvoker)DoSave); return; }
                closingOk = true;
            }
            if (searchTimer != null) searchTimer.Stop();
            base.OnFormClosing(e);
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                if (list.Editing) { list.StopEdit(); Note("", false); UpdateButtons(); return true; }
                Close();
                return true;
            }
            return base.ProcessDialogKey(keyData);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && searchTimer != null) { searchTimer.Dispose(); searchTimer = null; }
            base.Dispose(disposing);
        }

        // ---------------------------------------------------------------- for the developer's screenshots (--preview-words): nothing is saved
        public void Demo()
        {
            foreach (WordChange c in DemoChanges())
            {
                if (c.Kind == '+') edits.Add(c.Word);
                else if (c.Kind == '-') edits.Remove(c.Word);
                else edits.Fix(c.Word, c.NewWord);
            }
            search.Value = "בסוגיי";
            addBox.Value = "ודמסתייעא";
            RefreshAll("בסוגייין", -1);
            list.BeginEdit();
            list.EditText = "בסוגיין";
            UpdateButtons();
            Note("כתבו את האיות הנכון ולחצו \"אישור\" (או Enter). Esc – ביטול.", false);
        }

        public static List<WordChange> DemoChanges()
        {
            List<WordChange> l = new List<WordChange>();
            l.Add(new WordChange('+', "דמסתייעא", null));
            l.Add(new WordChange('-', "בבגמרא", null));
            l.Add(new WordChange('~', "בסוגיייתנו", "בסוגייתנו"));
            return l;
        }
    }

    // ------------------------------------------------------------------------------------------ after "save": the suggestion to the dictionary's owner
    class SuggestDialog : Form
    {
        readonly List<WordChange> items;
        readonly string edition;
        readonly bool hasKey;
        FlatTextBox note;
        FlatBtn btnSend, btnLater;
        LinkText link;
        TextBlock status;
        bool sending;
        public bool Sent;         // sent by the program (a new issue)
        public bool ViaBrowser;   // the GitHub page was opened instead
        public int Number;        // the number of the new issue

        public SuggestDialog(List<WordChange> unsent, string editionName)
        {
            items = unsent;
            edition = editionName;
            hasKey = Suggest.Token() != null;

            Text = MainForm.Title;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            if (Dpi.Forced) { StartPosition = FormStartPosition.Manual; Location = new Point(0, 0); }
            else StartPosition = FormStartPosition.CenterParent;
            BackColor = Theme.Card;
            Font = Theme.F(14, false);
            ClientSize = Dpi.Sz(560, 440);
            Icon = MainForm.AppIcon();

            TitleTick head = new TitleTick { Bounds = Dpi.R(28, 20, 504, 34), Text = "השינויים נשמרו" };
            TextBlock sub = new TextBlock
            {
                Bounds = Dpi.R(28, 60, 504, 62), Font = Theme.F(14, false), Color = Theme.TextSoft,
                Text = "אפשר לשלוח אותם כהצעה לשיפור המילון לכולם, ובעל המילון יחליט אם להכניס אותם לגרסה הבאה. "
                     + "ההצעה מתפרסמת בגלוי בגיטהב של המילון, בלי שם ובלי פרטים מזהים."
            };
            ChangeList list = new ChangeList(34) { Bounds = Dpi.R(28, 130, 504, 106), AllowUndo = false, TabStop = false };
            list.SetItems(items);
            TextBlock noteLabel = new TextBlock { Bounds = Dpi.R(28, 248, 504, 22), Font = Theme.F(14, true), Color = Theme.Text, VCenter = true, Wrap = false,
                Text = "הערה (לא חובה) – למשל מאיפה המילה:" };
            note = new FlatTextBox(15) { Bounds = Dpi.R(28, 274, 504, 36) };
            note.Cue = "למשל: גמרא, ברכות דף ב";
            status = new TextBlock { Bounds = Dpi.R(28, 316, 504, 24), Font = Theme.F(13, false), Color = Theme.TextQuiet, VCenter = true, Wrap = false,
                Text = hasKey ? "" : "השליחה היא דרך אתר גיטהב, וצריך חשבון גיטהב." };

            LowerPanel strip = new LowerPanel { Bounds = Dpi.R(0, 350, 560, 90) };
            btnSend = new FlatBtn { Bounds = Dpi.R(392, 24, 140, 42), Text = hasKey ? "שלח" : "שלח דרך גיטהב", Style = BtnStyle.Filled };
            btnLater = new FlatBtn { Bounds = Dpi.R(272, 24, 110, 42), Text = "לא עכשיו", Style = BtnStyle.Outlined, Light = true };
            link = new LinkText { Bounds = Dpi.R(28, 34, 236, 22), Text = "יש לי חשבון גיטהב – שלח דרך גיטהב", Visible = hasKey };
            btnSend.Click += delegate { DoSend(); };
            btnLater.Click += delegate { Close(); };
            link.Click += delegate { if (!sending) OpenBrowser(); };
            strip.Controls.AddRange(new Control[] { btnSend, btnLater, link });

            Controls.AddRange(new Control[] { head, sub, list, noteLabel, note, status, strip });
            ActiveControl = btnSend;
        }

        string MakeTitle() { return Suggest.Title(items); }
        string MakeBody() { return Suggest.Body(items, note.Value, edition); }

        void SetEnabled(bool on) { btnSend.Enabled = on; btnLater.Enabled = on; link.Enabled = on; note.Enabled = on; }

        void DoSend()
        {
            if (sending) return;
            if (!hasKey) { OpenBrowser(); return; }
            sending = true;
            SetEnabled(false);
            UseWaitCursor = true;
            status.Color = Theme.TextQuiet;
            status.Text = "שולח…";
            string title = MakeTitle(), body = MakeBody();
            Thread t = new Thread(delegate ()
            {
                int number; string address, error;
                bool ok = Suggest.Send(title, body, out number, out address, out error);
                try
                {
                    BeginInvoke((MethodInvoker)delegate
                    {
                        sending = false;
                        UseWaitCursor = false;
                        if (ok) { Sent = true; Number = number; Close(); return; }
                        SetEnabled(true);
                        status.Color = Theme.Error;
                        status.Text = "השליחה לא הצליחה.";
                        DialogResult a = MsgDialog.Show(this, "השליחה האוטומטית לא הצליחה (" + error + ").\r\n\r\nאפשר לשלוח את ההצעה דרך אתר גיטהב, אם יש לכם חשבון שם.",
                            "השליחה לא הצליחה", MsgKind.Warn, new string[] { "פתח את גיטהב", "לא עכשיו" }, new DialogResult[] { DialogResult.Yes, DialogResult.No }, 0, DialogResult.No);
                        if (a == DialogResult.Yes) OpenBrowser();
                    });
                }
                catch (InvalidOperationException) { }   // the window is gone
            });
            t.IsBackground = true;
            t.Start();
        }

        // the "new issue" page of GitHub, filled in; a long list goes through the clipboard (see Suggest.BrowserUrl)
        void OpenBrowser()
        {
            string title = MakeTitle(), body = MakeBody();
            bool cut;
            string url = Suggest.BrowserUrl(title, body, out cut);
            if (cut)
            {
                try { Clipboard.SetText(body); }
                catch (Exception ex) { Core.Log("Copying the suggestion failed: " + Core.Msg(ex)); }
                MsgDialog.Tell(this, "בדף של גיטהב שייפתח עכשיו, מחקו את השורה שבתיבת הטקסט והדביקו במקומה (Ctrl+V).", "הרשימה ארוכה, ולכן העתקתי אותה", MsgKind.Info);
            }
            try
            {
                Process.Start(url);
                Core.Log("The GitHub page for the suggestion was opened");
                ViaBrowser = true;
                Close();
            }
            catch (Exception ex)
            {
                Core.Log("Opening the browser failed: " + Core.Msg(ex));
                MsgDialog.Tell(this, Core.Msg(ex), "לא הצלחתי לפתוח את הדפדפן", MsgKind.Error);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (sending) { e.Cancel = true; return; }   // (at most 30 seconds: the time limit of the sending)
            base.OnFormClosing(e);
        }

        protected override bool ProcessDialogKey(Keys keyData)
        {
            if (keyData == Keys.Escape) { if (!sending) Close(); return true; }
            return base.ProcessDialogKey(keyData);
        }
    }
}
