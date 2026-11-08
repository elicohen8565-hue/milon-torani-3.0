using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

// Repair (a developer tool, 9-11-2026): Word would not add any custom dictionary any more (CustomDictionaries.Add -> E_FAIL),
// so the Torah dictionary could not be connected back after the size check.
// The cause, seen in the user's log: Word 365 keeps its list of custom dictionaries in the registry, one numbered entry per
// dictionary, for each signed-in account ("4_16_7bd116fa7f21d40b_LiveId" = the file, "..._state", "..._roamed", "..._external").
// A dictionary taken out of Word keeps its entry, and a new file name takes a new number. The list then had 10 entries:
// RoamingCustom.dic, two files of version 3.0.0.1, Milon_Torani.dic, a WordCheck trial and five size-check files. The sixth size-check
// file (500,000 words) was the first that Word refused, and from then on it refused every dictionary - most likely a limit of 10.
// This tool, with Word and the other Office programs closed:
//   1. keeps a copy of the list (reg export, and every value in the report),
//   2. deletes the entries of the program's own files (Milon_Torani.dic and the files of older versions), of the size check and of
//      the WordCheck trials - and nothing else (the person's own dictionaries stay; their numbers are closed up if a gap is left),
//   3. lets a hidden Word add the Torah dictionary again, and a new hidden Word checks that it knows its words (Verify).
// The report goes to the desktop: Milon_Repair_report.txt.  Built by tools/repair/build.sh; self-test: --selftest <folder> (Mono only).
namespace OtzariaDict
{
    public class DictEntry
    {
        public int Number;
        public string Identity = "";                          // "" = an entry without an account; else e.g. "16_7bd116fa7f21d40b_LiveId"
        public string Main;                                   // the value that holds the file (null: only settings are left)
        public string File;
        public List<string> Names = new List<string>();       // every value of the entry: the file and its settings
        public string Leaf
        {
            get
            {
                if (File == null) return null;
                try { return Path.GetFileName(File); } catch (Exception) { return File; }
            }
        }
    }

    public static class ListFix
    {
        static readonly string[] SettingNames = { "state", "roamed", "external", "culturetag" };

        // "4_16_x_LiveId_state" -> 4, "16_x_LiveId", setting;  "11" -> 11, "", main;  "1_state" -> 1, "", setting;  "UpdateComplete" -> false
        public static bool Parse(string vn, out int number, out string identity, out bool isMain)
        {
            number = -1; identity = ""; isMain = false;
            int cut = vn.IndexOf('_');
            string head = cut < 0 ? vn : vn.Substring(0, cut);
            if (head.Length == 0 || !int.TryParse(head, NumberStyles.None, CultureInfo.InvariantCulture, out number)) return false;
            string rest = cut < 0 ? "" : vn.Substring(cut + 1);
            if (rest.Length == 0) { isMain = true; return true; }
            foreach (string s in SettingNames)
            {
                if (rest == s) return true;                                       // a setting of the entry without an account
                if (rest.EndsWith("_" + s, StringComparison.Ordinal)) { identity = rest.Substring(0, rest.Length - s.Length - 1); return true; }
            }
            identity = rest;
            isMain = true;
            return true;
        }

        public static List<DictEntry> Read(RegistryKey k)
        {
            Dictionary<string, DictEntry> by = new Dictionary<string, DictEntry>(StringComparer.Ordinal);
            List<DictEntry> list = new List<DictEntry>();
            foreach (string vn in k.GetValueNames())
            {
                int n; string id; bool main;
                if (!Parse(vn, out n, out id, out main)) continue;
                string key = n + "|" + id;
                DictEntry e;
                if (!by.TryGetValue(key, out e)) { e = new DictEntry(); e.Number = n; e.Identity = id; by[key] = e; list.Add(e); }
                e.Names.Add(vn);
                if (main) { e.Main = vn; e.File = Convert.ToString(k.GetValue(vn)); }
            }
            list.Sort(delegate (DictEntry a, DictEntry b) { int c = string.CompareOrdinal(a.Identity, b.Identity); return c != 0 ? c : a.Number.CompareTo(b.Number); });
            return list;
        }

        // the program's own files, the size check's files and the WordCheck trials - never a dictionary of the person
        public static bool IsStale(DictEntry e)
        {
            string leaf = e.Leaf;
            if (leaf == null) return false;
            return Core.IsOurs(leaf)
                || leaf.StartsWith(SizeTestPrefix, StringComparison.OrdinalIgnoreCase)
                || (leaf.StartsWith("wordcheck", StringComparison.OrdinalIgnoreCase) && leaf.EndsWith(".dic", StringComparison.OrdinalIgnoreCase));
        }
        public const string SizeTestPrefix = "Milon_SizeTest_";

        static void Rename(RegistryKey k, string from, string to)
        {
            object v = k.GetValue(from, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            RegistryValueKind kind = k.GetValueKind(from);
            k.SetValue(to, v, kind);
            k.DeleteValue(from, false);
        }

        // deletes the stale entries and closes up the numbers of each account's list (1, 2, 3 ... like Word writes them);
        // returns what was done, one line each
        public static List<string> Clean(RegistryKey k)
        {
            List<string> done = new List<string>();
            List<DictEntry> all = Read(k);
            foreach (DictEntry e in all)
            {
                if (!IsStale(e)) continue;
                foreach (string vn in e.Names) k.DeleteValue(vn, false);
                done.Add("deleted " + e.Number + (e.Identity.Length > 0 ? " (" + e.Identity + ")" : "") + " " + e.File);
            }
            // close up the gaps: entries are moved down in order, so the number they move to is always free
            Dictionary<string, List<DictEntry>> byId = new Dictionary<string, List<DictEntry>>(StringComparer.Ordinal);
            foreach (DictEntry e in Read(k))
            {
                if (e.Main == null) continue;   // settings without a file (left by Office itself): not touched
                List<DictEntry> l;
                if (!byId.TryGetValue(e.Identity, out l)) { l = new List<DictEntry>(); byId[e.Identity] = l; }
                l.Add(e);
            }
            foreach (KeyValuePair<string, List<DictEntry>> kv in byId)
            {
                kv.Value.Sort(delegate (DictEntry a, DictEntry b) { return a.Number.CompareTo(b.Number); });
                for (int i = 0; i < kv.Value.Count; i++)
                {
                    DictEntry e = kv.Value[i];
                    int want = i + 1;
                    if (e.Number == want) continue;
                    string oldHead = e.Number.ToString(CultureInfo.InvariantCulture), newHead = want.ToString(CultureInfo.InvariantCulture);
                    foreach (string vn in e.Names) Rename(k, vn, newHead + vn.Substring(oldHead.Length));
                    done.Add("moved " + e.Number + " -> " + want + (e.Identity.Length > 0 ? " (" + e.Identity + ")" : "") + " " + e.File);
                    e.Number = want;
                }
            }
            return done;
        }

        public static string Dump(RegistryKey k)
        {
            StringBuilder sb = new StringBuilder();
            if (k == null) return "(the key does not exist)\r\n";
            List<string> names = new List<string>(k.GetValueNames());
            names.Sort(StringComparer.Ordinal);
            foreach (string vn in names)
            {
                object v = k.GetValue(vn, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                string text;
                byte[] b = v as byte[];
                if (b != null)
                {
                    StringBuilder h = new StringBuilder();
                    foreach (byte x in b) h.Append(x.ToString("x2", CultureInfo.InvariantCulture));
                    text = "hex:" + h;
                }
                else text = Convert.ToString(v, CultureInfo.InvariantCulture);
                sb.Append("  " + vn + " = " + text + "\r\n");
            }
            return sb.ToString();
        }
    }

    public class RepairJob
    {
        public static string StateDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MilonRepair");
        public static string ReportPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Milon_Repair_report.txt");
        public static string SizeCheckDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MilonSizeCheck");
        public static Func<List<string>> OfficeHook;   // the self-test: pretend Office programs that are open
        static readonly string[] Office = { "WINWORD", "OUTLOOK", "EXCEL", "POWERPNT", "ONENOTE", "MSACCESS", "MSPUB", "VISIO" };

        public Action<int, string> Progress = delegate { };
        public bool Ok, OursMissing;
        public string Error;
        public int Verdict = -1;
        readonly StringBuilder report = new StringBuilder();

        void Say(int p, string t) { Core.Log("Repair: " + t); try { Progress(p, t); } catch (Exception) { } }
        void R(string line) { report.Append(line + "\r\n"); }

        // the Office programs that are open now (they share Word's list of dictionaries, and could write their old copy back)
        public static List<string> OpenOffice()
        {
            if (OfficeHook != null) return OfficeHook();
            List<string> open = new List<string>();
            foreach (string name in Office)
            {
                try
                {
                    Process[] ps = Process.GetProcessesByName(name);
                    if (ps.Length > 0) open.Add(name);
                    foreach (Process p in ps) p.Dispose();
                }
                catch (Exception) { }
            }
            return open;
        }

        static string Backup()
        {
            Directory.CreateDirectory(StateDir);
            string file = Path.Combine(StateDir, "custom-dictionaries-" + DateTime.Now.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) + ".reg");
            string reg = string.IsNullOrEmpty(Environment.SystemDirectory) ? null : Path.Combine(Environment.SystemDirectory, "reg.exe");
            if (reg == null || !File.Exists(reg)) return "no reg.exe (the values are in this report)";
            ProcessStartInfo si = new ProcessStartInfo(reg, "export \"HKCU\\" + RegFallback.KeyPath + "\" \"" + file + "\" /y");
            si.UseShellExecute = false; si.CreateNoWindow = true;
            using (Process p = Process.Start(si))
            {
                if (!p.WaitForExit(30000)) { try { p.Kill(); } catch (Exception) { } return "reg export did not finish"; }
                return p.ExitCode == 0 && File.Exists(file) ? file : "reg export failed (" + p.ExitCode + ")";
            }
        }

        public void Run()
        {
            string ours = Path.Combine(Core.UProofDir(), Core.MainFile);
            try
            {
                R("#REPAIR 1");
                R("#DATE " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                List<string> open = OpenOffice();
                if (open.Count > 0) { Error = "office"; R("#STOP open: " + string.Join(", ", open.ToArray())); return; }

                Say(10, "שומר עותק של רשימת המילונים של וורד…");
                R("#BACKUP " + Backup());
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath, true))
                {
                    R("#BEFORE");
                    report.Append(ListFix.Dump(k));
                    if (k != null)
                    {
                        Say(25, "מוחק מהרשימה את הרשומות הישנות…");
                        foreach (string line in ListFix.Clean(k)) R("#CLEAN " + line);
                        R("#AFTER_CLEAN");
                        report.Append(ListFix.Dump(k));
                    }
                }

                if (!File.Exists(ours))
                {
                    OursMissing = true;
                    R("#NO_FILE " + ours);
                }
                else
                {
                    Say(45, "מחבר את המילון התורני לוורד…");
                    using (WordSession w = WordCom.OpenNew())
                    {
                        StringBuilder ds = new StringBuilder();
                        foreach (DictInfo di in WordCom.ListAll(w)) ds.Append((ds.Length > 0 ? " | " : "") + di.Name);
                        R("#WORD_LIST " + ds);
                        WordCom.RemoveOurs(w);
                        try { WordCom.Add(w, ours); R("#ADD ok"); }
                        catch (Exception ex) { R("#ADD FAILED " + Core.Msg(ex)); throw; }
                        WordCom.KeepOursNotDefault(w, Core.UProofDir());
                    }
                    WordCom.WaitForOurWordsToExit(15000);
                    using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath))
                    {
                        R("#AFTER_ADD");
                        report.Append(ListFix.Dump(k));
                    }

                    Say(70, "בודק בוורד חדש שהוא מכיר את המילים…");
                    using (WordSession v = WordCom.OpenNew())
                    {
                        string check = Verify.Run(v, Core.ReadDicWords(ours), true, out Verdict);
                        R("#VERIFY " + Verdict + " " + check.Replace("\r", " ").Replace("\n", " "));
                    }
                    WordCom.WaitForOurWordsToExit(15000);
                }

                // the size check's leftovers: its restore mark (its restore failed) and test files, if any are left
                try
                {
                    string mark = Path.Combine(SizeCheckDir, "restore.txt");
                    if (File.Exists(mark) && (Verdict == 1 || OursMissing)) { File.Delete(mark); R("#SIZECHECK_MARK deleted"); }
                    foreach (string f in Directory.GetFiles(Core.UProofDir(), ListFix.SizeTestPrefix + "*.dic")) { File.Delete(f); R("#DELETED " + f); }
                }
                catch (Exception ex) { R("#LEFTOVERS " + Core.Msg(ex)); }
                Ok = OursMissing || Verdict == 1;
            }
            catch (Exception ex)
            {
                Error = ex is WordMissingException ? "לא מצאתי את וורד במחשב." : Core.Msg(ex);
                R("#ERROR " + Core.Msg(ex).Replace("\r", " ").Replace("\n", " "));
                Core.Log("Repair failed: " + ex);
            }
            finally
            {
                R("#DONE " + (Ok ? "ok" : "FAILED"));
                try { File.WriteAllText(ReportPath, report.ToString(), new UTF8Encoding(true)); }
                catch (Exception ex)
                {
                    Core.Log("Repair report: " + Core.Msg(ex));
                    try { Directory.CreateDirectory(StateDir); ReportPath = Path.Combine(StateDir, "Milon_Repair_report.txt"); File.WriteAllText(ReportPath, report.ToString(), new UTF8Encoding(true)); }
                    catch (Exception) { }
                }
                Say(100, "הסתיים");
            }
        }
    }

    // ------------------------------------------------------------------------------------------ the window
    class RepairForm : Form
    {
        public const string Title = "תיקון המילון התורני בוורד";
        TextBlock body, stage;
        ProgressLine prog;
        FlatBtn btnFix, btnClose;
        bool running;
        Mutex mainMutex;

        public RepairForm()
        {
            Text = Title;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            if (Dpi.Forced) { StartPosition = FormStartPosition.Manual; Location = new Point(0, 0); }
            else StartPosition = FormStartPosition.CenterScreen;
            BackColor = Theme.Card;
            Font = Theme.F(14, false);
            ClientSize = Dpi.Sz(600, 420);
            try { Icon = MainForm.AppIcon(); } catch (Exception) { }

            TextBlock title = new TextBlock { Bounds = Dpi.R(28, 20, 544, 34), Font = Theme.F(22, true), Color = Theme.Text, VCenter = true, Wrap = false, Text = Title };
            body = new TextBlock
            {
                Bounds = Dpi.R(28, 66, 544, 210), Font = Theme.F(14, false), Color = Theme.Text,
                Text = "וורד מסרב עכשיו לקבל מילונים, כי ברשימת המילונים שלו הצטברו רשומות ישנות (רובן של בדיקת הגודל).\r\n\r\n"
                     + "התוכנה מוחקת מהרשימה רק את הרשומות של המילון התורני ושל הבדיקות, ומחברת את המילון התורני מחדש. "
                     + "המילון האישי שלכם נשאר כמו שהוא, ולפני השינוי נשמר עותק של הרשימה.\r\n\r\n"
                     + "לפני שמתחילים: סגרו את וורד ואת שאר תוכנות אופיס (גם אאוטלוק, אם הוא פתוח). זה לוקח פחות מדקה."
            };
            prog = new ProgressLine { Bounds = Dpi.R(28, 290, 544, 4) };
            stage = new TextBlock { Bounds = Dpi.R(28, 302, 544, 36), Font = Theme.F(14, false), Color = Theme.TextSoft, Text = "" };
            LowerPanel lower = new LowerPanel { Bounds = Dpi.R(0, 344, 600, 76) };
            btnFix = new FlatBtn { Bounds = Dpi.R(432, 16, 140, 44), Text = "תקן", Style = BtnStyle.Filled };
            btnClose = new FlatBtn { Bounds = Dpi.R(28, 16, 110, 44), Text = "סגור", Style = BtnStyle.Outlined, Light = true };
            btnFix.Click += delegate { Start(); };
            btnClose.Click += delegate { Close(); };
            lower.Controls.AddRange(new Control[] { btnFix, btnClose });
            Controls.AddRange(new Control[] { title, body, prog, stage, lower });
        }

        static void Info(IWin32Window owner, string text)
        {
            MessageBox.Show(owner, text, Title, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }

        void UI(MethodInvoker a)
        {
            try { if (!IsDisposed && IsHandleCreated) BeginInvoke(a); } catch (Exception) { }
        }

        public void PreviewDone()
        {
            RepairJob j = new RepairJob();
            j.Ok = true; j.Verdict = 1;
            body.Text = Summary(j);
            prog.Value = 100;
        }

        void Start()
        {
            if (running) return;
            List<string> open = RepairJob.OpenOffice();
            if (open.Count > 0)
            {
                Info(this, "תוכנות אופיס פתוחות עכשיו (" + string.Join(", ", open.ToArray()) + ").\r\n\r\nשמרו וסגרו את וורד ואת שאר תוכנות אופיס, גם את אאוטלוק, ולחצו שוב \"תקן\".\r\n\r\nאם אין חלון פתוח, חכו רגע ונסו שוב.");
                return;
            }
            bool created;
            try { mainMutex = new Mutex(true, @"Local\TorahDictionaryForWord.SingleInstance", out created); }
            catch (Exception) { mainMutex = null; created = true; }
            if (!created)
            {
                mainMutex.Close(); mainMutex = null;
                Info(this, "התוכנה \"המילון התורני\" פתוחה עכשיו.\r\n\r\nסגרו אותה ולחצו שוב \"תקן\".");
                return;
            }
            running = true;
            btnFix.Enabled = false;
            UseWaitCursor = true;
            RepairJob job = new RepairJob();
            job.Progress = delegate (int p, string s) { UI(delegate { prog.Value = p; stage.Text = s; }); };
            Thread t = new Thread(delegate ()
            {
                try { job.Run(); } catch (Exception ex) { Core.Log("Repair thread: " + ex); }
                UI(delegate { Finished(job); });
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
        }

        void ReleaseMain()
        {
            if (mainMutex == null) return;
            try { mainMutex.ReleaseMutex(); } catch (Exception) { }
            mainMutex.Close();
            mainMutex = null;
        }

        static string Summary(RepairJob job)
        {
            string file = "\r\n\r\nהדוח נשמר בשולחן העבודה, בקובץ Milon_Repair_report.txt. שלחו אותו בצ'אט.";
            if (job.Error == "office") return "תוכנות אופיס נפתחו בינתיים, ולכן לא שיניתי כלום. סגרו אותן ולחצו שוב \"תקן\"." + file;
            if (job.OursMissing && job.Error == null)
                return "הרשימה של וורד נוקתה, אבל קובץ המילון התורני לא נמצא. פתחו את \"המילון התורני 3.0\" ולחצו \"התקן\"." + file;
            if (job.Ok) return "תוקן ✓\r\n\r\nהמילון התורני מחובר לוורד שוב, ובבדיקה בוורד חדש כל מילות הבדיקה זוהו." + file;
            return "התיקון לא הצליח" + (job.Error != null ? ": " + job.Error : ".") + "\r\n\r\nשום דבר לא נמחק מלבד הרשומות הישנות, ונשמר עותק של הרשימה." + file;
        }

        void Finished(RepairJob job)
        {
            running = false;
            ReleaseMain();
            UseWaitCursor = false;
            btnFix.Enabled = !job.Ok;
            body.Text = Summary(job);
            stage.Text = "";
            prog.Value = 100;
            try { Process.Start("explorer.exe", "/select,\"" + RepairJob.ReportPath + "\""); } catch (Exception) { }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (running && e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; return; }   // less than a minute: it is not stopped in the middle
            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e) { ReleaseMain(); base.OnFormClosed(e); }
    }

    static class RepairProgram
    {
        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length >= 2 && args[0] == "--selftest")
            {
                if (Type.GetType("Mono.Runtime") == null) { Console.WriteLine("The self-test runs only in the developer's environment (Mono)."); return 2; }
                return RepairTest.Run(args[1]);
            }
            int forcedDpi = 0;
            for (int i = 0; i + 1 < args.Length; i++) if (args[i] == "--dpi") int.TryParse(args[i + 1], out forcedDpi);
            Dpi.Init(forcedDpi);
            bool first;
            Mutex single = null;
            try { single = new Mutex(true, @"Local\MilonRepair.SingleInstance", out first); }
            catch (Exception) { first = true; }
            if (!first) return 0;
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                RepairForm f = new RepairForm();
                if (Array.IndexOf(args, "--preview-done") >= 0) f.PreviewDone();
                Application.Run(f);
            }
            catch (Exception ex)
            {
                Core.Log("Repair FATAL " + ex);
                try { MessageBox.Show("אירעה שגיאה: " + Core.Msg(ex), RepairForm.Title); } catch (Exception) { }
            }
            finally { if (single != null) { try { single.ReleaseMutex(); } catch (Exception) { } single.Close(); } }
            return 0;
        }
    }

    // ------------------------------------------------------------------------------------------ self-test: a pretend Word 365
    // It keeps its list where the user's Word does (as the log shows): "<n>_<account>" = the file, "<n>_<account>_state" (1 = in use,
    // 0 = taken out), "_roamed", "_external". Taking a dictionary out keeps its entry; a new file name takes a new number; with 10 entries
    // it refuses to add any dictionary (what the user's Word did).
    public class RWorld
    {
        public const string Account = "16_7bd116fa7f21d40b_LiveId";
        public class Entry { public int N; public string File; public bool On; }

        public List<Entry> Read()
        {
            List<Entry> l = new List<Entry>();
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath))
            {
                if (k == null) return l;
                foreach (string vn in k.GetValueNames())
                {
                    int n; string id; bool main;
                    if (!ListFix.Parse(vn, out n, out id, out main) || !main || id != Account) continue;
                    byte[] st = k.GetValue(vn + "_state") as byte[];
                    Entry e = new Entry(); e.N = n; e.File = Convert.ToString(k.GetValue(vn)); e.On = st != null && st.Length > 0 && st[0] == 1;
                    l.Add(e);
                }
            }
            l.Sort(delegate (Entry a, Entry b) { return a.N.CompareTo(b.N); });
            return l;
        }

        public void Write(List<Entry> l)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RegFallback.KeyPath))
            {
                foreach (string vn in k.GetValueNames())
                {
                    int n; string id; bool main;
                    if (ListFix.Parse(vn, out n, out id, out main) && id == Account) k.DeleteValue(vn, false);
                }
                foreach (Entry e in l)
                {
                    string head = e.N.ToString(CultureInfo.InvariantCulture) + "_" + Account;
                    k.SetValue(head, e.File, RegistryValueKind.String);
                    k.SetValue(head + "_state", new byte[] { (byte)(e.On ? 1 : 0), 0, 0, 0 }, RegistryValueKind.Binary);
                    k.SetValue(head + "_roamed", new byte[] { 0, 0, 0, 0 }, RegistryValueKind.Binary);
                    k.SetValue(head + "_external", new byte[] { 0, 0, 0, 0 }, RegistryValueKind.Binary);
                }
            }
        }
    }

    public class RApp
    {
        public RWorld World;
        public List<RWorld.Entry> List;
        public RDicts Dicts;
        public HashSet<string> Known = new HashSet<string>(StringComparer.Ordinal);
        public RApp(RWorld w)
        {
            World = w;
            List = w.Read();
            Dicts = new RDicts(); Dicts.App = this;
            foreach (RWorld.Entry e in List)
            {
                if (!e.On) continue;
                string p = Path.IsPathRooted(e.File) ? e.File : Path.Combine(Core.UProofDir(), e.File);
                if (File.Exists(p)) Known.UnionWith(Core.ReadDicWords(p));
            }
        }
        public RDicts CustomDictionaries { get { return Dicts; } }
        public RDocs Documents { get { RDocs d = new RDocs(); d.App = this; return d; } }
        public void Quit(params object[] args) { World.Write(List); }
        public bool Marks(string word)
        {
            if (word == "קזקזקזק") return true;
            if (word == "שלום") return false;
            return !Known.Contains(word);
        }
    }

    public class RDicts
    {
        public RApp App;
        public RWorld.Entry Active;
        List<RDict> On()
        {
            List<RDict> l = new List<RDict>();
            foreach (RWorld.Entry e in App.List) if (e.On) l.Add(new RDict(this, e));
            return l;
        }
        public int Count { get { return On().Count; } }
        public RDict Item(int i) { return On()[i - 1]; }
        public RDict ActiveCustomDictionary
        {
            get { List<RDict> l = On(); if (Active != null && Active.On) return new RDict(this, Active); return l.Count > 0 ? l[0] : null; }
            set { Active = value == null ? null : value.E; }
        }
        public RDict Add(string path)
        {
            if (App.List.Count >= 10) throw new InvalidOperationException("Unspecified error (E_FAIL)");
            string uproof = Core.UProofDir();
            string file = string.Equals(Path.GetDirectoryName(path), uproof, StringComparison.OrdinalIgnoreCase) ? Path.GetFileName(path) : path;
            foreach (RWorld.Entry e in App.List)
                if (string.Equals(Path.GetFileName(e.File), Path.GetFileName(path), StringComparison.OrdinalIgnoreCase)) { e.On = true; return new RDict(this, e); }
            RWorld.Entry ne = new RWorld.Entry();
            ne.N = App.List.Count == 0 ? 1 : App.List[App.List.Count - 1].N + 1;
            ne.File = file; ne.On = true;
            App.List.Add(ne);
            return new RDict(this, ne);
        }
    }

    public class RDict
    {
        readonly RDicts owner;
        public readonly RWorld.Entry E;
        public RDict(RDicts owner, RWorld.Entry e) { this.owner = owner; E = e; }
        public string Name { get { return System.IO.Path.GetFileName(E.File); } }
        public string Path { get { string d = System.IO.Path.GetDirectoryName(E.File); return string.IsNullOrEmpty(d) ? Core.UProofDir() : d; } }
        public bool LanguageSpecific { get; set; }
        public void Delete() { E.On = false; if (owner.Active == E) owner.Active = null; }
    }

    public class RDocs { public RApp App; public RDoc Add(params object[] args) { RDoc d = new RDoc(); d.App = App; return d; } }
    public class RDoc
    {
        public RApp App;
        RRange r;
        public RRange Content { get { if (r == null) { r = new RRange(); r.App = App; } return r; } }
        public void Close(params object[] args) { }
    }
    public class RRange
    {
        public RApp App;
        string text = "";
        public string Text { get { return text; } set { text = value ?? ""; } }
        public int LanguageID { get; set; }
        public bool NoProofing { get; set; }
        public RErrors SpellingErrors { get { RErrors e = new RErrors(); foreach (string w in text.Split(' ')) if (w.Length > 0 && App.Marks(w)) e.Count++; return e; } }
    }
    public class RErrors { public int Count { get; set; } }

    static class RepairTest
    {
        static int fails;
        static void Check(bool ok, string what) { Console.WriteLine((ok ? "  ok    " : "  FAIL  ") + what); if (!ok) fails++; }

        static string Names(List<RWorld.Entry> l)
        {
            List<string> s = new List<string>();
            foreach (RWorld.Entry e in l) s.Add(e.N + ":" + Path.GetFileName(e.File) + (e.On ? "" : "(out)"));
            return string.Join(" ", s.ToArray());
        }

        // the list as the user's log shows it on 9-11-2026 (state bytes made up: 1 = in use)
        static void UsersList(RWorld w, string dir)
        {
            Registry.CurrentUser.DeleteSubKeyTree(RegFallback.KeyPath, false);
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RegFallback.KeyPath))
            {
                k.SetValue("UpdateComplete", 1, RegistryValueKind.DWord);
                k.SetValue("1_state", new byte[] { 1, 0, 0, 0 }, RegistryValueKind.Binary);
                k.SetValue("1_external", new byte[] { 0, 0, 0, 0 }, RegistryValueKind.Binary);
                k.SetValue("11", Path.Combine(dir, Core.MainFile), RegistryValueKind.String);   // the program's own registry fallback
            }
            string roaming = Path.Combine(Path.Combine(dir, "Proofing"), "RoamingCustom.dic");
            List<RWorld.Entry> l = new List<RWorld.Entry>();
            string[] files = { roaming, "Otzaria_Torah_Basic.dic", "Otzaria_Torah_Acronyms.dic", Core.MainFile, Path.Combine(dir, "wordcheck_trial.dic"),
                               "Milon_SizeTest_175000.dic", "Milon_SizeTest_200000.dic", "Milon_SizeTest_250000.dic", "Milon_SizeTest_300000.dic", "Milon_SizeTest_400000.dic" };
            for (int i = 0; i < files.Length; i++) { RWorld.Entry e = new RWorld.Entry(); e.N = i + 1; e.File = files[i]; e.On = i == 0; l.Add(e); }
            w.Write(l);
        }

        public static int Run(string d)
        {
            try { Console.OutputEncoding = new UTF8Encoding(false); } catch (Exception) { }
            string dir = Path.GetFullPath(d);
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
            Directory.CreateDirectory(dir);
            Directory.CreateDirectory(Path.Combine(dir, "Proofing"));
            Core.UProofOverride = dir;
            Core.LogPath = Path.Combine(dir, "log.txt");
            MyWords.FolderOverride = Path.Combine(dir, "settings");
            RepairJob.StateDir = Path.Combine(dir, "state");
            RepairJob.ReportPath = Path.Combine(dir, "report.txt");
            RepairJob.SizeCheckDir = Path.Combine(dir, "sizecheck");
            RepairJob.OfficeHook = delegate { return new List<string>(); };
            RWorld world = new RWorld();
            WordCom.OpenNewHook = delegate { WordSession s = new WordSession(); s.App = new RApp(world); s.Created = true; return s; };
            WordCom.OpenHook = WordCom.OpenNewHook;
            string ours = Path.Combine(dir, Core.MainFile);
            Core.WriteDic(ours, Core.BuildWords(true, true));
            Core.WriteDic(Path.Combine(Path.Combine(dir, "Proofing"), "RoamingCustom.dic"), new List<string> { "מילהשלי" });

            Console.WriteLine("-- reading a value name");
            int n; string id; bool main;
            Check(ListFix.Parse("4_16_7bd116fa7f21d40b_LiveId", out n, out id, out main) && n == 4 && id == RWorld.Account && main, "4_<account> = the file of entry 4");
            Check(ListFix.Parse("10_16_7bd116fa7f21d40b_LiveId_roamed", out n, out id, out main) && n == 10 && id == RWorld.Account && !main, "10_<account>_roamed = a setting of entry 10");
            Check(ListFix.Parse("1_state", out n, out id, out main) && n == 1 && id == "" && !main, "1_state = a setting of the entry without an account");
            Check(ListFix.Parse("11", out n, out id, out main) && n == 11 && id == "" && main, "11 = an entry without an account");
            Check(!ListFix.Parse("UpdateComplete", out n, out id, out main), "UpdateComplete is not an entry");

            Console.WriteLine("-- the user's list of 9-11: 10 entries, Word refuses every dictionary");
            UsersList(world, dir);
            bool refused = false;
            try { using (WordSession s = WordCom.OpenNew()) WordCom.Add(s, ours); } catch (Exception) { refused = true; }
            Check(refused, "before the repair the pretend Word refuses, as the real one did (" + Names(world.Read()) + ")");
            Directory.CreateDirectory(RepairJob.SizeCheckDir);
            File.WriteAllText(Path.Combine(RepairJob.SizeCheckDir, "restore.txt"), "version=1\r\n");
            RepairJob job = new RepairJob();
            job.Run();
            string rep = File.ReadAllText(RepairJob.ReportPath);
            List<RWorld.Entry> after = world.Read();
            Check(job.Ok && job.Verdict == 1 && job.Error == null, "repaired, and a new Word knows the words (" + job.Error + ")");
            Check(Names(after) == "1:RoamingCustom.dic 2:Milon_Torani.dic", "the list: the person's dictionary and the Torah dictionary (" + Names(after) + ")");
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath))
            {
                Check(k.GetValue("UpdateComplete") != null && k.GetValue("1_state") != null && k.GetValue("1_external") != null && k.GetValue("11") == null,
                      "Office's own values stay; the program's old fallback entry (11) is gone");
            }
            Check(rep.Contains("#BEFORE") && rep.Contains("4_16_7bd116fa7f21d40b_LiveId = Milon_Torani.dic") && rep.Contains("_state = hex:01000000"), "the report has every value before the change");
            Check(rep.Contains("#CLEAN deleted 6 (16_7bd116fa7f21d40b_LiveId) Milon_SizeTest_175000.dic") && rep.Contains("#ADD ok") && rep.Contains("#DONE ok"), "the report says what was deleted and that Word took the dictionary");
            Check(!File.Exists(Path.Combine(RepairJob.SizeCheckDir, "restore.txt")), "the size check's restore mark is gone (it would try again)");

            Console.WriteLine("-- a dictionary of the person after the old entries: it stays, and the numbers are closed up");
            UsersList(world, dir);
            List<RWorld.Entry> mine = world.Read();
            mine.RemoveAll(delegate (RWorld.Entry e) { return e.N > 6; });
            RWorld.Entry own = new RWorld.Entry(); own.N = 7; own.File = "MyOwn.dic"; own.On = true; mine.Add(own);
            world.Write(mine);
            Core.WriteDic(Path.Combine(dir, "MyOwn.dic"), new List<string> { "מילהאחרת" });
            RepairJob j2 = new RepairJob();
            j2.Run();
            after = world.Read();
            Check(j2.Ok && Names(after) == "1:RoamingCustom.dic 2:MyOwn.dic 3:Milon_Torani.dic", "the person's dictionary moved from 7 to 2 and is still in use (" + Names(after) + ")");
            Check(File.ReadAllText(RepairJob.ReportPath).Contains("#CLEAN moved 7 -> 2"), "the move is in the report");

            Console.WriteLine("-- an Office program is open: nothing is changed");
            UsersList(world, dir);
            string beforeOpen = Names(world.Read());
            RepairJob.OfficeHook = delegate { return new List<string> { "OUTLOOK" }; };
            RepairJob j3 = new RepairJob();
            j3.Run();
            Check(j3.Error == "office" && !j3.Ok && Names(world.Read()) == beforeOpen, "stopped before any change");
            RepairJob.OfficeHook = delegate { return new List<string>(); };

            Console.WriteLine("-- the dictionary file is missing: the list is cleaned, and the person is told to install");
            UsersList(world, dir);
            File.Move(ours, ours + ".away");
            RepairJob j4 = new RepairJob();
            j4.Run();
            File.Move(ours + ".away", ours);
            Check(j4.OursMissing && j4.Ok && Names(world.Read()) == "1:RoamingCustom.dic", "cleaned; the program's install is the next step (" + Names(world.Read()) + ")");

            Console.WriteLine(fails == 0 ? "SELFTEST PASSED" : "SELFTEST FAILED (" + fails + ")");
            return fails == 0 ? 0 : 1;
        }
    }
}
