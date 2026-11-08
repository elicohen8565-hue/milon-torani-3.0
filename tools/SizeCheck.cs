using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

// Size check - a developer tool (3-11-2026), not part of the program: how many words does Word take in ONE custom dictionary, and how much
// slower does it get? The person runs it once on his own computer. For every size it puts a test dictionary in place of the Torah dictionary
// (the 172,531 words of the extended edition with "chibur torani", then made-up words up to the size), starts a new hidden Word and asks it
// about words from the start, the middle and the very end of the file, and how long a block of known words and one of unknown words take.
// Sizes 175,000 ... 1,000,000; after the first size that fails, three more sizes narrow the limit down.
// At the end - also when the check stops in the middle, and at the next start if the computer went off - the Torah dictionary is put back
// as it was (with Word's default dictionary), and a new Word checks that it knows the words again.
// The results: Milon_SizeCheck_results.txt on the desktop (the person sends it back). One line per size, "SIZE<TAB>n<TAB>key=value...".
// Built from the program's own code (Core.cs, UI.cs, Words.cs) by tools/build_sizecheck.sh; self-test: --selftest <folder> (Mono only).
namespace OtzariaDict
{
    public class SizeResult
    {
        public int Size;
        public string Add = "-";      // "ok", or "failed:<why>" = Word itself would not add the file
        public string Via = "-";      // how the file got into Word's list: "word", or "registry" (the list Word reads when it starts)
        public bool Listed;           // the new Word lists the test dictionary
        public long StartMs = -1, FirstCheckMs = -1;
        public bool ControlsOk;       // nonsense is marked and an ordinary word is not: the check itself works
        public int BaseOk, BaseAll, MidOk, MidAll, EndOk, EndAll;
        public long OneMs = -1;       // one unknown word
        public int InN, OutN;         // words in the timed blocks: 50, fewer when Word is very slow (a block must not take minutes)
        public long InMs = -1, OutMs = -1;
        public int InMarked = -1, OutMarked = -1;
        public bool Works;            // the whole file is in effect, from its first word to its last
        public bool Stopped;          // the person stopped the check in the middle of this size
        public string Error;
    }

    public class SizeJob
    {
        public static readonly int[] Sizes = { 0, 175000, 200000, 250000, 300000, 400000, 500000, 750000, 1000000 };
        public const int NarrowSteps = 3;
        public const string TestPrefix = "Milon_SizeTest_";
        public const string NotAWord = "קזקזקזק";      // Word must mark it, otherwise the check itself does not work
        public const string CommonWord = "שלום";        // Word must not mark it
        const string Letters = "אבגדהוזחטיכלמנסעפצקרשת";   // 22 letters, no final forms
        const int Block = 50;                           // words in a timed block (fewer when one unknown word takes over a second)
        const int OutStart = 2000000;                   // the unknown block: made-up words that no test file has (the biggest needs 827,469)
        static readonly string[] Samples = { "מיבעי", "באנפא", "דבוקין", "דהעניין", "דמסתברא", "החששא", "ויוצרא", "וסרכא",
                                             "כדסבירא", "לוחתא", "סמיותא", "תכסיסא", "ודהנושא", "דהמכתב", "מג\"א" };   // as in Verify

        public static string StateDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "MilonSizeCheck");
        public static string ResultPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), "Milon_SizeCheck_results.txt");
        public static bool ResultsOnDesktop = true;
        public static Func<int> RunningWordsHook;      // the self-test: a pretend count of the person's open Words
        public static int PollMs = 2000;
        static string MarkerPath { get { return Path.Combine(StateDir, "restore.txt"); } }

        public Action<int, string> Progress = delegate { };
        public volatile bool StopRequested;
        public readonly List<SizeResult> Results = new List<SizeResult>();
        public bool RestoredOk, WordWasOpened, StoppedByPerson;
        public bool Changed;              // Word's list was changed (and so had to be put back)
        public string Error;
        public int LimitOk = -1, LimitBad = -1;
        public int ErrorAt = -1;          // the size at which Word stopped answering (an error), -1 = none
        public bool ControlsFailed;       // Word does not check Hebrew spelling here: nothing can be measured

        long lastPing = DateTime.Now.Ticks;
        public DateTime LastProgress { get { return new DateTime(Interlocked.Read(ref lastPing)); } }
        public void Ping() { Interlocked.Exchange(ref lastPing, DateTime.Now.Ticks); }

        readonly List<int> pids = new List<int>();
        readonly List<string> oursPaths = new List<string>();     // the Torah dictionary files that Word listed before the check
        readonly List<string> namesBefore = new List<string>();   // every dictionary Word listed before the check
        string defaultName;                                        // Word's default custom dictionary before the check
        List<string> baseWords, baseSamples;

        // a made-up word that is surely not Hebrew: "זקז" and five letters (22^5 = 5,153,632 different ones)
        public static string Synthetic(int i)
        {
            char[] c = new char[8];
            c[0] = 'ז'; c[1] = 'ק'; c[2] = 'ז';
            for (int k = 7; k >= 3; k--) { c[k] = Letters[i % Letters.Length]; i /= Letters.Length; }
            return new string(c);
        }

        public static bool Pending() { return File.Exists(MarkerPath); }
        public static bool IsTest(string name) { return name != null && name.StartsWith(TestPrefix, StringComparison.OrdinalIgnoreCase); }
        static string N(int n) { return n.ToString("N0", CultureInfo.CurrentCulture); }
        static string OneLine(string s) { return (s ?? "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " "); }

        void Say(int percent, string text) { Ping(); Core.Log("Size check: " + text); try { Progress(percent, text); } catch (Exception) { } }

        static int RunningWords() { return RunningWordsHook != null ? RunningWordsHook() : WordCom.RunningWordCount(); }

        // ---------------------------------------------------------------- the results file
        void StartResults()
        {
            try { File.WriteAllText(ResultPath, "", new UTF8Encoding(false)); return; }
            catch (Exception ex) { Core.Log("Size check, results on the desktop: " + Core.Msg(ex)); }
            ResultPath = Path.Combine(StateDir, "Milon_SizeCheck_results.txt");   // the desktop cannot be written: the check's own folder
            ResultsOnDesktop = false;
            try { File.WriteAllText(ResultPath, "", new UTF8Encoding(false)); } catch (Exception) { }
        }

        void Out(string line)
        {
            try { File.AppendAllText(ResultPath, line + "\r\n", new UTF8Encoding(false)); }
            catch (Exception ex) { Core.Log("Size check, result file: " + Core.Msg(ex)); }
        }

        // ---------------------------------------------------------------- the hidden Words this check starts (to stop one that hangs)
        static HashSet<int> WinwordPids()
        {
            HashSet<int> s = new HashSet<int>();
            try { foreach (Process p in Process.GetProcessesByName("WINWORD")) { try { s.Add(p.Id); } finally { p.Dispose(); } } }
            catch (Exception) { }
            return s;
        }

        WordSession OpenTracked()
        {
            HashSet<int> before = WinwordPids();
            WordSession s = WordCom.OpenNew();
            foreach (int p in WinwordPids()) if (!before.Contains(p)) lock (pids) pids.Add(p);
            Ping();
            return s;
        }

        // only processes this check started, and only while they are still Word (a number can be reused after a process ends)
        public void KillOurWords()
        {
            List<int> mine;
            lock (pids) mine = new List<int>(pids);
            foreach (int pid in mine)
            {
                try
                {
                    using (Process p = Process.GetProcessById(pid))
                        if (!p.HasExited && string.Equals(p.ProcessName, "WINWORD", StringComparison.OrdinalIgnoreCase)) { p.Kill(); Core.Log("Size check: stopped a hidden Word that did not answer, " + pid); }
                }
                catch (Exception) { }
            }
        }

        // ---------------------------------------------------------------- Word, asked about words
        static object Doc(WordSession s)
        {
            object docs = WordCom.Get(s.App, "Documents");
            return WordCom.Call(docs, "Add", Missing.Value, Missing.Value, Missing.Value, false);   // an invisible document
        }

        static int CountMarked(object content, string text)
        {
            WordCom.Put(content, "Text", text);
            WordCom.Put(content, "LanguageID", 1037);   // Hebrew
            try { WordCom.Put(content, "NoProofing", false); } catch (Exception) { }
            object errs = WordCom.Get(content, "SpellingErrors");
            return Convert.ToInt32(WordCom.Get(errs, "Count"));
        }

        static bool Marked(object content, string word) { return CountMarked(content, word) > 0; }

        static int RemoveTests(WordSession s)
        {
            int n = 0;
            object dicts = WordCom.Get(s.App, "CustomDictionaries");
            int count = Convert.ToInt32(WordCom.Get(dicts, "Count"));
            for (int i = count; i >= 1; i--)
            {
                object d = WordCom.Item(dicts, i);
                if (IsTest(Convert.ToString(WordCom.Get(d, "Name")))) { WordCom.Call(d, "Delete"); n++; }
            }
            return n;
        }

        static string ActiveName(WordSession s)
        {
            try
            {
                object a = WordCom.Get(WordCom.Get(s.App, "CustomDictionaries"), "ActiveCustomDictionary");
                return a == null ? null : Convert.ToString(WordCom.Get(a, "Name"));
            }
            catch (Exception) { return null; }
        }

        static void SetDefault(WordSession s, string name)
        {
            try
            {
                object dicts = WordCom.Get(s.App, "CustomDictionaries");
                int n = Convert.ToInt32(WordCom.Get(dicts, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object d = WordCom.Item(dicts, i);
                    if (string.Equals(Convert.ToString(WordCom.Get(d, "Name")), name, StringComparison.OrdinalIgnoreCase)) { WordCom.Put(dicts, "ActiveCustomDictionary", d); return; }
                }
            }
            catch (Exception ex) { Core.Log("Size check, default dictionary: " + Core.Msg(ex)); }
        }

        // the list Word reads when it starts: entries of test dictionaries go (a Word that was stopped by force could not take them out)
        static void CleanRegistry()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath, true))
                {
                    if (k == null) return;
                    foreach (string vn in k.GetValueNames())
                    {
                        if (vn.IndexOf('_') >= 0) continue;
                        string v = k.GetValue(vn) as string, leaf = null;
                        try { leaf = v == null ? null : Path.GetFileName(v); } catch (Exception) { }
                        if (!IsTest(leaf)) continue;
                        k.DeleteValue(vn, false);
                        foreach (string other in k.GetValueNames())
                            if (other.StartsWith(vn + "_", StringComparison.OrdinalIgnoreCase)) k.DeleteValue(other, false);
                        Core.Log("Size check: removed from the registry list " + v);
                    }
                }
            }
            catch (Exception ex) { Core.Log("Size check, registry: " + Core.Msg(ex)); }
        }

        static void DeleteTestFilesExcept(string keep)
        {
            try
            {
                foreach (string f in Directory.GetFiles(Core.UProofDir(), TestPrefix + "*.dic"))
                {
                    if (keep != null && string.Equals(Path.GetFileName(f), Path.GetFileName(keep), StringComparison.OrdinalIgnoreCase)) continue;
                    try { File.Delete(f); } catch (Exception ex) { Core.Log("Size check, delete " + f + ": " + Core.Msg(ex)); }
                }
            }
            catch (Exception) { }
        }

        // the person opened Word: the Torah dictionary can be put back only when every Word is closed (an open Word could write its old list over it)
        void WaitUntilWordClosed()
        {
            int waited = 0;
            while (RunningWords() > 0 && waited < 30 * 60 * 1000)
            {
                if (waited == 0) Out("#WAIT Word is open");
                Say(-1, "וורד פתוח. סגרו את כל חלונות וורד, ואז המילון התורני יחזור אליו.");
                Thread.Sleep(PollMs);
                waited += PollMs;
            }
        }

        // ---------------------------------------------------------------- the run
        public void Run()
        {
            DateTime t0 = DateTime.Now;
            WordCom.RetryMs = 120000;   // a Word busy with a very big dictionary may say "busy" for longer than the program's usual 15 seconds
            try
            {
                Directory.CreateDirectory(StateDir);
                StartResults();
                Out("#SIZECHECK 1");
                Out("#DATE " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                Out("#PROGRAM " + Core.AppVersion);
                if (Pending())
                {
                    Say(1, "מחזיר את המילון התורני לוורד (הבדיקה הקודמת נעצרה באמצע)…");
                    ReadMarker();
                    Out("#REPAIR_AT_START");
                    Changed = true;
                    WaitUntilWordClosed();
                    if (!Restore()) { Error = "לא הצלחתי להחזיר את המילון התורני לוורד."; return; }
                }
                if (StopRequested) return;   // only putting things back was asked for
                if (RunningWords() > 0) { Error = "וורד פתוח. סגרו את וורד ונסו שוב."; Out("#STOP Word is open"); return; }

                Say(2, "בודק את וורד ואת המילונים שלו…");
                ReadWordsList();
                WriteMarker();   // from here on, the Torah dictionary is put back whatever happens
                Changed = true;

                baseWords = Core.BuildWords(true, true);
                HashSet<string> inBase = new HashSet<string>(baseWords, StringComparer.Ordinal);
                baseSamples = new List<string>();
                foreach (string w in Samples) if (inBase.Contains(w)) baseSamples.Add(w);
                Out("#BASE " + baseWords.Count + " samples=" + baseSamples.Count);
                TestAllSizes();
            }
            catch (Exception ex)
            {
                Error = ex is WordMissingException ? "לא מצאתי את וורד במחשב." : Core.Msg(ex);
                Core.Log("Size check failed: " + ex);
                Out("#ERROR " + OneLine(Core.Msg(ex)));
            }
            finally
            {
                if (Pending())
                {
                    try
                    {
                        WaitUntilWordClosed();
                        Say(92, "מחזיר את המילון התורני לוורד ובודק שוורד מכיר אותו…");
                        Restore();
                    }
                    catch (Exception ex) { Core.Log("Size check restore: " + ex); Out("#RESTORED FAILED " + OneLine(Core.Msg(ex))); }
                }
                RestoredOk = !Pending();
                Out("#ELAPSED " + (int)(DateTime.Now - t0).TotalSeconds + " s");
                Out("#DONE");
                Say(100, "הסתיים");
            }
        }

        void ReadWordsList()
        {
            using (WordSession w = OpenTracked())
            {
                try { Out("#WORD " + Convert.ToString(WordCom.Get(w.App, "Version")) + " build " + Convert.ToString(WordCom.Get(w.App, "Build"))); }
                catch (Exception) { }
                StringBuilder ds = new StringBuilder();
                foreach (DictInfo di in WordCom.ListAll(w))
                {
                    string count = "?";
                    try { if (di.FullPath != null && File.Exists(di.FullPath)) count = Core.CountDicWords(di.FullPath).ToString(CultureInfo.InvariantCulture); }
                    catch (Exception) { }
                    ds.Append((ds.Length > 0 ? " | " : "") + di.Name + " (" + count + ")");
                    if (IsTest(di.Name) || di.FullPath == null) continue;   // a test left by a check stopped by force goes; a missing file is not checked
                    namesBefore.Add(di.Name);
                    if (Core.IsOurs(di.Name)) oursPaths.Add(di.FullPath);
                }
                defaultName = ActiveName(w);
                Out("#DICTS " + ds);
                Out("#DEFAULT " + (defaultName ?? "-"));
            }
            WordCom.WaitForOurWordsToExit(15000);
            Out("#OURS " + (oursPaths.Count > 0 ? "listed" : "not listed"));
        }

        void TestAllSizes()
        {
            int steps = Sizes.Length + NarrowSteps, step = 0, ok = 0, bad = -1;
            bool broke = false;
            foreach (int n in Sizes)
            {
                if (!CanGoOn()) { broke = true; break; }
                SizeResult r = TestAndWrite(n, step++, steps);
                if (r.Stopped) { StoppedByPerson = true; Out("#STOPPED by the person"); broke = true; break; }
                if (r.Error != null) { Out("#STOP error at " + n); ErrorAt = n; broke = true; if (n > 0) bad = n; break; }
                if (n == 0)
                {
                    if (!r.ControlsOk) { ControlsFailed = true; Out("#STOP the check itself does not work (controls)"); broke = true; break; }
                    continue;
                }
                if (r.Works) ok = n; else { bad = n; break; }   // a bigger file will not work either
            }
            // the limit is between the last size that worked and the first that did not: three more sizes narrow it down
            if (!broke && bad > 0 && ok >= Sizes[1])
            {
                for (int k = 0; k < NarrowSteps; k++)
                {
                    if (bad - ok <= 10000) break;
                    int m = (ok + bad) / 2 / 1000 * 1000;
                    if (m <= ok || m >= bad) break;
                    if (!CanGoOn()) break;
                    SizeResult r = TestAndWrite(m, step++, steps);
                    if (r.Stopped) { StoppedByPerson = true; Out("#STOPPED by the person"); break; }
                    if (r.Error != null) { Out("#STOP error at " + m); ErrorAt = m; break; }
                    if (r.Works) ok = m; else bad = m;
                }
            }
            LimitOk = ok;
            LimitBad = bad;
            Out("#LIMIT works=" + ok + " fails=" + (bad > 0 ? bad.ToString(CultureInfo.InvariantCulture) : "none") + (ErrorAt > 0 ? " error_at=" + ErrorAt : ""));
        }

        bool CanGoOn()
        {
            if (StopRequested) { StoppedByPerson = true; Out("#STOPPED by the person"); return false; }
            if (RunningWords() > 0) { WordWasOpened = true; Out("#STOP Word was opened during the check"); return false; }
            return true;
        }

        SizeResult TestAndWrite(int n, int step, int steps)
        {
            SizeResult r = new SizeResult();
            r.Size = n;
            try { TestSize(r, 5 + 85 * step / steps); }
            catch (Exception ex) { r.Error = OneLine(Core.Msg(ex)); r.Works = false; Core.Log("Size check, size " + n + ": " + ex); }
            Results.Add(r);
            Out(string.Format(CultureInfo.InvariantCulture,
                "SIZE\t{0}\tadd={1}\tvia={2}\tlisted={3}\tstart_ms={4}\tfirst_ms={5}\tcontrols={6}\tbase={7}/{8}\tmid={9}/{10}\tend={11}/{12}" +
                "\tone_ms={13}\tin_n={14}\tin_ms={15}\tin_marked={16}\tout_n={17}\tout_ms={18}\tout_marked={19}\tworks={20}{21}{22}",
                n, r.Add, r.Via, r.Listed ? 1 : 0, r.StartMs, r.FirstCheckMs, r.ControlsOk ? "ok" : "BAD", r.BaseOk, r.BaseAll, r.MidOk, r.MidAll,
                r.EndOk, r.EndAll, r.OneMs, r.InN, r.InMs, r.InMarked, r.OutN, r.OutMs, r.OutMarked, r.Works ? 1 : 0, r.Stopped ? "\tstopped=1" : "",
                r.Error != null ? "\terror=" + r.Error : ""));
            return r;
        }

        void TestSize(SizeResult r, int pct)
        {
            int n = r.Size, synth = 0;
            string path = null;
            if (n > 0)
            {
                synth = Math.Max(0, n - baseWords.Count);
                Say(pct, "מכין מילון בדיקה של " + N(n) + " מילים…");
                List<string> words = new List<string>(n);
                words.AddRange(n >= baseWords.Count ? baseWords : baseWords.GetRange(0, n));
                for (int i = 0; i < synth; i++) words.Add(Synthetic(i));
                Directory.CreateDirectory(Core.UProofDir());
                path = Path.Combine(Core.UProofDir(), TestPrefix + n.ToString(CultureInfo.InvariantCulture) + ".dic");
                Core.WriteDic(path, words);
            }

            // 1. Word gets the test dictionary in place of the Torah dictionary (a Word reads its dictionaries only when it starts)
            Say(pct + 1, n > 0 ? "מכניס לוורד את מילון הבדיקה (" + N(n) + " מילים)…" : "מודד את וורד בלי המילון התורני…");
            using (WordSession w = OpenTracked())
            {
                RemoveTests(w);
                WordCom.RemoveOurs(w);
                if (path != null)
                {
                    try { WordCom.Add(w, path); r.Add = "ok"; r.Via = "word"; }
                    catch (Exception ex) { r.Add = "failed:" + OneLine(Core.Msg(ex)); }
                }
                Ping();
            }
            WordCom.WaitForOurWordsToExit(15000);
            DeleteTestFilesExcept(path);   // the files of the sizes before: no longer in Word's list
            if (path != null && r.Add != "ok")
            {
                // Word itself would not take the file: the list that Word reads when it starts gets it (as the program's own fallback does)
                try { RegFallback.Add(path); r.Via = "registry"; }
                catch (Exception ex) { r.Via = "none"; Core.Log("Size check, registry: " + Core.Msg(ex)); return; }
            }
            if (StopRequested) { r.Stopped = true; return; }

            // 2. a NEW Word: does it know the words of the file, from its start to its very end - and how fast is it?
            Say(pct + 3, n > 0 ? "בודק בוורד חדש את מילון הבדיקה (" + N(n) + " מילים)…" : "מודד את וורד בלי המילון התורני…");
            Stopwatch sw = Stopwatch.StartNew();
            using (WordSession w = OpenTracked())
            {
                r.StartMs = sw.ElapsedMilliseconds;
                r.Listed = path == null || WordCom.IsListed(w, Path.GetFileName(path));
                object doc = Doc(w);
                try
                {
                    object content = WordCom.Get(doc, "Content");
                    sw.Reset(); sw.Start();
                    bool nonsense = Marked(content, NotAWord);
                    bool common = Marked(content, CommonWord);
                    r.FirstCheckMs = sw.ElapsedMilliseconds;
                    r.ControlsOk = nonsense && !common;
                    Ping();
                    r.BaseAll = baseSamples.Count;
                    foreach (string s in baseSamples) { if (!Marked(content, s)) r.BaseOk++; Ping(); }
                    if (StopRequested) { r.Stopped = true; return; }
                    sw.Reset(); sw.Start();
                    Marked(content, Synthetic(OutStart + 1000));   // one unknown word: how long does it take?
                    r.OneMs = sw.ElapsedMilliseconds;
                    Ping();
                    int bn = r.OneMs <= 1000 ? Block : (int)Math.Max(5, Math.Min(Block, 60000 / Math.Max(1, r.OneMs)));   // a block: about a minute at most
                    if (synth > 0)
                    {
                        r.MidAll = Math.Min(10, synth);
                        int from = Math.Max(0, synth / 2 - r.MidAll / 2);
                        for (int k = 0; k < r.MidAll; k++) { if (!Marked(content, Synthetic(from + k))) r.MidOk++; Ping(); }
                        r.EndAll = Math.Min(10, synth);
                        for (int k = synth - r.EndAll; k < synth; k++) { if (!Marked(content, Synthetic(k))) r.EndOk++; Ping(); }
                        List<string> inBlock = new List<string>();
                        for (int k = 0; k < bn; k++) inBlock.Add(Synthetic((int)((long)k * synth / bn)));   // spread over the whole file
                        r.InN = bn;
                        sw.Reset(); sw.Start();
                        r.InMarked = CountMarked(content, string.Join(" ", inBlock.ToArray()));
                        r.InMs = sw.ElapsedMilliseconds;
                        Ping();
                    }
                    if (StopRequested) { r.Stopped = true; return; }
                    List<string> outBlock = new List<string>();
                    for (int k = 0; k < bn; k++) outBlock.Add(Synthetic(OutStart + k));
                    r.OutN = bn;
                    sw.Reset(); sw.Start();
                    r.OutMarked = CountMarked(content, string.Join(" ", outBlock.ToArray()));
                    r.OutMs = sw.ElapsedMilliseconds;
                    Ping();
                }
                finally { try { WordCom.Call(doc, "Close", 0); } catch (Exception) { } }
            }
            WordCom.WaitForOurWordsToExit(15000);
            // works = Word knows the words of the file from its start (the base samples) to its very end (one miss in each group is let pass)
            r.Works = n == 0 || (r.Listed && r.ControlsOk && r.BaseOk >= r.BaseAll - 1 && r.MidOk >= r.MidAll - 1 && r.EndOk >= r.EndAll - 1);
        }

        // ---------------------------------------------------------------- putting the Torah dictionary back
        void WriteMarker()
        {
            StringBuilder sb = new StringBuilder("version=1\r\n");
            foreach (string p in oursPaths) sb.Append("ours=" + p + "\r\n");
            if (defaultName != null) sb.Append("default=" + defaultName + "\r\n");
            foreach (string nm in namesBefore) sb.Append("had=" + nm + "\r\n");
            Directory.CreateDirectory(StateDir);
            File.WriteAllText(MarkerPath, sb.ToString(), new UTF8Encoding(false));
        }

        void ReadMarker()
        {
            oursPaths.Clear(); namesBefore.Clear(); defaultName = null;
            try
            {
                foreach (string line in File.ReadAllLines(MarkerPath, Encoding.UTF8))
                {
                    if (line.StartsWith("ours=") && line.Length > 5) oursPaths.Add(line.Substring(5).Trim());
                    else if (line.StartsWith("default=") && line.Length > 8) defaultName = line.Substring(8).Trim();
                    else if (line.StartsWith("had=") && line.Length > 4) namesBefore.Add(line.Substring(4).Trim());
                }
            }
            catch (Exception ex)
            {
                // the mark cannot be read: when in doubt, the Torah dictionary goes back (only if its file is there)
                Core.Log("Size check, marker: " + Core.Msg(ex));
                string p = Path.Combine(Core.UProofDir(), Core.MainFile);
                if (File.Exists(p)) oursPaths.Add(p);
            }
        }

        // true = Word lists no test dictionary, every dictionary of before is back, and a new Word knows the words of the Torah dictionary
        bool Restore()
        {
            WordCom.WaitForOurWordsToExit(15000);
            KillOurWords();   // a hidden Word of this check that is still running is stuck: it must not write its list over ours later
            int removed;
            List<string> viaRegistry = new List<string>();
            using (WordSession w = OpenTracked())
            {
                removed = RemoveTests(w);
                string folder = null;
                foreach (string p in oursPaths)
                {
                    if (!File.Exists(p)) continue;
                    if (!WordCom.IsListed(w, Path.GetFileName(p)))
                    {
                        try { WordCom.Add(w, p); }
                        catch (Exception ex) { Core.Log("Size check, Word would not add " + p + ": " + Core.Msg(ex)); viaRegistry.Add(p); }
                    }
                    if (folder == null) folder = Path.GetDirectoryName(p);
                }
                if (defaultName != null && !Core.IsOurs(defaultName) && !IsTest(defaultName)) SetDefault(w, defaultName);
                if (folder != null) WordCom.KeepOursNotDefault(w, folder);
            }
            WordCom.WaitForOurWordsToExit(15000);
            CleanRegistry();
            foreach (string p in viaRegistry) { try { RegFallback.Add(p); } catch (Exception ex) { Core.Log("Size check, registry: " + Core.Msg(ex)); } }

            bool ok;
            using (WordSession v = OpenTracked())
            {
                HashSet<string> after = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                StringBuilder ds = new StringBuilder();
                bool testsGone = true;
                foreach (DictInfo di in WordCom.ListAll(v))
                {
                    after.Add(di.Name);
                    ds.Append((ds.Length > 0 ? " | " : "") + di.Name);
                    if (IsTest(di.Name)) testsGone = false;
                }
                bool allBack = true;
                foreach (string nm in namesBefore) if (!IsTest(nm) && !after.Contains(nm)) allBack = false;
                string main = null;
                foreach (string p in oursPaths)
                    if (File.Exists(p) && (main == null || string.Equals(Path.GetFileName(p), Core.MainFile, StringComparison.OrdinalIgnoreCase))) main = p;
                int verdict = 1;
                string check = "";
                if (main != null) check = Verify.Run(v, Core.ReadDicWords(main), true, out verdict);
                ok = testsGone && allBack && verdict != 0;   // -1 = Word does not check Hebrew here: the list alone can be checked
                Out("#RESTORED " + (ok ? "ok" : "FAILED") + " removed_tests=" + removed + " via_registry=" + viaRegistry.Count + " tests_gone=" + (testsGone ? 1 : 0) +
                    " all_back=" + (allBack ? 1 : 0) + " verdict=" + verdict + " default=" + (ActiveName(v) ?? "-"));
                Out("#DICTS_AFTER " + ds);
                if (check.Length > 0) Out("#VERIFY " + OneLine(check));
            }
            WordCom.WaitForOurWordsToExit(15000);
            DeleteTestFilesExcept(null);
            if (ok) { try { File.Delete(MarkerPath); } catch (Exception ex) { Core.Log("Size check, marker: " + Core.Msg(ex)); } }
            return ok;
        }
    }

    // ------------------------------------------------------------------------------------------ the window
    class SizeForm : Form
    {
        public const string Title = "בדיקת גודל המילון";
        const int WatchdogMinutes = 10;
        const string Intro =
            "התוכנה בודקת כמה מילים וורד יכול לקבל במילון אחד, ועד כמה הוא נעשה איטי כשהמילון גדול.\r\n\r\n" +
            "לזמן הבדיקה היא מחליפה את המילון התורני במילוני בדיקה גדולים (עד מיליון מילים), ובודקת אותם בוורד נסתר. " +
            "בסוף היא מחזירה את המילון התורני כמו שהיה, ובודקת שוורד מכיר אותו שוב.\r\n\r\n" +
            "לפני שמתחילים: שמרו את המסמכים וסגרו את וורד. במשך הבדיקה, כחצי שעה עד שעה, לא לפתוח את וורד ולא את המילון התורני. " +
            "אפשר להשאיר את המחשב ולחזור אחר כך.\r\n\r\n" +
            "בסוף יישמר בשולחן העבודה הקובץ Milon_SizeCheck_results.txt. אותו שולחים בצ'אט.";
        const string PendingNote =
            "הבדיקה הקודמת נעצרה באמצע, והמילון התורני עוד לא חזר לוורד. \"החזר את המילון\" מחזיר אותו (כמה דקות). " +
            "\"התחל\" מחזיר אותו ואחר כך בודק מחדש.\r\n\r\n";

        TextBlock body, stage, clock;
        ProgressLine prog;
        FlatBtn btnStart, btnRepair, btnClose;
        SizeJob job;
        bool running, closeWhenDone, repairOnly;
        DateTime started;
        Mutex mainMutex;
        System.Windows.Forms.Timer watchdog, tick;

        [DllImport("kernel32.dll")] static extern uint SetThreadExecutionState(uint flags);
        static void KeepAwake(bool on)
        {
            try { SetThreadExecutionState(on ? 0x80000001u : 0x80000000u); } catch (Exception) { }   // ES_CONTINUOUS | ES_SYSTEM_REQUIRED
        }

        static void Info(IWin32Window owner, string text)
        {
            MessageBox.Show(owner, text, Title, MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }

        public SizeForm()
        {
            Text = Title;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = false;   // the controls are placed by hand
            if (Dpi.Forced) { StartPosition = FormStartPosition.Manual; Location = new Point(0, 0); }
            else StartPosition = FormStartPosition.CenterScreen;
            BackColor = Theme.Card;
            Font = Theme.F(14, false);
            ClientSize = Dpi.Sz(620, 520);
            try { Icon = MainForm.AppIcon(); } catch (Exception) { }

            TextBlock title = new TextBlock { Bounds = Dpi.R(28, 20, 564, 34), Font = Theme.F(22, true), Color = Theme.Text, VCenter = true, Wrap = false, Text = Title };
            body = new TextBlock { Bounds = Dpi.R(28, 66, 564, 300), Font = Theme.F(14, false), Color = Theme.Text };
            prog = new ProgressLine { Bounds = Dpi.R(28, 376, 564, 4) };
            stage = new TextBlock { Bounds = Dpi.R(28, 388, 564, 46), Font = Theme.F(14, false), Color = Theme.TextSoft, Text = "" };
            LowerPanel lower = new LowerPanel { Bounds = Dpi.R(0, 444, 620, 76) };
            btnStart = new FlatBtn { Bounds = Dpi.R(452, 16, 140, 44), Text = "התחל", Style = BtnStyle.Filled };
            btnRepair = new FlatBtn { Bounds = Dpi.R(262, 16, 178, 44), Text = "החזר את המילון", Style = BtnStyle.Outlined };
            btnClose = new FlatBtn { Bounds = Dpi.R(28, 16, 110, 44), Text = "סגור", Style = BtnStyle.Outlined, Light = true };
            clock = new TextBlock { Bounds = Dpi.R(150, 16, 290, 44), Font = Theme.F(13, false), Color = Theme.TextQuiet, VCenter = true, Wrap = false, Text = "", Visible = false };
            btnStart.Click += delegate { Start(false); };
            btnRepair.Click += delegate { Start(true); };
            btnClose.Click += delegate { Close(); };
            lower.Controls.AddRange(new Control[] { btnStart, btnRepair, btnClose, clock });
            Controls.AddRange(new Control[] { title, body, prog, stage, lower });
            ShowIntro();

            // a Word that gives no answer for 10 minutes (stuck?) is stopped; the check then puts everything back
            watchdog = new System.Windows.Forms.Timer { Interval = 30000 };
            watchdog.Tick += delegate
            {
                if (!running || job == null || (DateTime.Now - job.LastProgress).TotalMinutes < WatchdogMinutes) return;
                Core.Log("Size check: no answer from Word for " + WatchdogMinutes + " minutes");
                job.Ping();
                job.KillOurWords();
            };
            tick = new System.Windows.Forms.Timer { Interval = 1000 };
            tick.Tick += delegate
            {
                TimeSpan t = DateTime.Now - started;
                clock.Text = "זמן: " + ((int)t.TotalMinutes).ToString(CultureInfo.InvariantCulture) + ":" + t.Seconds.ToString("00", CultureInfo.InvariantCulture);
            };
        }

        void ShowIntro()
        {
            bool pending = SizeJob.Pending();
            body.Text = (pending ? PendingNote : "") + Intro;
            btnRepair.Visible = pending;
        }

        // for the developer's screenshot: the end of a check, with made-up results (no Word)
        public void PreviewDone()
        {
            SizeJob j = new SizeJob();
            j.LimitOk = 256000; j.LimitBad = 262000; j.RestoredOk = true; j.Changed = true;
            body.Text = Summary(j, false);
            prog.Value = 100;
        }

        void UI(MethodInvoker a)
        {
            try { if (!IsDisposed && IsHandleCreated) BeginInvoke(a); } catch (Exception) { }   // the window is already gone
        }

        void Start(bool onlyRepair)
        {
            if (running) return;
            if (WordCom.RunningWordCount() > 0)
            {
                Info(this, "וורד פתוח עכשיו.\r\n\r\nשמרו את המסמכים, סגרו את כל חלונות וורד ולחצו שוב.\r\n\r\nאם אין חלון פתוח, חכו רגע ונסו שוב (וורד עוד נסגר ברקע).");
                return;
            }
            // the Torah dictionary program changes the same list of Word: while the check runs, it must not run too
            bool created;
            try { mainMutex = new Mutex(true, @"Local\TorahDictionaryForWord.SingleInstance", out created); }
            catch (Exception) { mainMutex = null; created = true; }
            if (!created)
            {
                mainMutex.Close();
                mainMutex = null;
                Info(this, "התוכנה \"המילון התורני\" פתוחה עכשיו.\r\n\r\nסגרו אותה ולחצו שוב.");
                return;
            }
            running = true;
            repairOnly = onlyRepair;
            btnStart.Enabled = false;
            btnRepair.Visible = false;
            clock.Text = "";
            clock.Visible = true;
            UseWaitCursor = true;
            KeepAwake(true);
            started = DateTime.Now;
            prog.Value = 0;
            stage.Text = "מתחיל…";
            job = new SizeJob();
            if (onlyRepair) job.StopRequested = true;
            SizeJob j = job;
            j.Progress = delegate (int p, string s) { UI(delegate { if (p >= 0) prog.Value = p; stage.Text = s; }); };
            watchdog.Start();
            tick.Start();
            Thread t = new Thread(delegate ()
            {
                try { j.Run(); }
                catch (Exception ex) { Core.Log("Size check thread: " + ex); }
                UI(Finished);
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);   // Word is asked from this thread
            t.Start();
        }

        void ReleaseMain()
        {
            if (mainMutex == null) return;
            try { mainMutex.ReleaseMutex(); } catch (Exception) { }
            mainMutex.Close();
            mainMutex = null;
        }

        void Finished()
        {
            running = false;
            watchdog.Stop();
            tick.Stop();
            KeepAwake(false);
            ReleaseMain();
            UseWaitCursor = false;
            btnStart.Enabled = true;
            btnRepair.Visible = SizeJob.Pending();
            if (closeWhenDone) { Close(); return; }
            body.Text = Summary(job, repairOnly);
            stage.Text = "";
            prog.Value = 100;
            if (!repairOnly) { try { Process.Start("explorer.exe", "/select,\"" + SizeJob.ResultPath + "\""); } catch (Exception) { } }
        }

        static string N(int n) { return n.ToString("N0", CultureInfo.CurrentCulture); }

        static string Summary(SizeJob job, bool onlyRepair)
        {
            string back = job.RestoredOk ? "המילון התורני חזר לוורד, ווורד מכיר אותו."
                                         : "לא הצלחתי לוודא שהמילון התורני חזר לוורד. פתחו שוב את בדיקת הגודל ולחצו \"החזר את המילון\".";
            if (onlyRepair) return (job.Error != null ? job.Error + "\r\n\r\n" : "") + back;
            StringBuilder sb = new StringBuilder();
            if (job.Error != null) sb.Append("הבדיקה נעצרה: " + job.Error + "\r\n\r\n");
            else if (job.WordWasOpened) sb.Append("הבדיקה נעצרה, כי וורד נפתח באמצע.\r\n\r\n");
            else if (job.StoppedByPerson) sb.Append("הבדיקה נעצרה לבקשתכם.\r\n\r\n");
            else if (job.ControlsFailed) sb.Append("וורד לא בודק איות בעברית במחשב הזה, ולכן אי אפשר היה לבדוק.\r\n\r\n");
            string limit = "";
            if (job.LimitOk >= SizeJob.Sizes[1]) limit = "וורד הכיר את כל המילים במילון של עד " + N(job.LimitOk) + " מילים. ";
            if (job.ErrorAt > 0) limit += "בבדיקה של " + N(job.ErrorAt) + " מילים וורד הפסיק לענות, והבדיקה נעצרה.";
            else if (job.LimitBad > 0) limit += "במילון של " + N(job.LimitBad) + " מילים הוא כבר לא הכיר את כולן.";
            else if (job.LimitOk == SizeJob.Sizes[SizeJob.Sizes.Length - 1]) limit += "זה הגודל הגדול ביותר שנבדק.";
            if (limit.Length > 0) sb.Append(limit.Trim() + "\r\n\r\n");
            if (job.Changed) sb.Append(back + "\r\n\r\n");
            sb.Append(SizeJob.ResultsOnDesktop ? "התוצאות נשמרו בשולחן העבודה, בקובץ Milon_SizeCheck_results.txt. שלחו אותו בצ'אט."
                                               : "התוצאות נשמרו בקובץ:\r\n" + SizeJob.ResultPath + "\r\nשלחו אותו בצ'אט.");
            return sb.ToString();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (running && e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                if (closeWhenDone) return;
                if (MessageBox.Show(this, "הבדיקה עדיין רצה.\r\n\r\nלעצור אותה? התוכנה תחזיר קודם את המילון התורני לוורד (זה לוקח כמה דקות), ואז תיסגר.", Title,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign) != DialogResult.Yes) return;
                closeWhenDone = true;
                job.StopRequested = true;
                stage.Text = "עוצר ומחזיר את המילון התורני…";
                return;
            }
            base.OnFormClosing(e);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            KeepAwake(false);
            ReleaseMain();
            base.OnFormClosed(e);
        }
    }

    // ------------------------------------------------------------------------------------------ start
    static class SizeCheckProgram
    {
        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length >= 2 && args[0] == "--selftest")
            {
                // the self-test writes the list of dictionaries in this user's registry: developer machine only
                if (Type.GetType("Mono.Runtime") == null) { Console.WriteLine("The self-test runs only in the developer's environment (Mono)."); return 2; }
                return SizeCheckTest.Run(args[1]);
            }
            int forcedDpi = 0;
            for (int i = 0; i + 1 < args.Length; i++) if (args[i] == "--dpi") int.TryParse(args[i + 1], out forcedDpi);
            Dpi.Init(forcedDpi);
            bool first;
            Mutex single = null;
            try { single = new Mutex(true, @"Local\MilonSizeCheck.SingleInstance", out first); }
            catch (Exception) { first = true; }
            if (!first) return 0;   // one window at a time
            try
            {
                Core.TrimLog();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.ThreadException += delegate (object s, ThreadExceptionEventArgs e) { Fatal(e.Exception); };
                AppDomain.CurrentDomain.UnhandledException += delegate (object s, UnhandledExceptionEventArgs e) { Fatal(e.ExceptionObject as Exception); };
                SizeForm f = new SizeForm();
                if (Array.IndexOf(args, "--preview-done") >= 0) f.PreviewDone();
                Application.Run(f);
            }
            catch (Exception ex) { Fatal(ex); }
            finally { if (single != null) { try { single.ReleaseMutex(); } catch (Exception) { } single.Close(); } }
            return 0;
        }

        static void Fatal(Exception ex)
        {
            Core.Log("Size check FATAL " + ex);
            try
            {
                MessageBox.Show("אירעה שגיאה: " + (ex == null ? "" : Core.Msg(ex)) + "\r\n\r\nקובץ יומן: " + Core.LogPath, SizeForm.Title,
                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            catch (Exception) { }
        }
    }

    // ------------------------------------------------------------------------------------------ the self-test: a pretend Word
    // WordCom talks to Word by member names (late binding), so a pretend Word needs only members with the same names.
    // The pretend Word keeps its list of dictionaries where the real one does - the registry list, which this user's Mono registry holds here -
    // reads it when it starts (and the files in it: only these words are known), and writes it back when it quits.
    public class SFWorld
    {
        public int IgnoreOver = int.MaxValue;   // a file with more words: Word does not use it at all
        public int UsesFirst = int.MaxValue;    // Word uses only this many words from the start of a file
        public int RefuseOver = int.MaxValue;   // Word itself refuses to add a file with more words
        public string Default;                  // the name of the default dictionary, kept between runs
        public int FailStart = -1;              // this start of Word (counted from 1) fails, as if Word stopped answering
        public int Starts;

        public List<string> ReadList()
        {
            SortedDictionary<int, string> byNumber = new SortedDictionary<int, string>();
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath))
            {
                if (k != null)
                    foreach (string vn in k.GetValueNames())
                    {
                        int n;
                        if (vn.IndexOf('_') >= 0 || !int.TryParse(vn, out n)) continue;
                        string v = k.GetValue(vn) as string;
                        if (v != null) byNumber[n] = v;
                    }
            }
            return new List<string>(byNumber.Values);
        }

        public void WriteList(List<string> paths)
        {
            Registry.CurrentUser.DeleteSubKeyTree(RegFallback.KeyPath, false);
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RegFallback.KeyPath))
                for (int i = 0; i < paths.Count; i++) k.SetValue((i + 1).ToString(CultureInfo.InvariantCulture), paths[i]);
        }

        public List<string> Names()
        {
            List<string> l = new List<string>();
            foreach (string p in ReadList()) l.Add(Path.GetFileName(p));
            return l;
        }
    }

    public class SFApp
    {
        public SFWorld World;
        public SFDicts Dicts;
        public HashSet<string> Known = new HashSet<string>(StringComparer.Ordinal);
        public SFApp(SFWorld w)
        {
            World = w;
            w.Starts++;
            if (w.Starts == w.FailStart) throw new InvalidOperationException("simulated: Word stopped answering");
            Dicts = new SFDicts();
            Dicts.App = this;
            foreach (string p in w.ReadList())
            {
                SFDict d = new SFDict(Dicts, Path.GetFileName(p), Path.GetDirectoryName(p));
                Dicts.Items.Add(d);
                if (w.Default != null && string.Equals(d.Name, w.Default, StringComparison.OrdinalIgnoreCase)) Dicts.Active = d;
                if (!File.Exists(p)) continue;
                List<string> words = Core.ReadDicWords(p);
                if (words.Count > w.IgnoreOver) continue;
                for (int i = 0; i < words.Count && i < w.UsesFirst; i++) Known.Add(words[i]);
            }
        }
        public SFDicts CustomDictionaries { get { return Dicts; } }
        public SFDocs Documents { get { SFDocs d = new SFDocs(); d.App = this; return d; } }
        public string Version { get { return "16.0"; } }
        public string Build { get { return "pretend"; } }
        public void Quit(params object[] args)
        {
            List<string> paths = new List<string>();
            foreach (SFDict d in Dicts.Items) paths.Add(Path.Combine(d.Path, d.Name));
            World.WriteList(paths);
            World.Default = Dicts.Active != null ? Dicts.Active.Name : null;
        }
        public bool Marks(string w)
        {
            if (w == SizeJob.NotAWord) return true;
            if (w == SizeJob.CommonWord) return false;
            return !Known.Contains(w);
        }
    }

    public class SFDicts
    {
        public SFApp App;
        public List<SFDict> Items = new List<SFDict>();
        public SFDict Active;
        public int Count { get { return Items.Count; } }
        public SFDict Item(int i) { return Items[i - 1]; }
        public SFDict ActiveCustomDictionary { get { return Active ?? (Items.Count > 0 ? Items[0] : null); } set { Active = value; } }
        public SFDict Add(string path)
        {
            if (File.Exists(path) && Core.CountDicWords(path) > App.World.RefuseOver) throw new InvalidOperationException("simulated: Word would not add so big a dictionary");
            SFDict d = new SFDict(this, Path.GetFileName(path), Path.GetDirectoryName(path));
            Items.Add(d);
            return d;   // not in effect in this Word: a Word reads its dictionaries only when it starts
        }
    }

    public class SFDict
    {
        readonly SFDicts owner;
        readonly string name, folder;
        public SFDict(SFDicts owner, string name, string folder) { this.owner = owner; this.name = name; this.folder = folder; }
        public string Name { get { return name; } }
        public string Path { get { return folder; } }   // like the real one: the folder, not the file
        public bool LanguageSpecific { get; set; }
        public void Delete() { owner.Items.Remove(this); if (owner.Active == this) owner.Active = null; }
    }

    public class SFDocs { public SFApp App; public SFDoc Add(params object[] args) { SFDoc d = new SFDoc(); d.App = App; return d; } }
    public class SFDoc
    {
        public SFApp App;
        SFRange content;
        public SFRange Content { get { if (content == null) { content = new SFRange(); content.App = App; } return content; } }
        public void Close(params object[] args) { }
    }
    public class SFRange
    {
        public SFApp App;
        string text = "";
        public string Text { get { return text; } set { text = value ?? ""; } }
        public int LanguageID { get; set; }
        public bool NoProofing { get; set; }
        public SFErrors SpellingErrors
        {
            get
            {
                int n = 0;
                foreach (string w in text.Split(' ')) if (w.Length > 0 && App.Marks(w)) n++;
                SFErrors e = new SFErrors();
                e.Count = n;
                return e;
            }
        }
    }
    public class SFErrors { public int Count { get; set; } }

    static class SizeCheckTest
    {
        static int fails;
        static string dir, ours, roaming;
        static SFWorld world;

        static void Check(bool ok, string what) { Console.WriteLine((ok ? "  ok    " : "  FAIL  ") + what); if (!ok) fails++; }

        static string Sizes(SizeJob j) { List<string> l = new List<string>(); foreach (SizeResult r in j.Results) l.Add(r.Size + (r.Works ? "+" : "-")); return string.Join(" ", l.ToArray()); }

        static bool SameNames(params string[] want)
        {
            List<string> have = world.Names();
            if (have.Count != want.Length) return false;
            foreach (string w in want) if (!have.Contains(w)) return false;
            return true;
        }

        static bool NothingLeft()
        {
            foreach (string n in world.Names()) if (SizeJob.IsTest(n)) return false;
            return Directory.GetFiles(dir, SizeJob.TestPrefix + "*").Length == 0 && !SizeJob.Pending();
        }

        // a new "computer": Word lists these files, with this default dictionary, and takes dictionaries of every size
        static void Fresh(string defaultName, params string[] listed)
        {
            world.IgnoreOver = world.UsesFirst = world.RefuseOver = int.MaxValue;
            world.FailStart = -1;
            world.Starts = 0;
            world.WriteList(new List<string>(listed));
            world.Default = defaultName;
            if (File.Exists(SizeJob.ResultPath)) File.Delete(SizeJob.ResultPath);
            SizeJob.RunningWordsHook = delegate { return 0; };
        }

        static SizeJob RunJob(bool onlyRepair)
        {
            SizeJob j = new SizeJob();
            if (onlyRepair) j.StopRequested = true;
            DateTime t0 = DateTime.Now;
            j.Run();
            Console.WriteLine("  (" + (int)(DateTime.Now - t0).TotalSeconds + " s) sizes: " + Sizes(j));
            return j;
        }

        static string Res() { return File.Exists(SizeJob.ResultPath) ? File.ReadAllText(SizeJob.ResultPath) : ""; }

        public static int Run(string d)
        {
            try { Console.OutputEncoding = new UTF8Encoding(false); } catch (Exception) { }
            dir = Path.GetFullPath(d);
            if (Directory.Exists(dir)) Directory.Delete(dir, true);
            Directory.CreateDirectory(dir);
            Core.UProofOverride = dir;
            Core.LogPath = Path.Combine(dir, "log.txt");
            MyWords.FolderOverride = Path.Combine(dir, "settings");
            SizeJob.StateDir = Path.Combine(dir, "state");
            SizeJob.ResultPath = Path.Combine(dir, "results.txt");
            SizeJob.PollMs = 10;
            world = new SFWorld();
            WordCom.OpenNewHook = delegate { WordSession s = new WordSession(); s.App = new SFApp(world); s.Created = true; return s; };
            WordCom.OpenHook = WordCom.OpenNewHook;
            ours = Path.Combine(dir, Core.MainFile);
            roaming = Path.Combine(dir, "RoamingCustom.dic");
            Core.WriteDic(ours, Core.BuildWords(true, true));
            Core.WriteDic(roaming, new List<string> { "מילהשלי" });

            Console.WriteLine("-- made-up words");
            HashSet<string> synth = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < 1000000; i++) synth.Add(SizeJob.Synthetic(i));
            for (int i = 2000000; i < 2000050; i++) synth.Add(SizeJob.Synthetic(i));
            Check(synth.Count == 1000050, "1,000,050 different made-up words");
            bool clash = false;
            foreach (string w in Core.BuildWords(true, true)) if (synth.Contains(w)) clash = true;
            Check(!clash && SizeJob.Synthetic(0) == "זקזאאאאא" && SizeJob.Synthetic(1) == "זקזאאאאב", "none is a word of the dictionary (" + SizeJob.Synthetic(0) + ", " + SizeJob.Synthetic(1) + ")");

            Console.WriteLine("-- Word does not use a dictionary of more than 260,000 words");
            Fresh("RoamingCustom.dic", roaming, ours);
            world.IgnoreOver = 260000;
            SizeJob a = RunJob(false);
            string res = Res();
            Check(a.Error == null, "no error (" + a.Error + ")");
            Check(Sizes(a) == "0+ 175000+ 200000+ 250000+ 300000- 275000- 262000- 256000+", "the sizes up to the first failure, then three that narrow the limit down");
            Check(a.LimitOk == 256000 && a.LimitBad == 262000 && res.Contains("#LIMIT works=256000 fails=262000"), "limit: works at 256,000, fails at 262,000");
            SizeResult r0 = a.Results[0], r175 = a.Results[1], r300 = a.Results[4];
            Check(r0.ControlsOk && r0.BaseOk == 0 && r0.BaseAll == 15 && r0.OutMarked == 50, "without a dictionary: the controls work, no sample word is known (" + r0.BaseOk + "/" + r0.BaseAll + ")");
            Check(r175.Add == "ok" && r175.Via == "word" && r175.Listed && r175.BaseOk == 15 && r175.MidOk == 10 && r175.EndOk == 10 && r175.InMarked == 0 && r175.OutMarked == 50,
                  "175,000: added by Word; the words of the start, middle and end known; the unknown block marked");
            Check(r300.Listed && r300.BaseOk == 0 && r300.EndOk == 0 && r300.InMarked == 50, "300,000: listed, but no word of it known");
            Check(res.Contains("SIZE\t175000\tadd=ok\tvia=word\tlisted=1\t") && res.Contains("\tin_n=50\t") && res.Contains("\tworks=1") && res.Contains("#ELAPSED") && res.Contains("#DONE"),
                  "the result file has a line per size");
            Check(res.Contains("#RESTORED ok") && a.RestoredOk, "the Torah dictionary is back, and a new Word knows its words");
            Check(SameNames("RoamingCustom.dic", Core.MainFile) && world.Default == "RoamingCustom.dic", "Word's list and default dictionary are as before (" + string.Join(", ", world.Names().ToArray()) + "; default " + world.Default + ")");
            Check(NothingLeft(), "no test file, no test entry in Word's list, no restore mark");

            Console.WriteLine("-- Word uses only the first 230,000 words of a file (and the Torah dictionary is its only one)");
            Fresh(null, ours);
            world.UsesFirst = 230000;
            SizeJob b = RunJob(false);
            Check(Sizes(b) == "0+ 175000+ 200000+ 250000- 225000+ 237000- 231000-", "a file cut short is found");
            SizeResult r250 = b.Results[3];
            Check(r250.Listed && r250.BaseOk == 15 && r250.MidOk == 10 && r250.EndOk == 0 && r250.InMarked > 0 && r250.InMarked < 50,
                  "250,000: the start and the middle known, the end not (block: " + r250.InMarked + " of 50 marked)");
            List<string> names = world.Names();
            Check(b.RestoredOk && names.Contains(Core.MainFile) && NothingLeft(), "back: " + string.Join(", ", names.ToArray()));
            Check(!Core.IsOurs(world.Default ?? Core.MainFile), "the Torah dictionary is not Word's default dictionary (" + world.Default + ")");

            Console.WriteLine("-- the check was stopped by force in the middle: at the next start everything is put back first");
            string test200 = Path.Combine(dir, SizeJob.TestPrefix + "200000.dic");
            Core.WriteDic(test200, new List<string> { "זקזאאאאא" });
            Fresh("RoamingCustom.dic", roaming, test200);
            Directory.CreateDirectory(SizeJob.StateDir);
            File.WriteAllText(Path.Combine(SizeJob.StateDir, "restore.txt"), "version=1\r\nours=" + ours + "\r\ndefault=RoamingCustom.dic\r\nhad=RoamingCustom.dic\r\nhad=" + Core.MainFile + "\r\n");
            Check(SizeJob.Pending(), "the restore mark is seen");
            SizeJob c = RunJob(true);   // the person asked only to put things back
            res = Res();
            Check(res.Contains("#REPAIR_AT_START") && res.Contains("#RESTORED ok") && c.RestoredOk && c.Results.Count == 0, "put back, and nothing more was done");
            Check(SameNames("RoamingCustom.dic", Core.MainFile) && world.Default == "RoamingCustom.dic" && NothingLeft(), "the Torah dictionary is back, the test dictionary and its file are gone");

            Console.WriteLine("-- the Torah dictionary was not installed, and Word takes every size up to 1,000,000");
            Fresh(null, roaming);
            SizeJob dj = RunJob(false);
            Check(dj.Results.Count == SizeJob.Sizes.Length && dj.Results.TrueForAll(delegate (SizeResult r) { return r.Works; }), "every size works; nothing is narrowed down");
            Check(dj.LimitOk == 1000000 && dj.LimitBad == -1 && Res().Contains("#LIMIT works=1000000 fails=none"), "limit: none found");
            Check(dj.RestoredOk && SameNames("RoamingCustom.dic") && NothingLeft(), "Word's list is still only the person's dictionary (" + string.Join(", ", world.Names().ToArray()) + ")");

            Console.WriteLine("-- Word itself will not add more than 180,000 words, but takes a file from its registry list (up to 260,000)");
            Fresh("RoamingCustom.dic", roaming, ours);
            world.RefuseOver = 180000;
            world.IgnoreOver = 260000;
            SizeJob e = RunJob(false);
            SizeResult e200 = e.Results.Count > 2 ? e.Results[2] : new SizeResult();
            Check(e.Results[1].Via == "word" && e200.Add.StartsWith("failed:") && e200.Via == "registry" && e200.Works, "200,000: refused by Word, put in the registry list, and it works (" + e200.Add + ")");
            Check(e.LimitOk == 256000 && e.LimitBad == 262000, "limit: works at 256,000, fails at 262,000 (" + Sizes(e) + ")");
            Check(e.RestoredOk && SameNames("RoamingCustom.dic", Core.MainFile) && NothingLeft(), "back as before, and no test entry is left in the registry list");

            Console.WriteLine("-- Word itself will not add more than 100,000 words: the Torah dictionary goes back through the registry list");
            Fresh("RoamingCustom.dic", roaming, ours);
            world.RefuseOver = 100000;
            world.IgnoreOver = 190000;
            SizeJob h = RunJob(false);
            res = Res();
            Check(Sizes(h) == "0+ 175000+ 200000- 187000+ 193000-", "every size through the registry list; the limit narrowed down to less than 10,000");
            Check(h.RestoredOk && res.Contains("via_registry=1") && res.Contains("#RESTORED ok") && SameNames("RoamingCustom.dic", Core.MainFile) && NothingLeft(),
                  "the Torah dictionary is back through the registry list, and a new Word knows its words");

            Console.WriteLine("-- the person opens Word in the middle of the check");
            Fresh("RoamingCustom.dic", roaming, ours);
            int calls = 0;
            SizeJob.RunningWordsHook = delegate { calls++; return calls >= 4 && calls <= 6 ? 1 : 0; };   // open from the 3rd size on, closed later
            SizeJob f = RunJob(false);
            res = Res();
            Check(f.WordWasOpened && f.Results.Count == 2 && res.Contains("#STOP Word was opened") && res.Contains("#WAIT Word is open"), "the check stops, and waits until Word is closed");
            Check(f.RestoredOk && res.Contains("#RESTORED ok") && SameNames("RoamingCustom.dic", Core.MainFile) && NothingLeft(), "then the Torah dictionary is put back");

            Console.WriteLine("-- Word stops answering in the middle");
            Fresh("RoamingCustom.dic", roaming, ours);
            world.FailStart = 5;   // 1 = the list, 2-3 = without a dictionary, 4-5 = 175,000
            SizeJob g = RunJob(false);
            res = Res();
            Check(g.Results.Count == 2 && g.Results[1].Error != null && g.ErrorAt == 175000 && res.Contains("#STOP error at 175000") && res.Contains("error_at=175000") && res.Contains("\terror="),
                  "the error is written down and the check stops");
            Check(g.RestoredOk && SameNames("RoamingCustom.dic", Core.MainFile) && NothingLeft(), "the Torah dictionary is back after that too");

            Console.WriteLine(fails == 0 ? "SELFTEST PASSED" : "SELFTEST FAILED (" + fails + ")");
            return fails == 0 ? 0 : 1;
        }
    }
}
