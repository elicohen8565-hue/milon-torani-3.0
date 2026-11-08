// WordCheck 4 - asks Microsoft Word's own spelling checker which words from the books (forms the Torah dictionary does not have yet,
//   found by scanning the libraries) Word marks as errors. A form Word marks is a candidate to be added to the dictionary; a form Word
//   does not mark is known to Word and is not needed. The forms are asked in order of value (how many red marks each would remove),
//   so a run stopped in the middle still gives the most useful part.
//   The program is WordCheck 3 (2-11-2026, run on the user's Word in 37 minutes) with another word list: for the time of the check
//   the Torah dictionary is disconnected from Word (a copy of its file is kept first) - not because the forms are in it (they are
//   not), but because Word checks about ten times faster without it. Word is then started again, because a dictionary taken out of
//   a Word that is already running is not reliably out until Word is started again. The dictionary is connected back at the end.
//   If the program stops in the middle, its next start connects the dictionary back (and puts the copy back if the file
//   disappeared) and then continues from the place reached.
// Nothing else is changed. C# 5, late binding (no Office files needed). Same technique as the "verify" step of the installer.
// Build: build_wc.sh (compiled against the .NET 4.5 reference assemblies, then checked with CheckApi.exe).  Test: WordCheck.exe --selftest
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace WordCheck
{
    public class WordMissingException : Exception
    {
        public WordMissingException() : base("Microsoft Word is not installed (Word.Application is not registered).") { }
    }

    public class WordDiedException : Exception
    {
        public WordDiedException(string m, Exception inner) : base(m, inner) { }
    }

    // ---------------------------------------------------------------- COM helpers
    public static class Com
    {
        public static object Get(object o, string name, params object[] args)
        { return o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, args); }
        public static void Put(object o, string name, object value)
        { o.GetType().InvokeMember(name, BindingFlags.SetProperty, null, o, new object[] { value }); }
        public static object Call(object o, string name, params object[] args)
        { return o.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, o, args); }

        public static Exception Unwrap(Exception ex)
        {
            while (ex is TargetInvocationException && ex.InnerException != null) ex = ex.InnerException;
            return ex;
        }

        static uint Code(Exception ex)
        {
            COMException ce = Unwrap(ex) as COMException;
            return ce == null ? 0u : unchecked((uint)ce.ErrorCode);
        }

        // Word is busy for a moment: wait and ask again
        public static bool IsBusy(Exception ex)
        {
            uint c = Code(ex);
            return c == 0x80010001u || c == 0x8001010Au;
        }

        // Word is gone (crashed or closed): a new Word has to be started
        public static bool IsDead(Exception ex)
        {
            uint c = Code(ex);
            return c == 0x800706BAu || c == 0x800706BEu || c == 0x80010108u || c == 0x80010012u || c == 0x800706BFu;
        }

        public static string Msg(Exception ex)
        {
            ex = Unwrap(ex);
            return ex.GetType().Name + ": " + ex.Message;
        }

        public static string OneLine(string s)
        {
            return (s ?? "").Replace('\r', ' ').Replace('\n', ' ').Replace('\t', ' ');
        }

        static int itemMode;   // 0 = not known yet, 1 = "Item" answers as a method, 2 = as a property
        public static object Item(object coll, int i)
        {
            if (itemMode == 1) return Call(coll, "Item", i);
            if (itemMode == 2) return Get(coll, "Item", i);
            try { object r = Call(coll, "Item", i); itemMode = 1; return r; }
            catch (Exception ex)
            {
                if (IsBusy(ex) || IsDead(ex)) throw;
                object r = Get(coll, "Item", i);
                itemMode = 2;
                return r;
            }
        }

        public static void Release(object o)
        {
            try { if (o != null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); } catch (Exception) { }
        }
    }

    public static class Log
    {
        public static string Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "wordcheck-log.txt");
        public static void Write(string s)
        {
            try { File.AppendAllText(Path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + s + "\r\n", Encoding.UTF8); }
            catch (Exception) { }
        }
    }

    // the file names the Torah dictionary installer uses (3.0.0.2 uses only the first one)
    public static class Names
    {
        static readonly string[] Ours = new string[] { "Milon_Torani.dic", "Otzaria_Torah_Basic.dic", "Otzaria_Torah_Extended.dic",
            "Otzaria_Torah_Extended2.dic", "Otzaria_Torah_Acronyms.dic", "Otzaria_Torah_Acronyms4.dic", "Otzaria_Torah_Composition.dic" };

        public static bool IsOurs(string fileName)
        {
            if (fileName == null) return false;
            foreach (string f in Ours) if (string.Equals(fileName, f, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        // the dictionary file on disk. Word's "Path" may be a folder, the whole file path, or nothing useful,
        // so the installer's folder (%APPDATA%\Microsoft\UProof) is tried too. Not found = the bare name is returned.
        public static string ResolveFile(string name, string raw)
        {
            try
            {
                if (!string.IsNullOrEmpty(raw))
                {
                    if (File.Exists(raw) && string.Equals(Path.GetFileName(raw), name, StringComparison.OrdinalIgnoreCase)) return raw;
                    if (Directory.Exists(raw) && File.Exists(Path.Combine(raw, name))) return Path.Combine(raw, name);
                }
                string uproof = Path.Combine(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft"), "UProof");
                string p = Path.Combine(uproof, name);
                if (File.Exists(p)) return p;
            }
            catch (Exception) { }
            return name;
        }

        public static bool IsRealFile(string path)
        {
            try { return Path.IsPathRooted(path) && File.Exists(path); } catch (Exception) { return false; }
        }
    }

    // ---------------------------------------------------------------- what the job needs from Word
    public interface IMarker
    {
        bool Marked(string text);              // words separated by \r; true = Word marks at least one of them
        List<int> ErrorStarts(string text);    // where every mark starts, in characters from the start of the text
    }

    public interface IWord : IMarker, IDisposable
    {
        string Version { get; }
        string Build { get; }
        string DictionaryNames();
        List<string> FindOurs();               // the Torah dictionary file(s) that Word uses now (full paths)
        List<string> FindOthers();             // the other custom dictionaries that Word uses (the person's own words, for example)
        bool HasDictionary(string fileName);
        void Detach(string fileName);          // takes a dictionary out of Word's list (the file itself stays)
        void Attach(string fullPath);          // puts it back, for every language
        int Pid { get; }                       // the Windows process of this Word (0 = not known); used to wait until it is really gone
    }

    public class WordMarker : IWord
    {
        object app, doc, content;
        string version = "?";
        string build = "?";
        int calls;
        int pid;

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        // several Words are started one after the other (never at the same moment), so that the new process can be told from the others
        static readonly object startLock = new object();

        public string Version { get { return version; } }
        public string Build { get { return build; } }
        public int Pid { get { return pid; } }

        static HashSet<int> WordProcesses()
        {
            HashSet<int> ids = new HashSet<int>();
            try
            {
                foreach (Process p in Process.GetProcessesByName("WINWORD")) { ids.Add(p.Id); p.Dispose(); }
            }
            catch (Exception) { }
            return ids;
        }

        // which process is this Word? First by the handle of its window, otherwise it is the one Word process that did not exist before
        int FindPid(HashSet<int> before)
        {
            try
            {
                long h = Convert.ToInt64(Com.Get(app, "Hwnd"));
                if (h != 0)
                {
                    uint p;
                    GetWindowThreadProcessId(new IntPtr(h), out p);
                    if (p != 0) return (int)p;
                }
            }
            catch (Exception) { }
            int found = 0, count = 0;
            foreach (int id in WordProcesses()) if (!before.Contains(id)) { found = id; count++; }
            return count == 1 ? found : 0;
        }

        public WordMarker()
        {
            Type t = Type.GetTypeFromProgID("Word.Application");
            if (t == null) throw new WordMissingException();
            lock (startLock)
            {
                HashSet<int> before = WordProcesses();
                app = Activator.CreateInstance(t);               // a NEW Word, hidden
                pid = FindPid(before);
            }
            try
            {
                try { Com.Put(app, "Visible", false); } catch (Exception) { }
                try { Com.Put(app, "DisplayAlerts", 0); } catch (Exception) { }
                try { version = Convert.ToString(Com.Get(app, "Version")); } catch (Exception) { }
                try { build = Convert.ToString(Com.Get(app, "Build")); } catch (Exception) { }
                object docs = Com.Get(app, "Documents");
                doc = Com.Call(docs, "Add", Missing.Value, Missing.Value, Missing.Value, false);   // invisible document
                Com.Release(docs);
                content = Com.Get(doc, "Content");
            }
            catch (Exception)
            {
                Dispose();          // do not leave a hidden Word behind
                throw;
            }
            Log.Write("Started a hidden Word, version " + version + " build " + build + ", process " + pid);
        }

        public string DictionaryNames()
        {
            StringBuilder sb = new StringBuilder();
            object dicts = null;
            try
            {
                dicts = Com.Get(app, "CustomDictionaries");
                int n = Convert.ToInt32(Com.Get(dicts, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object d = null;
                    try
                    {
                        d = Com.Item(dicts, i);
                        if (sb.Length > 0) sb.Append(", ");
                        sb.Append(Convert.ToString(Com.Get(d, "Name")));
                        string raw = "?";
                        try { raw = Convert.ToString(Com.Get(d, "Path")); } catch (Exception) { }
                        sb.Append(" [" + raw + "]");
                    }
                    catch (Exception) { }
                    finally { Com.Release(d); }
                }
            }
            catch (Exception ex) { sb.Append("? (" + Com.Msg(ex) + ")"); }
            finally { Com.Release(dicts); }
            return sb.ToString();
        }

        void Tick()
        {
            if (++calls % 150 == 0) { try { Com.Call(doc, "UndoClear"); } catch (Exception) { } }   // thousands of replacements would pile up
        }

        object Prepare(string text)
        {
            Com.Put(content, "Text", text);
            Com.Put(content, "LanguageID", 1037);                  // Hebrew
            try { Com.Put(content, "NoProofing", false); } catch (Exception) { }
            return Com.Get(content, "SpellingErrors");
        }

        public bool Marked(string text)
        {
            Tick();
            for (int attempt = 0; ; attempt++)
            {
                object errs = null;
                try
                {
                    errs = Prepare(text);
                    return Convert.ToInt32(Com.Get(errs, "Count")) > 0;
                }
                catch (Exception ex)
                {
                    if (Com.IsBusy(ex) && attempt < 40) { Thread.Sleep(250); continue; }
                    if (Com.IsDead(ex)) throw new WordDiedException(Com.Msg(ex), ex);
                    throw;
                }
                finally { Com.Release(errs); }
            }
        }

        public List<int> ErrorStarts(string text)
        {
            Tick();
            for (int attempt = 0; ; attempt++)
            {
                object errs = null;
                try
                {
                    errs = Prepare(text);
                    int start0 = Convert.ToInt32(Com.Get(content, "Start"));
                    int n = Convert.ToInt32(Com.Get(errs, "Count"));
                    List<int> res = new List<int>(n);
                    for (int i = 1; i <= n; i++)
                    {
                        object r = Com.Item(errs, i);
                        try { res.Add(Convert.ToInt32(Com.Get(r, "Start")) - start0); }
                        finally { Com.Release(r); }
                    }
                    return res;
                }
                catch (Exception ex)
                {
                    if (Com.IsBusy(ex) && attempt < 40) { Thread.Sleep(250); continue; }
                    if (Com.IsDead(ex)) throw new WordDiedException(Com.Msg(ex), ex);
                    throw;
                }
                finally { Com.Release(errs); }
            }
        }

        public List<string> FindOurs() { return FindDictionaries(true); }
        public List<string> FindOthers() { return FindDictionaries(false); }

        List<string> FindDictionaries(bool ours)
        {
            List<string> res = new List<string>();
            object dicts = Com.Get(app, "CustomDictionaries");
            try
            {
                int n = Convert.ToInt32(Com.Get(dicts, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object d = Com.Item(dicts, i);
                    try
                    {
                        string name = Convert.ToString(Com.Get(d, "Name"));
                        if (Names.IsOurs(name) != ours) continue;
                        string raw = null;
                        try { raw = Convert.ToString(Com.Get(d, "Path")); } catch (Exception) { }
                        res.Add(Names.ResolveFile(name, raw));
                    }
                    finally { Com.Release(d); }
                }
            }
            finally { Com.Release(dicts); }
            return res;
        }

        public bool HasDictionary(string fileName)
        {
            object dicts = Com.Get(app, "CustomDictionaries");
            try
            {
                int n = Convert.ToInt32(Com.Get(dicts, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object d = Com.Item(dicts, i);
                    try
                    {
                        if (string.Equals(Convert.ToString(Com.Get(d, "Name")), fileName, StringComparison.OrdinalIgnoreCase)) return true;
                    }
                    finally { Com.Release(d); }
                }
            }
            finally { Com.Release(dicts); }
            return false;
        }

        public void Detach(string fileName)
        {
            object dicts = Com.Get(app, "CustomDictionaries");
            try
            {
                int n = Convert.ToInt32(Com.Get(dicts, "Count"));
                for (int i = n; i >= 1; i--)
                {
                    object d = Com.Item(dicts, i);
                    try
                    {
                        if (string.Equals(Convert.ToString(Com.Get(d, "Name")), fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            Com.Call(d, "Delete");
                            Log.Write("Disconnected from Word for part 2: " + fileName);
                        }
                    }
                    finally { Com.Release(d); }
                }
            }
            finally { Com.Release(dicts); }
        }

        public void Attach(string fullPath)
        {
            object dicts = Com.Get(app, "CustomDictionaries");
            try
            {
                object d = Com.Call(dicts, "Add", fullPath);
                try { Com.Put(d, "LanguageSpecific", false); } catch (Exception) { }   // every language, Hebrew included
                Com.Release(d);
                Log.Write("Connected to Word: " + fullPath);
            }
            finally { Com.Release(dicts); }
        }

        public void Dispose()
        {
            bool quit = true;
            try { if (doc != null) Com.Call(doc, "Close", 0); } catch (Exception) { }
            try
            {
                if (app != null)
                {
                    object docs = Com.Get(app, "Documents");
                    int n = Convert.ToInt32(Com.Get(docs, "Count"));
                    Com.Release(docs);
                    if (n > 0)
                    {
                        // somebody's own documents were opened inside this Word: leave it open and show it
                        quit = false;
                        try { Com.Put(app, "Visible", true); } catch (Exception) { }
                        Log.Write("Word has " + n + " other open document(s); it was left open and shown.");
                    }
                }
            }
            catch (Exception) { }
            try { if (quit && app != null) Com.Call(app, "Quit", 0); } catch (Exception) { }     // 0 = do not save
            Com.Release(content);
            Com.Release(doc);
            Com.Release(app);
            content = null; doc = null; app = null;
            Log.Write("Closed the hidden Word.");
        }
    }

    // ---------------------------------------------------------------- the lists (embedded, gzip, UTF-8, one per line)
    //   @D<TAB>word   a form from the books that is not in the Torah dictionary, most valuable first: this is what Word is asked about
    //                 (WordCheck 3: the words of the dictionary itself; the program does not depend on which it is)
    //   @C<TAB>word   control word: a dictionary word that Word surely does not know by itself (shows that the dictionary is disconnected)
    //   @E<TAB>word   speed test: a word that Word marks as an error (it is not in the dictionary)
    //   @A<TAB>word   speed test: an ordinary Hebrew word that Word accepts (it is not in the dictionary)
    public class Pool
    {
        public List<string> DictWords = new List<string>();
        public List<string> Controls = new List<string>();
        public List<string> Flagged = new List<string>();
        public List<string> Accepted = new List<string>();
        public string Id = "";

        public static Pool Load(Stream raw)
        {
            Pool p = new Pool();
            using (GZipStream gz = new GZipStream(raw, CompressionMode.Decompress))
            using (StreamReader r = new StreamReader(gz, new UTF8Encoding(false)))
            {
                string line;
                while ((line = r.ReadLine()) != null)
                {
                    if (line.Length == 0) continue;
                    if (line.StartsWith("@D\t")) p.DictWords.Add(line.Substring(3));
                    else if (line.StartsWith("@C\t")) p.Controls.Add(line.Substring(3));
                    else if (line.StartsWith("@E\t")) p.Flagged.Add(line.Substring(3));
                    else if (line.StartsWith("@A\t")) p.Accepted.Add(line.Substring(3));
                }
            }
            p.MakeId();
            return p;
        }

        // the name of this list of words: a result file of another list is never continued
        public void MakeId()
        {
            long h = 17;
            foreach (string w in DictWords) { foreach (char c in w) h = unchecked(h * 31 + c); h = unchecked(h * 31 + 10); }
            Id = DictWords.Count + "-" + Controls.Count + "-" + Flagged.Count + "-" + Accepted.Count + "-" + ((ulong)h).ToString("x");
        }

        public static Pool LoadEmbedded()
        {
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream("WordCheck.pool.txt.gz"))
            {
                if (s == null) throw new InvalidOperationException("Missing embedded word list.");
                return Load(s);
            }
        }
    }

    // ---------------------------------------------------------------- ways to ask Word about a long list
    public interface IEngine
    {
        int Next { get; set; }      // first word that is not decided yet
        long Calls { get; }
        void Run(Func<bool> cancelled, Action<List<string>> committed);
    }

    // what the main loop needs: the answer for the words [lo, hi) of the list = the words of this block that Word marks
    public interface IBatchEngine
    {
        List<string> FindIn(int lo, int hi);
        long Calls { get; }
    }

    // fast: one question per block of words; Word lists where its marks are, and each mark is matched to its word.
    // If a mark does not fall on any word, the places cannot be trusted for this block: it is asked again, the careful way.
    public class EnumEngine : IEngine, IBatchEngine
    {
        readonly IMarker m;
        readonly List<string> words;
        int next;
        long calls;
        public int Chunk = 250;
        public long Unmapped;       // marks that did not fall on a word (should stay 0)
        public long Fallbacks;      // blocks that were asked again the careful way because of that

        public EnumEngine(IMarker m, List<string> words) { this.m = m; this.words = words; }
        public int Next { get { return next; } set { next = value; } }
        public long Calls { get { return calls; } }

        public void Run(Func<bool> cancelled, Action<List<string>> committed)
        {
            while (next < words.Count && !cancelled())
            {
                int n = Math.Min(Chunk, words.Count - next);
                List<string> flagged = FindIn(next, next + n);
                next += n;
                committed(flagged);
            }
        }

        public List<string> FindIn(int lo, int hi)
        {
            int count = hi - lo;
            int[] starts = new int[count];
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < count; i++)
            {
                if (i > 0) sb.Append('\r');
                starts[i] = sb.Length;
                sb.Append(words[lo + i]);
            }
            calls++;
            List<int> errs = m.ErrorStarts(sb.ToString());
            bool[] hit = new bool[count];
            int lost = 0;
            foreach (int s in errs)
            {
                int k = Array.BinarySearch(starts, s);
                if (k < 0) k = ~k - 1;                       // the last word that starts before this place
                if (k < 0 || s >= starts[k] + words[lo + k].Length) { lost++; continue; }
                hit[k] = true;
            }
            if (lost > 0)
            {
                Unmapped += lost;
                Fallbacks++;
                GroupEngine careful = new GroupEngine(m, words);
                List<string> again = careful.FindIn(lo, hi);
                calls += careful.Calls;
                return again;
            }
            List<string> res = new List<string>();
            for (int i = 0; i < count; i++) if (hit[i]) res.Add(words[lo + i]);
            return res;
        }
    }

    // careful: asks about a block; when Word marks something, splits the block in two and asks again.
    // SingleOnly = ask word by word (the slowest and safest way)
    public class GroupEngine : IEngine, IBatchEngine
    {
        readonly IMarker m;
        readonly List<string> words;
        int next;
        long calls;
        int chunk = 32;
        int winTested, winFlagged;
        public bool SingleOnly;

        public GroupEngine(IMarker m, List<string> words) { this.m = m; this.words = words; }
        public int Next { get { return next; } set { next = value; } }
        public long Calls { get { return calls; } }

        public void Run(Func<bool> cancelled, Action<List<string>> committed)
        {
            while (next < words.Count && !cancelled())
            {
                List<string> flagged = new List<string>();
                int n;
                if (SingleOnly)
                {
                    n = Math.Min(20, words.Count - next);
                    for (int i = next; i < next + n; i++) { calls++; if (m.Marked(words[i])) flagged.Add(words[i]); }
                }
                else
                {
                    n = Math.Min(chunk, words.Count - next);
                    Find(next, next + n, flagged);
                    winTested += n; winFlagged += flagged.Count;
                    if (winTested >= 300) Adapt();
                }
                next += n;
                committed(flagged);
            }
        }

        // the words [lo, hi) that Word marks
        public List<string> FindIn(int lo, int hi)
        {
            List<string> res = new List<string>();
            if (SingleOnly)
            {
                for (int i = lo; i < hi; i++) { calls++; if (m.Marked(words[i])) res.Add(words[i]); }
            }
            else Find(lo, hi, res);
            return res;
        }

        // the fewer words Word marks, the bigger the blocks
        void Adapt()
        {
            double rate = winFlagged / (double)winTested;
            chunk = rate < 0.01 ? 96 : rate < 0.03 ? 48 : rate < 0.08 ? 20 : rate < 0.20 ? 6 : 1;
            winTested = 0; winFlagged = 0;
        }

        void Find(int lo, int hi, List<string> outList)
        {
            string text = (hi - lo == 1) ? words[lo] : string.Join("\r", words.GetRange(lo, hi - lo).ToArray());
            calls++;
            if (!m.Marked(text)) return;
            if (hi - lo == 1) { outList.Add(words[lo]); return; }
            int mid = lo + (hi - lo) / 2;
            Find(lo, mid, outList);
            Find(mid, hi, outList);
        }
    }

    // ---------------------------------------------------------------- the result file (text, UTF-8): can be continued after a stop
    public class ResultFile
    {
        public bool SamePool;       // it was made with the same list of words
        public int NextB;           // "Q": everything before this place is answered
        public int FlaggedB;        // number of "G" lines
        public bool Done;           // a "DONE" line: the check is finished (a "BSKIP" line is NOT a finish)
        public bool HasProbes;      // "#PROBES_DONE"
        public bool HasOther;       // "#OTHERDICT": the other custom dictionaries were looked at
        public bool HasSpeedConn;   // "#SPEED_DONE connected" (or notconnected)
        public bool HasSpeedDet;    // "#SPEED_DONE detached" (or notconnected2)
        public string Mode;         // the last "#MODE" line

        public static ResultFile Read(string path, string poolId)
        {
            ResultFile rf = new ResultFile();
            if (!File.Exists(path)) return rf;
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
            {
                int n;
                if (line.StartsWith("#POOLID ")) rf.SamePool = line.Substring(8).Trim() == poolId;
                else if (line.StartsWith("Q\t")) { if (int.TryParse(line.Substring(2), out n)) rf.NextB = n; }
                else if (line.StartsWith("G\t")) rf.FlaggedB++;
                else if (line.StartsWith("DONE\t")) rf.Done = true;
                else if (line.StartsWith("#MODE ")) rf.Mode = line.Substring(6).Trim();
                else if (line.StartsWith("#PROBES_DONE")) rf.HasProbes = true;
                else if (line.StartsWith("#OTHERDICT")) rf.HasOther = true;
                else if (line.StartsWith("#SPEED_DONE "))
                {
                    string state = line.Substring(12).Trim();
                    if (state == "connected" || state == "notconnected") rf.HasSpeedConn = true;
                    else if (state == "detached" || state == "notconnected2") rf.HasSpeedDet = true;
                }
            }
            return rf;
        }

        // After a run that was killed in the middle of writing: cut off a last line that is not complete, and the "G" lines of a block whose
        // "Q" line was never written (they come before it), so that the block is simply asked again. Returns true if the file was changed.
        public static bool TrimTail(string path)
        {
            string text = new UTF8Encoding(false).GetString(File.ReadAllBytes(path));
            bool changed = false;
            if (text.Length > 0 && !text.EndsWith("\n")) { text = text.Substring(0, text.LastIndexOf('\n') + 1); changed = true; }
            List<string> lines = new List<string>(text.Split('\n'));
            if (lines.Count > 0 && lines[lines.Count - 1].Length == 0) lines.RemoveAt(lines.Count - 1);
            int lastQ = -1;
            for (int i = 0; i < lines.Count; i++) if (lines[i].StartsWith("Q\t")) lastQ = i;
            List<string> keep = new List<string>();
            for (int i = 0; i < lines.Count; i++)
            {
                if (i > lastQ && lines[i].StartsWith("G\t")) { changed = true; continue; }
                keep.Add(lines[i]);
            }
            if (!changed) return false;
            string tmp = path + ".tmp";
            File.WriteAllText(tmp, keep.Count == 0 ? "" : string.Join("\n", keep.ToArray()) + "\n", new UTF8Encoding(false));
            File.Copy(tmp, path, true);
            File.Delete(tmp);
            return true;
        }
    }

    // ---------------------------------------------------------------- putting the Torah dictionary back after the check
    public static class Restore
    {
        public static string DirOverride = null;   // used only by the self-test

        static string Dir
        {
            get
            {
                if (DirOverride != null) return DirOverride;
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordCheck");
            }
        }

        public static string MarkerPath { get { return Path.Combine(Dir, "restore.txt"); } }

        public static bool Pending() { return File.Exists(MarkerPath); }

        // before disconnecting: keep a copy of every file and write down what has to be put back
        public static void Prepare(List<string> paths)
        {
            Directory.CreateDirectory(Dir);
            StringBuilder sb = new StringBuilder();
            foreach (string p in paths)
            {
                string bak = Path.Combine(Dir, Path.GetFileName(p) + ".backup");
                if (!File.Exists(p)) throw new FileNotFoundException("dictionary file not found: " + p);
                File.Copy(p, bak, true);
                if (new FileInfo(bak).Length != new FileInfo(p).Length) throw new IOException("the backup copy is not complete: " + bak);
                sb.Append(p + "\t" + bak + "\r\n");
            }
            File.WriteAllText(MarkerPath, sb.ToString(), new UTF8Encoding(false));
            Log.Write("Restore list written: " + sb.ToString().Replace("\r\n", " | "));
        }

        // connects every dictionary back (puts the copy back first if the file disappeared) and checks it
        public static bool PutBack(IWord wm, StringBuilder report)
        {
            if (!File.Exists(MarkerPath)) return true;
            bool ok = true;
            foreach (string line in File.ReadAllLines(MarkerPath, Encoding.UTF8))
            {
                string[] f = line.Split('\t');
                if (f.Length < 2 || f[0].Length == 0) continue;
                string path = f[0], bak = f[1], name = Path.GetFileName(path);
                try
                {
                    long have = File.Exists(path) ? new FileInfo(path).Length : -1;
                    long saved = File.Exists(bak) ? new FileInfo(bak).Length : -1;
                    if (saved > 0 && have < saved / 2)
                    {
                        File.Copy(bak, path, true);
                        report.Append("file " + name + " copied back from the backup (" + have + " < " + saved + "); ");
                        Log.Write("Dictionary file copied back from the backup: " + path);
                    }
                    if (!wm.HasDictionary(name)) wm.Attach(path);
                    bool now = wm.HasDictionary(name);
                    report.Append(name + (now ? " connected; " : " NOT connected; "));
                    if (!now) ok = false;
                }
                catch (WordDiedException) { throw; }
                catch (Exception ex)
                {
                    ok = false;
                    report.Append(name + " failed: " + Com.OneLine(Com.Msg(ex)) + "; ");
                    Log.Write("Putting back " + path + " failed: " + Com.Msg(ex));
                }
            }
            if (ok)
            {
                try { File.Delete(MarkerPath); } catch (Exception) { }
            }
            return ok;
        }
    }
    // ---------------------------------------------------------------- the whole job
    // WordCheck 3 does one thing: it asks Word which words of the Torah dictionary Word also knows WITHOUT the dictionary.
    // What the first run (version 2) taught, and what this version does about it:
    //  * A dictionary that is added or taken away in a Word that is already running is not reliably in effect until Word is started
    //    again. So after the dictionary is disconnected, Word is closed and a NEW Word is started, and a control check proves the change.
    //  * The speed of the first run was uneven, so this version measures: the same kind of words with the dictionary connected and
    //    disconnected, and the time of every batch (lines SPEED and T in the result file).
    //  * Several hidden Word windows work at the same time (each is a separate Word process), each on its own batches.
    public class BatchDone
    {
        public int Lo, Hi, Worker;
        public long Ms;
        public List<string> Found;
    }

    public class Job
    {
        public Pool P;
        public string ResultPath;
        public volatile bool Cancel;
        public int Workers = 1;                               // Word windows working at the same time (the one of this thread is included)
        public int BatchSize = 250;
        public int PauseMs = 3000;                            // wait before a Word is started again after it failed (the self-test sets 0)
        public Func<IWord> CreateWord = delegate () { return new WordMarker(); };
        public Action<string> OnStatus;                       // text for the person
        public Action<int, int, int, double> OnProgress;      // part (always 2), done, total, words per second
        public string Outcome = "";                           // done / stopped / alreadydone / error
        public string ErrorText = "";
        public bool RestoreProblem;                           // the Torah dictionary could not be connected back
        public string FinalMode = "";
        public int WorkersUsed;

        IWord wm;                                             // the Word of this thread
        StreamWriter w;
        readonly object wlock = new object();
        readonly List<int> pids = new List<int>();
        bool detachedByUs;                                    // the dictionary was connected and this run took it out
        int controlsBefore;                                   // how many control words Word marked while the dictionary was connected
        volatile string curMode = "enum";

        void Status(string s) { if (OnStatus != null) OnStatus(s); }
        static string B(bool b) { return b ? "1" : "0"; }

        [DllImport("kernel32.dll")]
        static extern uint SetThreadExecutionState(uint flags);

        // while the check runs the computer must not fall asleep (the screen may switch off); off = back to normal
        static void KeepAwake(bool on)
        {
            try { SetThreadExecutionState(on ? 0x80000001u : 0x80000000u); }     // ES_CONTINUOUS | ES_SYSTEM_REQUIRED / ES_CONTINUOUS
            catch (Exception) { }
        }

        // the new Word must behave like a Word with Hebrew proofing: a nonsense word is marked, an ordinary word is not
        static bool Sane(IWord x)
        {
            return x.Marked("קזקזקזק") && !x.Marked("שלום");
        }
        void Out(string s) { lock (wlock) { if (w != null) w.WriteLine(s); } }
        void WriteRaw(string s) { lock (wlock) { if (w != null) w.Write(s); } }

        // ------------------------------------------------------------ Word windows
        IWord NewWord()
        {
            IWord x = CreateWord();
            int pid = 0;
            try { pid = x.Pid; } catch (Exception) { }
            if (pid > 0) lock (pids) { if (!pids.Contains(pid)) pids.Add(pid); }
            return x;
        }

        internal static void WaitGone(int pid, int ms)
        {
            if (pid <= 0) return;
            try
            {
                Process p = Process.GetProcessById(pid);
                if (!string.Equals(p.ProcessName, "WINWORD", StringComparison.OrdinalIgnoreCase)) return;   // the number belongs to another program now
                if (!p.WaitForExit(ms))
                {
                    Log.Write("Word " + pid + " did not close in time; closed by force.");
                    try { p.Kill(); } catch (Exception) { }
                }
            }
            catch (ArgumentException) { }            // the process is already gone
            catch (Exception ex) { Log.Write("Waiting for Word " + pid + ": " + Com.Msg(ex)); }
        }

        // closes a Word and waits until its process is really gone
        void CloseWordQuiet(IWord x)
        {
            if (x == null) return;
            int pid = 0;
            try { pid = x.Pid; } catch (Exception) { }
            try { x.Dispose(); } catch (Exception ex) { Log.Write("Closing Word failed: " + Com.Msg(ex)); }
            WaitGone(pid, 20000);
            if (pid > 0) lock (pids) { pids.Remove(pid); }
        }

        // the last resort, when the window is closed and Word does not answer: end every Word that this run started and did not close
        public void KillWords()
        {
            lock (pids)
            {
                foreach (int pid in pids)
                {
                    try
                    {
                        Process p = Process.GetProcessById(pid);
                        if (string.Equals(p.ProcessName, "WINWORD", StringComparison.OrdinalIgnoreCase)) p.Kill();
                    }
                    catch (Exception) { }
                }
            }
        }

        void CloseMain()
        {
            IWord x = wm;
            wm = null;
            CloseWordQuiet(x);
        }

        static int CountMarked(IWord x, List<string> words)
        {
            int n = 0;
            foreach (string s in words) if (x.Marked(s)) n++;
            return n;
        }

        // a NEW Word for one worker; it has to behave as "dictionary disconnected" (checked with the control words)
        IWord StartWorker(int id)
        {
            for (int attempt = 0; ; attempt++)
            {
                IWord x = NewWord();
                bool good = false;
                try
                {
                    if (!Sane(x)) good = false;
                    else if (!detachedByUs || P.Controls.Count == 0) good = true;
                    else good = x.FindOurs().Count == 0 && CountMarked(x, P.Controls) > controlsBefore;
                }
                catch (Exception) { CloseWordQuiet(x); throw; }
                if (good) return x;
                Log.Write("Worker " + id + ": the new Word does not behave as needed (attempt " + attempt + ")");
                CloseWordQuiet(x);
                if (attempt >= 1) throw new InvalidOperationException("a new Word does not behave as needed (it accepts the dictionary words after the dictionary was disconnected, or its Hebrew check does not work)");
                Thread.Sleep(PauseMs);
            }
        }

        IBatchEngine MakeEngine(string mode, IWord x, List<string> list)
        {
            if (mode == "enum") return new EnumEngine(x, list);
            GroupEngine g = new GroupEngine(x, list);
            g.SingleOnly = mode == "single";
            return g;
        }

        // ------------------------------------------------------------ putting the dictionary back
        // connects every dictionary back (in a Word that is then closed), and checks it in a second new Word
        bool RestoreAndVerify(StringBuilder rep)
        {
            bool ok = false;
            for (int attempt = 0; attempt < 2 && !ok; attempt++)
            {
                IWord a = null;
                try
                {
                    a = NewWord();
                    rep.Length = 0;
                    ok = Restore.PutBack(a, rep);
                }
                catch (WordDiedException) { Log.Write("Word closed while the dictionary was put back; trying again."); }
                finally { CloseWordQuiet(a); }
            }
            if (ok && P.Controls.Count > 0)
            {
                IWord b = null;
                try
                {
                    b = NewWord();
                    int after = CountMarked(b, P.Controls);
                    rep.Append(" controls after " + after + " of " + P.Controls.Count);
                }
                catch (Exception ex) { rep.Append(" (the final check did not run: " + Com.OneLine(Com.Msg(ex)) + ")"); }
                finally { CloseWordQuiet(b); }
            }
            return ok;
        }

        void FinishRestore()
        {
            if (!Restore.Pending()) return;
            Status("מחבר בחזרה את המילון התורני לוורד…");
            StringBuilder rep = new StringBuilder();
            bool ok = false;
            try { CloseMain(); ok = RestoreAndVerify(rep); }
            catch (Exception ex) { rep.Append(Com.OneLine(Com.Msg(ex))); Log.Write("Restore failed: " + Com.Msg(ex)); }
            Out("#RESTORED " + (ok ? "ok " : "FAILED ") + rep);
            Log.Write("Restore: " + (ok ? "ok " : "FAILED ") + rep);
        }

        // ------------------------------------------------------------ the run
        public void Run()
        {
            try
            {
                KeepAwake(true);
                RunCore();
            }
            catch (WordMissingException)
            {
                Outcome = "error";
                ErrorText = "לא מצאתי וורד מותקן במחשב הזה.";
            }
            catch (Exception ex)
            {
                Log.Write("Job failed: " + Com.Msg(ex));
                Outcome = "error";
                ErrorText = "אירעה שגיאה: " + Com.Msg(ex) + "\r\n\r\nהתקדמות הבדיקה נשמרה, אפשר להפעיל שוב ולהמשיך.";
                try { Out("#ERROR " + Com.OneLine(Com.Msg(ex))); } catch (Exception) { }
            }
            finally
            {
                // never leave the dictionary disconnected
                try { FinishRestore(); } catch (Exception ex) { Log.Write("FinishRestore: " + Com.Msg(ex)); }
                RestoreProblem = Restore.Pending();
                try { CloseMain(); } catch (Exception) { }
                lock (pids) { foreach (int pid in pids) WaitGone(pid, 3000); }
                lock (wlock) { try { if (w != null) w.Dispose(); } catch (Exception) { } w = null; }
                KeepAwake(false);
            }
        }

        void AppendLine(string s)
        {
            try { File.AppendAllText(ResultPath, s + "\n", new UTF8Encoding(false)); } catch (Exception) { }
        }

        void RunCore()
        {
            ResultFile rf = ResultFile.Read(ResultPath, P.Id);
            bool resume = rf.SamePool && File.Exists(ResultPath);
            if (resume && !rf.Done)
            {
                // a run that was killed while writing may have left half a line / the answers of a block that was not finished
                try { if (ResultFile.TrimTail(ResultPath)) { Log.Write("The end of the result file was cleaned before continuing."); rf = ResultFile.Read(ResultPath, P.Id); } }
                catch (Exception ex) { Log.Write("Cleaning the end of the result file failed: " + Com.Msg(ex)); }
            }

            // the dictionary was left disconnected by an earlier run that stopped in the middle: connect it back first
            string repairText = null;
            if (Restore.Pending())
            {
                Status("מחבר בחזרה את המילון התורני לוורד…");
                StringBuilder rep = new StringBuilder();
                bool ok = RestoreAndVerify(rep);
                repairText = (ok ? "ok " : "FAILED ") + rep;
                Log.Write("Repair at start: " + repairText);
                if (!ok) { Outcome = "error"; ErrorText = "לא הצלחתי לחבר בחזרה את המילון התורני לוורד."; return; }
            }

            if (rf.SamePool && rf.Done)
            {
                Outcome = "alreadydone";
                if (repairText != null) AppendLine("#REPAIRED " + repairText);
                return;
            }
            if (File.Exists(ResultPath) && !resume)
            {
                try { File.Move(ResultPath, ResultPath + ".old" + DateTime.Now.ToString("HHmmss")); } catch (Exception) { }
                rf = new ResultFile();
            }

            lock (wlock)
            {
                w = new StreamWriter(new FileStream(ResultPath, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false));
                w.NewLine = "\n";
                w.AutoFlush = true;
            }
            if (!resume)
            {
                Out("#WORDCHECK 4");
                Out("#POOLID " + P.Id);
                Out("#POOL " + P.DictWords.Count + " " + P.Controls.Count + " " + P.Flagged.Count + " " + P.Accepted.Count);
            }
            else Out("#RESUME " + DateTime.Now.ToString("yyyy-MM-dd HH:mm"));
            Out("#CLOCK " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " start");
            if (repairText != null) Out("#REPAIRED " + repairText);

            Status("מתחבר לוורד (בפעם הראשונה זה יכול לקחת כמה שניות)…");
            wm = NewWord();
            Out("#WORD " + wm.Version + " build " + wm.Build);
            Out("#DICTS " + wm.DictionaryNames());

            // the check itself must work: a nonsense word is marked, ordinary words are not
            bool c1 = wm.Marked("קזקזקזק");
            bool c2 = wm.Marked("שלום");
            bool c3 = wm.Marked("שלום\rבית\rספר");
            bool c4 = wm.Marked("שלום\rקזקזקזק\rבית");
            Out("#CONTROLS " + c1 + " " + c2 + " " + c3 + " " + c4);
            if (!c1 || c2 || c3 || !c4)
            {
                Outcome = "error";
                ErrorText = "וורד לא סימן את מילת הבקרה כמצופה, ולכן אי אפשר לסמוך על הבדיקה. ודאו שבוורד מותקנים כלי ההגהה בעברית.";
                return;
            }

            // ---- with the dictionary still connected: control words, a few more questions, and the speed of Word
            List<string> ours = wm.FindOurs();
            bool connected = ours.Count > 0;
            controlsBefore = CountMarked(wm, P.Controls);
            Out("#BCONTROLS before " + controlsBefore + " of " + P.Controls.Count + (connected ? " (dictionary connected)" : " (dictionary NOT connected)"));
            if (!connected && !Restore.Pending()) Out("#BNOTCONNECTED the Torah dictionary is not connected to Word");

            if (!rf.HasOther) NoteOthers();
            if (!rf.HasProbes && !Cancel)
            {
                Status("בודק צורות של מילים מהמילון…");
                RunProbes();
            }
            if (!rf.HasSpeedConn && !Cancel)
            {
                Status("מודד את מהירות וורד (המילון מחובר)…");
                RunSpeed(connected ? "connected" : "notconnected", wm, 0);
            }
            if (Cancel) { Outcome = "stopped"; Out("#STOPPED 2"); return; }

            // ---- take the dictionary out of Word, then start a NEW Word, which reads the list the way Word does when it opens
            if (connected)
            {
                Status("מנתק זמנית את המילון התורני…");
                foreach (string p in ours)
                {
                    if (!Names.IsRealFile(p))
                    {
                        Out("BSKIP\tdictionary file not found: " + p);
                        Log.Write("Skipped, file not found: " + p);
                        Outcome = "error";
                        ErrorText = "לא מצאתי את קובץ המילון התורני בתיקיית וורד, ולכן לא ניתקתי אותו.";
                        return;
                    }
                }
                try { Restore.Prepare(ours); }                  // a copy of the file and a note of what to connect back
                catch (Exception ex)
                {
                    Out("BSKIP\tnot disconnected: " + Com.OneLine(Com.Msg(ex)));
                    Log.Write("Skipped: " + Com.Msg(ex));
                    Outcome = "error";
                    ErrorText = "לא הצלחתי ליצור עותק גיבוי של המילון, ולכן לא ניתקתי אותו.";
                    return;
                }
                detachedByUs = true;
                foreach (string p in ours) wm.Detach(Path.GetFileName(p));
                List<string> left = wm.FindOurs();
                Out("#BDETACHED " + string.Join(" | ", ours.ToArray()) + " -> still connected: " + left.Count);
                if (left.Count > 0)
                {
                    Out("BSKIP\tcould not disconnect");
                    Outcome = "error";
                    ErrorText = "וורד לא הסכים לנתק את המילון.";
                    return;
                }
            }

            Status("מפעיל את וורד מחדש…");
            CloseMain();
            wm = NewWord();                                       // always a new process, whether or not the dictionary was connected
            List<string> still = wm.FindOurs();
            if (still.Count > 0)
            {
                Out("BSKIP\tstill connected after the restart");
                Outcome = "error";
                ErrorText = "המילון נשאר מחובר לוורד גם אחרי שהופעל מחדש.";
                return;
            }
            if (!Sane(wm))
            {
                Out("BSKIP\tthe restarted Word does not check Hebrew as expected");
                Outcome = "error";
                ErrorText = "אחרי ההפעלה מחדש וורד לא בודק עברית כמצופה, ולכן אי אפשר לסמוך על הבדיקה. הבדיקה לא בוצעה, והמילון התורני מחובר בחזרה.";
                return;
            }
            int during = CountMarked(wm, P.Controls);
            Out("#BCONTROLS during " + during + " of " + P.Controls.Count);
            if (detachedByUs && P.Controls.Count > 0 && during <= controlsBefore)
            {
                // Word still accepts the dictionary words: disconnecting had no effect, the answers would mean nothing
                Out("BSKIP\tno effect");
                Outcome = "error";
                ErrorText = "אחרי הניתוק וורד ממשיך להכיר את מילות המילון, ולכן הבדיקה לא תהיה אמינה. הבדיקה לא בוצעה, והמילון התורני מחובר בחזרה.";
                return;
            }

            if (!rf.HasSpeedDet && !Cancel)
            {
                Status("מודד את מהירות וורד (המילון מנותק)…");
                RunSpeed(detachedByUs ? "detached" : "notconnected2", wm, 1);
            }

            string mode = rf.Mode;
            if (string.IsNullOrEmpty(mode) && !Cancel)
            {
                Status("בוחר את דרך הבדיקה המהירה ביותר…");
                mode = Calibrate();
                Out("#MODE " + mode);
            }
            if (string.IsNullOrEmpty(mode)) mode = "enum";
            curMode = mode;
            FinalMode = mode;

            int flagged = rf.FlaggedB;
            if (!Cancel && rf.NextB < P.DictWords.Count)
            {
                Status("בודק את מילות המילון מול וורד…");
                RunParallel(rf.NextB, ref flagged);
                FinalMode = curMode;
            }
            if (Cancel) { Outcome = "stopped"; Out("#STOPPED 2"); return; }

            Out("BDONE\t" + P.DictWords.Count + "\t" + flagged);
            Out("DONE\t" + P.DictWords.Count + "\t" + flagged);
            Out("#CLOCK " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " end");
            Outcome = "done";
            Log.Write("Finished: " + flagged + " of " + P.DictWords.Count + " dictionary words are marked without the dictionary, mode " + curMode);
        }

        // ------------------------------------------------------------ the person's other dictionaries
        // The words of a custom dictionary file (Word's format: one word per line, normally UTF-16 with a byte order mark)
        static HashSet<string> ReadDic(string path)
        {
            byte[] raw = File.ReadAllBytes(path);
            Encoding enc = null;
            int skip = 0;
            if (raw.Length >= 2 && raw[0] == 0xFF && raw[1] == 0xFE) { enc = Encoding.Unicode; skip = 2; }
            else if (raw.Length >= 2 && raw[0] == 0xFE && raw[1] == 0xFF) { enc = Encoding.BigEndianUnicode; skip = 2; }
            else if (raw.Length >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF) { enc = Encoding.UTF8; skip = 3; }
            string text;
            if (enc != null) text = enc.GetString(raw, skip, raw.Length - skip);
            else
            {
                try { text = new UTF8Encoding(false, true).GetString(raw); }                 // no byte order mark: UTF-8 if it is valid UTF-8,
                catch (Exception) { text = Encoding.GetEncoding(1255).GetString(raw); }       // otherwise the old Hebrew code page
            }
            HashSet<string> words = new HashSet<string>();
            foreach (string line in text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
            {
                string w = line.Trim();
                if (w.Length > 0) words.Add(w);
            }
            return words;
        }

        // The other custom dictionaries (the person's own words, for example) stay connected during the whole check, so Word accepts their
        // words as well. The words of the Torah dictionary that are in them are listed (H lines): they must not be taken for words that
        // Word knows by itself.
        void NoteOthers()
        {
            List<string> others;
            try { others = wm.FindOthers(); }
            catch (WordDiedException) { throw; }
            catch (Exception ex) { Out("#OTHERDICT failed: " + Com.OneLine(Com.Msg(ex))); return; }
            if (others.Count == 0) { Out("#OTHERDICT none"); return; }
            foreach (string path in others)
            {
                try
                {
                    if (!Names.IsRealFile(path)) { Out("#OTHERDICT " + path + " not found"); continue; }
                    HashSet<string> words = ReadDic(path);
                    StringBuilder sb = new StringBuilder();
                    int hit = 0;
                    foreach (string w in P.DictWords) if (words.Contains(w)) { sb.Append("H\t" + w + "\n"); hit++; }
                    Out("#OTHERDICT " + path + " words " + words.Count + " also in the dictionary " + hit);
                    if (hit > 0) WriteRaw(sb.ToString());
                }
                catch (Exception ex) { Out("#OTHERDICT " + path + " failed: " + Com.OneLine(Com.Msg(ex))); }
            }
        }

        // ------------------------------------------------------------ questions with the dictionary connected
        // Does Word accept endings on a word that comes from a custom dictionary? (Word is asked about each ending separately.)
        void RunProbes()
        {
            HashSet<string> d = new HashSet<string>(P.DictWords);
            string[] suffixes = new string[] { "ים", "ות", "ו", "ה", "ם", "י", "ך", "נו", "ין", "א" };
            foreach (string b in P.Controls)
            {
                if (Cancel) return;
                Out("PROBE\t" + b + "\t" + B(wm.Marked(b)));
                foreach (string sf in suffixes)
                {
                    string x = b + sf;
                    if (d.Contains(x)) continue;
                    Out("PROBE\t" + x + "\t" + B(wm.Marked(x)));
                }
            }
            Out("#PROBES_DONE");
        }

        // The same kind of words, asked in the same way as in the real run, with the time of each batch.
        // Two classes: words Word marks and words Word accepts. 'slot' picks different words for the two states.
        void RunSpeed(string state, IWord x, int slot)
        {
            string[] names = new string[] { "flagged", "accepted" };
            List<string>[] lists = new List<string>[] { P.Flagged, P.Accepted };
            int per = 2;
            for (int c = 0; c < 2; c++)
            {
                List<string> src = lists[c];
                EnumEngine ee = new EnumEngine(x, src);
                for (int k = 0; k < per; k++)
                {
                    if (Cancel) return;
                    int lo = (slot * per + k) * BatchSize;
                    int hi = Math.Min(src.Count, lo + BatchSize);
                    if (hi - lo < 10) continue;
                    Stopwatch sw = Stopwatch.StartNew();
                    List<string> res = ee.FindIn(lo, hi);
                    Out("SPEED\t" + state + "\t" + names[c] + "\t" + k + "\t" + sw.ElapsedMilliseconds + "\t" + res.Count + "\t" + (hi - lo));
                }
            }
            Out("#SPEED_DONE " + state);
        }

        // decides how to ask Word: the fast way is used only if it gives exactly the answers of asking word by word
        string Calibrate()
        {
            List<string> sample = new List<string>();
            int step = Math.Max(1, P.DictWords.Count / 200);
            for (int i = 0; i < P.DictWords.Count && sample.Count < 200; i += step) sample.Add(P.DictWords[i]);
            sample.Insert(sample.Count / 2, "קזקזקזק");
            sample.Insert(0, "שלום"); sample.Insert(1, "בית"); sample.Insert(2, "ספר");   // some words are surely accepted
            HashSet<string> truth = new HashSet<string>();
            foreach (string s in sample) if (wm.Marked(s)) truth.Add(s);

            EnumEngine ee = new EnumEngine(wm, sample);
            ee.Chunk = sample.Count;
            List<string> got = new List<string>();
            try { ee.Run(delegate () { return false; }, delegate (List<string> f) { got.AddRange(f); }); }
            catch (WordDiedException) { throw; }
            catch (Exception ex) { Out("#CALIB enum failed " + Com.OneLine(Com.Msg(ex))); got = null; }
            if (got != null)
            {
                bool same = SameSet(got, truth);
                Out("#CALIB sample " + sample.Count + " single " + truth.Count + " enum " + got.Count + " unmapped " + ee.Unmapped + (same ? " same" : " different"));
                if (same && ee.Unmapped == 0) return "enum";
                HashSet<string> g = new HashSet<string>(got);
                int shown = 0;
                foreach (string s in sample)
                    if (truth.Contains(s) != g.Contains(s) && shown++ < 30) Out("#CALIBDIFF\t" + s + "\t" + B(truth.Contains(s)) + "\t" + B(g.Contains(s)));
            }

            GroupEngine ge = new GroupEngine(wm, sample);
            List<string> got2 = new List<string>();
            ge.Run(delegate () { return false; }, delegate (List<string> f) { got2.AddRange(f); });
            bool same2 = SameSet(got2, truth);
            Out("#CALIB group " + got2.Count + (same2 ? " same" : " different"));
            return same2 ? "group" : "single";
        }

        static bool SameSet(List<string> a, HashSet<string> b)
        {
            HashSet<string> s = new HashSet<string>(a);
            return s.SetEquals(b);
        }

        // ------------------------------------------------------------ the main loop: batches of words, one Word window per worker
        // Batch b holds the words [start + b*size, start + (b+1)*size). Results are written in order, so the place reached (Q) is
        // always "everything before it is answered", whatever the order in which the workers finish.
        void RunParallel(int start, ref int flaggedTotal)
        {
            List<string> list = P.DictWords;
            int n = list.Count;
            int size = BatchSize;
            int nb = (n - start + size - 1) / size;
            int workers = Math.Max(1, Math.Min(Workers, nb));
            WorkersUsed = workers;
            Out("#WORKERS " + workers);

            object gate = new object();
            int nextBatch = 0, nextWrite = 0, written = 0;
            Dictionary<int, BatchDone> done = new Dictionary<int, BatchDone>();
            int flagged = flaggedTotal;
            Exception fatal = null;
            DateTime t0 = DateTime.Now, lastUi = DateTime.MinValue;

            Action<int, IWord> loop = delegate (int id, IWord first)
            {
                IWord x = first;
                string engMode = null;
                IBatchEngine eng = null;
                int restarts = 0;
                try
                {
                    if (x == null)
                    {
                        try { x = StartWorker(id); }
                        catch (Exception ex)
                        {
                            Log.Write("Worker " + id + " could not start: " + Com.Msg(ex));
                            Out("#WORKERFAILED " + id + " " + Com.OneLine(Com.Msg(ex)));
                            return;
                        }
                    }
                    while (true)
                    {
                        int b;
                        lock (gate)
                        {
                            if (Cancel || fatal != null || nextBatch >= nb) break;
                            b = nextBatch++;
                        }
                        int lo = start + b * size, hi = Math.Min(n, lo + size);
                        Stopwatch sw = Stopwatch.StartNew();
                        List<string> found = null;
                        while (found == null)
                        {
                            if (eng == null || engMode != curMode) { engMode = curMode; eng = MakeEngine(engMode, x, list); }
                            try { found = eng.FindIn(lo, hi); }
                            catch (WordDiedException ex)
                            {
                                Log.Write("Word died (worker " + id + "): " + ex.Message);
                                Out("#WORDDIED " + id + " " + lo);
                                if (++restarts > 5) throw;
                                Status("וורד נסגר באמצע. מפעיל אותו מחדש…");
                                CloseWordQuiet(x);
                                x = null; eng = null;
                                Thread.Sleep(PauseMs);
                                x = StartWorker(id);
                                if (id == 0) wm = x;
                            }
                            catch (Exception ex)
                            {
                                if (engMode != "enum") throw;
                                // the fast way failed in the middle: go on the careful way
                                lock (gate)
                                {
                                    if (curMode == "enum")
                                    {
                                        Log.Write("Fast way failed, switching to the careful way: " + Com.Msg(ex));
                                        Out("#MODE group");
                                        Out("#SWITCH " + Com.OneLine(Com.Msg(ex)));
                                        curMode = "group";
                                    }
                                }
                                eng = null;
                            }
                        }
                        long ms = sw.ElapsedMilliseconds;
                        lock (gate)
                        {
                            done[b] = new BatchDone { Lo = lo, Hi = hi, Found = found, Ms = ms, Worker = id };
                            BatchDone bd;
                            while (done.TryGetValue(nextWrite, out bd))
                            {
                                done.Remove(nextWrite);
                                StringBuilder sb = new StringBuilder();
                                foreach (string f in bd.Found) sb.Append("G\t" + f + "\n");
                                sb.Append("Q\t" + bd.Hi + "\n");
                                sb.Append("T\t" + bd.Hi + "\t" + bd.Ms + "\t" + bd.Found.Count + "\t" + bd.Worker + "\n");
                                if (nextWrite % 10 == 0) sb.Append("#CLOCK " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + bd.Hi + "\n");
                                WriteRaw(sb.ToString());
                                flagged += bd.Found.Count;
                                written = bd.Hi - start;
                                nextWrite++;
                                if ((DateTime.Now - lastUi).TotalMilliseconds > 700 || bd.Hi >= n)
                                {
                                    lastUi = DateTime.Now;
                                    double secs = Math.Max(1.0, (DateTime.Now - t0).TotalSeconds);
                                    if (OnProgress != null) OnProgress(2, bd.Hi, n, written / secs);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Write("Worker " + id + " stopped with an error: " + Com.Msg(ex));
                    lock (gate) { if (fatal == null) fatal = ex; }
                }
                finally
                {
                    if (id == 0) { wm = x; }                     // the main Word (maybe a new one) is closed by the caller
                    else CloseWordQuiet(x);
                }
            };

            List<Thread> threads = new List<Thread>();
            for (int id = 1; id < workers; id++)
            {
                int myId = id;
                Thread t = new Thread(delegate () { loop(myId, null); });
                t.SetApartmentState(ApartmentState.STA);
                t.IsBackground = true;
                t.Start();
                threads.Add(t);
            }
            loop(0, wm);
            foreach (Thread t in threads) t.Join();
            flaggedTotal = flagged;
            if (fatal != null) throw new InvalidOperationException(Com.Msg(fatal), fatal);
        }
    }

    // ---------------------------------------------------------------- window
    public class MainForm : Form
    {
        Label intro, status;
        ProgressBar bar;
        Button btnStart, btnStop, btnFolder;
        Thread worker;
        Job job;
        Pool pool;
        string resultPath;
        int workers;
        bool closing;
        bool jobRunning;
        System.Windows.Forms.Timer forceTimer;

        public const string Title = "בדיקת מילים 4 מול וורד";

        public MainForm(Pool p, int workers)
        {
            pool = p;
            this.workers = workers;
            resultPath = ChoosePath();

            Text = Title;
            Font = new Font("Segoe UI", 10f);
            ClientSize = new Size(620, 350);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            RightToLeft = RightToLeft.Yes;

            intro = new Label();
            intro.SetBounds(16, 14, 588, 160);
            intro.RightToLeft = RightToLeft.Yes;
            intro.Text = "התוכנה בודקת אילו מילים מהספרים, שעוד אינן במילון התורני, וורד מסמן כשגיאה.\r\n"
                       + "יש לבדוק " + pool.DictWords.Count.ToString("N0") + " מילים, והבדיקה נמשכת בערך שעה.\r\n"
                       + "בזמן הבדיקה המילון התורני מנותק מוורד (כך וורד בודק מהר יותר), ובסוף התוכנה מחברת אותו בחזרה.\r\n\r\n"
                       + "לפני שמתחילים: שמרו וסגרו את וורד.\r\n"
                       + "בזמן הבדיקה: אל תפתחו את וורד, ואל תפעילו את תוכנת ההתקנה של המילון.\r\n"
                       + "המחשב לא יירדם בזמן הבדיקה. אפשר לעצור באמצע ולהמשיך אחר כך מאותו מקום.";
            Controls.Add(intro);

            bar = new ProgressBar();
            bar.SetBounds(16, 184, 588, 22);
            Controls.Add(bar);

            status = new Label();
            status.SetBounds(16, 214, 588, 70);
            status.RightToLeft = RightToLeft.Yes;
            status.Text = "מוכן. לחצו על \"התחל בדיקה\".";
            Controls.Add(status);

            btnStart = new Button();
            btnStart.SetBounds(478, 296, 126, 40);
            btnStart.Text = "התחל בדיקה";
            btnStart.RightToLeft = RightToLeft.Yes;
            btnStart.Click += delegate (object s, EventArgs e) { StartJob(); };
            Controls.Add(btnStart);

            btnStop = new Button();
            btnStop.SetBounds(342, 296, 126, 40);
            btnStop.Text = "עצור ושמור";
            btnStop.RightToLeft = RightToLeft.Yes;
            btnStop.Enabled = false;
            btnStop.Click += delegate (object s, EventArgs e) { StopJob(); };
            Controls.Add(btnStop);

            btnFolder = new Button();
            btnFolder.SetBounds(16, 296, 200, 40);
            btnFolder.Text = "הצג את הקובץ";
            btnFolder.RightToLeft = RightToLeft.Yes;
            btnFolder.Visible = false;
            btnFolder.Click += delegate (object s, EventArgs e) { ShowFile(); };
            Controls.Add(btnFolder);

            FormClosing += OnClosing;
        }

        static string ChoosePath()
        {
            string name = "Word_check_results_4.txt";
            string[] dirs = new string[] {
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Path.GetTempPath() };
            foreach (string d in dirs)
            {
                try
                {
                    if (string.IsNullOrEmpty(d) || !Directory.Exists(d)) continue;
                    string probe = Path.Combine(d, ".wordcheck-write-test");
                    File.WriteAllText(probe, "x");
                    File.Delete(probe);
                    return Path.Combine(d, name);
                }
                catch (Exception) { }
            }
            return Path.Combine(Path.GetTempPath(), name);
        }

        void SetStatus(string s)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try { BeginInvoke((MethodInvoker)delegate { status.Text = s; }); } catch (Exception) { }
        }

        static string TimeLeft(double secs)
        {
            if (secs < 90) return "פחות מדקה וחצי";
            if (secs < 5400) return "כ-" + Math.Round(secs / 60.0).ToString("0") + " דקות";
            return "כ-" + (secs / 3600.0).ToString("0.0") + " שעות";
        }

        void SetProgress(int part, int done, int total, double rate)
        {
            if (IsDisposed || !IsHandleCreated) return;
            try
            {
                BeginInvoke((MethodInvoker)delegate
                {
                    bar.Maximum = Math.Max(1, total);
                    bar.Value = Math.Max(0, Math.Min(bar.Maximum, done));
                    status.Text = "נבדקו " + done.ToString("N0") + " מתוך " + total.ToString("N0") + " מילים"
                                + (rate > 0.2 ? "\r\nנשאר: " + TimeLeft((total - done) / rate) : "");
                });
            }
            catch (Exception) { }
        }

        void StartJob()
        {
            if (jobRunning) return;
            try
            {
                if (Process.GetProcessesByName("WINWORD").Length > 0)
                {
                    DialogResult r = MessageBox.Show(this,
                        "נראה שוורד פתוח כרגע. שמרו וסגרו אותו, ואז לחצו \"אישור\" כדי להתחיל.\r\n(אם לא רואים חלון של וורד, אפשר פשוט ללחוץ \"אישור\".)",
                        Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    if (r != DialogResult.OK) return;
                }
            }
            catch (Exception) { }
            jobRunning = true;
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            btnFolder.Visible = false;
            job = new Job();
            job.P = pool;
            job.ResultPath = resultPath;
            job.Workers = workers;
            job.OnStatus = SetStatus;
            job.OnProgress = SetProgress;
            worker = new Thread(delegate ()
            {
                job.Run();
                try { BeginInvoke((MethodInvoker)Finished); } catch (Exception) { }
            });
            worker.SetApartmentState(ApartmentState.STA);
            worker.IsBackground = true;
            worker.Start();
        }

        void StopJob()
        {
            if (job != null) job.Cancel = true;
            btnStop.Enabled = false;
            status.Text = "עוצר ושומר… (אם המילון התורני נותק, הוא מחובר בחזרה)";
        }

        void Finished()
        {
            jobRunning = false;
            if (forceTimer != null) forceTimer.Stop();
            btnStop.Enabled = false;
            btnStart.Enabled = true;
            btnFolder.Visible = File.Exists(resultPath);
            string restoreWarning = job.RestoreProblem
                ? "\r\n\r\nשימו לב: לא הצלחתי לחבר בחזרה את המילון התורני לוורד. הפעילו את התוכנה הזאת שוב, והיא תנסה לחבר אותו. אם זה לא עוזר, התקינו מחדש את המילון התורני."
                : "";
            if (closing) { Close(); return; }
            if (job.Outcome == "done")
            {
                bar.Value = bar.Maximum;
                status.Text = "הבדיקה הסתיימה.";
                Say("הבדיקה הסתיימה.\r\n\r\nהקובץ נשמר כאן:\r\n" + resultPath + "\r\n\r\nגררו אותו לשיחה עם קלוד." + restoreWarning,
                    job.RestoreProblem ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            else if (job.Outcome == "alreadydone")
            {
                status.Text = "הבדיקה כבר הסתיימה בעבר.";
                Say("הבדיקה כבר הסתיימה בעבר. הקובץ נמצא כאן:\r\n" + resultPath + "\r\n\r\nגררו אותו לשיחה עם קלוד." + restoreWarning, MessageBoxIcon.Information);
            }
            else if (job.Outcome == "stopped")
            {
                status.Text = "הבדיקה נעצרה ונשמרה. הפעלה חוזרת של התוכנה תמשיך מאותו מקום.";
                if (job.RestoreProblem) Say(restoreWarning.Trim(), MessageBoxIcon.Warning);
            }
            else
            {
                status.Text = "הבדיקה לא הושלמה.";
                Say(job.ErrorText + restoreWarning, MessageBoxIcon.Warning);
            }
        }

        void Say(string text, MessageBoxIcon icon)
        {
            MessageBox.Show(this, text, Title, MessageBoxButtons.OK, icon, MessageBoxDefaultButton.Button1,
                MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
        }

        void ShowFile()
        {
            try { Process.Start("explorer.exe", "/select,\"" + resultPath + "\""); } catch (Exception) { }
        }

        void OnClosing(object sender, FormClosingEventArgs e)
        {
            if (jobRunning)
            {
                e.Cancel = true;
                if (!closing)
                {
                    DialogResult r = MessageBox.Show(this, "הבדיקה עדיין רצה. לעצור ולשמור את ההתקדמות?", Title,
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    if (r == DialogResult.Yes)
                    {
                        closing = true;
                        StopJob();
                        // if Word does not answer for a long while, give up waiting (the saved progress stays,
                        // and the next start connects the Torah dictionary back if it was disconnected)
                        forceTimer = new System.Windows.Forms.Timer();
                        forceTimer.Interval = 180000;
                        forceTimer.Tick += delegate (object s2, EventArgs e2)
                        {
                            Log.Write("forced exit");
                            try { if (job != null) job.KillWords(); } catch (Exception) { }
                            Process.GetCurrentProcess().Kill();
                        };
                        forceTimer.Start();
                    }
                }
            }
        }
    }

    // ---------------------------------------------------------------- test without Word:  mono WordCheck.exe --selftest
    // A pretend Word. It behaves the way the real one was seen to behave:
    //  * a Word keeps the list of dictionaries it had when it started. A dictionary that is taken out (or put in) while it runs is in effect
    //    only for the Words that are started afterwards, unless a flag says it is in effect at once (the program has to work either way)
    //  * Word knows some words by itself (Native) and accepts the words of the dictionary files it has loaded
    public class FakeWorld
    {
        public readonly object Lock = new object();
        public HashSet<string> Native = new HashSet<string>();                 // the words Word knows by itself: the truth the check is compared with
        public Dictionary<string, HashSet<string>> Files = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        public List<string> InRegistry = new List<string>();                    // dictionary files that a Word loads when it starts (kept between Words)
        public string DictPath;                                                  // the Torah dictionary file
        public bool RuntimeDeleteEffective, RuntimeAddEffective;                 // a change is in effect in a running Word at once
        public bool DictionaryIgnored;                                           // taking the dictionary out changes nothing, even after a restart
        public int ShiftStarts;                                                  // not 0: the marks are reported at wrong places
        public bool ErrorStartsFail;                                             // the list of places fails with an ordinary error
        public bool ErrorStartsDie;                                              // the list of places says "Word is gone", every time
        public string GroupBlindPrefix;                                          // a question about several words never finds words that start like this
        public long DieAtCall = -1;                                              // Word is gone at this question (counted over all Words)
        public HashSet<int> Stubborn = new HashSet<int>();                       // numbers (in the order they are started) of Words that keep the dictionary after it was taken out
        public HashSet<int> Blind = new HashSet<int>();                          // numbers of Words whose Hebrew check does not work: they mark nothing
        public int DelayMs;                                                      // time of every "list of places" question
        public long Calls, Created, Closed, Concurrent, MaxConcurrent;
    }

    public class FakeWord : IWord
    {
        readonly FakeWorld wd;
        readonly HashSet<string> loaded = new HashSet<string>();     // words of the dictionary files this Word has loaded
        readonly List<string> mine;                                   // the dictionary files this Word has loaded
        readonly int number;                                          // 1 for the first Word that was started
        bool disposed;

        public FakeWord(FakeWorld world)
        {
            wd = world;
            lock (wd.Lock)
            {
                mine = new List<string>(wd.InRegistry);              // a Word reads the list when it starts
                Reload();
                number = (int)(++wd.Created);
                wd.Concurrent++;
                if (wd.Concurrent > wd.MaxConcurrent) wd.MaxConcurrent = wd.Concurrent;
            }
        }

        public string Version { get { return "fake"; } }
        public string Build { get { return "0"; } }
        public int Pid { get { return 0; } }

        void Reload()
        {
            loaded.Clear();
            foreach (string path in mine)
            {
                HashSet<string> f;
                if (wd.Files.TryGetValue(path, out f)) loaded.UnionWith(f);
            }
        }

        bool Bad(string x)
        {
            if (wd.Blind.Contains(number)) return false;
            if (wd.Native.Contains(x)) return false;
            if (wd.Stubborn.Contains(number)) foreach (HashSet<string> f in wd.Files.Values) if (f.Contains(x)) return false;
            if (loaded.Contains(x)) return false;
            if (wd.DictionaryIgnored) foreach (HashSet<string> f in wd.Files.Values) if (f.Contains(x)) return false;
            return true;
        }

        void Tick(bool places)
        {
            if (disposed) throw new InvalidOperationException("the pretend Word was used after it was closed");
            long n;
            lock (wd.Lock) { n = ++wd.Calls; }
            if (wd.DieAtCall > 0 && n == wd.DieAtCall) throw new WordDiedException("pretend crash", null);
            if (places)
            {
                if (wd.ErrorStartsDie) throw new WordDiedException("pretend crash, every time", null);
                if (wd.ErrorStartsFail) throw new InvalidOperationException("pretend failure of the list of places");
                if (wd.DelayMs > 0) Thread.Sleep(wd.DelayMs);
            }
        }

        public bool Marked(string text)
        {
            Tick(false);
            bool several = text.IndexOf('\r') >= 0;
            foreach (string x in text.Split('\r'))
            {
                if (several && wd.GroupBlindPrefix != null && x.StartsWith(wd.GroupBlindPrefix)) continue;
                if (Bad(x)) return true;
            }
            return false;
        }

        public List<int> ErrorStarts(string text)
        {
            Tick(true);
            List<int> r = new List<int>();
            int pos = 0;
            foreach (string x in text.Split('\r'))
            {
                if (Bad(x)) r.Add(pos + wd.ShiftStarts);
                pos += x.Length + 1;
            }
            return r;
        }

        public string DictionaryNames()
        {
            StringBuilder sb = new StringBuilder();
            foreach (string p in mine) { if (sb.Length > 0) sb.Append(", "); sb.Append(Path.GetFileName(p)); }
            return sb.ToString();
        }

        public List<string> FindOurs()
        {
            List<string> l = new List<string>();
            foreach (string p in mine) if (Names.IsOurs(Path.GetFileName(p))) l.Add(p);
            return l;
        }

        public List<string> FindOthers()
        {
            List<string> l = new List<string>();
            foreach (string p in mine) if (!Names.IsOurs(Path.GetFileName(p))) l.Add(p);
            return l;
        }

        public bool HasDictionary(string fileName)
        {
            foreach (string p in mine) if (string.Equals(Path.GetFileName(p), fileName, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        public void Detach(string fileName)
        {
            lock (wd.Lock)
            {
                mine.RemoveAll(delegate (string p) { return string.Equals(Path.GetFileName(p), fileName, StringComparison.OrdinalIgnoreCase); });
                wd.InRegistry.RemoveAll(delegate (string p) { return string.Equals(Path.GetFileName(p), fileName, StringComparison.OrdinalIgnoreCase); });
                if (wd.RuntimeDeleteEffective) Reload();
            }
        }

        public void Attach(string fullPath)
        {
            if (!File.Exists(fullPath)) return;
            lock (wd.Lock)
            {
                if (!mine.Contains(fullPath)) mine.Add(fullPath);
                if (!wd.InRegistry.Contains(fullPath)) wd.InRegistry.Add(fullPath);
                if (wd.RuntimeAddEffective) Reload();
            }
        }

        public void Dispose()
        {
            if (disposed) return;
            disposed = true;
            lock (wd.Lock) { wd.Closed++; wd.Concurrent--; }
        }
    }

    public static class SelfTest
    {
        static int fails;
        static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);
        static void Check(bool ok, string what) { if (!ok) { fails++; Console.WriteLine("FAILED: " + what); } }

        class SimpleFake : IMarker
        {
            public HashSet<string> Bad;
            public long Calls;
            public bool Marked(string text) { Calls++; foreach (string x in text.Split('\r')) if (Bad.Contains(x)) return true; return false; }
            public List<int> ErrorStarts(string text)
            {
                Calls++;
                List<int> r = new List<int>();
                int pos = 0;
                foreach (string x in text.Split('\r')) { if (Bad.Contains(x)) r.Add(pos); pos += x.Length + 1; }
                return r;
            }
        }

        // every third list of places has one extra mark that sits on a line break, not on a word
        class WeirdFake : IMarker
        {
            public HashSet<string> Bad;
            int n;
            public bool Marked(string text) { foreach (string x in text.Split('\r')) if (Bad.Contains(x)) return true; return false; }
            public List<int> ErrorStarts(string text)
            {
                List<int> r = new List<int>();
                int pos = 0;
                foreach (string x in text.Split('\r')) { if (Bad.Contains(x)) r.Add(pos); pos += x.Length + 1; }
                int cr = text.IndexOf('\r');
                if (cr >= 0 && ++n % 3 == 0) r.Add(cr);
                return r;
            }
        }

        static List<string> RunEngine(IEngine e)
        {
            List<string> found = new List<string>();
            e.Run(delegate () { return false; }, delegate (List<string> f) { found.AddRange(f); });
            return found;
        }

        static List<string> RunBlocks(IBatchEngine e, int n, int size)
        {
            List<string> found = new List<string>();
            for (int lo = 0; lo < n; lo += size) found.AddRange(e.FindIn(lo, Math.Min(n, lo + size)));
            return found;
        }

        static bool Same(List<string> found, HashSet<string> truth)
        {
            return new HashSet<string>(found).SetEquals(truth) && found.Count == truth.Count;
        }

        static void EngineTests()
        {
            Random rnd = new Random(7);
            string[] shapes = new string[] { "מילה", "רש\"י", "תוס'", "ב", "ודאמרינן", "א\"ב" };
            double[] rates = new double[] { 0.0, 0.004, 0.02, 0.06, 0.15, 0.4, 1.0 };
            foreach (double rate in rates)
            {
                List<string> words = new List<string>();
                HashSet<string> bad = new HashSet<string>();
                for (int i = 0; i < 20000; i++)
                {
                    string x = shapes[rnd.Next(shapes.Length)] + i;
                    words.Add(x);
                    if (rnd.NextDouble() < rate) bad.Add(x);
                }
                SimpleFake f1 = new SimpleFake(); f1.Bad = bad;
                List<string> g = RunEngine(new GroupEngine(f1, words));
                SimpleFake f2 = new SimpleFake(); f2.Bad = bad;
                EnumEngine ee = new EnumEngine(f2, words);
                List<string> en = RunEngine(ee);
                SimpleFake f3 = new SimpleFake(); f3.Bad = bad;
                GroupEngine single = new GroupEngine(f3, words); single.SingleOnly = true;
                List<string> si = RunEngine(single);
                Check(Same(g, bad), "group engine, rate " + rate);
                Check(Same(en, bad) && ee.Unmapped == 0 && ee.Fallbacks == 0, "enum engine, rate " + rate);
                Check(Same(si, bad), "single engine, rate " + rate);
                Console.WriteLine("rate {0:0.000}: bad={1,6}  questions: careful={2,6}  fast={3,4}  word-by-word={4,6}", rate, bad.Count, f1.Calls, f2.Calls, f3.Calls);

                // the same through the block interface that the main loop uses (blocks of 250, the last one shorter)
                SimpleFake f4 = new SimpleFake(); f4.Bad = bad;
                SimpleFake f5 = new SimpleFake(); f5.Bad = bad;
                SimpleFake f6 = new SimpleFake(); f6.Bad = bad;
                GroupEngine single2 = new GroupEngine(f6, words); single2.SingleOnly = true;
                Check(Same(RunBlocks(new EnumEngine(f4, words), words.Count, 250), bad), "FindIn, enum engine, rate " + rate);
                Check(Same(RunBlocks(new GroupEngine(f5, words), words.Count, 250), bad), "FindIn, group engine, rate " + rate);
                Check(Same(RunBlocks(single2, words.Count, 250), bad), "FindIn, single engine, rate " + rate);
            }

            // a mark that sits on no word: the block is asked again the careful way, the answers stay right
            List<string> w2 = new List<string>();
            HashSet<string> b2 = new HashSet<string>();
            for (int i = 0; i < 3000; i++) { w2.Add("מלה" + i); if (i % 11 == 0) b2.Add("מלה" + i); }
            WeirdFake wf = new WeirdFake(); wf.Bad = b2;
            EnumEngine we = new EnumEngine(wf, w2);
            List<string> got = RunBlocks(we, w2.Count, 250);
            Check(Same(got, b2), "enum engine with marks off the words: answers (" + got.Count + " / " + b2.Count + ")");
            Check(we.Unmapped > 0 && we.Fallbacks > 0, "enum engine with marks off the words: noticed (" + we.Unmapped + ", " + we.Fallbacks + ")");

            // stop in the middle, then continue from the saved place
            List<string> ws = new List<string>();
            HashSet<string> bd = new HashSet<string>();
            for (int i = 0; i < 5000; i++) { ws.Add("x" + i); if (i % 37 == 0) bd.Add("x" + i); }
            foreach (int kind in new int[] { 0, 1 })
            {
                SimpleFake fm = new SimpleFake(); fm.Bad = bd;
                IEngine e1 = kind == 0 ? (IEngine)new GroupEngine(fm, ws) : new EnumEngine(fm, ws);
                List<string> found = new List<string>();
                int chunks = 0;
                e1.Run(delegate () { return chunks >= 10; }, delegate (List<string> f) { found.AddRange(f); chunks++; });
                int saved = e1.Next;
                Check(saved > 0 && saved < 5000, "stop: saved place " + saved);
                IEngine e2 = kind == 0 ? (IEngine)new GroupEngine(fm, ws) : new EnumEngine(fm, ws);
                e2.Next = saved;
                e2.Run(delegate () { return false; }, delegate (List<string> f) { found.AddRange(f); });
                Check(found.Count == bd.Count, "resume (" + kind + "): found " + found.Count + " expected " + bd.Count);
            }
        }

        static void PoolTests()
        {
            MemoryStream ms = new MemoryStream();
            using (GZipStream gz = new GZipStream(ms, CompressionMode.Compress, true))
            {
                byte[] b = Utf8.GetBytes("@C\tמיבעי\n@D\tתורה\n@D\tמיבעי\n@E\tקזקז\n@A\tשלום\n\n@D\tא\"ב\n");
                gz.Write(b, 0, b.Length);
            }
            ms.Position = 0;
            Pool p = Pool.Load(ms);
            Check(p.DictWords.Count == 3 && p.Controls.Count == 1 && p.Flagged.Count == 1 && p.Accepted.Count == 1, "pool: the four lists are read");
            Check(p.DictWords[0] == "תורה" && p.DictWords[2] == "א\"ב" && p.Controls[0] == "מיבעי", "pool: words and order");
            string id1 = p.Id;
            p.DictWords[1] = "מיבעה";
            p.MakeId();
            Check(id1.Length > 0 && id1 != p.Id, "pool: the name of the list changes with its words");
        }

        static string TempDir(string name)
        {
            string d = Path.Combine(Path.GetTempPath(), "wordcheck-selftest-" + name);
            if (Directory.Exists(d)) Directory.Delete(d, true);
            Directory.CreateDirectory(d);
            return d;
        }

        static string[] FileLines(string path) { return File.ReadAllLines(path, Encoding.UTF8); }

        static bool FileHas(string path, string start)
        {
            foreach (string line in FileLines(path)) if (line.StartsWith(start)) return true;
            return false;
        }

        static int CountLines(string path, string start)
        {
            int n = 0;
            foreach (string line in FileLines(path)) if (line.StartsWith(start)) n++;
            return n;
        }

        class Scenario
        {
            public FakeWorld World;
            public Pool Pool;
            public HashSet<string> Expect = new HashSet<string>();     // the dictionary words that Word does not know by itself
            public string Dir;
            public string ResultPath;
        }

        static Scenario MakeScenario(string name, int seed)
        {
            Scenario sc = new Scenario();
            sc.Dir = TempDir(name);
            Restore.DirOverride = Path.Combine(sc.Dir, "restore");
            sc.ResultPath = Path.Combine(sc.Dir, "results.txt");
            FakeWorld wd = new FakeWorld();
            wd.DictPath = Path.Combine(sc.Dir, "Milon_Torani.dic");
            Random rnd = new Random(seed);
            Pool p = new Pool();
            HashSet<string> dictFile = new HashSet<string>();
            for (int i = 0; i < 2000; i++)
            {
                string x = (i % 9 == 0 ? "ד\"" : "תורה") + i;
                p.DictWords.Add(x);
                dictFile.Add(x);
                if (rnd.NextDouble() < 0.3) wd.Native.Add(x); else sc.Expect.Add(x);
            }
            p.Controls.Add("מיבעי"); p.Controls.Add("באנפא");
            foreach (string c in p.Controls) { p.DictWords.Add(c); dictFile.Add(c); sc.Expect.Add(c); }
            for (int i = 0; i < 1000; i++) { p.Flagged.Add("פלג" + i); p.Accepted.Add("קבל" + i); wd.Native.Add("קבל" + i); }
            foreach (string x in new string[] { "שלום", "בית", "ספר" }) wd.Native.Add(x);
            wd.Files[wd.DictPath] = dictFile;
            wd.InRegistry.Add(wd.DictPath);
            p.MakeId();
            File.WriteAllText(wd.DictPath, "dictionary file content, long enough to be compared with its backup copy", Encoding.Unicode);
            sc.World = wd;
            sc.Pool = p;
            return sc;
        }

        static Job MakeJob(Scenario sc, int workers)
        {
            Job j = new Job();
            j.P = sc.Pool;
            j.ResultPath = sc.ResultPath;
            j.Workers = workers;
            j.BatchSize = 100;
            j.PauseMs = 0;
            FakeWorld wd = sc.World;
            j.CreateWord = delegate () { return new FakeWord(wd); };
            return j;
        }

        static void CheckResults(Scenario sc, string what, bool connectedAfter)
        {
            Dictionary<string, int> index = new Dictionary<string, int>();
            for (int i = 0; i < sc.Pool.DictWords.Count; i++) index[sc.Pool.DictWords[i]] = i;
            List<string> gl = new List<string>();
            List<int> pending = new List<int>();
            bool done = false, orderOk = true, qOk = true;
            int lastQ = 0;
            string doneLine = "";
            foreach (string line in FileLines(sc.ResultPath))
            {
                if (line.StartsWith("G\t"))
                {
                    string w = line.Substring(2);
                    gl.Add(w);
                    int ix;
                    if (index.TryGetValue(w, out ix)) pending.Add(ix); else orderOk = false;
                }
                else if (line.StartsWith("Q\t"))
                {
                    int q = int.Parse(line.Substring(2));
                    if (q <= lastQ) qOk = false;
                    foreach (int ix in pending) if (ix < lastQ || ix >= q) orderOk = false;
                    pending.Clear();
                    lastQ = q;
                }
                else if (line.StartsWith("DONE\t")) { done = true; doneLine = line; }
            }
            Check(done, what + ": DONE line");
            Check(doneLine == "DONE\t" + sc.Pool.DictWords.Count + "\t" + sc.Expect.Count, what + ": DONE numbers (" + doneLine.Replace('\t', ' ') + ")");
            HashSet<string> g = new HashSet<string>(gl);
            Check(g.SetEquals(sc.Expect) && gl.Count == sc.Expect.Count, what + ": answers (" + g.Count + " different, " + gl.Count + " lines, expected " + sc.Expect.Count + ")");
            Check(qOk && lastQ == sc.Pool.DictWords.Count, what + ": the places (Q) go up to the end (" + lastQ + ")");
            Check(orderOk && pending.Count == 0, what + ": every answer comes before the place that covers it, none twice");
            Check(sc.World.InRegistry.Contains(sc.World.DictPath) == connectedAfter, what + ": the dictionary is " + (connectedAfter ? "connected again" : "still not connected"));
            Check(!Restore.Pending(), what + ": no restore note is left");
        }

        static void CheckClosed(Scenario sc, string what)
        {
            Check(sc.World.Created == sc.World.Closed && sc.World.Concurrent == 0, what + ": every Word was closed (" + sc.World.Created + " started, " + sc.World.Closed + " closed)");
        }

        static void CheckSpeed(Scenario sc, string state)
        {
            int fl = 0, ac = 0;
            bool bad = false;
            foreach (string line in FileLines(sc.ResultPath))
            {
                if (!line.StartsWith("SPEED\t" + state + "\t")) continue;
                string[] f = line.Split('\t');                       // SPEED state class k ms marked asked
                int marked = int.Parse(f[5]), asked = int.Parse(f[6]);
                if (f[2] == "flagged") { fl++; if (marked != asked || asked < 10) bad = true; }
                else { ac++; if (marked != 0 || asked < 10) bad = true; }
            }
            Check(fl == 2 && ac == 2 && !bad, "speed lines (" + state + "): " + fl + " flagged, " + ac + " accepted" + (bad ? ", wrong numbers" : ""));
        }

        static void JobTests()
        {
            // 1. a normal run with one Word: the fast way is chosen, the controls prove the change, everything is right and closed
            Scenario sc = MakeScenario("normal", 1);
            Job j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done", "normal: outcome " + j.Outcome + " " + j.ErrorText);
            Check(FileHas(sc.ResultPath, "#MODE enum"), "normal: fast way chosen");
            Check(FileHas(sc.ResultPath, "#PROBES_DONE"), "normal: suffix questions done");
            Check(FileHas(sc.ResultPath, "#BCONTROLS before 0 of 2") && FileHas(sc.ResultPath, "#BCONTROLS during 2 of 2"), "normal: the control words prove the change");
            CheckSpeed(sc, "connected");
            CheckSpeed(sc, "detached");
            CheckResults(sc, "normal", true);
            CheckClosed(sc, "normal");
            Console.WriteLine("job normal: {0} dictionary words, {1} marked without the dictionary, {2} questions", sc.Pool.DictWords.Count, sc.Expect.Count, sc.World.Calls);
            Job again = MakeJob(sc, 1);
            again.Run();
            Check(again.Outcome == "alreadydone", "normal: the second run says already done (" + again.Outcome + ")");

            // 2. three Words at the same time: same answers, in order, never more than three Words, all closed
            sc = MakeScenario("parallel", 2);
            sc.World.DelayMs = 15;
            j = MakeJob(sc, 3);
            j.Run();
            Check(j.Outcome == "done", "parallel: outcome " + j.Outcome + " " + j.ErrorText);
            Check(FileHas(sc.ResultPath, "#WORKERS 3") && j.WorkersUsed == 3, "parallel: three workers");
            CheckResults(sc, "parallel", true);
            Check(sc.World.MaxConcurrent >= 2 && sc.World.MaxConcurrent <= 3, "parallel: Words at the same time: " + sc.World.MaxConcurrent);
            HashSet<string> who = new HashSet<string>();
            foreach (string line in FileLines(sc.ResultPath)) if (line.StartsWith("T\t")) who.Add(line.Split('\t')[4]);
            Check(who.Count >= 2, "parallel: the blocks were shared between Words (" + who.Count + ")");
            CheckClosed(sc, "parallel");

            // 3. a change of the dictionary list in a running Word may or may not be in effect at once: right either way
            for (int mask = 0; mask < 4; mask++)
            {
                sc = MakeScenario("flags" + mask, 20 + mask);
                sc.World.RuntimeDeleteEffective = (mask & 1) != 0;
                sc.World.RuntimeAddEffective = (mask & 2) != 0;
                j = MakeJob(sc, 1);
                j.Run();
                Check(j.Outcome == "done", "flags " + mask + ": outcome " + j.Outcome + " " + j.ErrorText);
                CheckResults(sc, "flags " + mask, true);
            }

            // 4. stop in the middle with three Words, then continue: nothing twice, nothing missing, the dictionary is back at the stop
            sc = MakeScenario("stop", 3);
            sc.World.DelayMs = 5;
            j = MakeJob(sc, 3);
            Job jj = j;
            j.OnProgress = delegate (int part, int done, int total, double rate) { if (done >= 600) jj.Cancel = true; };
            j.Run();
            Check(j.Outcome == "stopped", "stop: outcome " + j.Outcome + " " + j.ErrorText);
            Check(sc.World.InRegistry.Contains(sc.World.DictPath) && !Restore.Pending() && !j.RestoreProblem, "stop: the dictionary is connected back after the stop");
            Check(FileHas(sc.ResultPath, "#STOPPED 2") && !FileHas(sc.ResultPath, "DONE\t"), "stop: stopped, not finished");
            CheckClosed(sc, "stop");
            j = MakeJob(sc, 3);
            j.Run();
            Check(j.Outcome == "done", "stop: the continued run, outcome " + j.Outcome + " " + j.ErrorText);
            Check(FileHas(sc.ResultPath, "#RESUME"), "stop: the continued run says so");
            CheckResults(sc, "stop+continue", true);
            CheckClosed(sc, "stop+continue");

            // 4b. one block only: more workers are not started than there are blocks
            sc = MakeScenario("oneblock", 14);
            j = MakeJob(sc, 3);
            j.BatchSize = 5000;
            j.Run();
            Check(j.Outcome == "done" && FileHas(sc.ResultPath, "#WORKERS 1"), "oneblock: one worker (" + j.Outcome + " " + j.ErrorText + ")");
            CheckResults(sc, "oneblock", true);

            // 5a. Word is gone once in the middle: a new Word is started and the block is asked again
            sc = MakeScenario("crash", 4);
            FakeWorld w0 = sc.World;
            Job probe = MakeJob(sc, 1);
            long before = 0;
            probe.OnProgress = delegate (int part, int done, int total, double rate) { if (before == 0) before = w0.Calls; };
            probe.Run();
            sc = MakeScenario("crash", 4);
            sc.World.DieAtCall = before + 3;
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done", "crash: outcome " + j.Outcome + " " + j.ErrorText);
            Check(FileHas(sc.ResultPath, "#WORDDIED 0"), "crash: it happened in the main loop");
            CheckResults(sc, "crash", true);
            CheckClosed(sc, "crash");

            // 5b. Word is gone every time: the run ends with an error, the dictionary is connected back, the next run can continue
            sc = MakeScenario("crashes", 5);
            FakeWorld w1 = sc.World;
            j = MakeJob(sc, 1);
            j.OnProgress = delegate (int part, int done, int total, double rate) { w1.ErrorStartsDie = true; };
            j.Run();
            Check(j.Outcome == "error", "crashes: outcome " + j.Outcome);
            Check(CountLines(sc.ResultPath, "#WORDDIED") == 6 && !FileHas(sc.ResultPath, "DONE\t"), "crashes: six tries, no finish");
            Check(w1.InRegistry.Contains(w1.DictPath) && !Restore.Pending() && !j.RestoreProblem, "crashes: the dictionary is connected back");
            CheckClosed(sc, "crashes");
            w1.ErrorStartsDie = false;
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done", "crashes: the next run, outcome " + j.Outcome + " " + j.ErrorText);
            CheckResults(sc, "crashes+continue", true);

            // 5c. Word is gone while the dictionary is out (in the speed test): error, the dictionary is back, the next run finishes
            sc = MakeScenario("crash3", 15);
            FakeWorld w4 = sc.World;
            Job spy = MakeJob(sc, 1);
            long atSpeed = 0;
            spy.OnStatus = delegate (string s) { if (atSpeed == 0 && s.StartsWith("מודד את מהירות וורד (המילון מנותק)")) atSpeed = w4.Calls; };
            spy.Run();
            sc = MakeScenario("crash3", 15);
            sc.World.DieAtCall = atSpeed + 2;
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "error" && FileHas(sc.ResultPath, "#ERROR"), "crash3: outcome " + j.Outcome);
            Check(sc.World.InRegistry.Contains(sc.World.DictPath) && !Restore.Pending() && !j.RestoreProblem, "crash3: the dictionary is connected back");
            Check(FileHas(sc.ResultPath, "#RESTORED ok"), "crash3: the restore is written in the file");
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done", "crash3: the next run, outcome " + j.Outcome + " " + j.ErrorText);
            CheckResults(sc, "crash3+continue", true);

            // 6. an earlier run was killed while the dictionary was out (and its file was emptied): the next start repairs
            sc = MakeScenario("repair", 6);
            Restore.Prepare(new List<string> { sc.World.DictPath });
            sc.World.InRegistry.Clear();
            File.WriteAllText(sc.World.DictPath, "");
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done", "repair: outcome " + j.Outcome + " " + j.ErrorText);
            Check(FileHas(sc.ResultPath, "#REPAIRED ok"), "repair: written in the file");
            Check(new FileInfo(sc.World.DictPath).Length > 10, "repair: the file was copied back");
            CheckResults(sc, "repair", true);

            // 7. disconnecting has no effect (Word keeps accepting the dictionary words): nothing is claimed, the dictionary is connected
            sc = MakeScenario("noeffect", 7);
            sc.World.DictionaryIgnored = true;
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "error" && FileHas(sc.ResultPath, "BSKIP\tno effect") && !FileHas(sc.ResultPath, "DONE\t") && !FileHas(sc.ResultPath, "BDONE"), "noeffect: skipped, not finished (" + j.Outcome + ")");
            Check(sc.World.InRegistry.Contains(sc.World.DictPath) && !Restore.Pending() && !j.RestoreProblem, "noeffect: the dictionary is connected");
            Check(!MakeJobAlreadyDone(sc), "noeffect: a skipped run does not count as finished");

            // 8. the Torah dictionary is not connected to Word at all: the check still runs, nothing is connected afterwards
            sc = MakeScenario("notconnected", 8);
            sc.World.InRegistry.Clear();
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done" && FileHas(sc.ResultPath, "#BNOTCONNECTED") && FileHas(sc.ResultPath, "BDONE"), "notconnected: done (" + j.Outcome + " " + j.ErrorText + ")");
            Check(!j.RestoreProblem, "notconnected: no problem with the restore");
            CheckSpeed(sc, "notconnected");
            CheckSpeed(sc, "notconnected2");
            CheckResults(sc, "notconnected", false);

            // 9. Word reports a dictionary whose file cannot be found: it is not touched, the check does not start
            sc = MakeScenario("nofile", 9);
            string missing = Path.Combine(sc.Dir, "missing", "Milon_Torani.dic");
            sc.World.Files[missing] = sc.World.Files[sc.World.DictPath];
            sc.World.InRegistry.Clear();
            sc.World.InRegistry.Add(missing);
            sc.World.DictPath = missing;
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "error" && FileHas(sc.ResultPath, "BSKIP\tdictionary file not found") && !FileHas(sc.ResultPath, "DONE\t"), "nofile: skipped (" + j.Outcome + " " + j.ErrorText + ")");
            Check(sc.World.InRegistry.Count == 1 && sc.World.InRegistry[0] == missing && !Restore.Pending() && !j.RestoreProblem, "nofile: nothing was disconnected");

            // 10. the fast way breaks in the middle of the run: the careful way takes over, the answers stay right
            sc = MakeScenario("switch", 11);
            FakeWorld w2 = sc.World;
            j = MakeJob(sc, 1);
            j.OnProgress = delegate (int part, int done, int total, double rate) { w2.ErrorStartsFail = true; };
            j.Run();
            Check(j.Outcome == "done", "switch: outcome " + j.Outcome + " " + j.ErrorText);
            Check(FileHas(sc.ResultPath, "#MODE enum") && FileHas(sc.ResultPath, "#MODE group") && FileHas(sc.ResultPath, "#SWITCH") && j.FinalMode == "group", "switch: changed to the careful way (" + j.FinalMode + ")");
            CheckResults(sc, "switch", true);

            // 10b. Word reports wrong places from the start: the calibration chooses the careful way
            sc = MakeScenario("shifted", 12);
            sc.World.ShiftStarts = -1;
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done" && FileHas(sc.ResultPath, "#MODE group"), "shifted: careful way chosen (" + j.Outcome + " " + j.ErrorText + ")");
            CheckResults(sc, "shifted", true);

            // 10c. wrong places, and questions about several words miss some words: only word by word is right
            sc = MakeScenario("single", 13);
            sc.World.ShiftStarts = -1;
            sc.World.GroupBlindPrefix = "ד\"";
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done" && FileHas(sc.ResultPath, "#MODE single"), "single: word by word chosen (" + j.Outcome + " " + j.ErrorText + ")");
            CheckResults(sc, "single", true);

            // 11. an earlier run was killed in the middle of writing a block: half a line and the answers of an unfinished block are dropped
            sc = MakeScenario("killed", 16);
            j = MakeJob(sc, 1);
            Job j3 = j;
            j.OnProgress = delegate (int part, int done, int total, double rate) { if (done >= 700) j3.Cancel = true; };
            j.Run();
            Check(j.Outcome == "stopped", "killed: first run, outcome " + j.Outcome + " " + j.ErrorText);
            File.AppendAllText(sc.ResultPath, "G\tתורה1\nG\tתורה2\nG\tתו", Utf8);
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done", "killed: continued run, outcome " + j.Outcome + " " + j.ErrorText);
            CheckResults(sc, "killed", true);

            // 11b. a worker's new Word that still knows the dictionary (it must not be used), or whose Hebrew check does not work:
            //      it is closed and another one is started; if that does not help either, the worker drops out and the others go on
            sc = MakeScenario("stubborn", 21);
            sc.World.DelayMs = 15;
            sc.World.Stubborn.Add(3);
            j = MakeJob(sc, 2);
            j.Run();
            Check(j.Outcome == "done" && FileHas(sc.ResultPath, "#WORKERS 2") && !FileHas(sc.ResultPath, "#WORKERFAILED"), "stubborn: the second try worked (" + j.Outcome + " " + j.ErrorText + ")");
            Check(sc.World.Created >= 5, "stubborn: a replacement Word was started (" + sc.World.Created + ")");
            CheckResults(sc, "stubborn", true);
            CheckClosed(sc, "stubborn");

            sc = MakeScenario("stubborn2", 22);
            sc.World.DelayMs = 15;
            sc.World.Stubborn.Add(3);
            sc.World.Stubborn.Add(4);
            j = MakeJob(sc, 2);
            j.Run();
            Check(j.Outcome == "done" && FileHas(sc.ResultPath, "#WORKERFAILED 1"), "stubborn2: the second worker dropped out (" + j.Outcome + " " + j.ErrorText + ")");
            CheckResults(sc, "stubborn2", true);
            CheckClosed(sc, "stubborn2");

            sc = MakeScenario("blind", 23);
            sc.World.DelayMs = 15;
            sc.World.InRegistry.Clear();            // the dictionary is not connected: only the Hebrew check tells a good Word from a bad one
            sc.World.Blind.Add(3);
            j = MakeJob(sc, 2);
            j.Run();
            Check(j.Outcome == "done" && !FileHas(sc.ResultPath, "#WORKERFAILED"), "blind: the second try worked (" + j.Outcome + " " + j.ErrorText + ")");
            CheckResults(sc, "blind", false);
            CheckClosed(sc, "blind");

            // 12. the person's own dictionary stays connected: its words that are also words of the Torah dictionary are listed (H lines),
            //     so that the analysis does not take them for words that Word knows by itself
            sc = MakeScenario("other", 17);
            string own = Path.Combine(sc.Dir, "RoamingCustom.dic");
            string o1 = sc.Pool.DictWords[3], o2 = sc.Pool.DictWords[1500];
            File.WriteAllText(own, o1 + "\r\n" + o2 + "\r\nמילהשאינהבמילון\r\n", Encoding.Unicode);     // UTF-16 with a byte order mark
            sc.World.Files[own] = new HashSet<string> { o1, o2, "מילהשאינהבמילון" };
            sc.World.InRegistry.Add(own);
            sc.Expect.Remove(o1);
            sc.Expect.Remove(o2);
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done", "other: outcome " + j.Outcome + " " + j.ErrorText);
            HashSet<string> hs = new HashSet<string>();
            foreach (string line in FileLines(sc.ResultPath)) if (line.StartsWith("H\t")) hs.Add(line.Substring(2));
            Check(hs.Count == 2 && hs.Contains(o1) && hs.Contains(o2), "other: the words of the other dictionary are listed (" + hs.Count + ")");
            Check(FileHas(sc.ResultPath, "#OTHERDICT " + own + " words 3 also in the dictionary 2"), "other: the dictionary is described");
            CheckResults(sc, "other", true);
            Check(sc.World.InRegistry.Contains(own), "other: the other dictionary was not touched");

            // 12b. the same dictionary, written without a byte order mark (UTF-8) and with an unreadable (missing) second one
            sc = MakeScenario("other2", 18);
            string own2 = Path.Combine(sc.Dir, "Custom.dic");
            string gone = Path.Combine(sc.Dir, "gone", "Other.dic");
            string o3 = sc.Pool.DictWords[7];
            File.WriteAllText(own2, o3 + "\n", Utf8);
            sc.World.Files[own2] = new HashSet<string> { o3 };
            sc.World.InRegistry.Add(own2);
            sc.World.InRegistry.Add(gone);
            sc.Expect.Remove(o3);
            j = MakeJob(sc, 1);
            j.Run();
            Check(j.Outcome == "done" && FileHas(sc.ResultPath, "H\t" + o3) && FileHas(sc.ResultPath, "#OTHERDICT " + gone + " not found"), "other2: UTF-8 file read, missing file noted (" + j.Outcome + ")");
            CheckResults(sc, "other2", true);

            Restore.DirOverride = null;
        }

        // a process number that now belongs to another program (not Word) is never waited for / ended
        static void ProcessTests()
        {
            Process other = null;
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("sleep", "20");
                psi.UseShellExecute = false;
                other = Process.Start(psi);
            }
            catch (Exception ex) { Console.WriteLine("(process test skipped: " + ex.Message + ")"); return; }
            try
            {
                Job.WaitGone(other.Id, 300);
                other.Refresh();
                Check(!other.HasExited, "a process that is not Word is not ended by the waiting for Word");
            }
            finally { try { other.Kill(); } catch (Exception) { } }
        }

        // is the result file of this scenario taken as finished by a new run?
        static bool MakeJobAlreadyDone(Scenario sc)
        {
            return ResultFile.Read(sc.ResultPath, sc.Pool.Id).Done;
        }

        static void ResultFileTests()
        {
            string dir = TempDir("rf");
            string p = Path.Combine(dir, "r.txt");
            File.WriteAllText(p, "#WORDCHECK 3\n#POOLID abc\n#MODE enum\n#PROBES_DONE\n#SPEED_DONE connected\nG\tא\nG\tב\nQ\t3\nT\t3\t5\t2\t0\n#MODE group\nBSKIP\tno effect\nBDONE\t3\t2\n", Utf8);
            ResultFile rf = ResultFile.Read(p, "abc");
            Check(rf.SamePool && rf.NextB == 3 && rf.FlaggedB == 2 && rf.HasProbes && rf.HasSpeedConn && !rf.HasSpeedDet && rf.Mode == "group", "result file: read");
            Check(!rf.Done, "result file: BSKIP and BDONE are not a finish");
            Check(!ResultFile.Read(p, "zzz").SamePool, "result file: another list of words is detected");
            File.AppendAllText(p, "#SPEED_DONE notconnected2\nDONE\t3\t2\n", Utf8);
            rf = ResultFile.Read(p, "abc");
            Check(rf.Done && rf.HasSpeedDet, "result file: DONE and the second speed test");
            Check(!ResultFile.Read(Path.Combine(dir, "none.txt"), "abc").SamePool, "result file: a missing file");
            Check(!rf.HasOther, "result file: the other dictionaries are not looked at yet");
            File.AppendAllText(p, "#OTHERDICT none\n", Utf8);
            Check(ResultFile.Read(p, "abc").HasOther, "result file: the other dictionaries were looked at");

            string q = Path.Combine(dir, "t.txt");
            File.WriteAllText(q, "#POOLID abc\nG\tא\nQ\t1\nT\t1\t1\t1\t0\nG\tב\nG\tג\n#WORDDIED 0 5\nG\tדד", Utf8);
            Check(ResultFile.TrimTail(q), "trim: the file was changed");
            Check(File.ReadAllText(q, Utf8) == "#POOLID abc\nG\tא\nQ\t1\nT\t1\t1\t1\t0\n#WORDDIED 0 5\n", "trim: what is left");
            Check(!ResultFile.TrimTail(q), "trim: a clean file is not changed");
        }

        public static int Run()
        {
            EngineTests();
            PoolTests();
            ResultFileTests();
            JobTests();
            ProcessTests();
            try
            {
                Pool pool = Pool.LoadEmbedded();
                Console.WriteLine("embedded list: {0:N0} forms to ask about, {1} control words, {2} + {3} speed-test words, id {4}",
                    pool.DictWords.Count, pool.Controls.Count, pool.Flagged.Count, pool.Accepted.Count, pool.Id);
                HashSet<string> all = new HashSet<string>(pool.DictWords);
                Check(pool.DictWords.Count > 100000 && all.Count == pool.DictWords.Count, "embedded list: too small, or a word twice");
                // WordCheck 4: the forms asked about are NOT words of the dictionary, the control words are (make_pool4.py checks both
                // against the four editions) - so no control word may be among the forms
                bool controlsOk = pool.Controls.Count == 13;
                foreach (string c in pool.Controls) if (all.Contains(c)) controlsOk = false;
                Check(controlsOk, "embedded list: 13 control words, none of them among the forms asked about");
                Check(pool.Flagged.Count >= 1000 && pool.Accepted.Count >= 1000, "embedded list: speed-test words");
                bool overlap = false;
                foreach (string w in pool.Flagged) if (all.Contains(w)) overlap = true;
                foreach (string w in pool.Accepted) if (all.Contains(w)) overlap = true;
                Check(!overlap, "embedded list: speed-test words must not be dictionary words");
            }
            catch (Exception ex) { Check(false, "no embedded word list: " + ex.Message); }
            Console.WriteLine(fails == 0 ? "SELFTEST PASSED" : "SELFTEST FAILED (" + fails + ")");
            return fails == 0 ? 0 : 1;
        }
    }

    public static class Program
    {
        [STAThread]
        public static int Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "--selftest") return SelfTest.Run();
            int workers = Math.Max(1, Math.Min(3, Environment.ProcessorCount / 2));
            for (int i = 0; i + 1 < args.Length; i++)
            {
                int n;
                if (args[i] == "--workers" && int.TryParse(args[i + 1], out n) && n >= 1 && n <= 8) workers = n;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            bool first;
            Mutex single = new Mutex(true, "Local\\WordCheck4_SingleInstance", out first);
            if (!first)
            {
                MessageBox.Show("התוכנה כבר פועלת (אולי בחלון אחר). אי אפשר להפעיל אותה פעמיים.", MainForm.Title,
                    MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                return 1;
            }
            Pool pool;
            try { pool = Pool.LoadEmbedded(); }
            catch (Exception ex) { MessageBox.Show("שגיאה בטעינת רשימת המילים: " + ex.Message); return 1; }
            Application.Run(new MainForm(pool, workers));
            GC.KeepAlive(single);
            return 0;
        }
    }
}
