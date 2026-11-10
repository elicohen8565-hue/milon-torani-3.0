using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Win32;

namespace OtzariaDict
{
    // ---------------------------------------------------------------- files, words, log
    public static class Core
    {
        public const string AppVersion = "3.1.0.1";

        // ONE dictionary file (from 3.0.0.2). With five files Word 365 used only the first ones,
        // while one file of 102,573 words (version 2.0) works. Everything goes into this file, most important words first.
        public const string MainFile = "Milon_Torani.dic";

        // files of 3.0.0.0 / 3.0.0.1 - never written any more, only removed
        public const string BasicFile = "Otzaria_Torah_Basic.dic";
        public const string MoreFile = "Otzaria_Torah_Extended.dic";
        public const string More2File = "Otzaria_Torah_Extended2.dic";
        public const string AcroFile = "Otzaria_Torah_Acronyms.dic";
        public const string Acro4File = "Otzaria_Torah_Acronyms4.dic";
        public const string ComposeFile = "Otzaria_Torah_Composition.dic";
        public static readonly string[] LegacyFiles = new string[] { BasicFile, MoreFile, More2File, AcroFile, Acro4File, ComposeFile };
        public static readonly string[] OurFiles = new string[] { MainFile, BasicFile, MoreFile, More2File, AcroFile, Acro4File, ComposeFile };

        public const string BasicRes = "Otzaria.basic.txt.gz";
        public const string MoreRes = "Otzaria.more.txt.gz";
        public const string BaseRes = "Otzaria.hebrew_base.txt.gz";
        public const string CountsRes = "Otzaria.counts.txt";   // exact word count of each edition, written by make_assets.py
        public const string RetiredRes = "Otzaria.retired.txt.gz";   // the words earlier versions had and this one dropped (data/retired_list.txt)

        // Aramaic-style connecting letters put in front of ordinary Hebrew words
        public static readonly string[] Connectors = new string[] { "ד", "וד", "כד", "וכד", "מד", "ומד" };

        public static string LogPath = Path.Combine(Path.GetTempPath(), "milon-torani-log.txt");
        public static string UProofOverride = null;   // used only by the self-test

        public static string UProofDir()
        {
            if (UProofOverride != null) return UProofOverride;
            string appdata = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(Path.Combine(appdata, "Microsoft"), "UProof");
        }

        public static bool IsOurs(string fileName)
        {
            if (fileName == null) return false;
            foreach (string f in OurFiles)
                if (string.Equals(fileName, f, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        public static void Log(string s)
        {
            try { File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + "  " + s + "\r\n", Encoding.UTF8); }
            catch (Exception) { }
        }

        // the log never grows without limit: a big old log is simply started again
        public static void TrimLog()
        {
            try { FileInfo fi = new FileInfo(LogPath); if (fi.Exists && fi.Length > 1000000) fi.Delete(); }
            catch (Exception) { }
        }

        public static string Msg(Exception ex)
        {
            while (ex is TargetInvocationException && ex.InnerException != null) ex = ex.InnerException;
            string s = ex.GetType().Name + ": " + ex.Message;
            COMException ce = ex as COMException;
            if (ce != null) s += " (0x" + unchecked((uint)ce.ErrorCode).ToString("X8") + ")";   // the Windows error number helps to find the cause
            return s;
        }

        // Word's Dictionary.Path is often empty or not a real folder (Microsoft 365 keeps some dictionaries in the cloud),
        // so look for the file where Word usually keeps it: the folder Word reported first, then the UProof folder.
        public static string ResolveDicPath(string folderFromWord, string name)
        {
            if (string.IsNullOrEmpty(name)) return null;
            string p = null;
            if (!string.IsNullOrEmpty(folderFromWord)) p = IfExists(folderFromWord, name);
            if (p == null) p = IfExists(UProofDir(), name);
            if (p == null)
            {
                try { if (Path.IsPathRooted(name) && File.Exists(name)) p = name; }
                catch (Exception) { }
            }
            return p;
        }

        static string IfExists(string folder, string name)
        {
            try
            {
                string p = Path.Combine(folder, name);
                return File.Exists(p) ? p : null;
            }
            catch (Exception) { return null; }   // a name or folder with characters that are not allowed in a path
        }

        // ------------------------------------------------ exact number of words of each edition
        // Written into the program when it is built (make_assets.py), so the window does not have to open and count the
        // word lists every time it starts. The self-test builds every edition for real and compares.
        static Dictionary<string, int> editionCounts;

        public static int EditionCount(bool extended, bool compose)
        {
            string key = (extended ? "extended" : "basic") + (compose ? "_comp" : "");
            try
            {
                if (editionCounts == null)
                {
                    Dictionary<string, int> d = new Dictionary<string, int>();
                    using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(CountsRes))
                    {
                        if (s != null)
                            using (StreamReader r = new StreamReader(s, new UTF8Encoding(false)))
                            {
                                string line;
                                while ((line = r.ReadLine()) != null)
                                {
                                    string[] p = line.Trim().Split(new char[] { ' ' });   // an explicit array: Split(' ') can bind to a method that .NET Framework does not have
                                    int v;
                                    if (p.Length == 2 && int.TryParse(p[1], out v)) d[p[0]] = v;
                                }
                            }
                    }
                    editionCounts = d;
                }
                int n;
                if (editionCounts.TryGetValue(key, out n) && n > 0) return n;
            }
            catch (Exception ex) { Log("edition counts failed: " + Msg(ex)); }
            int real = BuildWords(extended, compose).Count;   // not expected (the counts file is missing): count for real
            if (editionCounts != null) editionCounts[key] = real;
            return real;
        }

        public static List<string> LoadWords(string resource)
        {
            List<string> list = new List<string>();
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
            {
                if (s == null) throw new InvalidOperationException("Missing embedded word list: " + resource);
                using (GZipStream gz = new GZipStream(s, CompressionMode.Decompress))
                using (StreamReader r = new StreamReader(gz, new UTF8Encoding(false)))
                {
                    string line;
                    while ((line = r.ReadLine()) != null)
                    {
                        line = line.Trim();
                        if (line.Length > 0) list.Add(line);
                    }
                }
            }
            return list;
        }

        // Counts lines without building the list (fast start-up)
        public static int CountWords(string resource)
        {
            int n = 0;
            using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
            {
                if (s == null) return 0;
                using (GZipStream gz = new GZipStream(s, CompressionMode.Decompress))
                using (StreamReader r = new StreamReader(gz, new UTF8Encoding(false)))
                {
                    string line;
                    while ((line = r.ReadLine()) != null) if (line.Trim().Length > 0) n++;
                }
            }
            return n;
        }

        // ------------------------------------------------ acronyms: every Hebrew letter combination with the gershayim (") before the last letter
        const string AcroFirst = "אבגדהוזחטיכלמנסעפצקרשת";    // 22 letters in their regular form
        const string AcroLast = "אבגדהוזחטיךלםןסעףץקרשת";      // 22 letters, final forms (ך ם ן ף ץ) where Hebrew writes them

        public static long AcronymCount(int letters)
        {
            long t = 1;
            for (int i = 0; i < letters; i++) t *= 22;
            return t;
        }

        // all acronyms with minLetters..maxLetters letters, e.g. 3 letters -> רש"י-style pattern: two letters, ", last letter
        public static List<string> MakeAcronyms(int minLetters, int maxLetters)
        {
            List<string> list = new List<string>((int)Math.Min(300000, AcronymCount(maxLetters) + 20000));
            for (int n = minLetters; n <= maxLetters; n++)
            {
                char[] buf = new char[n + 1];
                int[] idx = new int[n];
                long total = AcronymCount(n);
                for (long k = 0; k < total; k++)
                {
                    long r = k;
                    for (int p = n - 1; p >= 0; p--) { idx[p] = (int)(r % 22); r /= 22; }
                    for (int p = 0; p < n - 1; p++) buf[p] = AcroFirst[idx[p]];
                    buf[n - 1] = '"';
                    buf[n] = AcroLast[idx[n - 1]];
                    list.Add(new string(buf));
                }
            }
            return list;
        }

        // ------------------------------------------------ "chibur torani": connectors (ד, וד, כד ...) + common Hebrew words
        public static List<string> MakeComposition()
        {
            List<string> bases = LoadWords(BaseRes);
            List<string> list = new List<string>(bases.Count * Connectors.Length);
            foreach (string b in bases)
                foreach (string c in Connectors)
                    list.Add(c + b);
            return list;
        }

        // the words of the single dictionary file: basic, acronyms, extended, composition - in this order, without duplicates
        public static List<string> BuildWords(bool extended, bool compose)
        {
            List<string> all = new List<string>(extended ? 180000 : 70000);
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            AddUnique(all, seen, LoadWords(BasicRes));
            AddUnique(all, seen, MakeAcronyms(2, 3));
            if (extended) AddUnique(all, seen, LoadWords(MoreRes));
            if (compose) AddUnique(all, seen, MakeComposition());
            return all;
        }

        static void AddUnique(List<string> all, HashSet<string> seen, List<string> words)
        {
            foreach (string w in words) if (seen.Add(w)) all.Add(w);
        }

        // number of words in any custom dictionary file (Unicode, UTF-8 or the old ANSI kind)
        public static int CountDicWords(string path)
        {
            byte[] raw = File.ReadAllBytes(path);
            string txt;
            if (raw.Length >= 2 && raw[0] == 0xFF && raw[1] == 0xFE) txt = Encoding.Unicode.GetString(raw, 2, raw.Length - 2);
            else if (raw.Length >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF) txt = Encoding.UTF8.GetString(raw, 3, raw.Length - 3);
            else txt = Encoding.Default.GetString(raw);
            int n = 0;
            foreach (string line in txt.Split(new char[] { '\n' })) if (line.Trim().Length > 0) n++;
            return n;
        }

        // Only for the self-test: true = writing / deleting / copying over this file fails, as if another program held it.
        public static Func<string, bool> FailWriteHook, FailDeleteHook, FailCopyHook;

        // copies a file over another one
        public static void CopyFile(string from, string to)
        {
            if (FailCopyHook != null && FailCopyHook(to)) throw new IOException("simulated: the file is locked");
            File.Copy(from, to, true);
        }

        // deletes a file, also when it is marked read-only
        public static void DeleteFile(string path)
        {
            if (FailDeleteHook != null && FailDeleteHook(path)) throw new IOException("simulated: the file is in use");
            try
            {
                FileAttributes a = File.GetAttributes(path);
                if ((a & FileAttributes.ReadOnly) != 0) File.SetAttributes(path, a & ~FileAttributes.ReadOnly);
            }
            catch (Exception) { }
            File.Delete(path);
        }

        // the words of any custom dictionary file (Unicode with BOM - what Word writes -, UTF-8, or the old ANSI kind)
        public static List<string> ReadDicWords(string path)
        {
            byte[] raw = File.ReadAllBytes(path);
            string txt;
            if (raw.Length >= 2 && raw[0] == 0xFF && raw[1] == 0xFE) txt = Encoding.Unicode.GetString(raw, 2, raw.Length - 2);
            else if (raw.Length >= 3 && raw[0] == 0xEF && raw[1] == 0xBB && raw[2] == 0xBF) txt = Encoding.UTF8.GetString(raw, 3, raw.Length - 3);
            else txt = Encoding.Default.GetString(raw);
            List<string> words = new List<string>();
            foreach (string line in txt.Split(new char[] { '\n' }))
            {
                string w = line.Trim();   // also drops the \r of CRLF
                if (w.Length > 0) words.Add(w);
            }
            return words;
        }

        // Words in our dictionary file that are not ours. Word puts them there itself when the person clicks "Add to Dictionary" while our file
        // is Word's DEFAULT custom dictionary (it can become one when it is the only dictionary Word lists). They are the person's words:
        // an update keeps them in the new file, a removal moves them to CUSTOM.DIC - they are never deleted with our file.
        // "Ours" = every word of every edition. A version that drops words from the word lists must add the dropped ones to Retired,
        // otherwise they would be kept as if the person had added them.
        // alsoOurs = the words that the person added in the program's own editing window (MyWords): they are written by every
        // installation from the person's list, so in the file they are not "added by Word" - one taken back must go, not be kept.
        // 3.1: 11,034 words of 3.0's extended list that all the scanned libraries together do not have at least 3 times as written
        // (8,111 of them not even once): most are typos of online texts or two words run together; and 626 forms of the morning test
        // build of 3.1 that the final lists do not have. 3.1.0.1: 63 forms that are typos of the online libraries (הייינו) or words run
        // together, and the forms of 27 common words that "chibur torani" no longer builds on (tools/stage1/drop_typos.py).
        // The list is data/retired_list.txt.
        static string[] retired;
        public static string[] Retired
        {
            get
            {
                string[] r = retired;
                if (r == null) retired = r = LoadWords(RetiredRes).ToArray();
                return r;
            }
        }

        public static List<string> ForeignWords(string path, HashSet<string> alsoOurs = null)
        {
            List<string> found = new List<string>();
            if (!File.Exists(path)) return found;
            List<string> inFile = ReadDicWords(path);
            HashSet<string> ours = new HashSet<string>(BuildWords(true, true), StringComparer.Ordinal);
            foreach (string w in Retired) ours.Add(w);
            if (alsoOurs != null) foreach (string w in alsoOurs) ours.Add(w);
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string w in inFile) if (!ours.Contains(w) && seen.Add(w)) found.Add(w);
            return found;
        }

        // adds words to a custom dictionary file (made if it is missing), keeping every word already in it; returns how many were new
        public static int AddWordsToDic(string path, List<string> add)
        {
            List<string> all = File.Exists(path) ? ReadDicWords(path) : new List<string>();
            HashSet<string> have = new HashSet<string>(all, StringComparer.Ordinal);
            int n = 0;
            foreach (string w in add) if (have.Add(w)) { all.Add(w); n++; }
            WriteDic(path, all);
            return n;
        }

        // Word custom dictionaries: Unicode (UTF-16 little endian) with BOM, one word per line.
        // writes a dictionary file whole or not at all: first next to it, then in its place (File.Replace keeps the file's permissions).
        // A failure leaves the file as it was.
        public static void ReplaceDic(string path, List<string> words)
        {
            if (FailWriteHook != null && FailWriteHook(path)) throw new IOException("simulated: the file is locked");
            string tmp = path + ".new";
            WriteDic(tmp, words);
            try { File.Replace(tmp, path, null); }
            catch (Exception)
            {
                try { File.Delete(tmp); } catch (Exception) { }
                throw;
            }
        }

        public static void WriteDic(string path, List<string> words)
        {
            if (FailWriteHook != null && FailWriteHook(path)) throw new IOException("simulated: the file is locked");
            StringBuilder sb = new StringBuilder(words.Count * 8);
            foreach (string w in words) { sb.Append(w); sb.Append("\r\n"); }
            byte[] data = Encoding.Unicode.GetBytes(sb.ToString());
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                fs.WriteByte(0xFF); fs.WriteByte(0xFE);
                fs.Write(data, 0, data.Length);
            }
        }

        // ------------------------------------------------ what is installed on this computer
        public static InstalledInfo DetectInstalled()
        {
            InstalledInfo info = new InstalledInfo();
            try
            {
                string d = UProofDir();
                if (File.Exists(Path.Combine(d, MainFile)))
                {
                    bool ext, comp;
                    if (Stamp.TryReadEdition(out ext, out comp)) { info.Edition = ext ? 2 : 1; info.Compose = comp; }
                    else info.Edition = 1;
                }
                else
                {
                    // an older 3.0.0.x installation (several files)
                    bool b = File.Exists(Path.Combine(d, BasicFile));
                    bool m = File.Exists(Path.Combine(d, MoreFile));
                    info.Compose = File.Exists(Path.Combine(d, ComposeFile));
                    info.Edition = (b && m) ? 2 : (b ? 1 : 0);
                }
                if (info.Edition > 0 || info.Compose)
                {
                    string v, on;
                    if (Stamp.TryRead(out v, out on)) { info.Version = v; info.InstalledOn = on; }
                }
            }
            catch (Exception) { }
            return info;
        }

        // 0 = same, -1 = installed is older, 1 = installed is newer, 2 = unknown
        public static int CompareToCurrent(string installed)
        {
            try
            {
                Version a = new Version(installed), b = new Version(AppVersion);
                int c = a.CompareTo(b);
                return c < 0 ? -1 : (c > 0 ? 1 : 0);
            }
            catch (Exception) { return 2; }
        }

        // The Microsoft Word that is installed on this computer, read from the registry (no need to start Word)
        public static string WordVersion()
        {
            try
            {
                using (RegistryKey k = Registry.ClassesRoot.OpenSubKey(@"Word.Application\CurVer"))
                {
                    if (k == null) return null;
                    string v = k.GetValue(null) as string;   // e.g. Word.Application.16
                    if (string.IsNullOrEmpty(v)) return null;
                    string num = v.Substring(v.LastIndexOf('.') + 1);
                    return WordVersionLabel(num, num == "16" ? ReadOfficeReleaseIds() : null);
                }
            }
            catch (Exception) { return null; }
        }

        // Office 2016, 2019, 2021, 2024 and Microsoft 365 all call themselves "16"; the way Office was installed tells them apart
        static string ReadOfficeReleaseIds()
        {
            foreach (RegistryView view in new RegistryView[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                try
                {
                    using (RegistryKey hk = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                    using (RegistryKey k = hk.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration"))
                    {
                        if (k == null) continue;
                        string ids = k.GetValue("ProductReleaseIds") as string;
                        if (!string.IsNullOrEmpty(ids)) return ids;
                    }
                }
                catch (Exception) { }
            }
            return null;
        }

        // num = the number after "Word.Application.", releaseIds = Click-to-Run product ids (e.g. "O365ProPlusRetail"), or null
        public static string WordVersionLabel(string num, string releaseIds)
        {
            switch (num)
            {
                case "16":
                    if (!string.IsNullOrEmpty(releaseIds))
                    {
                        string ids = releaseIds.ToLowerInvariant();
                        if (ids.Contains("o365") || ids.Contains("microsoft365") || ids.Contains("m365")) return "Microsoft 365";
                        if (ids.Contains("2024")) return "2024";
                        if (ids.Contains("2021")) return "2021";
                        if (ids.Contains("2019")) return "2019";
                    }
                    return "2016 / 2019 / 2021 / 365";
                case "15": return "2013";
                case "14": return "2010";
                case "12": return "2007";
                case "11": return "2003";
                default: return "גרסה " + num;
            }
        }
    }

    public class InstalledInfo
    {
        public int Edition;          // 0 = none, 1 = basic, 2 = extended
        public bool Compose;         // "chibur torani" installed
        public string Version;       // null = installed by an older program (no version recorded)
        public string InstalledOn;
    }

    // Remembers which version/edition this program installed, so the next start can tell what is on the computer
    public static class Stamp
    {
        const string KeyPath = @"Software\TorahDictionaryForWord";

        public static void Write(bool extended, bool compose)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    k.SetValue("Version", Core.AppVersion, RegistryValueKind.String);
                    k.SetValue("Edition", extended ? "Extended" : "Basic", RegistryValueKind.String);
                    k.SetValue("Composition", compose ? 1 : 0, RegistryValueKind.DWord);
                    k.SetValue("InstalledOn", DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture), RegistryValueKind.String);
                }
            }
            catch (Exception ex) { Core.Log("Stamp write failed: " + Core.Msg(ex)); }
        }

        public static bool TryReadEdition(out bool extended, out bool compose)
        {
            extended = false; compose = false;
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (k == null) return false;
                    string e = k.GetValue("Edition") as string;
                    if (e == null) return false;
                    extended = e == "Extended";
                    object c = k.GetValue("Composition");
                    compose = c != null && Convert.ToInt32(c) == 1;
                    return true;
                }
            }
            catch (Exception) { return false; }
        }

        public static bool TryRead(out string version, out string installedOn)
        {
            version = null; installedOn = null;
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (k == null) return false;
                    version = k.GetValue("Version") as string;
                    installedOn = k.GetValue("InstalledOn") as string;
                    return version != null;
                }
            }
            catch (Exception) { return false; }
        }

        // forgets the installed dictionary (the separate "repeated words" note stays)
        public static void Clear()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath, true))
                {
                    if (k == null) return;
                    foreach (string n in new string[] { "Version", "Edition", "Composition", "InstalledOn" })
                        k.DeleteValue(n, false);
                }
            }
            catch (Exception) { }
        }

        // remembers that this program switched off Word's "flag repeated words" (so the button can offer to switch it back on)
        public static void WriteRepeatOff(bool off)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(KeyPath))
                    k.SetValue("RepeatedWordsOff", off ? 1 : 0, RegistryValueKind.DWord);
            }
            catch (Exception ex) { Core.Log("Stamp (repeat) failed: " + Core.Msg(ex)); }
        }

        public static bool ReadRepeatOff()
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (k == null) return false;
                    object v = k.GetValue("RepeatedWordsOff");
                    return v != null && Convert.ToInt32(v) == 1;
                }
            }
            catch (Exception) { return false; }
        }
    }

    // ---------------------------------------------------------------- Word automation (late binding, no Office files needed)
    public class WordMissingException : Exception
    {
        public WordMissingException() : base("Microsoft Word is not installed (Word.Application not registered).") { }
    }

    public class WordSession : IDisposable
    {
        public object App;
        public bool Created;   // true when we started a hidden Word ourselves

        public void Dispose()
        {
            object app = App;
            App = null;
            try { if (Created && app != null) WordCom.Call(app, "Quit", 0); } catch (Exception) { }   // 0 = do not save
            try { if (app != null && Marshal.IsComObject(app)) Marshal.FinalReleaseComObject(app); } catch (Exception) { }
            app = null;
            // let the runtime let go of every other Word object it still holds - otherwise a hidden Word can stay alive in the background
            try { GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect(); } catch (Exception) { }
        }
    }

    public class DictInfo
    {
        public string Name;
        public string Folder;     // the folder Word reported (may be empty)
        public string FullPath;
    }

    public static class WordCom
    {
        // ---- Word says "I am busy" while it starts up or while a dialog is open (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER).
        // That is not a failure: wait a moment and ask again, for up to RetryMs.
        public static int RetryMs = 15000;
        const int BusyWaitMs = 250;

        public static bool IsBusy(Exception ex)
        {
            while (ex is TargetInvocationException && ex.InnerException != null) ex = ex.InnerException;
            COMException ce = ex as COMException;
            if (ce == null) return false;
            uint hr = unchecked((uint)ce.ErrorCode);
            // 0x800AC472 = Word's own "not available right now", which it also answers while it is busy
            return hr == 0x80010001u || hr == 0x8001010Au || hr == 0x800AC472u;
        }

        public static T Retry<T>(Func<T> f)
        {
            int waited = 0;
            while (true)
            {
                try { return f(); }
                catch (Exception ex)
                {
                    if (!IsBusy(ex) || waited >= RetryMs) throw;
                    if (waited == 0) Core.Log("Word is busy, waiting a moment…");
                    Thread.Sleep(BusyWaitMs);
                    waited += BusyWaitMs;
                }
            }
        }

        public static object Get(object o, string name, params object[] args)
        { return Retry<object>(delegate { return o.GetType().InvokeMember(name, BindingFlags.GetProperty, null, o, args); }); }
        public static void Put(object o, string name, object value)
        { Retry<object>(delegate { o.GetType().InvokeMember(name, BindingFlags.SetProperty, null, o, new object[] { value }); return null; }); }
        public static object Call(object o, string name, params object[] args)
        { return Retry<object>(delegate { return o.GetType().InvokeMember(name, BindingFlags.InvokeMethod, null, o, args); }); }

        // ---- the hidden Words this program starts: remember their process, so we can wait until one has really gone
        // (a fixed pause was a guess: too long on a fast computer, too short on a slow one)
        static readonly List<int> ours = new List<int>();

        static int mySession = -1;   // the Windows session of this program (-2 = cannot be told)

        // the Word processes of THIS Windows user: on a shared computer another user's Word is in another session, and it neither
        // holds this person's dictionaries nor can this person close it
        static int[] WordPids()
        {
            try
            {
                if (mySession == -1)
                {
                    try { using (Process me = Process.GetCurrentProcess()) mySession = me.SessionId; }
                    catch (Exception) { mySession = -2; }
                }
                Process[] ps = Process.GetProcessesByName("WINWORD");
                List<int> ids = new List<int>();
                foreach (Process p in ps)
                {
                    try
                    {
                        bool mine = true;
                        if (mySession >= 0) { try { mine = p.SessionId == mySession; } catch (Exception) { mine = true; } }
                        if (mine) ids.Add(p.Id);
                    }
                    catch (Exception) { }
                    finally { p.Dispose(); }
                }
                return ids.ToArray();
            }
            catch (Exception) { return null; }   // cannot list processes: callers fall back to a fixed pause
        }

        static bool IsAlive(int pid)
        {
            try { using (Process p = Process.GetProcessById(pid)) return !p.HasExited; }
            catch (Exception) { return false; }   // no such process any more
        }

        // true = we knew the process of a hidden Word and waited for it; false = nothing known to wait for
        public static bool WaitForOurWordsToExit(int maxMs)
        {
            List<int> mine;
            lock (ours) mine = new List<int>(ours);
            if (mine.Count == 0) return false;
            int waited = 0;
            while (true)
            {
                bool alive = false;
                foreach (int pid in mine) if (IsAlive(pid)) alive = true;
                if (!alive || waited >= maxMs) break;
                Thread.Sleep(100);
                waited += 100;
            }
            lock (ours) ours.RemoveAll(delegate (int pid) { return !IsAlive(pid); });
            Core.Log("Waited " + waited + " ms for the hidden Word to close.");
            Thread.Sleep(300);   // a short moment for Word to finish letting go of its files and settings
            return true;
        }

        static WordSession Start(string logText)
        {
            Type t = Type.GetTypeFromProgID("Word.Application");
            if (t == null) throw new WordMissingException();
            WaitForOurWordsToExit(8000);   // an earlier hidden Word may still be closing: starting now could hand us that one

            int[] before = WordPids();
            WordSession s = new WordSession();
            s.App = Activator.CreateInstance(t);
            s.Created = true;
            int[] after = WordPids();
            if (before != null && after != null)
                foreach (int pid in after)
                    if (Array.IndexOf(before, pid) < 0) lock (ours) ours.Add(pid);
            try { Put(s.App, "Visible", false); } catch (Exception) { }
            try { Put(s.App, "DisplayAlerts", 0); } catch (Exception) { }
            Core.Log(logText);
            return s;
        }

        // Only for the self-test: a pretend Word in place of a real one.
        public static Func<WordSession> OpenHook, OpenNewHook;
        // Only for the self-test: whether a Word runs (null = ask Windows), and how long to wait for every Word to close
        public static Func<bool> AnyWordHook;
        public static Func<List<string>> OtherOfficeHook;   // (the self-test) the other Office programs that run now
        public static int CloseWaitMs = 15000;

        // true = a Word runs in this session - also one this program started that is still closing - or that cannot be told.
        // Word's own list of dictionaries in the registry is changed only when this is false: a running Word writes its copy of the list
        // back when it closes, and the change would be lost.
        public static bool AnyWordRunning()
        {
            if (AnyWordHook != null) return AnyWordHook();
            int[] ids = WordPids();
            return ids == null || ids.Length > 0;
        }

        // The other Office programs that run now in this Windows session. They too write Word's list of dictionaries and the personal
        // dictionary back when they close (Outlook even adds names to the personal dictionary by itself).
        public static List<string> OtherOfficeRunning()
        {
            if (OtherOfficeHook != null) return OtherOfficeHook();
            List<string> found = new List<string>();
            string[] exe = { "OUTLOOK", "EXCEL", "POWERPNT", "ONENOTE" }, name = { "Outlook", "Excel", "PowerPoint", "OneNote" };
            WordPids();   // (finds this program's Windows session)
            for (int i = 0; i < exe.Length; i++)
            {
                try
                {
                    bool here = false;
                    foreach (Process p in Process.GetProcessesByName(exe[i]))
                    {
                        try { if (mySession < 0 || p.SessionId == mySession) here = true; }
                        catch (Exception) { here = true; }
                        finally { p.Dispose(); }
                    }
                    if (here) found.Add(name[i]);
                }
                catch (Exception) { }
            }
            return found;
        }

        // waits until no Word runs, at most maxMs; true = none runs now
        public static bool WaitUntilNoWord(int maxMs)
        {
            for (int waited = 0; ; waited += 250)
            {
                if (!AnyWordRunning()) { if (waited > 0) Thread.Sleep(300); return true; }   // (a moment for Word to let go of its settings)
                if (waited >= maxMs) { Core.Log("A Word is still running after " + waited + " ms"); return false; }
                Thread.Sleep(250);
            }
        }

        // How many Word processes are running now, visible or hidden (a Word that this program started itself is not counted).
        // -1 = cannot be told. A running Word cannot show a dictionary that was added to it: it reads its list when it starts.
        public static int RunningWordCount()
        {
            int[] ids = WordPids();
            if (ids == null) return -1;
            int n = 0;
            lock (ours) { foreach (int id in ids) if (!ours.Contains(id)) n++; }
            return n;
        }

        // always a NEW hidden Word, which reads the custom dictionaries the way Word does when the user opens it
        public static WordSession OpenNew()
        {
            if (OpenNewHook != null) return OpenNewHook();
            return Start("Started a new hidden Word for the check.");
        }

        public static WordSession Open()
        {
            if (OpenHook != null) return OpenHook();
            WaitForOurWordsToExit(8000);   // a hidden Word of ours that is still closing must not be taken for the person's Word
            try
            {
                WordSession s = new WordSession();
                s.App = Marshal.GetActiveObject("Word.Application");   // Word is already open
                s.Created = false;
                Core.Log("Attached to a running Word.");
                return s;
            }
            catch (Exception) { }
            return Start("Started a hidden Word.");
        }

        public static object Item(object coll, int i)
        {
            try { return Get(coll, "Item", i); }
            catch (Exception) { return Call(coll, "Item", i); }
        }

        public static int RemoveOurs(WordSession s)
        {
            int removed = 0;
            object dicts = Get(s.App, "CustomDictionaries");
            int n = Convert.ToInt32(Get(dicts, "Count"));
            for (int i = n; i >= 1; i--)
            {
                object d = Item(dicts, i);
                string name = Convert.ToString(Get(d, "Name"));
                if (Core.IsOurs(name))
                {
                    Call(d, "Delete");
                    removed++;
                    Core.Log("Unregistered from Word: " + name);
                }
            }
            return removed;
        }

        public static void Add(WordSession s, string path)
        {
            object dicts = Get(s.App, "CustomDictionaries");
            object d = Call(dicts, "Add", path);
            try { Put(d, "LanguageSpecific", false); } catch (Exception) { }   // apply to every language, Hebrew included
            Core.Log("Registered in Word: " + path);
        }

        // Word puts the words the person adds ("Add to Dictionary") into its DEFAULT custom dictionary. That must never be ours: our file is
        // rewritten by every update and deleted by a removal. If ours is the default, another dictionary of the list becomes the default;
        // if Word lists no other one, Word's usual personal dictionary CUSTOM.DIC is made (empty) and added.
        // Never fails the installation: a problem here only goes to the log.
        public static void KeepOursNotDefault(WordSession s, string folder)
        {
            try
            {
                object dicts = Get(s.App, "CustomDictionaries");
                string active = null;
                bool known = false;
                try
                {
                    object a = Get(dicts, "ActiveCustomDictionary");
                    known = true;
                    if (a != null) active = Convert.ToString(Get(a, "Name"));
                }
                catch (Exception ex) { Core.Log("The default custom dictionary could not be read: " + Core.Msg(ex)); }
                Core.Log("Word's default custom dictionary: " + (active ?? (known ? "(none)" : "(unknown)")));
                if (active != null && !Core.IsOurs(active)) return;

                object other = null;
                int best = int.MaxValue;
                int n = Convert.ToInt32(Get(dicts, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object d = Item(dicts, i);
                    string name = Convert.ToString(Get(d, "Name"));
                    if (Core.IsOurs(name)) continue;
                    int rank = string.Equals(name, "RoamingCustom.dic", StringComparison.OrdinalIgnoreCase) ? 0
                             : string.Equals(name, "CUSTOM.DIC", StringComparison.OrdinalIgnoreCase) ? 1 : 2;
                    if (rank < best) { best = rank; other = d; }
                }
                // the default could not be read, and the person has dictionaries of their own: Word's choice is left alone
                if (!known && other != null) return;
                if (other == null)
                {
                    string custom = Path.Combine(folder, "CUSTOM.DIC");
                    if (!File.Exists(custom)) Core.WriteDic(custom, new List<string>());   // an empty personal dictionary, as Word itself starts one
                    other = Call(dicts, "Add", custom);
                    try { Put(other, "LanguageSpecific", false); } catch (Exception) { }
                    Core.Log("Added Word's usual personal dictionary: " + custom);
                }
                string wanted = Convert.ToString(Get(other, "Name"));
                try { Put(dicts, "ActiveCustomDictionary", other); }
                catch (Exception ex) { Core.Log("Setting the default dictionary failed: " + Core.Msg(ex)); }
                string now = ActiveName(dicts);
                // if the usual way of setting did not take, the two kinds of COM "set" are tried one after the other
                foreach (BindingFlags how in new BindingFlags[] { BindingFlags.PutRefDispProperty, BindingFlags.PutDispProperty })
                {
                    if (now != null && !Core.IsOurs(now)) break;
                    BindingFlags h = how;
                    object target = dicts, value = other;
                    try { Retry<object>(delegate { target.GetType().InvokeMember("ActiveCustomDictionary", h, null, target, new object[] { value }); return null; }); }
                    catch (Exception ex) { Core.Log("Setting the default dictionary (" + h + ") failed: " + Core.Msg(ex)); }
                    now = ActiveName(dicts);
                }
                Core.Log("Wanted as the default: " + wanted + "; Word's default custom dictionary is now: " + (now ?? "(unknown)"));
            }
            catch (Exception ex) { Core.Log("Could not keep our dictionary from being the default: " + Core.Msg(ex)); }
        }

        static string ActiveName(object dicts)
        {
            try { object a = Get(dicts, "ActiveCustomDictionary"); return a == null ? null : Convert.ToString(Get(a, "Name")); }
            catch (Exception) { return null; }
        }

        // is a dictionary with this file name in Word's list?
        public static bool IsListed(WordSession s, string leaf)
        {
            object dicts = Get(s.App, "CustomDictionaries");
            int n = Convert.ToInt32(Get(dicts, "Count"));
            for (int i = 1; i <= n; i++)
                if (string.Equals(Convert.ToString(Get(Item(dicts, i), "Name")), leaf, StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        static bool SameFolder(string a, string b)
        {
            try { return string.Equals(Path.GetFullPath(a).TrimEnd('\\', '/'), Path.GetFullPath(b).TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase); }
            catch (Exception) { return false; }
        }

        // makes sure Word lists this very file - same name AND same folder (removal: the person's words were moved into it)
        public static void EnsureListed(WordSession s, string path)
        {
            try
            {
                object dicts = Get(s.App, "CustomDictionaries");
                string leaf = Path.GetFileName(path), folder = Path.GetDirectoryName(path);
                int n = Convert.ToInt32(Get(dicts, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object d = Item(dicts, i);
                    if (!string.Equals(Convert.ToString(Get(d, "Name")), leaf, StringComparison.OrdinalIgnoreCase)) continue;
                    string f = null;
                    try { f = Convert.ToString(Get(d, "Path")); } catch (Exception) { }
                    if (string.IsNullOrEmpty(f) || SameFolder(f, folder)) return;   // no folder from Word: its usual folder is meant
                }
                object nd = Call(dicts, "Add", path);
                try { Put(nd, "LanguageSpecific", false); } catch (Exception) { }
                Core.Log("Listed in Word: " + path);
            }
            catch (Exception ex) { Core.Log("Could not list " + path + " in Word: " + Core.Msg(ex)); }
        }

        // every custom dictionary Word currently has, whoever added it
        public static List<DictInfo> ListAll(WordSession s)
        {
            List<DictInfo> list = new List<DictInfo>();
            object dicts = Get(s.App, "CustomDictionaries");
            int n = Convert.ToInt32(Get(dicts, "Count"));
            for (int i = 1; i <= n; i++)
            {
                try
                {
                    object d = Item(dicts, i);
                    DictInfo di = new DictInfo();
                    di.Name = Convert.ToString(Get(d, "Name"));
                    string p = null;
                    try { p = Convert.ToString(Get(d, "Path")); } catch (Exception) { }
                    di.Folder = p;
                    di.FullPath = Core.ResolveDicPath(p, di.Name);   // null when the file cannot be found
                    list.Add(di);
                }
                catch (Exception ex) { Core.Log("ListAll item failed: " + Core.Msg(ex)); }
            }
            return list;
        }

        // removes every custom dictionary from Word; names of the ones Word refused go to 'refused'
        public static int RemoveAll(WordSession s, List<string> refused)
        {
            int removed = 0;
            object dicts = Get(s.App, "CustomDictionaries");
            int n = Convert.ToInt32(Get(dicts, "Count"));
            for (int i = n; i >= 1; i--)
            {
                string name = "?";
                try
                {
                    object d = Item(dicts, i);
                    name = Convert.ToString(Get(d, "Name"));
                    Call(d, "Delete");
                    removed++;
                    Core.Log("Advanced removal, unregistered: " + name);
                }
                catch (Exception ex) { refused.Add(name); Core.Log("Advanced removal, Word refused " + name + ": " + Core.Msg(ex)); }
            }
            return removed;
        }

        // Word option "Flag repeated words" (File > Options > Proofing)
        public static bool SetRepeatWords(WordSession s, bool flag)
        {
            object opts = Get(s.App, "Options");
            Put(opts, "RepeatWord", flag);
            bool now = Convert.ToBoolean(Get(opts, "RepeatWord"));
            Core.Log("Options.RepeatWord set to " + flag + ", now " + now);
            return now == flag;
        }
    }

    // ---------------------------------------------------------------- fallback: write the registry list Word reads at start-up
    public static class RegFallback
    {
        public const string KeyPath = @"Software\Microsoft\Shared Tools\Proofing Tools\1.0\Custom Dictionaries";

        // writes what the registry list holds into the log (diagnostics only; shows whether Word itself writes its list there)
        public static void DumpToLog(string when)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath))
                {
                    if (k == null) { Core.Log("Registry list (" + when + "): the key does not exist"); return; }
                    StringBuilder sb = new StringBuilder();
                    foreach (string vn in k.GetValueNames()) sb.Append(vn + "=" + Convert.ToString(k.GetValue(vn)) + "; ");
                    Core.Log("Registry list (" + when + "): " + (sb.Length == 0 ? "empty" : sb.ToString()));
                }
            }
            catch (Exception ex) { Core.Log("Registry list dump failed: " + Core.Msg(ex)); }
        }

        // Every entry is a number ("3" = the file); Office may keep settings next to it under the same number ("3_state", "3_culturetag").
        static int NumberOf(string valueName)
        {
            int cut = valueName.IndexOf('_');
            string head = cut >= 0 ? valueName.Substring(0, cut) : valueName;
            int n;
            return int.TryParse(head, NumberStyles.None, CultureInfo.InvariantCulture, out n) ? n : -1;
        }

        // An entry goes together with its own settings, so that a later entry with the same number does not inherit them - by their exact
        // names only. Never by the prefix "<n>_": that would also hit Word 365's entry "<n>_16_..._LiveId" of another dictionary (3.1 review:
        // with "1_" it deleted the person's own RoamingCustom.dic from Word's list).
        static readonly string[] SettingSuffixes = { "_state", "_roamed", "_external", "_culturetag" };
        static void DeleteEntry(RegistryKey k, string vn)
        {
            k.DeleteValue(vn, false);
            foreach (string other in k.GetValueNames())
                foreach (string s in SettingSuffixes)
                    if (string.Equals(other, vn + s, StringComparison.OrdinalIgnoreCase)) k.DeleteValue(other, false);
        }

        static string Leaf(string path)
        {
            try { return Path.GetFileName(path); }
            catch (Exception) { return null; }   // characters that are not allowed in a path
        }

        public static void Add(string path)
        {
            using (RegistryKey k = Registry.CurrentUser.CreateSubKey(KeyPath))
            {
                int max = 0;
                foreach (string vn in k.GetValueNames())
                {
                    int n = NumberOf(vn);
                    if (n > max) max = n;   // also numbers that only left-over settings use
                    if (vn.IndexOf('_') >= 0) continue;
                    string v = k.GetValue(vn) as string;
                    if (v != null && string.Equals(v, path, StringComparison.OrdinalIgnoreCase)) return;   // already there
                }
                k.SetValue((max + 1).ToString(CultureInfo.InvariantCulture), path, RegistryValueKind.String);
                Core.Log("Registry fallback: added " + path);
            }
        }

        // every entry of ours
        public static void RemoveOurs() { Remove(true); }

        // only the entries of the old several-file versions; an entry of the one file of today stays
        public static void RemoveLegacy() { Remove(false); }

        static void Remove(bool includingMain)
        {
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath, true))
            {
                if (k == null) return;
                foreach (string vn in k.GetValueNames())
                {
                    if (vn.IndexOf('_') >= 0) continue;   // a setting: it goes with its entry
                    string v = k.GetValue(vn) as string;
                    string leaf = v == null ? null : Leaf(v);
                    if (leaf == null || !Core.IsOurs(leaf)) continue;
                    if (!includingMain && string.Equals(leaf, Core.MainFile, StringComparison.OrdinalIgnoreCase)) continue;
                    DeleteEntry(k, vn);
                    Core.Log("Registry fallback: removed " + v);
                }
            }
        }

        // advanced removal: drops every custom dictionary entry from the registry list; returns what it removed (for the backup note)
        public static List<string> RemoveAllDictionaries()
        {
            List<string> removed = new List<string>();
            using (RegistryKey k = Registry.CurrentUser.OpenSubKey(KeyPath, true))
            {
                if (k == null) return removed;
                foreach (string vn in k.GetValueNames())
                {
                    if (vn.IndexOf('_') >= 0) continue;
                    string v = k.GetValue(vn) as string;
                    if (v != null && v.EndsWith(".dic", StringComparison.OrdinalIgnoreCase))
                    {
                        removed.Add(v);
                        DeleteEntry(k, vn);
                        Core.Log("Registry (advanced): removed " + v);
                    }
                }
            }
            return removed;
        }
    }

    // ---------------------------------------------------------------- Word 365's own list of custom dictionaries (in the same registry key)
    // One numbered entry per dictionary for every signed-in account: "<n>_<account>" = the file (a name in Word's folder, or a full path),
    // with "<n>_<account>_state", "_roamed", "_external" next to it (e.g. "4_16_7bd116fa7f21d40b_LiveId" = "Milon_Torani.dic").
    // A dictionary taken out of Word KEEPS its entry, and a file of another name takes a new number. With 10 entries Word refuses to add
    // any dictionary (CustomDictionaries.Add -> E_FAIL): this happened on 9-11-2026, after a size check had added five test files to a list
    // that already held two files of version 3.0.0.1 and a WordCheck trial. Entries of files that are never needed again are therefore
    // deleted - only while no Word runs (a running Word writes its own copy of the list back) - and the numbers that are left are closed up
    // (1, 2, 3 ... as Word writes them; Word renumbers its list by itself too).
    public class DictEntry
    {
        public int Number;
        public string Identity = "";                      // "" = an entry without an account (older Office); else e.g. "16_..._LiveId"
        public string Main;                               // the value that holds the file (null: only settings are left)
        public string File;
        public List<string> Names = new List<string>();   // every value of the entry: the file and its settings
        public string Leaf
        {
            get
            {
                if (File == null) return null;
                try { return Path.GetFileName(File); } catch (Exception) { return File; }
            }
        }
    }

    public static class DictList
    {
        public const string SizeTestPrefix = "Milon_SizeTest_";
        static readonly string[] SettingNames = { "state", "roamed", "external", "culturetag" };

        // "4_16_x_LiveId_state" -> 4, "16_x_LiveId", a setting;  "11" -> 11, "", the file;  "1_state" -> 1, "", a setting;  "UpdateComplete" -> false
        public static bool Parse(string vn, out int number, out string identity, out bool isMain)
        {
            number = -1; identity = ""; isMain = false;
            int cut = vn.IndexOf('_');
            string head = cut < 0 ? vn : vn.Substring(0, cut);
            if (head.Length == 0 || !int.TryParse(head, NumberStyles.None, CultureInfo.InvariantCulture, out number)) return false;
            string rest = cut < 0 ? "" : vn.Substring(cut + 1);
            if (rest.Length == 0) { isMain = true; return true; }
            foreach (string st in SettingNames)
            {
                if (string.Equals(rest, st, StringComparison.OrdinalIgnoreCase)) return true;
                if (rest.EndsWith("_" + st, StringComparison.OrdinalIgnoreCase)) { identity = rest.Substring(0, rest.Length - st.Length - 1); return true; }
            }
            identity = rest;
            isMain = true;
            return true;
        }

        public static List<DictEntry> Read(RegistryKey k)
        {
            Dictionary<string, DictEntry> by = new Dictionary<string, DictEntry>(StringComparer.OrdinalIgnoreCase);
            List<DictEntry> list = new List<DictEntry>();
            foreach (string vn in k.GetValueNames())
            {
                int n; string id; bool main;
                if (!Parse(vn, out n, out id, out main)) continue;
                string key = n.ToString(CultureInfo.InvariantCulture) + "|" + id;
                DictEntry e;
                if (!by.TryGetValue(key, out e)) { e = new DictEntry(); e.Number = n; e.Identity = id; by[key] = e; list.Add(e); }
                e.Names.Add(vn);
                if (main) { e.Main = vn; e.File = Convert.ToString(k.GetValue(vn)); }
            }
            list.Sort(delegate (DictEntry a, DictEntry b) { int c = string.CompareOrdinal(a.Identity, b.Identity); return c != 0 ? c : a.Number.CompareTo(b.Number); });
            return list;
        }

        // files of older versions, of a size check and of WordCheck trials; with includeMain also our one file (Word adds it again at once)
        public static bool IsStale(DictEntry e, bool includeMain)
        {
            string leaf = e.Leaf;
            if (leaf == null) return false;
            foreach (string f in Core.LegacyFiles) if (string.Equals(leaf, f, StringComparison.OrdinalIgnoreCase)) return true;
            if (includeMain && string.Equals(leaf, Core.MainFile, StringComparison.OrdinalIgnoreCase)) return true;
            return leaf.StartsWith(SizeTestPrefix, StringComparison.OrdinalIgnoreCase)
                || (leaf.StartsWith("wordcheck", StringComparison.OrdinalIgnoreCase) && leaf.EndsWith(".dic", StringComparison.OrdinalIgnoreCase));
        }

        // a value is moved to a free name only: an existing value is never overwritten
        static bool Rename(RegistryKey k, string from, string to)
        {
            if (k.GetValue(to, null, RegistryValueOptions.DoNotExpandEnvironmentNames) != null) return false;
            object v = k.GetValue(from, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
            RegistryValueKind kind = k.GetValueKind(from);
            k.SetValue(to, v, kind);
            k.DeleteValue(from, false);
            return true;
        }

        // deletes the stale entries and closes up the numbers; returns what was done, one line each
        public static List<string> Clean(RegistryKey k, bool includeMain)
        {
            List<string> done = new List<string>();
            foreach (DictEntry e in Read(k))
            {
                if (!IsStale(e, includeMain)) continue;
                foreach (string vn in e.Names) k.DeleteValue(vn, false);
                done.Add("deleted " + e.Number + (e.Identity.Length > 0 ? " (" + e.Identity + ")" : "") + " " + e.File);
            }
            if (done.Count == 0) return done;
            // The numbers of each account's list are closed up, as Word writes them. An entry only moves DOWN, and only to a number that no
            // value of that account uses (settings that Office left without a file keep their number). Entries without an account (older
            // Office; Word 365 ignores them) are never moved: a moved one could take the number of an account's entry ("11" -> "1" next to
            // "1_16_..._LiveId"), and older versions removed an entry "<n>" together with every value named "<n>_...".
            Dictionary<string, List<DictEntry>> byId = new Dictionary<string, List<DictEntry>>(StringComparer.OrdinalIgnoreCase);
            foreach (DictEntry e in Read(k))
            {
                if (e.Identity.Length == 0) continue;
                List<DictEntry> l;
                if (!byId.TryGetValue(e.Identity, out l)) { l = new List<DictEntry>(); byId[e.Identity] = l; }
                l.Add(e);
            }
            foreach (KeyValuePair<string, List<DictEntry>> kv in byId)
            {
                HashSet<int> used = new HashSet<int>();
                int lowest = 1;
                foreach (DictEntry e in kv.Value) { used.Add(e.Number); if (e.Number < lowest) lowest = e.Number; }   // (a list that starts at 0 keeps 0)
                List<DictEntry> files = kv.Value.FindAll(delegate (DictEntry e) { return e.Main != null; });
                files.Sort(delegate (DictEntry a, DictEntry b) { return a.Number.CompareTo(b.Number); });
                foreach (DictEntry e in files)
                {
                    int want = lowest;
                    while (want < e.Number && used.Contains(want)) want++;
                    if (want >= e.Number) continue;
                    string oldHead = e.Number.ToString(CultureInfo.InvariantCulture), newHead = want.ToString(CultureInfo.InvariantCulture);
                    // every value of the entry ("<n>_<account>", "<n>_<account>_state" ...) moves, or none does
                    bool free = true;
                    foreach (string vn in e.Names)
                        if (k.GetValue(newHead + vn.Substring(oldHead.Length), null, RegistryValueOptions.DoNotExpandEnvironmentNames) != null) free = false;
                    if (!free) { done.Add("not moved (a name is taken) " + e.Number + " -> " + want + " (" + e.Identity + ") " + e.File); continue; }
                    foreach (string vn in e.Names) Rename(k, vn, newHead + vn.Substring(oldHead.Length));
                    done.Add("moved " + e.Number + " -> " + want + " (" + e.Identity + ") " + e.File);
                    used.Remove(e.Number); used.Add(want);
                    e.Number = want;
                }
            }
            return done;
        }

        // the same on the real list; never fails (a problem only goes to the log). Returns how many entries were deleted.
        public static int CleanStale(bool includeMain, string when)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath, true))
                {
                    if (k == null) return 0;
                    List<string> done = Clean(k, includeMain);
                    int deleted = 0;
                    foreach (string line in done) { Core.Log("Word's list (" + when + "): " + line); if (line.StartsWith("deleted")) deleted++; }
                    return deleted;
                }
            }
            catch (Exception ex) { Core.Log("Word's list could not be cleaned (" + when + "): " + Core.Msg(ex)); return 0; }
        }

        // the files of the person's Office personal dictionary (RoamingCustom.dic) that Word's list holds
        public static List<string> RoamingFiles()
        {
            List<string> found = new List<string>();
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath))
                {
                    if (k == null) return found;
                    foreach (DictEntry e in Read(k))
                    {
                        if (e.File == null || !string.Equals(e.Leaf, "RoamingCustom.dic", StringComparison.OrdinalIgnoreCase)) continue;
                        try { if (Path.IsPathRooted(e.File) && System.IO.File.Exists(e.File) && !found.Contains(e.File)) found.Add(e.File); }
                        catch (Exception) { }
                    }
                }
            }
            catch (Exception ex) { Core.Log("Word's list could not be read: " + Core.Msg(ex)); }
            return found;
        }
    }

    // 3.1: what the advanced removal does - the person chooses in its window
    public class AdvancedOptions
    {
        public bool RemoveOurs = true;        // the Torah dictionary leaves Word, as with "remove from Word"
        public bool CleanPersonal = true;     // its words leave Word's own personal dictionaries too (RoamingCustom.dic, CUSTOM.DIC)
        public bool RepairList = true;        // Word's own list of dictionaries: the entries of old versions and of tests go (a copy first)
        public bool RestoreRepeated = true;   // Word marks repeated words again
        public bool RemoveAll = false;        // every custom dictionary leaves Word - also the person's own and those of other programs
        public bool Any { get { return RemoveOurs || CleanPersonal || RepairList || RestoreRepeated || RemoveAll; } }
    }

    // ---------------------------------------------------------------- the whole install / remove procedure (UI-independent, so it can be tested)
    public class Result
    {
        public bool Ok;
        public string Title;
        public string Text;
        public bool WordWasOpen;
        public List<string> KeptWords;   // (removal) the person's own words that were in our file and were moved to CUSTOM.DIC
        public bool RepeatBack;          // (removal) Word marks repeated words again
        public string Details;           // (removal of every dictionary) its failures, as lines "• ...", for the answer of the choices
    }

    public delegate void ProgressFn(int percent, string text);

    public static class Installer
    {
        public static Result Install(bool extended, bool compose, ProgressFn progress, bool useWord)
        {
            Result res = new Result();
            string dir = Core.UProofDir();
            Core.Log("==== install " + Core.AppVersion + ", extended=" + extended + ", composition=" + compose);

            Stopwatch sw = Stopwatch.StartNew();
            progress(4, "מכין את רשימת המילים…");
            Directory.CreateDirectory(dir);
            string mainPath = Path.Combine(dir, Core.MainFile);
            string tmpPath = mainPath + ".tmp";

            // the word list is prepared in the background while Word starts (starting Word is the slow part): the edition with the
            // person's own changes (MyWords), and the words the person may have added into our file through Word: they are kept (Core.ForeignWords)
            List<string> words = null, theirs = null;
            List<WordChange> mine = null;
            Exception buildError = null;
            Thread builder = new Thread(delegate ()
            {
                try
                {
                    words = Core.BuildWords(extended, compose);
                    HashSet<string> alsoOurs = MyChanges(ref words, out mine);
                    try { theirs = Core.ForeignWords(mainPath, alsoOurs); }
                    catch (Exception ex) { theirs = new List<string>(); Core.Log("Reading the old dictionary file failed: " + Core.Msg(ex)); }
                }
                catch (Exception ex) { buildError = ex; }
            });
            builder.IsBackground = true;
            builder.Start();

            WordSession ws = null;
            string comProblem = null;
            bool wordMissing = false;
            if (useWord)
            {
                // entries of files that are never needed again fill Word's list (it takes no dictionary at all when it holds 10): they go,
                // but only while no Word runs - a running Word would write its own copy of the list back
                if (!WordCom.AnyWordRunning()) DictList.CleanStale(false, "before the installation");
                else Core.Log("A Word is running: Word's list is not cleaned before the installation");
                progress(15, "מתחבר לוורד (בפעם הראשונה זה יכול לקחת כמה שניות)…");
                try { ws = WordCom.Open(); res.WordWasOpen = !ws.Created; }
                catch (WordMissingException) { wordMissing = true; Core.Log("Word missing"); }
                catch (Exception ex) { comProblem = Core.Msg(ex); Core.Log("Word open failed: " + comProblem); }
            }
            bool startedHidden = ws != null && ws.Created;   // a hidden Word of our own: the result can be checked afterwards in a new one
            builder.Join();
            if (buildError != null)
            {
                if (ws != null) ws.Dispose();
                throw new InvalidOperationException("לא הצלחתי להכין את רשימת המילים: " + Core.Msg(buildError), buildError);
            }
            List<string> toWrite = words;
            if (theirs != null && theirs.Count > 0)
            {
                toWrite = new List<string>(words.Count + theirs.Count);
                toWrite.AddRange(words);
                toWrite.AddRange(theirs);
                Core.Log("Kept " + theirs.Count + " word(s) that the person had added into our file through Word");
            }
            Core.Log("timing: word list ready and Word connected after " + sw.ElapsedMilliseconds + " ms (" + words.Count + " words)");

            try
            {
                // 1. the new file is written NEXT TO the old one first: if that fails (no room, no permission), nothing has been changed yet
                progress(25, "כותב את קובץ המילון…");
                try { Core.WriteDic(tmpPath, toWrite); }
                catch (Exception ex)
                {
                    TryDelete(tmpPath);
                    if (!IsFileProblem(ex)) throw;
                    Core.Log("Write failed (new file): " + Core.Msg(ex));
                    return WriteFailed(res, ex);
                }

                // 2. Word lets go of the old dictionary
                if (ws != null)
                {
                    RegFallback.DumpToLog("before the old dictionary is removed");
                    progress(35, "מסיר גרסה קודמת של המילון (אם יש)…");
                    try { WordCom.RemoveOurs(ws); }
                    catch (Exception ex) { Core.Log("RemoveOurs failed: " + Core.Msg(ex)); }
                }

                // 3. the dictionary file itself gets the new words, overwritten in place - the way that has always worked, also while Word is open.
                //    If that fails, a second try copies the complete new file over it (this also mends an old file that was emptied before the
                //    first try failed, e.g. on a full disk). If both fail, the old dictionary goes back into Word, so Word keeps what it had.
                progress(50, "מעדכן את קובץ המילון…");
                Exception writeError = null;
                try { Core.WriteDic(mainPath, toWrite); }
                catch (Exception ex) { writeError = ex; Core.Log("Write failed: " + Core.Msg(ex) + " - a second try: the new file is copied over it"); }
                if (writeError != null && !Directory.Exists(mainPath))   // (a folder of that name: a copy would land inside it)
                {
                    try
                    {
                        Core.CopyFile(tmpPath, mainPath);
                        if (new FileInfo(mainPath).Length != new FileInfo(tmpPath).Length) throw new IOException("the copy is not complete");
                        writeError = null;
                        Core.Log("The second try worked");
                    }
                    catch (Exception ex) { Core.Log("The second try failed too: " + Core.Msg(ex)); }
                }
                if (writeError != null)
                {
                    if (ws != null && File.Exists(mainPath))
                    {
                        try
                        {
                            if (!WordCom.IsListed(ws, Core.MainFile)) { WordCom.Add(ws, mainPath); Core.Log("The old dictionary was given back to Word"); }
                        }
                        catch (Exception ex2) { Core.Log("Could not give the old dictionary back to Word: " + Core.Msg(ex2)); }
                    }
                    // the complete new file stays next to the old one (the old one may have been cut off); the next installation replaces both
                    if (!File.Exists(mainPath)) Stamp.Clear();   // nothing of ours is installed: the window must not say otherwise
                    if (!IsFileProblem(writeError)) throw new InvalidOperationException(Core.Msg(writeError), writeError);
                    return WriteFailed(res, writeError);
                }
                TryDelete(tmpPath);
                if (mine != null) MyWords.SaveApplied(MyWords.Additions(mine));   // what of the person's is in the file now
                foreach (string f in Core.LegacyFiles)   // files of older versions (several files)
                {
                    string p = Path.Combine(dir, f);
                    if (File.Exists(p)) { try { Core.DeleteFile(p); Core.Log("Deleted old file " + f); } catch (Exception ex) { Core.Log("Delete failed " + f + ": " + Core.Msg(ex)); } }
                }

                // 4. Word gets the new dictionary (a second try if it refuses once)
                bool registered = false;
                if (ws != null)
                {
                    progress(70, "מוסיף את המילון לוורד…");
                    try { WordCom.Add(ws, mainPath); registered = true; }
                    catch (Exception ex)
                    {
                        Core.Log("Word add failed: " + Core.Msg(ex) + " - trying once more");
                        try
                        {
                            Thread.Sleep(1000);
                            try { WordCom.RemoveOurs(ws); } catch (Exception) { }   // never two entries of ours
                            WordCom.Add(ws, mainPath);
                            registered = true;
                        }
                        catch (Exception ex2) { comProblem = Core.Msg(ex2); Core.Log("Word add failed again: " + comProblem); }
                    }
                    // Word refuses every dictionary when its list is full (10 entries, see DictList): our hidden Word is closed, the list is
                    // cleaned - our own entry too - and a new hidden Word is asked. Not with the person's Word: it would write the old list back.
                    if (!registered && startedHidden && WordCom.RunningWordCount() == 0)
                    {
                        try
                        {
                            ws.Dispose(); ws = null;
                            WordCom.WaitForOurWordsToExit(8000);
                            // the list is changed only when no Word at all runs any more - ours included
                            if (!WordCom.WaitUntilNoWord(WordCom.CloseWaitMs))
                                Core.Log("Word's list is not cleaned: a Word is still running, and it would write its old list back");
                            else
                            {
                                DictList.CleanStale(true, "Word refused the dictionary");
                                ws = WordCom.OpenNew();
                                try { WordCom.RemoveOurs(ws); } catch (Exception) { }
                                WordCom.Add(ws, mainPath);
                                registered = true;
                                comProblem = null;
                                Core.Log("Word took the dictionary after its list was cleaned");
                            }
                        }
                        catch (Exception ex3) { comProblem = Core.Msg(ex3); Core.Log("Word add failed after the list was cleaned: " + comProblem); }
                    }
                    Core.Log("timing: dictionary written and added to Word after " + sw.ElapsedMilliseconds + " ms");
                    if (registered)
                    {
                        // the person's Office dictionary first: then it, and not a new empty CUSTOM.DIC, becomes Word's default dictionary
                        ReEnableRoaming(ws);
                        WordCom.KeepOursNotDefault(ws, dir);
                        RegFallback.DumpToLog("right after Word added the dictionary");
                    }
                }

                // 5. Word did not take it: the list in the registry that Word reads when it starts (whether it does is checked below)
                bool viaRegistry = false;
                if (!registered)
                {
                    // Word may write its own list into the registry when it closes: our hidden Word is closed first, so it cannot wipe our entry
                    if (ws != null && ws.Created)
                    {
                        ws.Dispose(); ws = null;
                        if (!WordCom.WaitForOurWordsToExit(8000)) Thread.Sleep(1500);
                    }
                    progress(75, "רושם את המילון בהגדרות של וורד…");
                    try { RegFallback.RemoveOurs(); RegFallback.Add(mainPath); viaRegistry = true; }
                    catch (Exception ex) { Core.Log("Registry fallback failed: " + Core.Msg(ex)); }
                }

                if (registered || viaRegistry) Stamp.Write(extended, compose);

                // 6. ask Word itself whether it now knows the words - but only in a Word that this program started itself.
                // A Word that was already running does not know a dictionary that is added to it before it is started again,
                // so asking it would give a wrong "it does not recognize the words" (this is what happened on 2-11-2026 at 13:53).
                // After the registry route the same check tells whether Word reads that list.
                string check = null;
                int verdict = -1;   // 1 = all sample words known, 0 = some are not, -1 = the check did not run
                if ((registered || viaRegistry) && res.WordWasOpen)
                {
                    Core.Log("Word was already running: the automatic check is skipped (a running Word reads its dictionaries only when it starts)");
                }
                else if ((registered || viaRegistry) && startedHidden)
                {
                    progress(82, "בודק שוורד מזהה את המילים (כמה שניות)…");
                    WordSession vs = null;
                    try
                    {
                        if (ws != null)
                        {
                            ws.Dispose(); ws = null;
                            if (!WordCom.WaitForOurWordsToExit(8000)) Thread.Sleep(1500);   // let the first Word close completely
                        }
                        RegFallback.DumpToLog("after the first Word closed");
                        vs = WordCom.OpenNew();
                        check = Verify.Run(vs, words, true, out verdict);
                    }
                    catch (Exception ex) { check = "הבדיקה האוטומטית לא הצליחה (" + Core.Msg(ex) + ")."; Core.Log("Verify failed: " + Core.Msg(ex)); }
                    finally { if (vs != null) vs.Dispose(); }
                    Core.Log("timing: check finished after " + sw.ElapsedMilliseconds + " ms");
                }

                progress(100, "הסתיים");
                Core.Log("timing: install finished after " + sw.ElapsedMilliseconds + " ms");
                string name = (extended ? "המורחב" : "הבסיסי") + (compose ? " + חיבור תורני" : "");
                string countText = words.Count.ToString("N0");

                if (registered || (viaRegistry && verdict == 1))
                {
                    Describe(res, name, countText, check, verdict);
                }
                else if (viaRegistry && verdict == 0)
                {
                    res.Ok = false;
                    res.Title = "צריך להשלים ידנית";
                    res.Text = "המילון " + name + " (" + countText + " מילים) נשמר, אבל וורד לא קיבל אותו, ובבדיקה הוא לא זיהה את המילים.\r\n\r\n"
                             + "לחצו על \"התקנה ידנית\" ופעלו לפי ההוראות (לוקח חצי דקה)."
                             + (check != null ? "\r\n\r\n" + check : "")
                             + (comProblem != null ? "\r\n\r\nפרטי השגיאה: " + comProblem : "");
                }
                else if (viaRegistry)
                {
                    // registered only in the registry list, and nothing could check that Word reads it: never reported as done
                    res.Ok = false;
                    res.Title = "המילון נשמר – צריך לבדוק";
                    res.Text = "המילון " + name + " (" + countText + " מילים) נשמר ונרשם בהגדרות של וורד, אבל לא הצלחתי לוודא שוורד משתמש בו.\r\n\r\n"
                             + "סגרו את וורד אם הוא פתוח, פתחו אותו מחדש והקלידו מילה תורנית, למשל מיבעי.\r\n"
                             + "אם היא מסומנת בקו אדום, לחצו על \"התקנה ידנית\" ופעלו לפי ההוראות."
                             + (wordMissing ? "\r\n\r\n(לא מצאתי וורד מותקן במחשב הזה.)" : "")
                             + (comProblem != null ? "\r\n\r\nפרטי השגיאה: " + comProblem : "");
                }
                else
                {
                    res.Ok = false;
                    res.Title = "צריך להשלים ידנית";
                    res.Text = "קובץ המילון נשמר, אבל לא הצלחתי לרשום אותו בוורד אוטומטית.\r\n\r\n"
                             + "לחצו על \"התקנה ידנית\" ופעלו לפי ההוראות (לוקח חצי דקה).\r\n\r\n"
                             + "פרטי השגיאה: " + (comProblem ?? (wordMissing ? "וורד לא נמצא במחשב" : "לא ידוע"));
                }
                return res;
            }
            finally
            {
                if (ws != null) ws.Dispose();
            }
        }

        static bool IsFileProblem(Exception ex) { return ex is IOException || ex is UnauthorizedAccessException; }

        static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch (Exception ex) { Core.Log("Could not delete " + path + ": " + Core.Msg(ex)); }
        }

        static Result WriteFailed(Result res, Exception ex)
        {
            res.Ok = false;
            res.Title = "לא הצלחתי לכתוב את קובץ המילון";
            res.Text = "הקובץ " + Core.MainFile + " בתיקייה של וורד נעול, או שאין הרשאה לכתוב אליו.\r\n\r\n"
                     + "סגרו את וורד ונסו שוב. אם זה חוזר, הפעילו מחדש את המחשב ונסו שוב.\r\n\r\n"
                     + "פרטי השגיאה: " + Core.Msg(ex);
            return res;
        }

        // Office's personal dictionary of the person (RoamingCustom.dic: the words they added with "Add to Dictionary") is in Word's list but
        // not in use (on 9-11-2026 it was found so after Word's list had been full): it is put back into use. Never fails the installation.
        static void ReEnableRoaming(WordSession ws)
        {
            try
            {
                foreach (string path in DictList.RoamingFiles())
                {
                    if (WordCom.IsListed(ws, Path.GetFileName(path))) continue;
                    WordCom.Add(ws, path);
                    Core.Log("The person's Office dictionary was put back into use: " + path);
                }
            }
            catch (Exception ex) { Core.Log("The person's Office dictionary could not be put back into use: " + Core.Msg(ex)); }
        }

        // the person's changes applied to the words of the edition; returns the person's own words that may be in our file from before
        // (null = the person's list could not be read: then the edition is installed as it is, and nothing in the old file is lost)
        static HashSet<string> MyChanges(ref List<string> words, out List<WordChange> mine)
        {
            mine = null;
            try
            {
                List<WordChange> list = MyWords.Load();
                HashSet<string> alsoOurs = MyWords.LoadApplied();
                foreach (string w in MyWords.Additions(list)) alsoOurs.Add(w);
                if (list.Count > 0)
                {
                    words = MyWords.Apply(words, list);
                    Core.Log("The person's own changes: " + list.Count + " (now " + words.Count + " words)");
                }
                mine = list;
                return alsoOurs;
            }
            catch (Exception ex)
            {
                Core.Log("Reading the person's own changes failed: " + Core.Msg(ex) + " - the edition is installed as it is");
                return null;
            }
        }

        // The text after a dictionary was registered in Word. res.WordWasOpen = Word was already running: the check was not made,
        // and the dictionary will be known to that Word only after it is closed and opened again - never a "does not recognize".
        internal static void Describe(Result res, string name, string countText, string check, int verdict)
        {
            string head = "המילון " + name + " (" + countText + " מילים, בקובץ אחד) נוסף לוורד, גרסה " + Core.AppVersion + ".\r\n\r\n";
            if (res.WordWasOpen)
            {
                res.Ok = true;
                res.Title = "המילון הותקן – נשאר לסגור את וורד";
                res.Text = head
                         + "וורד היה פתוח, ולכן הוא עוד לא מכיר את המילון החדש.\r\n"
                         + "שמרו וסגרו את וורד (את כל החלונות שלו), פתחו אותו מחדש והקלידו מילה תורנית, למשל מיבעי.\r\n\r\n"
                         + "אם אין חלון וורד פתוח, ייתכן שוורד רץ ברקע (למשל אחרי תקלה). הפעלה מחדש של המחשב סוגרת אותו.";
                return;
            }
            res.Ok = verdict != 0;
            res.Title = verdict == 0 ? "המילון הותקן, אבל וורד לא מזהה את כל המילים" : "ההתקנה הושלמה";
            res.Text = head
                     + (check != null ? check + "\r\n\r\n" : "")
                     + "עכשיו פתחו את וורד והקלידו מילה תורנית כדי לבדוק.";
        }

        // removes only what this program installed
        public static Result Remove(bool useWord)
        {
            return Remove(useWord, true);
        }

        // restoreRepeated (3.1, the user): when this program switched off Word's marking of repeated words, the removal switches it back on
        public static Result Remove(bool useWord, bool restoreRepeated) { return Remove(useWord, restoreRepeated, false); }

        // forceRepeated: switched on also when this program did not switch it off (the person chose it in the advanced removal)
        public static Result Remove(bool useWord, bool restoreRepeated, bool forceRepeated)
        {
            Result res = new Result();
            Core.Log("==== remove " + Core.AppVersion);
            bool repeatBack = false;
            Stopwatch sw = Stopwatch.StartNew();
            string dir = Core.UProofDir();
            string mainPath = Path.Combine(dir, Core.MainFile);

            // words that the person added into our file through Word are not ours to delete: they move to Word's usual personal dictionary
            List<string> theirs = new List<string>();
            string customPath = null;
            bool keepMain = false;
            HashSet<string> alsoOurs = MyWords.LoadApplied();   // the person's own words stay in the person's list (MyWords) for the next installation
            try { foreach (string w in MyWords.Additions(MyWords.Load())) alsoOurs.Add(w); }
            catch (Exception ex) { Core.Log("Reading the person's own changes failed: " + Core.Msg(ex)); }
            try { theirs = Core.ForeignWords(mainPath, alsoOurs); }
            catch (Exception ex) { Core.Log("Reading our dictionary file failed: " + Core.Msg(ex)); }
            if (theirs.Count > 0)
            {
                string cp = Path.Combine(dir, "CUSTOM.DIC");
                try
                {
                    if (File.Exists(cp)) Core.CopyFile(cp, cp + ".bak");   // the person's own dictionary as it was, before it is changed
                    int added = Core.AddWordsToDic(cp, theirs);
                    customPath = cp;
                    Core.Log("Moved " + theirs.Count + " word(s) of the person to " + cp + " (" + added + " of them new there)");
                }
                catch (Exception ex) { keepMain = true; Core.Log("Could not save the person's words: " + Core.Msg(ex) + " - nothing is removed"); }
            }
            if (keepMain)
            {
                // their words could not be saved anywhere else: nothing is removed, not even from Word's list, so they go on working
                res.Ok = false;
                res.Title = "ההסרה לא הושלמה";
                res.Text = "לא הסרתי את המילון: בקובץ " + Core.MainFile + " יש מילים שהוספתם בעצמכם לוורד, ולא הצלחתי לשמור אותן במקום אחר.\r\n\r\n"
                         + "סגרו את וורד ולחצו שוב על \"הסר מוורד\".";
                return res;
            }

            int removedInWord = 0;
            string problem = null;
            WordSession ws = null;
            RegFallback.DumpToLog("before the removal");
            if (useWord)
            {
                try
                {
                    ws = WordCom.Open();
                    removedInWord = WordCom.RemoveOurs(ws);
                    if (customPath != null) { WordCom.EnsureListed(ws, customPath); WordCom.KeepOursNotDefault(ws, dir); }
                    if (restoreRepeated && (forceRepeated || Stamp.ReadRepeatOff()))
                    {
                        try { if (WordCom.SetRepeatWords(ws, true)) { Stamp.WriteRepeatOff(false); repeatBack = true; } }
                        catch (Exception ex) { Core.Log("The marking of repeated words could not be switched back on: " + Core.Msg(ex)); }
                    }
                }
                catch (WordMissingException) { }
                catch (Exception ex) { problem = Core.Msg(ex); Core.Log("Word remove failed: " + problem); }
                finally { if (ws != null) ws.Dispose(); }
                Core.Log("timing: Word step of the removal took " + sw.ElapsedMilliseconds + " ms");
            }
            try { RegFallback.RemoveOurs(); } catch (Exception ex) { Core.Log("Registry remove failed: " + Core.Msg(ex)); }
            int deleted = 0;
            List<string> stuck = new List<string>();
            foreach (string f in Core.OurFiles)
            {
                string p = Path.Combine(dir, f);
                if (!File.Exists(p)) continue;
                try { Core.DeleteFile(p); deleted++; }
                catch (Exception ex) { stuck.Add(f); Core.Log("Delete failed " + f + ": " + Core.Msg(ex)); }
            }
            if (stuck.Count == 0) { Stamp.Clear(); MyWords.SaveApplied(null); }   // otherwise the window goes on showing what is still on the computer

            if (stuck.Count > 0)
            {
                res.Ok = false;
                res.Title = "ההסרה לא הושלמה";
                res.Text = "לא הצלחתי למחוק את הקובץ " + string.Join(", ", stuck.ToArray()) + " (ייתכן שוורד או תוכנה אחרת משתמשים בו).\r\n\r\n"
                         + "סגרו את וורד ולחצו שוב על \"הסר מוורד\". אם זה חוזר, הפעילו מחדש את המחשב ונסו שוב.";
            }
            else
            {
                res.Ok = true;
                res.Title = "ההסרה הושלמה";
                res.Text = (removedInWord + deleted > 0)
                    ? "המילון התורני הוסר. אם וורד פתוח, סגרו אותו ופתחו מחדש."
                    : "לא נמצא מילון תורני מותקן.";
            }
            if (customPath != null)
            {
                res.Text += "\r\n\r\n" + KeptText(theirs.Count) + " למילון האישי של וורד (CUSTOM.DIC), כדי שלא יימחקו.";
                res.KeptWords = theirs;
            }
            res.RepeatBack = repeatBack;
            if (repeatBack) res.Text += "\r\n\r\nוורד יסמן שוב מילים כפולות" + (forceRepeated ? "." : ", כמו לפני שהתוכנה כיבתה את הסימון.");
            if (problem != null) res.Text += "\r\n\r\n(הערה: " + problem + ")";
            return res;
        }

        static string KeptText(int n)
        {
            return n == 1 ? "מילה אחת שהוספתם בעצמכם לוורד הועברה" : n.ToString("N0") + " מילים שהוספתם בעצמכם לוורד הועברו";
        }

        // "remove from Word" of the window (3.1.0.2, the user: "the removal should also remove what earlier versions put in"): Remove, and
        // then what earlier versions left behind: the words of the Torah dictionary in CUSTOM.DIC (version 2.0 wrote its words there; also
        // the words that have left the dictionary since) and the old entries in Word's list of dictionaries (old file names, the size
        // check). Office's RoamingCustom.dic - the words the person added in Word - is not touched here (the advanced removal offers it).
        // As in the advanced removal: only while no Office program runs, and a copy of everything that changes is saved first.
        public static Result RemoveFull(bool useWord, ProgressFn progress)
        {
            if (progress == null) progress = delegate (int p, string what) { };
            Result res = Remove(useWord);
            if (!res.Ok) return res;
            HashSet<string> kept = new HashSet<string>(StringComparer.Ordinal);
            if (res.KeptWords != null) foreach (string w in res.KeptWords) kept.Add(w);
            progress(70, "מנקה מה שנשאר מגרסאות קודמות…");
            bool wordClosed = WordCom.WaitUntilNoWord(WordCom.CloseWaitMs);
            List<string> office = wordClosed ? WordCom.OtherOfficeRunning() : new List<string>();
            if (!wordClosed || office.Count > 0)
            {
                string who = !wordClosed ? "וורד עדיין פתוח" : string.Join(", ", office.ToArray()) + (office.Count == 1 ? " פתוח" : " פתוחים");
                res.Text += "\r\n\r\n" + who + ", ולכן לא ניקיתי את מה שגרסאות קודמות השאירו (מילים ב-CUSTOM.DIC ורשומות ישנות ברשימת המילונים של וורד). "
                          + "כדי לנקות גם אותם: סגרו " + (!wordClosed ? "את וורד" : (office.Count == 1 ? "אותו" : "אותם")) + " ולחצו שוב על \"הסר מוורד\".";
                return res;
            }
            string backup = Path.Combine(Core.UProofDir(), "Backup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture));
            StringBuilder note = new StringBuilder();
            string problem;
            List<string> lines = RepairWordList(backup, true, out problem);
            int entries = lines == null ? 0 : lines.FindAll(delegate (string l) { return l.StartsWith("deleted", StringComparison.Ordinal); }).Count;
            if (entries > 0)
                note.Append("word-dictionary-list.reg, word-dictionary-list.txt – רשימת המילונים של וורד כפי שהייתה לפני הניקוי. כדי להחזיר אותה: סגרו את וורד ולחצו פעמיים על word-dictionary-list.reg.\r\n");
            List<string> failed;
            List<KeyValuePair<string, string>> copies;
            int words = CleanPersonal(backup, kept, false, out failed, out copies);
            foreach (KeyValuePair<string, string> c in copies)
                note.Append(Path.GetFileName(c.Key) + " – העותק של " + c.Value + " לפני שהמילים של המילון התורני יצאו ממנו. כדי להחזיר: סגרו את וורד והעתיקו אותו בחזרה למקומו.\r\n");
            if (words > 0 || entries > 0)
            {
                if (res.Text.StartsWith("לא נמצא מילון תורני מותקן", StringComparison.Ordinal)) res.Text = "המילון התורני עצמו לא היה מותקן.";
                res.Text += "\r\n\r\nנוקה גם מה שנשאר מגרסאות קודמות:"
                          + (words > 0 ? "\r\n• מ-CUSTOM.DIC " + (words == 1 ? "יצאה מילה אחת" : "יצאו " + words.ToString("N0") + " מילים") + " של המילון התורני. שאר המילים שלכם נשארו." : "")
                          + (entries > 0 ? "\r\n• מרשימת המילונים של וורד " + (entries == 1 ? "יצאה רשומה ישנה אחת." : "יצאו " + entries + " רשומות ישנות.") : "");
            }
            if (lines == null) res.Text += "\r\n\r\n(את הרשומות הישנות ברשימת המילונים של וורד לא ניקיתי: " + problem + ")";
            if (failed.Count > 0) res.Text += "\r\n\r\n(לא הצלחתי לנקות את " + string.Join(", ", failed.ToArray()) + "; הוא נשאר כמו שהיה.)";
            bool backedUp = false;
            try { backedUp = Directory.Exists(backup) && Directory.GetFileSystemEntries(backup).Length > 0; } catch (Exception) { }
            if (backedUp)
            {
                try
                {
                    File.WriteAllText(Path.Combine(backup, "README.txt"), "גיבוי של ההסרה – " + DateTime.Now.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture)
                                      + ", המילון התורני " + Core.AppVersion + "\r\n\r\n" + note.ToString(), new UTF8Encoding(true));
                }
                catch (Exception ex) { Core.Log("Could not write the backup note: " + Core.Msg(ex)); }
                res.Text += "\r\n\r\nעותק גיבוי של מה ששונה נשמר בתיקייה:\r\n" + backup;
            }
            else try { if (Directory.Exists(backup)) Directory.Delete(backup); } catch (Exception) { }
            return res;
        }

        // advanced removal: EVERY custom dictionary leaves Word (only Word's own built-in dictionary stays),
        // including dictionaries and words this program did not add. A backup copy is saved first.
        public static Result RemoveAdvanced(bool useWord) { return RemoveEverything(useWord, null); }

        // sharedBackup: the backup folder of the choices of the advanced removal (null = a folder of its own)
        static Result RemoveEverything(bool useWord, string sharedBackup)
        {
            Result res = new Result();
            Core.Log("==== advanced removal");
            string dir = Core.UProofDir();
            string backup = sharedBackup ?? Path.Combine(dir, "Backup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture));
            Directory.CreateDirectory(backup);
            int unregistered = 0, fromRegistry = 0;
            List<string> refused = new List<string>();
            string problem = null;
            // where every dictionary came from, so that it can be put back by hand (the backup holds the files, this says where they were)
            StringBuilder note = new StringBuilder();
            note.Append("Advanced removal " + DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) + ", program " + Core.AppVersion + "\r\n");

            // 1. ask Word itself (also saves a copy of every dictionary file it lists)
            if (useWord)
            {
                WordSession ws = null;
                try
                {
                    ws = WordCom.Open();
                    foreach (DictInfo di in WordCom.ListAll(ws))
                    {
                        note.Append("Word list: " + di.Name + " | folder from Word: " + (di.Folder ?? "") + " | file: " + (di.FullPath ?? "(not found)") + "\r\n");
                        try
                        {
                            if (di.FullPath != null && File.Exists(di.FullPath))
                                File.Copy(di.FullPath, Path.Combine(backup, Path.GetFileName(di.FullPath)), true);
                        }
                        catch (Exception ex) { Core.Log("Backup failed for " + di.Name + ": " + Core.Msg(ex)); }
                    }
                    unregistered = WordCom.RemoveAll(ws, refused);
                    foreach (string r in refused) note.Append("Word refused to remove: " + r + "\r\n");
                }
                catch (WordMissingException) { }
                catch (Exception ex) { problem = Core.Msg(ex); Core.Log("Advanced removal, Word failed: " + problem); }
                finally { if (ws != null) ws.Dispose(); }
            }

            // 2. registry list that Word reads at start-up
            try
            {
                List<string> gone = RegFallback.RemoveAllDictionaries();
                fromRegistry = gone.Count;
                foreach (string g in gone) note.Append("Registry list: " + g + "\r\n");
                if (gone.Count > 0) File.WriteAllLines(Path.Combine(backup, "registry-dictionary-list.txt"), gone.ToArray(), Encoding.UTF8);
            }
            catch (Exception ex) { Core.Log("Registry (advanced) failed: " + Core.Msg(ex)); }

            // 3. the dictionary files themselves (otherwise Word would pick CUSTOM.DIC up again by itself)
            int moved = 0;
            List<string> locked = new List<string>();
            try
            {
                foreach (string f in Directory.GetFiles(dir))
                {
                    if (!f.EndsWith(".dic", StringComparison.OrdinalIgnoreCase)) continue;   // CUSTOM.DIC, custom.dic ... any case
                    string leaf = Path.GetFileName(f);
                    try
                    {
                        File.Copy(f, Path.Combine(backup, leaf), true);
                        note.Append("File in the UProof folder: " + f + "\r\n");
                        try { Core.DeleteFile(f); moved++; }
                        catch (Exception)
                        {
                            Core.WriteDic(f, new List<string>());   // cannot delete (Word holds it): empty it instead, the copy is safe
                            moved++;
                        }
                    }
                    catch (Exception ex) { locked.Add(leaf); Core.Log("Could not clear " + leaf + ": " + Core.Msg(ex)); }
                }
            }
            catch (Exception ex) { Core.Log("Advanced removal, file step failed: " + Core.Msg(ex)); }

            Stamp.Clear();

            bool any = (unregistered + moved + fromRegistry) > 0;
            bool backedUp = false;   // copies were made (also when Word then refused to let a dictionary go)
            try { backedUp = Directory.GetFiles(backup).Length > 0; } catch (Exception) { }
            if (any || backedUp)
            {
                try { File.WriteAllText(Path.Combine(backup, "dictionaries.txt"), note.ToString(), new UTF8Encoding(true)); }   // with a BOM, so Notepad shows it right
                catch (Exception ex) { Core.Log("Could not write the backup note: " + Core.Msg(ex)); }
            }
            res.Ok = locked.Count == 0 && refused.Count == 0 && problem == null;
            res.Title = res.Ok ? "ההסרה המתקדמת הושלמה" : "ההסרה הושלמה חלקית";
            StringBuilder t = new StringBuilder();
            if (!any && res.Ok) t.Append("לא נמצאו מילונים מותאמים להסרה.\r\n\r\n");
            else if (res.Ok) t.Append("כל המילונים המותאמים אישית הוסרו מוורד. נשאר רק המילון המובנה של וורד.\r\n\r\n");
            else if (any) t.Append("חלק מהמילונים המותאמים אישית הוסרו מוורד.\r\n\r\n");
            else t.Append("לא הצלחתי להסיר את המילונים המותאמים אישית.\r\n\r\n");
            if (any || backedUp) t.Append("עותק גיבוי של כל מה שהוסר נשמר בתיקייה:\r\n" + backup + "\r\n(הקובץ dictionaries.txt שבה אומר מאיפה הגיע כל מילון.)\r\n\r\nאם וורד פתוח, סגרו אותו ופתחו מחדש.");
            else try { Directory.Delete(backup); } catch (Exception) { }
            if (refused.Count > 0) t.Append("\r\n\r\nוורד לא הסכים להסיר: " + string.Join(", ", refused.ToArray()) + ". סגרו את וורד ולחצו שוב.");
            if (locked.Count > 0) t.Append("\r\n\r\nלא הצלחתי לנקות: " + string.Join(", ", locked.ToArray()) + ". סגרו את וורד ולחצו שוב.");
            if (problem != null) t.Append("\r\n\r\n(הערה: " + problem + ")");
            res.Text = t.ToString();
            StringBuilder d = new StringBuilder();   // the same failures, for the answer of the choices (which names the backup folder itself)
            if (refused.Count > 0) d.Append("  וורד לא הסכים להסיר: " + string.Join(", ", refused.ToArray()) + ". סגרו את וורד ולחצו שוב.\r\n");
            if (locked.Count > 0) d.Append("  לא הצלחתי לנקות: " + string.Join(", ", locked.ToArray()) + ". סגרו את וורד ולחצו שוב.\r\n");
            if (problem != null) d.Append("  (הערה: " + problem + ")\r\n");
            res.Details = d.ToString();
            return res;
        }

        // 3.1 (the user, 9-11): the advanced removal is a set of choices. Each one is done only when chosen; a copy of everything that is
        // changed goes to one backup folder first, and the answer names it.
        public static Result RemoveAdvanced(bool useWord, AdvancedOptions o) { return RemoveAdvanced(useWord, o, null); }

        public static Result RemoveAdvanced(bool useWord, AdvancedOptions o, ProgressFn progress)
        {
            if (progress == null) progress = delegate (int p, string what) { };
            Core.Log("==== advanced removal (choices): ours " + o.RemoveOurs + ", personal " + o.CleanPersonal + ", list " + o.RepairList
                     + ", repeated " + o.RestoreRepeated + ", all " + o.RemoveAll);
            string backup = Path.Combine(Core.UProofDir(), "Backup_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture));
            StringBuilder t = new StringBuilder();
            StringBuilder note = new StringBuilder();   // README.txt of the backup folder: what each copy is and where it was
            bool ok = true, oursGone = false, repeatDone = false, allDone = false;
            HashSet<string> kept = new HashSet<string>(StringComparer.Ordinal);   // the person's words moved to CUSTOM.DIC: never taken out again

            // 1. the Torah dictionary leaves Word, as with "remove from Word" (the person's own words in it are kept in CUSTOM.DIC)
            if (o.RemoveOurs)
            {
                progress(10, "מסיר את המילון התורני מוורד…");
                Result r = Remove(useWord, o.RestoreRepeated, o.RestoreRepeated);
                ok &= r.Ok;
                oursGone = r.Ok;
                repeatDone = r.RepeatBack;
                if (r.KeptWords != null) foreach (string w in r.KeptWords) kept.Add(w);
                if (r.Ok)
                    t.Append("• המילון התורני הוסר מוורד." + (kept.Count > 0 ? " " + KeptText(kept.Count) + " ל-CUSTOM.DIC, כדי שלא יימחקו." : "") + "\r\n");
                else t.Append("• המילון התורני: " + r.Title + ". " + Installer.FirstParagraph(r.Text) + "\r\n");
            }

            // the marking of repeated words: as chosen, also when the person (not this program) switched it off
            if (o.RestoreRepeated)
            {
                if (!repeatDone && useWord)
                {
                    progress(25, "מחזיר את סימון המילים הכפולות…");
                    Result rr = SetRepeatedWords(true);
                    repeatDone = rr.Ok;
                }
                if (repeatDone) t.Append("• וורד יסמן שוב מילים כפולות.\r\n");
                else
                {
                    ok = false;
                    t.Append("• את סימון המילים הכפולות לא הצלחתי להחזיר. אפשר להחזיר אותו בוורד: קובץ ← אפשרויות ← הגהה, ולסמן את האפשרות של מילים חוזרות.\r\n");
                }
            }

            // 2. every custom dictionary (the old advanced removal), into the same backup folder
            if (o.RemoveAll)
            {
                progress(35, "מסיר את כל המילונים המותאמים אישית…");
                Result r = RemoveEverything(useWord, backup);
                ok &= r.Ok;
                allDone = true;
                t.Append("• " + r.Title + ": " + Installer.FirstParagraph(r.Text) + "\r\n");
                if (!string.IsNullOrEmpty(r.Details)) t.Append(r.Details);
            }

            // 3 + 4. Word's own list of dictionaries and its personal dictionaries: only while no Word runs - a running Word writes both back
            // when it closes. The other Office programs do the same (Outlook even adds names to the personal dictionary), so not while one
            // of them runs either.
            bool roamingChanged = false;
            if (o.RepairList || o.CleanPersonal)
            {
                progress(55, "ממתין שוורד ייסגר…");
                bool wordClosed = WordCom.WaitUntilNoWord(WordCom.CloseWaitMs);
                List<string> office = wordClosed ? WordCom.OtherOfficeRunning() : new List<string>();
                if (!wordClosed || office.Count > 0)
                {
                    ok = false;
                    string who = !wordClosed ? "וורד עדיין פתוח" : string.Join(", ", office.ToArray()) + (office.Count == 1 ? " פתוח" : " פתוחים");
                    t.Append("• " + who + ", ולכן לא " + (o.RepairList ? "תיקנתי את רשימת המילונים של וורד" : "") + (o.RepairList && o.CleanPersonal ? " ולא " : "")
                             + (o.CleanPersonal ? "ניקיתי את המילון האישי" : "") + (wordClosed ? " (תוכנות אופיס כותבות אותם מחדש כשהן נסגרות)" : "")
                             + ". סגרו " + (!wordClosed ? "את וורד" : (office.Count == 1 ? "אותו" : "אותם")) + " ולחצו שוב.\r\n");
                }
                else
                {
                    if (o.RepairList)
                    {
                        progress(70, "מתקן את רשימת המילונים של וורד…");
                        string problem;
                        List<string> lines = RepairWordList(backup, oursGone || o.RemoveAll, out problem);   // our entry goes only when our file went
                        if (lines == null) { ok = false; t.Append("• את רשימת המילונים של וורד לא תיקנתי: " + problem + "\r\n"); }
                        else
                        {
                            int deleted = lines.FindAll(delegate (string l) { return l.StartsWith("deleted", StringComparison.Ordinal); }).Count;
                            if (deleted > 0)
                            {
                                t.Append("• רשימת המילונים של וורד תוקנה: " + (deleted == 1 ? "יצאה ממנה רשומה ישנה אחת שחסמה" : "יצאו ממנה " + deleted + " רשומות ישנות שחסמו")
                                         + " הוספת מילונים.\r\n");
                                note.Append("word-dictionary-list.reg, word-dictionary-list.txt – רשימת המילונים של וורד כפי שהייתה לפני התיקון. כדי להחזיר אותה: סגרו את וורד ולחצו פעמיים על word-dictionary-list.reg.\r\n");
                            }
                            else t.Append("• רשימת המילונים של וורד תקינה: לא היו בה רשומות ישנות.\r\n");
                        }
                    }
                    if (o.CleanPersonal)
                    {
                        progress(85, "מנקה את המילון האישי של וורד…");
                        List<string> failed;
                        List<KeyValuePair<string, string>> copies;
                        int n = CleanPersonal(backup, kept, out failed, out copies);
                        foreach (KeyValuePair<string, string> c in copies)
                        {
                            note.Append(Path.GetFileName(c.Key) + " – העותק של " + c.Value + " לפני שהמילים של המילון התורני יצאו ממנו.\r\n");
                            if (Path.GetFileName(c.Value).StartsWith("RoamingCustom", StringComparison.OrdinalIgnoreCase)) roamingChanged = true;
                        }
                        if (failed.Count > 0)
                        {
                            ok = false;
                            t.Append("• לא הצלחתי לנקות את " + string.Join(", ", failed.ToArray()) + " (ייתכן שתוכנה אחרת משתמשת בו, או שהוא מוגן מכתיבה). הוא נשאר כמו שהיה.\r\n");
                        }
                        if (n > 0)
                            t.Append("• מהמילון האישי של וורד " + (n == 1 ? "יצאה מילה אחת" : "יצאו " + n.ToString("N0") + " מילים") + " של המילון התורני. שאר המילים שלכם נשארו.\r\n");
                        else if (failed.Count == 0) t.Append("• במילון האישי של וורד לא היו מילים של המילון התורני.\r\n");
                    }
                }
            }

            progress(97, "");
            bool backedUp = false;
            try { backedUp = Directory.Exists(backup) && Directory.GetFileSystemEntries(backup).Length > 0; } catch (Exception) { }
            if (backedUp)
            {
                if (allDone) note.Append("dictionaries.txt – מאיפה הגיע כל מילון שהוסר ב\"מחיקת כל המילונים המותאמים אישית\" (הקבצים עצמם כאן לידו).\r\n");
                note.Append("כדי להחזיר מילון: סגרו את וורד והעתיקו את הקובץ מכאן בחזרה למקום שבו היה.\r\n");
                try
                {
                    File.WriteAllText(Path.Combine(backup, "README.txt"), "גיבוי של ההסרה המתקדמת – " + DateTime.Now.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture)
                                      + ", המילון התורני " + Core.AppVersion + "\r\n\r\n" + note.ToString(), new UTF8Encoding(true));
                }
                catch (Exception ex) { Core.Log("Could not write the backup note: " + Core.Msg(ex)); }
                t.Append("\r\nעותק גיבוי של מה ששונה נשמר בתיקייה:\r\n" + backup + "\r\n(הקובץ README.txt שבה אומר מה כל קובץ ואיך מחזירים.)");
            }
            else try { if (Directory.Exists(backup)) Directory.Delete(backup); } catch (Exception) { }
            if (roamingChanged)
                t.Append("\r\n\r\nהמילון האישי של אופיס מסונכרן עם חשבון Microsoft. אם המילים חוזרות אליו, אפשר למחוק אותן גם בוורד: קובץ ← אפשרויות ← הגהה ← מילונים מותאמים אישית ← ערוך רשימת מילים.");
            t.Append("\r\n\r\nאם וורד פתוח, סגרו אותו ופתחו מחדש.");
            Result res = new Result();
            res.Ok = ok;
            res.Title = ok ? "ההסרה המתקדמת הושלמה" : "ההסרה המתקדמת הושלמה חלקית";
            res.Text = t.ToString().TrimStart();
            return res;
        }

        // the first paragraph of a result text
        internal static string FirstParagraph(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            int cut = text.IndexOf("\r\n\r\n", StringComparison.Ordinal);
            return (cut >= 0 ? text.Substring(0, cut) : text).Replace("\r\n", " ").Trim();
        }

        // Word's own list of dictionaries, as the repair tool of 9-11 did (3.1 advanced removal): first a copy of the list - every value in a
        // text file, and a .reg file that puts it back with a double click (where Windows has reg.exe) -, then the entries of old versions of
        // this program, of the size check and of word checks go (with includeMain also the entry of the dictionary file itself).
        // null = not done (problem says why): without a copy nothing is changed.
        internal static List<string> RepairWordList(string backup, bool includeMain, out string problem)
        {
            problem = null;
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegFallback.KeyPath, true))
                {
                    if (k == null) return new List<string>();
                    Directory.CreateDirectory(backup);
                    StringBuilder sb = new StringBuilder();
                    foreach (string vn in k.GetValueNames())
                    {
                        object v = k.GetValue(vn, null, RegistryValueOptions.DoNotExpandEnvironmentNames);
                        byte[] b = v as byte[];
                        sb.Append(vn + " = " + (b != null ? BitConverter.ToString(b) : Convert.ToString(v, CultureInfo.InvariantCulture)) + "\r\n");
                    }
                    string txt = Path.Combine(backup, "word-dictionary-list.txt"), reg = Path.Combine(backup, "word-dictionary-list.reg");
                    File.WriteAllText(txt, sb.ToString(), new UTF8Encoding(true));
                    ExportRegistryList(reg);
                    List<string> lines = DictList.Clean(k, includeMain);
                    foreach (string l in lines) Core.Log("Word's list (advanced removal): " + l);
                    if (lines.Count == 0)   // nothing changed: no copy is needed
                        foreach (string f in new string[] { txt, reg }) try { if (File.Exists(f)) File.Delete(f); } catch (Exception) { }
                    return lines;
                }
            }
            catch (Exception ex)
            {
                problem = Core.Msg(ex);
                Core.Log("Word's list could not be repaired: " + problem);
                return null;
            }
        }

        // a .reg copy of Word's list of dictionaries (Windows only; the text copy has every value anyway)
        static void ExportRegistryList(string file)
        {
            try
            {
                string reg = string.IsNullOrEmpty(Environment.SystemDirectory) ? null : Path.Combine(Environment.SystemDirectory, "reg.exe");
                if (reg == null || !File.Exists(reg)) return;
                ProcessStartInfo si = new ProcessStartInfo(reg, "export \"HKCU\\" + RegFallback.KeyPath + "\" \"" + file + "\" /y");
                si.UseShellExecute = false; si.CreateNoWindow = true;
                using (Process p = Process.Start(si))
                    if (!p.WaitForExit(30000)) { try { p.Kill(); } catch (Exception) { } Core.Log("reg export did not finish"); }
            }
            catch (Exception ex) { Core.Log("reg export failed: " + Core.Msg(ex)); }
        }

        // The words of the Torah dictionary leave Word's own personal dictionaries (3.1, the user: "also from Word's permanent dictionary"):
        // Office's RoamingCustom.dic - the dictionary of "Add to Dictionary" - and CUSTOM.DIC of the UProof folder. Every other word stays.
        // CUSTOM.DIC also loses the words that left the dictionary in 3.1 (Core.Retired): version 2.0 of the dictionary wrote its words
        // there, and those words came from it. RoamingCustom.dic keeps them: no version of this program wrote there, so the person put
        // them there. keep = words never taken out (the person's own words that the removal has just moved to CUSTOM.DIC).
        // A copy of each file that changes goes to the backup folder first (copies: copy -> the file it was), and a file is written whole
        // or not at all. Returns how many words left them; failed = the files that could not be changed (they are as they were).
        internal static int CleanPersonal(string backup, HashSet<string> keep, out List<string> failed, out List<KeyValuePair<string, string>> copies)
        { return CleanPersonal(backup, keep, true, out failed, out copies); }

        // roaming = false: CUSTOM.DIC only (the plain removal: what earlier versions of this program wrote)
        internal static int CleanPersonal(string backup, HashSet<string> keep, bool roaming, out List<string> failed, out List<KeyValuePair<string, string>> copies)
        {
            failed = new List<string>();
            copies = new List<KeyValuePair<string, string>>();
            HashSet<string> ours = new HashSet<string>(Core.BuildWords(true, true), StringComparer.Ordinal);
            HashSet<string> oursAndOld = new HashSet<string>(ours, StringComparer.Ordinal);
            foreach (string w in Core.Retired) oursAndOld.Add(w);
            List<string> files = roaming ? new List<string>(DictList.RoamingFiles()) : new List<string>();
            if (roaming) foreach (string f in PersonalFilesHook != null ? PersonalFilesHook() : new string[0]) if (!files.Contains(f)) files.Add(f);
            try
            {
                foreach (string f in Directory.GetFiles(Core.UProofDir()))
                    if (string.Equals(Path.GetFileName(f), "CUSTOM.DIC", StringComparison.OrdinalIgnoreCase) && !files.Contains(f)) files.Add(f);
            }
            catch (Exception) { }
            int removed = 0;
            foreach (string f in files)
            {
                try
                {
                    HashSet<string> set = string.Equals(Path.GetFileName(f), "CUSTOM.DIC", StringComparison.OrdinalIgnoreCase) ? oursAndOld : ours;
                    List<string> words = Core.ReadDicWords(f);
                    List<string> stay = words.FindAll(delegate (string w)
                    {
                        return (keep != null && keep.Contains(w)) || !set.Contains(w.Replace('״', '"').Replace('׳', '\''));
                    });
                    int n = words.Count - stay.Count;
                    if (n == 0) continue;
                    Directory.CreateDirectory(backup);
                    string copy = Path.Combine(backup, Path.GetFileName(f));
                    for (int i = 2; File.Exists(copy); i++) copy = Path.Combine(backup, Path.GetFileNameWithoutExtension(f) + " (" + i + ")" + Path.GetExtension(f));
                    File.Copy(f, copy);
                    copies.Add(new KeyValuePair<string, string>(copy, f));
                    Core.ReplaceDic(f, stay);
                    removed += n;
                    Core.Log("Personal dictionary " + f + ": " + n + " word(s) of the Torah dictionary out, " + stay.Count + " kept");
                }
                catch (Exception ex) { failed.Add(Path.GetFileName(f)); Core.Log("Personal dictionary " + f + " could not be cleaned: " + Core.Msg(ex)); }
            }
            return removed;
        }

        // Only for the self-test: more personal dictionary files (an Office RoamingCustom.dic outside Word's list)
        public static Func<string[]> PersonalFilesHook;

        // Word option "Flag repeated words": false = Word stops marking repeated words (all of them), true = marking is back on
        public static Result SetRepeatedWords(bool flag)
        {
            Result res = new Result();
            Core.Log("==== repeated words flag -> " + flag);
            WordSession ws = null;
            try
            {
                ws = WordCom.Open();
                res.WordWasOpen = !ws.Created;
                bool ok = WordCom.SetRepeatWords(ws, flag);
                res.Ok = ok;
                if (ok)
                {
                    Stamp.WriteRepeatOff(!flag);
                    res.Title = flag ? "סימון מילים כפולות הוחזר" : "סימון מילים כפולות כובה";
                    res.Text = flag
                        ? "וורד יסמן שוב מילים שחוזרות פעמיים ברצף."
                        : "וורד כבר לא יסמן כשגיאה מילים שחוזרות פעמיים ברצף - אף אחת מהן, בכל המסמכים.\r\n\r\n"
                          + (res.WordWasOpen ? "וורד היה פתוח, והשינוי חל עכשיו." : "השינוי יחול בפעם הבאה שתפתחו את וורד.");
                }
                else
                {
                    res.Title = "לא הצלחתי לשנות";
                    res.Text = "וורד לא קיבל את השינוי. אפשר לעשות זאת ידנית: בוורד, קובץ ← אפשרויות ← הגהה, ולהסיר את הסימון מהאפשרות של מילים כפולות.";
                }
            }
            catch (WordMissingException)
            {
                res.Ok = false; res.Title = "וורד לא נמצא";
                res.Text = "לא מצאתי וורד מותקן במחשב הזה, ולכן אי אפשר לשנות את ההגדרה.";
            }
            catch (Exception ex)
            {
                Core.Log("RepeatWord failed: " + Core.Msg(ex));
                res.Ok = false; res.Title = "לא הצלחתי לשנות";
                res.Text = "לא הצלחתי להתחבר לוורד.\r\n\r\nאפשר לעשות זאת ידנית: בוורד, קובץ ← אפשרויות ← הגהה, ולהסיר את הסימון מהאפשרות של מילים כפולות.\r\n\r\nפרטי השגיאה: " + Core.Msg(ex);
            }
            finally { if (ws != null) ws.Dispose(); }
            return res;
        }

        // Manual route: write the files and let the user add them from Word's own settings screen.
        public static List<string> WriteForManual(bool extended, bool compose)
        {
            string dir = Core.UProofDir();
            Directory.CreateDirectory(dir);
            string mainPath = Path.Combine(dir, Core.MainFile);
            List<string> words = Core.BuildWords(extended, compose);
            List<WordChange> mine;
            HashSet<string> alsoOurs = MyChanges(ref words, out mine);
            List<string> theirs = new List<string>();
            try { theirs = Core.ForeignWords(mainPath, alsoOurs); }
            catch (Exception ex) { Core.Log("Reading the old dictionary file failed: " + Core.Msg(ex)); }
            List<string> toWrite = words;
            if (theirs.Count > 0)
            {
                toWrite = new List<string>(words);
                toWrite.AddRange(theirs);
                Core.Log("Manual: kept " + theirs.Count + " word(s) the person had added");
            }
            Core.WriteDic(mainPath, toWrite);
            if (mine != null) MyWords.SaveApplied(MyWords.Additions(mine));
            foreach (string f in Core.LegacyFiles)
            {
                string p = Path.Combine(dir, f);
                try { if (File.Exists(p)) Core.DeleteFile(p); } catch (Exception) { }
            }
            try { RegFallback.RemoveLegacy(); } catch (Exception) { }   // an entry for our one file may be what makes Word use it: it stays
            Stamp.Write(extended, compose);
            return new List<string> { Core.MainFile };
        }
    }

    // ---------------------------------------------------------------- after installing: does Word itself now know the words?
    // Word is asked the same question that decides the red underline, for sample words spread over the whole file.
    // Their places in the file are reported too, so if Word stops at some size we can see where.
    public static class Verify
    {
        static readonly string[] Samples = new string[] {
            "מיבעי", "באנפא", "דבוקין", "דהעניין", "דמסתברא", "החששא", "ויוצרא", "וסרכא",
            "כדסבירא", "לוחתא", "סמיותא", "תכסיסא", "ודהנושא", "דהמכתב", "מג\"א" };
        const string NotAWord = "קזקזקזק";     // Word must mark it, otherwise the check itself does not work
        const string CommonWord = "שלום";       // Word must not mark it

        // true = Word marks the word as a spelling mistake
        static bool Marked(object content, string word)
        {
            WordCom.Put(content, "Text", word);
            WordCom.Put(content, "LanguageID", 1037);                 // Hebrew
            try { WordCom.Put(content, "NoProofing", false); } catch (Exception) { }
            object errs = WordCom.Get(content, "SpellingErrors");
            return Convert.ToInt32(WordCom.Get(errs, "Count")) > 0;
        }

        // verdict: 1 = every sample word is known, 0 = some are not, -1 = the check could not be trusted
        public static string Run(WordSession s, List<string> installed, bool fresh, out int verdict)
        {
            verdict = -1;
            Dictionary<string, int> place = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int i = 0; i < installed.Count; i++) if (!place.ContainsKey(installed[i])) place[installed[i]] = i + 1;

            List<string> samples = new List<string>();
            foreach (string w in Samples) if (place.ContainsKey(w)) samples.Add(w);
            for (int i = installed.Count - 1; i >= 0 && i > installed.Count - 5000; i--)   // and one of the very last words in the file
            {
                string w = installed[i];
                if (w.Length >= 5 && w.IndexOf('"') < 0 && !samples.Contains(w)) { samples.Add(w); break; }
            }

            StringBuilder dicts = new StringBuilder();
            try
            {
                foreach (DictInfo di in WordCom.ListAll(s))
                {
                    string count = "?";
                    try { if (di.FullPath != null && File.Exists(di.FullPath)) count = Core.CountDicWords(di.FullPath).ToString("N0"); } catch (Exception) { }
                    if (dicts.Length > 0) dicts.Append(", ");
                    dicts.Append(di.Name + " (" + count + ")");
                }
            }
            catch (Exception ex) { dicts.Append("? (" + Core.Msg(ex) + ")"); }

            object doc = null;
            List<string> unknown = new List<string>();
            bool controlsOk;
            try
            {
                object docs = WordCom.Get(s.App, "Documents");
                doc = WordCom.Call(docs, "Add", Missing.Value, Missing.Value, Missing.Value, false);   // invisible document
                object content = WordCom.Get(doc, "Content");
                bool notAWordMarked = Marked(content, NotAWord);
                bool commonMarked = Marked(content, CommonWord);
                controlsOk = notAWordMarked && !commonMarked;
                Core.Log("Verify controls: nonsense marked=" + notAWordMarked + ", common word marked=" + commonMarked);
                foreach (string w in samples)
                {
                    bool m = Marked(content, w);
                    Core.Log("Verify " + w + " (place " + place[w] + "): " + (m ? "MARKED" : "ok"));
                    if (m) unknown.Add(w);
                }
            }
            finally
            {
                if (doc != null) { try { WordCom.Call(doc, "Close", 0); } catch (Exception) { } }
            }

            string where = fresh ? "בוורד חדש שנפתח לבדיקה" : "בוורד שהיה פתוח";
            StringBuilder t = new StringBuilder();
            if (!controlsOk)
            {
                t.Append("בדיקה אוטומטית: לא הצלחתי לבדוק דרך וורד, צריך לבדוק ידנית.");
            }
            else if (unknown.Count == 0)
            {
                verdict = 1;
                t.Append("בדיקה אוטומטית " + where + ": כל " + samples.Count + " מילות הבדיקה מזוהות ✓");
            }
            else
            {
                verdict = 0;
                t.Append("בדיקה אוטומטית " + where + ": " + (samples.Count - unknown.Count) + " מתוך " + samples.Count + " מילות בדיקה מזוהות.\r\nלא זוהו: ");
                for (int i = 0; i < unknown.Count; i++)
                    t.Append((i > 0 ? ", " : "") + unknown[i] + " (מילה " + place[unknown[i]].ToString("N0") + ")");
            }
            t.Append("\r\nמילונים בוורד: " + dicts);
            Core.Log(t.ToString());
            return t.ToString();
        }
    }

    // ================================================================ stage 4: the person's own changes, and suggestions to the dictionary's owner

    // One change the person made: '+' = a word added, '-' = a word of the dictionary removed,
    // '~' = a word of the dictionary corrected (Word = the wrong spelling, NewWord = the right one: the wrong one goes, the right one comes).
    public class WordChange
    {
        public char Kind;
        public string Word;
        public string NewWord;
        public bool Sent;   // already sent as a suggestion to the dictionary's owner
        public WordChange(char kind, string word, string newWord) { Kind = kind; Word = word; NewWord = newWord; }
        public WordChange Copy() { WordChange c = new WordChange(Kind, Word, NewWord); c.Sent = Sent; return c; }
    }

    public enum EditResult { Done, Invalid, AlreadyThere, NotThere, Same }

    // The person's changes stay on this computer, outside Word's folder, and every installation (also of a new version) applies them:
    //   %APPDATA%\TorahDictionaryForWord\my-words.txt   one change per line: "+word", "-word", "~wrong<TAB>right"; "<TAB>sent" at the end = sent
    //   %APPDATA%\TorahDictionaryForWord\applied.txt    the person's words that the last installation wrote into our file
    //                                                   (so that one taken back later is not kept as a word that Word added by itself)
    public static class MyWords
    {
        public static string FolderOverride;   // the self-test
        public const string FileName = "my-words.txt";
        const string AppliedName = "applied.txt";

        public static string Folder()
        {
            if (FolderOverride != null) return FolderOverride;
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "TorahDictionaryForWord");
        }

        public static List<WordChange> Load()
        {
            List<WordChange> list = new List<WordChange>();
            string p = Path.Combine(Folder(), FileName);
            if (!File.Exists(p)) return list;
            foreach (string raw in File.ReadAllLines(p, Encoding.UTF8))
            {
                string line = raw.TrimEnd('\r');
                if (line.Length < 2 || line[0] == '#') continue;
                char kind = line[0];
                if (kind != '+' && kind != '-' && kind != '~') continue;
                string[] parts = line.Substring(1).Split(new char[] { '\t' });
                WordChange c = new WordChange(kind, WordEdits.Clean(parts[0]), null);
                if (c.Word == null) continue;   // not a word (the file was changed by hand?): left out
                int next = 1;
                if (kind == '~')
                {
                    c.NewWord = parts.Length > 1 ? WordEdits.Clean(parts[1]) : null;
                    if (c.NewWord == null || c.NewWord == c.Word) continue;
                    next = 2;
                }
                c.Sent = parts.Length > next && parts[next] == "sent";
                list.Add(c);
            }
            return list;
        }

        public static void Save(List<WordChange> changes)
        {
            string dir = Folder();
            Directory.CreateDirectory(dir);
            StringBuilder sb = new StringBuilder();
            sb.Append("# המילון התורני – השינויים שלכם במילון (+ הוספה, - הסרה, ~ תיקון). התוכנה כותבת את הקובץ הזה.\r\n");
            foreach (WordChange c in changes)
            {
                sb.Append(c.Kind).Append(c.Word);
                if (c.Kind == '~') sb.Append('\t').Append(c.NewWord);
                if (c.Sent) sb.Append("\tsent");
                sb.Append("\r\n");
            }
            string p = Path.Combine(dir, FileName), tmp = p + ".tmp";
            File.WriteAllText(tmp, sb.ToString(), new UTF8Encoding(true));   // written whole next to it first, then copied over
            File.Copy(tmp, p, true);
            File.Delete(tmp);
        }

        // the edition as the person has it: removed and corrected words out, the person's words in (at the end, no repeats)
        public static List<string> Apply(List<string> words, List<WordChange> changes)
        {
            if (changes == null || changes.Count == 0) return words;
            HashSet<string> drop = new HashSet<string>(StringComparer.Ordinal);
            List<string> add = new List<string>();
            foreach (WordChange c in changes)
            {
                if (c.Kind == '-') drop.Add(c.Word);
                else if (c.Kind == '~') { drop.Add(c.Word); add.Add(c.NewWord); }
                else if (c.Kind == '+') add.Add(c.Word);
            }
            List<string> result = new List<string>(words.Count + add.Count);
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string w in words) if (!drop.Contains(w) && seen.Add(w)) result.Add(w);
            foreach (string w in add) if (seen.Add(w)) result.Add(w);
            return result;
        }

        // the person's own words: added ones, and the right spelling of a correction
        public static List<string> Additions(List<WordChange> changes)
        {
            List<string> a = new List<string>();
            foreach (WordChange c in changes)
            {
                if (c.Kind == '+') a.Add(c.Word);
                else if (c.Kind == '~') a.Add(c.NewWord);
            }
            return a;
        }

        public static HashSet<string> LoadApplied()
        {
            HashSet<string> s = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                string p = Path.Combine(Folder(), AppliedName);
                if (File.Exists(p))
                    foreach (string l in File.ReadAllLines(p, Encoding.UTF8)) { string w = l.Trim(); if (w.Length > 0) s.Add(w); }
            }
            catch (Exception ex) { Core.Log("Reading the list of applied words failed: " + Core.Msg(ex)); }
            return s;
        }

        public static void SaveApplied(List<string> words)
        {
            try
            {
                string dir = Folder(), p = Path.Combine(dir, AppliedName);
                if (words == null || words.Count == 0) { if (File.Exists(p)) File.Delete(p); return; }
                Directory.CreateDirectory(dir);
                File.WriteAllLines(p, words.ToArray(), new UTF8Encoding(false));
            }
            catch (Exception ex) { Core.Log("Saving the list of applied words failed: " + Core.Msg(ex)); }
        }
    }

    // The changes made in the editing window, on top of an edition: knows the dictionary as it is with them, sorted for the list.
    public class WordEdits
    {
        public readonly List<WordChange> Changes = new List<WordChange>();
        readonly List<string> edition;
        readonly HashSet<string> inEdition;
        HashSet<string> current;
        List<string> sorted;

        public WordEdits(List<string> editionWords, List<WordChange> changes)
        {
            edition = editionWords;
            inEdition = new HashSet<string>(editionWords, StringComparer.Ordinal);
            if (changes != null) foreach (WordChange c in changes) Changes.Add(c.Copy());
            Rebuild();
        }

        void Rebuild()
        {
            current = new HashSet<string>(MyWords.Apply(edition, Changes), StringComparer.Ordinal);
            sorted = new List<string>(current);
            sorted.Sort(StringComparer.Ordinal);
        }

        public int Count { get { return current.Count; } }
        public bool Contains(string w) { return w != null && current.Contains(w); }
        public bool IsPersons(string w) { return w != null && current.Contains(w) && !inEdition.Contains(w); }   // added by the person
        public List<string> Sorted() { return sorted; }

        // every word that has this text in it (vowel points and invisible marks are not counted); empty text = all the words
        public List<string> Find(string text)
        {
            string q = Strip(text ?? "").Trim();
            if (q.Length == 0) return sorted;
            List<string> found = new List<string>();
            foreach (string w in sorted) if (w.IndexOf(q, StringComparison.Ordinal) >= 0) found.Add(w);
            return found;
        }

        void Put(string w)
        {
            if (!current.Add(w)) return;
            int i = sorted.BinarySearch(w, StringComparer.Ordinal);
            if (i < 0) sorted.Insert(~i, w);
        }

        void Take(string w)
        {
            if (!current.Remove(w)) return;
            int i = sorted.BinarySearch(w, StringComparer.Ordinal);
            if (i >= 0) sorted.RemoveAt(i);
        }

        // would the word be in the dictionary with the changes as they are now (the same rule as MyWords.Apply: an addition wins)
        bool Present(string w)
        {
            bool p = inEdition.Contains(w);
            foreach (WordChange c in Changes) if ((c.Kind == '-' || c.Kind == '~') && c.Word == w) p = false;
            foreach (WordChange c in Changes) if ((c.Kind == '+' && c.Word == w) || (c.Kind == '~' && c.NewWord == w)) p = true;
            return p;
        }

        public EditResult Add(string raw)
        {
            string w = Clean(raw);
            if (w == null) return EditResult.Invalid;
            if (current.Contains(w)) return EditResult.AlreadyThere;
            for (int i = Changes.Count - 1; i >= 0; i--)
                if (Changes[i].Kind == '-' && Changes[i].Word == w) Changes.RemoveAt(i);   // it was removed by the person: the removal is taken back
            if (!Present(w)) Changes.Add(new WordChange('+', w, null));
            Put(w);
            return EditResult.Done;
        }

        public EditResult Remove(string raw)
        {
            string w = Clean(raw);
            if (w == null) return EditResult.Invalid;
            if (!current.Contains(w)) return EditResult.NotThere;
            for (int i = Changes.Count - 1; i >= 0; i--)
            {
                WordChange c = Changes[i];
                if (c.Kind == '+' && c.Word == w) Changes.RemoveAt(i);                                  // the person's own word: its addition is taken back
                else if (c.Kind == '~' && c.NewWord == w) Changes[i] = new WordChange('-', c.Word, null); // the right spelling of a correction: only the removal stays
            }
            if (Present(w)) Changes.Add(new WordChange('-', w, null));   // a word of the edition
            Take(w);
            return EditResult.Done;
        }

        public EditResult Fix(string wrongRaw, string rightRaw)
        {
            string a = Clean(wrongRaw), b = Clean(rightRaw);
            if (a == null || b == null) return EditResult.Invalid;
            if (a == b) return EditResult.Same;
            if (!current.Contains(a)) return EditResult.NotThere;
            int own = -1, fixOfFix = -1;
            for (int i = 0; i < Changes.Count; i++)
            {
                if (Changes[i].Kind == '+' && Changes[i].Word == a) own = i;
                if (Changes[i].Kind == '~' && Changes[i].NewWord == a) fixOfFix = i;
            }
            if (own >= 0) { Changes.RemoveAt(own); if (!Present(b)) Changes.Add(new WordChange('+', b, null)); }   // the person's own word, spelled again
            else if (fixOfFix >= 0) Changes[fixOfFix] = new WordChange('~', Changes[fixOfFix].Word, b);            // a correction corrected again
            else Changes.Add(new WordChange('~', a, b));
            if (Present(a)) Changes.Add(new WordChange('-', a, null));
            if (!Present(b)) Changes.Add(new WordChange('+', b, null));
            Take(a);
            Put(b);
            return EditResult.Done;
        }

        public void Undo(int index)
        {
            if (index < 0 || index >= Changes.Count) return;
            Changes.RemoveAt(index);
            Rebuild();
        }

        static string Strip(string raw)
        {
            StringBuilder sb = new StringBuilder(raw.Length);
            foreach (char ch in raw)
            {
                if (ch >= '\u0591' && ch <= '\u05C7' && ch != '\u05BE') continue;   // vowel points and cantillation marks (the maqaf stays)
                if ((ch >= '\u200B' && ch <= '\u200F') || (ch >= '\u202A' && ch <= '\u202E') || (ch >= '\u2066' && ch <= '\u2069') || ch == '\uFEFF') continue;   // invisible marks (direction, joiners, BOM)
                // 3.1.0.1 (the deep check): the Hebrew gershayim and geresh, and the curly quotes that Word puts into documents, become the
                // plain marks that every word of the dictionary has - a word copied from a document is found, and added in the same form
                if (ch == '\u05F4' || ch == '\u201C' || ch == '\u201D' || ch == '\u201E' || ch == '\u2033') { sb.Append('"'); continue; }
                if (ch == '\u05F3' || ch == '\u2018' || ch == '\u2019' || ch == '\u2032') { sb.Append('\''); continue; }
                sb.Append(ch);
            }
            // two apostrophes are the way many people type the gershayim (no word of the dictionary has two in a row)
            return sb.ToString().Replace("''", "\"");
        }

        // one word as the person typed it, cleaned (no spaces around it, no vowel points, no invisible marks, the quote marks as the
        // dictionary writes them); null = not one word: empty, a space inside, a sign that does not belong in a Hebrew word, no Hebrew
        // letter at all, or longer than 60. (Every word of the editions is made of Hebrew letters, " and ' only, so each of them passes as it is.)
        public static string Clean(string raw)
        {
            if (raw == null) return null;
            string w = Strip(raw).Trim();
            // quote marks around the word are not part of it (a word copied with them: the window itself shows a word between quote
            // marks); a geresh at the end of an abbreviation (וכו') stays. No word starts with a quote mark or ends with gershayim.
            if (w.Length > 1 && (w[0] == '"' || w[0] == '\''))
            {
                char q = w[0];
                w = w.Substring(1);
                if (w.Length > 1 && w[w.Length - 1] == q) w = w.Substring(0, w.Length - 1);
            }
            w = w.TrimStart('"', '\'').TrimEnd('"');
            if (w.Length == 0 || w.Length > 60) return null;
            bool letter = false;
            foreach (char ch in w)
            {
                if (ch >= '\u05D0' && ch <= '\u05EA') { letter = true; continue; }   // Hebrew letters (with the final forms)
                bool ok = ch == '\u05BE' || ch == '"' || ch == '\'' || ch == '-';   // (the maqaf; the Hebrew geresh and gershayim are ' and " by now)
                if (!ok) return null;
            }
            return letter ? w : null;
        }
    }

    // A suggestion to the dictionary's owner = a new issue in the GitHub repository of the program. Sent by the program itself, with a key
    // built into it that may only open issues in that one repository; or (a person with a GitHub account, or when sending fails) through the
    // "new issue" page of GitHub in the browser, already filled in.
    public static class Suggest
    {
        public const string Repo = "elicohen8565-hue/milon-torani-3.0";
        const string KeyResource = "Milon.issue.bin";
        static readonly byte[] Mask = Encoding.ASCII.GetBytes("milon-torani-3.0/issues");

        static int Count(List<WordChange> ch, char kind)
        {
            int n = 0;
            foreach (WordChange c in ch) if (c.Kind == kind) n++;
            return n;
        }

        public static string Title(List<WordChange> ch)
        {
            List<string> parts = new List<string>();
            int a = Count(ch, '+'), r = Count(ch, '-'), f = Count(ch, '~');
            if (a > 0) parts.Add(a + " להוסיף");
            if (r > 0) parts.Add(r + " להסיר");
            if (f > 0) parts.Add(f + " לתקן");
            return "הצעה למילון: " + string.Join(", ", parts.ToArray());
        }

        // the free text of a person goes into a public page: no HTML, no "@" (it would call people), one line, not too long
        static string SafeNote(string note)
        {
            if (note == null) return "";
            StringBuilder sb = new StringBuilder();
            foreach (char ch in note)
            {
                if (ch == '<' || ch == '>') continue;
                if (ch == '@') { sb.Append('＠'); continue; }
                sb.Append(ch < ' ' ? ' ' : ch);
            }
            string s = sb.ToString().Trim();
            return s.Length > 500 ? s.Substring(0, 500) + "…" : s;
        }

        public static string Body(List<WordChange> ch, string note, string edition)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("<div dir=\"rtl\">\n\n");
            sb.Append("**הצעה לשיפור המילון**, נשלחה מתוך התוכנה (גרסה " + Core.AppVersion + (string.IsNullOrEmpty(edition) ? "" : ", מילון: " + edition) + ").\n\n");
            Section(sb, ch, '+', "להוסיף");
            Section(sb, ch, '-', "להסיר");
            Section(sb, ch, '~', "לתקן איות");
            string n = SafeNote(note);
            if (n.Length > 0) sb.Append("**הערה:** " + n + "\n\n");
            sb.Append("</div>\n");
            return sb.ToString();
        }

        // at most 1,000 lines a section: GitHub takes an issue text of up to 65,536 characters
        const int MaxLines = 1000;

        static void Section(StringBuilder sb, List<WordChange> ch, char kind, string head)
        {
            int n = Count(ch, kind);
            if (n == 0) return;
            sb.Append("### " + head + " (" + n + ")\n");
            int written = 0;
            foreach (WordChange c in ch)
            {
                if (c.Kind != kind) continue;
                if (written == MaxLines) { sb.Append("- (ועוד " + (n - MaxLines) + ")\n"); break; }
                sb.Append("- " + (kind == '~' ? c.Word + " ← " + c.NewWord : c.Word) + "\n");
                written++;
            }
            sb.Append("\n");
        }

        public static string Json(string title, string body)
        {
            return "{\"title\":\"" + JsonEscape(title) + "\",\"body\":\"" + JsonEscape(body) + "\"}";
        }

        public static string JsonEscape(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length + 16);
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ') sb.Append("\\u" + ((int)c).ToString("x4", CultureInfo.InvariantCulture));
                        else sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        // The "new issue" page of GitHub, already filled in. An address longer than 2,000 characters may not reach the browser whole
        // (a Hebrew letter takes 6 characters in an address), so a long text is not put in it: cut = true, and the page asks the person
        // to paste the text there (the window puts it on the clipboard first).
        public const int MaxUrl = 2000;
        public const string PasteHere = "<div dir=\"rtl\">\n\nהרשימה ארוכה, ולכן התוכנה העתיקה אותה. מחקו את השורה הזאת והדביקו כאן (Ctrl+V).\n\n</div>\n";

        public static string BrowserUrl(string title, string body)
        {
            bool cut;
            return BrowserUrl(title, body, out cut);
        }

        public static string BrowserUrl(string title, string body, out bool cut)
        {
            string head = "https://github.com/" + Repo + "/issues/new?title=" + Uri.EscapeDataString(title) + "&body=";
            cut = body.Length > MaxUrl;   // (checked first: a very long text is never escaped - .NET Framework refuses more than 32,766 characters)
            if (!cut)
            {
                string url = head + Uri.EscapeDataString(body);
                if (url.Length <= MaxUrl) return url;
                cut = true;
            }
            return head + Uri.EscapeDataString(PasteHere);
        }

        // the key is kept in the program a little hidden (not as plain text); the build writes it the same way (scripts/build_all.sh)
        public static byte[] Encode(string token)
        {
            byte[] b = Encoding.UTF8.GetBytes(token);
            for (int i = 0; i < b.Length; i++) b[i] = (byte)(b[i] ^ Mask[i % Mask.Length] ^ (byte)(i * 31 + 7));
            return b;
        }

        public static string Decode(byte[] data)
        {
            byte[] b = new byte[data.Length];
            for (int i = 0; i < data.Length; i++) b[i] = (byte)(data[i] ^ Mask[i % Mask.Length] ^ (byte)(i * 31 + 7));
            return Encoding.UTF8.GetString(b);
        }

        // null = this program was built without a key (then a suggestion goes through the browser)
        public static string Token()
        {
            try
            {
                using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(KeyResource))
                {
                    if (s == null) return null;
                    using (MemoryStream ms = new MemoryStream())
                    {
                        s.CopyTo(ms);
                        string t = Decode(ms.ToArray()).Trim();
                        return t.Length > 0 ? t : null;
                    }
                }
            }
            catch (Exception) { return null; }
        }

        // sends the suggestion as a new issue. true = sent (its number and address); false = not sent (error says why)
        public static bool Send(string title, string body, out int number, out string address, out string error)
        {
            number = 0; address = null; error = null;
            string token = Token();
            if (token == null) { error = "התוכנה נבנתה בלי מפתח לשליחה"; return false; }
            try
            {
                ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;   // GitHub takes only TLS 1.2 and newer; .NET 4.5 does not turn it on by itself
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create("https://api.github.com/repos/" + Repo + "/issues");
                req.Method = "POST";
                req.UserAgent = "milon-torani/" + Core.AppVersion;
                req.Accept = "application/vnd.github+json";
                req.ContentType = "application/json; charset=utf-8";
                req.Headers["Authorization"] = "Bearer " + token;
                req.Headers["X-GitHub-Api-Version"] = "2022-11-28";
                req.Timeout = 30000;
                req.AllowAutoRedirect = false;   // a page of an internet filter on the way is not GitHub's answer
                byte[] data = new UTF8Encoding(false).GetBytes(Json(title, body));
                req.ContentLength = data.Length;
                using (Stream s = req.GetRequestStream()) s.Write(data, 0, data.Length);
                string text;
                HttpStatusCode code;
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (StreamReader r = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) { code = resp.StatusCode; text = r.ReadToEnd(); }
                Match m = Regex.Match(text, "\"number\"\\s*:\\s*(\\d+)");
                if (m.Success) int.TryParse(m.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out number);
                Match u = Regex.Match(text, "\"html_url\"\\s*:\\s*\"(https://github\\.com/[^\"]+/issues/\\d+)\"");
                if (u.Success) address = u.Groups[1].Value;
                // sent = GitHub created the issue (201) and said its number; any other answer (e.g. a filter's page) is not "sent"
                if (code != HttpStatusCode.Created || number <= 0)
                {
                    error = "תשובה לא צפויה (" + (int)code + ")";
                    Core.Log("Sending the suggestion: unexpected answer " + (int)code + " (" + text.Length + " characters, " + (m.Success ? "with" : "no") + " issue number)");
                    number = 0; address = null;
                    return false;
                }
                Core.Log("Suggestion sent: issue #" + number);
                return true;
            }
            catch (WebException ex)
            {
                HttpWebResponse resp = ex.Response as HttpWebResponse;
                error = resp != null ? "GitHub: " + (int)resp.StatusCode + " " + resp.StatusDescription : Core.Msg(ex);
                Core.Log("Sending the suggestion failed: " + error);
                return false;
            }
            catch (Exception ex)
            {
                error = Core.Msg(ex);
                Core.Log("Sending the suggestion failed: " + error);
                return false;
            }
        }
    }
}
