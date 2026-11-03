// Developer tool (not part of the product).
//   mono DevDump.exe <folder>            writes the four editions exactly as the installer would write them,
//                                        so two builds can be compared byte by byte (sha256sum of the four files)
//   mono DevDump.exe <folder> --bench    also prints how long the main steps take
// Built by src/build.sh (see BUILD.md).
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace OtzariaDict
{
    static class DevDump
    {
        static int Main(string[] a)
        {
            string dir = a.Length > 0 ? a[0] : ".";
            bool bench = a.Length > 1 && a[1] == "--bench";
            Directory.CreateDirectory(dir);
            bool[] ext = { false, false, true, true };
            bool[] comp = { false, true, false, true };
            string[] names = { "basic", "basic_comp", "extended", "extended_comp" };
            for (int i = 0; i < 4; i++)
            {
                Stopwatch sw = Stopwatch.StartNew();
                List<string> w = Core.BuildWords(ext[i], comp[i]);
                long built = sw.ElapsedMilliseconds;
                sw.Restart();
                Core.WriteDic(Path.Combine(dir, names[i] + ".dic"), w);
                Console.WriteLine(names[i] + " " + w.Count + (bench ? "   build " + built + " ms, write " + sw.ElapsedMilliseconds + " ms" : ""));
            }
            if (bench)
            {
                Stopwatch sw = Stopwatch.StartNew();
                Core.CountWords(Core.BasicRes); Core.CountWords(Core.MoreRes); Core.CountWords(Core.BaseRes);
                Console.WriteLine("what the window used to do at start-up (count the three word lists): " + sw.ElapsedMilliseconds + " ms");
                sw.Restart();
                int sum = 0;
                for (int i = 0; i < 4; i++) sum += Core.EditionCount(ext[i], comp[i]);
                Console.WriteLine("what it does now (read the counts file; sum " + sum + "): " + sw.ElapsedMilliseconds + " ms");
            }
            return 0;
        }
    }
}
