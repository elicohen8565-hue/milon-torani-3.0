# Size check (a developer tool)

`SizeCheck.cs` is a one-time measuring tool. It is not part of the program and is not released with it.
It answers two questions on a real Windows computer with Word:

1. How many words does Word accept in **one** custom dictionary?
2. How much slower does spell checking get as the dictionary grows?

## What it does

For each size (0, 175,000, 200,000, 250,000, 300,000, 400,000, 500,000, 750,000, 1,000,000 words), it:

1. Writes a test dictionary `Milon_SizeTest_<size>.dic` to Word's folder (`%APPDATA%\Microsoft\UProof`). The file holds the 172,531 words of the extended edition with "chibur torani", then made-up words (`זקז` + 5 letters) up to the size.
2. In a hidden Word, removes the Torah dictionary and adds the test dictionary instead. If Word refuses the file, it is added to the registry list that Word reads when it starts (the program's own fallback).
3. In a **new** hidden Word, asks about:
   - two control words;
   - 15 sample words from the start of the file;
   - 10 made-up words from the middle and 10 from the very end;
   - one unknown word;
   - a timed block of known words spread over the file, and a timed block of unknown words. A block has 50 words, or fewer when Word is very slow.

The first size that fails stops the run. Three more sizes then narrow the limit down, to less than 10,000 words.

At the end the Torah dictionary is put back as it was, and so is Word's default dictionary. A new Word then checks that it knows the words again (`Verify.Run`). The restore also runs in these cases:
- an error;
- the person stops the check;
- Word is opened during the check (the tool waits until Word is closed).

If the computer goes off in the middle, a mark in `%LOCALAPPDATA%\MilonSizeCheck\restore.txt` makes the next start put everything back first. The button "החזר את המילון" does only that.

While the check runs:
- it holds the program's single-instance mutex, so the Torah dictionary program cannot change Word's list at the same time;
- it keeps the computer awake;
- it stops a hidden Word that gives no answer for 10 minutes.

## Results

The results go to `Milon_SizeCheck_results.txt` on the desktop. The file has:
- `#` lines about Word and its dictionaries before and after the check;
- one `SIZE` line per size, with tab-separated `key=value` fields:
  - `works=1` means Word knew the words of the file from the first to the last;
  - `one_ms`, `in_ms` and `out_ms` are the speed (`in_n` / `out_n` words per block);
- `#LIMIT works=<size> fails=<size>`;
- `#RESTORED ok`.

## Build and self-test

```
bash tools/build_sizecheck.sh
```

The script first runs `scripts/build_all.sh` (the whole program and all its checks). It then compiles the tool from `Core.cs`, `UI.cs`, `Words.cs` and `SizeCheck.cs` against the .NET Framework 4.5 reference assemblies, and runs `CheckApi`. Last, it runs the self-test, `--selftest <folder>` (Mono only), against a pretend Word. The pretend Word keeps its list where the real one does, in the registry list, and can:
- ignore a file over a size;
- use only the start of a file;
- refuse to add a file;
- stop answering.

The result is `dist/Milon-SizeCheck.exe`.
