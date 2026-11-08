# WordCheck 4 (a developer tool)

WordCheck 4 asks Word's own spelling checker about forms from the books that the Torah dictionary does not have yet. Word marks some of them as errors: those are the candidates to add to the dictionary. Word does not mark the rest, because it knows them already.

It is not part of the program and is not released with it.

## What it asks about

`pool4.txt.gz` holds 183,479 forms, made by `make_pool4.py` from the stage-1 scan (`candidates.tsv`).

**The forms:**
- **Taken:** forms that were found in the scanned libraries and that no earlier Word check covered.
- **Left out:** "chibur torani" forms, because that option of the dictionary covers them.
- **Order:** most valuable first, so a run stopped in the middle still gives the most useful part. The value of a form is how many red marks it would remove: its occurrences in all the libraries × the chance that Word marks it. Word marked 82% of the forms Hspell does not know, and 2.5% of the forms it knows.

**The best 180,000** are followed by the 3,479 untested forms of the four books (Tanakh, Mishna, Bavli, Yerushalmi) that were not among them. These forms decide the basic edition.

**What the list also holds:**
- The 13 control words and the speed-test words of WordCheck 3.
- No form that is in any of the four editions (`make_pool4.py` checks this).

`pool4_index.tsv.gz` has a line per form: its rank, sources, occurrences, files, whether Hspell knows it, and whether it is from the four books.

## The program

`WordCheck.cs` is WordCheck 3 (2-11-2026), which ran on the user's Word in 37 minutes and connected the dictionary back. Only these things changed:
- the word list;
- the names: `Word_check_results_4.txt`, `#WORDCHECK 4`, and the single-instance mutex;
- the texts of its window;
- the self-test's check of the embedded list: the control words must *not* be among the forms.

For the time of the check, the Torah dictionary is disconnected from Word. This is not because the forms are in it (they are not). It is because Word checks about ten times faster without it. The dictionary is connected back at the end. A run that stops in the middle connects it back at the next start, and then continues.

## Results

The results go to `Word_check_results_4.txt` on the desktop, in the format of WordCheck 3:
- `G<TAB>form` is a form Word marks;
- `Q<TAB>n` means every form before place n was answered;
- `DONE` ends the run;
- `#RESTORED ok` means the dictionary was connected back.

A form that was answered and has no `G` line is known to Word. If it also has an `H` line, it is in the user's own dictionary.

## Build

```
bash tools/wordcheck4/build.sh
```

The script builds the tool against the .NET Framework 4.5 reference assemblies, runs `CheckApi` (API check) and the self-test, and writes `dist/WordCheck4.exe`.
