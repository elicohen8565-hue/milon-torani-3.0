# Building and checking Torah Dictionary 3.0

The program is a WinForms installer in C# 5 for .NET Framework 4.5. It is built on Linux with Mono, and it runs on Windows 10/11 with Microsoft Word.

## In one step

```
sudo apt-get install -y mono-devel mono-mcs libmono-cecil-private-cil python3-pil fonts-dejavu-core
bash scripts/build_all.sh
```

The result is `dist/Milon-Torani-3.0.exe`. You only get it if every check passed. GitHub Actions (`.github/workflows/build.yml`) runs the same script on every push.

## What the script does, and the rules behind it

1. **Assets** (`src/make_assets.py`) – it makes:
   - the icon `app.ico`: 9 sizes, written by hand, as 32-bit bitmaps plus a PNG for 256;
   - the exact word count of each edition (`counts.txt`);
   - the compressed word lists that are embedded in the program.
2. **Build** (`src/build.sh`) – builds `TorahDictionary.exe` and the developer tool `DevDump.exe`.
   - **Rule:** the build compiles against the .NET Framework 4.5 **reference** assemblies (`/usr/lib/mono/4.5-api`), never against Mono's own libraries.
   - Mono's libraries have methods that .NET Framework lacks, such as `String.Split(char, StringSplitOptions)`. A program that uses one runs under Mono but dies on Windows with `MissingMethodException`.
   - Afterwards, `CheckApi.exe` (Mono.Cecil) checks every library member the finished program uses. It must print `API CHECK PASSED`.
   - Only C# 5 (`mcs -langversion:5`): no `?.`, no `$""`, no C# 6 initializers.
3. **Self-test:** `mono TorahDictionary.exe --selftest <folder>`.
   - It runs the install, removal and check logic against a temporary folder and a pretend Word.
   - It runs only under Mono. On Windows it refuses, because it writes this user's registry.
4. **Dictionary:** `DevDump.exe` writes the four editions exactly as the installer writes them, and their SHA-256 must equal `tests/dictionary.sha256`.
   - When the word lists in `data/` change on purpose, update that file in the same commit.
   - Also add every word that was dropped from the lists to `Core.Retired` (`src/Core.cs`). Otherwise a dropped word left in an installed file is kept as if the user had added it.

## A new version

1. Raise the version in **both** places:
   - `AppVersion` in `src/Core.cs`;
   - the three `Assembly*Version` attributes at the top of `src/UI.cs`.
2. Write the release notes in `RELEASE_NOTES.md`.
3. Push to `main`.

When the checks pass, `scripts/release.sh` publishes the release `v<version>` with the program. A push that keeps the same version number builds and tests the program, but publishes nothing.

The download name is ASCII on purpose (`Milon-Torani-3.0.exe`), because GitHub may rename release files whose names have other characters. Inside Windows, the program, its window and its file details are named "המילון התורני 3.0".

## Files

| Path | What |
|---|---|
| `src/Core.cs` | All logic: word lists, editions, writing the `.dic` file, Word automation (late-bound COM), registry, install/remove/advanced removal, the check through a hidden Word, log |
| `src/UI.cs` | The windows (drawn by hand, DPI-aware), `Program.Main`, the self-test and the pretend Word |
| `src/DevDump.cs`, `src/CheckApi.cs` | Developer tools |
| `data/basic_list.txt` | Basic edition (Tanakh, Mishnah, Bavli, Yerushalmi) |
| `data/extended_only_list.txt` | What the extended edition adds |
| `data/hebrew_base.txt` | Base words for "chibur torani" (6 Aramaic connecting letters × these words) |

## Hidden switches (developer only)

- `--dpi N` (50..400): pretends a display scale, for screenshots.
- `--preview`: shows the "installed" look.
- `--preview-wordopen [remove]`: shows the "Word is open" window.

To see the window on Linux:

```
xvfb-run -a -s "-screen 0 3200x2200x24" bash -c 'mono TorahDictionary.exe --preview --dpi 125 & sleep 6; import -window root shot.png; kill %1'
```

The log of the program on Windows is `%TEMP%\milon-torani-log.txt`.
