# Advisee Document Filler

A small tool for UNP CCIT advisers, available as a standalone desktop app
or a local web app. It reads Periodic Grades Listing PDFs from the student
portal and fills two official forms for every advisee.

- Appraisal Sheet (VPAA-CCIT-QF-05), one per student, all terms in one file,
  using the template that matches the student's course (BSCS, BSIT, BLIS
  or DCT), auto-detected from the PDF
- Report of Rating (VPAA-CCIT-QF-09), one per student per term

## Two ways to run it

### A. Standalone desktop app (recommended, no browser)

A step-by-step wizard. No web server, no network, nothing exposed.

From source:

```
pip install -r requirements.txt
python gui_app.py
```

On Linux you also need Tkinter once: `sudo apt install python3-tk`

As a single-file executable (no Python needed on the target machine):

- Windows: run `build_windows.bat`, then use `dist\AdviseeDocFiller.exe`
- Linux: run `./build_linux.sh`, then use `dist/AdviseeDocFiller`

Build on the OS you target. PyInstaller does not cross-compile, so the
Windows exe must be built on Windows and the Linux binary on Linux. A
Linux binary runs on distros with the same or newer glibc than the
build machine.

Generated documents default to an `AdviseeDocuments` folder inside the
user's real Documents folder (`Documents\AdviseeDocuments` on Windows,
resolved even if OneDrive or a domain policy has redirected it;
`~/Documents/AdviseeDocuments` on Linux, following the XDG user-dirs
setting). Use "Browse..." on the Generate step to save elsewhere.

### B. Web app (optional)

```
pip install -r requirements.txt
python app.py
```

It binds to 127.0.0.1 only (loopback, not visible to your network) and
picks a free port in the dynamic private range (49152-65535). The exact
address is printed on start.

## Build with GitHub (no local build needed)

Push this project to a GitHub repository. The included workflow at
`.github/workflows/build.yml` builds both executables on GitHub's own
Windows and Linux machines, so you never build on your own computer.

- Manual build: open the repo's Actions tab, pick "Build executables",
  click "Run workflow", wait a few minutes. This publishes/updates a
  rolling "latest" pre-release on the repo's Releases page with every
  build attached, so it's shareable without anyone needing to sign in
  (unlike raw workflow-run artifacts, which are private and expire).
- Release build: create and push a tag like `v1.0`
  (`git tag v1.0 && git push origin v1.0`). GitHub builds both binaries
  and attaches them to a proper versioned Release page instead of the
  rolling "latest" one.

## Verifying a release (Sigstore)

Every release includes `SHA256SUMS`, `SHA256SUMS.sig`, and
`SHA256SUMS.pem`. These prove the published files were built by this
repo's `build.yml` on GitHub's runners and haven't been altered since,
using [Sigstore](https://www.sigstore.dev/) keyless signing (no private
key to leak, verifiable against the public Rekor transparency log).
This is a supply-chain integrity check, not a Windows/Authenticode
signature - it won't change the SmartScreen or "Unknown Publisher"
prompt, which come from a separate, CA-based trust system.

To verify with [cosign](https://docs.sigstore.dev/cosign/system_config/installation/):

```
cosign verify-blob \
  --certificate SHA256SUMS.pem \
  --signature SHA256SUMS.sig \
  --certificate-identity-regexp "^https://github.com/Yaksharo/appraisal-filler/" \
  --certificate-oidc-issuer https://token.actions.githubusercontent.com \
  SHA256SUMS
```

Then confirm your downloaded file's hash appears in `SHA256SUMS`
(`sha256sum -c SHA256SUMS` on Linux, `certutil -hashfile <file> SHA256`
on Windows).

## Startup speed

Builds use PyInstaller's onedir mode: the app is a folder with the
executable inside, next to an `_internal` folder. It starts in about a
second. The old single-file mode had to unpack everything to a temp
directory on every launch, which is why it felt slow. Heavy libraries
also load lazily now, so the window appears immediately and the PDF
engine loads in the background the first time you click Next.

- Windows, two options from the same build:
  1. Installer: run `AdviseeDocFiller-Setup-1.0.exe`. It installs to
     `C:\Program Files\Yaksharo Solutions\Advisee Document Filler`,
     adds a Start Menu entry (and optional desktop icon), and registers
     a normal uninstaller in Windows Settings.
  2. Portable: download `AdviseeDocFiller-windows-portable.zip`, extract
     it once anywhere (USB stick included), run `AdviseeDocFiller.exe`
     inside. Copy-paste the folder to move it. Keep the exe next to its
     `_internal` folder.
- Linux portable: `AdviseeDocFiller-linux.tar.gz`, extract and run.
- Linux packages install the folder to /opt and handle everything.

## Linux packages (deb, rpm, Arch)

The GitHub workflow also builds proper installers for the three big
Linux families, uploaded as the `linux-packages` artifact and attached
to Releases:

- Debian / Ubuntu / Mint: `sudo apt install ./adviseedocfiller_*.deb`
- Fedora / RHEL / openSUSE: `sudo dnf install ./adviseedocfiller-*.rpm`
- Arch / Manjaro: `sudo pacman -U adviseedocfiller-*.pkg.tar.zst`

Each package installs the app to /opt/AdviseeDocFiller, adds an
`adviseedocfiller` command, and registers a menu entry with the app
icon. Note: the binary inside is built on GitHub's Ubuntu runner, so it
needs a distro with the same or newer glibc. Arch and current Fedora
are fine; very old LTS releases may not be.

## Window style

- Linux: the app keeps your desktop environment's native titlebar, so
  it automatically matches GNOME, KDE, XFCE, or whatever you run.
- Windows: the app draws its own borderless titlebar that follows the
  light/dark theme, with minimize, maximize/restore, and close buttons,
  drag-to-move, double-click to maximize, and a resize grip at the
  bottom-right.

## Themes and accessibility

- Desktop app: a flat, neutral white/grey/black look with a single blue
  accent, styled to blend in with the default light/dark themes of
  Windows 10/11, GNOME, and KDE. Follows your system light/dark theme on
  startup, with a toggle in the header. A- and A+ buttons resize all text.
- Web app: follows the system theme, with a toggle that remembers your
  choice.

## How to use

Both apps walk through the same steps; the desktop wizard breaks them
into pages (PDF Files &rarr; Students &rarr; Documents &rarr; Faculty &rarr;
Generate), and the web app does the same work on a single page with a
"Preview students" button.

1. Enter the Adviser name and Dean name (both required) and add one or
   more grade listing PDFs. You can add the 1st and 2nd term listings
   together. Students are matched across PDFs by ID number, and the
   course found in each PDF (BSCS, BSIT, BLIS or DCT) picks the matching
   Appraisal Sheet template automatically.
2. Preview/check the parsed students, and untick anyone you want to skip.
3. Pick which documents to generate.
4. Adviser and Dean are already filled in from step 1. Optionally map
   instructors per subject code (desktop: a dropdown/entry per code,
   pre-filled from names you've used before; web: one per line, like
   `CT103 = Juan Dela Cruz`). The PDF has no instructor names, so the
   Faculty and Instructor columns stay blank unless you map them — leaving
   a faculty field empty also leaves it blank in the generated documents,
   it does not fall back to a placeholder name.
5. Pick the output format. Individual gives one docx per student
   (zipped on the web app). One batch file merges every student into a
   single docx, one student per page, ready for bulk printing.
6. For Reports of Rating you can also untick terms you don't need.
7. Click Generate.

## Remembering names (local database)

The app keeps a small local SQLite database (`store.db`, in the OS's
standard per-user app-data folder — never bundled, never synced) of every
Adviser, Dean and per-subject-code faculty name you've typed. Next time,
those show up as suggestions: a dropdown you can pick from on the desktop
app, or autocomplete on the web app, and the faculty step pre-fills any
subject code it recognizes from a previous run. You can still type a new
name at any time — nothing is required to already be in the list. Nothing
in this database ever leaves the machine.

## How the data maps

- The periodic rating column in the PDF goes to the Midterm cell of the
  Report of Rating. The Grade column goes to the Final cell.
- The course code found in the PDF (`BSCS`, `BSIT`, `BLIS` or `DCT`) picks
  the Appraisal Sheet template. An unrecognized code falls back to the
  BSCS template.
- Each Appraisal Sheet template has one term table per year/term the
  course actually has (DCT: First and Second Year only; the others also
  have Third Year, Mid Year and Fourth Year). Tables are filled in
  chronological order — earliest school year first, 1st Term before 2nd —
  based on the Period line in each PDF, not on any year label in the PDF
  itself.
- Whole year blocks (heading and table) that come after the student's
  last parsed term are removed from the generated file, e.g. a student
  who only reached Second Year won't have Third Year, Mid Year or Fourth
  Year sections at all.
- Term/SY is written as `1st/25-26` style. Change `term_sy_label` in
  `filler.py` if your format differs.
- Total units per term is computed and placed in the totals row.

## Notes

- The templates live in `templates_docx/`. Replace them with updated forms
  as long as the table layout stays the same; add a new course by dropping
  in a docx and adding it to `APPRAISAL_TEMPLATES` in `filler.py`.
- Unused blank rows are removed automatically in both documents: on the
  Report of Rating you can untick the option in the UI to keep them, the
  Appraisal Sheet always trims them (that's what the extra blank rows in
  the BSCS/BSIT/BLIS templates are for).
- Trailing empty paragraphs in the templates are stripped, so the Report
  of Rating no longer produces a blank second page.
- The parser targets the exact layout of the "Periodic Grades Listing"
  report. If the registrar changes the report layout, adjust `parser.py`.
- The About section (desktop) and footer (web) show the version of the
  build you're running, resolved from `VERSION` (bundled by the build
  workflow from the git tag) and falling back to `git describe` when
  running from source. See `version.py`.

## App icon

The logo in `assets/` is used everywhere automatically:

- Windows: the exe file icon (embedded at build time) and the window
  and taskbar icon at runtime.
- Linux: the window and taskbar icon at runtime. Linux binaries cannot
  embed a file icon, so for a launcher/menu icon copy the binary and
  `assets/logo.png` somewhere permanent (e.g. `/opt/AdviseeDocFiller/`),
  adjust the paths inside `AdviseeDocFiller.desktop`, and copy that file
  to `~/.local/share/applications/`.
- Web app: the same logo is served as the favicon.

## Credits

Developed by Yaksharo a.k.a Ezer
