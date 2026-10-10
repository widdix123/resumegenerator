# Resume Builder

Edit one file (`resume.md`), run one command, and get the PDF.

## Quick start

1. Edit your content in **`resume.md`**.
2. Build:
   - **Double-click `build.command`** in Finder, or
   - in Terminal:
     ```bash
     cd ~/Downloads/resume
     ./build.command
     ```
3. `Abhishek_Gupta_Resume_Designed.pdf` opens when the build finishes.

### Build only the PDF

```bash
cd ~/Downloads/resume
python3 scripts/create_designed_resume.py   # designed PDF (photo + sidebar)
python3 scripts/create_ats_resume.py        # ATS PDF + ATS DOCX
```

## Build on GitHub (automatic)

Every push to `main` that changes `resume.md` (or `scripts/`) runs the **Build resume** pipeline
([.github/workflows/build-resume.yml](.github/workflows/build-resume.yml)). To get the files:

1. Open the repo on GitHub → **Actions** → **Build resume** → the latest run.
2. Under **Artifacts**, download **resume** (a zip with both PDFs and both DOCX files).

You can also start a build by hand: Actions → Build resume → **Run workflow**.

**Photo (optional, one-time):** the photo is not in the repo. To include it in pipeline builds,
store it as a repository secret named `RESUME_PHOTO_B64`:

```bash
base64 -i AbhishekGupta.jpg | gh secret set RESUME_PHOTO_B64 -R widdix123/resumegenerator
```

Without the secret, the pipeline still builds, just with no photo.

## Folder layout

```
resume/
├── README.md                  ← this file
├── resume.md                  ← EDIT THIS: all resume content
├── build.command              ← run this to build everything
├── AbhishekGupta.jpg          ← photo used in the designed resume
├── scripts/                   ← build scripts (no need to edit)
│   ├── resume_content.py        reads resume.md
│   ├── create_designed_resume.py  → Abhishek_Gupta_Resume_Designed.pdf
│   ├── create_ats_resume.py       → Abhishek_Gupta_Resume_ATS.pdf / .docx
│   └── create_designed_docx.py    → Abhishek_Gupta_Resume_Designed.docx
├── preview/                   ← intermediate HTML (open in a browser to preview)
└── assets/                    ← generated round photo for the .docx
```

## Outputs: which file to use

| File | Use it for |
|---|---|
| `Abhishek_Gupta_Resume_Designed.pdf` | Emailing recruiters and sharing directly. **This is the main one.** |
| `Abhishek_Gupta_Resume_ATS.pdf` / `.docx` | Uploading to job portals (Workday, Naukri, LinkedIn Easy Apply…). Single column, no photo, so ATS software parses it correctly. |
| `Abhishek_Gupta_Resume_Designed.docx` | Only when someone insists on an editable Word file. |

## Why the designed DOCX looks different from the PDF

The two files are produced in different ways:

- **The PDF** is a web page (HTML + CSS) printed to PDF by Google Chrome. CSS supports the full design: the sidebar that runs to the bottom of every page, rounded skill chips, the timeline dots, and exact spacing.
- **The DOCX** is rebuilt from scratch in Word's own format, which has none of those features. It mimics the layout with a two-column table and a shaded cell, and it uses only features that survive import into Google Docs. Word, Google Docs and Pages also each render the same `.docx` slightly differently (fonts, line spacing, page breaks).

So the DOCX is a close approximation, not a pixel copy. **When the look matters, send the PDF.** If you need to change wording, edit `resume.md` and rebuild instead of editing the DOCX. That way the PDF and DOCX stay in sync.

## Editing `resume.md`

The format rules are in the comment at the top of `resume.md`. In short:

```markdown
# Your Name
headline: One-line title under your name
tagline: Short italic line

## Contact
- email: you@example.com
- phone:                     ← leave empty to hide it

## Experience
### Job Title | Company | City, Country | Jan 2022 – Present
- Bold lead — the part before " — " is shown in bold in the designed PDF
- Plain bullet
```

Keep the `## Section` names as they are. If you make a typo, the build stops with an error that names the line number in `resume.md`.

## Requirements (one-time setup)

- macOS with **Google Chrome** installed at `/Applications/Google Chrome.app` (it renders the PDFs)
- **Python 3** (the built-in `python3` on macOS works)
- A photo named **`AbhishekGupta.jpg`** in this folder (needed by the designed resume; it is not stored in git)
- Python packages:
  ```bash
  pip3 install python-docx pillow
  ```

## Troubleshooting

- **"Permission denied" when running `./build.command`**: run `chmod +x build.command` once.
- **macOS blocks double-click ("unidentified developer")**: right-click `build.command` → Open → Open.
- **Chrome error / no PDF produced**: check that Chrome is in `/Applications`. The path is set in `CHROME` at the top of `scripts/create_ats_resume.py`.
- **`ModuleNotFoundError: docx` or `PIL`**: run the `pip3 install` command above.
- **Preview the layout without building a PDF**: open `preview/designed_resume.html` in Chrome.
