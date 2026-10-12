# HipoLaw V16.4 — Guided Prepare → ChatGPT → Generate (Windows)

Built on the **V15.1 Microsoft Word** TOC/TOA generator with the existing hippo icon and artwork, the V16 fixes, editable styles, original motion preservation, PDF/A-1 export and audit. **Windows Microsoft Word desktop is required** for Step 1 and Step 3. The ZIP is a **buildable source package**, not a prebuilt Windows EXE.

## Build once

1. Extract into a completely new folder.
2. Install Python 3.10 or newer and desktop Microsoft Word for Windows.
3. Double-click **`BUILD_V16_4.bat`**. This installs PyInstaller, pywin32, openpyxl, and PyMuPDF and generates **`dist\HipoLaw_v16_4.exe`**.
4. Run the EXE or `RUN_V16_4.bat`.
5. Keep **`dist\HipoLaw_Style_Settings.json`** and **`dist\HipoLaw_ChatGPT_Prompt.txt`** next to the EXE. Both are editable, and changing them NEVER requires a rebuild. Rebuilding preserves the existing edits.

## Step 1 — Prepare motion

Select the ORIGINAL Word .docx; confirm the exact first motion-body heading (the app suggests INTRODUCTION only if unambiguous); optionally check **Align motion body to 28-line pleading paper**; then click **Prepare Word working copy**. The original file is never overwritten. The app creates a separate `HipoLaw_Prepared_YYYYMMDD_HHMMSS` folder with:

- `<original>_HipoLaw_Prepared.docx` (the working copy for Steps 2 and 3)
- `<original>_HipoLaw_Prepared_PREVIEW_NOT_FOR_FILING.pdf` if Word's PDF/A export succeeds (preview only; DO NOT FILE)
- `<prepared>.docx.hipolaw_prepared.json` (SHA256 identity, first heading and Yes/No setting)
- `HipoLaw_PREPARATION_AUDIT.txt` (what was adjusted)

The preview PDF is **not** a completed filing. In Yes mode, main-body paragraphs are set to exact 24-point spacing on the prepared copy, but the final numerical pleading-line offset is calibrated when Step 3 adds separate caption/table/body sections. Footnotes, images and complex tables require visual review. No source content should change.

You may also **skip Step 1** and choose a legacy original motion on Step 3, for backward compatibility. Old V15 approved workbooks without fingerprint rows continue working with ordinary, unprepared Word files.

## Step 2 — ChatGPT spreadsheet

Click **Copy prompt + Open ChatGPT**. The app copies the comprehensive, reusable V15-schema spreadsheet prompt to the clipboard and opens `https://chatgpt.com/` in your browser. It **cannot upload a local Word file or send the prompt automatically**; paste the prompt and attach the prepared .docx yourself.

The prompt instructs ChatGPT to generate ONE review `.xlsx` with `Authorities`, `Contents`, `Setup`, and `Aliases`, preserving the exact seven-column approval schema from V15. It also includes a cryptographic SHA256 fingerprint of the Word file. Review all suggested citations and headings and set **Accept**, **Edit**, or **Reject** in the spreadsheet; blank decisions will block Step 3.

The included `05_Review_Template_V16_4.xlsx` contains the new Setup identity rows:

- `HipoLaw document SHA256`
- `HipoLaw document filename`
- `Align entire motion to 28-line pleading grid` (Yes/No)

If the input is a HipoLaw-prepared document, the review workbook must include a matching 64-hex-character SHA256 in Setup. This intentionally blocks mixing an approved spreadsheet with a different prepared version.

## Step 3 — Generate

The prepared Word file is automatically selected. Choose the APPROVED spreadsheet, confirm output folder, click **Preflight**, then **Generate filing**. The original V15-based Word engine builds native Word TC and TA fields, updates page numbers after Word pagination, preserves motion paragraphs and headings, and exports to PDF/A-1. The successful final output folder contains exactly:

1. One editable combined `.docx`
2. One matching `.pdf` (Word PDF/A-1 setting)
3. `HipoLaw_AUDIT.txt` (includes page-reference checks, formatting warnings, local and UTC timestamps and filing checks)

No additional PDF fragments should be left in the final folder. Independent PDF/A compliance requires installing veraPDF; if it is absent, the audit will clearly say the independent validation was NOT RUN. Do not assume a PDF/A declaration alone guarantees court acceptance. Review the Word/PDF visually before filing.

## Change styles or the prompt without rebuilding

Use **Edit styles or spreadsheet prompt (no rebuild)**, which opens BOTH editable files in Notepad. `HipoLaw_Style_Settings.json` controls TOC/TOA title boldness, alignment, category headings, sizes, and pleading grid offsets. `HipoLaw_ChatGPT_Prompt.txt` controls the instructions copied into ChatGPT. Save changes; the next run uses them. Only app logic changes require rebuilding.

## Compatibility and safeguards

- The original .docx is NEVER overwritten by Step 1 or Step 3.
- Prepared-motion SHA256 protects against mixing source versions and approved spreadsheets.
- Step 3 does **not** apply 24-point grid paragraph formatting again if Step 1 has already done it; final header calibration is still applied as needed.
- V15 Word-based editable TOC/TOA fields, hippo artwork, native pleading paper and caption fixes remain.
- Legacy V15 workbooks and direct source documents remain supported when there is no prepared manifest or fingerprint.
- The body and footnote text are compared against the selected input copy. Citation/heading correctness is still subject to human legal review.

If Word generation fails, send the error log and the output PDF/DOCX if created. If the build fails, send the last 30 lines from `BUILD_V16_4.bat`.
