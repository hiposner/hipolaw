# HipoLaw V16.4.8 — Citation start page + simple files

V16.4.8 builds on the original V15.1 Word TOC/TOA engine and V16.4.7's larger animated hippo and separate Copy Prompt / Open ChatGPT buttons. It makes two changes:

1. **Citations spanning pages use their first page.** The PDF cross-check searches both sides of a physical page break for a case name that begins at the end of page N and continues at the start of N+1. It verifies both case-name fragments, ignores unrelated headers/footers, and credits page N. It does **not** accept an unrelated continuation or simply disable citation/page checks. The native Word TOA markers remain authoritative and independent audit checks remain in place.
2. **Plain, predictable filenames, with no unnecessary saved preview or sidecar.** The user-facing preparation step writes just the Word working copy and a preparation audit. The final folder has Word, PDF, and audit. Folder names no longer contain timestamps.

## Step 1 — Prepare (only TWO files)

Choose the original Word document, confirm the first body heading, and optionally check **Align motion body to 28-line pleading paper**. The original is never overwritten. The app saves a folder named:

`<original motion name> - Prepared`

Inside that folder, there are exactly two files after a successful run:

- **`Prepared Motion.docx`** — the working Word copy for ChatGPT and Step 3
- **`Preparation Audit.txt`** — the preparation log **and** the document identity/fingerprint needed for Step 3

**There is no saved PDF in Step 1.** With grid alignment enabled, Word still generates a temporary PDF for the essential baseline check, but the program automatically deletes it once the check completes. With alignment disabled, no preview PDF is generated. Click **Open prepared Word file** to inspect the DOCX.

Keep the two preparation files together. Step 3 reads the SHA256 fingerprint embedded in the preparation audit and matches it against the approved Excel workbook. Do **not** edit the prepared Word copy after generating the review workbook. Previous prepared documents using a `.hipolaw_prepared.json` sidecar are still supported.

If the preparation folder already exists, HipoLaw asks before replacing its two files. It **refuses to overwrite unrelated files** in that folder. Repreparing a document may change its fingerprint, so a previously approved workbook may no longer match.

## Step 2 — ChatGPT review

Use **1. Copy Prompt**, then **2. Open ChatGPT**, and paste with **Ctrl+V**. Upload `Prepared Motion.docx`, download the generated review workbook, and approve the decisions. A browser's download location cannot generally be set by HipoLaw. In Step 3, select the workbook from Downloads. HipoLaw can copy it beside the original motion after verifying the document fingerprint.

## Step 3 — Generate (only THREE final files)

Select the existing prepared document and approved workbook. You can use *the same documents already prepared in an earlier version*; you do not need to redo Steps 1–2 merely to fix the pagination audit.

The default folder is named:

`<original motion name> - Final`

The completed folder contains exactly:

- **`Motion with Tables.docx`** — editable Word document, with the TOC and TOA
- **`Motion with Tables.pdf`** — corresponding Word-exported PDF/A, pending independent validation if required
- **`Final Audit.txt`** — consolidated legal-citation/page-reference and PDF safety audit

The audit includes citation coverage and diagnostic information. Temporary citations reports are deleted on success and failure. The program asks before reusing a final folder and will not overwrite unrelated files. **If any final audit fails, the generated document is NOT filing-ready**, even if the Word/PDF already exist for inspection.

### Citation pagination

When a case citation begins on printed page **7** and continues on **8**, the TOA references **7**. The audit can now recognize a genuine title split across two pages even when footer text or numbered pleading headers interrupts extracted PDF text. It still blocks false matches, missing authorities, incorrect page references, and ambiguous anchors.

### Build on Windows

Requires Windows 10/11, desktop Microsoft Word, Python 3.10+, and access to install build dependencies.

1. Extract the ZIP into a new folder.
2. Run **`BUILD_V16_4_8.bat`**.
3. Open **`dist\HipoLaw_v16_4_8.exe`**.
4. Leave **`HipoLaw_ChatGPT_Prompt.txt`** and **`HipoLaw_Style_Settings.json`** beside the EXE. Both can be edited without rebuilding.

Offline tests:

`python -m unittest discover -s tests -p "test_*.py"`

**Limitations:** A full Microsoft Word COM run, accurate pleadings layout, native TOC/TOA pagination, and complete PDF/A conformance cannot be verified from a non-Windows environment. Review the Word and PDF carefully before filing. The audit does not substitute for legal citation verification or independent PDF/A conformance testing (e.g., veraPDF).

**Privacy:** This public-source package includes application code, generic templates, assets, and tests only. Do not commit client motions, client-specific workbooks, generated court filings, or legal audit logs to a public repository.
