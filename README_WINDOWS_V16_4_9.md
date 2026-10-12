# HipoLaw V16.4.9 — Full-size graphic protection and corrected TOC/TOA

V16.4.9 builds on the original V15.1 Word TOC/TOA engine and V16.4.8's start-page citation audit, clean file naming, larger animated hippo, and separate Copy Prompt / Open ChatGPT buttons. It adds four corrections:

1. **Citations spanning pages still use their first page.** The PDF cross-check searches both sides of a physical page break for a case name that begins at the end of page N and continues at the start of N+1. It verifies both case-name fragments, ignores unrelated headers/footers, and credits page N. It does **not** accept an unrelated continuation or simply disable citation/page checks. The native Word TOA markers remain authoritative and independent audit checks remain in place.
2. **Plain, predictable filenames remain, with no unnecessary saved preview or sidecar.** The user-facing preparation step writes just the Word working copy and a preparation audit. The final folder has Word, PDF, and audit. Folder names no longer contain timestamps.

3. **Preserve embedded verdict graphics during 28-line alignment.** Standard paragraphs keep exact 24pt spacing, but a paragraph containing an inline picture or object uses Word's **At least 24pt** line spacing instead. This prevents a roughly 1.1-inch-tall inserted verdict form from being clipped to a 24pt-high line. Every saved DOCX undergoes a graphic-count, dimension, and unsafe fixed-line-height check; the app blocks potentially clipped or resized inline graphics.
4. **Correct table typography and ordering.** Native TOC entries use the actual source list number plus a period (e.g. `I. INTRODUCTION`, `A. The...`, `i. Applicable law`). The Table of Authorities groups appear in this order: **Cases, Statutes, Rules, Other Authorities** (omit empty groups). The **TABLE OF AUTHORITIES** title is explicitly centered, with inherited pleading-template indentation removed.

## Step 1 — Prepare (only TWO files)

Choose the original Word document, confirm the first body heading, and optionally check **Align motion body to 28-line pleading paper**. The original is never overwritten. The app saves a folder named:

`<original motion name> - Prepared`

Inside that folder, there are exactly two files after a successful run:

- **`Prepared Motion.docx`** — the working Word copy for ChatGPT and Step 3
- **`Preparation Audit.txt`** — the preparation log **and** the document identity/fingerprint needed for Step 3

**There is no saved PDF in Step 1.** With grid alignment enabled, Word still generates a temporary PDF for the essential baseline check, but the program automatically deletes it once the check completes. With alignment disabled, no preview PDF is generated. Click **Open prepared Word file** to inspect the DOCX.

Keep the two preparation files together. Step 3 reads the SHA256 fingerprint embedded in the preparation audit and matches it against the approved Excel workbook. Do **not** edit the prepared Word copy after generating the review workbook. Previous prepared documents using a `.hipolaw_prepared.json` sidecar are still supported.

**Important upgrade note:** If an earlier HipoLaw version has already clipped a verdict graphic in a prepared DOCX, that prepared file must NOT be reused. Re-run Step 1 **from the intact original Word document** with V16.4.9. Because the corrected prepared DOCX will have a new SHA-256 identity, the review workbook also needs to be recreated/updated against that new prepared copy before Step 3. Do not bypass the identity checks. Previously prepared documents WITHOUT cropped inline graphics can still be used when they pass the safety checks.

If the preparation folder already exists, HipoLaw asks before replacing its two files. It **refuses to overwrite unrelated files** in that folder. Repreparing a document may change its fingerprint, so a previously approved workbook may no longer match.

## Step 2 — ChatGPT review

Use **1. Copy Prompt**, then **2. Open ChatGPT**, and paste with **Ctrl+V**. Upload `Prepared Motion.docx`, download the generated review workbook, and approve the decisions. A browser's download location cannot generally be set by HipoLaw. In Step 3, select the workbook from Downloads. HipoLaw can copy it beside the original motion after verifying the document fingerprint.

## Step 3 — Generate (only THREE final files)

Select a prepared document and matching approved workbook. You can use the same inputs as an earlier version only if the graphic-safety check passes. **For the clipped-verdict-document example, reprepare from the original in V16.4.9 and generate a matching review workbook.** The motion text and original graphic content remain unchanged.

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
2. Run **`BUILD_V16_4_9.bat`**.
3. Open **`dist\HipoLaw_v16_4_9.exe`**.
4. Leave **`HipoLaw_ChatGPT_Prompt.txt`** and **`HipoLaw_Style_Settings.json`** beside the EXE. Both can be edited without rebuilding.

Offline tests:

`python -m unittest discover -s tests -p "test_*.py"`

**Layout trade-off:** Tall inline figures may occupy multiple pleading-number rows; HipoLaw prioritizes showing the *entire* exhibit over forcing all graphics into a single 24pt row. Check the line numbers and page transitions near inserted images. HipoLaw does not auto-modify floating images, complex tables, or footnotes to force them onto the grid.

**Limitations:** A full Microsoft Word COM run, accurate pleadings layout, native TOC/TOA pagination, and complete PDF/A conformance cannot be verified from a non-Windows environment. Review the Word and PDF carefully before filing. The audit does not substitute for legal citation verification or independent PDF/A conformance testing (e.g., veraPDF).

**Privacy:** This public-source package includes application code, generic templates, assets, and tests only. Do not commit client motions, client-specific workbooks, generated court filings, or legal audit logs to a public repository.
