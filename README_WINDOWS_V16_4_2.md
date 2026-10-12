# HipoLaw V16.4.2 — Rule 50 anchor-collision fix

This is a **source-build Windows package**, based on HipoLaw V16.4.1. It retains the V15.1 Word engine, hippo artwork, 3-step workflow, Word + PDF/A export, and editable formatting and ChatGPT prompt files.

## Why V16.4.1 stopped

The Word citation finder treated the text `Rule 50` at the start of `Rule 50(a)` and `Rule 50(b)` as an independent occurrence of the unsuffixed `Rule 50` authority. These overlapping proposed TA fields triggered a correct safety stop. Nothing was inserted.

V16.4.2 reads the characters immediately *after* an unsuffixed rule match in Word, excluding it if it is just the prefix of a longer rule number or a subsection citation. Real standalone `Rule 50` occurrences are still marked. The existing hard stop for genuinely conflicting or unmatched authorities remains.

The cross-platform preflight now applies the same rule-number boundary test. The ChatGPT review prompt has been updated to warn against generating ambiguous generic aliases. Existing prepared Word documents and workbooks can be reused without re-approving them **provided their source-document SHA256 remains unchanged**.

## Install

1. Extract the ZIP to a **new** folder.
2. On Windows with Word desktop and Python 3.10+ installed, double-click `BUILD_V16_4_2.bat`.
3. After the successful build, run `dist\HipoLaw_v16_4_2.exe` or `RUN_V16_4_2.bat`.
4. In Step 3, select your **existing prepared Word document** and **existing approved Excel workbook**. You do **not** need to prepare or approve the unchanged document again just for this engine fix.
5. Click Preflight, then Generate.

`dist\HipoLaw_Style_Settings.json` and `dist\HipoLaw_ChatGPT_Prompt.txt` remain editable with Notepad without rebuilding. The `.exe` itself must be rebuilt once for this citation-matching logic fix.

## What does and does not change

- The original motion, prepared document, TOC/TOA formatting preferences, and court-filing output formats are unchanged.
- Only the handling of **unqualified numeric rule anchors** changes.
- Generic `Rule 50` is **not** automatically mapped to `(a)` or `(b)`.
- Specific citations `Rule 50(a)` / `Rule 50(b)` get their own markers.
- Truly ambiguous overlaps, nonexistent approved citations, and unsafe aliases still stop the run **before inserting any TA fields**.
- The audit notes how many misleading generic-prefix matches were skipped.

## Testing and limitations

52 offline regression tests passed, including a simulated Word story containing `Rule 50(a)`, `Rule 50(b)`, a standalone `Rule 50`, and `Rule 500`. Integration testing using desktop Microsoft Word is **still required**; Python tests cannot guarantee court filing readiness.

If it still stops, send `HipoLaw_AUDIT.txt` and the review workbook. In particular, if the workbook maps one ambiguous alias to multiple authorities, that may require attorney review rather than software guessing.
