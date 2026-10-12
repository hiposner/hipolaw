# HipoLaw V16.4.7 — Clean workflow and larger hippo

V16.4.7 keeps the V15.1 Microsoft Word TOC/TOA engine and all V16.4.6 legal-citation, pagination and PDF/A checks. This release improves **where your files go** and **how Step 2 works**, without changing attorney approvals or source citations.

## Step 1 — Prepare your Word document

1. Select your original Word motion and the exact first body heading.
2. Choose whether the entire body should align with numbered pleading paper.
3. Click **Prepare Word working copy**. HipoLaw leaves your original unchanged and creates a distinct `HipoLaw_Prepared_...` working folder.

The prepared folder is **working material, not final filing output**. It contains the prepared DOCX, a matching `.hipolaw_prepared.json` fingerprint, a preparation audit, and optionally the preview PDF. Keep that fingerprint beside its prepared DOCX: Step 3 checks it against the approved Excel workbook. Do not file the preview PDF.

## Step 2 — ChatGPT review

There are now **two separate buttons**:

- **1. Copy Prompt** puts the complete, document-bound universal prompt on your clipboard. It does not open the browser.
- **2. Open ChatGPT** opens `https://chatgpt.com/` only. It does not copy or paste any text. Press **Ctrl+V** inside the chat, then upload the prepared DOCX.

A desktop app cannot make a browser save a ChatGPT-created spreadsheet directly into an arbitrary computer folder. Most browsers download to **Downloads**. After you review and approve the workbook, proceed to Step 3 and use **Browse** to choose that XLSX.

## Step 3 — Generate clean output

1. Choose the prepared Word DOCX. You can select it from its preparation folder.
2. Click **Browse** for the approved review XLSX (including from Downloads). HipoLaw checks the SHA256/document-name binding first, then **copies** the workbook next to the **original Word motion**, when that original can be found and its SHA256 verified. Otherwise, it copies next to the prepared Word document. The download stays unchanged. If copying fails (for example due to a read-only folder), you can choose to use the downloaded workbook in place. No extra copy is made when the XLSX is already in the preferred folder.
3. If a different workbook is already there, HipoLaw asks before replacing the review copy. Your original motion and the selected downloaded XLSX are not overwritten.
4. Generate. By default, HipoLaw makes one stable folder called `<motion name>_HipoLaw_Final` under **Save output in folder**. If an earlier HipoLaw-generated output folder has the same name, HipoLaw asks before reusing it. A folder containing unrelated files is never overwritten.

**On a completed run, the final folder contains exactly three deliverables:**

- `<motion name>_With_Tables.docx` — editable Word document
- `<motion name>_With_Tables.pdf` — matching PDF/A output (profile identified from Word XMP)
- `HipoLaw_AUDIT.txt` — one consolidated audit, including citation coverage

Temporary citation-audit intermediates are removed even if generation stops. If Word or an audit fails, the audit explains why the result is **not filing-ready**. The earlier safe original Word motion is always preserved.

## Larger animated hippo

The hippo in the lavender top banner is now displayed at approximately 250×175 pixels rather than about 149×104. Both animation frames were resized for a large crisp display; no new runtime image library is necessary.

## Build and test

Windows with installed desktop Microsoft Word and Python 3.10+ required:

1. Extract the ZIP into a **new folder**.
2. Double-click `BUILD_V16_4_7.bat`.
3. Launch `dist\HipoLaw_v16_4_7.exe`.
4. Put `HipoLaw_ChatGPT_Prompt.txt` and `HipoLaw_Style_Settings.json` beside the EXE (the build script copies them if missing).

Use `python -m unittest discover -s tests -p "test_*.py"` to run offline checks. Microsoft Word COM and genuine DOCX/PDF pagination cannot be executed on Linux. A successful Windows run, manual document review, and independent PDF/A validation where available are still required prior to filing.

**Privacy:** Do not upload client motions, approved client-specific workbooks, generated filing PDFs, or audits to the public GitHub repository. Only the reusable application source/templates/assets/tests belong there.
