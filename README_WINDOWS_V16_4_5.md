# HipoLaw V16.4.5 — Windows Word PDF/A compatibility

## Fix for the PDF/A identification error

Microsoft Word for Microsoft 365 may write a PDF with an XMP identification of **PDF/A-3a** when exported through its archival PDF option (`UseISO19005_1=True`). Previous HipoLaw versions required PDF/A-1 in the metadata, rejecting validly *identified* PDF/A-2 and PDF/A-3 profiles even before optional independent validation. We verified this exact mismatch using a Microsoft Word-created PDF/A preview from the same prepared motion.

V16.4.5 reads the actual PDF/A profile declared in the PDF's XMP, and supports PDF/A-1a/1b, PDF/A-2a/2b/2u, and PDF/A-3a/3b/3u. It does not alter, insert, fabricate, or repair a PDF/A declaration. Files without a supported declaration, encrypted PDFs, embedded attachments, active PDF content, and other disallowed features remain blocked. If veraPDF is installed, HipoLaw validates against the **detected profile** (e.g., `-f 3a`), stopping if the independent check fails. Without veraPDF, the audit explicitly reports that full ISO conformance has not been independently verified.

The underlying V16.4.4 citation anchoring, native Word TOA and TOC, formatting, 3-step workflow, 28-line pleading paper, external prompt/settings, and hippo artwork are unchanged.

## Installation on Windows

1. Extract the full source ZIP to a new folder; keep earlier working folders as backup.
2. Run `BUILD_V16_4_5.bat` to create `dist\HipoLaw_v16_4_5.exe` (requires Windows, Microsoft Word desktop, Python 3.10+).
3. Launch V16.4.5. Go straight to **Step 3**, selecting the **existing prepared Word document** and the **same approved review workbook**; do not repeat preparation or the attorney's Excel decisions unless the file content changed.
4. Inspect the complete `.docx`, PDF, and `HipoLaw_AUDIT.txt`. Look for `PDF/A declaration in XMP: PDF/A-3a` or the actual profile exported by Word. Note the independent validation status.

## Safety and testing

Run `python -m unittest discover -s tests -p 'test_*.py'`. Offline tests check PDF/A profile recognition, synthetic metadata, regression behavior, and mocked veraPDF invocation. The actual Microsoft Word COM export and final PDF/A ISO compliance **cannot** be certified without running the Windows application and, ideally, independent veraPDF validation. This is a software tool, not an assurance of court acceptance or legal correctness.

Northern District of California CM/ECF currently says PDF/A is not required, although it can reduce malformed-PDF rejections. HipoLaw retains the PDF/A integrity safeguard as a product safety goal.
