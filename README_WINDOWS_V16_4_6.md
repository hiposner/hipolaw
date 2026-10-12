# HipoLaw V16.4.6 — PDF citation-pagination audit correction

## What changed

V16.4.5 completed the Word TOA/TOC build and the PDF/A check on the reported motion but stopped during the **final independent PDF citation-page audit**. The reported audit errors were not Word TA-anchor failures:

- A parent citation (e.g. a rule without a child subdivision) could be confused with an immediately longer, narrower citation when reading **TOA entries** from extracted PDF text. V16.4.6 recognizes **whole TOA rows with dot leaders and page numbers** instead of accepting a substring embedded in a child rule.
- When searching the **motion body**, an unsuffixed rule could inadvertently match text belonging only to a subsection or sub-subsection. V16.4.6 applies the same subdivision-boundary check in the PDF audit that the Word marker matcher already uses.
- An abbreviated numeric citation appearing in a verified compound reference (e.g. a following evidence-rule number) could be absent from the body audit because the workbook's **source text** is abbreviated and the audit wrongly demanded the expanded TOA display text. V16.4.6 recognizes the exact contextual compound patterns that the Word marker matcher already accepts; it does **not** search for arbitrary bare numbers throughout the motion.
- A citation split between printed pages could be missed when the PDF audit searched each page separately. V16.4.6 searches a continuous, content-only stream while attributing each match to **the page on which the reference begins**. Numbered-pleading gutter and footer text are excluded to prevent false cross-page matches.

**Safety preserved:** V16.4.6 continues to stop when an approved authority is truly missing from the TOA or the printed pages do not match the cited body pages. It does not fabricate citations, silently adjust page numbers, or disable `ensure_pagination_matches`. The existing PDF/A, unsafe-PDF, source-document identity, TOC, and content-integrity checks remain.

## Installation and retry on Windows

1. Extract the complete V16.4.6 ZIP into its own new folder.
2. Double-click `BUILD_V16_4_6.bat` to build `dist\HipoLaw_v16_4_6.exe` (requires Windows, installed desktop Microsoft Word, and Python 3.10+).
3. Run the EXE and use **Step 3** with the existing **prepared Word working copy** and the **attorney-approved Excel workbook**. Do not redo Step 1 or change the motion solely because of this audit correction.
4. Carefully inspect the generated Word document, PDF, and `HipoLaw_AUDIT.txt`, including table page references, the 28-line grid, TOA/TOC typography and footer numbers, and legal citations.

The `HipoLaw_ChatGPT_Prompt.txt` and `HipoLaw_Style_Settings.json` remain editable beside the EXE and need no rebuild when changed. The prompt is universal for this class of legal motions, not tailored to the reported motion.

## Testing and limitations

Run `python -m unittest discover -s tests -p "test_*.py"` in the extracted source folder. Offline tests include realistic synthetic PDFs with distinct parent/child rule rows, safe compound shorthand, split-page citations, and a case where a truly missing TOA row must **continue** to stop the audit. The parsing is also checked against a previously available Word-generated PDF.

The final Microsoft Word COM workflow cannot be executed or independently validated on Linux. PDF/A declaration is checked but full ISO conformance requires independent veraPDF validation when available. **A PDF is not filing-ready until a successful Windows run plus attorney review.**
