# HipoLaw

Windows desktop workflow for preparing legal motions and generating tables of contents and authorities from approved Excel review workbooks.

**Current version:** V16.4.9 (source build; not a signed or independently certified filing application).

## Three-step workflow
1. **Prepare** a working Word document and one preparation audit (no saved preview PDF); optionally align the motion to a 28-line pleading-paper grid. The original document is never overwritten.
2. **Review** citations and headings using the editable ChatGPT prompt and Excel workbook.
3. **Generate** a final folder with exactly three files: `Motion with Tables.docx`, `Motion with Tables.pdf`, and `Final Audit.txt`. A citation spanning two physical pages is credited to its starting page. The 28-line alignment protects tall inline graphics, TOC list numbers retain their periods, and TOA categories place Other Authorities last. The review workbook can be copied beside the original motion after SHA256 verification.

## Build on Windows

Requires Windows, desktop Microsoft Word, Python 3.10+ and internet access to install dependencies.

Double-click `BUILD_V16_4_9.bat`, then launch `dist\HipoLaw_v16_4_9.exe`. See [V16.4.9 instructions](README_WINDOWS_V16_4_9.md).

`HipoLaw_Style_Settings.json` and `HipoLaw_ChatGPT_Prompt.txt` can be edited without rebuilding the executable. Changes to Python logic require a rebuild.

## Testing and filing

Run `python -m unittest discover -s tests -p 'test_*.py'` for offline regression tests. Microsoft Word automation, layout, citation correctness, and PDF/A conformance **must be verified independently on Windows** before court filing. HipoLaw's audit is not a legal citator or a substitute for attorney review.

## Privacy

**Do not commit client motions, approved case-specific workbooks, case PDFs, audit logs, temporary document copies, or generated filings.** This repository is public. The published repository contains only reusable application source, build materials, templates, and artwork.

No license has been granted by this README.
