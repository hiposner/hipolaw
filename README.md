# HipoLaw

Windows desktop workflow for preparing legal motions and generating tables of contents and authorities from approved Excel review workbooks.

**Current version:** V16.4.2 (source build; not a signed or independently certified filing application).

## Three-step workflow
1. **Prepare** a working Word document; optionally align text to a 28-line pleading-paper grid. The original document is not overwritten.
2. **Review** citations and headings using the editable ChatGPT prompt and Excel workbook.
3. **Generate** the combined Word document, PDF/A output, and audit report.

## Build on Windows

Requires Windows, desktop Microsoft Word, Python 3.10+ and internet access to install dependencies.

Double-click `BUILD_V16_4_2.bat`, then launch `dist\HipoLaw_v16_4_2.exe`. See [V16.4.2 instructions](README_WINDOWS_V16_4_2.md).

`HipoLaw_Style_Settings.json` and `HipoLaw_ChatGPT_Prompt.txt` can be edited without rebuilding the executable. Changes to Python logic require a rebuild.

## Testing and filing

Run `python -m unittest discover -s tests` for offline regression tests. Microsoft Word automation, layout, citation correctness, and PDF/A conformance **must be verified independently on Windows** before court filing. HipoLaw's audit is not a legal citator or a substitute for attorney review.

## Privacy

**Do not commit client motions, approved case-specific workbooks, case PDFs, audit logs, temporary document copies, or generated filings.** This repository is public. The published repository contains only reusable application source, build materials, templates, and artwork.

No license has been granted by this README.
