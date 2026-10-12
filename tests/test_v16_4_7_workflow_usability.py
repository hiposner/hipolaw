"""V16.4.7 regressions: repeatable output, safe XLSX filing, separated chat UI."""
from __future__ import annotations
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch, MagicMock
from types import SimpleNamespace

SRC = Path(__file__).resolve().parents[1] / 'generate_legal_tables_v16_4_9.py'
spec = importlib.util.spec_from_file_location('hipolaw_ui_1647', SRC)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)

class TestFileOrganization(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.orig = self.root / 'Sample Opposition.docx'
        self.orig.write_bytes(b'original Word bytes')
        self.prep_dir = self.root / 'HipoLaw_Prepared_20261011_123456'
        self.prep_dir.mkdir()
        self.prepared = self.prep_dir / 'Sample Opposition_HipoLaw_Prepared.docx'
        self.prepared.write_bytes(b'prepared Word bytes')
        manifest = {
            'schema_version': 1,
            'prepared_filename': self.prepared.name,
            'prepared_sha256': mod.sha256_file(self.prepared),
            'original_filename': self.orig.name,
            'original_sha256': mod.sha256_file(self.orig),
            'first_body_heading': 'INTRODUCTION',
            'align_grid': False,
        }
        mod.prepared_manifest_path(self.prepared).write_text(json.dumps(manifest), encoding='utf-8')
        self.downloads = self.root / 'Downloads'
        self.downloads.mkdir()
        self.download = self.downloads / 'Review from ChatGPT.xlsx'
        self.download.write_bytes(b'approved XLSX bytes')

    def test_default_final_folder_has_stable_short_name(self):
        self.assertEqual(mod.suggested_output_folder(self.prepared),
                         'Sample Opposition - Final')
        self.assertEqual(mod.suggested_output_folder(self.orig),
                         'Sample Opposition - Final')
        self.assertEqual(mod.generated_output_stem(self.prepared),'Sample Opposition')

    def test_filing_folder_cannot_overwrite_unrelated_files(self):
        folder=self.root/'Sample Opposition - Final'
        folder.mkdir()
        (folder/'Final Audit.txt').write_text('existing')
        expected={'Final Audit.txt','Motion with Tables.docx',
                  'Motion with Tables.pdf','_PDF_CITATION_AUDIT.tmp',
                  '_CITATION_COVERAGE.tmp'}
        with self.assertRaises(FileExistsError):
            mod.verify_reusable_filing_folder(folder,expected)
        self.assertTrue(mod.verify_reusable_filing_folder(folder,expected,replace_existing=True))
        (folder/'important-not-related.txt').write_text('preserve me')
        with self.assertRaisesRegex(RuntimeError,'unrelated files'):
            mod.verify_reusable_filing_folder(folder,expected,replace_existing=True)
        self.assertEqual((folder/'important-not-related.txt').read_text(),'preserve me')

    def test_failed_generation_cleans_internal_temp_audits(self):
        """A failed run keeps a useful audit, never leaked audit scratch files."""
        import openpyxl
        original=self.orig
        excel=self.download
        final=self.root/'Sample Opposition - Final'
        def progress(_line):
            (final/'_PDF_CITATION_AUDIT.tmp').write_text('scratch')
            (final/'_CITATION_COVERAGE.tmp').write_text('scratch')
        fake_book=MagicMock()
        fake_word=SimpleNamespace(DispatchEx=MagicMock(side_effect=RuntimeError('test Word failure')))
        with patch.object(mod,'win32',fake_word), \
             patch.object(mod,'preflight',return_value=([],[],{},'INTRODUCTION',False)), \
             patch.object(mod,'load_prepared_manifest',return_value=None), \
             patch.object(mod,'load_style_settings',return_value=mod.DEFAULT_STYLE), \
             patch.object(mod.openpyxl,'load_workbook',return_value=fake_book), \
             patch.object(mod,'read_setup',return_value={}), \
             patch.object(mod,'read_whole_motion_grid_option',return_value=False):
            with self.assertRaisesRegex(RuntimeError,'test Word failure'):
                mod.run(original,excel,output_parent=self.root,
                        progress_callback=progress)
        self.assertEqual({p.name for p in final.iterdir()},
                         {'Motion with Tables.docx','Final Audit.txt'})
        self.assertIn('ERROR:',(final/'Final Audit.txt').read_text())

    def test_source_original_is_located_by_hash(self):
        self.assertEqual(mod.preferred_review_folder(self.prepared), self.root)
        self.assertEqual(mod.preferred_review_folder(self.prepared,self.orig),self.root)

    def test_unknown_original_falls_back_to_prepared_folder(self):
        self.orig.write_bytes(b'modified original after prepare')
        self.assertEqual(mod.preferred_review_folder(self.prepared),self.prep_dir)
        self.assertEqual(mod.preferred_review_folder(self.prepared,self.orig),self.prep_dir)

    def test_downloaded_workbook_copied_beside_original_no_mutation(self):
        original_hash=mod.sha256_file(self.download)
        with patch.object(mod,'verify_spreadsheet_motion_binding',return_value={}):
            dest,copied=mod.file_into_motion_folder(self.prepared,self.download)
        self.assertTrue(copied)
        self.assertEqual(dest,self.root/'Sample Opposition_HipoLaw_Review.xlsx')
        self.assertEqual(mod.sha256_file(dest),original_hash)
        self.assertEqual(mod.sha256_file(self.download),original_hash)
        self.assertEqual(sorted(p.name for p in self.downloads.iterdir()),[self.download.name])
        self.assertFalse(any(p.name.startswith('.HipoLaw_Review_') for p in self.root.iterdir()))

    def test_identical_review_reuses_single_file(self):
        with patch.object(mod,'verify_spreadsheet_motion_binding',return_value={}):
            first,copied=mod.file_into_motion_folder(self.prepared,self.download)
            second,copied2=mod.file_into_motion_folder(self.prepared,self.download)
        self.assertTrue(copied)
        self.assertEqual((first,copied2),(second,False))
        self.assertEqual(len(list(self.root.glob('*_HipoLaw_Review.xlsx'))),1)

    def test_existing_different_review_requires_explicit_permission(self):
        dest=self.root/'Sample Opposition_HipoLaw_Review.xlsx'
        dest.write_bytes(b'attorney approved older sheet')
        with patch.object(mod,'verify_spreadsheet_motion_binding',return_value={}):
            with self.assertRaises(FileExistsError):
                mod.file_into_motion_folder(self.prepared,self.download)
        self.assertEqual(dest.read_bytes(),b'attorney approved older sheet')
        with patch.object(mod,'verify_spreadsheet_motion_binding',return_value={}):
            result,copied=mod.file_into_motion_folder(self.prepared,self.download,replace_existing=True)
        self.assertEqual(result,dest)
        self.assertTrue(copied)
        self.assertEqual(dest.read_bytes(),self.download.read_bytes())

    def test_xlsx_in_original_folder_is_not_duplicated(self):
        already=self.root/'My Approved Spreadsheet.xlsx'
        already.write_bytes(b'review workbook')
        with patch.object(mod,'verify_spreadsheet_motion_binding',return_value={}):
            target,copied=mod.file_into_motion_folder(self.prepared,already)
        self.assertFalse(copied)
        self.assertEqual(target,already)

    def test_invalid_workbook_never_copied(self):
        with patch.object(mod,'verify_spreadsheet_motion_binding',side_effect=ValueError('SHA256 mismatch')):
            with self.assertRaisesRegex(ValueError,'SHA256 mismatch'):
                mod.file_into_motion_folder(self.prepared,self.download)
        self.assertEqual(list(self.root.glob('*_HipoLaw_Review.xlsx')),[])

    def test_document_fingerprint_is_bound_before_workbook_copy(self):
        def assert_binding(prepared,download):
            self.assertEqual(prepared,self.prepared)
            self.assertEqual(download,self.download)
        with patch.object(mod,'verify_spreadsheet_motion_binding',side_effect=assert_binding) as check:
            mod.file_into_motion_folder(self.prepared,self.download)
            check.assert_called_once()

@unittest.skipUnless(os.environ.get('DISPLAY') or os.name=='nt','GUI tests need a display (xvfb-run on Linux)')
class TestUsabilityUI(unittest.TestCase):
    def setUp(self):
        self.app=mod.HipoLawApp()
        self.app.update_idletasks()
        self.addCleanup(self.app.destroy)

    def test_large_banner_mascot_uses_animated_full_size_images(self):
        self.assertGreaterEqual(self.app.images['idle'].height(),165)
        self.assertGreaterEqual(self.app.images['idle'].width(),240)
        self.assertEqual(self.app.images['run2'].width(),self.app.images['idle'].width())

    def test_two_explicit_prompt_and_browser_buttons(self):
        self.assertEqual(self.app.copy_prompt_btn.cget('text'),'1. Copy Prompt')
        self.assertEqual(self.app.chat_btn.cget('text'),'2. Open ChatGPT')
        with patch.object(self.app,'_copy_prompt') as cp, patch.object(mod.webbrowser,'open',return_value=True) as browser:
            self.app.open_chatgpt_prompt()
        cp.assert_not_called()
        browser.assert_called_once_with('https://chatgpt.com/')

    def test_copy_button_does_not_open_browser(self):
        with patch.object(self.app,'_copy_prompt') as cp, patch.object(mod.webbrowser,'open') as browser:
            self.app.copy_chatgpt_prompt()
        cp.assert_called_once()
        browser.assert_not_called()

if __name__=='__main__':unittest.main()
