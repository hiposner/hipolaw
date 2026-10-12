"""V16.4 offline tests. Windows Word COM integration must also be tested on the user's PC."""
import importlib.util, tempfile, unittest, json, shutil, hashlib
from pathlib import Path
from unittest.mock import patch

BASE=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('hipolaw_v164',BASE/'generate_legal_tables_v16_4_6.py')
mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
SOURCE=Path('/mnt/data/2026-10-11 Opp. to Motion for Judgment DRAFT 5 DC.docx')
APPROVED=Path('/mnt/data/La_Motte_Opposition_V15_Corrected_Headings.xlsx')

class FakeDocument:
    def __init__(self):
        self.Sections=type('Sections',(),{'Count':1})()
        self.closed=False
    def Repaginate(self):pass
    def SaveAs2(self, *_args, **_kwargs):pass
    def Close(self, **_kwargs):self.closed=True

class FakeWord:
    def __init__(self):
        self.doc=FakeDocument()
        self.Documents=type('Documents',(),{'Open':lambda _s,*a,**kw:self.doc})()
        self.quit=False
    def Quit(self):self.quit=True

@unittest.skipUnless(SOURCE.is_file() and APPROVED.is_file(),'La Motte test files not mounted')
class TestV164GuidedWorkflow(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
        self.temp=Path(self.tmp.name)
        self.source=self.temp/SOURCE.name
        shutil.copyfile(SOURCE,self.source)
        self.sourcehash=mod.sha256_file(self.source)

    def create_prepared(self,align=False):
        word=FakeWord()
        def make_preview(_doc,path):Path(path).write_bytes(b'%PDF-1.4 FAKE preview')
        with patch.object(mod,'normalize_caption_labels',return_value=[]), \
             patch.object(mod,'apply_whole_motion_grid',return_value=15) as grid, \
             patch.object(mod,'export_court_pdfa',side_effect=make_preview), \
             patch.object(mod,'calibrate_prepared_motion_numbers',return_value=True), \
             patch.object(mod,'verify_prepared_grid_preview',return_value=0.0):
            result=mod.prepare_motion(self.source,'INTRODUCTION',align_grid=align,
                                      output_parent=self.temp,word_factory=lambda:word)
        self.assertTrue(word.quit)
        self.assertTrue(word.doc.closed)
        self.assertEqual(mod.sha256_file(self.source),self.sourcehash)
        self.assertTrue(result['prepared'].exists())
        self.assertTrue(result['preview'].exists())
        self.assertTrue(result['manifest'].exists())
        grid.assert_called_once()
        self.assertEqual(grid.call_args.kwargs['enabled'],align)
        return result

    def make_binding_workbook(self,result,*,override_sha=None,override_name=None,grid='No'):
        import openpyxl
        dest=self.temp/'review.xlsx'
        wb=openpyxl.load_workbook(APPROVED)
        setup=wb['Setup']
        setup['A8']='Align entire motion to 28-line pleading grid'
        setup['B8']=grid
        setup['A9']='HipoLaw document SHA256'
        setup['B9']=override_sha if override_sha is not None else result['hash']
        setup['A10']='HipoLaw document filename'
        setup['B10']=override_name if override_name is not None else result['prepared'].name
        wb.save(dest);wb.close()
        return dest

    def test_prepared_copy_preserves_source_and_manifest(self):
        r=self.create_prepared(True)
        m=mod.load_prepared_manifest(r['prepared'])
        self.assertTrue(m['align_grid'])
        self.assertEqual(m['prepared_sha256'],r['hash'])
        self.assertEqual(m['original_sha256'],self.sourcehash)
        self.assertEqual(m['first_body_heading'],'INTRODUCTION')
        self.assertTrue(str(r['preview']).endswith('PREVIEW_NOT_FOR_FILING.pdf'))

    def test_generated_prompt_contains_exact_hash_filename_and_grid(self):
        r=self.create_prepared(True)
        prompt=mod.build_chatgpt_prompt(r['prepared'])
        self.assertIn(r['hash'],prompt)
        self.assertIn(r['prepared'].name,prompt)
        self.assertIn('Align entire motion to 28-line pleading grid | Yes',prompt)
        self.assertIn('HipoLaw document SHA256 | '+r['hash'],prompt)
        self.assertNotIn('<<MOTION_',prompt)

    def test_preflight_with_matching_prepared_motion_and_review(self):
        r=self.create_prepared(False)
        workbook=self.make_binding_workbook(r)
        auth,headings,aliases,first,grid=mod.preflight(r['prepared'],workbook)
        self.assertEqual((len(auth),len(headings),first),(37,14,'INTRODUCTION'))
        self.assertFalse(mod.read_whole_motion_grid_option(mod.read_spreadsheet_binding(workbook)[2]))

    def test_preflight_blocks_mismatched_docx(self):
        r=self.create_prepared(False)
        w=self.make_binding_workbook(r,override_sha='a'*64)
        with self.assertRaisesRegex(ValueError,'SHA256 mismatch'):
            mod.preflight(r['prepared'],w)

    def test_preflight_blocks_mismatched_grid(self):
        r=self.create_prepared(False)
        w=self.make_binding_workbook(r,grid='Yes')
        with self.assertRaisesRegex(ValueError,'Yes/No does not match'):
            mod.preflight(r['prepared'],w)

    def test_preflight_blocks_unbound_prepared_spreadsheet(self):
        r=self.create_prepared(False)
        with self.assertRaisesRegex(ValueError,'no "HipoLaw document SHA256" row'):
            mod.preflight(r['prepared'],APPROVED)

    def test_modified_prepared_doc_is_rejected(self):
        r=self.create_prepared(False)
        with open(r['prepared'],'ab') as f:f.write(b'changed')
        with self.assertRaisesRegex(ValueError,'changed after Step 1'):
            mod.load_prepared_manifest(r['prepared'])

    def test_preparation_requires_exact_unique_body_heading(self):
        with self.assertRaisesRegex(ValueError,'matches 0 Word paragraphs'):
            mod.prepare_motion(self.source,'WRONG HEADING',output_parent=self.temp,word_factory=FakeWord)
        self.assertEqual(mod.sha256_file(self.source),self.sourcehash)

    def test_suggested_heading_from_real_docx(self):
        self.assertEqual(mod.suggest_first_body_heading(self.source),'INTRODUCTION')

if __name__=='__main__':unittest.main()
