"""V16.4.8: two preparation files, three final files, retained SHA256 safety."""
from __future__ import annotations
import importlib.util
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('hipolaw_1648',ROOT/'generate_legal_tables_v16_4_8.py')
app=importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

class FakeDoc:
    def __init__(self):
        self.Sections=type('Sections',(),{'Count':1})()
        self.closed=False
    def Repaginate(self):pass
    def SaveAs2(self,*args,**kwargs):pass
    def Close(self,**kwargs):self.closed=True

class FakeWord:
    def __init__(self):
        self.doc=FakeDoc()
        self.Documents=type('Documents',(),{'Open':lambda _,*a,**kw:self.doc})()
        self.quit=False
    def Quit(self):self.quit=True

class CleanPreparation(unittest.TestCase):
    def setUp(self):
        tmp=tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root=Path(tmp.name)
        self.source=self.root/'Motion 1.docx'
        # A ZIP-backed DOCX is sufficient for mocked Microsoft Word tests.
        with zipfile.ZipFile(self.source,'w') as archive:
            archive.writestr('[Content_Types].xml','<Types/>')
            archive.writestr('word/document.xml','<document/>')
        self.original_sha=app.sha256_file(self.source)

    def prepare(self,align=False,replace=False):
        word=FakeWord()
        exports=[]
        def temporary_export(_document,path):
            path=Path(path)
            exports.append(path)
            path.write_bytes(b'%PDF-1.7 mock-only')
        def check_preview(path,heading,log):
            self.assertTrue(Path(path).is_file())
            self.assertEqual(heading,'INTRODUCTION')
            log.append('TEMPORARY GRID CHECK: passed')
        with patch.object(app,'document_paragraphs',return_value=[('body',['INTRODUCTION'])]), \
             patch.object(app,'count_source_sections',return_value=1), \
             patch.object(app,'has_numbered_pleading_header',return_value=True), \
             patch.object(app,'load_style_settings',return_value=app.DEFAULT_STYLE), \
             patch.object(app,'normalize_caption_labels'), \
             patch.object(app,'apply_whole_motion_grid'), \
             patch.object(app,'calibrate_prepared_motion_numbers'), \
             patch.object(app,'verify_motion_body_text_unchanged',return_value=10), \
             patch.object(app,'export_court_pdfa',side_effect=temporary_export), \
             patch.object(app,'verify_prepared_grid_preview',side_effect=check_preview):
            result=app.prepare_motion(self.source,'INTRODUCTION',align_grid=align,
                output_parent=self.root,word_factory=lambda:word,replace_existing=replace)
        self.assertTrue(word.quit)
        self.assertTrue(word.doc.closed)
        self.assertEqual(app.sha256_file(self.source),self.original_sha)
        return result,exports

    def test_step_one_delivers_only_prepared_word_and_audit(self):
        result,exports=self.prepare(align=False)
        self.assertEqual(exports,[])
        self.assertEqual(result['prepared'].name,'Prepared Motion.docx')
        self.assertEqual(result['folder'].name,'Motion 1 - Prepared')
        self.assertEqual(sorted(x.name for x in result['folder'].iterdir()),
                         ['Preparation Audit.txt','Prepared Motion.docx'])
        self.assertIsNone(result['preview'])
        manifest=app.load_prepared_manifest(result['prepared'])
        self.assertEqual(manifest['prepared_sha256'],result['hash'])
        self.assertEqual(manifest['original_sha256'],self.original_sha)
        self.assertEqual(manifest['prepared_filename'],'Prepared Motion.docx')
        self.assertFalse(app.prepared_manifest_path(result['prepared']).exists())
        self.assertIn(app.PREPARATION_MANIFEST_PREFIX,result['audit'].read_text())
        self.assertEqual(app.preferred_review_folder(result['prepared']),self.root)

    def test_aligned_step_one_verifies_pdf_without_saving_pdf(self):
        result,exports=self.prepare(align=True)
        self.assertEqual(len(exports),1)
        self.assertFalse(exports[0].exists())  # temp directory cleaned up
        self.assertEqual({x.suffix for x in result['folder'].iterdir()}, {'.docx','.txt'})
        self.assertTrue(app.load_prepared_manifest(result['prepared'])['align_grid'])
        self.assertIn('TEMPORARY GRID CHECK: passed',result['audit'].read_text())

    def test_repeat_preparation_requires_opt_in_and_no_duplicate_files(self):
        first,_=self.prepare()
        with self.assertRaises(FileExistsError):
            self.prepare()
        second,_=self.prepare(replace=True)
        self.assertEqual(first['prepared'],second['prepared'])
        self.assertEqual(len(list(second['folder'].iterdir())),2)
        self.assertEqual(app.sha256_file(self.source),self.original_sha)

    def test_foreign_files_prevent_replacement(self):
        result,_=self.prepare()
        foreign=result['folder']/'Client Note.txt'
        foreign.write_text('never replace')
        with self.assertRaisesRegex(RuntimeError,'unrelated files'):
            self.prepare(replace=True)
        self.assertEqual(foreign.read_text(),'never replace')

    def test_modified_word_and_lost_audit_remain_blocking(self):
        result,_=self.prepare()
        with result['prepared'].open('ab') as f:f.write(b'changed')
        with self.assertRaisesRegex(ValueError,'changed after Step 1'):
            app.load_prepared_manifest(result['prepared'])
        result['audit'].unlink()
        with self.assertRaisesRegex(ValueError,'matching Preparation Audit'):
            app.load_prepared_manifest(result['prepared'])

    def test_old_json_manifest_remains_supported(self):
        legacy_dir=self.root/'HipoLaw_Prepared_20261011_123456'
        legacy_dir.mkdir()
        legacy=legacy_dir/'Motion 1_HipoLaw_Prepared.docx'
        legacy.write_bytes(b'older prepared file')
        identity=dict(schema_version=1,prepared_filename=legacy.name,
                      prepared_sha256=app.sha256_file(legacy),
                      first_body_heading='INTRODUCTION',original_filename=self.source.name,
                      original_sha256=self.original_sha,align_grid=False)
        app.prepared_manifest_path(legacy).write_text(json.dumps(identity))
        self.assertEqual(app.load_prepared_manifest(legacy)['prepared_sha256'],identity['prepared_sha256'])

class FinalOutputNaming(unittest.TestCase):
    def test_final_folder_and_file_names_are_short(self):
        self.assertEqual(app.suggested_output_folder(Path('Motion 1.docx')),'Motion 1 - Final')
        self.assertEqual(app.generated_output_stem(Path('Motion 1_HipoLaw_Prepared.docx')),'Motion 1')
        self.assertEqual(app.generated_output_stem(Path('Motion 1 - Prepared.docx')),'Motion 1')
        # The final files are deliberately independent of a long source filename.
        source=Path(__file__).resolve().parents[1]/'generate_legal_tables_v16_4_8.py'
        text=source.read_text(encoding='utf-8')
        self.assertIn("output=folder/'Motion with Tables.docx'",text)
        self.assertIn("pdf=folder/'Motion with Tables.pdf'",text)
        self.assertIn("report=folder/'Final Audit.txt'",text)

    def test_replacement_rejects_foreign_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder=Path(tmp)/'Motion 1 - Final'
            folder.mkdir()
            for name in ('Motion with Tables.docx','Motion with Tables.pdf','Final Audit.txt'):
                (folder/name).write_text('test')
            expect={x.name for x in folder.iterdir()}|{'_PDF_CITATION_AUDIT.tmp','_CITATION_COVERAGE.tmp'}
            self.assertTrue(app.verify_reusable_filing_folder(folder,expect,replace_existing=True))
            (folder/'unrelated.txt').write_text('do not delete')
            with self.assertRaisesRegex(RuntimeError,'unrelated files'):
                app.verify_reusable_filing_folder(folder,expect,replace_existing=True)

if __name__=='__main__':unittest.main()

@unittest.skipIf(app.pymupdf is None,'PyMuPDF not installed')
class CrossPageCitationStartsOnSeven(unittest.TestCase):
    """The exact failure reported by the user: page 7 to page 8."""
    def make_pdf(self,target,second_page='Forsyth, 472 U.S. 511 (1985)'):
        pdf=app.pymupdf.open()
        prelim=pdf.new_page(width=612,height=792)
        prelim.insert_text((120,80),'TABLE OF AUTHORITIES',fontsize=12)
        prelim.insert_text((120,110),'Mitchell v. Forsyth, 472 U.S. 511 (1985) ............... 7',fontsize=12)
        for body_page in range(1,9):
            page=pdf.new_page(width=612,height=792)
            page.insert_text((120,90),f'Motion body page {body_page}',fontsize=12)
            if body_page==7:
                page.insert_text((120,706),'Mitchell v.',fontsize=12)
                page.insert_text((145,744),'Footer interrupts text extraction',fontsize=7)
            if body_page==8:
                page.insert_text((120,117),second_page,fontsize=12)
        pdf.save(target)
        pdf.close()

    def test_genuine_case_split_page_seven_to_eight_is_seven(self):
        with tempfile.TemporaryDirectory() as temp:
            target=Path(temp)/'split.pdf'
            self.make_pdf(target)
            item={'original':'Mitchell v. Forsyth, 472 U.S. 511 (1985)',
                  'final':'Mitchell v. Forsyth, 472 U.S. 511 (1985)','category':'Case'}
            with app.pymupdf.open(target) as pdf:
                self.assertEqual(app.case_title_split_at_page_boundary(pdf,1,item),[7])
                self.assertEqual(app.body_authority_page_hits(pdf,1,item),[7])
            report=Path(temp)/'audit.txt'
            issues=app.pdf_validate(target,[item],[],2,report,prelim_phys=1,numbered_grid=False)
            self.assertFalse([x for x in issues if ' | TOA=' in x],issues)
            self.assertIn('TOA=[7] | body=[7]',report.read_text())

    def test_unrelated_continuation_is_not_accepted_as_citation(self):
        with tempfile.TemporaryDirectory() as temp:
            target=Path(temp)/'unrelated.pdf'
            self.make_pdf(target,second_page='Other unrelated material')
            item={'original':'Mitchell v. Forsyth, 472 U.S. 511 (1985)',
                  'final':'Mitchell v. Forsyth, 472 U.S. 511 (1985)','category':'Case'}
            with app.pymupdf.open(target) as pdf:
                self.assertEqual(app.case_title_split_at_page_boundary(pdf,1,item),[])
                self.assertEqual(app.body_authority_page_hits(pdf,1,item),[])
            report=Path(temp)/'audit.txt'
            issues=app.pdf_validate(target,[item],[],2,report,prelim_phys=1,numbered_grid=False)
            self.assertTrue(any('TOA=[7] | body=[]' in x for x in issues))
            with self.assertRaisesRegex(RuntimeError,'PAGINATION MISMATCH'):
                app.ensure_pagination_matches(issues)
