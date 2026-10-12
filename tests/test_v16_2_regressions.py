"""Cross-platform safety tests for the V15.1-based Hipo Law V16 release.

Full Word COM and PDF pagination tests must also be performed on Windows.
"""
import importlib.util
import unittest
from pathlib import Path
from unittest.mock import patch, Mock
import json
import tempfile
import fitz

SCRIPT = Path(__file__).resolve().parents[1] / 'generate_legal_tables_v16_4_2.py'
spec = importlib.util.spec_from_file_location('hipolaw_v16', SCRIPT)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


class TestTocContract(unittest.TestCase):
    def test_first_level_number(self):
        self.assertEqual(mod.toc_display_label('I.', 'INTRODUCTION'), 'I INTRODUCTION')
        self.assertEqual(mod.toc_display_label('IV.', 'ARGUMENT'), 'IV ARGUMENT')

    def test_sublevel_number(self):
        self.assertEqual(mod.toc_display_label('A.', 'The Rule 50 motion fails.'),
                         'A The Rule 50 motion fails')
        self.assertEqual(mod.toc_display_label('ii.', 'Qualified immunity was waived.'),
                         'ii Qualified immunity was waived')

    def test_period_inside_closing_quote(self):
        self.assertEqual(mod.toc_display_label('iv.', 'Evidence of malice was not “none.”'),
                         'iv Evidence of malice was not “none”')

    def test_internal_punctuation_preserved(self):
        self.assertEqual(mod.toc_display_label('iii.', 'The U.S. Court ruled. Again.'),
                         'iii The U.S. Court ruled. Again')

    def test_number_is_not_generated_from_position(self):
        # A numbering gap/restart in the source is carried through literally.
        self.assertEqual(mod.toc_display_label('VII.', 'CONCLUSION'), 'VII CONCLUSION')

    def test_no_continued_heading_source(self):
        self.assertNotIn('(continued)', mod.make_prelim_text([{'category': 'Case'}]))


class DummyParagraph:
    def __init__(self,start,end):
        self.Range = DummyRange(start,end)

class DummyRange:
    def __init__(self,start,end):
        self.Start=start
        self.End=end
        self.Style=100
        self.ListFormat = DummyList()

class DummyList:
    def __init__(self):self.calls=0
    def RemoveNumbers(self):self.calls+=1

class DummyDoc:
    def __init__(self):
        self.prelim=[DummyParagraph(i*10,(i+1)*10) for i in range(7)]
        self.body=DummyParagraph(100,110)
        self.sections={2:type('Sec',(),{'Range':type('Range',(),{'Paragraphs':self.prelim + [self.body]})()})()}
    def Sections(self,n):return self.sections[n]

class TestNumberingProtection(unittest.TestCase):
    def test_only_new_preliminary_paragraphs_lose_numbering(self):
        doc=DummyDoc();notes=[]
        fake_heading=type('Heading',(),{'Start':100})()
        with patch.object(mod,'heading_ranges',return_value=[fake_heading]):
            mod.protect_preliminary_paragraphs(doc,2,'INTRODUCTION',notes)
        self.assertTrue(all(p.Range.ListFormat.calls==1 for p in doc.prelim))
        self.assertTrue(all(p.Range.Style==-1 for p in doc.prelim))
        self.assertEqual(doc.body.Range.ListFormat.calls,0)
        self.assertEqual(doc.body.Range.Style,100)
        self.assertTrue(any('NUMBERING ISOLATION' in x for x in notes))

    def test_numbering_change_is_fatal(self):
        class R:
            Paragraphs=lambda self,n: type('P',(),{'Range':type('R2',(),{'Text':'INTRODUCTION\r', 'ListFormat':type('L',(),{'ListString':'X.'})()})()})()
        with patch.object(mod,'heading_ranges',return_value=[R()]):
            with self.assertRaisesRegex(RuntimeError,'SOURCE MOTION NUMBERING CHANGED'):
                mod.check_heading_numbering(object(),[('INTRODUCTION','I.',1)],'test')

class TestExternalStyleConfig(unittest.TestCase):
    def test_config_is_read_each_run_and_case_offset_corrected(self):
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'HipoLaw_Style_Settings.json'
            path.write_text(json.dumps(mod.DEFAULT_STYLE))
            with patch.object(mod,'style_file_path',return_value=path):
                style=mod.load_style_settings()
                self.assertEqual(style['pleading_grid']['vertical_raise_pt']['cases'],3.5)
                style['heading_font_size_pt']=13
                path.write_text(json.dumps(style))
                self.assertEqual(mod.load_style_settings()['heading_font_size_pt'],13)

    def test_bad_config_is_rejected_before_creating_output(self):
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'HipoLaw_Style_Settings.json'
            settings=json.loads(json.dumps(mod.DEFAULT_STYLE))
            settings['entry_font_size_pt']='giant'
            path.write_text(json.dumps(settings))
            with patch.object(mod,'style_file_path',return_value=path):
                with self.assertRaisesRegex(ValueError,'entry_font_size_pt'):
                    mod.load_style_settings()

    def test_category_display_uses_uppercase_without_continuation(self):
        a=[{'category':'Case'}, {'category':'Statute'}, {'category':'Rule'}]
        result=mod.make_prelim_text(a,mod.DEFAULT_STYLE)
        for heading in ('CASES\r','STATUTES\r','RULES\r'):
            self.assertIn(heading,result)
        self.assertNotIn('(continued)',result)

    def test_pdf_page_bounds_reports_normal_pdf_clean(self):
        doc=fitz.open();pg=doc.new_page(width=612,height=792)
        pg.insert_text((120,120),'Safe text within page')
        self.assertEqual(mod.pdf_page_edge_audit(doc),[])
        doc.close()



class TestCourtFilingExport(unittest.TestCase):
    def test_word_pdf_a_export_parameters(self):
        d=Mock()
        mod.export_court_pdfa(d, Path('C:/temp/court.pdf'))
        d.ExportAsFixedFormat.assert_called_once()
        args,kwargs=d.ExportAsFixedFormat.call_args
        self.assertEqual(args[1],mod.PDF)
        self.assertTrue(kwargs['UseISO19005_1'])
        self.assertFalse(kwargs['IncludeDocProps'])
        self.assertFalse(kwargs['KeepIRM'])
        self.assertTrue(kwargs['DocStructureTags'])

    def test_caption_font_adjustment_only_before_motion(self):
        class RunFont:
            Name='Times New Roman Bold'
            Bold=None
        class ParaRange:
            def __init__(self,text): self.Text=text;self.Font=RunFont()
        class Para:
            def __init__(self,text): self.Range=ParaRange(text)
        class Table:
            def __init__(self,pos,texts):
                self.Range=type('Range',(),{})()
                self.Range.Start=pos
                self.Range.Paragraphs=[Para(text) for text in texts]
        class Tables:
            def __init__(self):
                self.objects=[Table(10,['MOTION TITLE','Date: November 5','Time: 2PM','Location: Oakland','Judge: The Honorable']),
                              Table(1000,['Date: Some evidence'])]
                self.Count=len(self.objects)
            def __call__(self,n):return self.objects[n-1]
        doc=type('Doc',(),{})();doc.Tables=Tables()
        anchor=type('Anchor',(),{'Start':900})()
        log=[]
        with patch.object(mod,'heading_ranges',return_value=[anchor]):
            changed=mod.normalize_caption_labels(doc,log,'INTRODUCTION')
        self.assertEqual(set(changed),{'Date','Time','Location','Judge'})
        font0=doc.Tables(1).Range.Paragraphs[0].Range.Font
        self.assertEqual(font0.Name,'Times New Roman Bold')
        for para in doc.Tables(1).Range.Paragraphs[1:]:
            self.assertEqual(para.Range.Font.Name,'Times New Roman')
            self.assertEqual(para.Range.Font.Bold,0)
        self.assertEqual(doc.Tables(2).Range.Paragraphs[0].Range.Font.Name,'Times New Roman Bold')

    def sample_pdf(self,path,with_declaration=True):
        doc=fitz.open()
        p=doc.new_page()
        p.insert_text((50,50),'Motion caption')
        if with_declaration:
            xmp='''<x:xmpmeta xmlns:x="adobe:ns:meta/" xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"><rdf:RDF><rdf:Description pdfaid:part="1" pdfaid:conformance="B"/></rdf:RDF></x:xmpmeta>'''
            doc.set_xml_metadata(xmp)
        doc.save(path)
        doc.close()

    def test_plain_pdf_blocked_not_misidentified_as_pdfa(self):
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'ordinary.pdf'
            self.sample_pdf(path,False)
            with self.assertRaisesRegex(RuntimeError,'PDF/A-1 declaration not found'):
                mod.pdfa_safety_audit(path)

    def test_xmp_declaration_is_not_misrepresented_as_full_validation(self):
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'declared.pdf'
            self.sample_pdf(path,True)
            with patch.dict(mod.os.environ,{'VERAPDF_CLI':''}),patch.object(mod.shutil,'which',return_value=None):
                results,warnings=mod.pdfa_safety_audit(path)
            self.assertTrue(any('PDF/A-1b' in x for x in results))
            self.assertTrue(any('NOT RUN' in x for x in results))
            self.assertTrue(warnings)

    def test_embedded_file_is_blocked_even_with_pdfa_marker(self):
        with tempfile.TemporaryDirectory() as td:
            path=Path(td)/'attached.pdf'
            d=fitz.open();d.new_page().insert_text((50,50),'Hi')
            d.set_xml_metadata('''<x:xmpmeta xmlns:x="adobe:ns:meta/" xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#" xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"><rdf:RDF><rdf:Description pdfaid:part="1" pdfaid:conformance="B"/></rdf:RDF></x:xmpmeta>''')
            d.embfile_add('attached.txt',b'data')
            d.save(path);d.close()
            with self.assertRaisesRegex(RuntimeError,'embedded attachments'):
                mod.pdfa_safety_audit(path)

if __name__=='__main__':unittest.main()
