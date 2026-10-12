"""V16.4.9: protect inline verdict graphics, preserve TOC dots, order TOA."""
from __future__ import annotations
import importlib.util
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
import zipfile

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('hipolaw_1649', ROOT/'generate_legal_tables_v16_4_9.py')
app=importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)


class FakeFormat:
    def __init__(self):
        self.SpaceBefore=15
        self.SpaceAfter=9
        self.LineSpacingRule=2
        self.LineSpacing=28
        self.Alignment=0
        self.LeftIndent=40
        self.RightIndent=20
        self.FirstLineIndent=25

class FakeRange:
    def __init__(self,start=100,inline_count=0,in_table=False):
        self.Start=start
        self.ParagraphFormat=FakeFormat()
        self.InlineShapes=type('Inlines',(),{'Count':inline_count})()
        self._table=in_table
    def Information(self,n):
        assert n==12
        return self._table

class FakeParagraph:
    def __init__(self,at,artwork=0,table=False):
        self.Range=FakeRange(at,artwork,table)

class FakeDoc:
    def __init__(self):
        self.Content=type('C',(),{'End':700})()
        self.body=[FakeParagraph(100),FakeParagraph(200,1),FakeParagraph(300),
                   FakeParagraph(400,2),FakeParagraph(500,0,True)]
    def Range(self,start,end):return type('R',(),{'Paragraphs':self.body})()

class TestArtPreservation(unittest.TestCase):
    def test_graphic_paragraphs_use_at_least_and_text_stays_exact(self):
        doc=FakeDoc();log=[]
        with patch.object(app,'heading_ranges',return_value=[type('Heading',(),{'Start':100})()]):
            count=app.apply_whole_motion_grid(doc,'INTRODUCTION',enabled=True,
                has_28_line_header=True,spacing_pt=24,log=log)
        self.assertEqual(count,2)
        self.assertEqual(doc.body[0].Range.ParagraphFormat.LineSpacingRule,4)
        self.assertEqual(doc.body[2].Range.ParagraphFormat.LineSpacingRule,4)
        self.assertEqual(doc.body[1].Range.ParagraphFormat.LineSpacingRule,app.WD_LINE_SPACE_AT_LEAST)
        self.assertEqual(doc.body[3].Range.ParagraphFormat.LineSpacingRule,app.WD_LINE_SPACE_AT_LEAST)
        self.assertEqual(doc.body[1].Range.ParagraphFormat.LineSpacing,24)
        self.assertEqual(doc.body[4].Range.ParagraphFormat.SpaceBefore,15)
        self.assertIn('2 inline-artwork paragraph(s)', '\n'.join(log))

    def test_unknown_artwork_inspection_stops_instead_of_clipping(self):
        class BrokenRange:
            @property
            def InlineShapes(self):
                raise RuntimeError('Office shape query failed')
        with self.assertRaisesRegex(RuntimeError,'Could not inspect paragraph inline artwork'):
            app.paragraph_has_inline_artwork(BrokenRange())

    def _docx(self,path, *, rule='atLeast',height_pt=80,size_pt=80):
        h=int(height_pt*12700)
        xml=('<?xml version="1.0" encoding="utf-8"?>'
             '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
             'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">'
             '<w:body><w:p><w:pPr><w:spacing w:line="480" w:lineRule="'+rule+'"/>'
             '</w:pPr><w:r><w:drawing><wp:inline><wp:extent cx="790000" cy="'+str(h)+'"/>'
             '</wp:inline></w:drawing></w:r></w:p></w:body></w:document>')
        with zipfile.ZipFile(path,'w') as z:z.writestr('word/document.xml',xml)

    def test_fails_if_image_is_clipped_by_exact_24pt_spacing(self):
        with tempfile.TemporaryDirectory() as td:
            a=Path(td)/'original.docx';b=Path(td)/'broken.docx'
            self._docx(a,rule='auto');self._docx(b,rule='exact')
            sizes,unsafe=app.inspect_inline_graphics_docx(b)
            self.assertEqual(len(sizes),1)
            self.assertEqual(len(unsafe),1)
            with self.assertRaisesRegex(RuntimeError,'GRAPHIC CLIPPING RISK'):
                app.verify_inline_graphics_preserved(a,b,[])

    def test_accepts_at_least_spacing_and_preserves_image_geometry(self):
        with tempfile.TemporaryDirectory() as td:
            a=Path(td)/'original.docx';b=Path(td)/'safe.docx'
            self._docx(a,rule='auto');self._docx(b,rule='atLeast')
            log=[];app.verify_inline_graphics_preserved(a,b,log)
            self.assertIn('1 graphic(s)',log[0])
            self.assertEqual(app.inspect_inline_graphics_docx(b)[1],[])

    def test_does_not_silently_rescale_figure(self):
        with tempfile.TemporaryDirectory() as td:
            a=Path(td)/'original.docx';b=Path(td)/'small.docx'
            self._docx(a,height_pt=80);self._docx(b,height_pt=50)
            with self.assertRaisesRegex(RuntimeError,'resized'):
                app.verify_inline_graphics_preserved(a,b,[])

    def test_real_original_motion_has_two_verdict_form_graphics(self):
        # Optional developer fixture only; never include confidential motion in ZIP.
        source=Path('/mnt/data/2026-10-11 Opp. to Motion for Judgment DRAFT 5 DC.docx')
        if not source.exists():self.skipTest('Development reference Word file unavailable')
        sizes,unsafe=app.inspect_inline_graphics_docx(source)
        self.assertEqual(len(sizes),2)
        self.assertEqual(unsafe,[])
        self.assertTrue(all(h>24*12700 for _,h in sizes))

class TestTocAndToa(unittest.TestCase):
    def test_toc_number_has_period_at_all_levels(self):
        self.assertEqual(app.toc_display_label('I.','INTRODUCTION'),'I. INTRODUCTION')
        self.assertEqual(app.toc_display_label('A.','Heading.'),'A. Heading')
        self.assertEqual(app.toc_display_label('iv.','Subheading.'),'iv. Subheading')
        self.assertEqual(app.toc_display_label('7)','Text'),'7. Text')
        self.assertEqual(app.toc_display_label('','Title.'),'Title')
        self.assertEqual(app.toc_display_label('I.','I. INTRODUCTION'),'I. INTRODUCTION')

    def test_other_authorities_always_last(self):
        authorities=[{'category':'Rule'}, {'category':'Other'},
                     {'category':'Case'},{'category':'Statute'}]
        text=app.make_prelim_text(authorities,app.DEFAULT_STYLE)
        headings=['CASES\r','STATUTES\r','RULES\r','OTHER AUTHORITIES\r']
        self.assertEqual([text.index(x) for x in headings],sorted(text.index(x) for x in headings))
        self.assertTrue(text.rfind('OTHER AUTHORITIES') > text.rfind('RULES'))
        # Ensure the TOA uses the existing category codes, not reassigned ones.
        self.assertEqual(app.category_code('Other'),3)
        self.assertEqual(app.category_code('Rule'),4)

    def test_toa_header_is_centered_without_inherited_indent(self):
        src=ROOT/'tests'/'test_v16_4_5_table_style.py'
        # The standalone COM-style mock from the existing style tests.
        spec=importlib.util.spec_from_file_location('hipo_styles_1649',src)
        mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)
        d=mod.FakeDoc(['TABLE OF CONTENTS','TABLE OF AUTHORITIES','CASES','STATUTES','RULES','OTHER AUTHORITIES'])
        fmt=d.paragraphs[1].Range.ParagraphFormat
        fmt.LeftIndent=45;fmt.RightIndent=12;fmt.FirstLineIndent=20
        log=[]
        app.standardize_table_results(d,log,lambda msg:None,True,2,app.DEFAULT_STYLE)
        self.assertEqual((fmt.Alignment,fmt.LeftIndent,fmt.RightIndent,fmt.FirstLineIndent),(1,0,0,0))
        self.assertTrue(any('TOA TITLE: centered' in x for x in log))

if __name__=='__main__':unittest.main()
