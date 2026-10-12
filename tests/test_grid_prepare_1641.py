"""Regression tests specifically for the V16.4 Step-1 alignment defect."""
import importlib.util
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest

BASE=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('hipolaw_gridfix',BASE/'generate_legal_tables_v16_4_2.py')
mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)

class CallableSections:
    def __init__(self,*sections):self._s=sections;self.Count=len(sections)
    def __call__(self,index):return self._s[index-1]

class GridParagraph:
    def __init__(self,before):
        self.format=SimpleNamespace(SpaceBefore=before)
        self.Range=SimpleNamespace(Text='\v'.join(map(str,range(1,29))),
                                   Start=100,End=200,ParagraphFormat=self.format)

class Header:
    def __init__(self,before):
        self.paragraph=GridParagraph(before)
        cell=SimpleNamespace(Range=SimpleNamespace(Paragraphs=[self.paragraph]))
        table=SimpleNamespace(Range=SimpleNamespace(Cells=[cell]))
        self.Range=SimpleNamespace(Tables=[table])
        self.LinkToPrevious=True

class HeaderCollection:
    def __init__(self,header):self.h=header
    def __call__(self,index):return self.h

class GridFixTests(unittest.TestCase):
    def test_step1_shifts_printed_numbers_but_not_first_page(self):
        body=Header(35.0); first=Header(35.0)
        sec=SimpleNamespace(Headers=HeaderCollection(body),
                            PageSetup=SimpleNamespace(DifferentFirstPageHeaderFooter=True),
                            first_header=first)
        doc=SimpleNamespace(Sections=CallableSections(sec))
        logs=[]
        self.assertTrue(mod.calibrate_prepared_motion_numbers(doc,enabled=True,
                          shift_pt=9.5,log=logs))
        self.assertEqual(body.paragraph.format.SpaceBefore,44.5)
        self.assertEqual(first.paragraph.format.SpaceBefore,35.0)
        self.assertIn('STEP 1 BODY GRID',logs[0])

    def test_no_alignment_selected_leaves_header_intact(self):
        body=Header(35.0);sec=SimpleNamespace(Headers=HeaderCollection(body),
                                 PageSetup=SimpleNamespace(DifferentFirstPageHeaderFooter=True))
        doc=SimpleNamespace(Sections=CallableSections(sec))
        self.assertFalse(mod.calibrate_prepared_motion_numbers(doc,enabled=False,
                         shift_pt=9.5,log=[]))
        self.assertEqual(body.paragraph.format.SpaceBefore,35.0)

    def test_no_separate_caption_header_fails_closed(self):
        body=Header(35.0);sec=SimpleNamespace(Headers=HeaderCollection(body),
                                 PageSetup=SimpleNamespace(DifferentFirstPageHeaderFooter=False))
        doc=SimpleNamespace(Sections=CallableSections(sec))
        with self.assertRaisesRegex(RuntimeError,'caption grid'):
            mod.calibrate_prepared_motion_numbers(doc,enabled=True,shift_pt=9.5,log=[])
        self.assertEqual(body.paragraph.format.SpaceBefore,35.0)

    def test_step3_restores_preliminary_without_double_shifting_body(self):
        prelim=Header(44.5);body=Header(44.5)
        sections=CallableSections(
          SimpleNamespace(Headers=HeaderCollection(Header(44.5))),
          SimpleNamespace(Headers=HeaderCollection(prelim)),
          SimpleNamespace(Headers=HeaderCollection(body)))
        doc=SimpleNamespace(Sections=sections)
        logs=[]
        mod.preserve_prepared_body_alignment(doc,2,3,shift_pt=9.5,log=logs)
        self.assertEqual(prelim.paragraph.format.SpaceBefore,35.0)
        self.assertEqual(body.paragraph.format.SpaceBefore,44.5)
        self.assertFalse(prelim.LinkToPrevious)
        self.assertFalse(body.LinkToPrevious)
        self.assertIn('no double adjustment',logs[-1])

    def test_existing_uploaded_preview_is_rejected_as_misaligned(self):
        p=Path('/mnt/data/2026-10-11 Opp. to Motion for Judgment FINAL_HipoLaw_Prepared_PREVIEW_NOT_FOR_FILING.pdf')
        if not p.is_file():self.skipTest('Uploaded PDF not available')
        logs=[]
        with self.assertRaisesRegex(RuntimeError,'still not aligned'):
            mod.verify_prepared_grid_preview(p,'INTRODUCTION',logs)
        self.assertTrue(any('9.5' in l for l in logs),logs)

    def test_matching_synthetic_grid_preview_is_accepted(self):
        if mod.pymupdf is None:self.skipTest('PyMuPDF not installed')
        pdf=mod.pymupdf.open();pdf.new_page(width=612,height=792)
        page=pdf.new_page(width=612,height=792)
        # Both have an identical top coordinate and therefore an aligned baseline.
        page.insert_text((80,80),'1',fontsize=12)
        page.insert_text((158,80),'INTRODUCTION',fontsize=12)
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/'mock.pdf';pdf.save(str(p));pdf.close()
            delta=mod.verify_prepared_grid_preview(p,'INTRODUCTION',[])
            self.assertLess(abs(delta),0.3)

if __name__=='__main__':unittest.main()
