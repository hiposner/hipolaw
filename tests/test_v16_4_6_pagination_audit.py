"""Regressions for genuine PDF TOA/body page comparisons after grid reflow.

No local Windows installation is required: these tests exercise the actual
PDF text extraction used by the final audit, including split-page citations.
"""
import importlib.util
import tempfile
import unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('hipolaw_1646',ROOT/'generate_legal_tables_v16_4_8.py')
mod=importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


def authority(raw,display,category='Rule'):
    return dict(original=raw,final=display,category=category)


def make_three_page_pdf(filename, *, include_parents=True):
    doc=mod.pymupdf.open()
    toa=doc.new_page(width=612,height=792)
    def line(page,txt,y):
        page.insert_text((105,y),txt,fontsize=12)
    line(toa,'TABLE OF AUTHORITIES',75)
    entries=[
        ('Fed. R. Evid. 606(b)', '1'),
        ('Fed. R. Evid. 606(b)(1)', '2'),
        ('Fed. R. Civ. P. 50', '1'),
        ('Fed. R. Civ. P. 50(a)', '2'),
        ('Mitchell v. Forsyth, 472 U.S. 511 (1985)', '1'),
        ('Fed. R. Evid. 802', '2'),
    ]
    for i,(label,pagenum) in enumerate(entries):
        if not include_parents and label in ('Fed. R. Evid. 606(b)','Fed. R. Civ. P. 50'):
            continue
        line(toa,f'{label} ............. {pagenum}',105+22*i)
    b1=doc.new_page(width=612,height=792)
    line(b1,'Federal Rule of Evidence 606(b)',95)
    line(b1,'Rule 50',120)
    line(b1,'Mitchell v.',710)
    # A 28-line pleading number and page footer must not get mixed into
    # the citation when PDFs are concatenated across physical page boundaries.
    b1.insert_text((80,360),'17',fontsize=12)
    b1.insert_text((145,749),'CASE FOOTER 123',fontsize=8)
    b2=doc.new_page(width=612,height=792)
    line(b2,'Forsyth, 472 U.S. 511 (1985) matters here.',94)
    line(b2,'Fed. R. Evid. 606(b)(1)',119)
    line(b2,'Rule 50(a)',144)
    line(b2,'Fed. R. Evid. 801(c), 802.',169)
    doc.save(filename)
    doc.close()


@unittest.skipIf(mod.pymupdf is None,'PyMuPDF not installed')
class PaginationAuditV1646(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.pdf=Path(self.tmp.name)/'check.pdf'
        make_three_page_pdf(self.pdf)
        self.authorities=[
            authority('Federal Rule of Evidence 606(b)','Fed. R. Evid. 606(b)'),
            authority('Fed. R. Evid. 606(b)(1)','Fed. R. Evid. 606(b)(1)'),
            authority('Rule 50','Fed. R. Civ. P. 50'),
            authority('Rule 50(a)','Fed. R. Civ. P. 50(a)'),
            authority('Mitchell v. Forsyth, 472 U.S. 511 (1985)','Mitchell v. Forsyth, 472 U.S. 511 (1985)','Case'),
            authority('802','Fed. R. Evid. 802'),
        ]

    def test_toa_parent_rows_not_inferred_from_child_rows(self):
        path=Path(self.tmp.name)/'no_parent.pdf'
        make_three_page_pdf(path,include_parents=False)
        with mod.pymupdf.open(path) as pdf:
            hits=mod.parse_toa_pdf_entries(pdf,1,self.authorities)
        self.assertIsNone(hits['Fed. R. Evid. 606(b)'])
        self.assertIsNone(hits['Fed. R. Civ. P. 50'])
        self.assertEqual(hits['Fed. R. Evid. 606(b)(1)'],[2])
        self.assertEqual(hits['Fed. R. Civ. P. 50(a)'],[2])

    def test_parent_and_specific_authorities_have_distinct_toa_rows(self):
        with mod.pymupdf.open(self.pdf) as pdf:
            hits=mod.parse_toa_pdf_entries(pdf,1,self.authorities)
        for a,page in zip(self.authorities,[1,2,1,2,1,2]):
            self.assertEqual(hits[a['final']],[page],a['final'])

    def test_pdf_body_rule_boundary_no_parent_child_false_positive(self):
        with mod.pymupdf.open(self.pdf) as pdf:
            for item,expect in zip(self.authorities[:4],[[1],[2],[1],[2]]):
                self.assertEqual(mod.body_authority_page_hits(pdf,1,item),expect,item['final'])

    def test_case_broken_across_pages_anchors_to_start(self):
        with mod.pymupdf.open(self.pdf) as pdf:
            self.assertEqual(mod.body_authority_page_hits(pdf,1,self.authorities[4]),[1])

    def test_compound_rule_number_reconstructed_from_context(self):
        with mod.pymupdf.open(self.pdf) as pdf:
            self.assertEqual(mod.body_authority_page_hits(pdf,1,self.authorities[5]),[2])
        self.assertNotIn(r'802',mod.authority_pdf_patterns(self.authorities[5]))

    def test_all_six_authorities_pass_full_pdf_audit(self):
        report=Path(self.tmp.name)/'audit.txt'
        issues=mod.pdf_validate(self.pdf,self.authorities,[],2,report,prelim_phys=1,
                                numbered_grid=False,min_table_font_pt=12)
        self.assertFalse([x for x in issues if ' | TOA=' in x],issues)
        contents=report.read_text(encoding='utf-8')
        self.assertEqual(contents.count('PASS: '),6)
        self.assertNotIn('REVIEW: Fed.',contents)

    def test_if_parent_row_missing_audit_remains_blocking(self):
        path=Path(self.tmp.name)/'no_parent.pdf'
        make_three_page_pdf(path,include_parents=False)
        report=Path(self.tmp.name)/'no_parent_audit.txt'
        issues=mod.pdf_validate(path,self.authorities,[],2,report,prelim_phys=1,numbered_grid=False)
        self.assertTrue(any('Fed. R. Evid. 606(b) | TOA=None' in x for x in issues))
        with self.assertRaisesRegex(RuntimeError,'PAGINATION MISMATCH'):
            mod.ensure_pagination_matches(issues)


if __name__=='__main__':unittest.main()
