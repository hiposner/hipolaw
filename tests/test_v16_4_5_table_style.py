"""Regression tests for formatting native Word-generated TOC/TOA heading categories."""
import importlib.util
import unittest
from pathlib import Path

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('hipolaw_table',ROOT/'generate_legal_tables_v16_4_8.py')
mod=importlib.util.module_from_spec(spec);spec.loader.exec_module(mod)

class FakeFont:
    def __init__(self):
        self.Size=None;self.Bold=None;self.Italic=None;self.Position=None
class FakeParagraphFormat:
    def __init__(self):
        self.SpaceBefore=None;self.SpaceAfter=None;self.LineSpacing=None
        self.LineSpacingRule=None;self.Alignment=None
class FakeRange:
    def __init__(self,text):
        self.Text=text+'\r';self.Font=FakeFont();self.ParagraphFormat=FakeParagraphFormat()
class FakePara:
    def __init__(self,text): self.Range=FakeRange(text)
class FakeField:
    def __init__(self,code):
        self.Code=type('Code',(),{'Text':code})()
        self.Result=FakeRange('')
        self.Result.Paragraphs=[]
class FakeDoc:
    def __init__(self,labels):
        self.Fields=[FakeField('TOC \\f C \\h \\z')]+[FakeField('TOA \\c '+str(c)) for c in (1,2,3,4)]
        self.paragraphs=[FakePara(x) for x in labels]
    def Sections(self,number):
        assert number==2
        return type('Section',(),{'Range':type('R',(),{'Paragraphs':self.paragraphs})()})()

class TableFormatting(unittest.TestCase):
    def test_uppercase_other_authorities_has_valid_other_group(self):
        labels=['TABLE OF CONTENTS','TABLE OF AUTHORITIES','CASES','STATUTES','OTHER AUTHORITIES','RULES']
        doc=FakeDoc(labels)
        log=[]
        mod.standardize_table_results(doc,log,lambda _:None,True,2,mod.DEFAULT_STYLE)
        mapping=dict(zip(labels,doc.paragraphs))
        self.assertEqual(mapping['OTHER AUTHORITIES'].Range.Font.Position,
                         mod.DEFAULT_STYLE['pleading_grid']['vertical_raise_pt']['other'])
        self.assertEqual(mapping['CASES'].Range.Font.Position,
                         mod.DEFAULT_STYLE['pleading_grid']['vertical_raise_pt']['cases'])
        self.assertTrue(any('TABLE TYPOGRAPHY' in line for line in log))

    def test_title_case_categories_and_unraised_body(self):
        labels=['TABLE OF CONTENTS','TABLE OF AUTHORITIES','Cases','Statutes','Other Authorities','Rules']
        doc=FakeDoc(labels)
        import copy
        style=copy.deepcopy(mod.DEFAULT_STYLE)
        style['category_headings']['uppercase']=False
        mod.standardize_table_results(doc,[],lambda _:None,True,2,style)
        self.assertEqual(doc.paragraphs[4].Range.Font.Position,style['pleading_grid']['vertical_raise_pt']['other'])
        # For an ordinary non-pleading source, generated titles must stay at baseline.
        plain=FakeDoc(labels)
        mod.standardize_table_results(plain,[],lambda _:None,False,2,style)
        self.assertTrue(all(p.Range.Font.Position==0 for p in plain.paragraphs))
        self.assertTrue(all(f.Result.Font.Position==0 for f in plain.Fields))

    def test_categories_have_exact_style_settings_keys(self):
        known={'toc','cases','statutes','other','rules'}
        self.assertEqual(set(mod.DEFAULT_STYLE['pleading_grid']['vertical_raise_pt']),known)

if __name__=='__main__':unittest.main()
