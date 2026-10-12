"""Generic rule subdivision and prefix regression tests (no named-case assumptions)."""
import importlib.util
import re
import unittest
from pathlib import Path
from unittest.mock import patch

BASE=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('hipolaw_nested',BASE/'generate_legal_tables_v16_4_8.py')
mod=importlib.util.module_from_spec(spec); spec.loader.exec_module(mod)

class Range:
    def __init__(self,txt,start=0,end=None):
        self._txt=txt;self.Start=start;self.End=len(txt) if end is None else end
    @property
    def Text(self):return self._txt[self.Start:self.End]
    @property
    def Duplicate(self):return Range(self._txt,self.Start,self.End)
    def SetRange(self,a,b):self.Start=a;self.End=b
class Fields:
    def __init__(self):self.calls=[]
    def Add(self,r,typ,code,p):self.calls.append((r.Start,code))
class Doc:
    def __init__(self,body):self.stories={1:Range(body),2:Range('')};self.Fields=Fields()
    def StoryRanges(self,n):return self.stories[n]
def find(doc,phrase):
    hits=[]
    for n,story in ((1,'body'),(2,'footnotes')):
        for m in re.finditer(re.escape(phrase),doc.stories[n]._txt,re.I):
            hits.append((story,m.start(),m.end()))
    return hits

def auth(cite):return {'original':cite,'final':cite,'category':'Evidence Rule'}

class NestedRules(unittest.TestCase):
    def test_entire_subdivision_chain_not_truncated(self):
        self.assertEqual(mod.variants_for_authority(auth('Fed. R. Evid. 606(b)(1)')),
            ['Fed. R. Evid. 606(b)(1)','Rule 606(b)(1)',
             'Federal Rule of Evidence 606(b)(1)'])
        self.assertEqual(mod.variants_for_authority(auth('Fed. R. Evid. 803(10)(A)(2)')),
            ['Fed. R. Evid. 803(10)(A)(2)','Rule 803(10)(A)(2)',
             'Federal Rule of Evidence 803(10)(A)(2)'])

    def test_parent_and_nested_child_can_both_be_anchored(self):
        body=('Federal Rule of Evidence 606(b) excludes certain testimony. '
              'Fed. R. Evid. 606(b)(1) contains a specific limitation. '
              'Rule 606(b)(1) separately. Rule 606(b) independently.')
        doc=Doc(body)
        log=[]
        with patch.object(mod,'find_text_in_stories',side_effect=find):
            okay=mod.insert_markers(doc,[auth('Federal Rule of Evidence 606(b)'),
                                           auth('Fed. R. Evid. 606(b)(1)')],log,lambda _:None,
                                      aliases={'federal rule of evidence 606(b)':['Rule 606(b)']})
        self.assertTrue(okay,'\n'.join(log))
        self.assertEqual(len(doc.Fields.calls),4)
        self.assertFalse(any('AMBIGUOUS OVERLAP' in x for x in log))
        self.assertTrue(mod.exact_safe_anchor_in_text(body,'Rule 606(b)'))
        self.assertTrue(mod.exact_safe_anchor_in_text(body,'Rule 606(b)(1)'))

    def test_shorter_rule_is_not_inferred_from_nested_child(self):
        only_child='Rule 606(b)(1). Federal Rule of Evidence 606(b)(1).'
        self.assertFalse(mod.exact_safe_anchor_in_text(only_child,'Rule 606(b)'))
        self.assertFalse(mod.exact_safe_anchor_in_text(only_child,'Federal Rule of Evidence 606(b)'))
        self.assertTrue(mod.exact_safe_anchor_in_text(only_child,'Rule 606(b)(1)'))
        doc=Doc(only_child);log=[]
        with patch.object(mod,'find_text_in_stories',side_effect=find):
            okay=mod.insert_markers(doc,[auth('Federal Rule of Evidence 606(b)'),
                                           auth('Fed. R. Evid. 606(b)(1)')],log,lambda _:None)
        self.assertFalse(okay)
        self.assertEqual(len(doc.Fields.calls),0)
        self.assertTrue(any('UNMATCHED' in x and '606(b)' in x for x in log))

    def test_ambiguous_manual_alias_still_stops(self):
        doc=Doc('Rule 606(b) is cited independently. Fed. R. Evid. 606(b)(1) elsewhere.')
        log=[]
        with patch.object(mod,'find_text_in_stories',side_effect=find):
            okay=mod.insert_markers(doc,[auth('Federal Rule of Evidence 606(b)'),
                                           auth('Fed. R. Evid. 606(b)(1)')],log,lambda _:None,
                                      aliases={'federal rule of evidence 606(b)':['Rule 606(b)'],
                                               'fed. r. evid. 606(b)(1)':['Rule 606(b)']})
        self.assertFalse(okay)
        self.assertEqual(len(doc.Fields.calls),0)
        self.assertTrue(any('AMBIGUOUS OVERLAP' in x for x in log))

if __name__=='__main__':unittest.main()
