"""Regression tests for the V16.4.2 Rule 50 / Rule 50(a) / Rule 50(b) collision."""
import importlib.util
import re
import unittest
from pathlib import Path
from unittest.mock import patch

BASE = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('hipolaw_v1642', BASE / 'generate_legal_tables_v16_4_8.py')
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)


class StoryRange:
    def __init__(self, data, start=0, end=None):
        self.data = data
        self.Start = start
        self.End = len(data) if end is None else end

    @property
    def Text(self):
        return self.data[self.Start:self.End]

    @property
    def Duplicate(self):
        return StoryRange(self.data, self.Start, self.End)

    def SetRange(self, start, end):
        self.Start, self.End = start, end


class FieldCollection:
    def __init__(self):
        self.insertions = []

    def Add(self, rng, field_type, code, preserve):
        self.insertions.append((rng.Start, code))


class MockWordDoc:
    def __init__(self, body, footnote=''):
        self.ranges = {1: StoryRange(body), 2: StoryRange(footnote)}
        self.Fields = FieldCollection()

    def StoryRanges(self, typ):
        return self.ranges[typ]


def literal_find(doc, phrase):
    results=[]
    for typ,name in ((1,'body'),(2,'footnotes')):
        for hit in re.finditer(re.escape(phrase),doc.ranges[typ].data,re.I):
            results.append((name,hit.start(),hit.end()))
    return results


def auth(cite):
    return {'original':cite, 'final':cite, 'category':'Rule'}


class RuleAnchorRegression(unittest.TestCase):
    def test_bare_rule_number_does_not_consume_subsections(self):
        full = ('Rule 50(a) motion, Rule 50(b) motion, Rule 50 motion, '
                'Rule 500 and Rule 50 (b) motion. Rule 50.')
        self.assertTrue(mod.is_bare_rule_variant('Rule 50'))
        self.assertFalse(mod.is_bare_rule_variant('Rule 50(a)'))
        self.assertTrue(mod.exact_safe_anchor_in_text(full,'Rule 50'))
        self.assertFalse(mod.exact_safe_anchor_in_text('Rule 50(a); Rule 50(b)','Rule 50'))
        self.assertFalse(mod.exact_safe_anchor_in_text('Rule 500','Rule 50'))
        self.assertTrue(mod.exact_safe_anchor_in_text('Rule 50 alone','Rule 50'))
        self.assertTrue(mod.disqualifying_rule_tail('(a) motion'))
        self.assertTrue(mod.disqualifying_rule_tail(' (b) motion'))
        self.assertFalse(mod.disqualifying_rule_tail(' motion'))

    def test_distinct_rule_references_can_all_be_marked(self):
        body=('Under Rule 50(a), the court acted. Rule 50(b) allows renewal. '
              'The general Rule 50 standard is discussed. Rule 50(a) again. '
              'Rule 500 is unrelated.')
        doc=MockWordDoc(body)
        authorities=[auth('Fed. R. Civ. P. 50'),auth('Fed. R. Civ. P. 50(a)'),auth('Fed. R. Civ. P. 50(b)')]
        log=[]
        with patch.object(mod,'find_text_in_stories',side_effect=lambda doc,phrase:literal_find(doc,phrase)):
            outcome=mod.insert_markers(doc,authorities,log,lambda _:None)
        self.assertTrue(outcome,'\n'.join(log))
        insertions=doc.Fields.insertions
        self.assertEqual(len(insertions),4)
        expected={body.index('Rule 50(a)'),body.index('Rule 50(b)'),body.index('Rule 50 standard'),body.index('Rule 50(a) again')}
        self.assertEqual({pos for pos,_ in insertions},expected)
        self.assertTrue(any('RULE BOUNDARY: ignored' in ln for ln in log))
        self.assertFalse(any('AMBIGUOUS OVERLAP' in ln for ln in log))

    def test_generic_with_no_standalone_reference_stops_safely(self):
        body='Rule 50(a) and Rule 50(b) only.'
        doc=MockWordDoc(body)
        authorities=[auth('Fed. R. Civ. P. 50'), auth('Fed. R. Civ. P. 50(a)'), auth('Fed. R. Civ. P. 50(b)')]
        log=[]
        with patch.object(mod,'find_text_in_stories',side_effect=lambda doc,phrase:literal_find(doc,phrase)):
            outcome=mod.insert_markers(doc,authorities,log,lambda _:None)
        self.assertFalse(outcome)
        self.assertFalse(doc.Fields.insertions)
        self.assertTrue(any('UNMATCHED' in ln and 'P. 50' in ln for ln in log))

    def test_cross_category_overlaps_still_stop(self):
        body='Rule 50(a) motion'
        doc=MockWordDoc(body)
        authorities=[auth('Fed. R. Civ. P. 50(a)'),auth('Rule 50(a) motion')]
        log=[]
        with patch.object(mod,'find_text_in_stories',side_effect=lambda doc,phrase:literal_find(doc,phrase)):
            outcome=mod.insert_markers(doc,authorities,log,lambda _:None)
        self.assertFalse(outcome)
        self.assertTrue(any('AMBIGUOUS OVERLAP' in ln for ln in log))
        self.assertFalse(doc.Fields.insertions)


if __name__=='__main__':
    unittest.main()
