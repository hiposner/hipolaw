"""Hipo Law Generator v16.4.6 (based on V15.1) (Windows + installed Microsoft Word).

Input: ORIGINAL .docx and approved V15-compatible .xlsx (Authorities, Contents, Setup, Aliases).
Creates a NEW editable .docx + matching .pdf + one audit log. Never writes the inputs.
IMPORTANT: Preflight refuses ambiguous headings, unapproved rows, and complex
existing section geometry. No automated program can guarantee all Word layouts.
Human legal-citation review is always necessary.

Run: py -u generate_legal_tables_v16_4_6.py
Requirements: py -m pip install pywin32 openpyxl pymupdf
"""
from __future__ import annotations
import re, shutil, traceback, sys, time, os, zipfile, tempfile, json, html, statistics, argparse, subprocess, hashlib, webbrowser
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import Tk, filedialog, ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
from collections import defaultdict
# Import Windows-only integrations only when running full Word automation;
# --preflight works on any operating system and never opens Word.
try:
    import openpyxl
except ImportError:
    openpyxl = None
try:
    import win32com.client as win32
except ImportError:
    win32 = None
try:
    import pymupdf
except ImportError:
    pymupdf = None

SECTION_NEXT_PAGE = 2
PAGE_BREAK = 7
PRIMARY = 1
LOWER_ROMAN = 2
ARABIC = 0
COLLAPSE_END = 0
DOCX = 16
PDF = 17
PAGE_COUNT = 2
ADJUSTED_PAGE = 1
PHYSICAL_PAGE = 3
STOP = 0
WD_FIELD_EMPTY = -1


# HipoLaw styling is deliberately NOT compiled into the EXE. Adjust the JSON
# next to HipoLaw_v16_4_6.exe and run the app again; no rebuild needed.
STYLE_FILE_NAME = 'HipoLaw_Style_Settings.json'
DEFAULT_STYLE = {
    'toc_title': {'centered': True, 'bold': True},
    'toa_title': {'centered': True, 'bold': True},
    'category_headings': {'uppercase': True, 'bold': True},
    'toc_entries': {'bold': False},
    'entry_font_size_pt': 12.0,
    'heading_font_size_pt': 12.0,
    'pleading_grid': {
        'line_spacing_pt': 24.0,
        # Shift printed numbers on generated BODY pages only; calibrated against
        # the standard pleading-paper template. Adjust for other page templates.
        'body_header_shift_pt': 9.5,
        # Lower case entries by 6pt compared with V16 to align them to the
        # actual 28-line grid and reduce raised-text clipping in Word.
        'vertical_raise_pt': {
            'toc': 9.5, 'cases': 3.5, 'statutes': 12.0,
            'other': 12.0, 'rules': 14.0,
        },
    },
}


def style_file_path():
    # sys.executable points to the user's .exe when frozen by PyInstaller.
    base = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
    return base / STYLE_FILE_NAME


def _validate_style(data):
    if not isinstance(data, dict):
        raise ValueError('Style settings must be one JSON object.')
    def strict_group(name, allowed):
        obj = data.get(name)
        if not isinstance(obj, dict) or set(obj) != set(allowed):
            raise ValueError(f'Style settings {name!r} must contain exactly: {", ".join(allowed)}')
        for k in allowed:
            if not isinstance(obj[k], bool):
                raise ValueError(f'Style settings {name}.{k} must be true or false')
    strict_group('toc_title', ('centered', 'bold'))
    strict_group('toa_title', ('centered', 'bold'))
    strict_group('category_headings', ('uppercase', 'bold'))
    strict_group('toc_entries', ('bold',))
    for k in ('entry_font_size_pt', 'heading_font_size_pt'):
        v = data.get(k)
        if not isinstance(v, (float, int)) or isinstance(v, bool) or not 10 <= v <= 14:
            raise ValueError(f'{k} must be a number between 10 and 14 points')
    grid = data.get('pleading_grid')
    if not isinstance(grid, dict) or not {'line_spacing_pt','vertical_raise_pt'} <= set(grid) or \
       not set(grid) <= {'line_spacing_pt','vertical_raise_pt','body_header_shift_pt'}:
        raise ValueError('pleading_grid requires line_spacing_pt and vertical_raise_pt; optional body_header_shift_pt')
    shift=grid.get('body_header_shift_pt',9.5)
    if not isinstance(shift,(int,float)) or isinstance(shift,bool) or not -18 <= shift <= 18:
        raise ValueError('pleading_grid.body_header_shift_pt must be between -18 and 18 points')
    spacing=grid['line_spacing_pt']
    if not isinstance(spacing, (int,float)) or isinstance(spacing,bool) or not 22 <= spacing <= 26:
        raise ValueError('line_spacing_pt must be between 22 and 26 points')
    pos=grid['vertical_raise_pt']
    if not isinstance(pos,dict) or set(pos)!={'toc','cases','statutes','other','rules'}:
        raise ValueError('vertical_raise_pt requires toc, cases, statutes, other, rules')
    for k,v in pos.items():
        if not isinstance(v, (int,float)) or isinstance(v, bool) or not -4 <= v <= 18:
            raise ValueError(f'vertical_raise_pt.{k} must be between -4 and 18 points')
    if set(data)!=set(DEFAULT_STYLE):
        raise ValueError('Unrecognized or missing style settings keys: '+repr(set(data)^set(DEFAULT_STYLE)))
    return data


def load_style_settings():
    path=style_file_path()
    if not path.exists():
        raise FileNotFoundError(f'Missing style settings file: {path}. Restore it from the V16.4.6 build ZIP.')
    try:
        data=json.loads(path.read_text(encoding='utf-8-sig'))
    except json.JSONDecodeError as exc:
        raise ValueError(f'Invalid JSON in {path.name}, line {exc.lineno}: {exc.msg}') from exc
    return _validate_style(data)


def pick(title, ext):
    root = Tk(); root.withdraw(); root.attributes('-topmost', True)
    path = filedialog.askopenfilename(title=title, filetypes=[(ext.upper()+' files', '*.'+ext)])
    root.destroy()
    return Path(path) if path else None


def read_sheet(sheet, label):
    """Strict seven-column approval contract (all blank decisions block filing)."""
    expected_a = 'Citation in motion' if label == 'Authorities' else 'Heading in motion'
    headers = list(next(sheet.values, []))
    expected_cols=([expected_a,'AI suggestion','Why','Decision','Your edit (if Edit)',
                    'FINAL entry used' if label=='Authorities' else 'FINAL heading used',
                    'Type' if label=='Authorities' else 'Level'])
    if [str(x or '').strip() for x in headers[:7]]!=expected_cols:
        raise ValueError(f'{label}: the first 7 headers must match the V15 template exactly.')
    rows, pending, seen = [], [], set()
    for number, row in enumerate(list(sheet.values)[1:], 2):
        if not row or not str(row[0] or '').strip(): continue
        original, suggestion, why, decision, edited, cached, category = (list(row)+[None]*7)[:7]
        original = str(original).strip()
        if original.casefold() in seen: raise ValueError(f'{label} row {number}: duplicate source {original!r}')
        seen.add(original.casefold())
        decision = str(decision or '').strip().casefold()
        if decision not in ('accept','edit','reject'):
            pending.append(f'row {number}: {original[:75]}')
            continue
        if decision == 'reject':continue
        if decision == 'edit':
            if not str(edited or '').strip():
                raise ValueError(f'{label} row {number}: Edit selected without replacement text')
            final=str(edited).strip()
        else:
            final=str(suggestion or original).strip()
        if not final:raise ValueError(f'{label} row {number}: empty final display entry')
        typ=str(category or '').strip()
        if label=='Authorities' and typ.lower() not in ('case','statute','rule','court rule','evidence rule','procedural rule','regulation','constitution','other'):
            raise ValueError(f'Authorities row {number}: Type {typ!r} must be Case/Statute/Rule/Regulation/Constitution/Other')
        if label=='Contents':
            try: level=int(typ)
            except ValueError: raise ValueError(f'Contents row {number}: invalid heading level {typ!r}')
            if level not in (1,2,3):raise ValueError(f'Contents row {number}: level must be 1–3')
            typ=str(level)
        rows.append(dict(original=original,final=final,category=typ,reason=str(why or '').strip()))
    if pending:
        raise ValueError(f'{label}: {len(pending)} undecided rows. Choose Accept, Edit, or Reject before running:\n'+'\n'.join(pending[:30]))
    if not rows:raise ValueError(f'{label}: no approved entries')
    return rows


def normalized_heading(s):
    s = str(s).replace('\r',' ').replace('\x07',' ').strip()
    s=re.sub(r'\s+',' ',s)
    return s.casefold()


def remove_terminal_heading_period(value):
    """Only the TOC display omits a final period, including before close quotes.

    The source Word paragraph and approved spreadsheet text are NOT changed.
    """
    label = re.sub(r'\s+', ' ', str(value)).strip()
    return re.sub(r'\.([”"\'’]*)$', r'\1', label)


def toc_display_label(number, heading):
    """Keep the source paragraph's list identifier but omit its final dot.

    No numbering is generated from TOC position: the actual Word ListString is
    captured BEFORE introducing any preliminary pages.
    """
    number = re.sub(r'\s+', ' ', str(number or '')).strip()
    number = re.sub(r'[.)]$', '', number).strip()
    label = remove_terminal_heading_period(heading)
    if number and not re.match(r'^' + re.escape(number) + r'(?=[.\s)]|$)', label, re.I):
        return number + ' ' + label
    return label


def original_heading_number(range_obj):
    """Read the real numbered heading; this is not the workbook's typed text."""
    paragraph = range_obj.Paragraphs(1)
    raw = paragraph.Range.Text.rstrip('\r\x07').strip()
    try:
        automatic = str(paragraph.Range.ListFormat.ListString or '').strip()
    except Exception:
        automatic = ''
    if automatic:
        return automatic
    # For source motions with typed prefixes (rather than Word list fields).
    m = re.match(r'^((?:[IVXLCDM]+|[ivxlcdm]+|[A-Za-z]|\d+)[.)])\s+', raw)
    return m.group(1) if m else ''


def capture_heading_numbering(doc, headings):
    """Snapshot each source heading's actual Word numbering and visible text."""
    snapshot = []
    for h in headings:
        found = heading_ranges(doc, h['original'])
        if len(found) != 1:
            raise RuntimeError(f'Heading must have one Word paragraph: {h["original"]!r}; got {len(found)}')
        number = original_heading_number(found[0])
        if not number:
            raise RuntimeError(f'Heading lacks a Word/typed list number: {h["original"]!r}. '
                               'The TOC must preserve original heading numbering; fix source or review manually.')
        snapshot.append((h['original'], number, int(h['category'])))
    return snapshot


def check_heading_numbering(doc, heading_snapshot, context):
    """Abort if a Word update has advanced or otherwise changed motion numerals."""
    issues = []
    for name, expected, _level in heading_snapshot:
        found = heading_ranges(doc, name)
        actual = original_heading_number(found[0]) if len(found) == 1 else '<missing>'
        if actual != expected:
            issues.append(f'{name!r}: original {expected!r}, now {actual!r}')
    if issues:
        raise RuntimeError('SOURCE MOTION NUMBERING CHANGED during ' + context +
                           '. Generation stopped without exporting a filing document. ' +
                           '; '.join(issues[:14]))


def heading_ranges(doc, name):
    """Match title text as typed, or text without its Word auto-numbering prefix."""
    want=normalized_heading(name)
    # Word ListFormat can supply numbering not contained in Range.Text.
    candidates={want,re.sub(r'^(?:[ivxlcdm]+|\d+)[.)]\s*','',want,flags=re.I)}
    found=[]
    for p in doc.Paragraphs:
        raw=p.Range.Text.strip('\r\x07').strip()
        t=normalized_heading(raw)
        variants={t, re.sub(r'^(?:[ivxlcdm]+|[a-z]|\d+)[.)]\s*', '', t, flags=re.I)}
        try:
            prefix=str(p.Range.ListFormat.ListString or '').strip()
            if prefix: variants.add(normalized_heading(prefix+' '+raw))
        except Exception:pass
        if candidates.intersection(variants):found.append(p.Range.Duplicate)
    return found


def escape_field_text(s):
    """Keep existing typographic punctuation in approved text.

    ASCII quotation marks used INSIDE a Word field's quoted argument must
    not be mistaken for the field's delimiters. Convert only ASCII quotes
    to typographic opening/closing double quotes; preserve existing smart quotes.
    """
    value = str(s).replace('\r', ' ').replace('\n', ' ')
    out, is_open = [], True
    for ch in value:
        if ch == '"':
            out.append('“' if is_open else '”')
            is_open = not is_open
        else:
            out.append(ch)
    return ''.join(out)


# Word's TC/TA quoted field arguments interpret even typographic quotes as
# delimiters in certain installations. Use plain, non-quote internal tokens
# and restore exact punctuation in the *visible result* after field updates.
def safe_tc_field_label(value):
    return (escape_field_text(value)
            .replace('“', 'QOPENQ').replace('”', 'QCLOSEQ')
            .replace('‘', 'SQOPENQ').replace('’', 'SQCLOSEQ'))


def safe_ta_sort_label(value):
    """Word lexical sorting: zero-pad one-digit numeric parentheses only.
    The visible result is restored to original by XML patch, without touching fields.
    """
    return re.sub(r'\((\d)\)',r'(0\1)',str(value))


def add_field(doc, loc, code):
    r=doc.Range(loc,loc)
    return doc.Fields.Add(r, WD_FIELD_EMPTY, code, False)


def find_matches(doc, phrase, story):
    """Find all non-overlapping literal matches in a Word story; returns COM range coordinates."""
    if not phrase:return []
    work=story.Duplicate; work.Start=story.Start; work.End=story.End
    f=work.Find; f.ClearFormatting(); f.Text=phrase
    f.Forward=True; f.Wrap=STOP; f.MatchCase=False; f.MatchWholeWord=False
    hits=[]; cursor=work.Start; limit=story.End
    while cursor < limit:
        work.SetRange(cursor,limit)
        if not work.Find.Execute(FindText=phrase, MatchCase=False, MatchWholeWord=False, Forward=True, Wrap=STOP):
            break
        start,end=work.Start,work.End
        hits.append((start,end))
        cursor=max(end,start+1)
        if len(hits)>300:raise RuntimeError('Unexpected excessive repeated matches')
    return hits


def list_story_ranges(doc):
    # Word 1 = body, 2 = footnotes. These are actual Word story ranges (not page text).
    stories=[]
    for typ,name in ((1,'body'),(2,'footnotes')):
        try:
            r=doc.StoryRanges(typ)
            if r is not None: stories.append((name,r.Duplicate))
        except Exception: pass
    return stories


def find_text_in_stories(doc, phrase):
    result=[]
    for name,story in list_story_ranges(doc):
        for start,end in find_matches(doc,phrase,story):
            result.append((name,start,end))
    return result


def category_code(s):
    kind=str(s).strip().lower()
    return {'case':1,'statute':2,'constitution':3,'regulation':3,'other':3,'rule':4,'court rule':4,'evidence rule':4,'procedural rule':4}[kind]


def variants_for_authority(item):
    """Conservative document anchors, without any named-case exceptions.

    The optional Aliases sheet supplies exact additional appearances when AI
    resolves a short form. Never mark bare 'Id.' automatically.
    """
    raw=item['original']
    group=category_code(item['category'])
    if group==1:
        # Split before a standard reporter citation, not at the first comma;
        # party names themselves may contain commas.
        title=re.split(r',\s*(?=\d{1,4}\s+(?:U\.S\.|S\.Ct\.|F\.|P\.|N\.E\.|Cal\.|WL\b|LEXIS\b))',raw, maxsplit=1,flags=re.I)[0]
        results=[raw,title]
        # Reporter-qualified short forms are supplied by ChatGPT in Aliases,
        # avoiding false attribution when two cases have similar short names.
        return list(dict.fromkeys(x for x in results if x and len(x)>=5))
    if group==2:
        # Exact statute citation plus normalized section form.
        m=re.search(r'(?i)\b(\d+)\s+U\.S\.C\.\s*§+\s*([\w.-]+)',raw)
        if m:return [f'{m.group(1)} U.S.C. § {m.group(2)}',raw]
    if group==4:
        # Keep the COMPLETE subdivision chain. Capturing only one parenthetical
        # would turn a citation to Rule N(a)(1) into a false Rule N(a) alias.
        m=re.search(r'(?i)(?:Fed\.\s*R\.\s*(Civ\.\s*P|Crim\.\s*P|Evid\.|App\.\s*P)\.?\s*|Rule\s+)(\d+)((?:\([a-z0-9]+\))*)',raw)
        if m:
            num=m.group(2); suffix=m.group(3) or ''
            variants=[raw,f'Rule {num}{suffix}']
            if 'Evid' in raw:
                variants += [f'Fed. R. Evid. {num}{suffix}',f'Federal Rule of Evidence {num}{suffix}']
            if 'Civ' in raw:
                variants += [f'Fed. R. Civ. P. {num}{suffix}',f'Federal Rule of Civil Procedure {num}{suffix}']
            return list(dict.fromkeys(variants))
    return [raw]


def is_rule_variant(phrase):
    """Recognize complete numbered rules, including arbitrary subdivision depth.

    Works for Word-text preflight and for generator-produced shorthand variants.
    This is a GENERIC boundary check: no particular rule or case is special.
    """
    phrase = str(phrase).strip()
    return bool(re.fullmatch(
        r'(?:Rule\s+|Fed\.\s*R\.\s*(?:Civ\.\s*P\.|Crim\.\s*P\.|Evid\.|App\.\s*P\.)\s*|'
        r'Federal\s+Rule\s+of\s+(?:Evidence|Civil\s+Procedure|Criminal\s+Procedure|Appellate\s+Procedure)\s+)'
        r'\d+(?:\([a-z0-9]+\))*',
        phrase, flags=re.I))


def is_bare_rule_variant(phrase):
    """Compatibility helper retained for unsuffixed rule tests."""
    return is_rule_variant(phrase) and '(' not in str(phrase)


def disqualifying_rule_tail(tail):
    """A rule anchor cannot consume a further subdivision or longer number."""
    return bool(re.match(r'[A-Za-z0-9_]|[ \t]*\([A-Za-z0-9]{1,3}\)', str(tail)))


def exact_safe_anchor_in_text(text, phrase):
    """Preflight must not count a parent rule only within a child reference."""
    if not is_rule_variant(phrase):
        return phrase.casefold() in text.casefold()
    for hit in re.finditer(re.escape(phrase), text, flags=re.I):
        if not disqualifying_rule_tail(text[hit.end():hit.end()+14]):
            return True
    return False


def matched_rule_is_independent(doc, story_name, end, phrase):
    """Check Word text after ANY rule level, including nested subdivisions."""
    if not is_rule_variant(phrase):
        return True
    story = doc.StoryRanges(1 if story_name == 'body' else 2)
    if end >= story.End:
        return True
    lookahead = story.Duplicate
    lookahead.SetRange(end, min(story.End, end + 14))
    return not disqualifying_rule_tail(lookahead.Text)


def context_qualified_evidence_matches(doc, item):
    """Find abbreviated evidence rule numbers only inside a verified compound citation.

    E.g. Fed. R. Evid. 801(c), 802; Rule 803(7) or 803(10).
    Do not search globally for 802 / 803(10), since those numbers can occur in
    transcript citations or unrelated places.
    """
    original = item['original']
    patterns = []
    if re.search(r'Fed\.\s*R\.\s*Evid\.\s*802\b', original, re.I):
        patterns = [('801(c), 802', '802'), ('801(c) and 802', '802')]
    elif re.search(r'Fed\.\s*R\.\s*Evid\.\s*803\(10\)', original, re.I):
        patterns = [('803(7) or 803(10)', '803(10)'),
                    ('803(7), 803(10)', '803(10)')]
    found = []
    for full_phrase, suffix in patterns:
        for story, start, end in find_text_in_stories(doc, full_phrase):
            # A Word Find span is native Word character positions. Subtracting
            # the suffix length is safe here because the suffix is plain ASCII.
            found.append((story, end - len(suffix), end, suffix))
    return found


def insert_markers(doc, authorities, log, progress, aliases=None):
    """Mark safe literal occurrences; reject competing/overlapping anchors."""
    aliases=aliases or {}
    markers=[]; misses=[]; all_claims=[]
    for idx,item in enumerate(authorities):
        progress(f"Matching authority {idx+1}/{len(authorities)}: {item['original'][:65]}")
        candidates=[]
        variants=list(dict.fromkeys(variants_for_authority(item)+aliases.get(item['original'].casefold(),[])))
        suppressed_prefix_hits = 0
        for phrase in variants:
            if phrase.strip().lower() in ('id.','ibid.','supra'):
                misses.append('Unsafe bare short form in Aliases: '+item['original']);continue
            for story,start,end in find_text_in_stories(doc,phrase):
                if not matched_rule_is_independent(doc,story,end,phrase):
                    suppressed_prefix_hits += 1
                    continue
                candidates.append((story,start,end,phrase))
        if suppressed_prefix_hits:
            log.append(f'RULE BOUNDARY: ignored {suppressed_prefix_hits} parent-prefix hit(s) '
                       f'within subsection citations for {item["original"]!r}')
        candidates += context_qualified_evidence_matches(doc, item)
        # A full cite and its contained case name are ONE occurrence: prefer
        # longest, preventing duplicate invisible TA fields in one citation.
        chosen=[]
        for hit in sorted(candidates,key=lambda x:(-(x[2]-x[1]),x[1])):
            if not any(hit[0]==prev[0] and hit[1]<prev[2] and hit[2]>prev[1] for prev in chosen):
                chosen.append(hit)
        if not chosen:misses.append(item['original'])
        else:
            markers += [dict(story=a,start=b,end=c,item=idx,phrase=d) for a,b,c,d in chosen]
            log.append(f'AUTH {idx+1}: {len(chosen)} anchored match(es): {item["original"]}')
    all_markers=sorted(markers,key=lambda x:(x['story'],x['start'],x['end']))
    overlaps=[]
    for i,m in enumerate(all_markers):
        for other in all_markers[i+1:]:
            if m['story']!=other['story']:break
            if other['start']>=m['end']:break
            if m['item']!=other['item'] and other['start']<m['end']:
                overlaps.append((m['phrase'],other['phrase'],
                                 authorities[m['item']]['original'],
                                 authorities[other['item']]['original']))
    if misses or overlaps:
        log.append('STOP: unresolved document anchors; nothing was inserted')
        log += ['  UNMATCHED: '+x for x in misses]
        log += ['  AMBIGUOUS OVERLAP: '+repr(x) for x in overlaps[:20]]
        return False
    progress(f"Matches found: {len(markers)}. Inserting native Word TA fields...")
    for n,m in enumerate(sorted(markers,key=lambda x:(x['story'],x['start']),reverse=True),1):
        progress(f"Inserting marker {n}/{len(markers)}: {m['phrase'][:55]}")
        it=authorities[m['item']]
        code=f'TA \\l "{escape_field_text(safe_ta_sort_label(it["final"]))}" \\c {category_code(it["category"])}'
        r=doc.StoryRanges(1 if m['story']=='body' else 2).Duplicate
        # Mark at the BEGINNING of the reference. A citation split across two
        # printed pages belongs to the page where the citation begins.
        r.SetRange(m['start'],m['start'])
        doc.Fields.Add(r, WD_FIELD_EMPTY, code, False)
    log.append(f'TOA: inserted {len(markers)} markers across {len(authorities)} approved authorities; start-page anchoring')
    log.append('Id. is not marked automatically; use citation coverage review and attorney approval.')
    return True


def replace_placeholder_with_field(doc, text, field_code, progress, preliminary=2):
    """Replace an exact placeholder paragraph without calling Word Find.Execute.

    Word's global Find can stall with many hidden TC/TA fields. We inspect only
    paragraphs in the preliminary section, then replace the exact text range.
    """
    progress(f"  {text}: scanning preliminary-section paragraphs (no Word Find)")
    matches = []
    section = doc.Sections(preliminary)
    for paragraph in section.Range.Paragraphs:
        candidate = paragraph.Range.Text.rstrip('\r\x07').strip()
        if candidate == text:
            matches.append((paragraph.Range.Start, paragraph.Range.End))
    if len(matches) != 1:
        raise RuntimeError(f'{text}: expected exactly one placeholder paragraph; found {len(matches)}')
    start, end = matches[0]
    progress(f"  {text}: unique placeholder at {start}-{end}")
    # Exclude the paragraph mark, preserving the surrounding layout.
    text_range = doc.Range(start, end - 1)
    text_range.Text = ''
    progress(f"  {text}: placeholder removed; inserting field")
    doc.Fields.Add(doc.Range(start, start), WD_FIELD_EMPTY, field_code, False)
    progress(f"  {text}: field inserted")


def make_prelim_text(authorities, style=None):
    style = style or DEFAULT_STYLE
    parts=['TABLE OF CONTENTS\r[GENERATE_TOC]\r\x0cTABLE OF AUTHORITIES\r']
    categories=[(1,'Cases','[GENERATE_CASES]'),(2,'Statutes','[GENERATE_STATUTES]'),
                (3,'Other Authorities','[GENERATE_OTHER]'),(4,'Rules','[GENERATE_RULES]')]
    used={category_code(a['category']) for a in authorities}
    for code,label,placeholder in categories:
        if code in used:
            if style['category_headings']['uppercase']:
                label=label.upper()
            parts.append(label+'\r'+placeholder+'\r')
    return ''.join(parts)


def protect_preliminary_paragraphs(doc, preliminary, body_heading, log):
    """Remove inherited motion list numbering from NEW table paragraphs only.

    V15's InsertBefore occurs at a numbered ListParagraph. Word can inherit the
    original heading's numId on inserted text, which would advance I/II/III...
    to X/XI/XII... . RemoveNumbers + Normal styling on PRELIM paragraphs alone
    prevents this; original body and caption are explicitly outside the range.
    """
    anchor = heading_ranges(doc, body_heading)
    if len(anchor) != 1:
        raise RuntimeError('Cannot protect TOC/TOA without unique original first heading')
    body_start = anchor[0].Start
    changed = 0
    for paragraph in list(doc.Sections(preliminary).Range.Paragraphs):
        rng = paragraph.Range
        if rng.Start >= body_start or rng.End > body_start:
            continue  # never touch even part of the first original body heading
        try:
            rng.ListFormat.RemoveNumbers()
        except Exception as exc:
            raise RuntimeError('Cannot safely isolate newly inserted TOC/TOA from motion numbering: ' + str(exc))
        # V15 inherited ListParagraph style (which can contain numbering).
        # Set only generated preliminary paragraphs to Word's built-in Normal.
        rng.Style = -1  # wdStyleNormal
        changed += 1
    log.append(f'NUMBERING ISOLATION: {changed} preliminary paragraphs made unnumbered '
               '(source motion paragraphs never edited)')
    if changed < 4:
        raise RuntimeError('Too few preliminary paragraphs were isolated; refusing to risk numbered source motion')


def insert_preliminary_sections(doc, log, body_heading, authorities, style=None):
    if doc.Sections.Count!=1:
        raise ValueError('This release supports original one-section motions only; document has '
                         f'{doc.Sections.Count} sections. No source modification performed.')
    headings=heading_ranges(doc,body_heading)
    if len(headings)!=1:raise ValueError(f'Cannot uniquely find first approved body heading: {body_heading!r}')
    start=headings[0].Start
    # An empty caption would create a useless blank first page. Support motions
    # that start immediately with the first body heading (two sections).
    empty_cover=start <= doc.Content.Start+2
    if not empty_cover:
        doc.Range(start,start).InsertBreak(SECTION_NEXT_PAGE)
    again=heading_ranges(doc,body_heading)[0]
    doc.Range(again.Start,again.Start).InsertBefore(make_prelim_text(authorities, style))
    again=heading_ranges(doc,body_heading)[0]
    doc.Range(again.Start,again.Start).InsertBreak(SECTION_NEXT_PAGE)
    expected=2 if empty_cover else 3
    if doc.Sections.Count!=expected:
        raise RuntimeError(f'Expected {expected} sections; got {doc.Sections.Count}')
    preliminary=expected-1
    protect_preliminary_paragraphs(doc, preliminary, body_heading, log)
    for index,fmt in ((preliminary,LOWER_ROMAN),(expected,ARABIC)):
        sec=doc.Sections(index)
        sec.PageSetup.DifferentFirstPageHeaderFooter=False
        footer=sec.Footers(PRIMARY);footer.LinkToPrevious=False
        footer.PageNumbers.NumberStyle=fmt
        footer.PageNumbers.RestartNumberingAtSection=True
        footer.PageNumbers.StartingNumber=1
        # Most pleadings have native PAGE fields already. For motions that do
        # not, add one in a new centered footer paragraph rather than failing.
        if not any(f.Code.Text.strip().upper().startswith('PAGE') for f in footer.Range.Fields):
            fr=footer.Range
            fr.InsertAfter('\r')
            fr=footer.Range
            fr.ParagraphFormat.Alignment=1
            pos=fr.End-1
            doc.Fields.Add(doc.Range(pos,pos),WD_FIELD_EMPTY,'PAGE',False)
            log.append(f'Created missing PAGE field in footer of section {index}')
    log.append(f'Inserted preliminary and body sections. First body heading: {body_heading!r}; '
               f'cover={not empty_cover}; {expected} sections.')
    return preliminary,expected


# --- Formatting and post-export validation ----------------------------------
MIN_TABLE_FONT_PT = 12.0
WD_LINE_SPACE_EXACTLY = 4
# Word can preserve existing pleading-paper geometry, but there is no universal
# line-number template. 24-point grid + baseline values apply ONLY when the
# source visibly contains the customary 28-number pleading-paper header.
PLEADING_GRID_PT=24.0
PLEADING_GRID_TOLERANCE_PT=4.0
GRID_TEXT_RAISE_PT={'toc':9.5,'cases':9.5,'statutes':12.0,'other':12.0,'rules':14.0}


def has_numbered_pleading_header(docx_path):
    """Inspect source Word headers, not client name or document title."""
    with zipfile.ZipFile(docx_path) as archive:
        for name in archive.namelist():
            if not re.fullmatch(r'word/header\d+\.xml',name): continue
            data=archive.read(name).decode('utf-8')
            values={int(html.unescape(x)) for x in re.findall(r'<w:t(?:\s+[^>]*)?>(\d{1,2})</w:t>',data)
                    if int(x)<=28}
            if len(values.intersection(set(range(1,29))))>=24:
                return True
    return False


def standardize_table_results(doc, log, progress, numbered_grid, preliminary, style=None):
    """Format ONLY inserted preliminary TOC/TOA elements; never the motion body."""
    style=style or DEFAULT_STYLE
    progress('Formatting generated tables (source-specific pleading grid only when detected)')
    fields=[]
    for f in list(doc.Fields):
        code=f.Code.Text.strip().upper()
        if code.startswith('TOA ') or code.startswith('TOC '):fields.append((code,f))
    if len(fields)<2:raise RuntimeError('Expected native TOC and TOA fields')
    categories={'1':'cases','2':'statutes','3':'other','4':'rules'}
    spacing=style['pleading_grid']['line_spacing_pt']
    raises=style['pleading_grid']['vertical_raise_pt']
    for code,f in fields:
        result=f.Result
        result.Font.Size=style['entry_font_size_pt']
        pf=result.ParagraphFormat
        pf.SpaceBefore=0;pf.SpaceAfter=0
        if numbered_grid:
            pf.LineSpacingRule=WD_LINE_SPACE_EXACTLY
            pf.LineSpacing=spacing
        if code.startswith('TOC '):
            group='toc'
            result.Font.Bold=style['toc_entries']['bold']
        else:
            m=re.search(r'\\c\s+([1234])(?:\s|$)',code,re.I)
            if not m:raise RuntimeError('Unknown TOA category: '+code)
            group=categories[m.group(1)]
        result.Font.Position=raises[group] if numbered_grid else 0
        if code.startswith('TOA '):result.Font.Italic=False
    # Heading formatting is done on isolated PRELIMINARY paragraphs only.
    # Printed category labels are NOT the keys in vertical_raise_pt.
    # In particular, "Other Authorities" must use the "other" baseline;
    # lowercasing the display text would produce the invalid key
    # "other authorities" and crash after all TA markers were inserted.
    category_groups=(('Cases','cases'),('Statutes','statutes'),
                     ('Other Authorities','other'),('Rules','rules'))
    category_names={(label.upper() if style['category_headings']['uppercase'] else label): group
                    for label,group in category_groups}
    headings={'TABLE OF CONTENTS':'toc','TABLE OF AUTHORITIES':'toc',**category_names}
    unknown_groups=set(headings.values())-set(raises)
    if unknown_groups:
        raise RuntimeError('Table heading styles have no configured baseline: '
                           +', '.join(sorted(unknown_groups)))
    heading_count=0
    for para in list(doc.Sections(preliminary).Range.Paragraphs):
        key=para.Range.Text.rstrip('\r\x07').strip()
        if key not in headings:continue
        group=headings[key]
        fmt=para.Range.ParagraphFormat
        font=para.Range.Font
        font.Position=(raises[group] if numbered_grid else 0)
        font.Size=style['heading_font_size_pt']
        font.Italic=False
        fmt.SpaceBefore=0;fmt.SpaceAfter=0
        if key=='TABLE OF CONTENTS':
            fmt.Alignment=1 if style['toc_title']['centered'] else 0
            font.Bold=style['toc_title']['bold']
        elif key=='TABLE OF AUTHORITIES':
            fmt.Alignment=1 if style['toa_title']['centered'] else 0
            font.Bold=style['toa_title']['bold']
        else:
            fmt.Alignment=0
            font.Bold=style['category_headings']['bold']
        heading_count+=1
    if heading_count<3:
        raise RuntimeError(f'Expected at least three TOC/TOA/category headings, found {heading_count}')
    case_fields=[f for code,f in fields if code.startswith('TOA ') and re.search(r'\\c\s+1(?:\s|$)',code,re.I)]
    case_count=0
    if case_fields:
        if len(case_fields)!=1:raise RuntimeError('Expected one Case TOA field')
        reporter=re.compile(r',?\s+\d+\s+(?:U\.S\.|S\.Ct\.|F\.\s*(?:2d|3d|4th|Supp\.)|\w+\s+App\.|\bWL\b|LEXIS\b)',re.I)
        for para in list(case_fields[0].Result.Paragraphs):
            text=para.Range.Text.rstrip('\r\x07')
            match=reporter.search(text)
            if not match:continue
            segment=para.Range.Duplicate
            segment.SetRange(para.Range.Start,para.Range.Start+match.start())
            if segment.Start<segment.End:
                segment.Font.Italic=True;case_count+=1
    log.append(f'TABLE TYPOGRAPHY: entries {style["entry_font_size_pt"]}pt; '
               f'titles centered/bold; categories uppercase/bold; '
               f'pleading grid={numbered_grid}; italicized {case_count} case names')
    log.append('TABLE GRID BASELINES: '+str(raises))


def _numeric_evidence_key(line):
    m=re.search(r'Fed\.\s*R\.\s*Evid\.\s*(\d+)(?:\((\d+)\))?',line,re.I)
    if m:
        return (int(m.group(1)), int(m.group(2)) if m.group(2) is not None else -1)
    return None


def _restore_cached_table_text_in_docx(path, log, progress, authorities, headings, numbered_grid):
    """Restore visible TOC punctuation and numeric labels without Word Find."""
    progress('Restoring Word field display text offline; no Word Find calls')
    target=Path(path).resolve()
    if not zipfile.is_zipfile(target):raise RuntimeError('Not a valid DOCX checkpoint')
    with zipfile.ZipFile(target,'r') as src:
        xml=src.read('word/document.xml').decode('utf-8')
    expected={}
    for h in headings:
        safe=safe_tc_field_label(h['toc_label'])
        for token in ('QOPENQ','QCLOSEQ','SQOPENQ','SQCLOSEQ'):
            expected[token]=expected.get(token,0)+safe.count(token)
    # XML text is split into runs; do not touch w:instrText fields.
    counts={token:0 for token in expected}
    counts['NUMERIC_PAD']=0
    def patch_tag(m):
        content=m.group(2)
        for old,new in [('QOPENQ','“'),('QCLOSEQ','”'),('SQOPENQ','‘'),('SQCLOSEQ','’')]:
            n=content.count(old)
            counts[old]=counts.get(old,0)+n
            content=content.replace(old,new)
        # If full display label contains a numeric subrule '(07)', restore
        # only inside TOA entries by direct string substitution. The XML
        # alteration leaves field instructions and all native markers intact.
        for a in authorities:
            final=a['final']
            for numeric in re.findall(r'\((\d)\)', final):
                old='('+numeric.zfill(2)+')'
                new='('+numeric+')'
                n=content.count(old)
                if n: content=content.replace(old,new);counts['NUMERIC_PAD']+=n
        return m.group(1)+content+m.group(3)
    updated=re.sub(r'(<w:t(?:\s+[^>]*)?>)(.*?)(</w:t>)',patch_tag,xml,flags=re.S)
    for token,n in expected.items():
        if counts.get(token,0)!=n:
            raise RuntimeError(f'Expected {n} visible {token} tokens, found {counts.get(token,0)}; no edit saved')
    # No automatic page-break cosmetics in V15: documents with different
    # case counts, page sizes or heading styles may need different splits.
    log.append('TOA CONTINUATION: natural Word page flow only; no continued heading')
    fd,tmp=tempfile.mkstemp(prefix='.hipolaw_v16_',suffix='.docx',dir=str(target.parent))
    os.close(fd)
    try:
        with zipfile.ZipFile(target,'r') as src,zipfile.ZipFile(tmp,'w') as dst:
            for entry in src.infolist():
                dst.writestr(entry,updated.encode('utf-8') if entry.filename=='word/document.xml' else src.read(entry.filename))
        with zipfile.ZipFile(tmp,'r') as check:
            if check.testzip() is not None:raise RuntimeError('Generated DOCX ZIP check failed')
        os.replace(tmp,target)
    finally:
        if os.path.exists(tmp):os.unlink(tmp)
    log.append(f'VISIBLE RESULTS RESTORED: {counts} (native Word fields preserved)')


def verify_table_field_results(doc, log, progress, headings, authorities):
    progress('Checking visible TOC headings and authority labels after DOCX reopen')
    fields=[f for f in list(doc.Fields) if f.Code.Text.strip().upper().startswith(('TOC ','TOA '))]
    toc=[f for f in fields if f.Code.Text.strip().upper().startswith('TOC ')]
    if len(toc)!=1:raise RuntimeError('Reopened Word document has no unique TOC field')
    visible=normalized_pdf_text(toc[0].Result.Text)
    for h in headings:
        label=normalized_pdf_text(escape_field_text(h.get('toc_label', h['final'])))
        if label not in visible:
            raise RuntimeError(f'TOC heading lost or altered: {label}')
    if any(x in visible for x in ('QOPENQ','QCLOSEQ','SQOPENQ','SQCLOSEQ')):
        raise RuntimeError('TOC contains an unreplaced quote placeholder')
    for a in authorities:
        group=category_code(a['category'])
        f=[x for x in fields if re.search(r'\\c\s+'+str(group)+r'(?:\s|$)',x.Code.Text,re.I)]
        if len(f)!=1:raise RuntimeError('Missing Word TOA field for '+a['category'])
        if normalized_pdf_text(a['final']) not in normalized_pdf_text(f[0].Result.Text):
            raise RuntimeError('Visible authority label missing: '+a['final'])
    log.append(f'WORD DISPLAY CHECK: {len(headings)} headings and {len(authorities)} authority labels present')


def compress_empty_preliminary_paragraphs(doc, log, progress, preliminary):
    progress('Compressing empty preliminary TOA field-end paragraphs')
    started=False;changed=0
    for p in list(doc.Sections(preliminary).Range.Paragraphs):
        text=p.Range.Text.strip('\r\x07').strip()
        if 'TABLE OF AUTHORITIES' in text:started=True
        if started and not text:
            p.Range.Font.Size=2
            fmt=p.Range.ParagraphFormat
            fmt.SpaceBefore=0;fmt.SpaceAfter=0
            fmt.LineSpacingRule=WD_LINE_SPACE_EXACTLY
            fmt.LineSpacing=2.5
            changed+=1
    log.append(f'PRELIMINARY EMPTY PARAGRAPHS COMPRESSED: {changed}')


def normalized_pdf_text(s):
    return re.sub(r'\s+', ' ', s).strip()


def authority_pdf_patterns(item, aliases=None):
    """Build boundary-safe patterns using actual source text and approved aliases.

    This is a PDF *cross-check*, not a substitute for the Word TA marker
    finder. Do not identify unsuffixed rules inside their subdivisions, or
    guess that a naked evidence-rule number occurs independently.
    """
    raw=item['original']
    category=category_code(item['category'])
    alias_values=(aliases or {}).get(raw.casefold(),[])
    patterns=[]

    def add_literal(phrase):
        phrase=str(phrase or '').strip()
        if len(phrase)<6:return
        pattern=re.escape(phrase).replace(r'\ ',r'\s+')
        if is_rule_variant(phrase):
            # A complete rule (including any subdivisions) is NOT part of a
            # narrower child rule. The boundary mirrors native Word matching.
            pattern+=r'(?![A-Za-z0-9_]|[ \t]*\([A-Za-z0-9]{1,3}\))'
        patterns.append(pattern)

    if category==1:
        title=re.split(r',\s*(?=\d{1,4}\s+(?:U\.S\.|S\.Ct\.|F\.|Cal\.|P\.|WL\b|LEXIS\b))',
                       raw,maxsplit=1,flags=re.I)[0]
        add_literal(title)
        for alias in alias_values:add_literal(alias)
    else:
        # A bare number supplied in a compound citation (for example the
        # second number in 'Fed. R. Evid. 801(c), 802') is NOT globally unique.
        # Only recognize it in a context established by the Word anchor logic.
        bare_number=bool(re.fullmatch(r'\d+(?:\([a-z0-9]+\))*',raw,re.I))
        for phrase in variants_for_authority(item)+alias_values:
            if bare_number and phrase==raw:continue
            add_literal(phrase)
        final=item['final']
        if re.search(r'\bEvid\.\s*802\b',final,re.I):
            patterns.extend([r'801\(c\)\s*,\s*802\b',r'801\(c\)\s+and\s+802\b'])
        if re.search(r'\bEvid\.\s*803\(10\)',final,re.I):
            patterns.extend([r'803\(7\)\s+or\s+803\(10\)',
                             r'803\(7\)\s*,\s*803\(10\)'])
    return list(dict.fromkeys(patterns))


def parse_toa_pdf_entries(doc_pdf,body_start_idx,authorities,prelim_start_idx=0):
    """Read complete TOA rows, not unbounded substrings of child authorities.

    A bare parent such as 'Rule N' cannot be counted from 'Rule N(a)'; the
    full approved label must be followed by actual dot leaders and page nums.
    Support both Word's contiguous leaders and spaced '. . .' dot leaders.
    """
    pages=[normalized_pdf_text(doc_pdf[i].get_text('text')) for i in range(prelim_start_idx,body_start_idx)]
    started=False;toa=[]
    for page in pages:
        if 'TABLE OF AUTHORITIES' in page:started=True
        if started:toa.append(page)
    joined=' '.join(toa)
    joined=re.sub(r'\s+([,.;])',r'\1',joined)
    leader=r'(?:\.{2,}|(?:\.\s*){3,}|(?:…\s*){2,})'
    results={}
    for item in authorities:
        label=normalized_pdf_text(item['final'])
        title=re.escape(label).replace(r'\ ',r'\s+')
        row=re.compile(r'(?<![A-Za-z0-9_])'+title+r'\s*'+leader+
                       r'\s*(\d{1,4}(?:\s*,\s*\d{1,4})*)(?!\d)',flags=re.I)
        found=list(row.finditer(joined))
        results[label]=[int(n) for n in re.findall(r'\d+',found[0].group(1))] if len(found)==1 else None
    return results


def pdf_motion_body_text(page):
    """Read visible body text without the pleading gutter or page footer.

    Filtering by PDF block geometry ensures a case name split at a physical
    page boundary remains contiguous when adjacent pages are joined below.
    All normal text (including footnotes within the body) stays eligible.
    """
    fragments=[]
    page_top,page_bottom=page.rect.y0,page.rect.y1
    for block in page.get_text('blocks'):
        if len(block)<7 or block[6]!=0:continue
        x0,y0,x1,y1,txt=block[:5]
        if y1<=page_top+50 or y0>=page_bottom-55:continue
        clean=txt.strip()
        if not clean:continue
        if (x1<=page.rect.x0+100 and
                re.fullmatch(r'(?:[1-9]|1\d|2[0-8])(?:\s+(?:[1-9]|1\d|2[0-8]))*',clean)):
            continue  # 28-line numbered pleading margin, not body content
        fragments.append(clean)
    return normalized_pdf_text(' '.join(fragments))


def body_authority_page_hits(pdf,body_idx,item,aliases=None):
    """Find appearances across PDF page breaks, attributing to START page.

    Matches are found in a single body-text stream (with removed pagination
    furniture) so a citation spanning pages can be counted on the first page.
    This mirrors the Word TA fields' start-page anchoring convention.
    """
    import bisect
    texts=[pdf_motion_body_text(pdf[i]) for i in range(body_idx,len(pdf))]
    starts=[];pieces=[];cursor=0
    for page_text in texts:
        starts.append(cursor)
        piece=page_text+'\n'
        pieces.append(piece)
        cursor+=len(piece)
    whole=''.join(pieces)
    hits=set()
    for pattern in authority_pdf_patterns(item,aliases):
        for m in re.finditer(pattern,whole,flags=re.I):
            index=bisect.bisect_right(starts,m.start())-1
            if index>=0:hits.add(index+1)
    return sorted(hits)


def pdf_pleading_grid_audit(pdf, body_start_idx, prelim_start_idx=0):
    """Only audit a numbered 28-line grid when actually detectable."""
    details=[];issues=[];total=0;worst=0.0;grid_pages=0
    for page_idx in range(prelim_start_idx,body_start_idx):
        page=pdf[page_idx]; numbers=[]; entries=[]
        for block in page.get_text('dict')['blocks']:
            if block.get('type')!=0:continue
            for line in block['lines']:
                spans=[s for s in line['spans'] if s['text'].strip()]
                for sp in spans:
                    x=sp['bbox'][0];t=sp['text'].strip()
                    if 35<=x<=95 and re.fullmatch(r'(?:[1-9]|1\d|2[0-8])',t):
                        numbers.append((sp['bbox'][1]+sp['bbox'][3])/2)
                content=[sp for sp in spans if 100<=sp['bbox'][0]<=590 and 65<=sp['bbox'][1]<=760]
                if not content:continue
                label=' '.join(sp['text'].strip() for sp in content)
                if not label or label.upper() in ('CASES','STATUTES','RULES','OTHER AUTHORITIES','CASES (CONTINUED)') or \
                   re.fullmatch(r'[ivxlcdm]+',label,re.I) or 'TABLE OF ' in label or 'PLAINTIFF' in label:continue
                center=(content[0]['bbox'][1]+content[0]['bbox'][3])/2
                entries.append((center,label[:75]))
        if len(set(round(y,1) for y in numbers))<24:
            details.append(f'PDF p.{page_idx+1}: no 28-line pleading grid detected (not audited)')
            continue
        grid_pages+=1
        for center,label in entries:
            offset=min(abs(center-y) for y in numbers)
            total+=1;worst=max(worst,offset)
            if offset>PLEADING_GRID_TOLERANCE_PT:
                issues.append(f'REVIEW: PDF p.{page_idx+1} entry {offset:.1f}pt off grid: {label!r}')
        details.append(f'PDF p.{page_idx+1}: checked {len(entries)} entries, {len(numbers)} number spans')
    details.append(f'Grid audit: {grid_pages} numbered pages, {total} entries; worst offset {worst:.1f}pt')
    return issues,details


def citation_coverage_diagnostic(pdf,body_idx,authorities):
    """Heuristic only; report any unlisted reporter signatures or ambiguous Id."""
    warnings=[];findings=[]
    bodies=[' '.join(pdf[i].get_text('text').split()) for i in range(body_idx,len(pdf))]
    cases=[x for x in authorities if category_code(x['category'])==1]
    reporter=re.compile(r'\b(\d{1,4})\s+(U\.S\.|F\.\s*(?:2d|3d|4th)|U\.S\.\s*Dist\.\s*LEXIS)\s+(\d{1,7})\b',re.I)
    key=lambda match:re.sub(r'\s+','', ''.join(match.groups())).lower()
    known={key(m) for a in cases for m in reporter.finditer(a['original'])}
    for number,body in enumerate(bodies,1):
        for m in reporter.finditer(body):
            if key(m) not in known:
                warnings.append(f'UNLISTED REPORTER on body p.{number}: {m.group(0)}')
        for m in re.finditer(r'\bId\.',body,re.I):
            window=body[max(0,m.start()-180):m.end()+35]
            findings.append(f'Id. on body p.{number} (resolve against preceding source): {window!r}')
    lines=['CITATION COVERAGE HEURISTIC — NOT LEGAL CITATOR VERIFICATION',
           f'{len(cases)} indexed cases, {len(bodies)} motion pages',
           'Short forms and Id. may be unmarked unless listed as Aliases.',
           '', 'REPORTER/ID OCCURRENCES:',*(findings or ['None']),
           '', 'WARNINGS:',*(warnings or ['None detected']),
           '', 'No web requests or paid AI APIs are made by the generator.']
    return warnings,lines



def pdf_page_edge_audit(pdf):
    """Detect text genuinely outside PDF media box; not a Word-view clipping guarantee."""
    issues=[]
    for index,page in enumerate(pdf,1):
        page_rect=page.rect
        for block in page.get_text('dict')['blocks']:
            if block.get('type')!=0:continue
            for line in block.get('lines',[]):
                for span in line.get('spans',[]):
                    text=span['text'].strip()
                    if not text:continue
                    rect=pymupdf.Rect(span['bbox'])
                    if (rect.x0 < page_rect.x0-0.5 or rect.x1 > page_rect.x1+0.5
                            or rect.y0 < page_rect.y0-0.5 or rect.y1 > page_rect.y1+0.5):
                        issues.append(f'REVIEW: PDF p.{index} text outside the page boundary: {text[:65]!r}')
    return issues

def pdf_validate(pdf_path,authorities,headings,body_start_phys,report_path,quality_path=None,prelim_phys=2,numbered_grid=False,aliases=None,min_table_font_pt=12.0):
    """PDF audit for arbitrary approved lists; never use motion-specific names."""
    if pymupdf is None:raise RuntimeError('PDF audit requires PyMuPDF')
    pdf=pymupdf.open(str(pdf_path))
    issues=[];details=[]
    try:
        body_idx=body_start_phys-1;prelim_start=prelim_phys-1
        if len(pdf)<=body_idx:raise RuntimeError('PDF has no body pages')
        printed=parse_toa_pdf_entries(pdf,body_idx,authorities,prelim_start)
        for a in authorities:
            label=normalized_pdf_text(a['final'])
            actual=body_authority_page_hits(pdf,body_idx,a,aliases)
            shown=printed.get(label)
            status='PASS' if shown is not None and set(shown)==set(actual) else 'REVIEW'
            line=f'{status}: {label} | TOA={shown} | body={actual}'
            details.append(line)
            if status!='PASS':issues.append(line)
        prelim=' '.join(normalized_pdf_text(pdf[i].get_text('text')) for i in range(prelim_start,body_idx))
        toc_part=prelim.split('TABLE OF AUTHORITIES',1)[0]
        if re.search(r'TABLE\s+OF\s+AUTHORITIES\s*\(continued\)',prelim,flags=re.I):
            issues.append('REVIEW: TOA continuation title is not allowed')
        for h in headings:
            label=normalized_pdf_text(escape_field_text(h.get('toc_label', h['final'])))
            if label not in toc_part:
                issues.append('REVIEW: missing or altered TOC label: '+label)
        for i in range(prelim_start,body_idx):
            ptext=normalized_pdf_text(pdf[i].get_text('text'))
            if len(ptext)<25:issues.append(f'REVIEW: possible blank preliminary page {i+1}')
        grid_issues,grid_details=pdf_pleading_grid_audit(pdf,body_idx,prelim_start) if numbered_grid else ([],['Unnumbered document: grid check not applicable'])
        issues.extend(grid_issues)
        for i in range(prelim_start,body_idx):
            for block in pdf[i].get_text('dict')['blocks']:
                if block.get('type')!=0:continue
                for line in block['lines']:
                    for span in line['spans']:
                        txt=span['text'].strip()
                        if not txt or txt.isdigit() or re.fullmatch(r'[ivxlcdm]+',txt,re.I):continue
                        if span['bbox'][0]<100 or span['bbox'][1]>710:continue  # skip preserved footer
                        if span['size']<min_table_font_pt-0.2:
                            issues.append(f'REVIEW: preliminary page {i+1} text below {min_table_font_pt:g}pt: {txt[:40]!r}')
                            break
        # Detect true PDF-edge overflow; a separate visual check of the Word
        # document is still required because Word draft/layout view can differ.
        edge_issues=pdf_page_edge_audit(pdf)
        issues.extend(edge_issues)
        warnings,cover=citation_coverage_diagnostic(pdf,body_idx,authorities)
        if quality_path:Path(quality_path).write_text('\n'.join(cover)+'\n',encoding='utf-8')
        issues.extend(warnings)
        report=['PDF AUDIT — Word-generated, not independent legal citator verification',
                f'PDF pages: {len(pdf)}; first motion body PDF page: {body_start_phys}',
                f'Approved authorities: {len(authorities)}; headings: {len(headings)}',
                '',*details,'','PLEADING GRID:',*grid_details,
                '',f'PDF EDGE CLIPPING: {len(edge_issues)} span(s) outside media box',
                'NOTE: PDF checks cannot prove the Word editor display is unclipped.',
                '','ISSUES:',*(issues or ['None detected']),
                '', 'Manual review: all Id./supra, legal validity, court-specific rules and long document layout.']
        Path(report_path).write_text('\n'.join(report)+'\n',encoding='utf-8')
        return issues
    finally:pdf.close()


def aliases_from_workbook(wb, authorities):
    """Optional Aliases sheet: original authority in A, exact occurrence B."""
    if 'Aliases' not in wb.sheetnames:return {}
    approved={a['original'].casefold() for a in authorities}
    alias_map=defaultdict(list)
    for n,row in enumerate(list(wb['Aliases'].values)[1:],2):
        if not row or not row[0]:continue
        key=str(row[0]).strip().casefold()
        phrase=str(row[1] or '').strip() if len(row)>1 else ''
        if key not in approved:raise ValueError(f'Aliases row {n}: no approved authority {row[0]!r}')
        if len(phrase)<5 or phrase.casefold() in ('id.','ibid.','supra'):
            raise ValueError(f'Aliases row {n}: use a distinctive exact passage (not bare Id.)')
        alias_map[key].append(phrase)
    return dict(alias_map)


def document_paragraphs(docx_path):
    """Cross-platform read-only preflight of body and footnote text."""
    import xml.etree.ElementTree as ET
    ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    all_parts=[]
    with zipfile.ZipFile(docx_path) as z:
        names=['word/document.xml']+[x for x in ('word/footnotes.xml','word/endnotes.xml') if x in z.namelist()]
        for name in names:
            root=ET.fromstring(z.read(name))
            ps=[]
            for p in root.findall('.//w:p',ns):
                text=''.join(t.text or '' for t in p.findall('.//w:t',ns))
                if text.strip():ps.append(text)
            all_parts.append((name,ps))
    return all_parts


def normalize_caption_labels(doc, log, first_body_heading=None):
    """Fix caption font *in the working copy only*, before PDF/A export.

    Some pleading templates name their caption font 'Times New Roman Bold'
    without setting Word's separate bold attribute. Word can therefore show
    the lines as ordinary text but export them in a bold font. Limit the
    remedy to the Date, Time, Location and Judge lines in caption tables.
    Do not touch motion titles or any body paragraphs.
    """
    changed=[]
    # Ignore caption-like words in any evidentiary tables in the motion body.
    before_pos=None
    if first_body_heading:
        matches=heading_ranges(doc,first_body_heading)
        if len(matches)!=1:
            raise RuntimeError('Cannot safely locate beginning of motion for caption-only formatting')
        before_pos=matches[0].Start
    for table_num in range(1,doc.Tables.Count+1):
        table=doc.Tables(table_num)
        if before_pos is not None and table.Range.Start >= before_pos:
            continue
        for paragraph in table.Range.Paragraphs:
            text=paragraph.Range.Text.replace('\r','').replace('\x07','').strip()
            match=re.match(r'^(Date|Time|Location|Judge)\s*:',text,re.I)
            if not match:
                continue
            paragraph.Range.Font.Name='Times New Roman'
            paragraph.Range.Font.Bold=0
            changed.append(match.group(1).title())
    log.append('CAPTION FONT: '+(', '.join(changed)+' set to regular Times New Roman in output only'
                                 if changed else 'No Date/Time/Location/Judge caption table labels found; unchanged'))
    return changed


def export_court_pdfa(doc, destination):
    """Use Word's native PDF/A-1 export, omitting source document properties.

    This still produces standards-required XMP metadata, which must NOT be
    stripped: removing it would invalidate PDF/A. No PDF postprocessing is done.
    """
    return doc.ExportAsFixedFormat(
        str(destination), PDF,
        OpenAfterExport=False,
        OptimizeFor=0,  # Print quality; retain vector/text content
        IncludeDocProps=False,
        KeepIRM=False,
        CreateBookmarks=0,
        DocStructureTags=True,
        UseISO19005_1=True,
    )


def recognized_pdfa_profile(xmp):
    """Read a genuine PDF/A-1/2/3 XMP identification, never fabricate one.

    Returns profile string suitable for veraPDF's -f flag, e.g. '3a'.
    An XMP claim is necessary but insufficient proof of ISO compliance.
    """
    import xml.etree.ElementTree as ET
    namespace='http://www.aiim.org/pdfa/ns/id/'
    part_tag='{'+namespace+'}part'
    conf_tag='{'+namespace+'}conformance'
    if not xmp:
        raise RuntimeError('FILING BLOCKED: PDF/A identification metadata is missing from XMP. '
                           'Word did not export a recognizable archival PDF.')
    try:
        root=ET.fromstring(xmp)
    except ET.ParseError as exc:
        raise RuntimeError('FILING BLOCKED: PDF/A XMP is not well-formed XML') from exc
    parts=set()
    conformance_levels=set()
    for element in root.iter():
        part=element.attrib.get(part_tag)
        conf=element.attrib.get(conf_tag)
        if part: parts.add(part.strip())
        if conf: conformance_levels.add(conf.strip().upper())
        if element.tag==part_tag and element.text:
            parts.add(element.text.strip())
        if element.tag==conf_tag and element.text:
            conformance_levels.add(element.text.strip().upper())
    allowed={'1':{'A','B'}, '2':{'A','B','U'}, '3':{'A','B','U'}}
    if len(parts)!=1 or len(conformance_levels)!=1:
        raise RuntimeError('FILING BLOCKED: PDF/A identification in XMP is missing or contradictory '
                           '(part='+repr(sorted(parts))+', conformance='+repr(sorted(conformance_levels))+').')
    part=next(iter(parts))
    conf=next(iter(conformance_levels))
    if conf not in allowed.get(part,set()):
        raise RuntimeError('FILING BLOCKED: unsupported PDF/A identification in XMP: '
                           f'PDF/A-{part}{conf.lower()}. Supported: PDF/A-1a/1b, -2a/2b/2u, -3a/3b/3u.')
    return part+conf.lower()


def pdfa_safety_audit(pdf_path, *, independent_validator=None):
    """Check PDF/A declaration and CM/ECF risk features; report limits.

    True standard conformance needs an independent PDF/A validator. If veraPDF
    is installed or VERAPDF_CLI is set, run it and stop on nonconformance.
    """
    if pymupdf is None:
        raise RuntimeError('PDF/A filing audit requires PyMuPDF')
    warnings=[]
    results=['COURT-FILING PDF/A SAFETY CHECK',
             'Word export: UseISO19005_1=True, IncludeDocProps=False, KeepIRM=False',
             'Actual ISO 19005 part is identified from the exported XMP (Word may output PDF/A-3a).',
             'PDF/A needs standards-required XMP and ICC profiles; this metadata must remain.']
    if Path(pdf_path).stat().st_size >= 50_000_000:
        raise RuntimeError('FILING BLOCKED: PDF is 50 MB or larger; Northern District of California ECF limit is 50 MB')
    d=pymupdf.open(str(pdf_path))
    try:
        if d.needs_pass or d.is_encrypted:
            raise RuntimeError('FILING BLOCKED: PDF is encrypted or password-protected')
        if d.embfile_count():
            raise RuntimeError('FILING BLOCKED: PDF contains embedded attachments')
        if d.is_form_pdf:
            raise RuntimeError('FILING BLOCKED: PDF contains interactive form fields')
        found=[]
        bad=re.compile(r'/(?:JavaScript|JS|Launch|EmbeddedFiles|RichMedia|Movie|Sound|GoToR|SubmitForm|ResetForm)\b')
        for xref in range(1,d.xref_length()):
            try:
                obj=d.xref_object(xref,compressed=False)
            except Exception:
                continue
            for token in sorted(set(bad.findall(obj))):
                found.append((xref,token))
        if found:
            raise RuntimeError('FILING BLOCKED: PDF contains active/prohibited objects: '+repr(found[:10]))
        xref=d.xref_xml_metadata()
        xmp=d.xref_stream(xref).decode('utf-8','replace') if xref else ''
        # Microsoft Word for Microsoft 365 can emit PDF/A-3a even with
        # UseISO19005_1=True. Read the ACTUAL XMP profile instead of
        # assuming PDF/A-1. Never modify or manufacture an XMP declaration.
        profile=recognized_pdfa_profile(xmp)
        results.append('PDF/A declaration in XMP: PDF/A-'+profile)
        results.append('PDF/A identification checked; full ISO conformance requires independent validation.')
        results.append(f'Pages: {len(d)}; PDF size: {Path(pdf_path).stat().st_size:,} bytes (court limit <50,000,000); encrypted: no; attachments: none; forms: none; active objects: none')
        props=d.metadata
        display={key:props.get(key,'') for key in ('title','author','subject','keywords') if props.get(key,'')}
        if display:
            warnings.append('Nonessential document properties remain in PDF metadata: '+repr(display))
        results.append('Optional source document properties: '+('found; manual privacy review recommended' if display else 'not present'))
        # Check caption page only, before altering anything else.
        bad_bold=[]
        for block in d[0].get_text('dict').get('blocks',[]):
            for line in block.get('lines',[]):
                for span in line.get('spans',[]):
                    m=re.match(r'^\s*(Date|Time|Location|Judge)\s*:',span.get('text',''),re.I)
                    if m and (span.get('flags',0) & 16 or 'Bold' in span.get('font','')):
                        bad_bold.append(m.group(1))
        if bad_bold:
            raise RuntimeError('FILING BLOCKED: caption lines still exported bold: '+', '.join(bad_bold))
        results.append('Caption Date/Time/Location/Judge lines: regular weight (where present)')
        if not d[0].get_text('text').strip():
            raise RuntimeError('FILING BLOCKED: caption page is not text-searchable')
        results.append('Text-searchable caption page: yes')
    finally:
        d.close()

    validator=independent_validator or os.environ.get('VERAPDF_CLI') or shutil.which('verapdf') or shutil.which('verapdf.bat')
    if validator:
        try:
            proc=subprocess.run([str(validator),'-f',profile,str(pdf_path)],capture_output=True,text=True,timeout=180)
        except Exception as exc:
            raise RuntimeError('FILING BLOCKED: veraPDF could not run successfully: '+str(exc)) from exc
        compliant=re.search(r'isCompliant=["\'](true|false)["\']',proc.stdout,re.I)
        if proc.returncode!=0 or not compliant or compliant.group(1).lower()!='true':
            raise RuntimeError('FILING BLOCKED: independent veraPDF PDF/A-'+profile+' validation failed. '+
                               (proc.stdout+' '+proc.stderr)[-1500:])
        results.append('Independent veraPDF PDF/A-'+profile+' validation: PASS')
    else:
        results.append('Independent veraPDF PDF/A-'+profile+' validation: NOT RUN (veraPDF not installed).')
        warnings.append('PDF/A-'+profile+' is declared in XMP, but full ISO conformance was not independently verified.')
    results.extend('REVIEW: '+w for w in warnings)
    return results, warnings


def verify_motion_body_text_unchanged(original_docx, generated_docx, first_heading):
    """Compare source and finished Word BODY content, excluding new front matter.

    The generated TOC and TA fields can add non-visible XML field runs, but they
    may not change source paragraph text, even by one character.
    """
    orig = document_paragraphs(original_docx)[0][1]
    final = document_paragraphs(generated_docx)[0][1]
    norm = normalized_heading(first_heading)
    ori_hits = [i for i, t in enumerate(orig) if normalized_heading(t) == norm]
    fin_hits = [i for i, t in enumerate(final) if normalized_heading(t) == norm]
    if len(ori_hits) != 1 or len(fin_hits) != 1:
        raise RuntimeError('Source motion text comparison cannot identify the unique first body heading')
    # The caption must also be preserved. New tables are inserted only BETWEEN
    # the original caption and the first body heading.
    if orig[:ori_hits[0]] != final[:ori_hits[0]]:
        raise RuntimeError('SOURCE MOTION CAPTION CHANGED; output is not filing-ready')
    before = orig[ori_hits[0]:]
    after = final[fin_hits[0]:]
    if before != after:
        from itertools import zip_longest
        for i, pair in enumerate(zip_longest(before, after, fillvalue='<missing>')):
            if pair[0] != pair[1]:
                raise RuntimeError(f'SOURCE MOTION CONTENT CHANGED at body paragraph {i+1}: '
                                   f'original {pair[0][:100]!r}; generated {pair[1][:100]!r}. '
                                   'Please do not file the output.')
    orig_parts = dict(document_paragraphs(original_docx))
    final_parts = dict(document_paragraphs(generated_docx))
    for part in ('word/footnotes.xml','word/endnotes.xml'):
        if orig_parts.get(part, []) != final_parts.get(part, []):
            raise RuntimeError(f'SOURCE MOTION {part} TEXT CHANGED; please do not file the output')
    return len(before)


def count_source_sections(source):
    with zipfile.ZipFile(source) as z:
        xml=z.read('word/document.xml').decode('utf-8')
    # The final sectPr element closes a normal Word document; each additional
    # sectPr signals an existing boundary that might disrupt footer numbering.
    return max(1,len(re.findall(r'<w:sectPr\b',xml)))


def read_setup(wb):
    if 'Setup' not in wb.sheetnames:
        raise ValueError('CLARIFICATION NEEDED: Setup sheet is missing. Use the V15 '
                         'template; specify the motion type and first numbered-body heading.')
    values={}
    for row in list(wb['Setup'].values)[1:]:
        if row and row[0]:values[str(row[0]).strip().casefold()]=str(row[1] or '').strip()
    answer=values.get('motion type','')
    if not answer or answer.casefold() in ('review','unknown','uncertain','tbd','?', 'not sure'):
        raise ValueError('CLARIFICATION NEEDED: What filing is this (motion, opposition, reply, etc.)? '
                         'Record the answer in Setup!B2.')
    # Backward compatible with early V15 workbooks labelled First body heading.
    body_answer=(values.get('first body anchor') or values.get('first body heading') or '').strip()
    if not body_answer or body_answer.casefold() in ('review','unknown','uncertain','tbd','?', 'not sure'):
        raise ValueError('CLARIFICATION NEEDED: Which precise paragraph begins the numbered motion body '
                         'after the cover/caption? Record its text in Setup!B3.')
    return values



# --- V16.4.6 guided desktop workflow and preparation provenance ---
CHATGPT_PROMPT_FILE = 'HipoLaw_ChatGPT_Prompt.txt'
PREPARATION_MANIFEST_SUFFIX = '.hipolaw_prepared.json'
BINDING_SHA_KEY = 'hipolaw document sha256'
BINDING_NAME_KEY = 'hipolaw document filename'


def external_file_path(name):
    base = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
    return base / name


def sha256_file(path):
    hasher=hashlib.sha256()
    with open(path,'rb') as stream:
        for chunk in iter(lambda: stream.read(1024*1024), b''):
            hasher.update(chunk)
    return hasher.hexdigest()


def prepared_manifest_path(docx_path):
    path=Path(docx_path)
    return path.with_name(path.name + PREPARATION_MANIFEST_SUFFIX)


def load_prepared_manifest(docx_path):
    path=prepared_manifest_path(docx_path)
    if not path.is_file():
        if Path(docx_path).stem.endswith('_HipoLaw_Prepared'):
            raise ValueError('This appears to be a HipoLaw prepared Word document but its .hipolaw_prepared.json file is missing. Keep the sidecar beside the DOCX or rerun Step 1.')
        return None
    try:
        manifest=json.loads(path.read_text(encoding='utf-8'))
    except (ValueError,OSError) as exc:
        raise ValueError(f'Invalid HipoLaw prepared-motion manifest {path}: {exc}') from exc
    if manifest.get('schema_version') != 1 or manifest.get('prepared_filename') != Path(docx_path).name:
        raise ValueError('Prepared motion manifest is invalid or does not match this Word file.')
    digest=sha256_file(docx_path)
    if digest != manifest.get('prepared_sha256'):
        raise ValueError('Prepared Word document changed after Step 1. Run Prepare again, then regenerate and review the Excel spreadsheet.')
    if not manifest.get('first_body_heading'):
        raise ValueError('Prepared manifest is missing the original first body heading.')
    return manifest


def read_spreadsheet_binding(excel):
    if openpyxl is None:
        raise RuntimeError('Spreadsheet binding verification requires openpyxl')
    wb=openpyxl.load_workbook(excel, read_only=True, data_only=True)
    try:
        setup=read_setup(wb)
        return setup.get(BINDING_SHA_KEY, '').strip().lower(), setup.get(BINDING_NAME_KEY, '').strip(), setup
    finally:
        wb.close()


def verify_spreadsheet_motion_binding(source, excel):
    """Prevent an approved workbook from being used against the wrong prepared motion.

    Legacy V15 workbooks remain usable with ordinary original .docx files.
    A HipoLaw-prepared .docx requires a matching hash in Setup, because the
    ChatGPT step must operate on that exact prepared file.
    """
    manifest=load_prepared_manifest(source)
    spreadsheet_hash, spreadsheet_name, setup=read_spreadsheet_binding(excel)
    actual_hash=manifest['prepared_sha256'] if manifest else sha256_file(source)
    if manifest and not spreadsheet_hash:
        raise ValueError('This is a HipoLaw prepared motion, but the workbook Setup sheet has no '
                         '"HipoLaw document SHA256" row. Use the Step 2 copied prompt, or add the '
                         'SHA256 and filename rows to Setup before approving the spreadsheet.')
    if spreadsheet_hash and not re.fullmatch('[0-9a-f]{64}', spreadsheet_hash):
        raise ValueError('Workbook Setup HipoLaw document SHA256 must be exactly 64 hexadecimal characters.')
    if spreadsheet_hash and spreadsheet_hash != actual_hash:
        raise ValueError('STOP: Approved spreadsheet belongs to a different Word motion or an earlier '
                         'prepared version (SHA256 mismatch). Do not reuse its TOA page references.')
    if manifest and not spreadsheet_name:
        raise ValueError('Prepared-motion workbook Setup is missing HipoLaw document filename.')
    if spreadsheet_name and spreadsheet_name != Path(source).name:
        raise ValueError('STOP: Workbook Setup HipoLaw document filename does not match selected motion.')
    if manifest:
        if normalized_heading(setup.get('first body anchor', setup.get('first body heading',''))) != normalized_heading(manifest['first_body_heading']):
            raise ValueError('STOP: Workbook first-body heading does not match prepared-motion manifest.')
        workbook_grid = read_whole_motion_grid_option(setup)
        if manifest['align_grid']:
            configured_shift=load_style_settings()['pleading_grid'].get('body_header_shift_pt',9.5)
            if float(manifest.get('body_header_shift_pt') or 0)!=float(configured_shift):
                raise ValueError('Prepared motion is from an older alignment build or uses a different pleading-number baseline setting. '
                                 'Rerun Step 1 with the current style settings.')
            expected_spacing=manifest.get('line_spacing_pt')
            configured_spacing=load_style_settings()['pleading_grid']['line_spacing_pt']
            if expected_spacing is not None and float(expected_spacing) != float(configured_spacing):
                raise ValueError('Prepared motion was aligned with a different grid line-spacing setting. '
                                 'Rerun Step 1 after changing style settings.')
        if workbook_grid != bool(manifest['align_grid']):
            raise ValueError('STOP: Workbook whole-motion grid Yes/No does not match the Step 1 prepared motion. '
                             'The setting must describe the prepared document you actually uploaded.')
    return manifest


def suggest_first_body_heading(docx_path):
    """A suggestion, not an auto-decision: UI displays editable value."""
    try:
        parts = document_paragraphs(docx_path)[0][1]
    except Exception:
        return ''
    candidates = [p.strip() for p in parts if normalized_heading(p).upper() == 'INTRODUCTION']
    return candidates[0] if len(candidates) == 1 else ''


def safe_preparation_target(original, output_parent=None, timestamp=None):
    original=Path(original)
    parent=Path(output_parent) if output_parent else original.parent
    stamp=timestamp or datetime.now().strftime('%Y%m%d_%H%M%S')
    name=re.sub('[^A-Za-z0-9_. -]','_', original.stem)[:80].rstrip()
    folder=parent/f'HipoLaw_Prepared_{stamp}'
    if folder.exists():
        raise FileExistsError(f'Prepared output folder already exists: {folder}')
    return folder, folder/(name+'_HipoLaw_Prepared.docx')


def _header_number_paragraph(header):
    """Locate the *one* 1–28 header paragraph, excluding law-firm artwork.

    The existing pleading paper stores the 28 numbers in a fixed-height Word
    header table cell. Word's Tables collection is more reliable than Find in
    the presence of the vertical VML firm-name box.
    """
    candidates=[]
    try:
        for table in header.Range.Tables:
            for cell in table.Range.Cells:
                for paragraph in cell.Range.Paragraphs:
                    line=str(paragraph.Range.Text or '')
                    numbers={int(v) for v in re.findall(r'(?<!\d)(?:[1-9]|1\d|2[0-8])(?!\d)',line)}
                    if numbers==set(range(1,29)):
                        key=(int(paragraph.Range.Start),int(paragraph.Range.End))
                        if all(item[0]!=key for item in candidates):
                            candidates.append((key,paragraph))
    except Exception as exc:
        raise RuntimeError('Unable to read the pleading-number header table: '+str(exc)) from exc
    if len(candidates)!=1:
        raise RuntimeError(f'Expected exactly one 28-number header paragraph; found {len(candidates)}. '
                           'Alignment stopped rather than changing the wrong header.')
    return candidates[0][1]


def _shift_number_column(header, shift_pt, log, label):
    """Shift the 1–28 column's vertical position; do not alter body text."""
    if not -18 <= float(shift_pt) <= 18:
        raise ValueError('Pleading-number baseline shift must be within -18pt to +18pt')
    paragraph=_header_number_paragraph(header)
    pf=paragraph.Range.ParagraphFormat
    old=float(pf.SpaceBefore or 0)
    updated=old+float(shift_pt)
    if not 0<=updated<=120:
        raise RuntimeError(f'Unsafe resulting pleading-header spacing: {updated:g}pt')
    pf.SpaceBefore=updated
    actual=float(pf.SpaceBefore or 0)
    if abs(actual-updated)>0.3:
        raise RuntimeError(f'Word did not apply the requested header correction: {actual:g}pt')
    log.append(f'{label}: 1–28 header-number column shifted {shift_pt:+g}pt '
               f'(paragraph before-spacing {old:g}pt -> {actual:g}pt)')
    return actual


def calibrate_prepared_motion_numbers(doc, *, enabled, shift_pt, log):
    """Calibrate the default header in Step 1, preserving caption first-page header.

    The source's title-page setting has a separate first-page header, so the
    cover retains its original 1–28 numbers. Regular pages receive the aligned
    baseline before the Word preview and SHA256 are generated.
    """
    if not enabled:return False
    if doc.Sections.Count!=1:
        raise RuntimeError('Step 1 grid calibration expects one Word section')
    section=doc.Sections(1)
    if not section.PageSetup.DifferentFirstPageHeaderFooter:
        raise RuntimeError('Cannot preserve the caption grid: source uses a shared first-page header. '
                           'Supply a pleading template with a distinct caption-page header.')
    _shift_number_column(section.Headers(PRIMARY),shift_pt,log,'STEP 1 BODY GRID')
    log.append('Caption first-page header unchanged; subsequent body pages aligned before preview')
    return True


def preserve_prepared_body_alignment(doc,preliminary,body_section, *, shift_pt,log):
    """Separate preliminary-table numbers from an already-aligned Step 1 body.

    Step 1 has shifted the default pleading header. After inserting the table
    sections, unlink the generated body header FIRST so it retains that shift;
    then unlink and restore the preliminary pages' original number baseline.
    Never shift the already-prepared body a second time.
    """
    body=doc.Sections(body_section).Headers(PRIMARY)
    body.LinkToPrevious=False
    _header_number_paragraph(body)  # fail closed if Word omitted the copied header
    prelim=doc.Sections(preliminary).Headers(PRIMARY)
    prelim.LinkToPrevious=False
    _shift_number_column(prelim,-float(shift_pt),log,'STEP 3 PRELIMINARY GRID')
    # Verify the two section headers are independent and correctly separated.
    body_before=float(_header_number_paragraph(body).Range.ParagraphFormat.SpaceBefore or 0)
    prelim_before=float(_header_number_paragraph(prelim).Range.ParagraphFormat.SpaceBefore or 0)
    if abs((body_before-prelim_before)-float(shift_pt))>0.5:
        raise RuntimeError('Generated preliminary and body headers were not isolated; '
                           'refusing to produce a mismatched filing.')
    log.append('STEP 3 BODY GRID: prealigned Step 1 header retained; no double adjustment')


def verify_prepared_grid_preview(pdf_path, first_body_heading, log, tolerance_pt=3.0):
    """Check first motion page's line-1 text versus the printed numeral 1.

    This is a targeted geometry check, not a general proof that every
    illustration, footnote, or floating object can follow a 28-line grid.
    """
    if pymupdf is None:
        raise RuntimeError('PyMuPDF is needed to verify prepared pleading-grid alignment')
    pdf=pymupdf.open(str(pdf_path))
    try:
        if len(pdf)<2:
            raise RuntimeError('Grid preview has no second page for the motion body')
        page=pdf[1]
        digits=[w for w in page.get_text('words') if w[4]=='1' and 65<=w[0]<=103 and 45<=w[1]<=110]
        if len(digits)!=1:
            raise RuntimeError('Cannot locate line number 1 on the first motion page')
        requested=normalized_heading(first_body_heading)
        # Search the first heading as printed, ignoring its numbering prefix.
        phrase=re.sub(r'^((?:[IVXLCDM]+|[A-Z]|[ivxlcdm]+)[.]?\s+)', '',str(first_body_heading)).strip()
        found=page.search_for(phrase) if phrase else []
        if not found:
            raise RuntimeError('Cannot locate the first body heading in the PDF preview')
        number_top=digits[0][1]
        heading_top=min(r.y0 for r in found if r.x0>120)
        offset=heading_top-number_top
        log.append(f'PDF PREVIEW GRID CHECK: first body heading versus printed line 1 '
                   f'= {offset:+.2f}pt (tolerance ±{tolerance_pt:g}pt)')
        if abs(offset)>float(tolerance_pt):
            raise RuntimeError('Prepared preview is still not aligned: first heading is '
                               f'{offset:+.1f}pt from pleading line 1. This Word template '
                               'requires a different header calibration; no success manifest created.')
        return offset
    finally:
        pdf.close()


def prepare_motion(original, first_body_heading, *, align_grid=False, output_parent=None,
                   progress_callback=None, word_factory=None):
    """Prepare a NEW working copy, never altering or opening the original for writing.

    Formatting occurs before any TOC/TOA citations are added. An optional PDF
    preview is intentionally labeled NOT_FOR_FILING. The final Word generator
    performs the body-only numbered-header calibration after separating the
    caption and body Word sections.
    """
    source=Path(original).resolve()
    if not source.is_file() or source.suffix.lower()!='.docx' or not zipfile.is_zipfile(source):
        raise ValueError('Choose an existing Microsoft Word .docx motion.')
    first_body_heading=str(first_body_heading).strip()
    if not first_body_heading:
        raise ValueError('Enter the EXACT first heading of the motion body (for example INTRODUCTION).')
    source_paragraphs=document_paragraphs(source)[0][1]
    target=normalized_heading(first_body_heading)
    hits=[p for p in source_paragraphs if normalized_heading(p)==target]
    if len(hits)!=1:
        raise ValueError(f'First body heading {first_body_heading!r} matches {len(hits)} Word paragraphs. '
                         'Enter unique exact heading text before preparing.')
    if count_source_sections(source)!=1:
        raise ValueError('HipoLaw preparation requires a single-section Word motion. No changes were made.')
    number_header=has_numbered_pleading_header(source)
    if align_grid and not number_header:
        raise ValueError('Grid alignment requested but no standard 28-line pleading header was found.')
    style=load_style_settings()
    folder, output=safe_preparation_target(source,output_parent)
    folder.mkdir(parents=True,exist_ok=False)
    preview=folder/(output.stem+'_PREVIEW_NOT_FOR_FILING.pdf')
    shutil.copy2(source,output)
    log=['PREPARATION MODE (V16.4.6): generated working copy; original left untouched',
         'ORIGINAL: '+str(source), 'PREPARED: '+str(output),
         'FIRST BODY HEADING: '+first_body_heading,
         'ALIGN ENTIRE MOTION GRID: '+('Yes' if align_grid else 'No')]
    def progress(msg):
        log.append(msg)
        if progress_callback:progress_callback(msg)
    word=None;doc=None
    try:
        if word_factory is None:
            if win32 is None:
                raise RuntimeError('Prepare requires Microsoft Word for Windows and pywin32')
            word=win32.DispatchEx('Word.Application')
        else:
            word=word_factory()
        word.Visible=False;word.DisplayAlerts=0;word.ScreenUpdating=False
        doc=word.Documents.Open(str(output),ReadOnly=False,AddToRecentFiles=False)
        if doc.Sections.Count != 1:
            raise ValueError('Prepared document has multiple Word sections; stopped safely.')
        progress('Making Word working-copy adjustments')
        normalize_caption_labels(doc,log,first_body_heading)
        apply_whole_motion_grid(doc,first_body_heading,enabled=align_grid,
                                has_28_line_header=number_header,
                                spacing_pt=style['pleading_grid']['line_spacing_pt'],log=log)
        calibrate_prepared_motion_numbers(doc,enabled=align_grid,
                        shift_pt=style['pleading_grid'].get('body_header_shift_pt',9.5),log=log)
        doc.Repaginate()
        doc.SaveAs2(str(output),FileFormat=DOCX)
        count=verify_motion_body_text_unchanged(source,output,first_body_heading)
        progress(f'Original motion body text preserved: {count} paragraphs')
        try:
            export_court_pdfa(doc,preview)
            if align_grid:
                verify_prepared_grid_preview(preview,first_body_heading,log)
            progress('Optional PDF/A preview generated (NOT FOR FILING): '+preview.name)
        except Exception as exc:
            if align_grid:
                raise RuntimeError('Pleading-grid Yes requires a verified PDF preview. '+str(exc)) from exc
            progress('PDF preview unavailable; Word working copy is still available: '+str(exc))
        doc.Close(SaveChanges=False);doc=None
        doc_hash=sha256_file(output)
        manifest={
            'schema_version':1,
            'prepared_filename':output.name,
            'prepared_sha256':doc_hash,
            'original_filename':source.name,
            'original_sha256':sha256_file(source),
            'first_body_heading':first_body_heading,
            'align_grid':bool(align_grid),
            'line_spacing_pt':float(style['pleading_grid']['line_spacing_pt']) if align_grid else None,
            'prepared_local_time':datetime.now().astimezone().isoformat(timespec='seconds'),
            'preview_filename':preview.name if preview.exists() else None,
            'grid_calibration':'Prepared Word preview passed a first-heading-to-numbered-line alignment check' if align_grid else 'Not requested',
            'body_header_shift_pt':float(style['pleading_grid'].get('body_header_shift_pt',9.5)) if align_grid else None
        }
        prepared_manifest_path(output).write_text(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
        (folder/'HipoLaw_PREPARATION_AUDIT.txt').write_text('\n'.join(log)+'\n',encoding='utf-8')
        return {'folder':folder,'prepared':output,'preview':preview if preview.is_file() else None,
                'manifest':prepared_manifest_path(output),'hash':doc_hash}
    except Exception:
        # Keep a failure log; never destroy source, and never create a success manifest.
        (folder/'HipoLaw_PREPARATION_ERROR.txt').write_text('\n'.join(log)+'\n'+traceback.format_exc(),encoding='utf-8')
        raise
    finally:
        if doc is not None:
            try:doc.Close(SaveChanges=False)
            except Exception:pass
        if word is not None:
            try:word.Quit()
            except Exception:pass


def build_chatgpt_prompt(docx_path, first_body_heading=None):
    source=Path(docx_path)
    if not source.is_file() or source.suffix.lower()!='.docx':
        raise ValueError('Choose the prepared Word document before opening ChatGPT.')
    manifest=load_prepared_manifest(source)
    template_file=external_file_path(CHATGPT_PROMPT_FILE)
    if not template_file.exists():
        raise FileNotFoundError(f'Missing editable ChatGPT prompt file alongside HipoLaw: {template_file}')
    template=template_file.read_text('utf-8-sig')
    details={
        'MOTION_FILENAME':source.name,
        'MOTION_SHA256':sha256_file(source),
        'FIRST_BODY_HEADING': manifest['first_body_heading'] if manifest else str(first_body_heading or '').strip(),
        'GRID_SETTING':'Yes' if manifest and manifest['align_grid'] else 'No',
    }
    if not details['FIRST_BODY_HEADING']:
        raise ValueError('First body heading is required. Prepare the motion or enter a heading first.')
    for k,v in details.items():
        template=template.replace('<<'+k+'>>',v)
    if re.search(r'<<[A-Z_]+>>',template):
        raise ValueError('Editable ChatGPT prompt contains unresolved placeholders.')
    return template


def preflight(source, excel):
    """Read-only checks. Unclear facts produce specific questions, not guessed output."""
    if not source.is_file() or source.suffix.lower()!='.docx' or not zipfile.is_zipfile(source):
        raise ValueError('Select an original Microsoft Word .docx motion.')
    if not excel.is_file() or excel.suffix.lower()!='.xlsx':
        raise ValueError('Select an AI-reviewed .xlsx workbook.')
    if openpyxl is None:raise RuntimeError('Preflight requires openpyxl: py -m pip install openpyxl')
    verify_spreadsheet_motion_binding(source, excel)
    wb=openpyxl.load_workbook(excel,read_only=True,data_only=False)
    try:
        for tab in ('Authorities','Contents'):
            if tab not in wb.sheetnames:raise ValueError(f'Missing required review tab: {tab}')
        auth=read_sheet(wb['Authorities'],'Authorities')
        toc=read_sheet(wb['Contents'],'Contents')
        alias=aliases_from_workbook(wb,auth)
        setup=read_setup(wb)
        grid_all=read_whole_motion_grid_option(setup)
    finally:wb.close()
    parts=document_paragraphs(source)
    paragraphs=parts[0][1]
    pnorm=[normalized_heading(x) for x in paragraphs]
    matches=[]
    for item in toc:
        target=normalized_heading(item['original'])
        stripped=re.sub(r'^(?:[ivxlcdm]+|\d+)[.)]\s*','',target,flags=re.I)
        found=[j for j,p in enumerate(pnorm) if p in (target,stripped)]
        if len(found)!=1:
            raise ValueError(f'CLARIFICATION NEEDED: Heading {item["original"]!r} matches {len(found)} Word paragraphs. '
                             'Which exact paragraph is the intended heading? Correct Contents column A or the motion.')
        matches.append((found[0],item))
    if matches!=sorted(matches,key=lambda x:x[0]):
        raise ValueError('Contents headings must appear in the same order as in the motion.')
    first=(setup.get('first body anchor') or setup.get('first body heading') or '').strip()
    if not first or first.upper() in ('REVIEW','UNKNOWN','UNCERTAIN','TBD'):
        suggestion=matches[0][1]['original']
        raise ValueError('CLARIFICATION NEEDED: Where exactly does the numbered motion body begin? '
                         f'Could it start at {suggestion!r}? Put the entire unique source paragraph in Setup!B3.')
    anchor=normalized_heading(first)
    stripped_anchor=re.sub(r'^(?:[ivxlcdm]+|\d+)[.)]\s*','',anchor,flags=re.I)
    anchor_hits=[j for j,p in enumerate(pnorm) if p in (anchor,stripped_anchor)]
    if len(anchor_hits)!=1:
        raise ValueError(f'CLARIFICATION NEEDED: Setup First body anchor {first!r} matches '
                         f'{len(anchor_hits)} readable source paragraphs. Use a distinctive full paragraph.')
    if anchor_hits[0]>matches[0][0]:
        raise ValueError('CLARIFICATION NEEDED: The first approved Contents heading occurs '
                         'BEFORE the declared start of the body. Confirm the caption/body boundary, '
                         'or remove pre-body headings from Contents.')
    text='\n'.join(x for _,ps in parts for x in ps)
    folded=text.casefold()
    unmatched=[]
    for item in auth:
        opts=variants_for_authority(item)+alias.get(item['original'].casefold(),[])
        if any(exact_safe_anchor_in_text(text,v) for v in opts):continue
        # Compound-citation possibilities are safe only with surrounding context.
        c=item['original']
        compound=[]
        if re.search(r'\bFed\.\s*R\.\s*Evid\.\s*802\b',c,re.I):
            compound=[r'801\(c\)\s*(?:,|and)\s*802\b']
        if re.search(r'\bFed\.\s*R\.\s*Evid\.\s*803\(10\)',c,re.I):
            compound=[r'803\(7\)\s*(?:or|,)\s*803\(10\)']
        if not any(re.search(pattern,text,re.I) for pattern in compound): unmatched.append(c)
    if unmatched:
        raise ValueError('CLARIFICATION NEEDED: Approved authority anchors were not found in the document. '
                         'Does the citation appear in a footnote, short form, or image? Correct column A or '
                         'add an exact occurrence on the Aliases sheet:\n'+'\n'.join(unmatched[:30]))
    if any(normalized_heading(x) in ('table of contents','table of authorities')
           for x in paragraphs):
        raise ValueError('CLARIFICATION NEEDED: Source already contains a Table of Contents or '
                         'Table of Authorities heading. Remove old tables from a COPY before running.')
    if setup.get('existing tables','No').strip().lower() not in ('no','none','false'):
        raise ValueError('CLARIFICATION NEEDED: Motion already contains a TOC or TOA. '
                         'Use a copy without those tables or confirm their replacement first.')
    # Fail safely if source has complex pre-existing section geometry.
    sects=count_source_sections(source)
    if sects>1:
        raise ValueError(f'CLARIFICATION NEEDED: Original document has {sects} Word sections. '
                         'This generator cannot safely preserve all section breaks yet. Ask for review or '
                         'supply a flattened single-section copy; do not silently alter original pagination.')
    grid_mode=setup.get('pleading grid','Auto').casefold()
    if grid_mode not in ('auto','none','28-line'):
        raise ValueError('Setup Pleading grid must be Auto, None, or 28-line.')
    grid_detected=has_numbered_pleading_header(source)
    if grid_mode=='28-line' and not grid_detected:
        raise ValueError('CLARIFICATION NEEDED: Setup says 28-line pleading paper, but the source does not '
                         'appear to have a 28-line numbered header. Correct Setup or supply a standard pleading copy.')
    grid=grid_detected if grid_mode=='auto' else grid_mode=='28-line'
    if grid_all and not grid_detected:
        raise ValueError('Setup whole-motion grid is Yes but the motion does not have a detected 28-line pleading header.')
    if grid_all and not grid:
        raise ValueError('Setup whole-motion grid is Yes but Pleading grid is None; choose Auto or 28-line.')
    print(f'PREFLIGHT: {len(auth)} approved authorities, {len(toc)} headings; '
          f'{len(alias)} authorities with aliases')
    print(f'PREFLIGHT: Motion type={setup["motion type"]!r}, body begins {first!r}; '
          f'numbered grid={grid}, whole-motion alignment={grid_all}, source sections={sects}')
    print('PREFLIGHT PASSED: all approved anchors appear in readable document text; '
          'Word pagination must still be checked on Windows')
    return auth,toc,alias,first,grid


def align_generated_body_header_numbers(doc, body_section, *, enabled, shift_pt, log):
    """Shift ONLY the printed 1–28 numbering baseline on the NEW body section.

    The source has a 24pt row rhythm but its actual header digits are ~9.5pt
    above the body text. Aligning paragraph rhythm alone is insufficient.
    Shift the number-list paragraph in the body section's unlinked header,
    leaving the caption, preliminary table-page header, horizontal margins,
    right pleading border and motion text unchanged. Never edit the input.
    """
    if not enabled:
        return 0
    if not -18 <= float(shift_pt) <= 18:
        raise ValueError('Body header grid shift must be within -18pt to +18pt.')
    header=doc.Sections(body_section).Headers(PRIMARY)
    header.LinkToPrevious=False
    paragraphs=list(header.Range.Paragraphs)
    # Depending on Word's nested header table implementation, paragraphs
    # may also be accessible through the table story; search both routes.
    try:
        for table in header.Range.Tables:
            paragraphs.extend(list(table.Range.Paragraphs))
    except Exception:
        pass
    for para in paragraphs:
        source=para.Range.Text.replace('\v',' ').replace('\r',' ').replace('\n',' ')
        parts=re.findall(r'(?<!\d)(?:[1-9]|1\d|2[0-8])(?!\d)',source)
        unique={int(n) for n in parts}
        if len(unique) < 24 or 1 not in unique or 28 not in unique:
            continue
        fmt=para.Range.ParagraphFormat
        before=float(fmt.SpaceBefore or 0)
        fmt.SpaceBefore=before+float(shift_pt)
        log.append(f'BODY NUMBER BASELINE: generated body-section header numbers 1–28 '
                   f'shifted {shift_pt:+g}pt; earlier pages and original source untouched')
        return 1
    raise RuntimeError('Cannot safely find the 28 printed pleading numbers in the GENERATED body '
                       'section header. No court-ready output produced. Choose No or supply a '
                       'compatible standard 28-line pleading template.')


# --- V16.4.6: spreadsheet-controlled whole-motion 28-line pleading alignment ---
WHOLE_MOTION_GRID_KEY = 'align entire motion to 28-line pleading grid'


def read_whole_motion_grid_option(setup):
    """Old V15/V16 spreadsheets continue to work (missing option means No).

    This setting controls source BODY paragraph spacing on the generated copy.
    It is not the pre-existing 'Pleading grid' switch, which controls whether
    the TOC/TOA uses the 28-line pleading layout.
    """
    raw = str(setup.get(WHOLE_MOTION_GRID_KEY, 'No')).strip().casefold()
    if raw not in ('yes', 'no'):
        raise ValueError('Setup \"Align entire motion to 28-line pleading grid\" must be Yes or No. '
                         'Use No to preserve the source motion spacing.')
    return raw == 'yes'


def apply_whole_motion_grid(doc, body_heading, *, enabled, has_28_line_header, spacing_pt, log):
    """Apply 24pt exact paragraph baseline spacing to body paragraphs only.

    This uses native Word paragraph formatting; it does not edit the caption,
    pleadings-paper geometry, list styles, numbering, headers, footers,
    Word field content or source file. It leaves footnotes and floating
    graphics alone because altering them can break notes or visual exhibits.
    """
    if not enabled:
        log.append('WHOLE MOTION GRID: No (original body paragraph formatting preserved)')
        return 0
    if not has_28_line_header:
        raise ValueError('Whole-motion alignment requires a detected 28-line pleading-paper header. '
                         'Please supply the numbered pleading DOCX, or choose No.')
    anchors = heading_ranges(doc, body_heading)
    if len(anchors) != 1:
        raise RuntimeError('Cannot apply body grid: first body heading is not unique.')
    start = anchors[0].Start
    # Word paragraph changes are confined to the main story. Footnotes,
    # captions and section headers/footers are never traversed.
    main_body = doc.Range(start, doc.Content.End)
    updated = 0
    skipped_tables = 0
    for paragraph in list(main_body.Paragraphs):
        rng = paragraph.Range
        if rng.Start < start:
            continue
        try:
            in_table = bool(rng.Information(12))  # wdWithInTable
        except Exception:
            in_table = False
        if in_table:
            skipped_tables += 1
            continue  # complex/exhibit tables must be audited visually
        fmt = rng.ParagraphFormat
        fmt.SpaceBefore = 0
        fmt.SpaceAfter = 0
        fmt.LineSpacingRule = WD_LINE_SPACE_EXACTLY
        fmt.LineSpacing = float(spacing_pt)
        updated += 1
    if updated < 1:
        raise RuntimeError('Whole-motion alignment changed no body paragraphs; stopped safely.')
    log.append(f'WHOLE MOTION GRID: Yes; {updated} body paragraphs set to exactly '
               f'{spacing_pt:g} pt line spacing and zero before/after; '
               f'{skipped_tables} paragraphs inside tables left unchanged')
    log.append('GRID LIMITATIONS: Footnotes, inline/floating graphics and complex tables '
               'are not forcibly reflowed; review all pagination and visual exhibits.')
    return updated


def audit_body_grid(pdf, body_idx, tolerance_pt=4.0):
    """Compare visible main-body lines with the source 28-number grid.

    This is a diagnostic, not proof of perfect Word line layout. Excludes
    footnotes (smaller than 11pt) and decorative/footer text. Report concise
    page summaries to avoid one warning per line.
    """
    findings = []
    sampled = 0
    if pymupdf is None:
        return ['Whole-motion grid PDF audit unavailable (PyMuPDF not installed).']
    for idx in range(body_idx, len(pdf)):
        page = pdf[idx]
        nums, body = [], []
        for block in page.get_text('dict')['blocks']:
            if block.get('type') != 0:
                continue
            for line in block.get('lines', []):
                for span in line.get('spans', []):
                    label = span['text'].strip()
                    if not label:
                        continue
                    x0,y0,x1,y1 = span['bbox']
                    center = (y0 + y1)/2
                    if 35 <= x0 <= 95 and re.fullmatch(r'(?:[1-9]|1\d|2[0-8])', label):
                        nums.append(center)
                    elif 115 <= x0 <= 590 and 65 <= y0 <= 733 and span['size'] >= 11:
                        if not ('PLAINTIFF' in label and y0 > 690):
                            body.append((center, label[:65]))
        if len({round(n, 1) for n in nums}) < 24:
            findings.append(f'REVIEW: Body PDF page {idx+1} has no recognizable 28-line grid.')
            continue
        if not body:
            continue
        deviations=[(min(abs(y-g) for g in nums),txt) for y,txt in body]
        outside=[(round(d,1),txt) for d,txt in deviations if d>tolerance_pt]
        sampled += len(deviations)
        if outside:
            worst=max(d for d,_ in outside)
            example=', '.join(repr(t) for _,t in outside[:2])
            findings.append(f'REVIEW: Body PDF page {idx+1}: {len(outside)}/{len(deviations)} '
                            f'text spans over {tolerance_pt:g}pt off grid (worst {worst:.1f}pt); {example}')
    if not findings:
        findings.append(f'BODY GRID CHECK: {sampled} sampled text spans within {tolerance_pt:g}pt '
                        'of the numbered grid (footnotes/images excluded).')
    return findings


def ensure_pagination_matches(issues):
    """Reject unresolved TOA/TOC page discrepancies; allow cosmetic warnings.

    The user explicitly chose to repaginate the entire motion. Page references
    must be verified after the LAST Word pagination and PDF export.
    """
    broken=[v for v in issues if v.startswith('REVIEW: ') and
            (' | TOA=' in v or 'missing or altered TOC label' in v)]
    if broken:
        raise RuntimeError('PAGINATION MISMATCH after whole-motion grid adjustment. '
                           'The generated PDF is NOT filing-ready; review the audit: '
                           + '; '.join(broken[:4]))


def run(source=None,excel=None, output_parent=None, folder_name=None, progress_callback=None):
    if win32 is None or openpyxl is None:
        raise RuntimeError('Full generator requires Windows, desktop Microsoft Word, pywin32 and openpyxl')
    if source is None:source=pick('Select ORIGINAL Word motion (.docx)','docx')
    if not source:return
    if excel is None:excel=pick('Select AI-approved review workbook (.xlsx)','xlsx')
    if not excel:return
    authorities,headings,aliases,body_heading,numbered_grid=preflight(source,excel)
    style=load_style_settings()
    preparation=load_prepared_manifest(source)
    setup_book=openpyxl.load_workbook(excel,read_only=True,data_only=False)
    try:
        grid_all=read_whole_motion_grid_option(read_setup(setup_book))
    finally:
        setup_book.close()
    ts=datetime.now().strftime('%Y%m%d_%H%M%S')
    if output_parent is None:
        output_parent = source.parent
    output_parent = Path(output_parent)
    if folder_name:
        clean_folder = re.sub(r'[^-._ a-zA-Z0-9]', '_', str(folder_name).strip())[:120].strip()
        if not clean_folder:
            raise ValueError('Output folder name cannot be blank')
    else:
        clean_folder = f'HipoLaw_V16_4_{ts}'
    folder = output_parent / clean_folder
    if folder.exists():
        raise FileExistsError(f'Output folder already exists: {folder}')
    folder.mkdir(parents=True,exist_ok=False)
    stem=re.sub(r'[^-._ a-zA-Z0-9]','_',source.stem)[:75]
    output=folder/(stem+'_With_Tables.docx')
    pdf=folder/(stem+'_With_Tables.pdf')
    report=folder/'HipoLaw_AUDIT.txt'
    audit_report=folder/'_PDF_CITATION_AUDIT.tmp'
    coverage_report=folder/'_CITATION_COVERAGE.tmp'
    shutil.copy2(source,output)
    started=time.monotonic()
    log=[f'ORIGINAL: {source}',f'APPROVALS: {excel}',f'OUTPUT: {output}',
         f'APPROVED: {len(authorities)} authorities, {len(headings)} TOC headings',
         'FIRST BODY HEADING: '+body_heading,
         'STYLES: '+str(style_file_path()),
         'GUIDED PREPARED DOCUMENT: '+('Yes — verified manifest and SHA256' if preparation else 'No — legacy/direct input'),
         'WHOLE MOTION GRID SETTING: '+('Yes' if grid_all else 'No'),
         'GENERATED LOCAL TIME: '+datetime.now().astimezone().isoformat(timespec='seconds'),
         'GENERATED UTC TIME: '+datetime.now(__import__('datetime').timezone.utc).isoformat(timespec='seconds')]
    def progress(message):
        line=f'[{time.monotonic()-started:7.1f}s] {message}'
        print(line,flush=True);log.append(line)
        report.write_text('\n'.join(log)+'\n',encoding='utf-8')
        if progress_callback is not None:
            try:
                progress_callback(line)
            except Exception:
                pass
    word=None;doc=None
    try:
        progress('Opening a separate Word application and the working copy')
        word=win32.DispatchEx('Word.Application')
        word.Visible=False;word.DisplayAlerts=0;word.ScreenUpdating=False
        doc=word.Documents.Open(str(output),ReadOnly=False,AddToRecentFiles=False)
        if doc.Sections.Count!=1:
            raise ValueError('Original document has multiple Word sections. V15 safely stops rather than changing existing section formatting.')
        original_pages=doc.ComputeStatistics(PAGE_COUNT)
        log.append(f'Original physical pages: {original_pages}')
        normalize_caption_labels(doc,log,body_heading)
        # Capture the ORIGINAL numbering before changing even paragraph spacing.
        heading_snapshot = capture_heading_numbering(doc, headings)
        if grid_all:
            progress('Aligning generated motion body to numbered 28-line pleading grid')
        if preparation and grid_all:
            log.append('GRID STEP 1: Paragraph spacing already prepared and SHA256 verified; no double formatting')
        else:
            apply_whole_motion_grid(doc,body_heading,enabled=grid_all,
                                    has_28_line_header=numbered_grid,
                                    spacing_pt=style['pleading_grid']['line_spacing_pt'],log=log)
        if grid_all:
            doc.Repaginate()
            check_heading_numbering(doc,heading_snapshot,'whole-motion grid formatting')
            log.append('RE-PAGINATION: whole-motion formatting applied BEFORE TOC/TOA markers and field generation')
        for h, (_,number,level) in zip(headings, heading_snapshot):
            h['toc_label'] = toc_display_label(number, h['final'])
            log.append(f'TOC LEVEL {level}: {number} {h["original"]} => {h["toc_label"]}')
        anchors=[];errors=[]
        progress(f'Locating {len(headings)} approved TOC headings')
        for h in headings:
            found=heading_ranges(doc,h['original'])
            if len(found)!=1:errors.append(f'{h["original"]!r}: found {len(found)} paragraphs')
            else:anchors.append((found[0].End-1,h))
        if errors:raise RuntimeError('Heading anchors are not unique: '+'; '.join(errors))
        for pos,h in sorted(anchors,key=lambda x:x[0],reverse=True):
            code=f'TC "{safe_tc_field_label(h["toc_label"])}" \\f C \\l {int(h["category"])}'
            add_field(doc,pos,code)
        progress('Inserting native Table of Authorities citation markers')
        if not insert_markers(doc,authorities,log,progress,aliases):
            raise RuntimeError('Some approved citations cannot be safely anchored; see HipoLaw_AUDIT.txt')
        progress('Inserting preliminary Word sections before first body heading')
        prelim,body_section=insert_preliminary_sections(doc,log,body_heading,authorities,style)
        if grid_all:
            progress('Calibrating printed pleading numbers to the motion body baseline')
        if preparation and grid_all:
            preserve_prepared_body_alignment(doc,prelim,body_section,
                     shift_pt=style['pleading_grid'].get('body_header_shift_pt',9.5),log=log)
        else:
            align_generated_body_header_numbers(doc,body_section,enabled=grid_all,
                     shift_pt=style['pleading_grid'].get('body_header_shift_pt',9.5),log=log)
        check_heading_numbering(doc,heading_snapshot,'preliminary insertion')
        doc.Save()
        fields_to_build=[('[GENERATE_TOC]',r'TOC \f C \h \z')]
        active={category_code(a['category']) for a in authorities}
        fields_to_build += [(placeholder,r'TOA \c '+str(cat))
                            for cat,placeholder in [(1,'[GENERATE_CASES]'),(2,'[GENERATE_STATUTES]'),
                                                    (3,'[GENERATE_OTHER]'),(4,'[GENERATE_RULES]')]
                            if cat in active]
        progress(f'Building {len(fields_to_build)} native TOC/TOA fields')
        for number,(placeholder,code) in enumerate(fields_to_build,1):
            progress(f'Field {number}/{len(fields_to_build)}: {placeholder}')
            replace_placeholder_with_field(doc,placeholder,code,progress,prelim)
            doc.Save()
        for f in list(doc.Fields):
            if f.Code.Text.strip().upper().startswith(('TOC ','TOA ')):
                progress('Updating '+f.Code.Text.strip()[:70]);f.Update()
        standardize_table_results(doc,log,progress,numbered_grid,prelim,style)
        check_heading_numbering(doc,heading_snapshot,'table field updates')
        compress_empty_preliminary_paragraphs(doc,log,progress,prelim)
        doc.Save();doc.Close(SaveChanges=False);doc=None
        _restore_cached_table_text_in_docx(output,log,progress,authorities,headings,numbered_grid)
        progress('Reopening corrected DOCX for Word pagination and PDF export')
        doc=word.Documents.Open(str(output),ReadOnly=False,AddToRecentFiles=False)
        verify_table_field_results(doc,log,progress,headings,authorities)
        check_heading_numbering(doc,heading_snapshot,'final Word reopen')
        for section_idx in range(1,doc.Sections.Count+1):
            doc.Sections(section_idx).Footers(PRIMARY).Range.Fields.Update()
        doc.Repaginate()
        check_heading_numbering(doc,heading_snapshot,'repagination')
        anchor=heading_ranges(doc,body_heading)
        if len(anchor)!=1:raise RuntimeError('First motion heading changed during generation')
        body_print=int(anchor[0].Information(ADJUSTED_PAGE))
        body_phys=int(anchor[0].Information(PHYSICAL_PAGE))
        if body_print!=1:
            raise RuntimeError(f'Body page numbering should restart at 1, got {body_print}')
        log.append(f'Motion begins: physical PDF page {body_phys}, printed page 1')
        progress('Saving Word file and exporting PDF')
        doc.SaveAs2(str(output),FileFormat=DOCX)
        protected_count=verify_motion_body_text_unchanged(source,output,body_heading)
        log.append(f'SOURCE TEXT PRESERVATION: {protected_count} original body paragraphs unchanged')
        export_court_pdfa(doc,pdf)
        progress('Checking Word PDF/A profile and CM/ECF unsafe features')
        filing_report,filing_warnings=pdfa_safety_audit(pdf)
        log.extend(filing_report)

        progress('Running independent, generic PDF audit')
        issues=pdf_validate(pdf,authorities,headings,body_phys,audit_report,coverage_report,
                            prelim_phys=prelim if prelim==1 else 2,numbered_grid=numbered_grid,aliases=aliases,
                            min_table_font_pt=style['entry_font_size_pt'])
        if grid_all:
            with pymupdf.open(str(pdf)) as generated_pdf:
                whole_grid_findings=audit_body_grid(generated_pdf,body_phys-1)
            log.extend(whole_grid_findings)
            # A failed page-reference comparison blocks successful completion.
            # Audit file remains available for diagnosis on exception.
            if any(msg.startswith('REVIEW:') for msg in whole_grid_findings):
                issues.extend(whole_grid_findings)
            ensure_pagination_matches(issues)
        log.append(f'PDF AUDIT: {len(issues)} layout/citation warning(s); {len(filing_warnings)} filing-verification warning(s)')
        log.append('LEGAL REVIEW REQUIRED: correctness/current treatment and ambiguous short forms are not certified')
        progress('Finished. Files are in '+str(folder))
        # Three deliverables only: editable DOCX, matching PDF, consolidated audit.
        citation_audit = audit_report.read_text(encoding='utf-8')
        coverage_audit = coverage_report.read_text(encoding='utf-8')
        report.write_text('\n'.join(log) + '\n\n' + '='*78 +
                          '\nPDF CITATION AUDIT\n' + citation_audit + '\n' + '='*78 +
                          '\nCITATION COVERAGE REVIEW\n' + coverage_audit,encoding='utf-8')
        audit_report.unlink()
        coverage_report.unlink()
        print('FINAL PDF AUDIT: '+('PASS (supported checks)' if not issues else f'REVIEW {len(issues)} issue(s)'),flush=True)
        if filing_warnings:
            print('PDF/A profile identified; full veraPDF validation not performed unless validator installed.',flush=True)
        return {'folder': folder, 'output': output, 'pdf': pdf, 'report': report, 'audit_report': report, 'coverage_report': report, 'issues': issues}
    except BaseException as exc:
        log.append('ERROR: '+repr(exc))
        log.append(traceback.format_exc())
        if grid_all and audit_report.exists():
            log.append('LAST PDF CITATION AUDIT (errors prevented filing-ready completion):')
            log.append(audit_report.read_text(encoding='utf-8'))
        report.write_text('\n'.join(log)+'\n',encoding='utf-8')
        raise
    finally:
        try:
            if doc is not None:doc.Close(SaveChanges=False)
        except Exception:pass
        try:
            if word is not None:word.Quit()
        except Exception:pass




def suggested_output_folder(source_path: Path) -> str:
    stem = re.sub(r'[^-._ a-zA-Z0-9]', '_', source_path.stem).strip(' ._') or 'motion'
    return f'{stem}_HipoLaw_Output'


def resource_path(*parts):
    base = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))
    return str(base.joinpath(*parts))


class HipoLawApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('HipoLaw V16.4.6 — Prepare · Review · Generate')
        self.geometry('1190x855')
        self.minsize(1030, 760)
        self.configure(bg='#f6f0ea')

        self.original_motion_var = tk.StringVar()
        self.first_heading_var = tk.StringVar()
        self.align_grid_var = tk.BooleanVar(value=False)
        self.prepared_var = tk.StringVar()
        self._last_prepared_result = None
        self.motion_var = tk.StringVar()
        self.review_var = tk.StringVar()
        self.save_dir_var = tk.StringVar()
        self.folder_var = tk.StringVar()
        self.status_var = tk.StringVar(value='Step 1: choose the motion and prepare a working copy.')
        self.hero_status_var = tk.StringVar(value='Ready to chew through documents.')
        self.folder_autofill = True
        self.is_running = False
        self.last_output_folder = None
        self._anim_tick = 0
        self._frame_index = 0
        self._status_dots = 0
        self.images = {}
        self._build_ui()
        self._load_mascot_assets()
        self._pulse_animation()
        self.motion_var.trace_add('write', self._motion_changed)
        self.original_motion_var.trace_add('write', self._original_motion_changed)

    def _style_audit(self):
        return {
            'bg': '#f6f0ea',
            'card': '#fffaf6',
            'lavender': '#efe3ff',
            'lavender_2': '#e4d2ff',
            'lavender_3': '#c8a6ff',
            'purple': '#6f46b5',
            'purple_dark': '#4c2c86',
            'text': '#3e3150',
            'muted': '#6e6078',
            'line': '#e3d6f0',
            'entry_bg': '#ffffff',
            'button_gold': '#f3bf42',
            'button_gold_dark': '#d39d27',
            'success_bg': '#fff4d9',
            'success_text': '#5d4a1f',
            'log_bg': '#ffffff',
            'log_border': '#d7c8eb',
        }

    def _build_ui(self):
        c=self._style_audit()
        self.colors=c
        self.option_add('*Font', '{Segoe UI} 10')
        outer=tk.Frame(self,bg=c['bg'])
        outer.pack(fill='both',expand=True,padx=14,pady=12)
        outer.grid_columnconfigure(0,weight=7)
        outer.grid_columnconfigure(1,weight=6)
        outer.grid_rowconfigure(1,weight=1)

        hero=tk.Frame(outer,bg=c['lavender'],highlightbackground=c['line'],highlightthickness=1)
        hero.grid(row=0,column=0,columnspan=2,sticky='ew',pady=(0,10))
        hero.grid_columnconfigure(0,weight=1)
        left=tk.Frame(hero,bg=c['lavender'])
        left.grid(row=0,column=0,sticky='w',padx=20,pady=11)
        tk.Label(left,text='HipoLaw',font=('Segoe UI',24,'bold'),bg=c['lavender'],fg=c['purple_dark']).pack(anchor='w')
        tk.Label(left,text='1. Prepare  →  2. ChatGPT spreadsheet  →  3. Generate court filing',
                 font=('Segoe UI',11,'bold'),bg=c['lavender'],fg=c['text']).pack(anchor='w',pady=(4,2))
        tk.Label(left,text='Original V15 Word engine · PDF/A output · Editable formatting · Hippo stays!',
                 bg=c['lavender'],fg=c['muted']).pack(anchor='w')
        self.hero_status=tk.Label(left,textvariable=self.hero_status_var,bg='#fff8e8',
                                  fg=c['success_text'],padx=10,pady=4)
        self.hero_status.pack(anchor='w',pady=(7,0))
        right=tk.Frame(hero,bg=c['lavender'])
        right.grid(row=0,column=1,padx=16,pady=7,sticky='e')
        self.mascot_label=tk.Label(right,bg=c['lavender'])
        self.mascot_label.pack()

        form_card=tk.Frame(outer,bg=c['card'],highlightbackground=c['line'],highlightthickness=1)
        form_card.grid(row=1,column=0,sticky='nsew',padx=(0,10))
        log_card=tk.Frame(outer,bg=c['card'],highlightbackground=c['line'],highlightthickness=1)
        log_card.grid(row=1,column=1,sticky='nsew')
        log_card.grid_columnconfigure(0,weight=1)
        log_card.grid_rowconfigure(0,weight=1)

        tk.Label(form_card,text='Guided motion workflow',font=('Segoe UI',15,'bold'),
                 bg=c['card'],fg=c['purple_dark']).pack(anchor='w',padx=16,pady=(12,5))
        tabs_style=ttk.Style(self)
        try:tabs_style.configure('Hipo.TNotebook.Tab',padding=(10,9),font=('Segoe UI',10,'bold'))
        except Exception:pass
        self.workflow_tabs=ttk.Notebook(form_card,style='Hipo.TNotebook')
        self.workflow_tabs.pack(fill='both',expand=True,padx=12,pady=(0,8))
        self.tab_prepare=tk.Frame(self.workflow_tabs,bg=c['card'])
        self.tab_chat=tk.Frame(self.workflow_tabs,bg=c['card'])
        self.tab_generate=tk.Frame(self.workflow_tabs,bg=c['card'])
        self.workflow_tabs.add(self.tab_prepare,text='1 · Prepare')
        self.workflow_tabs.add(self.tab_chat,text='2 · ChatGPT')
        self.workflow_tabs.add(self.tab_generate,text='3 · Generate')

        prep=tk.Frame(self.tab_prepare,bg=c['card'])
        prep.pack(fill='x',padx=9,pady=12)
        prep.grid_columnconfigure(0,weight=1)
        tk.Label(prep,text='Prepare a safe working copy',font=('Segoe UI',12,'bold'),
                 bg=c['card'],fg=c['purple_dark']).grid(row=0,column=0,columnspan=3,sticky='w')
        tk.Label(prep,text='Your original Word file is never overwritten. Choose the first motion heading.',
                 bg=c['card'],fg=c['muted'],wraplength=390,justify='left').grid(row=1,column=0,columnspan=3,sticky='w',pady=(2,10))
        self._make_field(prep,2,'Original motion (.docx)',self.original_motion_var,self.browse_original)
        self._make_field(prep,4,'First heading of motion body (confirm)',self.first_heading_var)
        self.align_box=tk.Checkbutton(prep,text='Align motion body to 28-line pleading paper',
                                      variable=self.align_grid_var,bg=c['card'],fg=c['text'],
                                      activebackground=c['card'],anchor='w')
        self.align_box.grid(row=6,column=0,columnspan=3,sticky='ew',pady=(14,5))
        self.prepare_btn=self._make_button(prep,'Prepare Word working copy',self.start_prepare,
                                            bg=c['purple'],fg='white')
        self.prepare_btn.grid(row=7,column=0,columnspan=3,sticky='ew',ipady=9,pady=5)
        self.preview_btn=self._make_button(prep,'Preview prepared document',self.preview_prepared,
                                            bg=c['button_gold'],fg='#4c3a05')
        self.preview_btn.grid(row=8,column=0,columnspan=3,sticky='ew',ipady=7,pady=5)
        self.prepared_label=tk.Label(prep,textvariable=self.prepared_var,wraplength=390,justify='left',
                                    bg=c['card'],fg=c['muted'])
        self.prepared_label.grid(row=9,column=0,columnspan=3,sticky='w',pady=(6,0))
        tk.Label(prep,text='Note: Final numbered-header calibration occurs after table pages are inserted.',
                 wraplength=390,justify='left',bg=c['card'],fg=c['muted']).grid(row=10,column=0,columnspan=3,sticky='w',pady=(10,0))

        chat=tk.Frame(self.tab_chat,bg=c['card'])
        chat.pack(fill='x',padx=12,pady=16)
        tk.Label(chat,text='Build your review spreadsheet',font=('Segoe UI',13,'bold'),
                 bg=c['card'],fg=c['purple_dark']).pack(anchor='w',pady=(0,12))
        tk.Label(chat,text=('HipoLaw copies the reusable spreadsheet prompt to your clipboard and '
                             'opens ChatGPT. You attach the prepared Word file yourself.'),
                 bg=c['card'],fg=c['text'],wraplength=390,justify='left').pack(anchor='w',pady=(0,13))
        self.chat_btn=self._make_button(chat,'Copy prompt + Open ChatGPT',self.open_chatgpt_prompt,
                                         bg=c['purple'],fg='white')
        self.chat_btn.pack(fill='x',ipady=11,pady=(0,10))
        self.copy_prompt_btn=self._make_button(chat,'Copy prompt only',self.copy_chatgpt_prompt,
                                                bg='#e9ddff',fg=c['purple_dark'])
        self.copy_prompt_btn.pack(fill='x',ipady=8,pady=5)
        tk.Label(chat,text=('In ChatGPT, upload the prepared .docx. Review every citation and heading '
                             'in the generated XLSX, then return to Step 3.'),
                 bg=c['card'],fg=c['muted'],wraplength=390,justify='left').pack(anchor='w',pady=(20,8))
        tk.Label(chat,text=('The prompt requests the exact V15 Authorities, Contents, Setup, '
                             'and Aliases sheets, plus a document fingerprint to prevent mismatches.'),
                 bg=c['card'],fg=c['muted'],wraplength=390,justify='left').pack(anchor='w')

        gen=tk.Frame(self.tab_generate,bg=c['card'])
        gen.pack(fill='both',expand=True,padx=9,pady=7)
        gen.grid_columnconfigure(0,weight=1)
        self._make_field(gen,0,'Prepared/selected motion (.docx)',self.motion_var,self.browse_motion)
        self._make_field(gen,2,'Approved review spreadsheet (.xlsx)',self.review_var,self.browse_review)
        self._make_field(gen,4,'Save output in folder',self.save_dir_var,self.browse_save_dir)
        self._make_field(gen,6,'Output folder name',self.folder_var,None,bind_edit=True)
        button_bar=tk.Frame(gen,bg=c['card'])
        button_bar.grid(row=8,column=0,columnspan=3,sticky='ew',pady=(15,5))
        for i in range(3):button_bar.grid_columnconfigure(i,weight=1)
        self.preflight_btn=self._make_button(button_bar,'Preflight',self.start_preflight,bg='#e9ddff',fg=c['purple_dark'])
        self.preflight_btn.grid(row=0,column=0,sticky='ew',padx=(0,4),ipady=8)
        self.run_btn=self._make_button(button_bar,'Generate filing',self.start_run,bg=c['purple'],fg='white')
        self.run_btn.grid(row=0,column=1,sticky='ew',padx=4,ipady=8)
        self.open_btn=self._make_button(button_bar,'Open output',self.open_last_output,bg=c['button_gold'],fg='#4c3a05')
        self.open_btn.grid(row=0,column=2,sticky='ew',padx=(4,0),ipady=8)
        tk.Label(gen,text='Final output: one DOCX, one PDF/A, one consolidated audit.',
                 wraplength=390,bg=c['card'],fg=c['muted']).grid(row=9,column=0,columnspan=3,sticky='w',pady=(7,0))

        self.styles_btn=self._make_button(form_card,'Edit styles or spreadsheet prompt (no rebuild)',
                                           self.edit_external_settings,bg='#efe6ff',fg=c['purple_dark'])
        self.styles_btn.pack(fill='x',padx=12,pady=(2,4),ipady=3)
        self.status_box=tk.Label(form_card,textvariable=self.status_var,anchor='w',justify='left',
                                 wraplength=490,bg=c['success_bg'],fg=c['success_text'],padx=12,pady=6)
        self.status_box.pack(fill='x',padx=12,pady=(0,5))

        log_inner=tk.Frame(log_card,bg=c['card'])
        log_inner.pack(fill='both',expand=True,padx=14,pady=12)
        tk.Label(log_inner,text='Processing status',font=('Segoe UI',15,'bold'),
                 bg=c['card'],fg=c['purple_dark']).pack(anchor='w')
        tk.Label(log_inner,text='Preparation, verification and filing progress.',
                 bg=c['card'],fg=c['muted']).pack(anchor='w',pady=(3,12))
        self.log=ScrolledText(log_inner,wrap='word',height=20,font=('Consolas',9),state='disabled',
                              relief='flat',bd=0,padx=10,pady=10,background=c['log_bg'],
                              foreground='#3a3150',insertbackground='#3a3150',highlightthickness=1,
                              highlightbackground=c['log_border'],highlightcolor=c['purple'])
        self.log.pack(fill='both',expand=True)
        tk.Label(log_inner,text='Always review legal citations, Word pagination and PDF/A before filing.',
                 bg=c['card'],fg=c['muted'],wraplength=390,justify='left').pack(anchor='w',pady=(10,0))
        self.append_log('Welcome to HipoLaw V16.4.6 — Prepare → ChatGPT spreadsheet → Generate filing.')
        self.append_log('The original motion is never overwritten. Use the prepared file in Steps 2 and 3.')

    def _make_field(self, parent, row, label_text, variable, browse_cmd=None, note=None, bind_edit=False):
        c = self.colors
        tk.Label(parent, text=label_text, font=('Segoe UI', 10, 'bold'), bg=c['card'], fg=c['text']).grid(row=row, column=0, columnspan=3, sticky='w', pady=(7, 3))
        entry_wrap = tk.Frame(parent, bg=c['entry_bg'], highlightbackground=c['line'], highlightthickness=1)
        entry_wrap.grid(row=row+1, column=0, sticky='ew', padx=(0, 8))
        entry_wrap.grid_columnconfigure(0, weight=1)
        entry = tk.Entry(entry_wrap, textvariable=variable, relief='flat', bd=0, bg=c['entry_bg'], fg=c['text'],
                         font=('Segoe UI', 10))
        entry.grid(row=0, column=0, sticky='ew', padx=10, pady=7)
        if bind_edit:
            entry.bind('<Key>', self._folder_manually_edited)
        if browse_cmd:
            btn = self._make_button(parent, 'Browse…', browse_cmd, bg='#efe6ff', fg=c['purple_dark'])
            btn.grid(row=row+1, column=1, sticky='ew', ipady=3)
        # Keep each field to two rows so Run stays visible on laptop screens.

    def _make_button(self, parent, text, command, bg, fg):
        btn = tk.Button(parent, text=text, command=command, relief='flat', bd=0,
                        bg=bg, fg=fg, activebackground=bg, activeforeground=fg,
                        font=('Segoe UI', 10, 'bold'), cursor='hand2', padx=12, pady=6)
        return btn

    def _load_mascot_assets(self):
        self.images['idle'] = tk.PhotoImage(file=resource_path('assets', 'mascot_frame1.png')).subsample(2,2)
        self.images['run1'] = tk.PhotoImage(file=resource_path('assets', 'mascot_frame1.png')).subsample(2,2)
        self.images['run2'] = tk.PhotoImage(file=resource_path('assets', 'mascot_frame2.png')).subsample(2,2)
        self.mascot_label.configure(image=self.images['idle'])

    def _original_motion_changed(self,*_args):
        value=self.original_motion_var.get().strip()
        if not value:return
        path=Path(value)
        if path.is_file():
            suggestion=suggest_first_body_heading(path)
            if suggestion and not self.first_heading_var.get().strip():
                self.first_heading_var.set(suggestion)
        if not self.motion_var.get().strip():
            self.motion_var.set(value)
        if not self.save_dir_var.get().strip():
            self.save_dir_var.set(str(path.parent))

    def browse_original(self):
        path=filedialog.askopenfilename(title='Select ORIGINAL Word motion',filetypes=[('Word files','*.docx')])
        if path:
            self.first_heading_var.set('')
            self.prepared_var.set('')
            self._last_prepared_result=None
            self.original_motion_var.set(path)
            self.motion_var.set(path)

    def start_prepare(self):
        if self.is_running:return
        original=Path(self.original_motion_var.get().strip())
        heading=self.first_heading_var.get().strip()
        if not original.is_file() or original.suffix.lower()!='.docx':
            messagebox.showerror('HipoLaw','Choose an existing original .docx motion.',parent=self);return
        if not heading:
            messagebox.showerror('HipoLaw','Enter and confirm the first body heading.',parent=self);return
        savedir=Path(self.save_dir_var.get().strip() or str(original.parent))
        if not savedir.is_dir():
            messagebox.showerror('HipoLaw','Choose an existing save folder.',parent=self);return
        self.set_running(True)
        self.append_log('')
        self.append_log('STEP 1 — Preparing a new motion copy')
        threading.Thread(target=self._prepare_worker,args=(original,heading,bool(self.align_grid_var.get()),savedir),daemon=True).start()

    def _prepare_worker(self,original,heading,align,savedir):
        try:
            import pythoncom
            pythoncom.CoInitialize()
            try:
                result=prepare_motion(original,heading,align_grid=align,output_parent=savedir,
                            progress_callback=lambda msg:self.after(0,self.append_log,msg))
            finally:
                pythoncom.CoUninitialize()
            self.after(0,self._prepare_success,result)
        except Exception as exc:
            details=traceback.format_exc()
            self.after(0,self.append_log,'PREPARE ERROR: '+details)
            self.after(0,self._worker_failure,str(exc))

    def _prepare_success(self,result):
        self.set_running(False)
        self._last_prepared_result=result
        self.prepared_var.set(str(result['prepared']))
        self.motion_var.set(str(result['prepared']))
        self.append_log('STEP 1 FINISHED: '+str(result['prepared']))
        self.status_var.set('Working copy ready. Preview it; then open ChatGPT in Step 2.')
        self.workflow_tabs.select(self.tab_chat)
        messagebox.showinfo('HipoLaw','Prepared working copy saved. Review it, then open ChatGPT in Step 2.\n\n'+str(result['prepared']),parent=self)

    def preview_prepared(self):
        source=(self._last_prepared_result or {}).get('prepared') or self.prepared_var.get().strip()
        if not source or not Path(source).is_file():
            messagebox.showinfo('HipoLaw','Prepare the Word working copy first.',parent=self);return
        preview=(self._last_prepared_result or {}).get('preview')
        target=Path(preview) if preview and Path(preview).is_file() else Path(source)
        try:
            if os.name=='nt':os.startfile(str(target))
            else:subprocess.Popen(['xdg-open',str(target)])
            self.append_log('PREVIEW: '+str(target))
        except Exception as exc:
            messagebox.showerror('HipoLaw',f'Cannot open preview: {exc}',parent=self)

    def _copy_prompt(self):
        target=Path(self.motion_var.get().strip())
        prompt=build_chatgpt_prompt(target,self.first_heading_var.get().strip())
        self.clipboard_clear()
        self.clipboard_append(prompt)
        self.update_idletasks()
        self.append_log('STEP 2: Reusable prompt copied for '+target.name)
        self.append_log('Motion SHA256: '+sha256_file(target))

    def copy_chatgpt_prompt(self):
        try:
            self._copy_prompt()
            messagebox.showinfo('HipoLaw','Prompt copied. Paste it into ChatGPT and attach your selected Word motion.',parent=self)
        except Exception as exc:messagebox.showerror('HipoLaw',str(exc),parent=self)

    def open_chatgpt_prompt(self):
        try:
            self._copy_prompt()
            if not webbrowser.open('https://chatgpt.com/'):
                self.append_log('Browser did not open automatically: https://chatgpt.com/')
            messagebox.showinfo('HipoLaw','ChatGPT opened (or use chatgpt.com). Paste the copied prompt, attach the prepared Word file, and download the review XLSX.',parent=self)
        except Exception as exc:messagebox.showerror('HipoLaw',str(exc),parent=self)

    def edit_external_settings(self):
        # Editable plain-text files live beside the executable, not inside the EXE.
        if os.name != 'nt':
            messagebox.showinfo('HipoLaw','Edit these files with any text editor:\n'+str(style_file_path())+'\n'+str(external_file_path(CHATGPT_PROMPT_FILE)),parent=self)
            return
        try:
            for path in (style_file_path(),external_file_path(CHATGPT_PROMPT_FILE)):
                if not path.is_file():raise FileNotFoundError(f'Missing external file: {path}')
                subprocess.Popen(['notepad.exe',str(path)])
                self.append_log('OPEN EDITABLE FILE: '+str(path))
        except Exception as exc:messagebox.showerror('HipoLaw',str(exc),parent=self)

    def _folder_manually_edited(self, _event=None):
        self.folder_autofill = False

    def _motion_changed(self, *_args):
        value = self.motion_var.get().strip()
        if not value:
            return
        source = Path(value)
        if self.folder_autofill or not self.folder_var.get().strip():
            self.folder_var.set(suggested_output_folder(source))
        if not self.save_dir_var.get().strip():
            self.save_dir_var.set(str(source.parent))

    def browse_motion(self):
        path = filedialog.askopenfilename(title='Select ORIGINAL Word motion (.docx)', filetypes=[('Word files', '*.docx')])
        if path:
            self.motion_var.set(path)

    def browse_review(self):
        path = filedialog.askopenfilename(title='Select AI-approved review workbook (.xlsx)', filetypes=[('Excel files', '*.xlsx')])
        if path:
            self.review_var.set(path)

    def browse_save_dir(self):
        path = filedialog.askdirectory(title='Choose where the output folder should be created')
        if path:
            self.save_dir_var.set(path)

    def append_log(self, message):
        self.log.configure(state='normal')
        self.log.insert('end', message.rstrip() + '\n')
        self.log.see('end')
        self.log.configure(state='disabled')

    def edit_style_settings(self):
        try:
            path=style_file_path()
            if not path.is_file():
                raise FileNotFoundError(f'Style file missing beside EXE: {path}')
            if os.name=='nt':
                import subprocess
                subprocess.Popen(['notepad.exe', str(path)])
            else:
                messagebox.showinfo('Hipo Law', f'Edit {path} in a text editor.', parent=self)
            self.append_log('EDIT STYLES: '+str(path))
            self.append_log('Changes take effect the next time Run generator starts; no rebuild needed.')
        except Exception as exc:
            messagebox.showerror('Hipo Law', f'Cannot open style file: {exc}', parent=self)

    def set_running(self, running):
        self.is_running = running
        state = 'disabled' if running else 'normal'
        for widget in (self.prepare_btn, self.preview_btn, self.chat_btn, self.copy_prompt_btn, self.preflight_btn, self.run_btn, self.open_btn, self.styles_btn):
            widget.configure(state=state)
        if running:
            self.status_var.set('Hippo is chewing through the documents… please wait.')
            self.hero_status_var.set('Chomp chomp… processing legal paperwork.')
        else:
            self.hero_status_var.set('Ready to chew through documents.')

    def validate_inputs(self):
        motion = self.motion_var.get().strip()
        review = self.review_var.get().strip()
        save_dir = self.save_dir_var.get().strip()
        folder = self.folder_var.get().strip()
        if not motion:
            raise ValueError('Choose the original motion .docx file.')
        if not review:
            raise ValueError('Choose the approved review workbook .xlsx file.')
        if not save_dir:
            raise ValueError('Choose the folder where the output folder should be created.')
        if not folder:
            raise ValueError('Output folder name cannot be blank.')
        motion_path = Path(motion)
        review_path = Path(review)
        save_dir_path = Path(save_dir)
        if motion_path.suffix.lower() != '.docx':
            raise ValueError('Motion document must be a .docx file.')
        if review_path.suffix.lower() != '.xlsx':
            raise ValueError('Review workbook must be an .xlsx file.')
        if not motion_path.exists():
            raise ValueError(f'Motion document not found: {motion_path}')
        if not review_path.exists():
            raise ValueError(f'Review workbook not found: {review_path}')
        if not save_dir_path.exists() or not save_dir_path.is_dir():
            raise ValueError(f'Save folder not found: {save_dir_path}')
        return motion_path, review_path, save_dir_path, folder

    def start_preflight(self):
        self._start_worker(mode='preflight')

    def start_run(self):
        self._start_worker(mode='run')

    def _start_worker(self, mode='run'):
        if self.is_running:
            return
        try:
            motion_path, review_path, save_dir_path, folder = self.validate_inputs()
        except Exception as exc:
            messagebox.showerror('Hipo Law', str(exc), parent=self)
            return
        self.last_output_folder = None
        self.set_running(True)
        self.status_var.set('Starting…')
        self.append_log('')
        self.append_log('=' * 78)
        self.append_log(('RUN GENERATOR' if mode == 'run' else 'PREFLIGHT ONLY') + f' — {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        worker = threading.Thread(target=self._worker_body, args=(mode, motion_path, review_path, save_dir_path, folder), daemon=True)
        worker.start()

    def _worker_body(self, mode, motion_path, review_path, save_dir_path, folder):
        try:
            if mode == 'preflight':
                auth, toc, alias, first, grid = preflight(motion_path, review_path)
                summary = (f'Preflight passed: {len(auth)} authorities, {len(toc)} headings, '
                           f'{len(alias)} aliased authorities, first body heading {first!r}, numbered grid={grid}.')
                self.after(0, lambda: self.append_log(summary))
                self.after(0, lambda: self._worker_success(None, preflight_only=True))
            else:
                # COM must be initialized in the worker thread before using Word.
                import pythoncom
                pythoncom.CoInitialize()
                try:
                    result = run(motion_path, review_path, output_parent=save_dir_path, folder_name=folder,
                                 progress_callback=lambda line: self.after(0, self.append_log, line))
                finally:
                    pythoncom.CoUninitialize()
                self.after(0, lambda: self._worker_success(result, preflight_only=False))
        except Exception as exc:
            details = traceback.format_exc()
            self.after(0, self.append_log, 'ERROR: ' + str(exc))
            self.after(0, self.append_log, details)
            error_message = str(exc)
            self.after(0, lambda message=error_message: self._worker_failure(message))

    def _worker_success(self, result, preflight_only=False):
        self.set_running(False)
        if preflight_only:
            self.status_var.set('Preflight passed. No files were generated.')
            self.hero_status_var.set('Preflight done. Hippo is ready for the real run.')
            messagebox.showinfo('Hipo Law', 'Preflight passed. No files were generated.', parent=self)
            return
        self.last_output_folder = result['folder'] if result else None
        warning_count = len(result.get('issues', [])) if result else 0
        if warning_count:
            self.status_var.set(f'Generation finished with {warning_count} audit warning(s). Review the output carefully.')
            self.hero_status_var.set('Finished chewing — please inspect the polished results.')
            messagebox.showwarning('Hipo Law', f'Generation finished with {warning_count} audit warning(s).\nPlease review all outputs before filing.', parent=self)
        else:
            self.status_var.set('Generation finished. Review the Word, PDF, and audit reports before filing.')
            self.hero_status_var.set('Finished chewing — clean results are ready for review.')
            messagebox.showinfo('Hipo Law', 'Generation finished successfully. Review the Word, PDF, and audit reports before filing.', parent=self)
        if self.last_output_folder:
            self.append_log('OUTPUT FOLDER: ' + str(self.last_output_folder))

    def _worker_failure(self, message):
        self.set_running(False)
        self.status_var.set('Stopped. Fix the issue and try again.')
        self.hero_status_var.set('Hippo hit a snag. Fix the issue and try another bite.')
        messagebox.showerror('Hipo Law', f'Generation stopped:\n\n{message}', parent=self)

    def open_last_output(self):
        if not self.last_output_folder:
            messagebox.showinfo('Hipo Law', 'No generated output folder has been created in this session yet.', parent=self)
            return
        folder = Path(self.last_output_folder)
        if not folder.exists():
            messagebox.showerror('Hipo Law', f'Output folder not found: {folder}', parent=self)
            return
        try:
            os.startfile(str(folder))
        except Exception as exc:
            messagebox.showerror('Hipo Law', f'Could not open folder:\n{exc}', parent=self)

    def _pulse_animation(self):
        self._anim_tick += 1
        if self.is_running:
            self._frame_index = (self._frame_index + 1) % 2
            self._status_dots = (self._status_dots + 1) % 4
            self.mascot_label.configure(image=self.images['run1'] if self._frame_index == 0 else self.images['run2'])
            self.hero_status_var.set('Chomp chomp' + '.' * self._status_dots + ' processing legal paperwork')
        else:
            self.mascot_label.configure(image=self.images['idle'])
        self.after(220, self._pulse_animation)

def launch_gui():
    app = HipoLawApp()
    app.mainloop()


def main():
    parser=argparse.ArgumentParser(description='Hipo Law V16.4.6 Guided Legal Tables Generator (V15.1 Word engine)')
    parser.add_argument('--preflight',nargs=2,metavar=('MOTION.docx','REVIEW.xlsx'),
                        help='Read-only validation; works without Windows Word')
    parser.add_argument('--run',nargs=2,metavar=('MOTION.docx','REVIEW.xlsx'),
                        help='Generate using installed Microsoft Word on Windows')
    parser.add_argument('--prepare', nargs=2, metavar=('MOTION.docx','FIRST_HEADING'),
                        help='Step 1: prepare a working Word copy (add --align for grid formatting)')
    parser.add_argument('--align', action='store_true', help='With --prepare, align body to 28-line grid')
    parser.add_argument('--gui', action='store_true', help='Launch the Hipo Law graphical interface')
    args=parser.parse_args()
    if args.prepare:
        prepare_motion(Path(args.prepare[0]),args.prepare[1],align_grid=args.align)
    elif args.preflight:
        preflight(Path(args.preflight[0]),Path(args.preflight[1]))
    elif args.run:
        run(Path(args.run[0]),Path(args.run[1]))
    elif args.gui or len(sys.argv) == 1:
        launch_gui()
    else:
        run()


if __name__=='__main__':
    try:
        main()
    except Exception as exc:
        details = traceback.format_exc()
        print('ERROR:', exc, flush=True)
        try:
            error_path = Path.home() / 'HipoLaw_startup_error.txt'
            error_path.write_text(details, encoding='utf-8')
            if sys.platform == 'win32' and not (('--preflight' in sys.argv) or ('--run' in sys.argv) or ('--prepare' in sys.argv)):
                # Windowed PyInstaller builds have no visible console. Show the error.
                import ctypes
                ctypes.windll.user32.MessageBoxW(
                    None, f'Hipo Law could not start:\n\n{exc}\n\nDetails saved to:\n{error_path}',
                    'Hipo Law startup error', 0x10)
        except Exception:
            print(details, flush=True)
        if sys.stdin and sys.stdin.isatty():
            try:input('Press Enter to close...')
            except EOFError:pass
        sys.exit(1)
