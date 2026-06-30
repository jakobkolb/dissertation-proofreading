#!/usr/bin/env python3
"""Apply proofreading corrections as real Word Track-Changes (w:ins/w:del),
preserving formatting by splitting exactly at original run boundaries.
"""
import zipfile, re, json, sys, shutil, difflib, datetime, os
from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NSMAP = {'w': W}

def qn(tag):
    return f'{{{W}}}{tag}'

AUTHOR = 'Claude Lektorat'
DATE = datetime.datetime.now(datetime.timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

_rev_id_counter = [1000000]
def next_rev_id():
    _rev_id_counter[0] += 1
    return _rev_id_counter[0]

TOKEN_RE = re.compile(r'\w+|[^\w\s]|\s+', re.UNICODE)

def tokenize(text):
    return TOKEN_RE.findall(text)

def clone_run_with_text(run_elem, text):
    new_r = etree.Element(qn('r'))
    rpr = run_elem.find(qn('rPr'))
    if rpr is not None:
        new_r.append(etree.fromstring(etree.tostring(rpr)))
    t = etree.SubElement(new_r, qn('t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return new_r

def make_del_run(run_elem, text):
    new_r = etree.Element(qn('r'))
    rpr = run_elem.find(qn('rPr'))
    if rpr is not None:
        new_r.append(etree.fromstring(etree.tostring(rpr)))
    dt = etree.SubElement(new_r, qn('delText'))
    dt.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    dt.text = text
    return new_r

def wrap_ins(run_list):
    ins = etree.Element(qn('ins'))
    ins.set(qn('id'), str(next_rev_id()))
    ins.set(qn('author'), AUTHOR)
    ins.set(qn('date'), DATE)
    for r in run_list:
        ins.append(r)
    return ins

def wrap_del(run_list):
    de = etree.Element(qn('del'))
    de.set(qn('id'), str(next_rev_id()))
    de.set(qn('author'), AUTHOR)
    de.set(qn('date'), DATE)
    for r in run_list:
        de.append(r)
    return de

def get_runs_with_text(p):
    """Direct-order list of (run_elem, parent_elem, text) for every w:r containing w:t, in document order."""
    out = []
    for r in p.iter(qn('r')):
        t_elems = r.findall(qn('t'))
        if not t_elems:
            continue
        text = ''.join(t.text or '' for t in t_elems)
        out.append((r, r.getparent(), text))
    return out

def build_segments(orig_tokens, new_tokens):
    """Return list of (kind, orig_start_char, orig_end_char, new_text) in original-char coordinates."""
    sm = difflib.SequenceMatcher(a=orig_tokens, b=new_tokens, autojunk=False)
    segments = []
    orig_pos = 0
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        orig_chunk = ''.join(orig_tokens[i1:i2])
        new_chunk = ''.join(new_tokens[j1:j2])
        start = orig_pos
        end = orig_pos + len(orig_chunk)
        if tag == 'equal':
            segments.append(('equal', start, end, orig_chunk))
        elif tag == 'delete':
            segments.append(('delete', start, end, ''))
        elif tag == 'insert':
            segments.append(('insert', start, start, new_chunk))
        elif tag == 'replace':
            segments.append(('replace', start, end, new_chunk))
        orig_pos = end
    return segments

def apply_paragraph(p, corrected_text):
    runs = get_runs_with_text(p)
    if not runs:
        return False
    orig_text = ''.join(t for _, _, t in runs)
    if orig_text == corrected_text:
        return False
    orig_tokens = tokenize(orig_text)
    new_tokens = tokenize(corrected_text)
    segments = build_segments(orig_tokens, new_tokens)

    # compute run boundaries in char coords
    run_bounds = []
    pos = 0
    for r, parent, text in runs:
        run_bounds.append((pos, pos + len(text)))
        pos += len(text)

    # for each run, collect replacement element list
    replacements = {}  # id(run_elem) -> list of elements
    ins_emitted_for_segment = set()  # segment index already emitted its ins text

    for ridx, (r, parent, text) in enumerate(runs):
        rstart, rend = run_bounds[ridx]
        elems = []
        for sidx, (kind, s, e, new_text) in enumerate(segments):
            if kind == 'insert':
                # attach pure insertion at the run whose start == s (insert *before* this run)
                if s == rstart:
                    elems.append(wrap_ins([clone_run_with_text(r, new_text)]))
                continue
            # overlap with this run
            ov_start = max(s, rstart)
            ov_end = min(e, rend)
            if ov_start >= ov_end:
                continue
            local_s = ov_start - rstart
            local_e = ov_end - rstart
            piece = text[local_s:local_e]
            if kind == 'equal':
                elems.append(clone_run_with_text(r, piece))
            elif kind == 'delete':
                elems.append(wrap_del([make_del_run(r, piece)]))
            elif kind == 'replace':
                elems.append(wrap_del([make_del_run(r, piece)]))
                if sidx not in ins_emitted_for_segment:
                    elems.append(wrap_ins([clone_run_with_text(r, new_text)]))
                    ins_emitted_for_segment.add(sidx)
        replacements[id(r)] = (parent, elems)

    # handle a trailing pure-insert segment that starts at end of paragraph (s == total length)
    total_len = run_bounds[-1][1] if run_bounds else 0
    for sidx, (kind, s, e, new_text) in enumerate(segments):
        if kind == 'insert' and s == total_len and total_len > 0:
            last_r, last_parent, last_text = runs[-1]
            replacements[id(last_r)][1].append(wrap_ins([clone_run_with_text(last_r, new_text)]))

    # apply replacements: for each run, replace it in its parent with the new element list
    for r, parent, text in runs:
        parent_elems = replacements[id(r)][1]
        idx = list(parent).index(r)
        parent.remove(r)
        for offset, el in enumerate(parent_elems):
            parent.insert(idx + offset, el)
    return True

def ensure_track_changes(settings_root):
    if settings_root.find(qn('trackChanges')) is None:
        tc = etree.Element(qn('trackChanges'))
        settings_root.insert(0, tc)

def main(docx_in, docx_out, corrections_path, manifest_path):
    with open(corrections_path, encoding='utf-8') as f:
        corrections = json.load(f)  # { "id": "corrected text", ... }
    with open(manifest_path, encoding='utf-8') as f:
        manifest = json.load(f)
    by_id = {m['id']: m for m in manifest}

    z = zipfile.ZipFile(docx_in)
    names = z.namelist()
    contents = {n: z.read(n) for n in names}

    part_files = {
        'document': 'word/document.xml',
        'footnotes': 'word/footnotes.xml',
        'endnotes': 'word/endnotes.xml',
    }
    part_roots = {}
    part_paras = {}
    for part, path in part_files.items():
        if path in contents:
            root = etree.fromstring(contents[path])
            part_roots[part] = root
            part_paras[part] = root.findall('.//' + qn('p'))

    applied = 0
    skipped_unchanged = 0
    for sid, corrected_text in corrections.items():
        pid = int(sid)
        if pid not in by_id:
            print(f'WARN: id {pid} not in manifest, skipping')
            continue
        m = by_id[pid]
        part = m['part']
        idx = m['idx_in_part']
        p = part_paras[part][idx]
        changed = apply_paragraph(p, corrected_text)
        if changed:
            applied += 1
            m['status'] = 'edited'
        else:
            skipped_unchanged += 1
            m['status'] = 'reviewed_no_change'

    for part, path in part_files.items():
        if part in part_roots:
            contents[path] = etree.tostring(part_roots[part], xml_declaration=True, encoding='UTF-8', standalone=True)

    if 'word/settings.xml' in contents:
        sroot = etree.fromstring(contents['word/settings.xml'])
        ensure_track_changes(sroot)
        contents['word/settings.xml'] = etree.tostring(sroot, xml_declaration=True, encoding='UTF-8', standalone=True)

    with zipfile.ZipFile(docx_out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents[n])

    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)

    print(f'Applied {applied} edits, {skipped_unchanged} unchanged, out -> {docx_out}')

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4])
