#!/usr/bin/env python3
"""Extract paragraph manifest from document.xml / footnotes.xml / endnotes.xml."""
import zipfile, re, json, sys
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}
W = NS['w']

def qn(tag):
    return f'{{{W}}}{tag}'

def get_paragraphs(root):
    """Return all w:p elements in document order, including inside tables, excluding those in headers/footers (root passed is already the right part)."""
    return root.findall('.//' + qn('p'))

def para_runs_info(p):
    """Return list of run dicts: {text, rPr_xml, elem} for w:r children directly under p (also handles w:hyperlink wrapping runs by descending)."""
    runs = []
    # iterate direct text-bearing runs, in document order, including ones inside w:hyperlink/w:smartTag
    for r in p.iter(qn('r')):
        # skip runs inside footnoteReference-only / nested differently is fine since w:r itself holds w:t
        t_elems = r.findall(qn('t'))
        text = ''.join(t.text or '' for t in t_elems)
        if t_elems:
            runs.append({'elem': r, 'text': text})
    return runs

def para_text(p):
    return ''.join(t.text or '' for t in p.iter(qn('t')))

def has_noneditable_content(p):
    """Flag paragraphs containing fields, drawings, equations etc - still editable but flag for caution."""
    tags = set(el.tag for el in p.iter())
    flags = []
    for tag in ('fldSimple', 'drawing', 'object', 'oMath', 'oMathPara'):
        if qn(tag) in tags:
            flags.append(tag)
    return flags

def extract(docx_path, out_manifest, out_dir):
    z = zipfile.ZipFile(docx_path)
    parts = {
        'document': 'word/document.xml',
        'footnotes': 'word/footnotes.xml',
        'endnotes': 'word/endnotes.xml',
    }
    manifest = []
    pid = 0
    import os
    os.makedirs(out_dir, exist_ok=True)
    for partname, path in parts.items():
        if path not in z.namelist():
            continue
        data = z.read(path)
        root = etree.fromstring(data)
        # body paragraphs: for document.xml, restrict to <w:body> descendants (excludes nothing extra anyway)
        paras = get_paragraphs(root)
        for idx, p in enumerate(paras):
            text = para_text(p)
            if text.strip() == '':
                continue
            flags = has_noneditable_content(p)
            manifest.append({
                'id': pid,
                'part': partname,
                'idx_in_part': idx,
                'text': text,
                'nwords': len(text.split()),
                'flags': flags,
                'status': 'pending',
            })
            pid += 1
    with open(out_manifest, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    print(f'Extracted {len(manifest)} paragraphs to {out_manifest}')

if __name__ == '__main__':
    extract('dissertation.docx', 'manifest.json', 'batches')
