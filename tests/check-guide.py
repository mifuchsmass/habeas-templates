#!/usr/bin/env python3
"""Check the editable Word guide against the HTML guide and its current links."""
from pathlib import Path
from html.parser import HTMLParser
from urllib.parse import unquote, urljoin
import zipfile
import xml.etree.ElementTree as E

ROOT = Path(__file__).resolve().parent.parent
BASENAME = 'Word Template Automation Guide for Conditional Text'
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

class GuideHTML(HTMLParser):
    def __init__(self):
        super().__init__()
        self.stack, self.blocks, self.links = [], [], []
        self.tables = 0

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            self.links.append(dict(attrs))
        if tag == 'table':
            self.tables += 1
        if tag in ('p', 'h1', 'h2', 'h3', 'td', 'th'):
            for parent in self.stack:
                parent['child'] = True
            self.stack.append({'tag': tag, 'text': '', 'child': False})

    def handle_data(self, text):
        for block in self.stack:
            block['text'] += text

    def handle_endtag(self, tag):
        if tag in ('p', 'h1', 'h2', 'h3', 'td', 'th'):
            block = self.stack.pop()
            assert block['tag'] == tag, 'Unexpected guide block nesting'
            if not block['child']:
                self.blocks.append(block['text'])

html_path = ROOT / 'documentation' / (BASENAME + '.html')
word_path = html_path.with_suffix('.docx')
html = GuideHTML()
html.feed(html_path.read_text(encoding='utf-8'))
assert not html.stack
with zipfile.ZipFile(word_path) as z:
    assert z.testzip() is None
    document = E.fromstring(z.read('word/document.xml'))
    text = [''.join(t.text or '' for t in p.iter(W + 't')) for p in document.iter(W + 'p')]
    assert text == [p for p in html.blocks if p != 'Download Word version (.docx)'], 'Word/HTML wording differs'
    assert text[0] == 'HABEAS PETITION GENERATOR — WORD TEMPLATE AUTHORING GUIDE'
    assert text[1] == 'Updated October 13, 2026'
    assert html.tables == len(list(document.iter(W + 'tbl'))) == 2
    for name in z.namelist():
        if name == 'word/document.xml' or name.startswith(('word/header', 'word/footer')) and name.endswith('.xml'):
            contents = ''.join(E.fromstring(z.read(name)).itertext()).lower()
            assert not any(term in contents for term in ('review draft', 'review-draft', 'attorney review draft', 'not published', 'unpublished'))
    tables = list(document.iter(W + 'tbl'))
    rows = [[ ''.join(t.text or '' for t in cell.iter(W + 't')) for cell in row.findall(W + 'tc')] for row in tables[0].findall(W + 'tr')]
    assert rows == [
        ['Abbreviated name', 'Full name inserted into the Word document'],
        ['Plymouth', 'Plymouth County Correctional Facility'],
        ['Burlington', 'Burlington ICE Office'],
        ['Boston', 'Boston Field Office'],
        ['Berlin', 'FCI Berlin'],
        ['Strafford', 'Strafford County House of Corrections'],
        ['Wyatt', 'Donald W. Wyatt Detention Facility'],
        ['Northwest', 'Northwest State Correctional Facility']
    ]
assert not html_path.with_suffix('.txt').exists()
assert any('You must include spaces before and after each comparison symbol.' in p for p in text)
assert any('The application generally inserts answers as entered.' in p for p in text)
word_link = next(a for a in html.links if a.get('href', '').endswith('.docx'))
assert 'download' in word_link
assert html_path.parent / unquote(word_link['href']) == word_path
studio = GuideHTML()
studio.feed((ROOT / 'drafting-studio.html').read_text())
guide_link = next(a for a in studio.links if a.get('href', '').endswith('Conditional%20Text.html'))
assert guide_link['target'] == '_blank' and 'noopener' in guide_link['rel']
assert ROOT / unquote(guide_link['href']) == html_path
base = 'https://mifuchsmass.github.io/habeas-templates/'
assert urljoin(base + 'drafting-studio.html', guide_link['href']) == base + guide_link['href']
assert urljoin(base + guide_link['href'], word_link['href']) == base + 'documentation/' + word_link['href']
print(f'PASS {len(text)} identical substantive paragraphs/cells, date, no draft notices, two tables, seven facility names, and guide/download links.')
