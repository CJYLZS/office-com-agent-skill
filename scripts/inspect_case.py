"""Bulk live OOXML inspection; no paragraph-by-paragraph COM traversal."""
import argparse
import json
import sys
import time
import xml.etree.ElementTree as ET
from pathlib import Path
from word_com import WordBridge

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'pkg': 'http://schemas.microsoft.com/office/2006/xmlPackage'}
W = '{' + NS['w'] + '}'

def parse_structure(xml):
    root = ET.fromstring(xml)
    styles = {}
    for style in root.findall('.//w:styles/w:style', NS):
        sid = style.get(W+'styleId')
        name = style.find('w:name', NS)
        outline = style.find('w:pPr/w:outlineLvl', NS)
        based = style.find('w:basedOn', NS)
        styles[sid] = {'name': name.get(W+'val') if name is not None else sid,
                       'level': int(outline.get(W+'val')) if outline is not None else None,
                       'based': based.get(W+'val') if based is not None else None}
    def style_level(sid, seen=None):
        seen = set() if seen is None else seen
        if sid in seen or sid not in styles:
            return None
        seen.add(sid)
        item = styles[sid]
        return item['level'] if item['level'] is not None else style_level(item['based'], seen)
    body = root.find('.//w:document/w:body', NS)
    if body is None:
        raise ValueError('WordOpenXML 中没有正文')
    rows = []
    for index, p in enumerate(body.findall('.//w:p', NS), 1):
        text = ''.join(e.text or '' for e in p.iter(W+'t'))
        ps = p.find('w:pPr/w:pStyle', NS)
        sid = ps.get(W+'val') if ps is not None else None
        outline = p.find('w:pPr/w:outlineLvl', NS)
        level = int(outline.get(W+'val')) if outline is not None else style_level(sid)
        numbering = p.find('w:pPr/w:numPr', NS)
        row = {'xml_paragraph_index': index, 'text': text, 'style_id': sid,
               'style': styles.get(sid, {}).get('name'), 'outline_level': level}
        if numbering is not None:
            row['numbering'] = {e.tag.split('}')[-1]:e.get(W+'val') for e in numbering}
        rows.append(row)
    return rows


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--pid', type=int, required=True)
    parser.add_argument('--hwnd', type=int, required=True)
    parser.add_argument('--output', type=Path, default=Path('artifacts/structure.json'))
    parser.add_argument('--xml-output', type=Path, default=Path('artifacts/live-document.xml'))
    parser.add_argument('--worker', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()
    if not args.worker:
        # Parent bounds the entire child, including a blocked COM call. Read-only only.
        import subprocess
        try:
            result = subprocess.run([sys.executable, str(Path(__file__).resolve()), *sys.argv[1:], '--worker'], timeout=9)
            return result.returncode
        except subprocess.TimeoutExpired:
            print(json.dumps({'error':'检查超过 9 秒，已终止只读检查进程；未执行文档修改'}, ensure_ascii=False), file=sys.stderr)
            return 2
    t0 = time.perf_counter()
    bridge = WordBridge()
    doc, identity = bridge.select(args.pid, args.hwnd)
    t1 = time.perf_counter()
    xml = str(doc.Content.WordOpenXML)
    t2 = time.perf_counter()
    rows = parse_structure(xml)
    t3 = time.perf_counter()
    report = {'target': identity, 'source': 'live COM Range.WordOpenXML',
              'note': 'XML 段落索引不是 Word Range 坐标；自动编号需局部 COM 核验',
              'paragraph_count': len(rows), 'headings': [r for r in rows if r['outline_level'] is not None and r['outline_level'] < 9],
              'paragraphs': rows,
              'timing_seconds': {'attach': round(t1-t0,3), 'bulk_xml': round(t2-t1,3), 'local_parse':round(t3-t2,3), 'total':round(t3-t0,3)}}
    for path in (args.output, args.xml_output):
        path.parent.mkdir(parents=True, exist_ok=True)
    args.xml_output.write_text(xml, encoding='utf-8')
    args.output.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k:v for k,v in report.items() if k != 'paragraphs'}, ensure_ascii=False, indent=2))
    return 0

if __name__ == '__main__':
    raise SystemExit(main())
