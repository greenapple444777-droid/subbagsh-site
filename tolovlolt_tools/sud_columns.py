"""Хэрэглээ: python3 sud_columns.py 7ХОНОГ.docx
7 хоногийн хуудасны СҮД жагсаалтыг хүрээгүй хүснэгтэд ангилал тус бүрийг (Шинэ | Уялдаа | Давталт) тусдаа баганад байрлуулна."""
import sys, zipfile, re
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from lxml import etree
from split_weeks import W, txt
ITEM = re.compile(r'^\s*[А-ЯӨҮЁ]{2,4}\d+\.\d+[а-яөүё]{0,2}\.\s')
LAB = re.compile(r'^\s*(Уялдаа|Давталт) СҮД:\s*$')
f = sys.argv[1]
z = zipfile.ZipFile(f); parts = {n: z.read(n) for n in z.namelist()}
doc = etree.fromstring(parts['word/document.xml']); body = doc.find(W + 'body')
sect = body[-1]
avail = int(sect.find(W + 'pgSz').get(W + 'w')) - int(sect.find(W + 'pgMar').get(W + 'left')) - int(sect.find(W + 'pgMar').get(W + 'right'))
def E(tag, **a):
    e = etree.Element(W + tag)
    for k, v in a.items(): e.set(W + k, str(v))
    return e
n = 0
for p in [e for e in body if e.tag == W + 'p' and 'Шинэ СҮД:' in txt(e)]:
    cols = [[p]]
    nx = p.getnext()
    while nx is not None and nx.tag == W + 'p' and (ITEM.match(txt(nx)) or LAB.match(txt(nx))):
        if LAB.match(txt(nx)): cols.append([])
        cols[-1].append(nx)
        nx = nx.getnext()
    anchor = p.getprevious()
    tbl = E('tbl'); pr = etree.SubElement(tbl, W + 'tblPr')
    pr.append(E('tblW', w=avail, type='dxa'))
    b = etree.SubElement(pr, W + 'tblBorders')
    for s in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'): b.append(E(s, val='nil'))
    pr.append(E('tblLayout', type='fixed'))
    m = etree.SubElement(pr, W + 'tblCellMar'); m.append(E('left', w=0, type='dxa')); m.append(E('right', w=113, type='dxa'))
    g = etree.SubElement(tbl, W + 'tblGrid'); ws = [avail // len(cols)] * len(cols); ws[-1] = avail - sum(ws[:-1])
    for w in ws: g.append(E('gridCol', w=w))
    tr = etree.SubElement(tbl, W + 'tr'); trpr = etree.SubElement(tr, W + 'trPr'); trpr.append(E('cantSplit'))
    for col, w in zip(cols, ws):
        tc = etree.SubElement(tr, W + 'tc'); tcpr = etree.SubElement(tc, W + 'tcPr'); tcpr.append(E('tcW', w=w, type='dxa'))
        for e in col:
            body.remove(e)
            for k in e.findall(f'{W}pPr/{W}keepNext'): k.getparent().remove(k)
            tc.append(e)
        if not col: tc.append(E('p'))
    anchor.addnext(tbl)
    sp = E('p'); spp = etree.SubElement(sp, W + 'pPr'); spp.append(E('keepNext')); spp.append(E('spacing', before=0, after=0, line=20, lineRule='exact'))
    tbl.addnext(sp)   # хоёр хүснэгтийг тусгаарлах маш нимгэн мөр
    n += 1
parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
with zipfile.ZipFile(f, 'w', zipfile.ZIP_DEFLATED) as o:
    for k, d in parts.items(): o.writestr(k, d)
print(n)
