"""Хэрэглээ: python3 sud_inline_first.py 7ХОНОГ.docx
СҮД шошгын ("Шинэ СҮД:" г.м.) ард эхний СҮД-ийг нэг мөрөнд залгаж, 7 хоногийн хүснэгтийг нэг хуудсанд багтаахад туслана."""
import sys, zipfile, re, copy
sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.abspath(__file__)))
from lxml import etree
from split_weeks import W, txt
ITEM = re.compile(r'^\s*[А-ЯӨҮЁ]{2,4}\d+\.\d+[а-яөүё]{0,2}\.\s')
f = sys.argv[1]
z = zipfile.ZipFile(f); parts = {n: z.read(n) for n in z.namelist()}
doc = etree.fromstring(parts['word/document.xml']); body = doc.find(W + 'body')
n = 0
for p in list(body.iter(W + 'p')):
    if re.search(r'(Шинэ|Уялдаа|Давталт) СҮД:\s*$', txt(p)):
        q = p.getnext()
        if q is not None and q.tag == W + 'p' and ITEM.match(txt(q)):
            sp = etree.SubElement(p, W + 'r'); r0 = q.find(W + 'r')
            if r0 is not None and r0.find(W + 'rPr') is not None: sp.append(copy.deepcopy(r0.find(W + 'rPr')))
            t = etree.SubElement(sp, W + 't'); t.text = ' '; t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            for r in q.findall(W + 'r'): p.append(r)
            q.getparent().remove(q); n += 1
parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
with zipfile.ZipFile(f, 'w', zipfile.ZIP_DEFLATED) as o:
    for k, d in parts.items(): o.writestr(k, d)
print(n)
