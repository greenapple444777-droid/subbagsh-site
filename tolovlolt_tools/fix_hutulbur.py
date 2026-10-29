"""Сарын хөтөлбөрийн (I–IV долоо хоногийн үйл ажиллагаа) Word файлыг PROMPT.md-ийн загвараар засна:
A4, захын зай 1.5/1.5/0.5/0.5 см, Arial 11, долоо хоног бүр шинэ хуудаснаас,
гарчиг хуудасны доор ганцаараа үлдэхгүй (keepNext зөвхөн гарчигт), зураг асуултаасаа салахгүй.
Бичвэр өөрчлөгдөхгүй.

Хэрэглээ: python3 fix_hutulbur.py ЭХ.docx ГАРАХ.docx
"""
import os, re, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from lxml import etree
from split_weeks import W, TOP, LEFT, BOTTOM, RIGHT, apply_font, set_flag, txt

WEEK = re.compile(r'^\d+ (дугаар|дүгээр) сарын [IV]+ долоо хоног$')
HEAD = ('Шинэ үг', 'Үйл ажиллагаа', 'Хэрэглэгдэхүүн', 'Багшийн чиглүүлэх асуулт')


def main(src, out):
    z = zipfile.ZipFile(src)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = etree.fromstring(parts['word/document.xml'])
    body = doc.find(W + 'body')
    kids = [e for e in body if e.tag == W + 'p']
    for e in body:
        apply_font(e, True)
    for k in list(body.iter(W + 'keepNext')):
        k.getparent().remove(k)
    for i, p in enumerate(kids):
        t = txt(p)
        if WEEK.match(t):
            set_flag(p, 'pageBreakBefore')
            set_flag(p, 'keepNext')
            continue
        is_title = t and t == t.upper() and len(t) < 120 and re.search('[А-ЯӨҮЁ]', t)
        is_label = t.endswith(':') and len(t) < 60
        if is_title or is_label or t.startswith(HEAD):
            set_flag(p, 'keepNext')
        nxt = kids[i + 1] if i + 1 < len(kids) else None
        if t.startswith('-') and nxt is not None and txt(nxt).startswith('-'):
            set_flag(p, 'keepNext')                    # асуултууд хоорондоо салахгүй
        if nxt is not None and nxt.find('.//' + W + 'drawing') is not None and not txt(nxt):
            set_flag(p, 'keepNext')                    # зураг өмнөх мөрөөсөө салахгүй
    for sp in body.iter(W + 'sectPr'):
        sz = sp.find(W + 'pgSz')
        land = sz.get(W + 'orient') == 'landscape'
        sz.set(W + 'w', '16838' if land else '11906'); sz.set(W + 'h', '11906' if land else '16838')
        pm = sp.find(W + 'pgMar')
        for k, v in (('top', TOP), ('left', LEFT), ('bottom', BOTTOM), ('right', RIGHT),
                     ('header', TOP // 2), ('footer', BOTTOM)):
            pm.set(W + k, str(v))
    parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as o:
        for n, d in parts.items():
            o.writestr(n, d)
    print(out)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
