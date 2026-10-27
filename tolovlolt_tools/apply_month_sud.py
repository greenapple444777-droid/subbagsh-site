"""Сарын төлөвлөгөөний (.xlsx) 7 хоног бүрийн Шинэ / Уялдаа / Давталт СҮД-ийг
7 хоногийн Word төлөвлөлтөд (7 хоногийн хуудас + 5 өдөр) яг тааруулж сольно.

Хэрэглээ: python3 apply_month_sud.py САР.xlsx 7ХОНОГ.docx ДОЛОО_ХОНОГИЙН_ДУГААР [ГАРАХ.docx]
  ГАРАХ-ыг өгөөгүй бол файлыг өөрийг нь шинэчилнэ.

Сарын төлөвлөгөөний хүснэгт: B баганад «КОД: тайлбар», C–F баганад 1–4-р 7 хоногийн төрөл
(Шинэ / Уялдаа / Давталт).
"""
import copy, re, sys, zipfile
import openpyxl
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
XS = '{http://www.w3.org/XML/1998/namespace}space'
KINDS = ('Шинэ', 'Уялдаа', 'Давталт')
ITEM = re.compile(r'^\s*[А-ЯӨҮЁ]{2,4}\d+\.\d+[а-яөүёa-z]{0,2}\.\s')
LABEL = re.compile(r'^\s*(Шинэ|Уялдаа|Давталт) СҮД:\s*$')


def month_plan(path):
    ws = openpyxl.load_workbook(path, data_only=True).active
    head = next(r for r in range(1, ws.max_row + 1) if str(ws.cell(r, 2).value or '').startswith('Суралцахуйн үр дүн'))
    weeks = {w: {k: [] for k in KINDS} for w in range(1, 5)}
    for r in range(head + 1, ws.max_row + 1):
        b = str(ws.cell(r, 2).value or '').strip()
        m = re.match(r'^([А-ЯӨҮЁ]{2,4}\d+\.\d+[а-яөүёa-z]{0,2})\s*[:.]\s*(.+)$', b, re.S)
        if not m:
            continue
        code = m.group(1)[:-1] + m.group(1)[-1].replace('a', 'а')   # латин a → кирилл а
        for w in range(1, 5):
            kind = str(ws.cell(r, 2 + w).value or '').strip()
            if kind in KINDS:
                weeks[w][kind].append(f'{code}. {m.group(2).strip()}')
    return weeks


def text(p):
    return ''.join(p.itertext())


def set_text(p, value, bold_code=False):
    """Догол мөрийн run-уудыг нэг (эсвэл код + бичвэр) run-аар солино, форматыг хадгална."""
    runs = p.findall(W + 'r')
    base = None
    for r in runs:
        b = r.find(f'{W}rPr/{W}b')
        if b is None or b.get(W + 'val') in ('0', 'false'):
            base = r.find(W + 'rPr'); break
    if base is None and runs:
        base = runs[0].find(W + 'rPr')
    for r in runs:
        p.remove(r)

    def add(t, bold):
        r = etree.SubElement(p, W + 'r')
        rp = copy.deepcopy(base) if base is not None else etree.Element(W + 'rPr')
        for x in rp.findall(W + 'b') + rp.findall(W + 'bCs'):
            rp.remove(x)
        if bold:
            rp.insert(1 if rp.find(W + 'rFonts') is not None else 0, etree.Element(W + 'b'))
        r.append(rp)
        te = etree.SubElement(r, W + 't'); te.text = t; te.set(XS, 'preserve')

    if bold_code:
        code, rest = value.split(' ', 1)
        add(code + ' ', True); add(rest, False)
    else:
        add(value, False)


def apply(doc, plan):
    body = doc.find(W + 'body')
    blocks = 0
    for p in list(body.iter(W + 'p')):
        if 'Шинэ СҮД:' not in text(p):
            continue
        # хуучин жагсаалт (Шинэ-ийн мөрүүд, Уялдаа/Давталт шошго ба мөрүүд)-ыг цуглуулж хасна
        item_tpl = label_tpl = None
        nxt = p.getnext()
        while nxt is not None and nxt.tag == W + 'p' and (ITEM.match(text(nxt)) or LABEL.match(text(nxt))):
            if ITEM.match(text(nxt)) and item_tpl is None:
                item_tpl = copy.deepcopy(nxt)
            if LABEL.match(text(nxt)) and label_tpl is None:
                label_tpl = copy.deepcopy(nxt)
            cur, nxt = nxt, nxt.getnext()
            cur.getparent().remove(cur)
        if item_tpl is None:
            continue
        r0 = item_tpl.find(W + 'r')
        code_bold = r0 is not None and r0.find(f'{W}rPr/{W}b') is not None and \
            r0.find(f'{W}rPr/{W}b').get(W + 'val') not in ('0', 'false') and ITEM.match(text(r0) + ' x')
        if label_tpl is None:                      # шошгын загвар: Шинэ шошгоны тод run
            label_tpl = copy.deepcopy(item_tpl)
        prev = p
        for kind in KINDS:
            if not plan[kind]:
                continue
            if kind != 'Шинэ':
                lab = copy.deepcopy(label_tpl)
                set_text(lab, f'{kind} СҮД:')
                for r in lab.findall(W + 'r'):
                    rp = r.find(W + 'rPr')
                    if rp.find(W + 'b') is None:
                        rp.insert(1 if rp.find(W + 'rFonts') is not None else 0, etree.Element(W + 'b'))
                prev.addnext(lab); prev = lab
            for it in plan[kind]:
                q = copy.deepcopy(item_tpl)
                set_text(q, it, bool(code_bold))
                prev.addnext(q); prev = q
        blocks += 1
    return blocks


def main(xlsx, docx, week, out=None):
    plan = month_plan(xlsx)[int(week)]
    z = zipfile.ZipFile(docx)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = etree.fromstring(parts['word/document.xml'])
    n = apply(doc, plan)
    parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(out or docx, 'w', zipfile.ZIP_DEFLATED) as o:
        for k, d in parts.items():
            o.writestr(k, d)
    print(f'{out or docx}: {n} хэсэгт СҮД сольсон — ' +
          '; '.join(f'{k}: {", ".join(x.split(".")[0] + "." + x.split(".")[1] for x in plan[k])}' for k in KINDS if plan[k]))


if __name__ == '__main__':
    main(*sys.argv[1:])
