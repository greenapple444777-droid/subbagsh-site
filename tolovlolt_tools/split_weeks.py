"""Сарын төлөвлөлтийг (I–IV долоо хоног) долоо хоног бүрээр тусдаа .docx болгоно.

Хэрэглээ: python3 split_weeks.py ЭХ.docx ГАРАХ_ХАВТАС
Бичвэр, бүтцийг өөрчлөхгүй; зөвхөн хуудаслалт, захын зай, фонтыг засна.
"""
import copy, os, re, sys, zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
CM = lambda cm: round(cm * 567)
TOP, LEFT, BOTTOM, RIGHT = CM(1.5), CM(1.5), CM(0.5), CM(0.5)
FONT, SIZE = 'Arial', 22          # Arial 11pt (half-points)
DAY_RE = re.compile(r'\d+ (дэх|дахь) өдөр\s*-\s*\S+ гараг')


def txt(e):
    return ''.join(e.itertext()).strip()


def ppr_of(p):
    ppr = p.find(W + 'pPr')
    if ppr is None:
        ppr = etree.Element(W + 'pPr')
        p.insert(0, ppr)
    return ppr


def set_flag(p, name):
    """Insert an empty pPr child (keepNext / pageBreakBefore) in schema order."""
    order = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore']
    ppr = ppr_of(p)
    if ppr.find(W + name) is not None:
        return
    pos = 0
    for i, c in enumerate(ppr):
        if c.tag.replace(W, '') in order[:order.index(name)]:
            pos = i + 1
    ppr.insert(pos, etree.Element(W + name))


def set_font(rpr, resize):
    rf = rpr.find(W + 'rFonts')
    if rf is None:
        rf = etree.Element(W + 'rFonts')
        rpr.insert(1 if rpr.find(W + 'rStyle') is not None else 0, rf)
    for a in list(rf.attrib):
        if a.endswith('Theme'):
            del rf.attrib[a]
    for a in ('ascii', 'hAnsi', 'cs', 'eastAsia'):
        rf.set(W + a, FONT)
    if not resize:
        return
    for tag in ('sz', 'szCs'):
        el = rpr.find(W + tag)
        if el is not None and int(el.get(W + 'val')) > SIZE:
            continue                      # гарчгийн том үсгийг хэвээр үлдээнэ
        if el is None:
            el = etree.SubElement(rpr, W + tag)
        el.set(W + 'val', str(SIZE))
    # schema order: sz, szCs must come after most props and before highlight/u/... — re-sort
    ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
             'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
             'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
             'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']
    kids = list(rpr)
    kids.sort(key=lambda c: ORDER.index(c.tag.replace(W, '')) if c.tag.replace(W, '') in ORDER else len(ORDER))
    for c in kids:
        rpr.append(c)


def apply_font(el, resize):
    for r in el.iter(W + 'r'):
        rpr = r.find(W + 'rPr')
        if rpr is None:
            rpr = etree.Element(W + 'rPr')
            r.insert(0, rpr)
        set_font(rpr, resize)
    for ppr in el.iter(W + 'pPr'):
        rpr = ppr.find(W + 'rPr')
        if rpr is not None:
            set_font(rpr, resize)


def main(src, outdir):
    z = zipfile.ZipFile(src)
    names = z.namelist()
    doc = etree.fromstring(z.read('word/document.xml'))
    rels = etree.fromstring(z.read('word/_rels/document.xml.rels'))
    styles = etree.fromstring(z.read('word/styles.xml'))
    body = doc.find(W + 'body')
    els = list(body)
    final_sect = els[-1]
    assert final_sect.tag == W + 'sectPr'

    weeks = [i for i, e in enumerate(els) if txt(e).startswith('БАТЛАВ')]
    assert len(weeks) == 4, weeks
    outro = [i for i, e in enumerate(els) if txt(e).lower().startswith('сонсъё, дэмжье')]
    outro_start = outro[0] if outro else len(els) - 1
    bounds = weeks + [outro_start]
    # файлын нэр: «АХЛАХ БҮЛЭГ - ХӨГЖИЛТЭЙ ТОГЛООМУУД»
    title = next((txt(e) for e in els[weeks[0]:weeks[0] + 4] if 'БҮЛГИЙН' in txt(e)), '')
    m = re.match(r'(\S+) БҮЛГИЙН', title)
    group = (f'{m.group(1)} БҮЛЭГ - ' if m else '') + txt(els[1])
    ref_grid = None

    # styles: default font Arial 11
    rpd = styles.find(f'{W}docDefaults/{W}rPrDefault/{W}rPr')
    if rpd is not None:
        set_font(rpd, True)

    for k in range(4):
        intro = [copy.deepcopy(e) for e in els[:weeks[0]]]
        week = [copy.deepcopy(e) for e in els[bounds[k]:bounds[k + 1]]]
        tail = [copy.deepcopy(e) for e in els[outro_start:-1]]
        # outro (Сонсъё, дэмжье ...) өдрийн хэсгийн төгсгөлд, тухайн хэсгийн section дотор орно
        if week and week[-1].find('.//' + W + 'sectPr') is not None:
            week = week[:-1] + tail + [week[-1]]
        else:
            week = week + tail
        for e in list(body):
            body.remove(e)
        for e in intro + week + [copy.deepcopy(final_sect)]:
            body.append(e)

        # 1) захын зай: дээр 1.5, зүүн 1.5, доор 0.5, баруун 0.5 см
        for sp in body.iter(W + 'sectPr'):
            pm = sp.find(W + 'pgMar')
            pm.set(W + 'top', str(TOP)); pm.set(W + 'left', str(LEFT))
            pm.set(W + 'bottom', str(BOTTOM)); pm.set(W + 'right', str(RIGHT))
            pm.set(W + 'header', str(TOP // 2)); pm.set(W + 'footer', str(BOTTOM))

        # 2) 7 хоногийн хүснэгт: толгой мөр дараагийн хуудсанд давтагдахгүй, мөр хуваагдахгүй, өргөн хуудсанд багтана
        kids = list(body)
        weekly_zone = set()
        for tbl in body.iter(W + 'tbl'):
            top = tbl
            while top.getparent() is not body:
                top = top.getparent()
            ti = kids.index(top)
            # хүснэгтийн хуудасны элементүүд (БАТЛАВ ... гарын үсэг) — үсгийн хэмжээг хэвээр үлдээнэ
            j = ti
            while j > 0 and not txt(kids[j]).startswith('БАТЛАВ'):
                j -= 1
            n = ti
            while kids[n].find('.//' + W + 'sectPr') is None:
                n += 1
            weekly_zone.update(range(j, n + 1))

            for h in list(tbl.iter(W + 'tblHeader')):
                h.getparent().remove(h)
            for tr in tbl.findall(W + 'tr'):
                trpr = tr.find(W + 'trPr')
                if trpr is None:
                    trpr = etree.Element(W + 'trPr'); tr.insert(1 if tr.find(W + 'tblPrEx') is not None else 0, trpr)
                if trpr.find(W + 'cantSplit') is None:
                    trpr.insert(0, etree.Element(W + 'cantSplit'))
            sect = kids[n].find('.//' + W + 'sectPr')
            avail = int(sect.find(W + 'pgSz').get(W + 'w')) - LEFT - RIGHT
            grid = tbl.find(W + 'tblGrid')
            gw = [int(g.get(W + 'w')) for g in grid]
            if ref_grid is None and len(set(gw)) > 1:
                ref_grid = gw
            if len(set(gw)) == 1 and ref_grid and len(ref_grid) == len(gw):
                gw = ref_grid[:]
            f = min(1, (avail - gw[0]) / (sum(gw) - gw[0]))
            nw = [gw[0]] + [int(x * f) for x in gw[1:]]
            for g, x in zip(grid, nw):
                g.set(W + 'w', str(x))
            for tr in tbl.findall(W + 'tr'):
                col = 0
                for tc in tr.findall(W + 'tc'):
                    pr = tc.find(W + 'tcPr'); gs = pr.find(W + 'gridSpan')
                    span = int(gs.get(W + 'val')) if gs is not None else 1
                    tw = pr.find(W + 'tcW')
                    if tw is not None and tw.get(W + 'type') == 'dxa':
                        tw.set(W + 'w', str(sum(nw[col:col + span])))
                    col += span
            for p in tbl.findall(W + 'tr')[-1].iter(W + 'p'):
                set_flag(p, 'keepNext')        # сүүлийн мөр гарын үсгийн мөртэйгөө хамт

        # 3) Фонт: бүх бичвэр Arial; 7 хоногийн хүснэгтийн хуудаснаас бусад нь 11pt
        for i, e in enumerate(kids):
            apply_font(e, i not in weekly_zone)

        # 4) Өдөр бүр шинэ хуудаснаас; гарчиг хуудасны доор ганцаараа үлдэхгүй
        in_day = in_head = False
        for i, e in enumerate(kids):
            t = txt(e)
            if e.tag != W + 'p':
                continue
            if e.find('.//' + W + 'sectPr') is not None:
                in_day = False
                continue
            is_day = bool(DAY_RE.search(t)) and len(t) < 200
            if is_day:
                prev = next((x for x in reversed(kids[:i]) if txt(x) or x.find('.//' + W + 'drawing') is not None), None)
                if prev is not None and not txt(prev).endswith('долоо хоног'):
                    set_flag(e, 'pageBreakBefore')
                in_day = in_head = True
            if not in_day:
                continue
            if not t and e.find('.//' + W + 'br[@' + W + 'type="page"]') is not None \
                    and e.find('.//' + W + 'drawing') is None:
                body.remove(e)                 # хуучин хоосон page break — давхар хуудас үүсгэхгүй
                continue
            if t.startswith('Үйл ажиллагаа'):
                in_head = False
            nxt = next((x for x in kids[i + 1:] if x.getparent() is body), None)
            nxt_t = txt(nxt) if nxt is not None else ''
            heading = t.startswith(('Үйл ажиллагаа', 'Хэрэглэгдэхүүн', 'Багшийн чиглүүлэх асуулт')) \
                or t.lower().startswith(('сонсъё, дэмжье', 'эцэг эхтэйгээ', 'эргэцүүлэл'))
            if (in_head or heading) and not (nxt is not None and DAY_RE.search(nxt_t)):
                set_flag(e, 'keepNext')

        data = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
        used = set(re.findall(rb'r:(?:embed|id|link)="([^"]+)"', data))
        nrels = copy.deepcopy(rels)
        drop = set()
        for r in list(nrels):
            if r.get('Type').endswith('/image') and r.get('Id').encode() not in used:
                drop.add('word/' + r.get('Target')); nrels.remove(r)
        fn = os.path.join(outdir, f'{k + 1}. {group} - XI сарын {"I II III IV".split()[k]} долоо хоног.docx')
        with zipfile.ZipFile(fn, 'w', zipfile.ZIP_DEFLATED) as o:
            for n_ in names:
                if n_ in drop:
                    continue
                if n_ == 'word/document.xml':
                    o.writestr(n_, data)
                elif n_ == 'word/_rels/document.xml.rels':
                    o.writestr(n_, etree.tostring(nrels, xml_declaration=True, encoding='UTF-8', standalone=True))
                elif n_ == 'word/styles.xml':
                    o.writestr(n_, etree.tostring(styles, xml_declaration=True, encoding='UTF-8', standalone=True))
                else:
                    o.writestr(n_, z.read(n_))
        print(fn)


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
