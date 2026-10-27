"""Өдрийн төгсгөлийн зураг дараагийн хуудсанд ганцаараа шилжсэн бол тухайн өдрийн зургийг
бага зэрэг (алхам бүрт 8%, хамгийн багадаа 60%) жижигрүүлж, өдөр нэг хуудсандаа багтахаар болгоно.
LibreOffice-оор хуудаслалтыг шалгана.

Хэрэглээ: python3 fit_images.py ФАЙЛ.docx [ФАЙЛ2.docx ...]
"""
import math, os, re, subprocess, sys, tempfile, zipfile
from lxml import etree

W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
DAY_RE = re.compile(r'\d+ (дэх|дахь) өдөр\s*-\s*\S+ гараг')
STEP, MIN = 0.92, 0.60


def page_texts(docx, tmp):
    subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', tmp, docx],
                   check=True, capture_output=True)
    pdf = os.path.join(tmp, os.path.splitext(os.path.basename(docx))[0] + '.pdf')
    out = subprocess.run(['pdftotext', '-layout', pdf, '-'], check=True, capture_output=True, text=True).stdout
    return out.split('\f')


def scale(drawing, f):
    for ext in drawing.iter(WP + 'extent', A + 'ext'):
        for k in ('cx', 'cy'):
            if ext.get(k):
                ext.set(k, str(int(int(ext.get(k)) * f)))


def fit(path):
    z = zipfile.ZipFile(path)
    parts = {n: z.read(n) for n in z.namelist()}
    doc = etree.fromstring(parts['word/document.xml'])
    body = doc.find(W + 'body')
    done = {}
    for _ in range(int(math.log(MIN) / math.log(STEP)) + 1):
        with tempfile.TemporaryDirectory() as tmp:
            probe = os.path.join(tmp, 'probe.docx')
            parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
            with zipfile.ZipFile(probe, 'w', zipfile.ZIP_DEFLATED) as o:
                for n, d in parts.items():
                    o.writestr(n, d)
            pages = page_texts(probe, tmp)
        # хоосон (зөвхөн зурагтай) хуудасны өмнөх хуудасны өдрийг олно
        heads = []
        for i in range(1, len(pages)):
            if not pages[i].strip() and i < len(pages) - 1:   # сүүлийн элемент нь \f-ийн араас гарах хоосон мөр
                m = None
                for line in pages[i - 1].splitlines():
                    if DAY_RE.search(line):
                        m = line.strip(); break
                if m:
                    heads.append(m)
        if not heads:
            break
        changed = False
        kids = list(body)
        for h in heads:
            idx = next((j for j, e in enumerate(kids) if ''.join(e.itertext()).strip() == h), None)
            if idx is None:
                continue
            f_now = done.get(h, 1.0)
            if f_now * STEP < MIN:
                continue
            for e in kids[idx + 1:]:
                t = ''.join(e.itertext()).strip()
                if DAY_RE.search(t) or e.find('.//' + W + 'sectPr') is not None:
                    break
                for d in e.iter(W + 'drawing'):
                    scale(d, STEP)
            done[h] = f_now * STEP
            changed = True
        if not changed:
            break
    parts['word/document.xml'] = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as o:
        for n, d in parts.items():
            o.writestr(n, d)
    for h, f in done.items():
        print(f'  {os.path.basename(path)}: {h} — зураг {round(f * 100)}%')


if __name__ == '__main__':
    for p in sys.argv[1:]:
        fit(p)
