/*
 * Excel (.xlsx) → Word (.docx) хөрвүүлэгч — хөтөч болон Node.js дээр ажиллана.
 * Хэрэглээ: const { blob/bytes, log } = await excelToWord(arrayBuffer, { fit: true });
 * Хамаарал: ExcelJS 4.4.0, JSZip 3.10.1 (хөтөч дээр глобал, Node дээр require).
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) {
    module.exports = factory(require("exceljs"), require("jszip"));
  } else {
    root.excelToWord = factory(root.ExcelJS, root.JSZip);
  }
})(typeof self !== "undefined" ? self : this, function (ExcelJS, JSZip) {
  "use strict";

  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

  // Excel paperSize → [өргөн, өндөр] twip-ээр
  const PAPER = { 1: [12240, 15840], 5: [12240, 20160], 8: [16838, 23811],
                  9: [11906, 16838], 11: [8391, 11906], 13: [10318, 14570] };

  const BORDER_MAP = {
    hair: ["single", 2], thin: ["single", 4], medium: ["single", 12], thick: ["single", 18],
    double: ["double", 4], dotted: ["dotted", 4], dashed: ["dashed", 4],
    mediumDashed: ["dashed", 12], dashDot: ["dotDash", 4], mediumDashDot: ["dotDash", 12],
    dashDotDot: ["dotDotDash", 4], mediumDashDotDot: ["dotDotDash", 12], slantDashDot: ["dotDash", 12],
  };
  const BORDER_WEIGHT = { hair: 1, dotted: 2, dashDotDot: 3, dashDot: 4, dashed: 5, thin: 6,
    mediumDashDotDot: 7, slantDashDot: 8, mediumDashDot: 9, mediumDashed: 10, medium: 11,
    thick: 12, double: 13 };
  const H_ALIGN = { left: "left", center: "center", centerContinuous: "center", right: "right",
    justify: "both", distributed: "distribute", fill: "left" };
  const V_ALIGN = { top: "top", middle: "center", bottom: "bottom", justify: "center",
    distributed: "center" };

  const INDEXED = ["000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF","000000","FFFFFF","FF0000","00FF00","0000FF","FFFF00","FF00FF","00FFFF","800000","008000","000080","808000","800080","008080","C0C0C0","808080","9999FF","993366","FFFFCC","CCFFFF","660066","FF8080","0066CC","CCCCFF","000080","FF00FF","FFFF00","00FFFF","800080","800000","008080","0000FF","00CCFF","CCFFFF","CCFFCC","FFFF99","99CCFF","FF99CC","CC99FF","FFCC99","3366FF","33CCCC","99CC00","FFCC00","FF9900","FF6600","666699","969696","003366","339966","003300","333300","993300","993366","333399","333333"];
  const THEME_DEFAULT = ["FFFFFF","000000","E7E6E6","44546A","4472C4","ED7D31","A5A5A5","FFC000","5B9BD5","70AD47","0563C1","954F72"];
  const THEME_ORDER = ["lt1","dk1","lt2","dk2","accent1","accent2","accent3","accent4","accent5","accent6","hlink","folHlink"];

  const esc = (s) => String(s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

  // ------------------------------------------------------------ өнгө ---
  function parseTheme(xml) {
    if (!xml) return THEME_DEFAULT.slice();
    const out = {};
    for (const key of THEME_ORDER) {
      const m = new RegExp("<a:" + key + ">([\\s\\S]*?)</a:" + key + ">").exec(xml);
      if (!m) continue;
      const v = /srgbClr val="([0-9A-Fa-f]{6})"/.exec(m[1]) || /lastClr="([0-9A-Fa-f]{6})"/.exec(m[1]);
      if (v) out[key] = v[1];
    }
    return THEME_ORDER.map((k, i) => out[k] || THEME_DEFAULT[i]);
  }

  function rgbToHls(r, g, b) {
    const max = Math.max(r, g, b), min = Math.min(r, g, b), l = (max + min) / 2;
    if (max === min) return [0, l, 0];
    const d = max - min, s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    let h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [h / 6, l, s];
  }
  function hlsToRgb(h, l, s) {
    if (s === 0) return [l, l, l];
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const f = (t) => { t = (t + 1) % 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
    return [f(h + 1 / 3), f(h), f(h - 1 / 3)];
  }
  function applyTint(hex, tint) {
    if (!tint) return hex;
    let [h, l, s] = rgbToHls(...[0, 2, 4].map((i) => parseInt(hex.substr(i, 2), 16) / 255));
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    return hlsToRgb(h, Math.max(0, Math.min(1, l)), s)
      .map((v) => Math.round(v * 255).toString(16).padStart(2, "0")).join("").toUpperCase();
  }
  function makeColorResolver(theme) {
    return function (c) {
      if (!c) return null;
      let hex = null;
      if (typeof c.argb === "string") hex = c.argb.slice(-6).toUpperCase();
      else if (typeof c.theme === "number") hex = theme[c.theme] || null;
      else if (typeof c.indexed === "number" && c.indexed < 64) hex = INDEXED[c.indexed];
      return hex ? applyTint(hex, c.tint) : null;
    };
  }

  // ---------------------------------------------------- тооны формат ---
  const pad = (n, w = 2) => String(n).padStart(w, "0");
  const MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const DAYS = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];

  function fmtDate(d, fmt) {
    let f = (fmt || "").replace(/\[[^\]]*\]/g, "").replace(/"([^"]*)"/g, "$1").replace(/\\/g, "");
    if (!/[ymdhs]/i.test(f)) f = "yyyy-mm-dd";
    const tokens = f.match(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM|am\/pm|[\s\S]/gi) || [];
    const ampm = /am\/pm/i.test(f);
    let prevH = false, out = "";
    tokens.forEach((t, i) => {
      const tl = t.toLowerCase(), next = tokens.slice(i + 1, i + 3).join("").toLowerCase();
      if (tl === "yyyy") out += d.getUTCFullYear();
      else if (tl === "yy") out += pad(d.getUTCFullYear() % 100);
      else if (tl === "mm" || tl === "m") {
        const isMin = prevH || next.startsWith(":s") || next.startsWith("s");
        const v = isMin ? d.getUTCMinutes() : d.getUTCMonth() + 1;
        out += tl === "mm" ? pad(v) : v;
      } else if (tl === "mmm") out += MONTHS[d.getUTCMonth()].slice(0, 3);
      else if (tl === "mmmm") out += MONTHS[d.getUTCMonth()];
      else if (tl === "dd") out += pad(d.getUTCDate());
      else if (tl === "d") out += d.getUTCDate();
      else if (tl === "ddd") out += DAYS[d.getUTCDay()].slice(0, 3);
      else if (tl === "dddd") out += DAYS[d.getUTCDay()];
      else if (tl === "hh" || tl === "h") {
        let h = d.getUTCHours(); if (ampm) h = h % 12 || 12;
        out += tl === "hh" ? pad(h) : h;
      } else if (tl === "ss" || tl === "s") out += tl === "ss" ? pad(d.getUTCSeconds()) : d.getUTCSeconds();
      else if (tl === "am/pm") out += d.getUTCHours() < 12 ? "AM" : "PM";
      else out += t;
      if (tl === "hh" || tl === "h") prevH = true; else if (tl !== ":") prevH = false;
    });
    return out;
  }

  function fmtGeneral(v) {
    if (Number.isInteger(v)) return String(v);
    return String(Number(v.toPrecision(10)));
  }

  function withThousands(s) {
    const [i, f] = s.split(".");
    const neg = i.startsWith("-");
    const body = (neg ? i.slice(1) : i).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
    return (neg ? "-" : "") + body + (f !== undefined ? "." + f : "");
  }

  function formatValue(v, fmt) {
    if (v === null || v === undefined) return "";
    if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
    if (v instanceof Date) return fmtDate(v, fmt);
    if (typeof v !== "number") return String(v);
    if (!fmt || fmt === "General" || fmt === "@") return fmtGeneral(v);

    const sections = fmt.split(";");
    let sec = sections[0], negSec = false;
    if (v < 0 && sections.length > 1) { sec = sections[1]; negSec = true; }
    else if (v === 0 && sections.length > 2) sec = sections[2];
    if (!sec.trim()) return "";
    sec = sec.replace(/\[[^\]]*\]/g, "").replace(/_./g, " ").replace(/\*/g, "");
    const m = /[#0?][#0?,]*(\.[#0?]+)?(E[+-]0+)?/i.exec(sec);
    const clean = (s) => s.replace(/"([^"]*)"/g, "$1").replace(/\\/g, "");
    if (!m) return clean(sec) || fmtGeneral(v);
    const numFmt = m[0];
    const decimals = m[1] ? m[1].length - 1 : 0;
    const minDec = m[1] ? (m[1].match(/0/g) || []).length : 0;
    let val = negSec ? Math.abs(v) : v;
    if (sec.includes("%")) val *= 100;
    let body;
    if (m[2]) {
      body = val.toExponential(decimals).toUpperCase().replace(/E([+-])(\d)$/, "E$10$2");
    } else {
      body = val.toFixed(decimals);
      if (decimals > minDec && body.includes(".")) {
        let [ip, fp] = body.split(".");
        fp = fp.replace(/0+$/, "");
        fp = fp + "0".repeat(Math.max(0, minDec - fp.length));
        body = ip + (fp ? "." + fp : "");
      }
      if (numFmt.includes(",")) body = withThousands(body);
      const intZeros = (numFmt.split(".")[0].match(/0/g) || []).length;
      if (intZeros === 0 && /^-?0\./.test(body)) body = body.replace("0.", ".");
    }
    return clean(sec.slice(0, m.index)) + body + clean(sec.slice(m.index + numFmt.length));
  }

  // ------------------------------------------------------ туслахууд ---
  function colToNum(s) { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + ch.charCodeAt(0) - 64; return n; }
  function parseRef(ref) {
    const m = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(ref.trim());
    if (!m) return null;
    const c1 = colToNum(m[1]), r1 = +m[2], c2 = m[3] ? colToNum(m[3]) : c1, r2 = m[4] ? +m[4] : r1;
    return { minC: Math.min(c1, c2), maxC: Math.max(c1, c2), minR: Math.min(r1, r2), maxR: Math.max(r1, r2) };
  }
  const colWidthToTwips = (w) => (w > 0 ? Math.floor(w * 7 + 5) * 0.75 * 20 : 0);

  function hasVisual(cell) {
    const s = cell && cell.style;
    if (!s) return false;
    const b = s.border;
    if (b && ["left", "right", "top", "bottom"].some((k) => b[k] && b[k].style)) return true;
    return !!(s.fill && s.fill.type === "pattern" && s.fill.pattern && s.fill.pattern !== "none");
  }

  // ------------------------------------------------- хуудас хөрвүүлэх ---
  function convertSheet(ws, color, opts, stats) {
    const ps = ws.pageSetup || {};
    const props = ws.properties || {};
    const cellAt = (r, c) => (r >= 1 && c >= 1 ? ws.findCell(r, c) : undefined);

    // Хуудасны тохиргоо
    let [pw, ph] = PAPER[ps.paperSize] || PAPER[9];
    const landscape = ps.orientation === "landscape";
    if (landscape) [pw, ph] = [ph, pw];
    const mg = ps.margins || {};
    const tw = (v, d) => Math.round((v != null ? v : d) * 1440);
    const margins = { left: tw(mg.left, 0.7), right: tw(mg.right, 0.7), top: tw(mg.top, 0.75),
      bottom: tw(mg.bottom, 0.75), header: tw(mg.header, 0.3), footer: tw(mg.footer, 0.3) };
    const usableW = pw - margins.left - margins.right;
    const usableH = ph - margins.top - margins.bottom;

    // Муж
    let range = null;
    if (ps.printArea) range = parseRef(String(ps.printArea).split(/&&|,/)[0].split("!").pop());
    if (!range) {
      let maxR = 0, maxC = 0;
      ws.eachRow({ includeEmpty: false }, (row, r) => {
        row.eachCell({ includeEmpty: false }, (cell, c) => {
          if ((cell.value !== null && cell.value !== undefined && cell.value !== "") || hasVisual(cell)) {
            maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
          }
        });
      });
      range = { minR: 1, minC: 1, maxR, maxC };
    }
    const merges = (ws.model.merges || []).map(parseRef).filter(Boolean);
    if (!ps.printArea) {
      for (const mr of merges) {
        if (mr.minR <= range.maxR && mr.minC <= range.maxC) {
          range.maxR = Math.max(range.maxR, mr.maxR); range.maxC = Math.max(range.maxC, mr.maxC);
        }
      }
    }
    if (range.maxR < 1 || range.maxC < 1) return null;

    const defaultW = props.defaultColWidth || 8.43;
    const defaultH = props.defaultRowHeight || 15;
    const cols = [], colTw = {};
    for (let c = range.minC; c <= range.maxC; c++) {
      const col = ws.getColumn(c);
      const w = col.width != null ? col.width : defaultW;
      if (col.hidden || w <= 0) continue;
      cols.push(c); colTw[c] = colWidthToTwips(w);
    }
    const rows = [], rowH = {};
    for (let r = range.minR; r <= range.maxR; r++) {
      const row = ws.findRow(r);
      if (row && row.hidden) continue;
      rows.push(r); rowH[r] = (row && row.height) || defaultH;
    }
    if (!cols.length || !rows.length) return null;

    // Масштаб
    const totalW = cols.reduce((s, c) => s + colTw[c], 0);
    let scale = 1;
    if (!ps.fitToPage && ps.scale && ps.scale >= 10 && ps.scale < 100) scale = ps.scale / 100;
    if (opts.fit && totalW * scale > usableW) scale = usableW / totalW;
    if (opts.fit && ps.fitToPage) {
      const pagesH = ps.fitToHeight != null ? ps.fitToHeight : 1;
      if (pagesH) {
        const totalH = rows.reduce((s, r) => s + rowH[r], 0) * 20 * 1.03;
        scale = Math.min(scale, (usableH * pagesH) / totalH);
      }
    }
    for (const c of cols) colTw[c] = Math.max(20, Math.floor(colTw[c] * scale));
    const tableW = cols.reduce((s, c) => s + colTw[c], 0);
    const fontScale = opts.scaleFonts ? scale : 1;

    const colSet = new Set(cols), rowSet = new Set(rows);
    const mergeAt = new Map();
    const key = (r, c) => r * 100000 + c;
    for (const mr of merges) {
      const vc = [], vr = [];
      for (let c = mr.minC; c <= mr.maxC; c++) if (colSet.has(c)) vc.push(c);
      for (let r = mr.minR; r <= mr.maxR; r++) if (rowSet.has(r)) vr.push(r);
      if (!vc.length || !vr.length) continue;
      for (const r of vr) {
        for (const c of vc) mergeAt.set(key(r, c), { skip: true });
        mergeAt.set(key(r, vc[0]), {
          start: true, cols: vc, mr, top: r === vr[0],
          vmerge: vr.length > 1 ? (r === vr[0] ? "restart" : "continue") : null,
        });
      }
    }

    // Хүрээ
    const side = (r, c, s) => {
      const cell = cellAt(r, c);
      const b = cell && cell.style && cell.style.border && cell.style.border[s];
      return b && b.style ? b : null;
    };
    const OPP = { left: ["right", 0, -1], right: ["left", 0, 1], top: ["bottom", -1, 0], bottom: ["top", 1, 0] };
    const edge = (r, c, s) => {
      const [os, dr, dc] = OPP[s];
      const cands = [side(r, c, s), side(r + dr, c + dc, os)].filter(Boolean);
      if (!cands.length) return null;
      return cands.reduce((a, b) => ((BORDER_WEIGHT[b.style] || 0) > (BORDER_WEIGHT[a.style] || 0) ? b : a));
    };
    const borderXml = (tag, b) => {
      if (!b) return `<w:${tag} w:val="nil"/>`;
      const [val, sz] = BORDER_MAP[b.style] || ["single", 4];
      return `<w:${tag} w:val="${val}" w:sz="${sz}" w:space="0" w:color="${color(b.color) || "000000"}"/>`;
    };

    // Текст
    const rPr = (font, over) => {
      const pick = (k) => (over && over[k] != null ? over[k] : font ? font[k] : undefined);
      const name = esc(pick("name") || "Calibri");
      let x = `<w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:cs="${name}" w:eastAsia="${name}"/>`;
      if (pick("bold")) x += "<w:b/><w:bCs/>";
      if (pick("italic")) x += "<w:i/><w:iCs/>";
      if (pick("strike")) x += "<w:strike/>";
      const col = color(pick("color"));
      if (col) x += `<w:color w:val="${col}"/>`;
      const size = (Number(pick("size")) || 11) * fontScale;
      const half = Math.max(2, Math.round(size * 2));
      x += `<w:sz w:val="${half}"/><w:szCs w:val="${half}"/>`;
      const u = pick("underline");
      if (u && u !== "none") x += `<w:u w:val="${String(u).includes("double") ? "double" : "single"}"/>`;
      const va = pick("vertAlign");
      if (va === "superscript" || va === "subscript") x += `<w:vertAlign w:val="${va}"/>`;
      return `<w:rPr>${x}</w:rPr>`;
    };
    const runs = (text, rp) => String(text).replace(/\r\n?/g, "\n").split("\n").map((line, i) =>
      (i ? `<w:r>${rp}<w:br/></w:r>` : "") +
      (line ? `<w:r>${rp}<w:t xml:space="preserve">${esc(line)}</w:t></w:r>` : "")).join("");

    const paragraph = (cell) => {
      const style = (cell && cell.style) || {};
      const font = style.font;
      const al = style.alignment || {};
      const baseRpr = rPr(font);
      let value = cell ? cell.value : null;
      let body = "";
      if (value && typeof value === "object" && !(value instanceof Date)) {
        if ("formula" in value || "sharedFormula" in value) {
          if (value.result === undefined) { stats.missingFormula++; value = null; }
          else value = value.result;
        }
        if (value && typeof value === "object" && "hyperlink" in value) value = value.text;
        if (value && typeof value === "object" && "error" in value) value = value.error;
      }
      if (value && typeof value === "object" && Array.isArray(value.richText)) {
        body = value.richText.map((p) => runs(p.text || "", rPr(font, p.font))).join("");
      } else if (value !== null && value !== undefined) {
        body = runs(formatValue(value, style.numFmt), baseRpr);
      }
      const isNum = typeof value === "number";
      let jc = H_ALIGN[al.horizontal];
      if (!jc) jc = isNum || value instanceof Date ? "right" : typeof value === "boolean" ? "center" : "left";
      let ind = "";
      if (al.indent) ind = `<w:ind w:${jc === "right" ? "right" : "left"}="${Math.round(al.indent * 180 * fontScale)}"/>`;
      return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>${ind}` +
        `<w:jc w:val="${jc}"/>${baseRpr}</w:pPr>${body}</w:p>`;
    };

    // Хэвлэх гарчиг мөр
    let titleRows = new Set();
    const tm = /\$?(\d+):\$?(\d+)/.exec(ps.printTitlesRow || "");
    if (tm) for (let r = +tm[1]; r <= +tm[2]; r++) titleRows.add(r);

    const jcTbl = ps.horizontalCentered ? "center" : "left";
    const mar = Math.round(30 * Math.max(scale, 0.5));
    const emit = (segRows) => {
      let headerOk = true;
      const x = [`<w:tbl><w:tblPr><w:tblW w:w="${tableW}" w:type="dxa"/><w:jc w:val="${jcTbl}"/>` +
        `<w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar>` +
        `<w:top w:w="0" w:type="dxa"/><w:left w:w="${mar}" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/>` +
        `<w:right w:w="${mar}" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr><w:tblGrid>` +
        cols.map((c) => `<w:gridCol w:w="${colTw[c]}"/>`).join("") + "</w:tblGrid>"];
      for (const r of segRows) {
        let trPr = `<w:cantSplit/><w:trHeight w:val="${Math.max(1, Math.floor(rowH[r] * 20 * scale))}" w:hRule="atLeast"/>`;
        if (titleRows.has(r) && headerOk) trPr += "<w:tblHeader/>"; else headerOk = false;
        x.push(`<w:tr><w:trPr>${trPr}</w:trPr>`);
        for (const c of cols) {
          const info = mergeAt.get(key(r, c));
          if (info && info.skip) continue;
          let spanCols = [c], styleCell, bT, bB, bL, bR, contentCell;
          if (info && info.start) {
            const mr = info.mr;
            spanCols = info.cols;
            styleCell = cellAt(mr.minR, mr.minC);
            bT = null; bB = null; bL = null; bR = null;
            for (const cc of spanCols) bT = bT || edge(r, cc, "top");
            const bottomRow = info.vmerge === "restart" ? mr.maxR : r;
            for (const cc of spanCols) bB = bB || edge(bottomRow, cc, "bottom");
            for (let rr = mr.minR; rr <= mr.maxR; rr++) {
              bL = bL || edge(rr, mr.minC, "left");
              bR = bR || edge(rr, mr.maxC, "right");
            }
            contentCell = info.top ? styleCell : null;
          } else {
            styleCell = cellAt(r, c);
            bT = edge(r, c, "top"); bB = edge(r, c, "bottom"); bL = edge(r, c, "left"); bR = edge(r, c, "right");
            contentCell = styleCell;
          }
          const width = spanCols.reduce((s, cc) => s + colTw[cc], 0);
          let tcPr = `<w:tcW w:w="${width}" w:type="dxa"/>`;
          if (spanCols.length > 1) tcPr += `<w:gridSpan w:val="${spanCols.length}"/>`;
          if (info && info.vmerge) tcPr += info.vmerge === "restart" ? `<w:vMerge w:val="restart"/>` : "<w:vMerge/>";
          tcPr += "<w:tcBorders>" + borderXml("top", bT) + borderXml("left", bL) +
            borderXml("bottom", bB) + borderXml("right", bR) + "</w:tcBorders>";
          const st = (styleCell && styleCell.style) || {};
          const fill = st.fill;
          if (fill) {
            let fc = null;
            if (fill.type === "pattern" && fill.pattern && fill.pattern !== "none") fc = color(fill.fgColor) || color(fill.bgColor);
            else if (fill.type === "gradient" && fill.stops && fill.stops.length) fc = color(fill.stops[0].color);
            if (fc) tcPr += `<w:shd w:val="clear" w:color="auto" w:fill="${fc}"/>`;
          }
          const al = st.alignment || {};
          const rot = al.textRotation;
          if (rot === "vertical" || (typeof rot === "number" && rot < 0)) tcPr += `<w:textDirection w:val="tbRl"/>`;
          else if (typeof rot === "number" && rot >= 45) tcPr += `<w:textDirection w:val="btLr"/>`;
          if (al.shrinkToFit) tcPr += "<w:tcFitText/>";
          tcPr += `<w:vAlign w:val="${V_ALIGN[al.vertical] || "bottom"}"/>`;
          const para = contentCell ? paragraph(contentCell)
            : paragraph(styleCell ? { style: styleCell.style, value: null } : null);
          x.push(`<w:tc><w:tcPr>${tcPr}</w:tcPr>${para}</w:tc>`);
        }
        x.push("</w:tr>");
      }
      x.push("</w:tbl>");
      return x.join("");
    };

    // Давтагдах гарчиг хүснэгтийн эхэнд байх ёстой → өмнөх мөрүүдийг салгана
    let segments = [rows];
    if (titleRows.size) {
      const first = Math.min(...titleRows);
      const pre = rows.filter((r) => r < first);
      const crosses = merges.some((mr) => mr.minR < first && first <= mr.maxR);
      if (pre.length && pre.length < rows.length && !crosses) segments = [pre, rows.filter((r) => r >= first)];
    }
    const sectPr = `<w:sectPr><w:pgSz w:w="${pw}" w:h="${ph}"${landscape ? ' w:orient="landscape"' : ""}/>` +
      `<w:pgMar w:top="${margins.top}" w:right="${margins.right}" w:bottom="${margins.bottom}" ` +
      `w:left="${margins.left}" w:header="${margins.header}" w:footer="${margins.footer}" w:gutter="0"/>` +
      `<w:cols w:space="720"/></w:sectPr>`;
    return { tables: segments.map(emit), sectPr, scale };
  }

  // ------------------------------------------------------- docx багц ---
  const TAIL_PPR = `<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>`;

  async function buildDocx(sheets) {
    let body = "";
    sheets.forEach((s, i) => {
      const last = i === sheets.length - 1;
      s.tables.forEach((t, j) => {
        body += t;
        const breakHere = !last && j === s.tables.length - 1;
        body += `<w:p><w:pPr>${TAIL_PPR}${breakHere ? s.sectPr : ""}</w:pPr></w:p>`;
      });
      if (last) body += s.sectPr;
    });
    const zip = new JSZip();
    zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`);
    zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`);
    zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`);
    zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W_NS}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri" w:eastAsia="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>`);
    zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}"><w:body>${body}</w:body></w:document>`);
    return zip;
  }

  // -------------------------------------------------------- үндсэн ---
  async function excelToWord(data, opts) {
    opts = Object.assign({ fit: true, scaleFonts: true, sheets: null }, opts || {});
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data);
    const theme = parseTheme(wb._themes && wb._themes.theme1);
    const color = makeColorResolver(theme);
    const log = [];
    const stats = { missingFormula: 0 };
    const out = [];
    wb.eachSheet((ws) => {
      if (opts.sheets && !opts.sheets.includes(ws.name)) return;
      if (!opts.sheets && ws.state && ws.state !== "visible") return;
      const res = convertSheet(ws, color, opts, stats);
      if (!res) { log.push({ sheet: ws.name, skipped: true }); return; }
      out.push(res);
      log.push({ sheet: ws.name, scale: res.scale });
    });
    if (!out.length) throw new Error("Хөрвүүлэх өгөгдөлтэй хуудас олдсонгүй.");
    const zip = await buildDocx(out);
    const result = { log, missingFormula: stats.missingFormula, zip };
    return result;
  }
  excelToWord.formatValue = formatValue;
  excelToWord.sheetNames = async function (data) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(data);
    const names = [];
    wb.eachSheet((ws) => { if (!ws.state || ws.state === "visible") names.push(ws.name); });
    return names;
  };
  return excelToWord;
});
