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
    // Цаас: opts.paper = "A4" бол үргэлж A4, үгүй бол Excel-ийнхээр
    let [pw, ph] = opts.paper === "A4" ? PAPER[9] : PAPER[ps.paperSize] || PAPER[9];
    const landscape = opts.orientation === "landscape" ? true
      : opts.orientation === "portrait" ? false : ps.orientation === "landscape";
    if (landscape) [pw, ph] = [ph, pw];
    let margins;
    if (opts.margins) {
      // см → twip (1 см = 567 twip)
      const cm = (v) => Math.round(Number(v) * 567);
      margins = { left: cm(opts.margins.left), right: cm(opts.margins.right), top: cm(opts.margins.top),
        bottom: cm(opts.margins.bottom), header: 0, footer: 0 };
    } else {
      const mg = ps.margins || {};
      const tw = (v, d) => Math.round((v != null ? v : d) * 1440);
      margins = { left: tw(mg.left, 0.7), right: tw(mg.right, 0.7), top: tw(mg.top, 0.75),
        bottom: tw(mg.bottom, 0.75), header: tw(mg.header, 0.3), footer: tw(mg.footer, 0.3) };
    }
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

    // Нүдний текстийг Word run болгоно
    const cellText = (cell) => {
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
      return { body, jc, baseRpr, indent: al.indent || 0 };
    };
    const paragraph = (cell) => {
      const t = cellText(cell);
      let ind = "";
      if (t.indent) ind = `<w:ind w:${t.jc === "right" ? "right" : "left"}="${Math.round(t.indent * 180 * fontScale)}"/>`;
      return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>${ind}` +
        `<w:jc w:val="${t.jc}"/>${t.baseRpr}</w:pPr>${t.body}</w:p>`;
    };

    // Нүд бүрийн (merge тооцсон) хүрээ ба агуулгын нүд
    const cellBox = (r, c, info) => {
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
      return { spanCols, styleCell, contentCell, bT, bB, bL, bR };
    };

    // Хэвлэх гарчиг мөр
    let titleRows = new Set();
    const tm = /\$?(\d+):\$?(\d+)/.exec(ps.printTitlesRow || "");
    if (tm) for (let r = +tm[1]; r <= +tm[2]; r++) titleRows.add(r);

    const jcTbl = ps.horizontalCentered ? "center" : "left";
    const tblOffset = jcTbl === "center" ? Math.max(0, Math.floor((usableW - tableW) / 2)) : 0;
    const colLeft = {};
    { let x = tblOffset; for (const c of cols) { colLeft[c] = x; x += colTw[c]; } }
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
          const { spanCols, styleCell, contentCell, bT, bB, bL, bR } = cellBox(r, c, info);
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

    // Хүрээгүй мөрийг хүснэгтгүй энгийн бичвэр болгоно. Нүд бүрийн байрлалыг
    // догол (indent) болон tab-аар хадгална.
    const lineRpr = rPr(null);
    const emitPlain = (r) => {
      const hTw = Math.max(1, Math.floor(rowH[r] * 20 * scale));
      const items = [];
      for (const c of cols) {
        const info = mergeAt.get(key(r, c));
        if (info && info.skip) continue;
        if (info && info.start && !info.top) continue;
        const spanCols = info && info.start ? info.cols : [c];
        const cell = info && info.start ? cellAt(info.mr.minR, info.mr.minC) : cellAt(r, c);
        if (!cell) continue;
        const t = cellText(cell);
        if (!t.body) continue;
        const left = colLeft[spanCols[0]];
        const right = left + spanCols.reduce((s, cc) => s + colTw[cc], 0);
        items.push(Object.assign(t, { left, right }));
      }
      if (!items.length) {
        return `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${hTw}" w:lineRule="exact"/>` +
          `<w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>`;
      }
      const spacing = `<w:spacing w:before="0" w:after="0" w:line="${hTw}" w:lineRule="atLeast"/>`;
      if (items.length === 1) {
        const it = items[0];
        const indL = it.left + (it.jc !== "right" && it.indent ? Math.round(it.indent * 180 * fontScale) : 0);
        const indR = Math.max(0, usableW - it.right);
        return `<w:p><w:pPr>${spacing}<w:ind w:left="${indL}" w:right="${indR}"/>` +
          `<w:jc w:val="${it.jc === "distribute" ? "both" : it.jc}"/>${it.baseRpr}</w:pPr>${it.body}</w:p>`;
      }
      // Олон нүд → нэг мөр, нүд бүр өөрийн байрлалд tab-аар
      const first = items[0];
      const tabs = items.slice(1).map((it) => it.jc === "right" ? `<w:tab w:val="right" w:pos="${it.right - first.left}"/>`
        : it.jc === "center" ? `<w:tab w:val="center" w:pos="${Math.round((it.left + it.right) / 2) - first.left}"/>`
        : `<w:tab w:val="left" w:pos="${it.left - first.left}"/>`);
      const body = items.map((it, i) => (i ? `<w:r>${it.baseRpr}<w:tab/></w:r>` : "") + it.body).join("");
      return `<w:p><w:pPr><w:tabs>${tabs.join("")}</w:tabs>${spacing}<w:ind w:left="${first.left}"/>` +
        `<w:jc w:val="left"/>${first.baseRpr}</w:pPr>${body}</w:p>`;
    };

    // Мөр бүрийг "хүснэгт" эсвэл "энгийн" гэж ангилна
    const isTableRow = {};
    for (const r of rows) {
      if (!opts.plainOutside) { isTableRow[r] = true; continue; }
      let boxed = false;
      for (const c of cols) {
        const info = mergeAt.get(key(r, c));
        if (info && info.skip) continue;
        const b = cellBox(r, c, info);
        if ((b.bL && b.bR) || (b.bT && b.bB)) { boxed = true; break; }
      }
      isTableRow[r] = boxed;
    }
    // Нэгтгэсэн нүд хоёр төрлийн мөрийг огтолж болохгүй
    for (let changed = true; changed;) {
      changed = false;
      for (const mr of merges) {
        const span = rows.filter((r) => r >= mr.minR && r <= mr.maxR);
        if (span.some((r) => isTableRow[r]) && span.some((r) => !isTableRow[r])) {
          span.forEach((r) => { isTableRow[r] = true; }); changed = true;
        }
      }
    }

    // Дараалсан хүснэгтийн мөрүүдийг нэг хүснэгт болгоно. Давтагдах гарчиг
    // хүснэгтийн эхэнд байх ёстой тул гарчгийн өмнө салгана.
    const firstTitle = titleRows.size ? Math.min(...titleRows) : null;
    const canSplitAtTitle = firstTitle != null && !merges.some((mr) => mr.minR < firstTitle && firstTitle <= mr.maxR);
    const parts = [];
    let run = [];
    const flush = () => { if (run.length) { parts.push({ type: "tbl", xml: emit(run) }); run = []; } };
    for (const r of rows) {
      if (!isTableRow[r]) { flush(); parts.push({ type: "p", xml: emitPlain(r) }); continue; }
      if (r === firstTitle && canSplitAtTitle) flush();
      run.push(r);
    }
    flush();
    const sectPr = `<w:sectPr><w:pgSz w:w="${pw}" w:h="${ph}"${landscape ? ' w:orient="landscape"' : ""}/>` +
      `<w:pgMar w:top="${margins.top}" w:right="${margins.right}" w:bottom="${margins.bottom}" ` +
      `w:left="${margins.left}" w:header="${margins.header}" w:footer="${margins.footer}" w:gutter="0"/>` +
      `<w:cols w:space="720"/></w:sectPr>`;
    return { parts, sectPr, scale };
  }

  // ------------------------------------------------------- docx багц ---
  const TAIL_PPR = `<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>`;

  async function buildDocx(sheets) {
    let body = "";
    sheets.forEach((s, i) => {
      const last = i === sheets.length - 1;
      s.parts.forEach((p, j) => {
        body += p.xml;
        const next = s.parts[j + 1];
        // Хүснэгтийн араас paragraph хэрэгтэй (хоёр хүснэгт нийлэхгүй, sectPr байрлах)
        if (p.type === "tbl" && (!next || next.type === "tbl")) {
          body += `<w:p><w:pPr>${TAIL_PPR}${!next && !last ? s.sectPr : ""}</w:pPr></w:p>`;
        } else if (!next && !last) {
          body += `<w:p><w:pPr>${TAIL_PPR}${s.sectPr}</w:pPr></w:p>`;
        }
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

  // ------------------------------------------------- файл ачаалах ---
  // ExcelJS зураг, график, тайлбар (comment), "Format as Table" хүснэгт, pivot
  // зэрэгт алдаа заадаг. Эдгээр нь Word-д ордоггүй тул уншихаас өмнө хасна.
  const DROP_REL = /\/(drawing|vmlDrawing|comments|threadedComment|table|pivotTable|image|chart|ctrlProp|control|oleObject|package)$/;
  // Стандарт бус xlsx-ийг (x: угтвартай XML, Strict OOXML, өөр нэртэй
  // хэсгүүд — вэб систем, .NET программаас гаргасан файлд түгээмэл) ExcelJS-ийн
  // уншдаг энгийн бүтэц рүү шилжүүлнэ.
  const MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const STRICT_NS = [
    ["http://purl.oclc.org/ooxml/spreadsheetml/main", MAIN_NS],
    ["http://purl.oclc.org/ooxml/officeDocument/relationships", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"],
    ["http://purl.oclc.org/ooxml/drawingml/main", "http://schemas.openxmlformats.org/drawingml/2006/main"],
  ];
  function fixXml(x) {
    for (const [a, b] of STRICT_NS) x = x.split(a).join(b);
    for (const ns of [MAIN_NS, PKG_REL_NS]) {
      const re = new RegExp('xmlns:([A-Za-z_][\\w.-]*)="' + ns.replace(/[.\/]/g, "\\$&") + '"', "g");
      const prefixes = [];
      let m;
      while ((m = re.exec(x))) prefixes.push(m[1]);
      for (const p of prefixes) {
        x = x.replace(new RegExp("<(/?)" + p + ":", "g"), "<$1");
        const decl = 'xmlns:' + p + '="' + ns + '"';
        x = x.includes('xmlns="' + ns + '"') ? x.split(" " + decl).join("") : x.split(decl).join('xmlns="' + ns + '"');
      }
    }
    return x;
  }
  const resolvePath = (base, target) => {
    if (target.startsWith("/")) return target.slice(1);
    const parts = base.split("/"); parts.pop();
    for (const seg of target.split("/")) { if (seg === "..") parts.pop(); else if (seg && seg !== ".") parts.push(seg); }
    return parts.join("/");
  };
  const relsPathFor = (p) => { const i = p.lastIndexOf("/"); return p.slice(0, i + 1) + "_rels/" + p.slice(i + 1) + ".rels"; };
  const parseRels = (x) => (x.match(/<Relationship\b[^>]*>/g) || []).map((tag) => {
    const a = (n) => { const m = new RegExp(n + '="([^"]*)"').exec(tag); return m ? m[1] : ""; };
    return { Id: a("Id"), Type: a("Type"), Target: a("Target"), TargetMode: a("TargetMode") };
  });
  const relsXml = (rels) => '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + PKG_REL_NS + '">' +
    rels.map((r) => '<Relationship Id="' + r.Id + '" Type="' + r.Type + '" Target="' + esc(r.Target) + '"' +
      (r.TargetMode ? ' TargetMode="' + r.TargetMode + '"' : "") + "/>").join("") + "</Relationships>";

  async function normalizePackage(zip) {
    // 1) Бүх XML-ийн namespace-ийг засна
    const texts = [];
    zip.forEach((path, f) => { if (!f.dir && /\.(xml|rels)$/i.test(path)) texts.push(path); });
    for (const path of texts) zip.file(path, fixXml(await zip.file(path).async("string")));

    // 2) Хэсгүүдийн байрлалыг стандарт болгоно
    const rootRelsFile = zip.file("_rels/.rels");
    const rootRels = rootRelsFile ? parseRels(await rootRelsFile.async("string")) : [];
    const docRel = rootRels.find((r) => /\/officeDocument$/.test(r.Type));
    const wbPath = docRel ? resolvePath("", docRel.Target) : "xl/workbook.xml";
    if (!zip.file(wbPath)) return zip;
    const wbRelsFile = zip.file(relsPathFor(wbPath));
    const wbRels = wbRelsFile ? parseRels(await wbRelsFile.async("string")) : [];
    const sheetRels = wbRels.filter((r) => /\/worksheet$/.test(r.Type));
    const standard = wbPath === "xl/workbook.xml" && !!wbRelsFile &&
      wbRels.every((r) => r.TargetMode === "External" || !/^\//.test(r.Target)) &&
      sheetRels.every((r) => /^xl\/worksheets\/sheet\d+\.xml$/.test(resolvePath(wbPath, r.Target)));
    if (standard) return zip;

    const out = new JSZip();
    const copy = async (from, to) => { const f = zip.file(from); if (f) out.file(to, await f.async("uint8array")); return !!f; };
    await copy("[Content_Types].xml", "[Content_Types].xml");
    for (const p of ["docProps/core.xml", "docProps/app.xml"]) await copy(p, p);
    out.file("_rels/.rels", relsXml([{ Id: "rId1",
      Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Target: "xl/workbook.xml" }]));
    await copy(wbPath, "xl/workbook.xml");
    const newRels = [];
    let n = 0;
    for (const r of wbRels) {
      if (r.TargetMode === "External") continue;
      const src = resolvePath(wbPath, r.Target);
      let dst = null;
      if (/\/worksheet$/.test(r.Type)) dst = "xl/worksheets/sheet" + (++n) + ".xml";
      else if (/\/sharedStrings$/.test(r.Type)) dst = "xl/sharedStrings.xml";
      else if (/\/styles$/.test(r.Type)) dst = "xl/styles.xml";
      else if (/\/theme$/.test(r.Type)) dst = "xl/theme/theme1.xml";
      if (!dst || !(await copy(src, dst))) continue;
      newRels.push({ Id: r.Id, Type: r.Type, Target: dst.slice(3) });
      if (/\/worksheet$/.test(r.Type)) {
        const sr = zip.file(relsPathFor(src));
        if (sr) {
          const rels = parseRels(await sr.async("string")).map((x) => x.TargetMode === "External" ? x
            : Object.assign(x, { Target: "/" + resolvePath(src, x.Target) }));
          out.file(relsPathFor(dst), relsXml(rels));
        }
      }
    }
    out.file("xl/_rels/workbook.xml.rels", relsXml(newRels));
    return out;
  }

  async function sanitize(data) {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    if (bytes[0] === 0xD0 && bytes[1] === 0xCF) {
      throw new Error("Энэ файл хуучин .xls (Excel 97–2003) хэлбэртэй эсвэл нууц үгээр хамгаалагдсан байна. " +
        "Excel-ээр нээгээд (нууц үгтэй бол нууц үгийг нь авч) Файл → Save As → Excel Workbook (.xlsx) гэж хадгалаад дахин оруулна уу.");
    }
    if (bytes[0] !== 0x50 || bytes[1] !== 0x4B) {
      throw new Error("Файл Excel (.xlsx) бүтэцтэй биш байна. Excel-ээр нээгээд Excel Workbook (.xlsx) гэж дахин хадгална уу.");
    }
    let zip;
    try { zip = await JSZip.loadAsync(bytes); }
    catch (e) { throw new Error("Файл эвдэрсэн эсвэл нууц үгээр хамгаалагдсан байна."); }
    zip = await normalizePackage(zip);
    if (!zip.file("xl/workbook.xml")) {
      if (zip.file("EncryptedPackage")) throw new Error("Файл нууц үгээр хамгаалагдсан байна. Excel-ээр нээж нууц үгийг нь авч хадгалаад дахин оруулна уу.");
      throw new Error("Файл Excel (.xlsx) бүтэцтэй биш байна.");
    }
    const tasks = [];
    zip.forEach((path, f) => {
      if (/^xl\/worksheets\/_rels\/[^/]+\.rels$/.test(path)) {
        tasks.push(f.async("string").then((x) => zip.file(path,
          x.replace(/<Relationship\b[^>]*\/>/g, (rel) => {
            const t = /Type="([^"]*)"/.exec(rel);
            return t && DROP_REL.test(t[1]) ? "" : rel;
          }))));
      } else if (/^xl\/worksheets\/[^/]+\.xml$/.test(path)) {
        tasks.push(f.async("string").then((x) => zip.file(path, x
          .replace(/<(drawing|legacyDrawing|legacyDrawingHF|picture)\b[^>]*\/>/g, "")
          .replace(/<tableParts\b[\s\S]*?<\/tableParts>|<tableParts\b[^>]*\/>/g, "")
          .replace(/<controls\b[\s\S]*?<\/controls>/g, "")
          .replace(/<oleObjects\b[\s\S]*?<\/oleObjects>/g, ""))));
      } else if (path === "xl/workbook.xml") {
        tasks.push(f.async("string").then((x) => zip.file(path,
          x.replace(/<pivotCaches\b[\s\S]*?<\/pivotCaches>/g, ""))));
      }
    });
    await Promise.all(tasks);
    // Хамаарлаас салгасан хэсгүүдийг өөрсдийг нь ч устгана (ExcelJS тэднийг шууд уншдаг)
    zip.forEach((path) => {
      if (/^xl\/(drawings|charts|chartsheets|comments\d*\.xml|threadedComments|persons|tables|pivotTables|pivotCache|ctrlProps|activeX|embeddings)/.test(path)) zip.remove(path);
    });
    return zip.generateAsync({ type: "uint8array" });
  }

  async function loadWorkbook(data) {
    const clean = await sanitize(data);
    const wb = new ExcelJS.Workbook();
    try { await wb.xlsx.load(clean); }
    catch (e) { throw new Error("Excel файлыг уншиж чадсангүй (" + (e && e.message ? e.message : e) + ")."); }
    return wb;
  }

  // -------------------------------------------------------- үндсэн ---
  async function excelToWord(data, opts) {
    opts = Object.assign({ fit: true, scaleFonts: true, sheets: null, paper: null,
      orientation: "auto", margins: null, plainOutside: false }, opts || {});
    const wb = await loadWorkbook(data);
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
    const wb = await loadWorkbook(data);
    const names = [];
    wb.eachSheet((ws) => { if (!ws.state || ws.state === "visible") names.push(ws.name); });
    return names;
  };
  return excelToWord;
});
