'use strict';
/**
 * 4-9 呼吸器リスク評価
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。q()/brk() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 図はすべてネイティブ図形（箱・表）。図のラベルは図の外に置く。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fits, fitInfo, heightFor, measure } = require('../_lib/decklib');

const NAME = '4-9_pulmonary_risk_assessment';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.RENDER_DIR || process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
const SPEC_MD = fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8');
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const lines = part.split('\n');
    let notes = null, sources = null;
    const quoted = new Set();
    let inOn = false;
    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i];
      if (/^notes:\s*\|\s*$/.test(ln)) {
        const buf = [];
        for (let j = i + 1; j < lines.length && (lines[j].startsWith('  ') || lines[j] === ''); j++) {
          buf.push(lines[j].replace(/^ {2}/, ''));
        }
        while (buf.length && buf[buf.length - 1] === '') buf.pop();
        notes = buf.join('\n');
      }
      const m = ln.match(/^sources:\s*(.*)$/);
      if (m) sources = m[1].trim();
      if (/^on-slide:/.test(ln)) { inOn = true; continue; }
      if (inOn && !/^ {2}- /.test(ln)) inOn = false;
      // on-slide の「…」（入れ子の『』はそのまま）と、layout 行の「…」（表の列名）
      if (inOn || /^layout:/.test(ln)) for (const qm of ln.matchAll(/「([^」]*)」/g)) quoted.add(qm[1]);
    }
    out[id] = { notes, sources: sources && sources !== '—' ? sources : null, quoted };
  }
  return out;
}
const SPEC = parseSpec(SPEC_MD);
function notesOf(id) {
  const n = SPEC[id] && SPEC[id].notes;
  if (!n) throw new Error(`spec.md の ${id} に notes がありません`);
  return n;
}
/** spec の文言そのまま（照合のみ） */
function q(id, str) {
  if (!SPEC[id].quoted.has(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
  return str;
}
/**
 * spec の文言 str を、breaks の各語の直前で改行して行配列にする（文言は変えない）。
 * 改行位置の半角スペースは落とす。
 */
function brk(id, str, ...breaks) {
  q(id, str);
  const lines = [];
  let from = 0;
  for (const b of breaks) {
    const k = str.indexOf(b, from + 1);
    if (k < 0) throw new Error(`${id}: 「${str}」に改行位置「${b}」が見つかりません`);
    lines.push(str.slice(from, k).replace(/ $/, ''));
    from = k;
  }
  lines.push(str.slice(from));
  return lines;
}
/**
 * 1行の文字列に部分装飾を付けて run 配列にする。marks: [[語, 色], ...]（太字）。
 * 語は行の中に1回だけ現れること。
 */
function mark(line, ...marks) {
  let runs = [line];
  for (const [w, color] of marks) {
    const next = [];
    let found = 0;
    for (const r of runs) {
      if (typeof r !== 'string') { next.push(r); continue; }
      const k = r.indexOf(w);
      if (k < 0) { next.push(r); continue; }
      found++;
      if (k > 0) next.push(r.slice(0, k));
      next.push({ text: w, color, bold: true });
      if (k + w.length < r.length) next.push(r.slice(k + w.length));
    }
    if (found !== 1) throw new Error(`mark: 「${line}」に「${w}」が ${found} 回`);
    runs = next;
  }
  return runs;
}
const runText = l => (Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(l));
const toRuns = l => (Array.isArray(l) ? l : [l]).map(r => (typeof r === 'string' ? { text: r } : r));

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 文言はそのままで、文献の途中の語の切れ目で改行する（区切り > 著者「et al.」の後 > 誌名と年の間 > その他の空白）。
//  2) 入らなければ区切りの前後の空白を詰める（' / ' → '/'）。
//  3) それでも入らなければ（）内の補足を後ろの文献から外す。
//  4) それでも入らないときだけ、drop で指定した文献を落とす（DEVIATIONS に記録）。
// ---------------------------------------------------------------------------
const SRC = LAYOUT.source;
const DEVIATIONS = [];
function libLines(items) {
  const lines = [];
  for (const it of items) {
    const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
    if (cur && fits(cur, SRC.size, SRC.w)) lines[lines.length - 1] = cur;
    else lines.push(it);
  }
  return lines;
}
function split2(str, sep) {
  const fitsLine = l => fits(l, SRC.size, SRC.w);
  const cands = [];
  for (let i = 1; i < str.length; i++) {
    let a, b, rank;
    if (str.startsWith(sep, i)) { a = str.slice(0, i); b = str.slice(i + sep.length); rank = 0; }
    else if (str[i] === ' ') {
      a = str.slice(0, i); b = str.slice(i + 1);
      if (a.endsWith('/') || b.startsWith('/')) continue;
      if (/ et$/.test(a)) continue;                   // 「et」と「al.」の間では切らない
      if (a.endsWith('al.')) rank = 1;                 // 著者名（et al.）の後
      else if (/^\d/.test(b)) rank = 2;                // 誌名と年の間
      else rank = 3;
    } else continue;
    a = a.trim(); b = b.trim();
    if (a && b && fitsLine(a) && fitsLine(b)) cands.push({ a, b, rank });
  }
  if (!cands.length) return null;
  cands.sort((x, y) => x.rank - y.rank || y.a.length - x.a.length);
  return [cands[0].a, cands[0].b];
}
function sourcesFor(id, o = {}) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  let items = spec.split(' / ');
  if (o.drop) {
    const k = items.findIndex(it => it.startsWith(o.drop));
    if (k < 0) throw new Error(`${id}: 出典に「${o.drop}」がありません`);
    // 1)〜3) で 2 行に入らないことを確かめてから落とす
    const all = items.join(' / ');
    if (libLines(items).length <= 2 || split2(all, ' / ') || split2(items.join('/'), '/') ||
      split2(items.map(it => it.replace(/（[^）]*）$/, '')).join('/'), '/')) {
      throw new Error(`${id}: 出典は文献を落とさずに 2 行に入ります（drop 不要）`);
    }
    DEVIATIONS.push(`${id} 出典: 「${items[k]}」を落とした（${items.length} 件は区切りの空白を詰め（）内を外しても 16pt 2 行に入らない。全文はノートと前のスライドの出典にある）`);
    items = items.filter((_, i) => i !== k);
  }
  const base = items.join(' / ');
  if (libLines(items).length <= 2) return items;
  let r = split2(base, ' / ');
  if (r) { console.log(`  ${id} 出典: 文献の途中で改行（文言はそのまま）`); return r; }
  r = split2(items.join('/'), '/');
  if (r) { DEVIATIONS.push(`${id} 出典: 区切りの空白を詰めた（' / ' → '/'）`); return r; }
  const cut = items.slice();
  for (let k = cut.length - 1; k >= 0; k--) {
    if (!/（[^）]*）$/.test(cut[k])) continue;
    const was = cut[k];
    cut[k] = cut[k].replace(/（[^）]*）$/, '');
    for (const sep of [' / ', '/']) {
      r = split2(cut.join(sep), sep);
      if (r) {
        DEVIATIONS.push(`${id} 出典: ${was} → ${cut[k]}${sep === '/' ? '、区切りの空白を詰めた' : ''}`);
        return r;
      }
    }
  }
  throw new Error(`${id}: 出典が 2 行に収まりません`);
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const TOP = 1.85;                // 本文の上端
const BOT = 6.46;                // 出典 2 行（枠 y6.5）のときの本文下端（全スライド出典 2 行）
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const CM = 2.54;
const MEIRYO_EM = 1.5;           // 実機メイリオの 1 行高（em、行間 100%）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;

// ---------------------------------------------------------------------------
// 行頭記号つきの項目群（1 つのテキスト枠。S2 のカード・1 行の強調ボックス用）
//  - 1 項目を 2 行以上に割ったときは、2 行目以降を記号の右にぶら下げる（見えない記号 U+3000 で字下げ）。
//  - ls: 行間（倍率。1 = 既定）。gap: 項目の間の段落後の間隔（pt）。
//  幅の検査は decklib と同じ推定器（行頭記号ぶんを差し引く）。高さは実機メイリオ推定（1.5 em × ls）で取り、
//  lint の推定（pt × ls × 1.18、段落後 未指定 = 6pt）が枠高 + 0.3cm を超えないことも確かめる。
//  item = 行 | {lines:[行...], color, bold}（行 = 文字列 | run 配列）
// ---------------------------------------------------------------------------
function normItems(items) {
  return items.map(it => (it && typeof it === 'object' && !Array.isArray(it) && it.lines ? it : { lines: [it] }));
}
function itemTextH(items, o) {
  const size = o.size || 22, ls = o.ls || 1, gap = o.gap != null ? o.gap : 8;
  const its = normItems(items);
  const n = its.reduce((a, it) => a + it.lines.length, 0);
  return n * size / 72 * MEIRYO_EM * ls + (its.length - 1) * gap / 72;
}
function itemText(s, items, o) {
  const size = o.size || 22, ls = o.ls || 1, gap = o.gap != null ? o.gap : 8;
  const marker = o.marker !== false;
  const ind = marker ? (o.indent != null ? o.indent : 0.28) : 0;
  const indPt = Math.round(ind * 72);
  const its = normItems(items);
  const where = `p${s.page}「${s.title}」 項目「${runText(its[0].lines[0])}」`;
  its.forEach(it => it.lines.forEach(l => {
    const fi = fitInfo(runText(l), size, o.w - ind);
    if (!fi.ok) throw new Error(`${where}: 「${runText(l)}」が幅に収まりません（推定 ${fi.needIn}in > ${r3(o.w - ind)}in、約${fi.overChars}字オーバー）`);
  }));
  const need = itemTextH(its, o);
  const h = o.h != null ? o.h : Math.ceil(need * 100) / 100;
  if (need > h + 1e-6) throw new Error(`${where}: 高さ不足（必要 ${r3(need)}in > 枠 ${r3(h)}in）`);
  const lintLs = ls !== 1 ? ls : 1.2;
  let est = 0;
  its.forEach((it, i) => it.lines.forEach((l, j) => {
    const sa = (j === it.lines.length - 1 && i < its.length - 1 && gap > 0) ? gap : 6;
    est += size / 72 * lintLs * 1.18 + sa / 72;
  }));
  if (est > h + 0.3 / CM) throw new Error(`${where}: lint 推定 ${r3(est)}in > 枠 ${r3(h)}in + 0.12in`);
  s._bounds(where, o.x, o.y, o.w, h);
  const total = its.reduce((a, it) => a + it.lines.length, 0);
  const runs = [];
  let pi = 0;
  its.forEach((it, i) => it.lines.forEach((l, j) => {
    const rr = toRuns(l);
    rr.forEach((r, k) => {
      const opt = {
        fontFace: FONT, lang: 'ja-JP', fontSize: size,
        color: r.color || it.color || o.color || C.dark,
      };
      if (r.bold || it.bold || o.bold) opt.bold = true;
      if (k === 0) {
        opt.bullet = marker ? (j === 0 ? { indent: indPt } : { indent: indPt, characterCode: '3000' }) : false;
        if (ls !== 1) opt.lineSpacingMultiple = ls;
        if (j === it.lines.length - 1 && i < its.length - 1 && gap > 0) opt.paraSpaceAfter = gap;
      }
      if (k === rr.length - 1 && pi < total - 1) opt.breakLine = true;
      runs.push({ text: r.text, options: opt });
    });
    pi++;
  }));
  s.raw.addText(runs, {
    x: o.x, y: o.y, w: o.w, h, margin: 0, isTextBox: true,
    fontFace: FONT, lang: 'ja-JP', fontSize: size,
    align: o.align || 'left', valign: o.valign || 'top',
  });
  s._reg('txt', o.x, o.y, o.w, h, `items「${runText(its[0].lines[0]).slice(0, 12)}」`);
  return { x: o.x, y: o.y, w: o.w, h, bottom: r3(o.y + h) };
}

// ---------------------------------------------------------------------------
// 行送りを固定した項目リスト（S5・S6・S7 用）。枠線・塗りなしの表（2 列: 行頭記号 | 本文、1 項目 = 1 行）で組む。
//  テキスト枠の行間（倍率）は描画系で 1 行の高さが変わる（実機メイリオ 1.5 em、LibreOffice の代替フォント約 1.2 em）ため、
//  2〜3 行に割れる項目が多いカードは、3-11 の tightTable と同じく行送りを固定値（1.2 em = 26.4pt）にできる表のセルで組む。
//  2 行目以降は本文の列に入るので、行頭記号の右にぶら下がる。項目の間は行の下余白（gap pt）。
//  幅は decklib と同じ推定器で、高さは行数 × 行送り＋下余白で検査する。重なり検査用の登録も行う。
//  （pptxgenjs はセル余白の先頭値が 1 未満だとインチとみなすので、余白はインチで渡す）
// ---------------------------------------------------------------------------
function itemTableH(items, o) {
  const size = o.size || 22, lsEm = o.lsEm || 1.2, gap = o.gap != null ? o.gap : 6;
  const its = normItems(items);
  const n = its.reduce((a, it) => a + it.lines.length, 0);
  return n * size * lsEm / 72 + (its.length - 1) * gap / 72;
}
function itemTable(s, items, o) {
  const size = o.size || 22, lsEm = o.lsEm || 1.2, gap = o.gap != null ? o.gap : 6;
  const LS = size * lsEm;   // pt
  const marker = o.marker !== false;
  const ind = marker ? (o.indent != null ? o.indent : 0.28) : 0;
  const its = normItems(items);
  const where = `p${s.page}「${s.title}」 項目「${runText(its[0].lines[0])}」`;
  its.forEach(it => it.lines.forEach(l => {
    const fi = fitInfo(runText(l), size, o.w - ind);
    if (!fi.ok) throw new Error(`${where}: 「${runText(l)}」が幅に収まりません（推定 ${fi.needIn}in > ${r3(o.w - ind)}in、約${fi.overChars}字オーバー）`);
  }));
  const none = () => [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }];
  const rowH = [];
  const rows = its.map((it, i) => {
    const padB = i < its.length - 1 ? r3(gap / 72) : 0;
    const cellOpt = { margin: [0, 0, padB, 0], valign: 'top', border: none(), fontFace: FONT, fontSize: size };
    const text = [];
    it.lines.forEach((l, j) => {
      const rr = toRuns(l);
      rr.forEach((r, k) => text.push({
        text: r.text,
        options: {
          fontFace: FONT, lang: 'ja-JP', fontSize: size, color: r.color || it.color || o.color || C.dark,
          bold: (r.bold || it.bold || o.bold) || undefined, lineSpacing: LS,
          breakLine: (k === rr.length - 1 && j < it.lines.length - 1) || undefined,
        },
      }));
    });
    rowH.push(r3(it.lines.length * LS / 72 + padB));
    const body = { text, options: { ...cellOpt, align: o.align || 'left' } };
    if (!marker) return [body];
    const mk = {
      text: [{ text: '•', options: { fontFace: FONT, lang: 'ja-JP', fontSize: size, color: C.dark, lineSpacing: LS } }],
      options: { ...cellOpt, align: 'left' },
    };
    return [mk, body];
  });
  const h = r3(rowH.reduce((a, b) => a + b, 0));
  if (o.h != null && h > o.h + 1e-6) throw new Error(`${where}: 高さ不足（必要 ${h}in > 枠 ${r3(o.h)}in）`);
  const colW = marker ? [ind, r3(o.w - ind)] : [o.w];
  s._bounds(where, o.x, o.y, o.w, h);
  s.raw.addTable(rows, { x: o.x, y: o.y, w: o.w, colW, rowH, fontFace: FONT, fontSize: size });
  s._reg('box', o.x, o.y, o.w, h, `items表「${runText(its[0].lines[0]).slice(0, 12)}」`);
  return { x: o.x, y: o.y, w: o.w, h, bottom: r3(o.y + h) };
}

/**
 * 見出し付きカード（角丸の灰色地 + 24pt 太字の見出し + 行頭記号つき項目 22pt）。
 * table: true なら本文を itemTable（行送り固定 1.2 em）、false なら itemText（テキスト枠、行間 ls 倍）で組む。
 */
const CARD = { padX: 0.14, padY: 0.1, titleH: 0.45, titleGap: 0.04, titleSize: 24, size: 22, radius: 0.08 };
const bodyH = o => (o.table ? itemTableH(o.items, { size: CARD.size, gap: o.gap }) : itemTextH(o.items, { size: CARD.size, ls: o.ls, gap: o.gap }));
function listCardNeed(o) {
  const padY = o.padY != null ? o.padY : CARD.padY;
  return 2 * padY + CARD.titleH + CARD.titleGap + bodyH(o);
}
function listCard(s, o) {
  const padY = o.padY != null ? o.padY : CARD.padY;
  if (listCardNeed(o) > o.h + 1e-6) throw new Error(`p${s.page} カード「${o.title}」: 高さ不足（必要 ${r3(listCardNeed(o))}in > ${o.h}in）`);
  s.box({ x: r3(o.x), y: r3(o.y), w: r3(o.w), h: r3(o.h), fill: C.bg, radius: CARD.radius });
  const ix = r3(o.x + CARD.padX), iw = r3(o.w - 2 * CARD.padX);
  s.text(o.title, { x: ix, y: r3(o.y + padY), w: iw, h: CARD.titleH, size: CARD.titleSize, bold: true, color: o.titleColor || C.dark });
  const by = r3(o.y + padY + CARD.titleH + CARD.titleGap);
  const bh = r3(o.y + o.h - padY - by);
  if (o.table) return itemTable(s, o.items, { x: ix, y: by, w: iw, h: bh, size: CARD.size, gap: o.gap });
  return itemText(s, o.items, { x: ix, y: by, w: iw, h: bh, size: CARD.size, ls: o.ls, gap: o.gap });
}

/**
 * 強調ボックス（白地＋ティールの枠・太字・中央揃え）。lines は 1〜2 行（各行 文字列 | run 配列）。
 * table: true なら文字を itemTable（行送り固定 1.2 em）で箱の上下中央に置く。
 */
const emphTextH = o => (o.table ? itemTableH(o.lines, { size: o.size || 24, gap: 0 }) : itemTextH(o.lines, { size: o.size || 24, ls: 1, gap: 0 }));
const emphNeed = o => emphTextH(o) + 2 * (o.margin != null ? o.margin : 0.08);
function emphBox(s, o) {
  const m = o.margin != null ? o.margin : 0.08;
  s.box({ x: r3(o.x), y: r3(o.y), w: r3(o.w), h: r3(o.h), fill: C.white, radius: 0.08, line: { color: o.lineColor || C.teal, width: 2.5 } });
  const common = { x: r3(o.x + m), w: r3(o.w - 2 * m), size: o.size || 24, gap: 0, marker: false, align: 'center', bold: true, color: o.color || C.dark };
  if (o.table) {
    const th = emphTextH(o);
    return itemTable(s, o.lines, { ...common, y: r3(o.y + (o.h - th) / 2) });
  }
  return itemText(s, o.lines, { ...common, y: r3(o.y + m), h: r3(o.h - 2 * m), ls: 1, valign: 'middle' });
}

/**
 * 行の多い表（S3・S8 用。3-11 の tightTable を写したもの）。decklib の s.table と同じ見た目
 * （見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、上下中央）で、行間だけを固定値 LS pt（= 1.2 em）に詰める。
 * s.table は行高を実機メイリオの 1.5 em で取るため、行の多い表が本文ゾーン（出典 2 行時 1.85〜6.45in）に入らない。
 * セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、重なり検査用の登録も同じく行う。
 * cell: 文字列（'\n' で改行）| {text, color, bold, align}
 */
function tightTable(s, rows, o) {
  const size = o.size || 22, LS = o.lineSpacing || size * 1.2;
  const m = o.margin || [3, 6, 3, 6];   // pt: 上, 右, 下, 左
  const colW = o.colW, nCol = colW.length;
  const w = colW.reduce((a, b) => a + b, 0);
  const rowH = [];
  const out = rows.map((row, ri) => {
    if (row.length !== nCol) throw new Error(`p${s.page} 表 行${ri + 1}: 列数 ${row.length} ≠ ${nCol}`);
    const head = ri === 0;
    let maxLines = 1;
    const cells = row.map((cell, ci) => {
      const c = typeof cell === 'string' ? { text: cell } : cell;
      const lines = c.text.split('\n');
      const avail = colW[ci] - (m[1] + m[3]) / 72;
      lines.forEach(l => {
        const fi = fitInfo(l, size, avail);
        if (!fi.ok) {
          throw new Error(`p${s.page} 表 行${ri + 1}列${ci + 1}「${l}」がセル幅に収まりません` +
            `（推定 ${fi.needIn}in > ${avail.toFixed(3)}in、約${fi.overChars}字オーバー）`);
        }
      });
      maxLines = Math.max(maxLines, lines.length);
      const color = c.color || C.dark, bold = c.bold != null ? c.bold : head;
      return {
        text: lines.map((l, li) => ({
          text: l,
          options: {
            fontFace: FONT, lang: 'ja-JP', fontSize: size, color, bold: bold || undefined,
            lineSpacing: LS, breakLine: li < lines.length - 1 || undefined,
          },
        })),
        options: {
          fill: head ? { color: C.lightGray } : undefined,
          align: c.align || 'left', valign: 'middle',
          fontFace: FONT, fontSize: size, color, bold: bold || undefined,
        },
      };
    });
    rowH.push(Math.ceil((maxLines * LS / 72 + (m[0] + m[2]) / 72) * 100 - 1e-6) / 100);
    return cells;
  });
  const h = rowH.reduce((a, b) => a + b, 0);
  s._bounds(`p${s.page} 表`, o.x, o.y, w, h);
  s.raw.addTable(out, {
    x: o.x, y: o.y, w, colW, rowH, fontFace: FONT, fontSize: size, margin: m,
    border: { type: 'solid', pt: 1, color: C.border }, valign: 'middle',
  });
  s._reg('box', o.x, o.y, w, h, 'table');
  return { x: o.x, y: o.y, w, h: r2(h), rowH };
}

function assertBottom(id, what, bottom, limit = BOT) {
  if (bottom > limit + 1e-6) throw new Error(`${id} ${what}の下端 ${r3(bottom)}in > ${limit}in`);
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['リスク', '禁煙', 'COPD\n喘息', 'OSA・\n小児', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-9 呼吸器リスク評価',
  });

  // ===== S1 表紙 =====
  deck.cover({
    title: '呼吸器リスク評価',
    subtitle: '4-9 術前評価',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 術後肺合併症のリスク（2 列カード + 強調ボックス） =====
  {
    const id = 'S2';
    const s = deck.slide({ title: '術後肺合併症のリスク', chapter: 0, notes: notesOf(id), sources: sourcesFor(id) });
    const cw = r3((BW - 0.25) / 2);
    const L = {
      title: q(id, '患者因子'), titleColor: C.blue, table: true, gap: 14, padY: 0.12,
      items: ['高齢（60 歳超で上がる）', 'ASA-PS 2 以上、介助が必要', 'COPD、うっ血性心不全', '1 か月以内の呼吸器感染', 'アルブミン 3.0 g/dL 未満、貧血'].map(t => q(id, t)),
    };
    const R = {
      title: q(id, '手術・麻酔因子'), titleColor: C.vermilion, table: true, gap: 14, padY: 0.12,
      items: ['上腹部・胸腔内の切開', '長時間手術（2 時間超から）', '緊急手術', '全身麻酔（区域麻酔より）', '残存筋弛緩（長時間作用型の薬）'].map(t => q(id, t)),
    };
    const ch = r2(Math.max(listCardNeed(L), listCardNeed(R)) + 0.01);
    listCard(s, { ...L, x: X0, y: TOP, w: cw, h: ch });
    listCard(s, { ...R, x: r3(XR - cw), y: TOP, w: cw, h: ch });
    const E = { lines: [mark(q(id, 'PPC が起きると 30 日死亡率は 19.5%（起きなければ 0.5%）'), ['19.5%', C.red])], size: 26 };
    const eh = r2(emphNeed(E) + 0.04);
    const ey = r2(TOP + ch + Math.max(GAP + 0.06, (BOT - TOP - ch - eh) / 2));
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(id, '強調ボックス', ey + eh);
    console.log(`  ${id}: カード h ${ch}（〜${r2(TOP + ch)}）、強調 ${ey}〜${r2(ey + eh)}`);
  }

  // ===== S3 ARISCAT スコア（表 + 下に 1 行） =====
  {
    const id = 'S3';
    const s = deck.slide({ title: 'ARISCAT スコア', chapter: 0, notes: notesOf(id), sources: sourcesFor(id) });
    const ctr = text => ({ text: q(id, text), align: 'center' });
    const rows = [
      [q(id, '因子'), q(id, '区分'), { text: q(id, '点数'), align: 'center' }],
      [q(id, '年齢'), q(id, '51〜80 歳 / 80 歳超'), ctr('3 / 16')],
      [q(id, '術前 SpO2（室内気・仰臥位）'), q(id, '91〜95% / 90% 以下'), ctr('8 / 24')],
      [q(id, '1 か月以内の呼吸器感染'), q(id, 'あり'), ctr('17')],
      [q(id, '術前貧血'), q(id, 'Hb 10 g/dL 以下'), ctr('11')],
      [q(id, '切開部位'), q(id, '上腹部 / 胸腔内'), ctr('15 / 24')],
      [q(id, '手術時間'), q(id, '2〜3 時間 / 3 時間超'), ctr('16 / 23')],
      [q(id, '緊急手術'), q(id, 'あり'), ctr('8')],
    ];
    // 列幅: spec の目安 3.6 / 6.0 / 2.6 では「術前 SpO2（室内気・仰臥位）」「1 か月以内の呼吸器感染」が 2 行になり
    // 表＋下の 1 行が本文ゾーンに入らないため、1 列目を広げて全行 1 行にする（合計 12.2in）
    const t = tightTable(s, rows, { x: X0, y: TOP, colW: [4.85, 4.75, 2.6] });
    const line = mark(q(id, 'PPC：低 26 点未満 1.6%、中 26〜44 点 13.3%、高 45 点以上 42.1%'),
      ['1.6%', C.blue], ['13.3%', C.vermilion], ['42.1%', C.red]);
    const lh = heightFor(1, 22, 0);
    const ly = r2(t.y + t.h + Math.max(GAP + 0.06, (BOT - t.y - t.h - lh) / 2));
    s.text([line], { x: X0, y: ly, w: BW, h: lh, size: 22 });
    assertBottom(id, '下の 1 行', ly + lh);
    console.log(`  ${id}: 表 ${t.y}〜${r2(t.y + t.h)}、下の 1 行 ${ly}〜${r2(ly + lh)}`);
  }

  // ===== S4 禁煙は遅すぎない（箇条書き 5 行 + 強調ボックス） =====
  {
    const id = 'S4';
    const s = deck.slide({ title: '禁煙は遅すぎない', chapter: 1, notes: notesOf(id), sources: sourcesFor(id) });
    const items = [
      q(id, '喫煙者は術後の肺炎、再挿管、心筋梗塞、死亡が多い'),
      mark(q(id, '4 週以上の禁煙で呼吸器合併症が減る（8 週以上でさらに）'), ['4 週以上', C.vermilion]),
      q(id, '3〜4 週以上で創の合併症も減る'),
      q(id, '直前の禁煙でも合併症は増えない（『痰が増えて危ない』は誤り）'),
      q(id, '加熱式たばこもニコチンを含む。周術期は中止の対象'),
    ];
    const size = 24, gap = 14;
    const bh = heightFor(items.length, size, gap);
    s.bullets(items, { y: TOP, h: bh, size, gap });
    const E = { lines: [mark(q(id, 'JSA：禁煙はいつ始めても意義がある。できれば 4 週以上'), ['4 週以上', C.vermilion])], size: 28 };
    const eh = r2(emphNeed(E) + 0.06);
    const ey = r2(TOP + bh + (BOT - TOP - bh - eh) / 2);
    if (ey - (TOP + bh) < GAP - 1e-6) throw new Error(`${id}: 箇条書きと強調ボックスの間 ${r3(ey - TOP - bh)}in < ${GAP}`);
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(id, '強調ボックス', ey + eh);
    console.log(`  ${id}: 箇条書き h ${bh}（〜${r2(TOP + bh)}）、強調 ${ey}〜${r2(ey + eh)}`);
  }

  // ===== S5 COPD と喘息（2 列カード） =====
  {
    const id = 'S5';
    const s = deck.slide({ title: 'COPD と喘息', chapter: 2, notes: notesOf(id), sources: sourcesFor(id) });
    // 喘息のカードは 2〜3 行に割れる項目が多く 10 行になる。行送り固定 1.2 em（itemTable）で本文ゾーンに収める
    const L = {
      title: q(id, 'COPD'), titleColor: C.blue, table: true, gap: 12,
      items: [
        { lines: brk(id, '吸入薬は当日も続ける（持参させる）', '（持参') },
        q(id, '増悪中は延期。症状が元に戻るまで'),
        { lines: brk(id, '未治療の COPD は術前に治療を始める', '術前に') },
        q(id, '肺機能検査はルーチンには不要'),
        q(id, '上腹部・長時間手術でリスクが高い'),
      ],
    };
    const nsaids = brk(id, 'NSAIDs 過敏喘息：NSAIDs とコハク酸エステル型ステロイドの急速静注を避ける', 'NSAIDs と', 'ステロイドの');
    const R = {
      title: q(id, '喘息'), titleColor: C.vermilion, table: true, gap: 4,
      items: [
        { lines: brk(id, 'コントロール良好な時期に手術。ICS は続ける', 'ICS') },
        q(id, '不良なら延期して呼吸器内科へ'),
        { lines: brk(id, '延期できなければ術前 5 日の全身ステロイド＋β2 刺激薬', '全身') },
        { lines: [[{ text: nsaids[0], color: C.red, bold: true }], nsaids[1], nsaids[2]] },
        { lines: brk(id, 'できれば挿管を避ける（区域麻酔、声門上器具）', '（区域') },
      ],
    };
    const ch = r2(BOT - TOP);
    const wL = 6.04, wR = r3(BW - 0.25 - wL);
    listCard(s, { ...L, x: X0, y: TOP, w: wL, h: ch });
    listCard(s, { ...R, x: r3(XR - wR), y: TOP, w: wR, h: ch });
    console.log(`  ${id}: カード h ${ch}（必要 COPD ${r3(listCardNeed(L))} / 喘息 ${r3(listCardNeed(R))}）`);
  }

  // ===== S6 OSA と STOP-Bang（左: 8 項目、右: 箇条書き 4 項目） =====
  {
    const id = 'S6';
    const s = deck.slide({ title: 'OSA と STOP-Bang', chapter: 3, notes: notesOf(id), sources: sourcesFor(id) });
    const sb = ['S いびき（ドア越しに聞こえる）', 'T 日中の眠気・疲労', 'O 無呼吸を指摘された', 'P 高血圧',
      'B BMI 35 超', 'A 50 歳超', 'N 首回り 40 cm 超', 'G 男性'].map(t => {
      q(id, t);
      const k = t.indexOf(' ');
      return [t.slice(0, k), t.slice(k + 1)];
    });
    // 左: 頭文字（24pt 太字・ブルー）と項目（22pt）を行ごとに並べる。灰色の地にのせる
    const LET_W = 0.42, PADL = 0.18, PADR = 0.16, LGAP = 0.1, PADY = 0.12;
    const txtW = r2(Math.max(...sb.map(([, t]) => measure(t, 22))) + 0.02);
    const pw = r2(PADL + LET_W + LGAP + txtW + PADR);
    const ph = r2(BOT - TOP);
    s.box({ x: X0, y: TOP, w: pw, h: ph, fill: C.bg, radius: 0.08 });
    const pitch = (ph - 2 * PADY) / sb.length;
    const rowH = 0.5;
    sb.forEach(([letter, text], i) => {
      const y = r3(TOP + PADY + i * pitch + (pitch - rowH) / 2);
      s.text(letter, { x: r3(X0 + PADL), y, w: LET_W, h: rowH, size: 24, bold: true, color: C.blue, align: 'center', valign: 'middle' });
      s.text(text, { x: r3(X0 + PADL + LET_W + LGAP), y, w: txtW, h: rowH, size: 22, valign: 'middle' });
    });
    // 右: 4 項目（2〜3 行に割れる項目は行送り 1.2 em でぶら下げ）
    const rx = r2(X0 + pw + 0.32), rw = r3(XR - rx);
    const items = [
      { lines: brk(id, '3 点以上で OSA の可能性、5 点以上で高リスク', '5 点').map((l, i) => (i === 0 ? mark(l, ['3 点以上', C.blue]) : mark(l, ['5 点以上で高リスク', C.vermilion]))) },
      { lines: brk(id, 'OSA で呼吸不全（オッズ比 2.4）、ICU 転床（2.8）', 'ICU') },
      { lines: brk(id, '疑いだけでは延期しない。低換気・重症肺高血圧・原因不明の安静時低酸素なら追加評価', '低換気', '原因不明') },
      { lines: brk(id, 'CPAP は持参。困難気道として備え、区域麻酔を優先', '困難気道') },
    ];
    const o = { size: 22, gap: 18 };
    const ih = r3(itemTableH(items, o));
    const iy = r2(TOP + (ph - ih) / 2);
    itemTable(s, items, { ...o, x: rx, y: iy, w: rw });
    console.log(`  ${id}: 左の地 w ${pw}（〜${r2(X0 + pw)}）、右 x ${rx} w ${rw}、右の高さ ${ih}（${iy}〜${r2(iy + ih)}）`);
  }

  // ===== S7 小児の上気道炎（2 列カード + 強調ボックス） =====
  {
    const id = 'S7';
    const s = deck.slide({ title: '小児の上気道炎', chapter: 3, notes: notesOf(id), sources: sourcesFor(id) });
    const L = {
      title: q(id, '延期を考える'), titleColor: C.vermilion, table: true, gap: 8, padY: 0.08,
      items: [
        q(id, '38℃ 超の発熱、ぐったりしている'),
        q(id, '膿性鼻汁、湿性咳嗽'),
        q(id, '喘鳴・ラ音（下気道の所見）'),
        { lines: brk(id, '上気道炎が今ある、または 2 週以内ならリスク約 2 倍', '2 週') },
        q(id, '気管挿管が要る、気道の手術'),
      ],
    };
    const R = {
      title: q(id, '進める・工夫する'), titleColor: C.blue, table: true, gap: 4.5, padY: 0.08,
      items: [
        q(id, '透明な鼻汁だけなら多くは可'),
        { lines: brk(id, '上気道炎から 2〜4 週たてばリスクは下がる', 'リスク') },
        q(id, 'フェイスマスクか声門上器具で'),
        q(id, 'プロポフォールで静脈導入'),
        { lines: brk(id, '術前にサルブタモール吸入を考える（導入 10〜30 分前）', '（導入') },
      ],
    };
    const E = {
      lines: brk(id, '喉頭痙攣・気管支攣縮・低酸素に備える。迷えば延期し、保護者と主治医に共有', '迷えば')
        .map((l, i) => (i === 1 ? mark(l, ['迷えば延期し', C.red]) : l)),
      size: 22, table: true, margin: 0.07,
    };
    const eh = r2(emphNeed(E) + 0.02);
    const ey = r2(BOT - eh);
    const ch = r2(ey - GAP - TOP);
    const wL = 5.93, wR = r3(BW - 0.25 - wL);
    listCard(s, { ...L, x: X0, y: TOP, w: wL, h: ch });
    listCard(s, { ...R, x: r3(XR - wR), y: TOP, w: wR, h: ch });
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    console.log(`  ${id}: カード h ${ch}（必要 ${r3(listCardNeed(L))} / ${r3(listCardNeed(R))}）、強調 ${ey}〜${r2(ey + eh)}`);
  }

  // ===== S8 肺機能検査と相談先（表 3 列 × 5 行＋見出し行） =====
  {
    const id = 'S8';
    const s = deck.slide({ title: '肺機能検査と相談先', chapter: 3, notes: notesOf(id), sources: sourcesFor(id) });
    const cell = (str, ...breaks) => brk(id, str, ...breaks).join('\n');
    const rows = [
      [q(id, '状況'), q(id, 'どうする'), q(id, '相談先')],
      [q(id, '心臓・胸部以外の手術'), cell('肺機能検査・胸部 X 線の値だけで可否を決めない（ACP）', '可否'),
        cell('原因不明の呼吸困難なら呼吸器内科', '呼吸困難', '呼吸器内科')],
      [q(id, '肺切除'), cell('FEV1・DLCO の予測術後値が 60% 超なら低リスク', '60%'), cell('呼吸器外科・呼吸器内科', '呼吸器内科')],
      [cell('コントロール不良の喘息、COPD 増悪', '喘息'), cell('延期して治療。症状が元に戻るまで', '症状'), q(id, '呼吸器内科')],
      [cell('OSA に低換気・重症肺高血圧・安静時低酸素', '重症', '安静時'), q(id, '術前に追加評価（必要なら延期）'),
        cell('呼吸器・睡眠の専門医', '専門医')],
      [q(id, '成人の呼吸器感染後'), q(id, '約 6 週の延期を勧める総説がある'), q(id, '主科と相談')],
    ];
    // 列幅は spec の目安 3.8 / 5.2 / 3.2 を、12 行（見出し 1 ＋本文 11）で収まる配分に見直した（合計 12.2in）。
    // 行送り 1.2 em の 12 行（4.4in）を本文ゾーン（1.85〜6.47in）に入れるため、セルの上下余白を 1pt にする。
    const t = tightTable(s, rows, { x: X0, y: TOP, colW: [3.59, 5.36, 3.25], margin: [1, 6, 1, 6] });
    assertBottom(id, '表', t.y + t.h);
    console.log(`  ${id}: 表 ${t.y}〜${r2(t.y + t.h)}（行高 ${t.rowH.join(', ')}）`);
  }

  // ===== S9 まとめ（箇条書き 5 行） =====
  {
    const id = 'S9';
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: notesOf(id), sources: sourcesFor(id, { drop: 'Canet J' }) });
    const items = [
      mark(q(id, 'PPC 5%、起きると 30 日死亡 19.5%（なし 0.5%）。ARISCAT で評価'), ['19.5%', C.red]),
      mark(q(id, '禁煙はいつでも遅くない。4 週以上で呼吸器、3〜4 週で創'), ['いつでも遅くない', C.vermilion]),
      mark(q(id, '吸入薬は続ける。増悪・コントロール不良なら延期して相談'), ['延期して相談', C.vermilion]),
      mark(q(id, 'STOP-Bang 3 点以上で OSA を疑う。CPAP 持参、挿管に備える'), ['3 点以上', C.blue]),
      mark(q(id, '小児の上気道炎：発熱・膿性鼻汁・湿性咳嗽なら延期を考える'), ['延期を考える', C.vermilion]),
    ];
    const size = 24, gap = 24;
    const hh = heightFor(items.length, size, gap);
    if (hh > BOT - TOP + 1e-6) throw new Error(`${id} 箇条書きの高さ ${hh} > ${r2(BOT - TOP)}`);
    s.bullets(items, { y: TOP, h: r2(BOT - 0.02 - TOP), size, gap, valign: 'middle' });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
