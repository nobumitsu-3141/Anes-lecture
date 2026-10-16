'use strict';
/**
 * 4-8 心疾患別の注意点
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。q()/brk() が spec と照合する）、
 * タイトル・ノート・出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 図はすべてネイティブ図形（箱・線・表）。棒グラフも箱で描く。図のラベルは図の外に置く。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fits, fitInfo, heightFor, measure } = require('../_lib/decklib');

const NAME = '4-8_cardiac_conditions';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.RENDER_DIR || process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（タイトル・ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
const SPEC_MD = fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8');
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const lines = part.split('\n');
    const head = lines[0];
    const title = (head.match(/title=「(.+)」\s*$/) || [])[1];
    const ch = (head.match(/\bch=(\d+)/) || [])[1];
    let notes = null, sources = null, layout = '';
    const quoted = new Set();
    const onSlide = [];
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
      if (/^layout:/.test(ln)) layout = ln;
      if (/^on-slide:/.test(ln)) { inOn = true; continue; }
      if (inOn && !/^ {2}- /.test(ln)) inOn = false;
      if (inOn) onSlide.push(ln);
      // on-slide の「…」（入れ子の『』はそのまま）
      if (inOn) for (const qm of ln.matchAll(/「([^」]*)」/g)) quoted.add(qm[1]);
    }
    out[id] = { title, ch: ch != null ? Number(ch) : null, notes, sources: sources && sources !== '—' ? sources : null, quoted, layout, onSlide: onSlide.join('\n') };
  }
  return out;
}
const SPEC = parseSpec(SPEC_MD);
function notesOf(id) {
  const n = SPEC[id] && SPEC[id].notes;
  if (!n) throw new Error(`spec.md の ${id} に notes がありません`);
  return n;
}
function titleOf(id) {
  const t = SPEC[id] && SPEC[id].title;
  if (!t) throw new Error(`spec.md の ${id} に title がありません`);
  return t;
}
/** spec の文言そのまま（照合のみ） */
function q(id, str) {
  if (!SPEC[id].quoted.has(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
  return str;
}
/** on-slide 行の「」の外にある文言（例: グラフの（Hawn 2013））の照合 */
function qOn(id, str) {
  if (!SPEC[id].onSlide.includes(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
  return str;
}
/** layout 行の文言（図の見出しに使う）の照合 */
function qLayout(id, str) {
  if (!SPEC[id].layout.includes(str)) throw new Error(`${id}: 「${str}」は spec の layout にありません`);
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
/** 行配列のうち、指定した行に部分装飾を付ける。marks: {行番号: [[語, 色], ...]} */
function markLines(lines, marks) {
  return lines.map((l, i) => (marks[i] ? mark(l, ...marks[i]) : l));
}
const runText = l => (Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(l));
const toRuns = l => (Array.isArray(l) ? l : [l]).map(r => (typeof r === 'string' ? { text: r } : r));

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める（4-7 と同じ手順）
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 文言はそのままで、文献の途中の語の切れ目で改行する（区切り > 著者「et al.」の後 > 誌名と年の間 > その他の空白）。
//     数値と単位（「2022 年」など）の間では切らない。
//  2) 入らなければ区切りの前後の空白を詰める（' / ' → '/'）。
//  3) それでも入らなければ（）内の補足を後ろの文献から外す。
//  4) それでも入らないときだけ、JCS 2022 の長い表題を最小限短くする（文献は落とさない。DEVIATIONS に記録）。
// ---------------------------------------------------------------------------
const SRC = LAYOUT.source;
const DEVIATIONS = [];
const NOTES_LOG = [];
const JCS_FULL = '日本循環器学会 2022 年改訂版 非心臓手術における合併心疾患の評価と管理に関するガイドライン';
const JCS_SHORT = [
  '日本循環器学会 2022 年改訂版 非心臓手術における合併心疾患ガイドライン',   // 「の評価と管理に関する」を外す（4-7 と同じ）
  '日本循環器学会 2022 年改訂版 非心臓手術ガイドライン',
];
function libLines(items) {
  const lines = [];
  for (const it of items) {
    const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
    if (cur && fits(cur, SRC.size, SRC.w)) lines[lines.length - 1] = cur;
    else lines.push(it);
  }
  return lines;
}
/** 項目 items を sep でつないだ文字列を 2 行に分ける。区切りの位置（文献の境目）か、文献の中の空白で切る */
function split2(items, sep) {
  const str = items.join(sep);
  const bounds = new Set();
  let pos = 0;
  items.forEach((it, i) => { pos += it.length; if (i < items.length - 1) { bounds.add(pos); pos += sep.length; } });
  const inSep = new Set();
  for (const b of bounds) for (let k = b; k < b + sep.length; k++) inSep.add(k);
  const fitsLine = l => fits(l, SRC.size, SRC.w);
  const cands = [];
  for (let i = 1; i < str.length; i++) {
    let a, b, rank;
    if (bounds.has(i)) { a = str.slice(0, i); b = str.slice(i + sep.length); rank = 0; }
    else if (str[i] === ' ' && !inSep.has(i)) {
      a = str.slice(0, i); b = str.slice(i + 1);
      if (/ et$/.test(a)) continue;                   // 「et」と「al.」の間では切らない
      if (/\d$/.test(a) && /^[^\x00-\x7F]/.test(b)) continue;   // 数値と単位（2022 年 など）の間では切らない
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
function sourcesFor(id) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  const items = spec.split(' / ');
  if (libLines(items).length <= 2) return items;
  let r = split2(items, ' / ');
  if (r) { NOTES_LOG.push(`${id} 出典: 文献の途中（「${r[0].slice(-14)}」の後）で改行（文言はそのまま）`); return r; }
  r = split2(items, '/');
  if (r) { DEVIATIONS.push(`${id} 出典: 区切りの前後の空白を詰めた（' / ' → '/'）`); return r; }
  // 3) （）内の補足を後ろの文献から 1 つずつ外す（stages[n] = n 個外した状態）
  const stages = [{ cut: items.slice(), removed: [] }];
  for (let k = items.length - 1; k >= 0; k--) {
    const prev = stages[stages.length - 1];
    const m = prev.cut[k].match(/（[^）]*）$/);
    if (!m) continue;
    const c = prev.cut.slice();
    c[k] = c[k].slice(0, -m[0].length);
    stages.push({ cut: c, removed: [...prev.removed, m[0]] });
  }
  for (const st of stages.slice(1)) {
    r = split2(st.cut, '/');
    if (r) {
      DEVIATIONS.push(`${id} 出典: 区切りの前後の空白を詰め（' / ' → '/'）、（）内の補足 ${st.removed.join('・')} を外した`);
      return r;
    }
  }
  // 4) 最後の手段: JCS 2022 の長い表題を最小限短くする。（）内の補足は外す数が少ない順に試し、残せるものは残す
  for (const short of JCS_SHORT) {
    for (const st of stages) {
      const k = st.cut.indexOf(JCS_FULL);
      if (k < 0) continue;
      const c2 = st.cut.slice();
      c2[k] = short;
      r = split2(c2, '/');
      if (r) {
        const rm = st.removed.length ? `、（）内の補足 ${st.removed.join('・')} を外し` : '';
        DEVIATIONS.push(`${id} 出典: 区切りの前後の空白を詰め（' / ' → '/'）${rm}、` +
          `「${JCS_FULL}」→「${short}」（それでも 16pt 2 行に入らないため。文献は落としていない）`);
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
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;
const up3 = v => Math.ceil(v * 1000 - 1e-6) / 1000;

/** 本文の下端: 出典が 2 行（枠 y6.5）なら 6.465in、1 行（y6.85）なら 6.75in */
function botOf(s) {
  return s.hasSources ? r3(Math.min(LAYOUT.body.bottom, s.sourceTop - 0.035)) : LAYOUT.body.bottom;
}
function assertBottom(s, id, what, bottom) {
  const lim = botOf(s);
  if (bottom > lim + 1e-6) throw new Error(`${id} ${what}の下端 ${r3(bottom)}in > ${lim}in`);
}

// ---------------------------------------------------------------------------
// 行送りを固定した項目リスト（4-9 と同じ作り）。枠線・塗りなしの表（2 列: 行頭記号 | 本文、1 項目 = 1 行）で組む。
//  テキスト枠の行間（倍率）は描画系で 1 行の高さが変わる（実機メイリオ 1.5 em、LibreOffice の代替フォント約 1.2 em）ため、
//  2〜3 行に割れる項目が多いカードは、3-11 の tightTable と同じく行送りを固定値（既定 1.2 em）にできる表のセルで組む。
//  2 行目以降は本文の列に入るので、行頭記号の右にぶら下がる。項目の間は行の下余白（gap pt）。
//  幅は decklib と同じ推定器で、高さは行数 × 行送り＋下余白で検査する。重なり検査用の登録も行う。
//  （pptxgenjs はセル余白の先頭値が 1 未満だとインチとみなすので、余白はインチで渡す）
//  item = 行 | {lines:[行...], color, bold}（行 = 文字列 | run 配列）
// ---------------------------------------------------------------------------
function normItems(items) {
  return items.map(it => (it && typeof it === 'object' && !Array.isArray(it) && it.lines ? it : { lines: [it] }));
}
function itemTableH(items, o) {
  const size = o.size || 22, lsEm = o.lsEm || 1.2, gap = o.gap != null ? o.gap : 6;
  const its = normItems(items);
  const n = its.reduce((a, it) => a + it.lines.length, 0);
  return n * size * lsEm / 72 + (its.length - 1) * gap / 72;
}
/** 項目群の最長行の幅（行頭記号ぶんを含む） */
function itemsNeedW(items, o) {
  const size = o.size || 22;
  const ind = o.marker === false ? 0 : (o.indent != null ? o.indent : 0.28);
  return Math.max(...normItems(items).flatMap(it => it.lines).map(l => measure(runText(l), size))) + ind;
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

// ---------------------------------------------------------------------------
// 見出し付きカード（角丸の灰色地 + 太字の見出し + 行頭記号つき項目 22pt、行送り固定）
//  title: 文字列 | 行配列（2 行の見出しは行送り固定の表で組む）
// ---------------------------------------------------------------------------
const CARD = { padX: 0.14, padY: 0.1, titleGap: 0.04, titleSize: 24, size: 22, radius: 0.08, indent: 0.28, lsEm: 1.2 };
/** padT / padB（上・下の余白）を省略すると padY */
const cardOpt = o => {
  const c = { ...CARD, ...o };
  return { ...c, padT: o.padT != null ? o.padT : c.padY, padB: o.padB != null ? o.padB : c.padY };
};
function titleLines(o) { return Array.isArray(o.title) ? o.title : [o.title]; }
function cardTitleH(o) {
  const c = cardOpt(o);
  return c.titleH != null ? c.titleH : titleLines(c).length * c.titleSize * 1.2 / 72;
}
function cardBodyH(o) {
  const c = cardOpt(o);
  return itemTableH(c.items, { size: c.size, lsEm: c.lsEm, gap: c.gap });
}
function listCardNeed(o) {
  const c = cardOpt(o);
  return c.padT + c.padB + cardTitleH(c) + c.titleGap + cardBodyH(c);
}
function listCardNeedW(o) {
  const c = cardOpt(o);
  const title = Math.max(...titleLines(c).map(t => measure(t, c.titleSize)));
  return Math.max(title, itemsNeedW(c.items, { size: c.size, indent: c.indent })) + 2 * c.padX;
}
function listCard(s, o) {
  const c = cardOpt(o);
  if (listCardNeed(c) > c.h + 1e-6) throw new Error(`p${s.page} カード「${titleLines(c).join('')}」: 高さ不足（必要 ${r3(listCardNeed(c))}in > ${c.h}in）`);
  s.box({ x: r3(c.x), y: r3(c.y), w: r3(c.w), h: r3(c.h), fill: C.bg, radius: c.radius });
  const ix = r3(c.x + c.padX), iw = r3(c.w - 2 * c.padX);
  const th = cardTitleH(c);
  const tl = titleLines(c);
  // 見出しも行送り固定（1.2 em）の表で組む（テキスト枠だと lint 推定の高さが 0.44in 未満の枠に入らない）
  itemTable(s, tl, { x: ix, y: r3(c.y + c.padT), w: iw, h: th, size: c.titleSize, gap: 0, marker: false, bold: true,
    color: c.titleColor || C.dark, align: c.titleAlign || 'left' });
  const by = r3(c.y + c.padT + th + c.titleGap);
  const bh = r3(c.y + c.h - c.padB - by);
  return itemTable(s, c.items, { x: ix, y: by, w: iw, h: bh, size: c.size, lsEm: c.lsEm, gap: c.gap, indent: c.indent });
}
/**
 * 2 列カード。幅は各カードの最長行に合わせ、余りを等分。高さは h（省略時は高い方に合わせる）。
 * 低い方は余った高さを項目の間に配る。戻り値: カードの下端
 */
function twoCards(s, id, L, R, y, h, G = 0.25) {
  const wl = listCardNeedW(L), wr = listCardNeedW(R);
  const slack = BW - G - wl - wr;
  if (slack < 0) throw new Error(`${id}: 2 列カードの幅が ${r3(-slack)}in 足りません`);
  const wL = r3(wl + slack / 2), wR = r3(BW - G - wL);
  const ch = h != null ? h : r2(Math.max(listCardNeed(L), listCardNeed(R)) + 0.01);
  const spread = c => {
    const n = normItems(c.items).length;
    const free = ch - listCardNeed(c);
    return n > 1 && free > 0.01 ? { ...c, gap: (c.gap || 0) + Math.floor(free / (n - 1) * 72 * 10) / 10 } : c;
  };
  const Ls = spread(L), Rs = spread(R);
  listCard(s, { ...Ls, x: X0, y, w: wL, h: ch });
  listCard(s, { ...Rs, x: r3(XR - wR), y, w: wR, h: ch });
  console.log(`  ${id}: カード幅 ${wL} / ${wR}（余り ${r3(slack)}）、h ${ch}（必要 ${r3(listCardNeed(L))} / ${r3(listCardNeed(R))}、` +
    `項目間 ${Ls.gap} / ${Rs.gap}pt）`);
  return r3(y + ch);
}

/** 強調ボックス（白地＋ティールの枠・太字・中央揃え）。lines は 1〜2 行（各行 文字列 | run 配列）。行送り固定の表で組む */
const emphTextH = o => itemTableH(o.lines, { size: o.size || 24, lsEm: o.lsEm, gap: 0 });
const emphNeed = o => emphTextH(o) + 2 * (o.margin != null ? o.margin : 0.08);
function emphBox(s, o) {
  const m = o.margin != null ? o.margin : 0.08;
  if (emphNeed(o) > o.h + 1e-6) throw new Error(`p${s.page} 強調ボックス: 高さ不足（必要 ${r3(emphNeed(o))}in > ${o.h}in）`);
  s.box({ x: r3(o.x), y: r3(o.y), w: r3(o.w), h: r3(o.h), fill: C.white, radius: 0.08, line: { color: o.lineColor || C.teal, width: 2.5 } });
  const th = emphTextH(o);
  return itemTable(s, o.lines, {
    x: r3(o.x + m), y: r3(o.y + (o.h - th) / 2), w: r3(o.w - 2 * m), size: o.size || 24, lsEm: o.lsEm, gap: 0,
    marker: false, align: 'center', bold: true, color: o.color || C.dark,
  });
}

/** spec の見出し行（ch=・title=）と一致するスライドを作る */
function specSlide(deck, id, o = {}) {
  const sp = SPEC[id];
  if (sp.ch == null) throw new Error(`${id}: ch がありません`);
  return deck.slide({ title: titleOf(id), chapter: sp.ch, notes: notesOf(id), sources: sourcesFor(id), ...o });
}

/** 箇条書き（1 項目 1 行）の文字サイズ: sizes のうち全項目が safety in の余裕を残して 1 行に入る最大 */
function bulletSize(items, sizes, w = 12.2, ind = 0.32, safety = 0.1) {
  for (const sz of sizes) {
    if (items.every(it => measure(runText(it), sz) + ind + safety <= w)) return sz;
  }
  throw new Error(`箇条書き「${runText(items[0])}」…が ${sizes[sizes.length - 1]}pt でも 1 行に入りません`);
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['心不全', 'AS', 'ステント', 'CIED・\nAF', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-8 心疾患別の注意点',
  });

  // ===== S1 表紙 =====
  {
    const subtitle = '4-8 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: titleOf('S1'), subtitle, meta, notes: notesOf('S1') });
  }

  // ===== S2 心不全：代償か、非代償か（2 列カード + 強調ボックス） =====
  {
    const id = 'S2';
    const s = specSlide(deck, id);
    const BOT = botOf(s);
    const cardBase = { padX: 0.1, padT: 0.09, padB: 0.03, titleH: 0.4, titleGap: 0.02, indent: 0.24, gap: 0 };
    const L = {
      ...cardBase, title: q(id, 'どれだけ危ないか'), titleColor: C.vermilion,
      items: [
        { lines: brk(id, '心不全は冠動脈疾患より危険（死亡 HR 1.63 対 1.08）', '（死亡') },
        { lines: brk(id, '90 日死亡 5.5%（心不全なし 1.2%）、症状があれば 10%', '（心不全', '症状') },
        { lines: brk(id, 'EF が保たれていても（HFpEF）リスクは上がる', 'リスク') },
        { lines: markLines(brk(id, '非代償（起座呼吸、浮腫、体重増加）は延期', '体重'), { 1: [['延期', C.vermilion]] }) },
      ],
    };
    const R = {
      ...cardBase, title: q(id, '術前にすること'), titleColor: C.blue,
      items: [
        { lines: brk(id, '新しい息切れ・所見があれば心エコー（I）。安定なら不要（III）', '心エコー') },
        { lines: brk(id, 'BNP/NT-proBNP でリスクを補う（2a）', 'リスク') },
        { lines: brk(id, 'β遮断薬・RAAS 阻害薬（HFrEF）・スタチンは継続', '（HFrEF）') },
        { lines: brk(id, 'SGLT2 阻害薬は 3 日前から休薬（→ 4-13）', '（→') },
      ],
    };
    const E = {
      lines: [mark(q(id, '『心不全の既往』ではなく『今、代償されているか』を問診と診察で確かめる'), ['『今、代償されているか』', C.blue])],
      size: 22, margin: 0.04,
    };
    const eh = up3(emphNeed(E));
    const ey = r3(BOT - eh);
    const ch = r3(ey - GAP - TOP);
    twoCards(s, id, L, R, TOP, ch, 0.24);
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(s, id, '強調ボックス', ey + eh);
    console.log(`  ${id}: カード ${TOP}〜${r3(TOP + ch)}、強調 ${ey}〜${r3(ey + eh)}`);
  }

  // ===== S3 AS：重症の定義と危険（2 列カード + 強調ボックス） =====
  {
    const id = 'S3';
    const s = specSlide(deck, id);
    const BOT = botOf(s);
    const cardBase = { padX: 0.14, padY: 0.08, titleH: 0.42, titleGap: 0.04, indent: 0.26, gap: 4 };
    const L = {
      ...cardBase, title: q(id, '重症 AS の定義'), titleColor: C.blue,
      items: [
        q(id, '最高血流速度 4.0 m/s 以上'),
        q(id, '平均圧較差 40 mmHg 以上'),
        { lines: brk(id, '弁口面積 1.0 cm² 未満（係数 0.6 cm²/m² 未満）', '（係数') },
        { lines: markLines(brk(id, '超重症：5.0 m/s 以上、60 mmHg 以上、0.6 cm² 未満', '60 mmHg'), { 0: [['超重症：', C.vermilion]] }) },
      ],
    };
    const R = {
      ...cardBase, title: q(id, '非心臓手術のリスク'), titleColor: C.vermilion,
      items: [
        { lines: brk(id, '中等度以上の AS：30 日死亡 2.1%（対照 1.0%）、心筋梗塞 3.0%', '（対照') },
        { lines: brk(id, '国内：未治療の重症 AS で 30 日死亡 4.3%、AVR 後は 0%', '30 日') },
        q(id, '症状があれば 7.2%、なくても 3.1%'),
        { lines: markLines(brk(id, '術後の心不全が多く、緊急手術が最も危ない', '緊急'), { 1: [['緊急手術が最も危ない', C.vermilion]] }) },
      ],
    };
    const E = {
      lines: markLines(brk(id, '収縮期雑音を聴いたら心エコー。症候性の重症 AS、EF 50% 未満の重症 AS は手術前にハートチームへ', '症候性'),
        { 1: [['ハートチームへ', C.vermilion]] }),
      size: 24, margin: 0.08,
    };
    const eh = up3(emphNeed(E));
    const ey = r3(BOT - eh);
    const ch = r3(ey - GAP - 0.06 - TOP);
    twoCards(s, id, L, R, TOP, ch);
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(s, id, '強調ボックス', ey + eh);
    console.log(`  ${id}: カード ${TOP}〜${r3(TOP + ch)}、強調 ${ey}〜${r3(ey + eh)}`);
  }

  // ===== S4 AS の麻酔：4 つの原則（4 列カード + 強調ボックス） =====
  {
    const id = 'S4';
    const s = specSlide(deck, id);
    const BOT = botOf(s);
    const G = 0.25, n = 4;
    const cw = r3((BW - (n - 1) * G) / n);
    const cardBase = { padX: 0.1, padY: 0.14, titleGap: 0.16, indent: 0.24, gap: 14, titleColor: C.blue };
    const cards = [
      { title: brk(id, '前負荷を保つ', '保つ'),
        items: [q(id, '脱水を避ける'), { lines: brk(id, '急な静脈拡張を避ける', '避ける') }] },
      { title: brk(id, '洞調律・徐脈寄り', '徐脈'),
        items: [q(id, '頻脈を避ける'), { lines: brk(id, '心房細動は早く治療', '早く') }] },
      { title: brk(id, '体血管抵抗を保つ', '保つ'),
        items: [{ lines: brk(id, '低血圧にはフェニレフリン', 'フェニレフリン') }, { lines: brk(id, 'エフェドリンは頻脈に注意', '頻脈') }] },
      { title: brk(id, '低血圧を放置しない', '放置'),
        items: [{ lines: brk(id, '肥大心筋の冠灌流が落ちる', '冠灌流') }, { lines: brk(id, '動脈ラインで早く気づく', '早く') }] },
    ].map(c => ({ ...cardBase, ...c }));
    const E = {
      lines: markLines(brk(id, '単回大量の脊髄くも膜下麻酔は避ける。低用量・持続硬膜外の分割投与なら選択肢（根拠は観察研究のみ）', '低用量'),
        { 0: [['単回大量の脊髄くも膜下麻酔は避ける', C.vermilion]] }),
      size: 24, margin: 0.1,
    };
    const need = Math.max(...cards.map(listCardNeed));
    const ch = r2(need + 0.12);
    const eh = up3(emphNeed(E));
    const sep = 0.45;
    const total = ch + sep + eh;
    const y0 = r2(TOP + Math.max(0, (BOT - TOP - total) / 2));
    cards.forEach((c, i) => {
      const x = r3(X0 + i * (cw + G));
      if (listCardNeedW(c) > cw + 1e-6) throw new Error(`${id}: カード「${titleLines(c).join('')}」の幅 ${r3(listCardNeedW(c))} > ${cw}`);
      listCard(s, { ...c, x, y: y0, w: cw, h: ch });
    });
    const ey = r3(y0 + ch + sep);
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(s, id, '強調ボックス', ey + eh);
    console.log(`  ${id}: カード w ${cw} h ${ch}（必要 ${r3(need)}）y ${y0}、強調 ${ey}〜${r3(ey + eh)}`);
  }

  // ===== S5 ステント後：いつ手術か（左: 横棒グラフ、右: 箇条書き 4 項目） =====
  {
    const id = 'S5';
    const s = specSlide(deck, id);
    const BOT = botOf(s);
    // 右の箇条書き（行送り固定の表）
    const items = [
      { lines: brk(id, 'AHA/ACC 2024：P2Y12 中止なら ACS 12 か月、安定 6 か月、急ぐ手術は 3 か月', 'ACS', '急ぐ') },
      { lines: markLines(brk(id, 'どのステントも 30 日以内の待機手術はしない。バルーンのみは 14 日', '待機'), { 0: [['30 日以内', C.vermilion]] }) },
      { lines: brk(id, 'JCS 2022 は PCI 後 1〜1.5 か月以降の手術を CQ2 で検討', '1〜1.5') },
      { lines: markLines(brk(id, 'アスピリンは続ける。P2Y12 阻害薬を止めるかは PCI 施行医と', 'P2Y12'), { 0: [['アスピリンは続ける。', C.blue]] }) },
    ];
    const IND = 0.28, RGAP = 16;
    const rw = r2(itemsNeedW(items, { size: 22, indent: IND }) + 0.06);
    const rx = r3(XR - rw);
    const rh = itemTableH(items, { size: 22, gap: RGAP });

    // 左: 横棒グラフ（30 日 MACE、Hawn 2013）
    const bars = [
      ['6 週未満', '11.6%', 11.6, C.vermilion],
      ['6 週〜6 か月', '6.4%', 6.4, C.blue],
      ['6〜12 か月', '4.2%', 4.2, C.blue],
      ['12〜24 か月', '3.5%', 3.5, C.blue],
    ];
    bars.forEach(([lab, val]) => q(id, `${lab} ${val}`));
    const head = [qLayout(id, '留置からの期間別 30 日 MACE'), qOn(id, '（Hawn 2013）')];
    const cR = r3(rx - 0.4);                  // グラフの右端（右の箇条書きと 0.4in あける）
    const labW = r2(Math.max(...bars.map(b => measure(b[0], 22))) + 0.08);
    const valW = r2(Math.max(...bars.map(b => measure(b[1], 22))) + 0.08);
    const LG = 0.18;                          // ラベル↔軸、棒↔数値の間
    const axisX = r3(X0 + labW + LG);
    const barMax = r3(cR - valW - LG - axisX);
    const barH = 0.5, pitch = 0.82, lineH = 22 * 1.5 / 72;
    const headH = heightFor(head.length, 22, 0);
    const headGap = 0.3;
    const chartH = headH + headGap + (bars.length - 1) * pitch + barH;
    const blockH = Math.max(chartH, rh);
    const y0 = r2(TOP + Math.max(0, (BOT - TOP - blockH) / 2));
    const cy = r2(y0 + (blockH - chartH) / 2);
    s.text(head, { x: X0, y: cy, w: r3(cR - X0), h: headH, size: 22, bold: false, color: C.dark });
    const by0 = r3(cy + headH + headGap);
    s.vline({ x: axisX, y: r3(by0 - 0.12), h: r3((bars.length - 1) * pitch + barH + 0.24), color: C.gray, width: 1.5 });
    bars.forEach(([lab, val, v, color], i) => {
      const by = r3(by0 + i * pitch);
      const len = r3(barMax * v / bars[0][2]);
      const ty = r3(by + (barH - lineH) / 2);
      s.text(lab, { x: X0, y: ty, w: labW, h: r3(lineH), size: 22, align: 'right' });
      s.box({ x: axisX, y: by, w: len, h: barH, fill: color });
      s.text(val, { x: r3(axisX + len + LG), y: ty, w: valW, h: r3(lineH), size: 22, bold: true, color });
    });
    const ry = r2(y0 + (blockH - rh) / 2);
    itemTable(s, items, { x: rx, y: ry, w: rw, size: 22, gap: RGAP, indent: IND });
    assertBottom(s, id, 'グラフ', by0 + (bars.length - 1) * pitch + barH + 0.12);
    assertBottom(s, id, '右の箇条書き', ry + rh);
    console.log(`  ${id}: グラフ x ${X0}〜${cR}（ラベル ${labW}、棒 最大 ${barMax}）y ${cy}〜${r2(cy + chartH)}、右 x ${rx} w ${rw} y ${ry}〜${r2(ry + rh)}`);
  }

  // ===== S6 CIED と心房細動（2 列カード + 強調ボックス） =====
  {
    const id = 'S6';
    const s = specSlide(deck, id);
    const BOT = botOf(s);
    const cardBase = { padX: 0.1, padT: 0.09, padB: 0.03, titleH: 0.4, titleGap: 0.02, indent: 0.24, gap: 0 };
    const L = {
      ...cardBase, title: q(id, 'CIED（手帳で確かめる）'), titleColor: C.blue,
      items: [
        q(id, '機種、植込みの理由、ペーシング依存か'),
        { lines: brk(id, '最終点検：ペースメーカー 12 か月、ICD 6 か月以内', 'ICD') },
        q(id, '依存＋臍より上の手術：非同期に再設定か磁石'),
        { lines: brk(id, 'ICD に磁石：頻拍治療が止まる。ペーシングは多くの機種で不変', 'ペーシング') },
        { lines: brk(id, 'バイポーラ優先、短い通電、対極板の位置、除細動パッド', '対極板') },
      ],
    };
    const R = {
      ...cardBase, title: q(id, '心房細動'), titleColor: C.vermilion,
      items: [
        { lines: brk(id, '心拍数 100 超は術前に薬で整える', '術前') },
        { lines: brk(id, '新規の心房細動は術後に外来フォローと血栓リスク評価（I）', '外来', '血栓') },
        q(id, '抗凝固薬の休薬は 4-12'),
        { lines: brk(id, 'β遮断薬・Ca 拮抗薬は当日も続ける', '当日') },
      ],
    };
    const E = {
      lines: markLines(brk(id, '設定を変えたら、監視下を離れる前に戻す。磁石の置き場所を全員が知っておく', '磁石'),
        { 0: [['監視下を離れる前に戻す', C.vermilion]] }),
      size: 22, margin: 0.04,
    };
    const eh = up3(emphNeed(E));
    const ey = r3(BOT - eh);
    const ch = r3(ey - GAP - TOP);
    twoCards(s, id, L, R, TOP, ch, 0.24);
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(s, id, '強調ボックス', ey + eh);
    console.log(`  ${id}: カード ${TOP}〜${r3(TOP + ch)}、強調 ${ey}〜${r3(ey + eh)}`);
  }

  // ===== S7 肺高血圧：相談して、送る（箇条書き 5 行 + 強調ボックス） =====
  {
    const id = 'S7';
    const s = specSlide(deck, id);
    const BOT = botOf(s);
    const items = [
      q(id, '定義：安静時の平均肺動脈圧 20 mmHg 超（2022 ESC/ERS）'),
      mark(q(id, '非心臓手術の早期死亡 7%。肺動脈性肺高血圧の緊急手術は 15%'), ['緊急手術は 15%', C.red]),
      q(id, '合併症を増やす：右房圧上昇、6 分間歩行 399 m 未満、昇圧薬、緊急'),
      q(id, '肺高血圧の治療薬は続ける（I）。心エコー・右心カテで層別化'),
      mark(q(id, '避ける：低酸素、高 CO2、アシドーシス、低体温、低血圧'), ['避ける：', C.vermilion]),
    ];
    const size = bulletSize(items, [24, 22]);
    const E = {
      lines: markLines(brk(id, '待機手術は肺高血圧の診療経験がある施設へ。緊急なら専門医に電話してから', '緊急'),
        { 1: [['専門医に電話してから', C.vermilion]] }),
      size: 24, margin: 0.1,
    };
    const eh = up3(emphNeed(E));
    const gap = 12;
    const bh = heightFor(items.length, size, gap);
    const sep = r2(Math.max(GAP + 0.1, (BOT - TOP - bh - eh) / 2));
    const by0 = r2(TOP + Math.max(0, (BOT - TOP - bh - sep - eh) / 2));
    s.bullets(items, { y: by0, h: bh, size, gap });
    const ey = r3(by0 + bh + sep);
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(s, id, '強調ボックス', ey + eh);
    console.log(`  ${id}: 箇条書き ${size}pt ${by0}〜${r2(by0 + bh)}、強調 ${ey}〜${r3(ey + eh)}`);
  }

  // ===== S8 まとめ（箇条書き 5 行） =====
  {
    const id = 'S8';
    const s = specSlide(deck, id);
    const BOT = botOf(s);
    const items = [
      mark(q(id, '心不全：非代償は延期。β遮断薬は続け、SGLT2 は 3 日前に休薬'), ['非代償は延期', C.vermilion]),
      mark(q(id, 'AS：雑音→心エコー。症候性・EF 50% 未満の重症はハートチームへ'), ['ハートチームへ', C.blue]),
      mark(q(id, 'AS の麻酔：前負荷・洞調律・血管抵抗を保ち、低血圧を放置しない'), ['低血圧を放置しない', C.vermilion]),
      mark(q(id, 'ステント後：30 日以内は不可。P2Y12 中止なら安定 6・ACS 12 か月'), ['30 日以内は不可', C.vermilion]),
      mark(q(id, 'CIED は手帳で確認し磁石の意味を知る。肺高血圧は専門施設へ'), ['専門施設へ', C.blue]),
    ];
    const size = bulletSize(items, [26, 24, 22]);
    const gap = 24;
    const hh = heightFor(items.length, size, gap);
    if (hh > BOT - TOP + 1e-6) throw new Error(`${id} 箇条書きの高さ ${hh} > ${r2(BOT - TOP)}`);
    s.bullets(items, { y: TOP, h: r2(BOT - 0.02 - TOP), size, gap, valign: 'middle' });
    console.log(`  ${id}: 箇条書き ${size}pt（必要 ${hh}）`);
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (NOTES_LOG.length) console.log('[notes]\n  ' + NOTES_LOG.join('\n  '));
  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  if (res.render) console.log('[render] ' + JSON.stringify(res.render));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.stack || e.message || e); process.exit(1); });
