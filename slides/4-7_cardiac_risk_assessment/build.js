'use strict';
/**
 * 4-7 心血管リスク評価の手順
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。q()/brk() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 図はすべてネイティブ図形（箱・矢印・表）。図のラベルは図の外に置く。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fits, fitInfo, heightFor, measure } = require('../_lib/decklib');

const NAME = '4-7_cardiac_risk_assessment';
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
    out[id] = { title, ch: ch != null ? Number(ch) : null, notes, sources: sources && sources !== '—' ? sources : null, quoted };
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
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 文言はそのままで、文献の途中の語の切れ目で改行する（区切り > 著者「et al.」の後 > 誌名と年の間 > その他の空白）。
//  2) 入らなければ区切りの前後の空白を詰める（' / ' → '/'）。
//  3) それでも入らなければ（）内の補足を後ろの文献から外す。外した段階ごとに、まず区切り ' / ' のままで試し、
//     入らないときだけ空白を詰める（補足を外して入るようになったら、空白は詰めない）。
//  4) それでも入らないときだけ、JCS 2022 の長い表題を最小限短くする（文献は落とさない。DEVIATIONS に記録）。
//  S5・S6・S8・S9 は、どの 2 文献を 1 行に並べても 16pt で 1 行（46.6 字）を超えるため、
//  文献の区切りだけでは 2 行に入らない（3 行になりライブラリが Error）。文献の途中での改行は避けられない。
// ---------------------------------------------------------------------------
const SRC = LAYOUT.source;
const DEVIATIONS = [];
const NOTES_LOG = [];
const JCS_FULL = '日本循環器学会 2022 年改訂版 非心臓手術における合併心疾患の評価と管理に関するガイドライン';
const JCS_SHORT = [
  '日本循環器学会 2022 年改訂版 非心臓手術における合併心疾患ガイドライン',   // 「の評価と管理に関する」を外す
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
      if (a.endsWith('al.')) rank = 1;                 // 著者名（et al.）の後
      else if (/^\d/.test(b)) rank = 2;                // 誌名と年の間
      else rank = 3;
    } else continue;
    a = a.trim(); b = b.trim();
    if (a && b && fitsLine(a) && fitsLine(b)) cands.push({ a, b, rank });
  }
  if (!cands.length) return null;
  cands.sort((x, y) => x.rank - y.rank || y.a.length - x.a.length);
  return { lines: [cands[0].a, cands[0].b], rank: cands[0].rank };
}
function sourcesFor(id) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  const items = spec.split(' / ');
  if (libLines(items).length <= 2) return items;
  // （）内の補足を後ろの文献から 1 つずつ外した段階（stages[n] = n 個外した状態。stages[0] は外さない）
  const stages = [{ cut: items.slice(), removed: [] }];
  for (let k = items.length - 1; k >= 0; k--) {
    const prev = stages[stages.length - 1];
    const m = prev.cut[k].match(/（[^）]*）$/);
    if (!m) continue;
    const c = prev.cut.slice();
    c[k] = c[k].slice(0, -m[0].length);
    stages.push({ cut: c, removed: [...prev.removed, m[0]] });
  }
  // 表題は原文 → 短縮 1 → 短縮 2 の順。それぞれで（）を外す数の少ない順、区切りは ' / ' → '/' の順に試す
  for (const short of [null, ...JCS_SHORT]) {
    for (const st of stages) {
      let cut = st.cut;
      if (short) {
        const k = cut.indexOf(JCS_FULL);
        if (k < 0) continue;
        cut = cut.slice();
        cut[k] = short;
      }
      for (const sep of [' / ', '/']) {
        const r = split2(cut, sep);
        if (!r) continue;
        const how = r.rank === 0 ? '文献の区切りで改行' : `文献の途中（「${r.lines[0].slice(-16)}」の後）で改行`;
        const parts = [];
        if (sep === '/') parts.push(`区切りの前後の空白を詰めた（' / ' → '/'）`);
        if (st.removed.length) parts.push(`（）内の補足 ${st.removed.join('・')} を外した`);
        if (short) parts.push(`表題を短縮「${JCS_FULL}」→「${short}」`);
        if (parts.length) {
          const keep = sep === ' / ' ? `区切りは ' / ' のまま。` : '';
          DEVIATIONS.push(`${id} 出典: ${parts.join('。')}（16pt 2 行に入らないため。文献は落としていない。${keep}${how}）`);
        } else {
          NOTES_LOG.push(`${id} 出典: ${how}（文言はそのまま）`);
        }
        return r.lines;
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
const TOP_T = 1.82;              // 表の上端（金下線 8pt の下端 1.576in から 0.24in 以上。4-10・4-12 と同じ）
const BOT = 6.46;                // 出典 2 行（枠 y6.5）のときの本文下端（全スライド出典 2 行）
const ZONE_BOT = 6.44;           // 表の下端の上限（出典 2 行は枠 y6.5・下揃えで、メイリオ 16pt の文字の上端が約 6.68in。その 0.24in 上）
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
// 表のセル余白（pt: 上, 右, 下, 左）。行間を固定値（1.2 em）にすると字面が行の上に寄り、字の下に約 3pt の空きが残る
// （描画で実測: 上下とも 1pt の余白で、上の罫線まで 2.2pt・下の罫線まで 4.0pt）。上を厚く下を薄くして、
// 字と罫線の間を上下とも約 3.6pt（0.05in）にそろえる。
const TM = [2.6, 6, 0.4, 6];
const CM = 2.54;
const MEIRYO_EM = 1.5;           // 実機メイリオの 1 行高（em、行間 100%）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;

function assertBottom(id, what, bottom, limit = BOT) {
  if (bottom > limit + 1e-6) throw new Error(`${id} ${what}の下端 ${r3(bottom)}in > ${limit}in`);
}

// ---------------------------------------------------------------------------
// 行頭記号つきの項目群（1 つのテキスト枠。4-9 と同じ）
//  - 1 項目を 2 行に割ったときは、2 行目以降を記号の右にぶら下げる（見えない記号 U+3000 で字下げ）。
//  - ls: 行間（倍率。0.8 = 実機メイリオで約 1.2 em。3-11 の tightTable と同じ行送り）。
//  - gap: 項目の間の段落後の間隔（pt）。
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

/** 見出し付きカード（角丸の灰色地 + 24pt 太字の見出し + 行頭記号つき項目 22pt） */
const CARD = { padX: 0.14, padY: 0.1, titleH: 0.45, titleGap: 0.04, titleSize: 24, size: 22, radius: 0.08, indent: 0.28 };
function listCardNeed(o) {
  const padY = o.padY != null ? o.padY : CARD.padY;
  return 2 * padY + CARD.titleH + CARD.titleGap + itemTextH(o.items, { size: CARD.size, ls: o.ls, gap: o.gap });
}
/** カードに必要な幅（見出し 24pt・項目 22pt の最長行 + 行頭記号 + 左右の余白） */
function listCardNeedW(o) {
  const lines = normItems(o.items).flatMap(it => it.lines);
  const body = Math.max(...lines.map(l => measure(runText(l), CARD.size))) + CARD.indent;
  return Math.max(measure(o.title, CARD.titleSize), body) + 2 * CARD.padX;
}
function listCard(s, o) {
  const padY = o.padY != null ? o.padY : CARD.padY;
  if (listCardNeed(o) > o.h + 1e-6) throw new Error(`p${s.page} カード「${o.title}」: 高さ不足（必要 ${r3(listCardNeed(o))}in > ${o.h}in）`);
  s.box({ x: r3(o.x), y: r3(o.y), w: r3(o.w), h: r3(o.h), fill: C.bg, radius: CARD.radius });
  const ix = r3(o.x + CARD.padX), iw = r3(o.w - 2 * CARD.padX);
  s.text(o.title, { x: ix, y: r3(o.y + padY), w: iw, h: CARD.titleH, size: CARD.titleSize, bold: true, color: o.titleColor || C.dark });
  const by = r3(o.y + padY + CARD.titleH + CARD.titleGap);
  return itemText(s, o.items, { x: ix, y: by, w: iw, h: r3(o.y + o.h - padY - by), size: CARD.size, ls: o.ls, gap: o.gap, indent: CARD.indent });
}
/**
 * 2 列カード。幅は各カードの最長行に合わせ、余りを等分。高さは高い方に合わせる。
 * 項目の間隔（gap）と行間（ls）は左右で同じにし、低い方のカードは下に空きを残す
 * （余った高さを項目の間に配ると、左右で項目の並びのリズムがそろわないため）。戻り値: カードの下端
 */
function twoCards(s, id, L, R, y) {
  if (L.gap !== R.gap || L.ls !== R.ls) throw new Error(`${id}: 2 列カードの gap・ls は左右で同じにしてください`);
  const wl = listCardNeedW(L), wr = listCardNeedW(R);
  const slack = BW - 0.25 - wl - wr;
  if (slack < 0) throw new Error(`${id}: 2 列カードの幅が ${r3(-slack)}in 足りません`);
  const wL = r3(wl + slack / 2), wR = r3(BW - 0.25 - wL);
  const ch = r2(Math.max(listCardNeed(L), listCardNeed(R)) + 0.01);
  listCard(s, { ...L, x: X0, y, w: wL, h: ch });
  listCard(s, { ...R, x: r3(XR - wR), y, w: wR, h: ch });
  console.log(`  ${id}: カード幅 ${wL} / ${wR}（余り ${r3(slack)}）、h ${ch}（必要 ${r3(listCardNeed(L))} / ${r3(listCardNeed(R))}）、項目間 ${L.gap}pt・行間 ${L.ls}`);
  return r2(y + ch);
}

/** 強調ボックス（白地＋ティールの枠・太字）。lines は 1〜2 行（各行 文字列 | run 配列）。 */
function emphBox(s, o) {
  const m = o.margin != null ? o.margin : 0.08;
  s.box({ x: r3(o.x), y: r3(o.y), w: r3(o.w), h: r3(o.h), fill: C.white, radius: 0.08, line: { color: o.lineColor || C.teal, width: 2.5 } });
  return itemText(s, o.lines, {
    x: r3(o.x + m), y: r3(o.y + m), w: r3(o.w - 2 * m), h: r3(o.h - 2 * m), size: o.size || 24, ls: o.ls || 1, gap: 0,
    marker: false, align: 'center', valign: 'middle', bold: true, color: o.color || C.dark,
  });
}
const emphNeed = o => itemTextH(o.lines, { size: o.size || 24, ls: o.ls || 1, gap: 0 }) + 2 * (o.margin != null ? o.margin : 0.08);

/**
 * 行の多い表（S4・S6 用。3-11 の tightTable を写したもの）。decklib の s.table と同じ見た目
 * （見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、上下中央）で、行間だけを固定値 LS pt（= 1.2 em）に詰める。
 * s.table は行高を実機メイリオの 1.5 em で取るため、行の多い表が本文ゾーン（出典 2 行時 1.85〜6.46in）に入らない。
 * セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、重なり検査用の登録も同じく行う。
 * cell: 文字列（'\n' で改行）| {text, color, bold, align}
 * o.round: 行の高さの切り上げ単位（既定 0.01in。S4・S6 は 0.001in にして、余白を本文ゾーンの限度まで使う。4-10 と同じ）
 */
function tightTable(s, rows, o) {
  const size = o.size || 22, LS = o.lineSpacing || size * 1.2;
  const m = o.margin || [3, 6, 3, 6];   // pt: 上, 右, 下, 左
  // pptxgenjs は margin[0] が 1 未満だと 4 値すべてをインチとして扱うため、上の余白は 1pt 以上にする
  if (!(m[0] >= 1)) throw new Error(`p${s.page} 表: 上の余白 ${m[0]}pt は 1pt 以上にしてください`);
  const RND = o.round || 0.01;
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
    rowH.push(r3(Math.ceil((maxLines * LS / 72 + (m[0] + m[2]) / 72) / RND - 1e-6) * RND));
    return cells;
  });
  const h = rowH.reduce((a, b) => a + b, 0);
  s._bounds(`p${s.page} 表`, o.x, o.y, w, h);
  s.raw.addTable(out, {
    x: o.x, y: o.y, w, colW, rowH, fontFace: FONT, fontSize: size, margin: m,
    border: { type: 'solid', pt: 1, color: C.border }, valign: 'middle',
  });
  s._reg('box', o.x, o.y, w, h, 'table');
  return { x: o.x, y: o.y, w, h: r3(h), rowH };
}

/**
 * 水平の矢印（軸＋塗りの三角の矢じり。3-11 の flowArrow と同じ作り）。
 * pptxgenjs の矢じりは線幅に比例して小さいので、矢じりを自前の三角形（freeform）で描く。
 */
function flowArrow(s, x1, y, x2, o = {}) {
  const color = o.color || '404040', width = o.width || 3, hl = o.hl || 0.16, hw = o.hw || 0.24;
  const bx = x2 - hl;
  s.hline({ x: r3(x1), y: r3(y), w: r3(bx - x1 + 0.02), color, width });
  s.polyline([[x2, y], [bx, y - hw / 2], [bx, y + hw / 2]], { color, width: 1, closed: true, fill: color });
}

/**
 * 行頭記号（•）の色を本文色 262626 に固定する（buClr）。pptxgenjs は buClr を出せず、記号の色が行頭の run の色
 * （S7「例外：」の青、S8「93% は無症状」の橙など）になってしまうため、書き出した XML に足す。
 * deck.pptx.write を包み、decklib の後処理（postProcess）の前に適用する。buClr は buSzPct の直前（スキーマの順）。
 */
function fixBulletColor(deck) {
  const JSZip = require('jszip');
  const write = deck.pptx.write.bind(deck.pptx);
  deck.pptx.write = async opts => {
    const zip = await JSZip.loadAsync(await write(opts));
    let n = 0;
    for (const name of Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))) {
      const x = await zip.file(name).async('string');
      const y = x.replace(/<a:buSzPct val="100000"\/><a:buChar /g, m => { n++; return `<a:buClr><a:srgbClr val="${C.dark}"/></a:buClr>${m}`; });
      zip.file(name, y);
    }
    if (!n) throw new Error('fixBulletColor: 行頭記号が見つかりません');
    console.log(`  行頭記号の色を ${C.dark} に固定（${n} 段落）`);
    return zip.generateAsync({ type: 'nodebuffer' });
  };
}

/** spec の見出し行（ch=・title=）と一致するスライドを作る */
function specSlide(deck, id, o = {}) {
  const sp = SPEC[id];
  if (sp.ch == null) throw new Error(`${id}: ch がありません`);
  return deck.slide({ title: titleOf(id), chapter: sp.ch, notes: notesOf(id), sources: sourcesFor(id), ...o });
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['手順', 'RCRI', '運動\n耐容能', '検査', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-7 心血管リスク評価の手順',
  });
  fixBulletColor(deck);

  // ===== S1 表紙 =====
  {
    const subtitle = '4-7 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: titleOf('S1'), subtitle, meta, notes: notesOf('S1') });
  }

  // ===== S2 評価の流れ（横並びのフロー 5 箱 + 各箱の下に補足） =====
  {
    const id = 'S2';
    const s = specSlide(deck, id);
    // 箱は左右に 0.15in 内側へ寄せ、補足（箱より 0.3in 広い）が本文ゾーンからはみ出さないようにする
    const INSET = 0.15, G = 0.36, SUPX = 0.15;
    const bw = r3((BW - 2 * INSET - 4 * G) / 5);
    const bh = 1.3, size = 24;
    const boxes = [
      [q(id, '緊急か')],
      brk(id, '活動性の心疾患', '心疾患'),
      [q(id, '手術リスク')],
      brk(id, 'RCRI・運動耐容能', '運動'),
      [q(id, '追加検査')],
    ];
    const sups = [
      markLines(brk(id, '2 時間未満（AHA）なら検査せず手術へ', '（AHA）', '検査'), { 2: [['手術へ', C.blue]] }),
      markLines(brk(id, 'あれば延期して治療（→ S3）', '治療'), { 0: [['延期して', C.vermilion]] }),
      markLines(brk(id, '低リスク手術なら原則そのまま手術', 'なら原則', 'そのまま'), { 2: [['そのまま手術', C.blue]] }),
      markLines(brk(id, '低リスクなら手術へ', '手術へ'), { 1: [['手術へ', C.blue]] }),
      markLines(brk(id, '管理が変わるときだけ', 'ときだけ'), { 1: [['ときだけ', C.vermilion]] }),
    ];
    const supSize = 22, supGap = 0.3;
    const supH = Math.max(...sups.map(l => heightFor(l.length, supSize, 0)));
    const groupH = bh + supGap + supH;
    const by = r2(TOP + (BOT - TOP - groupH) * 0.42);
    boxes.forEach((lines, i) => {
      const bx = r3(X0 + INSET + i * (bw + G));
      s.box({ x: bx, y: by, w: bw, h: bh, fill: C.bg, radius: 0.1, line: { color: C.blue, width: 2.5 },
        text: lines, size, bold: true, color: C.dark });
      if (i < boxes.length - 1) flowArrow(s, bx + bw + 0.05, r3(by + bh / 2), bx + bw + G - 0.05);
      const sl = sups[i];
      s.text(sl, { x: r3(bx - SUPX), y: r2(by + bh + supGap), w: r3(bw + 2 * SUPX), h: heightFor(sl.length, supSize, 0),
        size: supSize, align: 'center' });
    });
    assertBottom(id, '補足', by + bh + supGap + supH);
    console.log(`  ${id}: 箱 w ${bw} h ${bh}（y ${by}〜${r2(by + bh)}）、補足 〜${r2(by + bh + supGap + supH)}`);
  }

  // ===== S3 延期すべき活動性の心疾患（2 列カード + 強調ボックス） =====
  {
    const id = 'S3';
    const s = specSlide(deck, id);
    // 2 行にまたがる項目の行どうしが詰まって見えないよう行間を 0.85（実機メイリオで 1.28 em）にする。
    // 強調ボックス（2 行）を本文ゾーンに残すため、カードの高さは変えず、項目の間を 3.5pt に詰めて吸収する。
    const L = {
      title: q(id, '見つけたら延期（待機手術）'), titleColor: C.vermilion, ls: 0.85, gap: 3.5,
      items: [
        q(id, '急性冠症候群、MI 後 60 日以内'),
        { lines: brk(id, '非代償性心不全（起座呼吸、浮腫の悪化）', '（起座') },
        { lines: brk(id, '重症不整脈（新しい AF 高心拍、VT、高度ブロック）', 'VT') },
        q(id, '症候性の重症弁膜症（失神、狭心痛）'),
      ],
    };
    const R = {
      title: q(id, '時期と検査を相談する因子'), titleColor: C.blue, ls: 0.85, gap: 3.5,
      items: [
        { lines: brk(id, '冠動脈ステント・CABG の既往（→ 4-8）', '（→') },
        q(id, '脳卒中・TIA から 3 か月以内'),
        q(id, '重症弁膜症・重症肺高血圧'),
        q(id, '先天性心疾患、植込みデバイス'),
        q(id, 'フレイル（→ 4-11）'),
      ],
    };
    const cb = twoCards(s, id, L, R, TOP);
    // 強調ボックスの文言は 22pt でも 1 行に入らない（推定 12.32in > 12.07in）ため、「＝」の後で 2 行に割る（24pt・行送り 1.35 em）
    const E = {
      lines: markLines(brk(id, '新しい症状（胸痛、息切れの悪化、失神）＝評価が終わるまで待機手術はしない', '評価が'), { 1: [['待機手術はしない', C.red]] }),
      size: 24, ls: 0.9,
    };
    const eh = r2(emphNeed(E) + 0.06);
    const ey = r2(cb + Math.max(GAP, (BOT - cb - eh) / 2));
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(id, '強調ボックス', ey + eh);
    console.log(`  ${id}: 強調 ${ey}〜${r2(ey + eh)}`);
  }

  // ===== S4 RCRI：6 つの因子（表 2 列 × 6 行＋見出し行 + 下に 1 行） =====
  {
    const id = 'S4';
    const s = specSlide(deck, id);
    const cell = (str, ...breaks) => brk(id, str, ...breaks).join('\n');
    const f = str => ({ text: q(id, str), bold: true });
    const rows = [
      [q(id, '因子（各 1 点）'), q(id, '中身')],
      [f('高リスク手術'), q(id, '腹腔内、胸腔内、鼠径靱帯より中枢の血管手術')],
      [f('虚血性心疾患'), cell('心筋梗塞、負荷試験陽性、虚血による胸痛、硝酸薬、異常 Q 波', '硝酸薬')],
      [f('心不全'), cell('既往、肺水腫、夜間発作性呼吸困難、両側ラ音、III 音、X 線の血管再分布', 'III 音')],
      [f('脳血管疾患'), q(id, '脳卒中または一過性脳虚血発作の既往')],
      [f('インスリン治療中の糖尿病'), q(id, '経口薬だけの糖尿病は数えない')],
      [f('クレアチニン 2.0 mg/dL 超'), q(id, '177 µmol/L 超')],
    ];
    // 列幅: spec の目安 4.2 / 8.0 では「インスリン治療中の糖尿病」（推定 4.11in）「クレアチニン 2.0 mg/dL 超」（推定 4.24in）が
    // 1 列目（使える幅 4.03in）に 1 行で入らず、表が 2 行ぶん伸びて下の 1 行が本文ゾーンに入らない。
    // 1 列目を 4.41in に広げて全因子を 1 行にする（2 列目 7.79in でも中身は 1〜2 行のまま。合計 12.2in）。
    // 表（9 行ぶん）は上端 TOP_T・セル余白 TM（字と罫線の間 上下とも約 0.05in）。
    const t = tightTable(s, rows, { x: X0, y: TOP_T, colW: [4.41, 7.79], margin: TM, round: 0.001 });
    // 下の 1 行（全角 41 字ぶんで 2 行に割れる）は 22pt。「（Lee 1999 検証群）。近年は 3.9〜15%」を独立した 2 行目の段落にし、
    // 表の下 0.24in から本文下端までの高さで、行間 0.8 のまま 2 段落の間（段落後）をできるだけ空ける（2 行が詰まって見えないように）。
    const line = brk(id, '主要心合併症：0 点 0.4%、1 点 0.9%、2 点 7%、3 点以上 11%（Lee 1999 検証群）。近年は 3.9〜15%', '（Lee');
    const items = markLines(line, { 1: [['近年は 3.9〜15%', C.vermilion]] }).map(l => ({ lines: [l] }));
    const size = 22, ls = 0.8;
    const ly = r3(t.y + t.h + GAP);
    const lh = r3(BOT - ly);
    const gap = Math.min(12, Math.floor((lh - itemTextH(items, { size, ls, gap: 0 })) * 72 * 10) / 10);
    if (gap < 3) throw new Error(`${id}: 下の 1 行の 2 行の間 ${gap}pt < 3pt（表が高すぎます）`);
    itemText(s, items, { size, ls, gap, marker: false, x: X0, y: ly, w: BW, h: lh });
    assertBottom(id, '表', t.y + t.h, ZONE_BOT);
    assertBottom(id, '下の 1 行', ly + lh);
    console.log(`  ${id}: 表 ${t.y}〜${r3(t.y + t.h)}（行高 ${t.rowH.join(', ')}）、下の 1 行 ${size}pt 行間 ${ls}＋段落間 ${gap}pt ${ly}〜${r3(ly + lh)}`);
  }

  // ===== S5 METs と DASI で測る（2 列カード + 強調ボックス） =====
  {
    const id = 'S5';
    const s = specSlide(deck, id);
    // 左右とも項目間 8pt・行間 0.8 でそろえる。高さは右（7 行）に合わせ、左（5 行）は下に空きを残す。
    const L = {
      title: q(id, '4 METs の目安'), titleColor: C.blue, ls: 0.8, gap: 8,
      items: [
        q(id, '階段で上の階まで休まず上がれる'),
        q(id, '坂道を歩いて上がれる'),
        q(id, '床掃除、庭仕事'),
        { lines: markLines(brk(id, '主観で聞くと 4 METs 未満の 8 割を見逃す', '8 割'), { 1: [['8 割を見逃す', C.vermilion]] }) },
      ],
    };
    const R = {
      title: q(id, 'DASI（12 項目、0〜58 点）'), titleColor: C.vermilion, ls: 0.8, gap: 8,
      items: [
        { lines: brk(id, 'AHA 2024：上昇リスク手術では DASI で評価（2a）', 'DASI で') },
        mark(q(id, '34 点以下＝不良（AHA）'), ['不良', C.vermilion]),
        { lines: brk(id, 'JCS：7 METs（DASI 34）以上は追加検査なしで手術を考慮（IIa）', '追加検査') },
        { lines: brk(id, '34 点未満は 1 点下がるごとに死亡・MI のオッズが 1.05 倍', '死亡') },
      ],
    };
    const cb = twoCards(s, id, L, R, TOP);
    const E = { lines: [mark(q(id, '『階段は大丈夫ですか』ではなく、具体的な活動で聞く。迷えば DASI を使う'), ['具体的な活動で聞く', C.blue])], size: 22 };
    const eh = r2(emphNeed(E) + 0.06);
    const ey = r2(cb + Math.max(GAP + 0.06, (BOT - cb - eh) / 2));
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(id, '強調ボックス', ey + eh);
    console.log(`  ${id}: 強調 ${ey}〜${r2(ey + eh)}`);
  }

  // ===== S6 どの検査を、誰に（表 3 列 × 4 行＋見出し行） =====
  {
    const id = 'S6';
    const s = specSlide(deck, id);
    const cell = (str, ...breaks) => brk(id, str, ...breaks).join('\n');
    const f = (str, ...breaks) => ({ text: cell(str, ...breaks), bold: true });
    const rows = [
      [q(id, '検査'), q(id, '出す相手'), q(id, '出さない・注意')],
      [f('BNP / NT-proBNP', 'NT-proBNP'),
        cell('上昇リスク手術で心疾患・65 歳以上など（AHA 2a）。JCS は 4 METs 未満・不明で（IIb）', '65 歳', '4 METs'),
        cell('BNP 92、NT-proBNP 200〜300 pg/mL が目安。腎機能で上がる', '200〜300', '腎機能')],
      [f('12 誘導心電図', '心電図'),
        cell('心疾患・症状（AHA 2a）。リスク因子＋中〜高リスク手術（ESC）', 'リスク因子', '中〜高'),
        cell('低リスク手術の無症状者には出さない（III）', '無症状者', '出さない')],
      [f('心エコー'),
        cell('新しい息切れ、心不全の所見、重い弁膜症の疑い（I）', '重い'),
        cell('無症状で安定していればルーチンにしない（3）', 'ルーチン')],
      [f('負荷試験'),
        cell('上昇リスク手術で運動耐容能が不良・不明、結果で管理が変わる（2b）', '運動耐容能', '結果で'),
        cell('低リスク手術・低リスク患者・機能良好には行わない（3）', '低リスク患者', '機能良好')],
    ];
    // 列幅: spec の目安 2.4 / 5.4 / 4.4 では、行 1 の 2 列目がどこで割っても 4 行になる
    // （「JCS は 4 METs 未満・不明で（IIb）」は推定 5.60in、「65 歳以上など（AHA 2a）。JCS は」は 5.30in で、
    // どちらも 5.4in 列の使える幅 5.23in を超える）。行 4 の 3 列目「機能良好には行わない（3）」（4.30in）も 4.4in 列に入らない。
    // 表の高さが最小（12 行ぶん）になる配分に見直した（合計 12.2in）。3 列目は「機能良好には行わない（3）」が入る 4.47in、
    // 1 列目は「NT-proBNP」（推定 1.73in）が入る 2.0in。2 列目 5.73in は、「JCS は」から始まる行（5.60in）が入る幅
    // （5.77in 以上）には届かないため、行 1 は「。JCS は」の後で割り、最長の行（5.30in）に 0.26in の余裕を残す。
    // 12 行（行送り 1.2 em）を上端 TOP_T・セル余白 TM で本文ゾーン（〜6.47in）に入れる。
    const t = tightTable(s, rows, { x: X0, y: TOP_T, colW: [2.0, 5.73, 4.47], margin: TM, round: 0.001 });
    assertBottom(id, '表', t.y + t.h, ZONE_BOT);
    console.log(`  ${id}: 表 ${t.y}〜${r3(t.y + t.h)}（行高 ${t.rowH.join(', ')}）`);
  }

  // ===== S7 血行再建とステント後（箇条書き 4 行 + 強調ボックス） =====
  {
    const id = 'S7';
    const s = specSlide(deck, id);
    const items = [
      mark(q(id, 'CARP 試験：血管手術前の予防的血行再建で死亡も MI も減らず'), ['死亡も MI も減らず', C.vermilion]),
      mark(q(id, '例外：急性冠症候群（I）、左主幹部 50% 以上の狭窄（2a）'), ['例外：', C.blue]),
      mark(q(id, '血行再建をすると手術が遅れる（CARP：中央値 54 日 vs 18 日）'), ['54 日 vs 18 日', C.blue]),
      mark(q(id, 'DAPT 中断は ACS 後 12 か月、安定 6 か月（AHA）。PCI 医と相談'), ['PCI 医と相談', C.vermilion]),
    ];
    // 4 行目が 26pt では行頭記号ぶんで 1 行に入らないため 24pt（1 バレット 1 行を優先）
    const size = 24, gap = 22, by0 = 1.95;
    const bh = heightFor(items.length, size, gap);
    s.bullets(items, { y: by0, h: bh, size, gap });
    const E = { lines: [mark(q(id, '適応は非手術時と同じ。負荷試験は『管理が変わるとき』だけ'), ['『管理が変わるとき』だけ', C.vermilion])], size: 26 };
    const eh = r2(emphNeed(E) + 0.12);
    const ey = r2(by0 + bh + Math.max(GAP + 0.06, (BOT - by0 - bh - eh) / 2));
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(id, '強調ボックス', ey + eh);
    console.log(`  ${id}: 箇条書き ${by0}〜${r2(by0 + bh)}、強調 ${ey}〜${r2(ey + eh)}`);
  }

  // ===== S8 術後の心筋傷害（MINS）（箇条書き 5 行 + 強調ボックス） =====
  {
    const id = 'S8';
    const s = specSlide(deck, id);
    const items = [
      mark(q(id, 'MINS：術中〜術後 30 日の心筋傷害。45 歳以上の入院手術の 8〜18%'), ['8〜18%', C.blue]),
      mark(q(id, '93% は無症状。胸痛がないから大丈夫、ではない'), ['93% は無症状', C.vermilion]),
      q(id, 'hs-TnT 20〜65 ng/L で 5 以上の上昇、または 65 以上（VISION）'),
      // 3 値は青・黒・赤（いずれも太字）。橙と赤は色覚によって近く見えるので並べない
      mark(q(id, '30 日死亡：20〜65 で 3%、65〜1,000 で 9%、1,000 以上で 30%'), ['3%', C.blue], ['9%', C.dark], ['30%', C.red]),
      q(id, 'ESC：65 歳以上・心疾患などの中〜高リスク手術で術前・術後（I）'),
    ];
    const size = 24, gap = 10, by0 = 1.9;
    const bh = heightFor(items.length, size, gap);
    s.bullets(items, { y: by0, h: bh, size, gap });
    // 強調ボックスの文言は 22pt でも 1 行に入らない（推定 12.32in > 12.07in）ため、「：」の後で 2 行に割る（24pt・行送り 1.5 em。下に余裕があるので行間は詰めない）
    const E = {
      lines: markLines(brk(id, '上がっていたら：可逆的な原因（出血、低血圧、敗血症）を直し、循環器内科へ', '可逆的'), { 1: [['可逆的な原因', C.vermilion]] }),
      size: 24, ls: 1,
    };
    const eh = r2(emphNeed(E) + 0.06);
    const ey = r2(by0 + bh + Math.max(GAP, (BOT - by0 - bh - eh) / 2));
    emphBox(s, { ...E, x: X0, y: ey, w: BW, h: eh });
    assertBottom(id, '強調ボックス', ey + eh);
    console.log(`  ${id}: 箇条書き ${by0}〜${r2(by0 + bh)}、強調 ${ey}〜${r2(ey + eh)}`);
  }

  // ===== S9 まとめ（箇条書き 5 行） =====
  {
    const id = 'S9';
    const s = specSlide(deck, id);
    const items = [
      mark(q(id, '緊急なら検査せず手術。活動性の心疾患なら延期して相談'), ['延期して相談', C.vermilion]),
      mark(q(id, '低リスク手術や RCRI 低リスクには検査を足さない'), ['検査を足さない', C.blue]),
      mark(q(id, '運動耐容能は主観で聞かない。具体的な活動と DASI'), ['主観で聞かない', C.vermilion]),
      mark(q(id, 'BNP・負荷試験は上昇リスク手術で、管理が変わるときだけ'), ['管理が変わるときだけ', C.blue]),
      mark(q(id, '血行再建は手術のためにしない。MINS の 9 割は無症状'), ['9 割は無症状', C.vermilion]),
    ];
    const size = 26, gap = 24;
    const hh = heightFor(items.length, size, gap);
    if (hh > BOT - TOP + 1e-6) throw new Error(`${id} 箇条書きの高さ ${hh} > ${r2(BOT - TOP)}`);
    s.bullets(items, { y: TOP, h: r2(BOT - 0.02 - TOP), size, gap, valign: 'middle' });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (NOTES_LOG.length) console.log('[notes]\n  ' + NOTES_LOG.join('\n  '));
  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
