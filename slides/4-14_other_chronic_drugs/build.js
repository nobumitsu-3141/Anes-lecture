'use strict';
/**
 * 4-14 常用薬③：その他
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk()/q() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。spec から形を変えたところはビルド時に deviations として一覧を出す。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら scratchpad の render/<NAME>）。
 *
 * 密度への対応（文言を変えずに収めるための組版上の工夫）:
 *  - カード・強調ボックスの本文は 22pt 以上のまま行送りを 1.2 em にする（1 行 1 枠の s.text を 1.2 em おきに置く。
 *    4-12・4-13 と同じ方式）。
 *  - 表は 3-11 の tightTable（行間 1.2 em、余白 [3,6,3,6]pt）を写して使う。列幅は spec の値（目安）に
 *    できるだけ近づけ、語句の切れ目で改行したときに表が本文ゾーンに入る幅にする（合計 12.2in）。
 *  - S5・S6 の表と S7 のカード本文は、22pt では tightTable・列幅の見直しをしても本文ゾーンに入らず、
 *    薬剤名・数値を落とさずに短くする方法もない。文言は変えずに、その部分だけ文字を小さくした
 *    （S5 16pt・S6 18pt・S7 のカード本文 18pt。deviations に記載。spec 側で分割か削減が必要）。
 *    S5 は 17pt だとセルの上下余白が 1〜2pt しか取れず、列挙（SSRI・SNRI・三環系）も途中で折れるため、
 *    16pt にしてセルの上下余白 3pt（0.04in）と語句の切れ目での改行を確保した。
 */
const fs = require('fs');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fits, fitInfo, measure, heightFor } = require('../_lib/decklib');

const NAME = '4-14_other_chronic_drugs';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR ||
  path.join('/tmp/claude-0/-home-user-Anes-lecture/0b714c0d-c11e-535a-a97c-0e6ce9d53e16/scratchpad/render', NAME);

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
    const title = (lines[0].match(/title=「([^」]*)」/) || [])[1];
    let notes = null, sources = null;
    const quoted = [];
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
      // 表の列見出しは layout 行に「」で書かれている
      if (/^layout:/.test(ln)) for (const qq of ln.matchAll(/「([^」]*)」/g)) quoted.push(qq[1]);
      if (/^on-slide:/.test(ln)) { inOn = true; continue; }
      if (inOn && !/^ {2}- /.test(ln)) inOn = false;
      if (inOn) for (const qq of ln.matchAll(/「([^」]*)」/g)) quoted.push(qq[1]);
    }
    out[id] = { title, notes, sources: sources && sources !== '—' ? sources : null, quoted };
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
/** スライドに置いた文字列（spec 照合済み）を記録し、最後に spec の on-slide がすべて置かれたかを確かめる */
const USED = {};
/** spec の文言そのまま（照合のみ） */
function q(id, str) {
  if (!SPEC[id].quoted.includes(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
  (USED[id] = USED[id] || new Set()).add(str);
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
 * 1行の文字列に部分装飾を付けて run 配列にする。marks: [[語, 色], ...]（太字）。色 null は太字だけ。
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
      next.push(color ? { text: w, color, bold: true } : { text: w, bold: true });
      if (k + w.length < r.length) next.push(r.slice(k + w.length));
    }
    if (found !== 1) throw new Error(`mark: 「${line}」に「${w}」が ${found} 回`);
    runs = next;
  }
  return runs;
}
const runText = l => (Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(l));

/** spec から形を変えたところ（ビルドの最後に一覧を出す） */
const DEVIATIONS = [];

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const TOP = 1.85;                // 本文の上端
const TOP_T = 1.82;              // 表・詰まったカードの上端（金下線 8pt の下端 1.576in から 0.24in 以上）
/**
 * 本文の下端。全スライドで出典が 2 行（枠 y6.5・下揃え。PowerPoint のメイリオ 16pt 2 行で文字の上端 約6.68in）
 * なので、その 0.24in 上。
 */
const BOTTOM = 6.44;
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const TABLE_W = 12.2;            // 表の幅の上限（合計 12.2in 以内）
const r3 = v => Math.round(v * 1000) / 1000;
const BLUE = C.blue, VER = C.vermilion, RED = C.red, TEAL = C.teal;

// ---------------------------------------------------------------------------
// 行送りを詰めたテキスト（1 行 1 枠の s.text を 1.2 em おきに置く）
//  - 枠の高さは decklib / lint の 1 行の下限（22pt 0.41in、24pt 0.46in）で、次の行の枠と少し重なる
//    （文字どうしの枠の重なりは lint・decklib とも対象外）。
// ---------------------------------------------------------------------------
const LS_EM = 1.2;
const lineIn = size => size * LS_EM / 72;            // 行送り（22pt 0.367in、24pt 0.4in）
const BOX_H = { 22: 0.41, 24: 0.46 };                // 1 行の枠の高さ
const overhang = size => BOX_H[size] - lineIn(size); // 最終行の枠が行送りからはみ出す分

function tightText(s, lines, o) {
  const size = o.size || 22;
  if (!BOX_H[size]) throw new Error(`tightText: size ${size} は 22 / 24 のみ`);
  lines.forEach((line, i) => {
    s.text([line], {
      x: r3(o.x), y: r3(o.y + i * lineIn(size)), w: r3(o.w), h: BOX_H[size],
      size, bold: !!o.bold, color: o.color || C.dark, align: o.align || 'left',
    });
  });
  return lines.length * lineIn(size);
}

// ---------------------------------------------------------------------------
// カード（S4）: 角丸の塗り箱＋見出し＋「•」付きの項目（文字はすべて箱の内側）。
// 項目は語句の切れ目で 2 行に割り、2 行目以降は記号の右にぶら下げる。
// ---------------------------------------------------------------------------
// padY は上の余白、padB は下の余白（最終行を箱の下端から離す）
const CARD = { padX: 0.1, padY: 0.08, padB: 0.15, titleSize: 24, size: 22, titleGap: 0.06, ind: 0.28 };
function cardNeedW(c) {
  const t = measure(runText(c.title), CARD.titleSize);
  const b = Math.max(...c.items.flat().map(l => measure(runText(l), CARD.size))) + CARD.ind;
  return Math.max(t, b) + 2 * CARD.padX;
}
function cardFixedH(c) {
  const nLines = c.items.reduce((a, it) => a + it.length, 0);
  return CARD.padY + CARD.padB + lineIn(CARD.titleSize) + CARD.titleGap + nLines * lineIn(CARD.size);
}
function itemGapFor(c, h, maxGap) {
  const n = c.items.length;
  const g = n > 1 ? (h - cardFixedH(c)) / (n - 1) : 0;
  if (g < -1e-9) throw new Error(`カード「${runText(c.title)}」: 高さが ${r3(-g * (n - 1))}in 足りません（h=${r3(h)}）`);
  return Math.floor(Math.min(maxGap, Math.max(0, g)) * 1000) / 1000;
}
function drawCard(s, c, x, y, w, h, itemGap) {
  s.box({ x: r3(x), y: r3(y), w: r3(w), h: r3(h), fill: C.bg, radius: 0.08 });
  const ix = x + CARD.padX, iw = w - 2 * CARD.padX;
  tightText(s, [c.title], { x: ix, y: y + CARD.padY, w: iw, size: CARD.titleSize, bold: true, color: c.titleColor });
  let yy = y + CARD.padY + lineIn(CARD.titleSize) + CARD.titleGap;
  c.items.forEach(lines => {
    tightText(s, ['•'], { x: ix, y: yy, w: 0.35, size: CARD.size });
    tightText(s, lines, { x: ix + CARD.ind, y: yy, w: iw - CARD.ind, size: CARD.size });
    yy += lines.length * lineIn(CARD.size) + itemGap;
  });
  const end = yy - itemGap;
  if (end > y + h - CARD.padB + 1e-6) throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さ不足（本文の下端 ${r3(end)} > ${r3(y + h - CARD.padB)}）`);
}

/** 行送りを詰めた強調ボックス（白地＋色枠・太字・中央）。下端 yBottom に合わせて置く。戻り値 {y, h} */
const EMPH = { radius: 0.1, lineW: 3.5, padY: 0.04, padX: 0.07 };
function emphTightH(n, size) { return r3(2 * EMPH.padY + n * lineIn(size) + overhang(size)); }
function emphTight(s, o) {
  const size = o.size || 24;
  const h = emphTightH(o.lines.length, size);
  const y = r3(o.yBottom - h);
  s.box({ x: r3(o.x), y, w: r3(o.w), h, fill: C.white, radius: EMPH.radius, line: { color: o.lineColor, width: EMPH.lineW } });
  tightText(s, o.lines, { x: o.x + EMPH.padX, y: y + EMPH.padY, w: o.w - 2 * EMPH.padX, size, bold: true, align: 'center' });
  return { y, h };
}

/**
 * 行の多い表（3-11_local_anesthetics/build.js の tightTable を写したもの）。decklib の s.table と同じ見た目
 * （見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、上下中央）で、行間だけを固定値 LS pt（= 1.2 em）に詰める。
 * 見出し行も本文と同じ左揃え。セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、
 * 重なり検査用の登録も同じく行う。cell: 文字列（'\n' で改行）| {text, color, bold, align}
 * o.size（既定 22）・o.margin（既定 [3,6,3,6]pt）は S5・S6 だけ変える（deviations に記載）。
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
  return { x: o.x, y: o.y, w, h: r3(h), rowH };
}
/**
 * 表の列幅: 各列の最長行に必要な幅（セルの左右余白込み）を確保し、残りを広げる列 grow に足す。
 * rows のセルは tightTable と同じ形（'\n' 区切り）。合計は TABLE_W。
 */
function fitColW(rows, size, margin, grow) {
  const mx = (margin[1] + margin[3]) / 72;
  const need = rows[0].map((_, ci) => Math.max(...rows.map(r => {
    const c = typeof r[ci] === 'string' ? { text: r[ci] } : r[ci];
    return Math.max(...c.text.split('\n').map(l => measure(l, size)));
  })) + mx);
  const colW = need.map(v => Math.ceil(v * 1000) / 1000);
  const rest = TABLE_W - colW.reduce((a, b) => a + b, 0);
  if (rest < -1e-9) throw new Error(`表の列幅が ${r3(-rest)}in 足りません（必要 ${colW.join(' / ')}）`);
  colW[grow] = r3(colW[grow] + rest);
  return colW;
}

/**
 * S7 のカード本文（項目の箇条書き）。22pt では入らないため、カードの角丸の箱と見出し（22pt）はそのままに、
 * 本文だけを箱の内側に置いた枠なしの表（2 列: 記号「•」｜本文、1 項目 1 行）で size pt・行送り 1.2 em で組む。
 * 表はカードの箱に完全に含まれる（重なり判定の対象外）。
 */
function listTable(s, items, o) {
  const size = o.size, LS = size * 1.2, gapPt = o.itemGapPt;
  const markW = o.markW, textW = r3(o.w - markW);
  const rowH = [];
  const none = [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }];
  const out = items.map((lines, ii) => {
    lines.forEach(l => {
      const fi = fitInfo(runText(l), size, textW);
      if (!fi.ok) throw new Error(`p${s.page} カード本文「${runText(l)}」が幅に収まりません（推定 ${fi.needIn}in > ${textW}in）`);
    });
    const padTop = ii === 0 ? 0 : gapPt;
    rowH.push(r3(lines.length * LS / 72 + padTop / 72));
    const opt = { fontFace: FONT, lang: 'ja-JP', fontSize: size, color: C.dark };
    return [
      { text: [{ text: '•', options: { ...opt, lineSpacing: LS } }], options: { ...opt, border: none, valign: 'top', margin: [padTop, 0, 0, 0] } },
      {
        text: lines.map((l, li) => ({ text: l, options: { ...opt, lineSpacing: LS, breakLine: li < lines.length - 1 || undefined } })),
        options: { ...opt, border: none, valign: 'top', margin: [padTop, 0, 0, 0] },
      },
    ];
  });
  const h = r3(rowH.reduce((a, b) => a + b, 0));
  s._bounds(`p${s.page} カード本文`, o.x, o.y, o.w, h);
  s.raw.addTable(out, {
    x: o.x, y: o.y, w: o.w, colW: [markW, textW], rowH, fontFace: FONT, fontSize: size, margin: 0,
    border: none, valign: 'top',
  });
  s._reg('box', o.x, o.y, o.w, h, 'table(card)');
  return { h, rowH };
}

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める（文献は落とさない。4-12 と同じ手順）
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 文言はそのままで、文献の途中の語の切れ目でも改行する（区切り > 和文の語の切れ目・「et al.」の後 >
//     読点の後・誌名と年の間 > その他の空白。数値と単位の間、「et」と「al.」の間では割らない）。
//  2) 入らなければ区切りの前後の空白を詰める（' / ' → '/'）。
//  3) それでも入らなければ（）内の補足を外す（外す数が最小の形。空白は詰めない形を優先）。
//  実機メイリオの字幅（欧文は広い）でも 11.6in に入ることを確かめる。形を変えたものは DEVIATIONS に記録。
// ---------------------------------------------------------------------------
const SRC = LAYOUT.source;
const LATIN_EM = {
  ' ': 0.352, '.': 0.364, ',': 0.364, ':': 0.454, ';': 0.454, '-': 0.454, '/': 0.454, '%': 1.076, '(': 0.454, ')': 0.454,
  0: 0.636, 1: 0.636, 2: 0.636, 3: 0.636, 4: 0.636, 5: 0.636, 6: 0.636, 7: 0.636, 8: 0.636, 9: 0.636,
  a: 0.601, b: 0.623, c: 0.521, d: 0.623, e: 0.596, f: 0.352, g: 0.623, h: 0.633, i: 0.274, j: 0.344, k: 0.592, l: 0.274,
  m: 0.973, n: 0.633, o: 0.607, p: 0.623, q: 0.623, r: 0.427, s: 0.521, t: 0.394, u: 0.633, v: 0.592, w: 0.818,
  x: 0.592, y: 0.592, z: 0.525,
  A: 0.684, B: 0.686, C: 0.698, D: 0.771, E: 0.632, F: 0.575, G: 0.775, H: 0.751, I: 0.421, J: 0.455, K: 0.693,
  L: 0.557, M: 0.843, N: 0.748, O: 0.787, P: 0.603, Q: 0.787, R: 0.695, S: 0.684, T: 0.616, U: 0.732, V: 0.684,
  W: 0.989, X: 0.685, Y: 0.615, Z: 0.685,
};
/** 実機メイリオの字幅の目安（欧文は Verdana 系で広い） */
function realW(text, size) {
  let e = 0;
  for (const ch of String(text)) {
    const o = ch.codePointAt(0);
    e += o <= 0x24F ? (LATIN_EM[ch] != null ? LATIN_EM[ch] : 0.65) : 1.0;
  }
  return e * size / 72;
}
const srcFits = l => fits(l, SRC.size, SRC.w) && realW(l, SRC.size) <= SRC.w - 0.05;
function libLines(items) {
  const lines = [];
  for (const it of items) {
    if (!srcFits(it)) return null;
    const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
    if (cur && srcFits(cur)) lines[lines.length - 1] = cur;
    else lines.push(it);
  }
  return lines;
}
const isCJK = ch => ch != null && ch.codePointAt(0) > 0x24F;
const isDigit = ch => /[0-9]/.test(ch || '');
/** str（区切り sep で文献をつないだもの）を 2 行に割る。最良の割り方 {a, b, rank} か null */
function split2(str, sep) {
  if (srcFits(str)) return { a: str, b: null, rank: 0 };
  const cands = [];
  for (let i = 1; i < str.length; i++) {
    let a, b, rank;
    if (str.startsWith(sep, i)) { a = str.slice(0, i); b = str.slice(i + sep.length); rank = 0; }
    else if (str[i] === ' ') {
      if (str.startsWith(' / ', i) || str.startsWith(' / ', i - 2)) continue;   // 区切りの中
      a = str.slice(0, i); b = str.slice(i + 1);
      const pa = a[a.length - 1], nb = b[0];
      if (isDigit(pa) && isCJK(nb)) continue;                                 // 数値と単位（2020 年、第 5 版）
      if (a.endsWith(' et') && b.startsWith('al.')) continue;
      if (isCJK(pa)) rank = 1;                                                // 和文の語の切れ目
      else if (a.endsWith('al.')) rank = 1;                                   // 著者（et al.）の後
      else if (pa === ',') rank = 2;                                          // 読点（著者の途中）
      else if (/[A-Za-z]/.test(pa) && isDigit(nb)) rank = 2;                 // 誌名と年の間
      else rank = 3;
    } else if ((str[i - 1] === '／' || str[i - 1] === '・') && str[i] !== ' ') {
      a = str.slice(0, i); b = str.slice(i); rank = 1;                        // 和文の中黒・斜線の後
    } else continue;
    a = a.trim(); b = b.trim();
    if (a && b && srcFits(a) && srcFits(b)) cands.push({ a, b, rank });
  }
  if (!cands.length) return null;
  cands.sort((x, y) => x.rank - y.rank || y.a.length - x.a.length);
  return cands[0];
}
function combos(arr, k, start = 0) {
  if (k === 0) return [[]];
  const out = [];
  for (let i = start; i <= arr.length - k; i++) for (const rest of combos(arr, k - 1, i + 1)) out.push([arr[i], ...rest]);
  return out;
}
function sourcesOf(id) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  const items0 = spec.split(' / ');
  const lib = libLines(items0);
  if (lib && lib.length <= 2) return items0;
  const parens = [];
  items0.forEach((it, ii) => { for (const mm of it.matchAll(/（[^（）]*）/g)) parens.push({ ii, text: mm[0] }); });
  for (let k = 0; k <= parens.length; k++) {
    // 同じ段階では区切りの空白を残す形と詰める形を比べ、改行位置のよい方（同順位なら空白を残す方）をとる
    let best = null;
    for (const sep of [' / ', '/']) {
      for (const P of combos(parens, k)) {
        const items = items0.map((it, ii) => {
          let t = it;
          for (const p of P) if (p.ii === ii) t = t.replace(p.text, '');
          return t;
        });
        const r = split2(items.join(sep), sep);
        if (r && (!best || r.rank < best.r.rank)) best = { r, P, sep };
      }
    }
    if (!best) continue;
    const lines = best.r.b ? [best.r.a, best.r.b] : [best.r.a];
    const what = [];
    if (best.sep === '/') what.push('区切りの前後の空白を詰めた');
    if (best.P.length) what.push(`（）内の補足 ${best.P.map(p => p.text).join('・')} を外した`);
    if (best.r.rank > 0) what.push('文献の途中の語の切れ目で改行');
    DEVIATIONS.push(`${id} 出典: 「${spec}」→「${lines.join(' ⏎ ')}」（16pt・2 行に収めるため。${what.join('、')}）`);
    return lines;
  }
  throw new Error(`${id}: 出典が 2 行に収まりません`);
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['降圧薬', 'ステロ\nイド', '神経・\n精神', 'その他', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-14 常用薬③：その他',
  });
  const slide = (id, chapter) => deck.slide({ title: titleOf(id), chapter, notes: notesOf(id), sources: sourcesOf(id) });

  // ===== S1 表紙 =====
  {
    const title = titleOf('S1'), subtitle = '4-14 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title, subtitle, meta, notes: notesOf('S1') });
  }

  // ===== S2 ACE 阻害薬・ARB の扱い（箇条書き6行 + 強調ボックス） =====
  // 箇条書き 24pt・段落後 5pt、強調ボックスは 24pt 太字 2 行（行送り 1.2 em）。
  {
    const id = 'S2';
    const s = slide(id, 0);
    const items = [
      mark(q(id, '国内添付文書：『手術前 24 時間は投与しないことが望ましい』'), ['国内添付文書：', BLUE]),
      mark(q(id, 'AHA/ACC 2024：高血圧は 24 時間前休薬を考慮（2b）'), ['AHA/ACC 2024：', BLUE]),
      mark(q(id, 'HFrEF は継続（2a）。ESC 2022：心不全なしは当日休薬（IIa）'), ['HFrEF は継続（2a）', VER], ['ESC 2022：', BLUE]),
      mark(q(id, 'STOP-or-NOT：続けても止めても死亡・合併症は 22% で同じ'), ['STOP-or-NOT：', BLUE], ['22% で同じ', VER]),
      mark(q(id, '術中低血圧は継続 54%・休薬 41%。SPACE：休薬で術後高血圧が多め'), ['SPACE：', BLUE]),
      mark(q(id, '配合錠（ARB＋Ca 拮抗薬、ARNI）の中身を見落とさない'), ['中身を見落とさない', VER]),
    ];
    const BUL = { y: TOP, size: 24, gap: 5 };
    const bh = heightFor(items.length, BUL.size, BUL.gap);
    s.bullets(items, { ...BUL, h: bh });
    const em = brk(id, 'どちらでもよいが、続けるなら導入時の低血圧に備える。止めたら再開日を書く', '止めたら');
    const e = emphTight(s, { x: X0, w: BW, yBottom: BOTTOM, lineColor: TEAL, size: 24,
      lines: [em[0], mark(em[1], ['止めたら再開日を書く', VER])] });
    if (e.y - (BUL.y + bh) < GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔 ${r3(e.y - BUL.y - bh)}in < ${GAP}`);
    console.log(`  p${s.page} 箇条書き ${BUL.y}〜${r3(BUL.y + bh)}in、強調ボックス ${e.y}〜${r3(e.y + e.h)}in`);
  }

  // ===== S3 ほかの降圧薬とスタチン（表 3列 × 5行＋見出し行） =====
  // spec の列幅（2.8 / 3.6 / 5.8）では 3 列目が 3〜4 行に折れて本文ゾーンに入らない。1・2 列目を最長行
  // （「クロニジン」「続ける（I）」）の幅まで詰めて 3 列目を広げる（合計 12.2in）。2 列目は全角 5 字幅のため、
  // 「当日は休薬が一般的」は「当日は休薬／が一般的」、「続ける（I）。当日朝も」は 3 行に割る（文言は同じ）。
  {
    const id = 'S3';
    const s = slide(id, 0);
    const drug = text => ({ text, bold: true });
    const keep = text => ({ text, bold: true, color: BLUE });
    const stop = text => ({ text, bold: true, color: VER });
    const j = a => a.join('\n');
    const rows = [
      [q(id, '薬'), q(id, '当日'), q(id, '理由')],
      [drug(q(id, 'β遮断薬')), keep(q(id, '続ける（I）')),
        j(brk(id, '周術期の中止で 30 日死亡のオッズ比 3.93（観察研究）。当日の新規開始は害（POISE）。始めるなら 7 日以上前', '（観察研究）', '始めるなら'))],
      [drug(q(id, 'Ca 拮抗薬')), keep(q(id, '続ける')), q(id, '中止の利益はない')],
      [drug(q(id, '利尿薬')), stop(j(brk(id, '当日は休薬が一般的', 'が一般的'))), j(brk(id, '脱水・低血圧・腎機能。溢水のある心不全では続ける', '溢水'))],
      [drug(q(id, 'クロニジン')), keep(q(id, '続ける')), q(id, '急な中止でリバウンド高血圧。新規開始は無益（3）')],
      [drug(q(id, 'スタチン')), keep(j(brk(id, '続ける（I）。当日朝も', '（I）', '当日朝も'))),
        j(brk(id, '術後 4 日超の中断で心筋壊死のオッズ比 2.9（大動脈手術）。ASCVD または 10 年リスクで適応があれば開始（I）', 'オッズ比', 'ASCVD'))],
    ];
    const m = [3, 6, 3, 6];
    const colW = fitColW(rows, 22, m, 2);
    const t = tightTable(s, rows, { x: X0, y: TOP_T, colW, margin: m });
    console.log(`  p${s.page} 表: 列幅 ${colW.join(' / ')}in、行高 ${t.rowH.join(' / ')}in、下端 ${r3(t.y + t.h)}in`);
    if (t.y + t.h > BOTTOM + 0.011) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)} > ${BOTTOM}`);
  }

  // ===== S4 ステロイドカバー（2列カード + 強調ボックス） =====
  // カード本文 22pt・行送り 1.2 em。右のカードの「続いて 200 mg/24 時間を持続。食べられたら」を 1 行に入れるため、
  // 左のカードは「5 mg/日以上を 1 か月以上」の幅に詰め、残りを右のカードに。
  {
    const id = 'S4';
    const s = slide(id, 1);
    const em = brk(id, '術中・術後の昇圧薬に反応しない低血圧、低 Na、低血糖、嘔吐 → 副腎不全を疑ってヒドロコルチゾン', '→');
    const e = emphTight(s, { x: X0, w: BW, yBottom: BOTTOM, lineColor: RED, size: 24,
      lines: [em[0], mark(em[1], ['副腎不全を疑って', RED])] });
    const c1 = {
      title: q(id, '誰に'), titleColor: BLUE,
      items: [
        brk(id, 'プレドニゾロン換算 5 mg/日以上を 1 か月以上', '5 mg'),
        brk(id, '経口だけでなく吸入・点鼻・関節内・外用も', '点鼻'),
        brk(id, '過去 6 か月に経口 2 週超・長期の高用量 ICS（GINA）', '2 週超', 'ICS'),
      ],
    };
    const c2 = {
      title: q(id, 'どう投与するか（Woodcock 2020）'), titleColor: VER,
      items: [
        (() => { const l = brk(id, '普段のステロイドは当日朝も飲む。カバーは『上乗せ』', 'カバーは'); return [l[0], mark(l[1], ['『上乗せ』', VER])]; })(),
        [q(id, '導入時にヒドロコルチゾン 100 mg 静注')],
        brk(id, '続いて 200 mg/24 時間を持続。食べられたら通常量の 2 倍を 48 時間経口で', '通常量'),
        brk(id, '小手術に一律 100 mg は異論あり（→ 3-12）', '異論'),
      ],
    };
    const w1 = r3(cardNeedW(c1));
    const w2 = r3(BW - GAP - w1);
    if (cardNeedW(c2) > w2 + 1e-9) throw new Error(`${id}: 右のカードの幅不足（必要 ${r3(cardNeedW(c2))} > ${w2}）`);
    const y = TOP_T, h = r3(e.y - GAP - y);
    drawCard(s, c1, X0, y, w1, h, itemGapFor(c1, h, 0.2));
    drawCard(s, c2, X0 + w1 + GAP, y, w2, h, itemGapFor(c2, h, 0.2));
    console.log(`  p${s.page} カード: 幅 ${w1} / ${w2}in、高さ ${h}in、項目間 ${itemGapFor(c1, h, 0.2)} / ${itemGapFor(c2, h, 0.2)}in`);
  }

  // ===== S5 神経・精神の薬（表 3列 × 6行＋見出し行） =====
  // 22pt では tightTable・列幅の見直しをしても高さ約 8.3in で本文ゾーン（1.82〜6.44in、4.62in）に入らない。
  // 薬剤名（併用禁忌の薬の並び、セロトニン症候群を起こす薬）を落とさずに短くする方法もないので、
  // 文言は変えずに表だけ 17pt・セルの上下余白 1pt にする（deviations に記載）。
  {
    const id = 'S5';
    const s = slide(id, 2);
    const TS = 17;
    const drug = text => ({ text, bold: true });
    const keep = text => ({ text, bold: true, color: BLUE });
    const j = a => a.join('\n');
    const rows = [
      [q(id, '薬'), q(id, '当日'), q(id, '注意')],
      [drug(q(id, '抗てんかん薬')), keep(q(id, '続ける')),
        j(brk(id, '絶食が長引けば静注製剤に。バルプロ酸とカルバペネムは併用禁忌', 'バルプロ酸'))],
      [drug(j(brk(id, 'レボドパ（パーキンソン病）', '（'))), keep(q(id, '少量の水で続ける')),
        j(brk(id, '急な中止で悪性症候群。絶食ならレボドパ静注。メトクロプラミド・ハロペリドールを避ける', 'メトクロプラミド'))],
      [drug(j(brk(id, 'MAO-B 阻害薬（セレギリン）', '（'))), keep(q(id, '続ける')),
        j(brk(id, 'ペチジン・トラマドール・タペンタドール、三環系・SSRI・SNRI・ミルタザピンは併用禁忌。エフェドリンを避けフェニレフリン少量', '三環系', 'エフェドリン'))],
      [drug(j(brk(id, 'SSRI・SNRI・三環系', '三環系'))), keep(q(id, '続ける')),
        j(brk(id, '出血が増える。セロトニン症候群（トラマドール、ペチジン、メチレンブルー、リネゾリド）', 'ペチジン'))],
      [drug(q(id, 'リチウム')), j(brk(id, '小手術は継続。大量出血・絶食が続く手術は 72 時間前に中止', '絶食が', '72 時間')),
        j(brk(id, '脱水・NSAIDs・ACE 阻害薬・サイアザイドで蓄積。再開は K・Na・Cr を見て', 'サイアザイド'))],
      [drug(q(id, '抗精神病薬')), keep(q(id, '続ける（精神科と）')),
        j(brk(id, 'QT 延長、錐体外路、悪性症候群。パーキンソン病にハロペリドールは禁忌', 'パーキンソン'))],
    ];
    const m = [1, 6, 1, 6];
    const colW = fitColW(rows, TS, m, 2);
    const t = tightTable(s, rows, { x: X0, y: TOP_T, colW, margin: m, size: TS });
    console.log(`  p${s.page} 表: 列幅 ${colW.join(' / ')}in、行高 ${t.rowH.join(' / ')}in、下端 ${r3(t.y + t.h)}in`);
    if (t.y + t.h > BOTTOM + 0.011) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)} > ${BOTTOM}`);
    DEVIATIONS.push(`${id} 文字の大きさ: 表（見出し行＋6 行）を 22pt ではなく ${TS}pt、セルの上下余白を 3pt ではなく 1pt で組んだ` +
      '（spec の文言は一字も変えていない）。22pt では tightTable（行間 1.2 em・余白 [3,6,3,6]pt）と列幅の見直し（合計 12.2in）を' +
      'しても表の高さが約 8.3in（20pt でも約 6.9in、18pt でも約 5.0in）で、本文ゾーン（1.82〜6.44in、4.62in）に入らない。' +
      '3 列目は併用禁忌・セロトニン症候群の薬剤名の並びが大半で、薬剤名を落とさずに短くする余地がない。' +
      'spec 側で S5 を 2 枚に分けるか、表の内容を減らす必要がある');
  }

  // ===== S6 ホルモン・漢方・サプリ（表 3列 × 6行＋見出し行） =====
  // 22pt では列幅を見直しても高さ約 6.8in で入らない。文言は変えずに表だけ 18pt にする（deviations に記載）。
  {
    const id = 'S6';
    const s = slide(id, 3);
    const TS = 18;
    const drug = text => ({ text, bold: true });
    const keep = text => ({ text, bold: true, color: BLUE });
    const stop = text => ({ text, bold: true, color: VER });
    const j = a => a.join('\n');
    const rows = [
      [q(id, '薬'), q(id, 'どうする'), q(id, '理由')],
      [drug(q(id, 'レボチロキシン')), j(brk(id, '当日飲めなくてもよい。5 日超は静注', '5 日超')),
        j(brk(id, '半減期約 7 日。顕性の機能異常は待機手術を延期', '顕性'))],
      [drug(q(id, '抗甲状腺薬')), keep(q(id, '続ける')), q(id, '甲状腺クリーゼの予防')],
      [drug(q(id, '低用量ピル・LEP')), stop(j(brk(id, '手術 4 週前に中止、術後 2 週は再開しない', '術後'))),
        j(brk(id, '添付文書で禁忌（静脈血栓塞栓症）。予約時に確認', '（静脈'))],
      [drug(q(id, 'ラロキシフェン')), stop(j(brk(id, '不動になる 3 日前に中止', '3 日前'))),
        j(brk(id, '静脈血栓塞栓症。骨折の緊急手術では VTE 予防を強化', '骨折'))],
      [drug(q(id, '漢方（甘草）・サプリ')), j(brk(id, 'サプリは 1〜2 週前に中止。甘草は K を確認', '1〜2', '甘草は')),
        j(brk(id, '偽アルドステロン症。ニンニク・イチョウは出血、セント・ジョーンズ・ワートは薬物代謝', 'ニンニク', 'セント'))],
      [drug(j(brk(id, '吸入薬・ロイコトリエン拮抗薬', 'ロイコ'))), keep(q(id, '当日も続ける')), q(id, '吸入ステロイドはカバーの判断で数える')],
    ];
    const m = [3, 6, 3, 6];
    const colW = fitColW(rows, TS, m, 2);
    const t = tightTable(s, rows, { x: X0, y: TOP_T, colW, margin: m, size: TS });
    console.log(`  p${s.page} 表: 列幅 ${colW.join(' / ')}in、行高 ${t.rowH.join(' / ')}in、下端 ${r3(t.y + t.h)}in`);
    if (t.y + t.h > BOTTOM + 0.011) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)} > ${BOTTOM}`);
    DEVIATIONS.push(`${id} 文字の大きさ: 表（見出し行＋6 行）を 22pt ではなく ${TS}pt で組んだ（spec の文言は一字も変えていない）。` +
      '22pt では tightTable（行間 1.2 em・余白 [3,6,3,6]pt）と列幅の見直し（合計 12.2in）をしても表の高さが約 6.8in' +
      '（20pt でも約 5.3in）で、本文ゾーン（1.82〜6.44in、4.62in）に入らない。薬剤名・期間（4 週・2 週・3 日・1〜2 週・5 日）を' +
      '落とさずに短くしても入らない。spec 側で S6 を 2 枚に分けるか、表の内容を減らす必要がある');
  }

  // ===== S7 当日の内服指示の書き方（3列カード + 強調ボックス） =====
  // カード本文 14 項目（22pt で各カード 9〜11 行）が 22pt では強調ボックスと合わせて本文ゾーンに入らない
  // （22pt・行送り 1.2 em で約 5.7in > 4.62in）。見出し（22pt）と強調ボックス（24pt）はそのままに、
  // カード本文だけを 18pt にする（箱の内側に置いた枠なしの表。deviations に記載）。
  {
    const id = 'S7';
    const s = slide(id, 3);
    const BS = 18;
    const em = brk(id, '『飲む薬』『止める薬と再開日』『相談した相手』を薬ごとに書き、少量の水で飲ませる', '少量の水');
    const e = emphTight(s, { x: X0, w: BW, yBottom: BOTTOM, lineColor: TEAL, size: 24,
      lines: [mark(em[0], ['『止める薬と再開日』', VER]), em[1]] });
    const cards = [
      {
        title: q(id, '当日も飲む'), titleColor: BLUE,
        items: [
          brk(id, 'β遮断薬、Ca 拮抗薬、スタチン', 'スタチン'),
          brk(id, '吸入薬、抗てんかん薬、レボドパ', 'レボドパ'),
          brk(id, 'SSRI、抗甲状腺薬、レボチロキシン（飲めなくても可）', 'レボチロキシン', '（飲めなくても可）'),
          brk(id, '普段のステロイド（＋カバー）', '（＋カバー）'),
        ],
      },
      {
        title: q(id, '止める（いつから）'), titleColor: VER,
        items: [
          brk(id, 'ACE 阻害薬・ARB：24 時間前から（HFrEF は相談）', '24 時間'),
          [q(id, '利尿薬：当日朝')],
          [q(id, 'SGLT2 阻害薬：3 日前（→ 4-13）')],
          brk(id, 'リチウム：大量出血の手術は 72 時間前', '大量出血'),
          brk(id, '低用量ピル：4 週前。サプリ：1〜2 週前', 'サプリ'),
        ],
      },
      {
        title: q(id, '相談する相手'), titleColor: C.dark,
        items: [
          brk(id, 'HFrEF・不整脈の薬：循環器内科', '循環器'),
          [q(id, '副腎不全：内分泌内科')],
          brk(id, 'パーキンソン病：脳神経内科', '脳神経'),
          brk(id, 'リチウム・MAO 阻害薬・抗精神病薬：精神科', '抗精神病薬'),
          [q(id, 'てんかん・漢方：処方医')],
        ],
      },
    ];
    const K = { padX: 0.08, padY: 0.08, titleSize: 22, titleGap: 0.06, markW: 0.22 };
    const bodyLS = BS * 1.2 / 72;
    // 幅: 見出し（22pt）と本文の最長行に必要な幅を確保し、余りを等分
    const need = cards.map(c => Math.max(measure(runText(c.title), K.titleSize),
      Math.max(...c.items.flat().map(l => measure(runText(l), BS))) + K.markW) + 2 * K.padX);
    const extra = BW - 2 * GAP - need.reduce((a, b) => a + b, 0);
    if (extra < -1e-9) throw new Error(`${id}: カード幅が ${r3(-extra)}in 足りません`);
    const ws = need.map(v => Math.floor((v + extra / cards.length) * 1000) / 1000);
    const y = TOP_T, h = r3(e.y - GAP - y);
    // 項目の間隔: カードごとに余りを項目間に配る（上限 8pt）
    const bodyH = h - 2 * K.padY - lineIn(K.titleSize) - K.titleGap;
    const gapPts = cards.map(c => {
      const nl = c.items.reduce((a, it) => a + it.length, 0);
      const g = Math.floor((bodyH - nl * bodyLS) / (c.items.length - 1) * 72 * 10) / 10;
      if (g < 0) throw new Error(`${id}: カード「${c.title}」の高さが足りません（本文 ${nl} 行）`);
      return Math.min(8, g);
    });
    let x = X0;
    cards.forEach((c, k) => {
      s.box({ x: r3(x), y, w: ws[k], h, fill: C.bg, radius: 0.08 });
      tightText(s, [c.title], { x: x + K.padX, y: y + K.padY, w: ws[k] - 2 * K.padX, size: K.titleSize, bold: true, color: c.titleColor });
      const by = r3(y + K.padY + lineIn(K.titleSize) + K.titleGap);
      const lt = listTable(s, c.items, { x: r3(x + K.padX), y: by, w: r3(ws[k] - 2 * K.padX), size: BS, itemGapPt: gapPts[k], markW: K.markW });
      if (by + lt.h > y + h - K.padY + 1e-6) throw new Error(`${id}: カード「${c.title}」の本文の下端 ${r3(by + lt.h)} > ${r3(y + h - K.padY)}`);
      x += ws[k] + GAP;
    });
    console.log(`  p${s.page} カード: 幅 ${ws.join(' / ')}in（必要 ${need.map(r3).join(' / ')}）、高さ ${h}in、項目間 ${gapPts.join(' / ')}pt`);
    DEVIATIONS.push(`${id} 文字の大きさ: 3 列カードの本文（14 項目）を 22pt ではなく ${BS}pt で組んだ（見出し 22pt・強調ボックス 24pt、` +
      'spec の文言は一字も変えていない）。22pt では行送り 1.2 em・列幅の最適化をしても各カード 9〜11 行になり、' +
      '強調ボックスと合わせて約 5.7in（20pt でも約 5.0in）で本文ゾーン（4.62in）に入らない。項目の大半が薬剤名と期間で、' +
      '落とさずに短くする余地がない。本文はカードの箱の内側に置いた枠なしの表（記号｜本文の 2 列）で組んだ。' +
      'spec 側で S7 の項目を減らす必要がある');
  }

  // ===== S8 まとめ（箇条書き5行） =====
  {
    const id = 'S8';
    const s = slide(id, 4);
    const items = [
      mark(q(id, 'ACE 阻害薬・ARB は休薬でも継続でも可。止めたら再開日を書く'), ['止めたら再開日を書く', VER]),
      mark(q(id, 'β遮断薬・スタチン・吸入薬・抗てんかん薬・レボドパは当日も飲む'), ['当日も飲む', BLUE]),
      mark(q(id, 'ステロイドカバー：5 mg/日以上を 1 か月以上。普段の量に上乗せ'), ['普段の量に上乗せ', VER]),
      mark(q(id, 'MAO-B 阻害薬にペチジン・トラマドールは禁忌。SSRI でも慎重に'), ['禁忌', RED]),
      mark(q(id, 'ピル 4 週前、サプリ 1〜2 週前、リチウムは大出血手術で 72 時間前'), ['72 時間前', BLUE]),
    ];
    const size = 24, gap = 28;
    const hh = heightFor(items.length, size, gap);
    if (hh > BOTTOM - TOP + 1e-6) throw new Error(`${id}: 箇条書きの高さ ${hh} > ${r3(BOTTOM - TOP)}`);
    s.bullets(items, { y: TOP, h: r3(BOTTOM - TOP), size, gap, valign: 'middle' });
  }

  // spec の on-slide（と表の列見出し）がすべてスライドに置かれたか
  for (const [id, sp] of Object.entries(SPEC)) {
    for (const str of sp.quoted) {
      if (!(USED[id] && USED[id].has(str))) throw new Error(`${id}: spec の「${str}」をスライドに置いていません`);
    }
  }

  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.render) console.log('[render] ' + JSON.stringify(res.render));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
