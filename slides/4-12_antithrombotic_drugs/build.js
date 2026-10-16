'use strict';
/**
 * 4-12 常用薬①：抗血栓薬
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk()/q() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。spec から形を変えたところはビルド時に deviations として一覧を出す。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら scratchpad の render/<NAME>）。
 *
 * 密度への対応（文言を変えずに収めるための組版上の工夫）:
 *  - カード・強調ボックスの本文は 22pt のまま行送りを 1.2 em にする（1 行 1 枠の s.text を 1.2 em おきに置く。
 *    4-13 と同じ方式。spcPts は slide_lint の高さ推定を壊し、spcPct は代替フォントで行送りが変わるため）。
 *  - 行の多い表（S5・S6）は 3-11 の tightTable（行間 1.2 em、余白 [3,6,3,6]pt）を写して使う。
 *  - S2 は 2 列カード＋全幅の強調ボックスでは 22pt で入らないため、右のカードを本文ゾーンの高さいっぱいにし、
 *    強調ボックスを左のカードの下に置く（文言は変えない）。
 *  - S6 の表は 22pt では tightTable・列幅調整をしても高さ 6.9in（本文ゾーン 4.6in）になり、
 *    数値・薬剤名を落とさずに短くする方法もない。数値を欠いた区域麻酔の休薬表は危険なので、
 *    文言は変えずに表だけ 18pt にした（deviations に記載。spec 側で S6 の分割か削減が必要）。
 */
const fs = require('fs');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fits, fitInfo, measure, heightFor } = require('../_lib/decklib');

const NAME = '4-12_antithrombotic_drugs';
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
// カード: 角丸の塗り箱＋見出し＋項目（文字はすべて箱の内側）。
// marker: true なら項目の頭に「•」を置き、2 行目以降は記号の右にぶら下げる。
// padY は上の余白。padB は最終行の枠（高さ BOX_H）の下端からカードの下端までの余白で、
// 実機メイリオ（行高が大きい）でも最終行がカードの縁に寄らないよう 0.1in 以上とる。
// ---------------------------------------------------------------------------
const CARD = { padX: 0.1, padY: 0.08, padB: 0.12, titleSize: 22, size: 22, titleGap: 0.06, ind: 0.28 };
function cardOpt(o) { return { ...CARD, ...(o || {}) }; }
/** カードの内側に必要な幅（推定器ベース、余白を含む） */
function cardNeedW(c, o) {
  const k = cardOpt(o);
  const t = measure(runText(c.title), k.titleSize);
  const ind = c.marker === false ? 0 : k.ind;
  const b = Math.max(...c.items.flat().map(l => measure(runText(l), k.size))) + ind;
  return Math.max(t, b) + 2 * k.padX;
}
/** 項目の間隔を除いた高さ */
function cardFixedH(c, o) {
  const k = cardOpt(o);
  const nLines = c.items.reduce((a, it) => a + it.length, 0);
  return k.padY + lineIn(k.titleSize) + k.titleGap + nLines * lineIn(k.size) + overhang(k.size) + k.padB;
}
function drawCard(s, c, x, y, w, h, itemGap, o) {
  const k = cardOpt(o);
  s.box({ x: r3(x), y: r3(y), w: r3(w), h: r3(h), fill: C.bg, radius: 0.08 });
  const ix = x + k.padX, iw = w - 2 * k.padX;
  tightText(s, [c.title], { x: ix, y: y + k.padY, w: iw, size: k.titleSize, bold: true, color: c.titleColor });
  let yy = y + k.padY + lineIn(k.titleSize) + k.titleGap;
  const ind = c.marker === false ? 0 : k.ind;
  c.items.forEach(lines => {
    if (ind) tightText(s, ['•'], { x: ix, y: yy, w: 0.35, size: k.size });
    tightText(s, lines, { x: ix + ind, y: yy, w: iw - ind, size: k.size });
    yy += lines.length * lineIn(k.size) + itemGap;
  });
  const end = yy - itemGap;
  const lastBox = end + overhang(k.size);   // 最終行の枠の下端
  if (lastBox + k.padB > y + h + 1e-6) {
    throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さ不足（最終行の枠の下端 ${r3(lastBox)} + 余白 ${k.padB} > ${r3(y + h)}）`);
  }
  return end;
}
/** 項目の間隔（余りを項目間に配る。上限 maxGap） */
function itemGapFor(c, h, maxGap, o) {
  const n = c.items.length;
  const g = n > 1 ? (h - cardFixedH(c, o)) / (n - 1) : 0;
  if (g < -1e-9) throw new Error(`カード「${runText(c.title)}」: 高さが ${r3(-g * (n - 1))}in 足りません（h=${r3(h)}）`);
  return Math.floor(Math.min(maxGap, Math.max(0, g)) * 1000) / 1000;
}
/**
 * 横並び 2 列のカード（高さ y〜bottom 共通）。幅は各カードの必要幅を確保して余りを等分。
 * 項目の間隔は詰まった方に合わせる（上限 maxGap）。
 */
function cardRow(s, cards, o) {
  const k = o.card;
  const need = cards.map(c => cardNeedW(c, k));
  const avail = BW - GAP * (cards.length - 1);
  const extra = avail - need.reduce((a, b) => a + b, 0);
  if (extra < -1e-9) throw new Error(`p${s.page}: カード幅が ${r3(-extra)}in 足りません`);
  const ws = need.map(v => Math.floor((v + extra / cards.length) * 1000) / 1000);
  const h = r3(o.bottom - o.y);
  const gap = Math.min(...cards.map(c => itemGapFor(c, h, o.maxGap, k)));
  let x = X0;
  const below = [];   // 最終行の枠の下端からカードの下端まで
  cards.forEach((c, i) => {
    const end = drawCard(s, c, x, o.y, ws[i], h, gap, k);
    below.push(r3(o.y + h - end - overhang(cardOpt(k).size)));
    x += ws[i] + GAP;
  });
  console.log(`  p${s.page} カード: 幅 ${ws.join(' / ')}in（必要 ${need.map(r3).join(' / ')}）、高さ ${h}in、項目間 ${gap}in、最終行の下 ${below.join(' / ')}in`);
}

/** 強調ボックス（白地＋色枠・太字・中央）。1 行は decklib の box（通常の行送り） */
const EMPH = { radius: 0.1, lineW: 3.5 };
function emphBox(s, o) {
  s.box({
    x: r3(o.x), y: r3(o.y), w: r3(o.w), h: o.h, fill: C.white, radius: EMPH.radius,
    line: { color: o.lineColor, width: EMPH.lineW },
    text: o.lines, size: o.size || 24, bold: true, color: C.dark,
  });
}
/** 行送りを詰めた強調ボックス（22pt・2 行）。下端 yBottom に合わせて置く。戻り値 {y, h} */
const EMPH_T = { padY: 0.04, padX: 0.07 };
function emphTightH(n, size) { return r3(2 * EMPH_T.padY + n * lineIn(size) + overhang(size)); }
function emphTight(s, o) {
  const size = o.size || 22;
  const h = emphTightH(o.lines.length, size);
  const y = o.yBottom - h;
  s.box({ x: r3(o.x), y: r3(y), w: r3(o.w), h, fill: C.white, radius: EMPH.radius, line: { color: o.lineColor, width: EMPH.lineW } });
  tightText(s, o.lines, { x: o.x + EMPH_T.padX, y: y + EMPH_T.padY, w: o.w - 2 * EMPH_T.padX, size, bold: true, align: 'center' });
  return { y, h };
}

/**
 * 行の多い表（3-11_local_anesthetics/build.js の tightTable を写したもの）。decklib の s.table と同じ見た目
 * （見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、上下中央）で、行間だけを固定値 LS pt（= 1.2 em）に詰める。
 * 見出し行も本文と同じ左揃え。セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、
 * 重なり検査用の登録も同じく行う。cell: 文字列（'\n' で改行）| {text, color, bold, align}
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

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める（文献は落とさない）
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 文言はそのままで、文献の途中の語の切れ目でも改行する（区切り > 和文の語の切れ目・「et al.」の後 >
//     誌名と年の間 > その他の欧文の空白 > 著者名の途中（読点の後）。数値と単位の間、「et」と「al.」の間では割らない）。
//  2) 入らなければ区切りの前後の空白を詰める（' / ' → '/'）。
//  3) それでも入らなければ（）内の補足を外す（外す数が最小の形。空白は詰めない形を優先）。
//  4) それでも入らないスライドだけ、extra（S6 の doi、S8 の頁の略記と和文の空白 1 か所）を当てる。
//  1)〜3) は改行位置の段（誌名と年の間まで → その他の欧文の空白まで → 著者名の途中まで）ごとに試し、
//  著者名の途中（「Douketis JD,」⏎「et al.」など）で割るよりは（）内の補足を外して区切りで割る方をとる。
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
function split2(str, sep, maxRank = Infinity) {
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
      else if (pa === ',') rank = 4;                                          // 読点（著者名の途中）
      else if (/[A-Za-z]/.test(pa) && isDigit(nb)) rank = 2;                 // 誌名と年の間
      else rank = 3;                                                          // その他の欧文の空白
    } else if ((str[i - 1] === '／' || str[i - 1] === '・') && str[i] !== ' ') {
      a = str.slice(0, i); b = str.slice(i); rank = 1;                        // 和文の中黒・斜線の後
    } else continue;
    a = a.trim(); b = b.trim();
    if (rank <= maxRank && a && b && srcFits(a) && srcFits(b)) cands.push({ a, b, rank });
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
/** extra: [[元, 後], ...]（（）を外すだけでは入らないスライドの追加の縮め方） */
function sourcesOf(id, extra = []) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  const items0 = spec.split(' / ');
  const lib = libLines(items0);
  if (lib && lib.length <= 2) return items0;
  const parens = [];
  items0.forEach((it, ii) => { for (const mm of it.matchAll(/（[^（）]*）/g)) parens.push({ ii, text: mm[0] }); });
  for (let ne = 0; ne <= extra.length; ne++) {
    for (const E of combos(extra, ne)) for (const maxRank of [2, 3, 4]) {
      for (let k = 0; k <= parens.length; k++) {
        // 同じ段階では区切りの空白を残す形と詰める形を比べ、改行位置のよい方（同順位なら空白を残す方）をとる
        let best = null;
        for (const sep of [' / ', '/']) {
          for (const P of combos(parens, k)) {
            const items = items0.map((it, ii) => {
              let t = it;
              for (const p of P) if (p.ii === ii) t = t.replace(p.text, '');
              for (const [from, to] of E) t = t.replace(from, to);
              return t;
            });
            const r = split2(items.join(sep), sep, maxRank);
            if (r && (!best || r.rank < best.r.rank)) best = { r, P, E, sep };
          }
        }
        {
          if (!best) continue;
          const sep = best.sep;
          const lines = best.r.b ? [best.r.a, best.r.b] : [best.r.a];
          const what = [];
          if (sep === '/') what.push('区切りの前後の空白を詰めた');
          if (best.P.length) what.push(`（）内の補足 ${best.P.map(p => p.text).join('・')} を外した`);
          for (const [from, to] of best.E) what.push(to ? `「${from}」→「${to}」` : `「${from}」を外した`);
          if (best.r.rank > 0) what.push('文献の途中の語の切れ目で改行');
          DEVIATIONS.push(`${id} 出典: 「${spec}」→「${lines.join(' ⏎ ')}」（16pt・2 行に収めるため。${what.join('、')}）`);
          return lines;
        }
      }
    }
  }
  throw new Error(`${id}: 出典が 2 行に収まりません`);
}

async function main() {
  const deck = new Deck({
    file: OUT,
    // spec の章チップ表記は「抗血小\n板薬」だが、「血小板」の途中で割らないよう改行位置だけ「抗血小板\n薬」にする
    // （1 行 4 字は「抗凝固薬」と同じ。チップの位置・寸法・文字サイズ・色はライブラリ既定のまま）。
    chapters: ['考え方', '抗凝固薬', '抗血小板\n薬', '区域\n麻酔', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-12 常用薬①：抗血栓薬',
  });
  if (!SPEC_MD.includes('「抗血小\\n板薬」')) throw new Error('spec の章チップ表記が変わりました（build.js の chapters を見直す）');
  DEVIATIONS.push('章チップ: 「抗血小\\n板薬」→「抗血小板\\n薬」（改行位置のみ。「血小板」の途中で割らないため。1 行 4 字は「抗凝固薬」と同じ）');
  const slide = o => deck.slide(o);   // 章チップはライブラリ既定（4 字ラベルの 10.5pt 調整も decklib 側）

  // ===== S1 表紙 =====
  {
    const title = '常用薬①：抗血栓薬', subtitle = '4-12 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes(`title=「${title}」`) || !SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) {
      throw new Error('S1: 表紙の文言が spec と違います');
    }
    deck.cover({ title, subtitle, meta, notes: notesOf('S1') });
  }

  // ===== S2 出血と血栓の天秤で決める（2列カード＋強調ボックス） =====
  // 右のカードは 5 項目が狭い幅で各 2 行（10 行）になるので本文ゾーンの高さいっぱいに置き、
  // 強調ボックスは左のカード（3 項目 7 行）の下に置く。
  {
    const id = 'S2';
    const s = slide({ title: '出血と血栓の天秤で決める', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    const c1 = {
      title: q(id, '手技の出血リスク（JCS 2020）'), titleColor: BLUE,
      items: [
        (() => { const l = brk(id, 'minor：抜歯、白内障、体表の小手術、生検なしの内視鏡 → 休薬しない（I）', '生検なし'); return [mark(l[0], ['minor：', null]), mark(l[1], ['休薬しない（I）', BLUE])]; })(),
        (() => { const l = brk(id, '低：関節鏡、乳房切除、経尿道手術など → 継続を考慮（IIa）', '経尿道'); return [mark(l[0], ['低：', null]), mark(l[1], ['継続を考慮（IIa）', BLUE])]; })(),
        (() => { const l = brk(id, '高：脊髄幹麻酔、開頭・脊椎、胸・腹・骨盤（鏡視下を含む）、大きな整形外科 → 中断（IIa）', '胸・腹', '大きな'); return [mark(l[0], ['高：', null]), l[1], mark(l[2], ['中断（IIa）', VER])]; })(),
      ],
    };
    const c2 = {
      title: q(id, '患者の血栓リスク（高い例）'), titleColor: VER,
      items: [
        brk(id, '機械弁（特に僧帽弁位）、リウマチ性僧帽弁狭窄', 'リウマチ'),
        brk(id, '3 か月以内の脳卒中・TIA・静脈血栓塞栓症', 'TIA'),
        brk(id, 'CHA2DS2-VASc が高い心房細動', '高い'),
        brk(id, '抗リン脂質抗体症候群などの血栓性素因', 'などの'),
        brk(id, '冠動脈ステント留置後の早い時期（→ 4-8）', '早い'),
      ],
    };
    const emLines = brk(id, '止血できる手技は休薬しない。休薬は最小限、止血を確認したらすぐ再開。処方医に相談する', '止血を確認');
    const w2 = r3(cardNeedW(c2));
    const x2 = r3(XR - w2);
    const wl = r3(x2 - GAP - X0);
    // 右: 本文ゾーンの高さいっぱい
    const h2 = r3(BOTTOM - TOP);
    drawCard(s, c2, x2, TOP, w2, h2, itemGapFor(c2, h2, 0.2));
    // 左下: 強調ボックス（22pt・2 行）
    const em = emphTight(s, {
      x: X0, w: wl, yBottom: BOTTOM, lineColor: TEAL,
      lines: [mark(emLines[0], ['止血できる手技は休薬しない。', BLUE]), mark(emLines[1], ['処方医に相談する', BLUE])],
    });
    // 左上: カード 1
    if (cardNeedW(c1) > wl + 1e-9) throw new Error(`${id}: 左のカードの幅不足`);
    const h1 = r3(em.y - GAP - TOP);
    const g1 = itemGapFor(c1, h1, 0.2);
    drawCard(s, c1, X0, TOP, wl, h1, g1);
    console.log(`  p${s.page} 左 ${wl}in（カード ${TOP}〜${r3(TOP + h1)}、項目間 ${g1}、強調 ${r3(em.y)}〜${BOTTOM}）、右 ${w2}in`);
    DEVIATIONS.push(`${id} 配置: 強調ボックスを全幅ではなく左のカードの下に置き、右のカードを本文ゾーンの高さいっぱいにした` +
      '（全幅の強調ボックスの上に 2 列カードを置くと、22pt では右のカード 5 項目が 9 行、左が 8 行になり入らないため。文言は変えていない）');
  }

  // ===== S3 ワルファリンとブリッジ（箇条書き5行＋強調ボックス） =====
  {
    const id = 'S3';
    const s = slide({ title: 'ワルファリンとブリッジ', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    // 行頭の「•」は本文の色（s.bullets だと行頭記号が先頭の run の色になり、強調色が混ざる）。
    // 5 行を一定の間隔で置き、強調ボックスとの間を空けすぎない。
    const BUL = { y: 1.95, size: 24, pitch: 0.7, ind: 0.383 };
    const bulItems = [
      mark(q(id, '5 日前に中止。手術前に INR が戻ったことを確かめる（1.5 未満）'), ['5 日前に中止', VER]),
      mark(q(id, '再開は術後 24 時間以内（止血を確認して）'), ['24 時間以内', BLUE]),
      // 同じ行に橙と赤を並べない（色覚多様性で区別しにくい）。警告色は橙 1 色にし、数値は太字で示す
      mark(q(id, '原則ブリッジしない。BRIDGE：血栓は減らず、大出血 1.3%→3.2%'), ['原則ブリッジしない', VER], ['1.3%→3.2%', null]),
      mark(q(id, '考慮：機械弁・リウマチ性僧帽弁狭窄・3 か月以内の脳梗塞など'), ['考慮：', BLUE]),
      mark(q(id, '機械弁のブリッジは指針で違う。循環器・心臓血管外科と決める'), ['循環器・心臓血管外科と決める', BLUE]),
    ];
    const lineH = heightFor(1, BUL.size, 0);
    bulItems.forEach((runs, i) => {
      const y = r3(BUL.y + i * BUL.pitch);
      s.text(['•'], { x: X0, y, w: BUL.ind, h: lineH, size: BUL.size });
      s.text([runs], { x: r3(X0 + BUL.ind), y, w: r3(BW - BUL.ind), h: lineH, size: BUL.size });
    });
    const bulBottom = BUL.y + (bulItems.length - 1) * BUL.pitch + lineH;
    const EH = 0.7;
    const ey = r3(BOTTOM - EH);
    if (ey < bulBottom + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(ey - bulBottom)}in`);
    emphBox(s, { x: X0, y: ey, w: BW, h: EH, lineColor: RED, size: 24,
      lines: [mark(q(id, '緊急手術：ビタミン K、プロトロンビン複合体製剤。INR を測って判断'), ['緊急手術：', RED])] });
  }

  // ===== S4 DOAC は 1〜2 日休薬（2列カード＋強調ボックス） =====
  // 左 8 行・右 8 行と 2 行の強調ボックスを 22pt で入れるため、カードの余白を詰め、項目の頭を太字にして
  // 項目の切れ目を示す（「•」を置く幅がない）。
  // カードの高さ（1.82〜5.34in）は 8 行で埋まり、項目間に間隔をとる余りがない（左は「ダビガトラン」の
  // CrCl 別の時間を 3 行より少なくできず、右は左に幅を譲ると 3 行に増える。総当たりで 8 行が下限）。
  // 余り（約 0.11in）は、実機メイリオで最終行がカードの縁に寄らないよう、すべて最終行の下に回す。
  // 右のカードの太字は「主文を太字・（）内の補足と後に続く文は通常」にそろえる。
  {
    const id = 'S4';
    const s = slide({ title: 'DOAC は 1〜2 日休薬', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    const K = { padX: 0.08, padY: 0.05, titleGap: 0.02, padB: 0.1 };
    const emLines = brk(id, 'PAUSE：30 日の大出血 0.90〜1.85%、動脈血栓塞栓症 0.16〜0.60%。ブリッジなし・凝固検査なしで低率だった', 'ブリッジなし');
    const em = emphTight(s, { x: X0, w: BW, yBottom: BOTTOM, lineColor: TEAL,
      lines: [mark(emLines[0], ['PAUSE：', BLUE]), mark(emLines[1], ['低率だった', BLUE])] });
    cardRow(s, [
      {
        title: q(id, '休薬する場合（出血リスク別）'), titleColor: BLUE, marker: false,
        items: [
          [mark(q(id, '低リスク：1 日休薬、翌日に再開'), ['低リスク：', null])],
          (() => { const l = brk(id, '高リスク：2 日休薬（最終内服から約 60 時間。JCS は 48 時間以上）、2〜3 日後に再開', '（最終', 'JCS'); return [mark(l[0], ['高リスク：', null]), l[1], l[2]]; })(),
          (() => { const l = brk(id, 'ダビガトラン（高リスク、JCS 2020）：CrCl 80 以上 48 時間、50〜79 で 72 時間、30〜49 で 96 時間以上', 'CrCl', '30〜49'); return [mark(l[0], [l[0], null]), l[1], l[2]]; })(),
          [mark(q(id, 'ブリッジも凝固検査も要らない'), ['ブリッジも凝固検査も要らない', VER])],
        ],
      },
      {
        title: q(id, '落とし穴'), titleColor: VER, marker: false,
        items: [
          (() => { const l = brk(id, '『2 日休薬』と『48 時間』は違う。最終内服の日時で数える', '違う'); return [mark(l[0], [l[0], null]), mark(l[1], ['違う。', null])]; })(),
          (() => { const l = brk(id, '腎機能は eGFR ではなく CrCl（Cockcroft-Gault）', '（Cockcroft'); return [mark(l[0], [l[0], null]), l[1]]; })(),
          (() => { const l = brk(id, '脊髄幹麻酔なら 72 時間（→ S6）', '（→'); return [mark(l[0], [l[0], null]), l[1]]; })(),
          (() => { const l = brk(id, '添付文書の記載（24/48 時間など）も確かめる', '（24'); return [mark(l[0], [l[0], null]), mark(l[1], ['も確かめる', null])]; })(),
        ],
      },
    ], { y: TOP_T, bottom: r3(em.y - GAP), maxGap: 0, card: K });
  }

  // ===== S5 抗血小板薬（表 3列 × 4行＋見出し行、下に1行） =====
  // spec の列幅（3.0 / 4.4 / 4.8）では「DAPT 中の待機手術は延期が第一。」が 1 行に入らないため、
  // 1 列目を「クロピドグレル」、2 列目を「（添付文書は 14 日以上）」の幅まで詰めて 3 列目を広げる（合計 12.23in）。
  {
    const id = 'S5';
    const s = slide({ title: '抗血小板薬', chapter: 2, notes: notesOf(id), sources: sourcesOf(id) });
    const drug = text => ({ text, bold: true });
    // 合計 12.23in（本文ゾーンの全幅）。3 列は順に「クロピドグレル」「（添付文書は 14 日以上）」「DAPT 中の待機手術は延期が第一。」が
    // 1 行に入る最小の幅で、合計 12.2in には収まらない（どれかを 2 行にすると表の下の 1 行が本文ゾーンからはみ出す）。
    const S5_COLW = [2.565, 4.18, 5.485];
    const t = tightTable(s, [
      [q(id, '薬'), q(id, '休薬の目安（手術）'), q(id, '注意')],
      [drug(brk(id, 'アスピリン（低用量）', '（低用量）').join('\n')), { text: brk(id, '原則続ける。特にステント後', '特に').join('\n'), color: BLUE, bold: true }, brk(id, '頭蓋内・脊柱管内は別。一次予防なら処方医と中止も', '一次予防').join('\n')],
      [drug(q(id, 'クロピドグレル')), brk(id, '5 日前（添付文書は 14 日以上）', '（添付').join('\n'), brk(id, 'DAPT 中の待機手術は延期が第一。PCI 施行医に相談', 'PCI').join('\n')],
      [drug(q(id, 'プラスグレル')), brk(id, '7 日前（添付文書は 14 日以上）', '（添付').join('\n'), brk(id, '国内用量でも目安は 7 日（海外データ）', '（海外').join('\n')],
      [drug(q(id, 'チカグレロル')), q(id, '3〜5 日前'), brk(id, '可逆的。DAPT 中は PCI 施行医に相談', 'DAPT').join('\n')],
    ], { x: X0, y: TOP_T, colW: S5_COLW });
    console.log(`  p${s.page} 表: 行高 ${t.rowH.join(' / ')}in、下端 ${r3(t.y + t.h)}in`);
    DEVIATIONS.push(`${id} 列幅: spec の目安 3.0 / 4.4 / 4.8in ではなく ${S5_COLW.join(' / ')}in（合計 12.23in）にした。` +
      'spec の幅では 3 列目の「DAPT 中の待機手術は延期が第一。」（22pt で 5.32in）が 1 行に入らず、その行が 3 行になって' +
      '表の下の 1 行が本文ゾーン（〜6.44in）からはみ出すため。1 列目は最長の「クロピドグレル」、2 列目は「（添付文書は 14 日以上）」が入る幅');
    const fh = heightFor(1, 24, 0);
    const fy = r3(t.y + t.h + GAP + 0.1);
    if (fy + fh > BOTTOM + 1e-9) throw new Error(`${id}: 表の下の 1 行の下端 ${r3(fy + fh)} > ${BOTTOM}`);
    s.text([mark(q(id, 'NSAIDs・SSRI・抗凝固薬の併用で出血は増え、単剤の表は使えない'), ['単剤の表は使えない', VER])],
      { x: X0, y: fy, w: BW, h: fh, size: 24 });
  }

  // ===== S6 区域麻酔は長く休む（表 3列 × 7行＋見出し行） =====
  // 22pt では tightTable・列幅の見直し後も高さ 6.9in になり本文ゾーン（4.62in）に入らない。数値・薬剤名を
  // 落とさずに短くしても 6.2in。文言は変えず、表の文字だけ 18pt にする（deviations に記載）。
  {
    const id = 'S6';
    const s = slide({ title: '区域麻酔は長く休む', chapter: 3, notes: notesOf(id),
      sources: sourcesOf(id, [['. doi:10.1136/rapm-2024-105766', '']]) });
    const drug = text => ({ text, bold: true });
    const key = text => ({ text, bold: true, color: VER });
    const TS = 18;
    // 列幅（合計 12.23in。S5 の表と右端をそろえる）: 2 列目は見出し「脊髄幹麻酔・深部ブロックの前（ASRA 第 5 版）」、3 列目は「24 時間後（国内は高用量）」が
    // 1 行に入る最小の幅（推定器ちょうど。実機メイリオの字幅ではセル幅の 8〜9 割）にし、残りをすべて 1 列目に回す。
    // どちらかを 2 行にするとその行が 1 行ぶん（0.3in）高くなり、表が本文ゾーン（1.82〜6.44in）に入らない。
    const S6_COLW = [2.34, 6.245, 3.645];
    const t = tightTable(s, [
      [q(id, '薬'), q(id, '脊髄幹麻酔・深部ブロックの前（ASRA 第 5 版）'), q(id, 'カテーテル抜去後の再開')],
      [drug(q(id, 'ワルファリン')), q(id, '約 5 日。INR 正常化'), brk(id, '抜去は INR 1.5 未満。抜去後すぐ', '抜去後').join('\n')],
      [drug(q(id, '抗 Xa 薬 高用量')), key(q(id, '72 時間以上')), q(id, '24 時間後')],
      [drug(brk(id, '抗 Xa 薬 低用量（VTE 予防量）', '（VTE').join('\n')), brk(id, 'リバーロキサバン 24 時間、アピキサバン 36 時間', 'アピキサバン').join('\n'), q(id, '6 時間後')],
      [drug(q(id, 'ダビガトラン')), q(id, 'CrCl 50 以上 72 時間、30〜49 で 120 時間'), q(id, '24 時間後（国内は高用量）')],
      [drug(q(id, 'LMWH')), q(id, '低用量 12 時間、高用量 24 時間'), q(id, '4 時間後')],
      [drug(q(id, '未分画ヘパリン')), brk(id, '皮下 低用量 4〜6 時間、高用量 12〜24 時間。静注 4〜6 時間＋aPTT 正常', '静注').join('\n'), q(id, '1 時間後')],
      [drug(q(id, '抗血小板薬')), brk(id, 'アスピリン・NSAIDs は不要。クロピドグレル 5〜7 日、チカグレロル 5 日、プラスグレル 7〜10 日', 'クロピドグレル', 'プラスグレル').join('\n'), brk(id, '負荷なし：すぐ、負荷あり：6 時間後', '負荷あり').join('\n')],
    ], { x: X0, y: TOP_T, colW: S6_COLW, size: TS });
    console.log(`  p${s.page} 表: 行高 ${t.rowH.join(' / ')}in、下端 ${r3(t.y + t.h)}in`);
    DEVIATIONS.push(`${id} 列幅: spec の目安 3.2 / 5.6 / 3.4in ではなく ${S6_COLW.join(' / ')}in（合計 12.23in。S5 と右端をそろえた）にした。` +
      'spec の 2 列目 5.6in では見出し「脊髄幹麻酔・深部ブロックの前（ASRA 第 5 版）」、皮下 UFH の行、抗血小板薬の行、' +
      'ダビガトランの行が 18pt でも 1 行に入らず、表が 4 行ぶん（約 1.2in）高くなって本文ゾーンに入らないため。' +
      '2 列目は見出し（最長）が、3 列目は「24 時間後（国内は高用量）」が 1 行に入る最小の幅とし、残りを 1 列目に回した' +
      '（1 列目の「抗 Xa 薬 高用量」は推定でセル幅の約 9 割。1 列目を 2.8in 以上にする余地はない）');
    if (t.y + t.h > BOTTOM + 0.011) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)} > ${BOTTOM}`);
    DEVIATIONS.push(`${id} 文字の大きさ: 表（見出し行＋7 行）を 22pt ではなく ${TS}pt で組んだ（spec の文言は一字も変えていない）。` +
      '22pt では tightTable（行間 1.2 em・余白 [3,6,3,6]pt）と列幅の見直し（合計 12.23in）をしても表の高さが約 6.9in で、' +
      '本文ゾーン（1.82〜6.44in、4.62in）に入らない。数値・薬剤名を残したまま語を削っても約 6.2in で入らず、' +
      '数値を欠いた区域麻酔の休薬表は誤読の危険がある。spec 側で S6 を 2 枚に分けるか、表の内容を減らす必要がある');
  }

  // ===== S7 相談先と緊急時（2列カード＋強調ボックス） =====
  {
    const id = 'S7';
    const s = slide({ title: '相談先と緊急時', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    const EH = 0.66;
    const ey = r3(BOTTOM - EH);
    emphBox(s, { x: X0, y: ey, w: BW, h: EH, lineColor: TEAL, size: 24,
      lines: [mark(q(id, '休薬・再開の判断と相談した相手を記録に残し、患者にも説明する'), ['記録に残し', VER])] });
    cardRow(s, [
      {
        title: q(id, '誰に相談するか'), titleColor: BLUE,
        items: [
          [q(id, '抗凝固薬の処方医（診療所・内科）')],
          (() => { const l = brk(id, 'ステント後：PCI 施行医・循環器内科', 'PCI'); return [mark(l[0], [l[0], null]), l[1]]; })(),
          [mark(q(id, '機械弁：循環器内科・心臓血管外科'), ['機械弁：', null])],
          (() => { const l = brk(id, '3 か月以内の静脈血栓塞栓症：延期を主科と', '延期'); return [mark(l[0], [l[0], null]), l[1]]; })(),
          [mark(q(id, '迷ったら上級の麻酔科医に電話'), ['上級の麻酔科医に電話', BLUE])],
        ],
      },
      {
        title: q(id, '緊急手術'), titleColor: VER,
        items: [
          brk(id, '休薬表は使えない。最終内服の時刻を聞く', '最終'),
          (() => { const l = brk(id, '中和薬の在庫：ビタミン K、PCC、イダルシズマブ', 'ビタミン'); return [mark(l[0], [l[0], null]), l[1]]; })(),
          [q(id, '脊髄幹麻酔を避けて全身麻酔に')],
          [q(id, '血液製剤の準備（→ 4-4）')],
          [q(id, 'DOAC は CrCl で残り方が違う')],
        ],
      },
    ], { y: TOP, bottom: r3(ey - GAP), maxGap: 0.16, card: {} });
  }

  // ===== S8 術前指示の書き方（大きなボックス1つに指示例4行＋下に1行） =====
  // 指示例 4 項目は 22pt で 9 行。行送り 1.2 em で箱に入れ、項目の 2 行目以降は 1 字ぶん下げる。
  // 患者 A（項目 1・2）と患者 B（項目 3・4）の間を少し広く空け、細い線で区切る。
  // 日付の「◯/◯」（U+25EF。spec の文言のまま）は字面が大きく、上の行に寄るので、その行の上だけ 0.04in 広げる。
  {
    const id = 'S8';
    // 出典は（）内の補足を外し区切りの空白を詰めても 2 行に入らないため、頁の範囲を略記し、和文の題の空白を 1 か所詰める
    // （「2020 年」の前後の空白は他のスライドと同じに残す）。
    const s = slide({ title: '術前指示の書き方', chapter: 3, notes: notesOf(id),
      sources: sourcesOf(id, [
        ['2020 年改訂版 不整脈薬物治療ガイドライン', '2020 年改訂版不整脈薬物治療ガイドライン'],
        ['2019;179:1469-1478', '2019;179:1469-78'],   // 頁の範囲の略記（Vancouver 式。情報は落ちない）
      ]) });
    const lab = BLUE;
    const A = brk(id, '患者 A（心房細動）エリキュース 5 mg 1 日 2 回：最終内服は手術 4 日前の夜（◯/◯ 20 時）。脊髄くも膜下麻酔の予定のため 72 時間以上あける（9 時入室で約 85 時間）。ブリッジなし', '手術 4 日前', '72 時間以上');
    const Bl = q(id, '再開：術後 2 日目の夜、止血を確認してから。処方医（◯◯内科）に連絡済み');
    const Cl = q(id, '患者 B（ステント留置 2 年前）バイアスピリン 100 mg：継続');
    const D = brk(id, 'クロピドグレル 75 mg：PCI 施行医（循環器 ◯◯医師）と相談し、最終内服は手術 7 日前の朝（◯/◯ 8 時）。脊髄くも膜下麻酔のため 7 日以上あける。再開は止血を確認してから（負荷投与なら穿刺から 6 時間以上後）', '最終内服', '脊髄くも膜下', '（負荷');
    const items = [
      [mark(A[0], ['患者 A（心房細動）', lab], ['エリキュース 5 mg 1 日 2 回', null]), A[1], mark(A[2], ['72 時間以上', VER])],
      [mark(Bl, ['再開：', null])],
      [mark(Cl, ['患者 B（ステント留置 2 年前）', lab], ['バイアスピリン 100 mg', null])],
      [mark(D[0], ['クロピドグレル 75 mg', null]), D[1], mark(D[2], ['7 日以上', VER]), D[3]],
    ];
    const PAD = { x: 0.09, y: 0.1 }, IND = 0.34, G_IN = 0.06, G_GROUP = 0.17, G_MARU = 0.04;
    const L = lineIn(22);
    const nLines = items.reduce((a, it) => a + it.length, 0);
    const hasMaru = l => runText(l).includes('◯/◯');
    const nMaru = items.reduce((a, it) => a + it.filter((l, j) => j > 0 && hasMaru(l)).length, 0);
    if (nMaru !== 2) throw new Error(`${id}: 「◯/◯」の行が ${nMaru} 行（2 行のはず）`);
    const boxH = r3(2 * PAD.y + nLines * L + 2 * G_IN + G_GROUP + nMaru * G_MARU);
    const footH = heightFor(1, 24, 0);
    const y0 = TOP_T;
    if (y0 + boxH + GAP + footH > BOTTOM + 1e-9) throw new Error(`${id}: 箱 ${boxH} + 下の 1 行が ${BOTTOM} を超えます`);
    s.box({ x: X0, y: y0, w: BW, h: boxH, fill: C.bg, radius: 0.08 });
    const iw = BW - 2 * PAD.x;
    let y = y0 + PAD.y;
    items.forEach((lines, i) => {
      lines.forEach((l, j) => {
        if (j > 0 && hasMaru(l)) y += G_MARU;
        const dx = j === 0 ? 0 : IND;
        tightText(s, [l], { x: X0 + PAD.x + dx, y, w: iw - dx, size: 22 });
        y += L;
      });
      if (i === 1) {
        // 患者 A と患者 B の間の区切り線
        s.hline({ x: X0 + PAD.x, y: r3(y + G_GROUP / 2), w: r3(iw), color: C.border, width: 1 });
        y += G_GROUP;
      } else if (i < items.length - 1) y += G_IN;
    });
    const fy = r3(y0 + boxH + GAP + (BOTTOM - (y0 + boxH + GAP) - footH) / 2);
    s.text([mark(q(id, '『◯日前から中止』ではなく、最終内服の日時と再開の条件を書く'), ['最終内服の日時と再開の条件', VER])],
      { x: X0, y: fy, w: BW, h: footH, size: 24, bold: true, align: 'center' });
    console.log(`  p${s.page} 指示例: 箱 ${y0}〜${r3(y0 + boxH)}in（${nLines} 行）、下の 1 行 ${fy}〜${r3(fy + footH)}in`);
  }

  // ===== S9 まとめ（箇条書き5行） =====
  {
    const id = 'S9';
    const s = slide({ title: 'まとめ', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    const items = [
      mark(q(id, '休薬は出血リスク × 血栓リスクで決める。止血できる手技は続ける'), ['出血リスク × 血栓リスク', BLUE]),
      mark(q(id, 'ワルファリン 5 日、DOAC 1〜2 日。ブリッジは原則しない'), ['ブリッジは原則しない', VER]),
      mark(q(id, 'アスピリンは原則続ける。P2Y12 阻害薬の中止は PCI 施行医と'), ['PCI 施行医と', BLUE]),
      mark(q(id, '区域麻酔は手術より長い：高用量 DOAC 72 時間（ASRA 第 5 版）'), ['高用量 DOAC 72 時間', VER]),
      // 警告色は橙 1 色にそろえる（2 行目の橙と赤が色覚多様性で区別しにくいため）
      mark(q(id, '最終内服の日時・再開の条件・相談相手を記録する。緊急は中和薬'), ['緊急は中和薬', VER]),
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
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
