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
 * 組版の方針（文言を変えずに 22pt 以上で収める）:
 *  - カード・強調ボックスの本文は 22pt 以上のまま行送りを 1.2 em にする（1 行 1 枠の s.text を 1.2 em おきに置く）。
 *  - 表は 3-11 の tightTable（行間 1.2 em、余白 [3,6,3,6]pt、22pt）を写して使う（行内の部分装飾にだけ対応を足した）。
 *    セルの余白は左右の配分だけを [3,3,3,9]pt に変える（左 0.125in。合計は同じなのでセル幅の検査は変わらない）。
 *    列幅は spec の値（目安）で入らないときだけ、語句の切れ目で改行して本文ゾーンに入る幅に見直す（合計 12.2in）。
 *  - 本文ゾーン: x 0.55〜12.78in、y 1.85in〜（出典 2 行の文字の上端 約 6.68in の 0.24in 上 = 6.44in）。
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
  if (/【要確認/.test(str)) throw new Error(`${id}: 【要確認】の注記はスライド面に出さない`);
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
/**
 * 本文の下端。出典が 2 行（枠 y6.5・下揃え。PowerPoint のメイリオ 16pt 2 行で文字の上端 約6.68in）なら
 * その 0.24in 上の 6.44in。出典が 1 行（y6.85）のスライドでも同じ値にそろえる。
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
// カード（S4・S9）: 角丸の塗り箱＋見出し＋「•」付きの項目（文字はすべて箱の内側）。
// 項目は語句の切れ目で 2 行以上に割り、2 行目以降は記号の右にぶら下げる。
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
  return end;
}

/** 行送りを詰めた強調ボックス（白地＋色枠・太字・中央）。下端 yBottom に合わせて置く。戻り値 {y, h} */
const EMPH = { radius: 0.1, lineW: 3.5, padY: 0.06, padX: 0.07 };
function emphTightH(n, size) { return r3(2 * EMPH.padY + n * lineIn(size) + overhang(size)); }
function emphTight(s, o) {
  const size = o.size || 24;
  const h = r3(Math.max(emphTightH(o.lines.length, size), o.minH || 0));
  const y = r3(o.yBottom - h);
  s.box({ x: r3(o.x), y, w: r3(o.w), h, fill: C.white, radius: EMPH.radius, line: { color: o.lineColor, width: EMPH.lineW } });
  const th = emphTightH(o.lines.length, size) - 2 * EMPH.padY;
  tightText(s, o.lines, { x: o.x + EMPH.padX, y: y + (h - th) / 2, w: o.w - 2 * EMPH.padX, size, bold: true, align: 'center' });
  return { y, h };
}

/**
 * 行の多い表（3-11_local_anesthetics/build.js の tightTable を写したもの）。decklib の s.table と同じ見た目
 * （見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、上下中央）で、行間だけを固定値 LS pt（= 1.2 em）に詰める。
 * 見出し行も本文と同じ左揃え。セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、
 * 重なり検査用の登録も同じく行う。
 * cell: 文字列（'\n' で改行）| {text: 文字列 | 行の配列（行は文字列か mark() の run 配列）, color, bold, align}
 * （3-11 からの変更は、行内の部分装飾（run 配列）を受け付けるようにしたことだけ）
 */
function tightTable(s, rows, o) {
  const size = o.size || 22, LS = o.lineSpacing || size * 1.2;
  const m = o.margin || [3, 6, 3, 6];   // pt: 上, 右, 下, 左
  const colW = o.colW, nCol = colW.length;
  const w = colW.reduce((a, b) => a + b, 0);
  if (w > TABLE_W + 1e-6) throw new Error(`p${s.page} 表の幅 ${r3(w)}in > ${TABLE_W}in`);
  const rowH = [];
  const out = rows.map((row, ri) => {
    if (row.length !== nCol) throw new Error(`p${s.page} 表 行${ri + 1}: 列数 ${row.length} ≠ ${nCol}`);
    const head = ri === 0;
    let maxLines = 1;
    const cells = row.map((cell, ci) => {
      const c = typeof cell === 'string' ? { text: cell } : cell;
      const lines = Array.isArray(c.text) ? c.text : c.text.split('\n');
      const avail = colW[ci] - (m[1] + m[3]) / 72;
      lines.forEach(l => {
        const fi = fitInfo(runText(l), size, avail);
        if (!fi.ok) {
          throw new Error(`p${s.page} 表 行${ri + 1}列${ci + 1}「${runText(l)}」がセル幅に収まりません` +
            `（推定 ${fi.needIn}in > ${avail.toFixed(3)}in、約${fi.overChars}字オーバー）`);
        }
      });
      maxLines = Math.max(maxLines, lines.length);
      const color = c.color || C.dark, bold = c.bold != null ? c.bold : head;
      const runs = [];
      lines.forEach((l, li) => {
        const rl = (Array.isArray(l) ? l : [l]).map(r => (typeof r === 'string' ? { text: r } : r));
        rl.forEach((r, k) => runs.push({
          text: r.text,
          options: {
            fontFace: FONT, lang: 'ja-JP', fontSize: size, color: r.color || color,
            bold: (r.bold != null ? r.bold : bold) || undefined,
            lineSpacing: LS, breakLine: (k === rl.length - 1 && li < lines.length - 1) || undefined,
          },
        }));
      });
      return {
        text: runs,
        options: {
          fill: head ? { color: C.lightGray } : undefined,
          align: c.align || 'left', valign: 'middle',
          fontFace: FONT, fontSize: size, color, bold: bold || undefined,
        },
      };
    });
    rowH.push(Math.ceil((maxLines * LS / 72 + (m[0] + m[2]) / 72) * 100 - 1e-6) / 100 + (ri > 0 ? (o.extraRowH || 0) : 0));
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
 * 表を置いて、本文ゾーンの下端を超えないことを確かめる（ログも出す）。
 * 本文ゾーンに余りが大きいとき（行の少ない表）は、本文の行の高さを少しずつ足して（1 行あたり最大 0.16in。
 * 行間 1.2 em とセルの余白はそのまま、文字は上下中央）、表の下に 0.3in 程度の余白を残す。
 */
function rowLines(rows) {
  return rows.map(r => Math.max(...r.map(c => {
    const t = typeof c === 'string' ? c : c.text;
    return Array.isArray(t) ? t.length : t.split('\n').length;
  })));
}
function placeTable(s, id, rows, o) {
  const m = o.margin || [3, 6, 3, 6];
  const natural = rowLines(rows).reduce((a, n) => a + Math.ceil((n * 26.4 / 72 + (m[0] + m[2]) / 72) * 100 - 1e-6) / 100, 0);
  const room = BOTTOM - TOP - natural - 0.3;
  const extraRowH = room > 0 ? Math.min(0.16, Math.floor(room / (rows.length - 1) * 100) / 100) : 0;
  const t = tightTable(s, rows, { x: X0, y: TOP, extraRowH, ...o });
  console.log(`  p${s.page} 表: 列幅 ${o.colW.join(' / ')}in、行高 ${t.rowH.map(r3).join(' / ')}in（本文行 +${extraRowH}in）、${TOP}〜${r3(t.y + t.h)}in`);
  if (t.y + t.h > BOTTOM + 1e-6) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)} > ${BOTTOM}`);
  return t;
}

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める（文献は落とさない）
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 文言はそのままで、文献の途中の語の切れ目でも改行する（区切り > 和文の語の切れ目・「et al.」の後 >
//     読点の後・誌名と年の間 > その他の空白。数値と単位の間、「et」と「al.」の間では割らない）。
//  2) 入らなければ区切りの前後の空白を詰める（' / ' → '/'）。
//  3) それでも入らなければ（）内の補足を外す（外す数が最小の形）。
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
      if (isDigit(pa) && isCJK(nb)) continue;                                 // 数値と単位
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
  const done = (best) => {
    const lines = best.r.b ? [best.r.a, best.r.b] : [best.r.a];
    const what = [];
    if (best.sep === '/') what.push('区切りの前後の空白を詰めた');
    if (best.P && best.P.length) what.push(`（）内の補足 ${best.P.map(p => p.text).join('・')} を外した`);
    if (best.r.rank > 0) what.push('文献の途中の語の切れ目で改行');
    // 改行位置だけの調整（文言そのまま）は deviations に入れない
    if (best.sep === '/' || (best.P && best.P.length)) {
      DEVIATIONS.push(`${id} 出典: 「${spec}」→「${lines.join(' ⏎ ')}」（16pt・2 行に収めるため。${what.join('、')}）`);
    } else {
      console.log(`  ${id} 出典: ${what.join('、')}`);
    }
    return lines;
  };
  // 同じ段階では区切りの空白を残す形と詰める形を比べ、改行位置のよい方（同順位なら空白を残す方）をとる
  const bestOf = (variants) => {
    let best = null;
    for (const sep of [' / ', '/']) {
      for (const v of variants) {
        const r = split2(v.items.join(sep), sep);
        if (r && (!best || r.rank < best.r.rank)) best = { r, P: v.P, sep };
      }
    }
    return best;
  };
  let best = bestOf([{ items: items0, P: [] }]);
  if (best) return done(best);
  for (let k = 1; k <= parens.length; k++) {
    best = bestOf(combos(parens, k).map(P => ({
      P,
      items: items0.map((it, ii) => {
        let t = it;
        for (const p of P) if (p.ii === ii) t = t.replace(p.text, '');
        return t;
      }),
    })));
    if (best) return done(best);
  }
  throw new Error(`${id}: 出典が 2 行に収まりません`);
}

/**
 * 表のセルの余白（pt: 上, 右, 下, 左）。tightTable 既定の [3,6,3,6] から左右の配分だけを変え、左を 9pt（0.125in）に
 * して 1 文字目を左の罫線から離す（左右の合計 12pt は同じなので、セル幅の検査・列幅・改行は変わらない）。
 * 文字は左揃えで、推定器（全角 ×1.12）より実際のメイリオは 1 割ほど狭いので、右 3pt でも右の罫線には届かない。
 */
const TABLE_MARGIN = [3, 3, 3, 9];

// 表のセルの書式
const drug = text => ({ text, bold: true });                       // 1 列目（薬）
const keep = text => ({ text, bold: true, color: BLUE });          // 続ける・飲む
const stop = text => ({ text, bold: true, color: VER });           // 止める・休薬
const cell = lines => ({ text: lines });                           // 行の配列（mark() の run 可）

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
  // 箇条書き 24pt、強調ボックスは 24pt 太字 2 行（行送り 1.2 em）。箇条書きの枠は強調ボックスの 0.24in 上までとり、
  // 上下中央に置く。
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
    const em = brk(id, 'どちらでもよいが、続けるなら導入時の低血圧に備える。止めたら再開日を書く', '止めたら');
    const e = emphTight(s, { x: X0, w: BW, yBottom: BOTTOM, lineColor: TEAL, size: 24,
      lines: [em[0], mark(em[1], ['止めたら再開日を書く', VER])] });
    const BUL = { y: TOP, size: 24, gap: 5 };
    const need = heightFor(items.length, BUL.size, BUL.gap);
    const bh = r3(e.y - GAP - BUL.y);
    if (need > bh + 1e-9) throw new Error(`${id}: 箇条書きの高さ ${need} > ${bh}`);
    s.bullets(items, { ...BUL, h: bh, valign: 'middle' });
    console.log(`  p${s.page} 箇条書き枠 ${BUL.y}〜${r3(BUL.y + bh)}in（必要 ${need}in、上下中央）、強調ボックス ${e.y}〜${r3(e.y + e.h)}in`);
  }

  // ===== S3 ほかの降圧薬とスタチン（表 3列 × 5行＋見出し行） =====
  // spec の列幅（2.8 / 3.6 / 5.8）では行 1・行 5 の理由が 4 行、行 3・行 4 が 2 行に折れて計 14 行・約 5.6in になり、
  // 本文ゾーン（1.85〜6.44in、4.59in）に入らない。22pt・行送り 1.2 em・上下余白 3pt で入るのは 11 行（約 4.53in）まで。
  // 行 1・行 5 の理由（約 51 字）はどう割っても 3 行、行 3 の理由も 2 行なので、行 4 の理由
  // 「急な中止でリバウンド高血圧。新規開始は無益（3）」を 1 行にするしかなく、3 列目は 8.06in＋余白が要る。
  // 残る 4.1in を 1 列目（「クロニジン」1.71in）と 2 列目に分ける。2 列目は「続ける（I）。」が 1 行に入る 2.25in をとり、
  // スタチンの当日を「続ける（I）。／当日朝も」の 2 行にする（「（I）。」だけの行を作らない）。
  // そのため左右の余白は 3 列で計 13pt しかとれず、この表だけ余白を [3,0,3,4]pt にする（左 4pt は従来どおり、右 0pt）。
  // 推定器より実際のメイリオは 1 割ほど狭いので、右 0pt でも文字は右の罫線に届かない。
  {
    const id = 'S3';
    const s = slide(id, 0);
    const rows = [
      [q(id, '薬'), q(id, '当日'), q(id, '理由')],
      [drug(q(id, 'β遮断薬')), keep(q(id, '続ける（I）')),
        cell(brk(id, '周術期の中止で 30 日死亡のオッズ比 3.93（観察研究）。当日の新規開始は害（POISE）。始めるなら 7 日以上前', '（観察研究）', '始めるなら'))],
      [drug(q(id, 'Ca 拮抗薬')), keep(q(id, '続ける')), q(id, '中止の利益はない')],
      [drug(q(id, '利尿薬')), stop(brk(id, '当日は休薬が一般的', '一般的')), cell(brk(id, '脱水・低血圧・腎機能。溢水のある心不全では続ける', '溢水'))],
      [drug(q(id, 'クロニジン')), keep(q(id, '続ける')), q(id, '急な中止でリバウンド高血圧。新規開始は無益（3）')],
      [drug(q(id, 'スタチン')), keep(brk(id, '続ける（I）。当日朝も', '当日朝も')),
        cell(brk(id, '術後 4 日超の中断で心筋壊死のオッズ比 2.9（大動脈手術）。ASCVD または 10 年リスクで適応があれば開始（I）', '（大動脈', 'ASCVD'))],
    ];
    const m = [3, 0, 3, 4];
    const colW = [1.77, 2.31, 8.12];
    placeTable(s, id, rows, { colW, margin: m });
    DEVIATIONS.push(`${id} 表: 列幅を spec の 2.8 / 3.6 / 5.8in から ${colW.join(' / ')}in に、セルの余白を [3,6,3,6]pt から ${'[' + m.join(',') + ']'}pt にした` +
      '（文言・22pt・行送り 1.2 em はそのまま）。spec の列幅では計 14 行・約 5.6in で本文ゾーン（4.59in）に入らない。' +
      '入るのは 11 行までで、そのためには行 4 の理由を 1 行にする 3 列目の幅（8.06in＋余白）が要る。' +
      '残りで 2 列目に「続ける（I）。」が 1 行に入る幅をとると、左右の余白は 3 列で計 13pt まで');
  }

  // ===== S4 ステロイドカバー（2列カード + 強調ボックス） =====
  // カード本文 22pt・行送り 1.2 em。右のカードの 3 項目めは文の切れ目（「持続。」の後）で割り、2 行目
  // 「食べられたら通常量の 2 倍を 48 時間経口で」が入る幅を右のカードに残す。左のカードは
  // 「過去 6 か月に経口 2 週超・」が 1 行に入る幅にし、「長期の高用量 ICS（GINA）」を 2 行目にまとめる。
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
        brk(id, '過去 6 か月に経口 2 週超・長期の高用量 ICS（GINA）', '長期'),
      ],
    };
    const c2 = {
      title: q(id, 'どう投与するか（Woodcock 2020）'), titleColor: VER,
      items: [
        (() => { const l = brk(id, '普段のステロイドは当日朝も飲む。カバーは『上乗せ』', 'カバーは'); return [l[0], mark(l[1], ['『上乗せ』', VER])]; })(),
        [q(id, '導入時にヒドロコルチゾン 100 mg 静注')],
        brk(id, '続いて 200 mg/24 時間を持続。食べられたら通常量の 2 倍を 48 時間経口で', '食べられたら'),
        brk(id, '小手術に一律 100 mg は異論あり（→ 3-12）', '異論'),
      ],
    };
    const w1 = Math.ceil(cardNeedW(c1) * 1000 + 1) / 1000;   // 推定幅 + 0.001in（丸めで幅不足にしない）
    const w2 = r3(BW - GAP - w1);
    if (cardNeedW(c2) > w2 + 1e-9) throw new Error(`${id}: 右のカードの幅不足（必要 ${r3(cardNeedW(c2))} > ${w2}）`);
    const y = TOP, h = r3(e.y - GAP - y);
    const g1 = itemGapFor(c1, h, 0.2), g2 = itemGapFor(c2, h, 0.2);
    drawCard(s, c1, X0, y, w1, h, g1);
    drawCard(s, c2, X0 + w1 + GAP, y, w2, h, g2);
    console.log(`  p${s.page} カード: 幅 ${w1} / ${w2}in、${y}〜${r3(y + h)}in、項目間 ${g1} / ${g2}in、強調ボックス ${e.y}〜${r3(e.y + e.h)}in`);
  }

  // ===== S5 神経・精神の薬①（表 3列 × 3行＋見出し行） =====
  // spec の列幅（3.0 / 3.2 / 6.0）では 1 列目に「（パーキンソン病）」が入らず、3 列目は MAO-B 阻害薬の行が 5 行、
  // レボドパの行が 4 行に折れて表の高さが約 4.7in になり本文ゾーン（4.59in）に入らない。
  // 1 列目は「（パーキンソン病）」が 1 行に入る幅、2 列目は「少量の水で／続ける」の幅、残りを 3 列目にとり、
  // 3 列目に「ペチジン・トラマドール・タペンタドール、」「メトクロプラミド・ハロペリドールを避ける」が 1 行に入るようにする。
  // 2 列目を「少量の水で続ける」が 1 行に入る幅（2.91in）にすると 3 列目は 6.0in 以下になり、行 3 の注意が 5 行に折れて
  // 薬剤の並びが 3 か所で切れる（行 2 も「絶食なら／レボドパ静注」で割れる）ので、3 列目を広くとる方をとる。
  // 行 3 は抗うつ薬の並び「三環系・SSRI・SNRI・ミルタザピンは」を 1 行にまとめ、「併用禁忌。」を次の行に置く
  // （並びの途中や「・」の後で割らない。「三環系〜併用禁忌。」の 1 行は 7.86in 要り、この幅では入らない）。
  {
    const id = 'S5';
    const s = slide(id, 2);
    const rows = [
      [q(id, '薬'), q(id, '当日'), q(id, '注意')],
      [drug(q(id, '抗てんかん薬')), keep(q(id, '続ける')),
        cell((l => [l[0], mark(l[1], ['併用禁忌', RED])])(brk(id, '絶食が長引けば静注製剤に。バルプロ酸とカルバペネムは併用禁忌', 'バルプロ酸')))],
      [drug(brk(id, 'レボドパ（パーキンソン病）', '（')), keep(brk(id, '少量の水で続ける', '続ける')),
        cell((l => [mark(l[0], ['悪性症候群', RED]), l[1], l[2]])(brk(id, '急な中止で悪性症候群。絶食ならレボドパ静注。メトクロプラミド・ハロペリドールを避ける', '絶食なら', 'メトクロプラミド')))],
      [drug(brk(id, 'MAO-B 阻害薬（セレギリン）', '（')), keep(q(id, '続ける')),
        cell((l => [l[0], l[1], mark(l[2], ['併用禁忌', RED]), l[3]])(brk(id, 'ペチジン・トラマドール・タペンタドール、三環系・SSRI・SNRI・ミルタザピンは併用禁忌。エフェドリンを避けフェニレフリン少量', '三環系', '併用禁忌', 'エフェドリン')))],
    ];
    const colW = [3.25, 1.88, 7.07];
    placeTable(s, id, rows, { colW, margin: TABLE_MARGIN });
    DEVIATIONS.push(`${id} 表: 列幅を spec の 3.0 / 3.2 / 6.0in から ${colW.join(' / ')}in にした（文言・22pt・行送り 1.2 em はそのまま）。` +
      'spec の 1 列目（3.0in）には「（パーキンソン病）」が入らない（余白込み 3.25in 要る）。2 列目を「少量の水で続ける」が 1 行に入る幅にすると、' +
      '3 列目が 6.0in 以下になり、行 3 の注意が 5 行に折れて薬剤の並びが 3 か所で切れる');
  }

  // ===== S6 神経・精神の薬②（表 3列 × 3行＋見出し行） =====
  // spec の列幅（3.0 / 3.2 / 6.0）では 2 列目の「絶食が続く手術は」、3 列目の「パーキンソン病にハロペリドールは禁忌」が
  // 1 行に入らず（語の途中で折れる）、表が本文ゾーンに入らない。2 列目は「絶食が続く手術は」、3 列目は
  // 「パーキンソン病にハロペリドールは禁忌」が 1 行に入る幅にし、残りを 1 列目にとる（spec の値に近い）。
  {
    const id = 'S6';
    const s = slide(id, 2);
    const rows = [
      [q(id, '薬'), q(id, '当日'), q(id, '注意')],
      [drug(brk(id, 'SSRI・SNRI・三環系', '三環系')), keep(q(id, '続ける')),
        cell(brk(id, '出血が増える。セロトニン症候群（トラマドール、ペチジン、メチレンブルー、リネゾリド）', '（トラマドール', 'メチレンブルー'))],
      [drug(q(id, 'リチウム')), cell((l => [l[0], l[1], l[2], mark(l[3], ['72 時間前に中止', VER])])(brk(id, '小手術は継続。大量出血・絶食が続く手術は 72 時間前に中止', '大量出血', '絶食が', '72 時間'))),
        cell(brk(id, '脱水・NSAIDs・ACE 阻害薬・サイアザイドで蓄積。再開は K・Na・Cr を見て', 'サイアザイド', '再開は'))],
      [drug(q(id, '抗精神病薬')), keep(brk(id, '続ける（精神科と）', '（精神科と）')),
        cell((l => [l[0], mark(l[1], ['禁忌', RED])])(brk(id, 'QT 延長、錐体外路、悪性症候群。パーキンソン病にハロペリドールは禁忌', 'パーキンソン病')))],
    ];
    const colW = [2.96, 2.91, 6.33];
    placeTable(s, id, rows, { colW, margin: TABLE_MARGIN });
    DEVIATIONS.push(`${id} 表: 列幅を spec の 3.0 / 3.2 / 6.0in から ${colW.join(' / ')}in にした（文言・22pt・行送り 1.2 em はそのまま）。` +
      'spec の列幅では 3 列目の「パーキンソン病にハロペリドールは禁忌」が語の途中で折れ、表が本文ゾーンに入らない');
  }

  // ===== S7 甲状腺薬と女性ホルモン（表 3列 × 4行＋見出し行） =====
  // spec の列幅（3.4 / 4.0 / 4.8）で語句の切れ目で改行して入るが、2 列目の「不動になる 3 日前に中止」（太字）が
  // 推定でセル幅の 0.01in 手前まで来るので、3 列目（最長の行「甲状腺クリーゼの予防」「（静脈血栓塞栓症）。」は 3.4in）
  // から 0.2in を 2 列目に回す。行数・行の高さは変わらない。
  {
    const id = 'S7';
    const s = slide(id, 3);
    const rows = [
      [q(id, '薬'), q(id, 'どうする'), q(id, '理由')],
      [drug(q(id, 'レボチロキシン')), cell(brk(id, '当日飲めなくてもよい。5 日超は静注', '5 日超')),
        cell(brk(id, '半減期約 7 日。顕性の機能異常は待機手術を延期', '顕性', '待機手術'))],
      [drug(q(id, '抗甲状腺薬')), keep(q(id, '続ける')), q(id, '甲状腺クリーゼの予防')],
      [drug(q(id, '低用量ピル・LEP')), stop(brk(id, '手術 4 週前に中止、術後 2 週は再開しない', '術後')),
        cell((l => [mark(l[0], ['禁忌', RED]), l[1], l[2]])(brk(id, '添付文書で禁忌（静脈血栓塞栓症）。予約時に確認', '（静脈', '予約時')))],
      [drug(q(id, 'ラロキシフェン')), stop(q(id, '不動になる 3 日前に中止')),
        cell(brk(id, '静脈血栓塞栓症。骨折の緊急手術では VTE 予防を強化', '骨折', 'VTE'))],
    ];
    const colW = [3.4, 4.2, 4.6];
    placeTable(s, id, rows, { colW, margin: TABLE_MARGIN });
    DEVIATIONS.push(`${id} 表: 列幅を spec の 3.4 / 4.0 / 4.8in から ${colW.join(' / ')}in にした（文言・22pt・行数はそのまま）。` +
      '2 列目の「不動になる 3 日前に中止」（太字）が実機のメイリオで折れないよう、余裕のある 3 列目から 0.2in 回した');
  }

  // ===== S8 漢方・サプリ・吸入薬（表 3列 × 3行＋見出し行） =====
  // spec の列幅（3.4 / 4.0 / 4.8）では 3 列目の「セント・ジョーンズ・ワートは」が入らず薬剤名の途中で折れ、
  // 1 列目の「ロイコトリエン拮抗薬」も入らない。1 列目は「ロイコトリエン拮抗薬」、3 列目は「セント・ジョーンズ・ワートは」が
  // 1 行に入る幅にし、残りを 2 列目にとる。
  {
    const id = 'S8';
    const s = slide(id, 3);
    const rows = [
      [q(id, '薬'), q(id, 'どうする'), q(id, '理由')],
      [drug(q(id, '漢方（甘草・麻黄）')), cell(brk(id, '一律に止めない。麻黄は中止を考え、甘草は K を確認', '麻黄は', '甘草は')),
        cell(brk(id, '甘草は偽アルドステロン症（低 K）。麻黄はエフェドリン', '（低 K）'))],
      [drug(q(id, 'サプリ・ハーブ')), stop(brk(id, '1〜2 週前に中止。具体的に尋ねる', '具体的に')),
        cell(brk(id, 'ニンニク・イチョウは出血、セント・ジョーンズ・ワートは薬物代謝', 'セント', '薬物代謝'))],
      [drug(brk(id, '吸入薬・ロイコトリエン拮抗薬', 'ロイコ')), keep(brk(id, '当日も続ける。吸入器を持参', '吸入器')),
        cell(brk(id, '吸入ステロイドはカバーの判断で数える', 'カバー'))],
    ];
    const colW = [3.6, 3.64, 4.96];
    placeTable(s, id, rows, { colW, margin: TABLE_MARGIN });
    DEVIATIONS.push(`${id} 表: 列幅を spec の 3.4 / 4.0 / 4.8in から ${colW.join(' / ')}in にした（文言・22pt・行送り 1.2 em はそのまま）。` +
      'spec の列幅では「セント・ジョーンズ・ワート」「ロイコトリエン拮抗薬」が薬剤名の途中で折れる');
  }

  // ===== S9 当日の内服指示（2列カード + 強調ボックス） =====
  // カード本文 22pt・行送り 1.2 em。右のカード（止める）は「リチウム：大量出血の手術は 72 時間前」が 1 行に入る幅、
  // 左のカード（飲む）は残り（「β遮断薬、Ca 拮抗薬、スタチン」が 1 行に入る）。左右とも本文 7 行。強調ボックスは 24pt 太字 1 行。
  {
    const id = 'S9';
    const s = slide(id, 3);
    const e = emphTight(s, { x: X0, w: BW, yBottom: BOTTOM, lineColor: TEAL, size: 24, minH: 0.62,
      lines: [mark(q(id, '飲む薬・止める薬と再開日を薬ごとに書き、少量の水で飲ませる'), ['止める薬と再開日', VER])] });
    const c1 = {
      title: q(id, '当日も飲む'), titleColor: BLUE,
      items: [
        [q(id, 'β遮断薬、Ca 拮抗薬、スタチン')],
        brk(id, '吸入薬、抗てんかん薬、レボドパ', 'レボドパ'),
        brk(id, 'SSRI、抗甲状腺薬、レボチロキシン（飲めなくても可）', 'レボチロキシン', '（飲めなくても可）'),
        [q(id, '普段のステロイド（＋カバー）')],
      ],
    };
    const c2 = {
      title: q(id, '止める（いつから）'), titleColor: VER,
      items: [
        brk(id, 'ACE 阻害薬・ARB：24 時間前から（HFrEF は相談）', '24 時間'),
        [q(id, '利尿薬：当日朝')],
        [q(id, 'SGLT2 阻害薬：3 日前（→ 4-13）')],
        [q(id, 'リチウム：大量出血の手術は 72 時間前')],
        brk(id, '低用量ピル：4 週前。サプリ：1〜2 週前', 'サプリ'),
      ],
    };
    const w2 = Math.ceil(cardNeedW(c2) * 1000 + 1) / 1000;
    const w1 = r3(BW - GAP - w2);
    if (cardNeedW(c1) > w1 + 1e-9) throw new Error(`${id}: 左のカードの幅不足（必要 ${r3(cardNeedW(c1))} > ${w1}）`);
    const y = TOP, h = r3(e.y - GAP - y);
    const g1 = itemGapFor(c1, h, 0.22), g2 = itemGapFor(c2, h, 0.22);
    drawCard(s, c1, X0, y, w1, h, g1);
    drawCard(s, c2, X0 + w1 + GAP, y, w2, h, g2);
    console.log(`  p${s.page} カード: 幅 ${w1} / ${w2}in、${y}〜${r3(y + h)}in、項目間 ${g1} / ${g2}in、強調ボックス ${e.y}〜${r3(e.y + e.h)}in`);
  }

  // ===== S10 相談する相手と記録（箇条書き5行 + 強調ボックス） =====
  // 箇条書き 24pt。相談先（：の後）を青太字にする。強調ボックスは 24pt 太字 1 行。
  {
    const id = 'S10';
    const s = slide(id, 3);
    const dept = (str, d) => mark(q(id, str), [d, BLUE]);
    const items = [
      dept('HFrEF・不整脈の薬：循環器内科', '循環器内科'),
      dept('副腎不全の疑い、ステロイドカバー：内分泌内科', '内分泌内科'),
      dept('パーキンソン病：脳神経内科', '脳神経内科'),
      dept('リチウム・MAO 阻害薬・抗精神病薬：精神科', '精神科'),
      dept('てんかん・漢方：処方医', '処方医'),
    ];
    const e = emphTight(s, { x: X0, w: BW, yBottom: BOTTOM, lineColor: RED, size: 24, minH: 0.62,
      lines: [mark(q(id, '自己判断で中止しない。相談した相手と事実を記録する'), ['自己判断で中止しない。', RED])] });
    const BUL = { y: TOP, size: 24, gap: 18 };
    const need = heightFor(items.length, BUL.size, BUL.gap);
    const bh = r3(e.y - GAP - BUL.y);
    if (need > bh + 1e-9) throw new Error(`${id}: 箇条書きの高さ ${need} > ${bh}`);
    s.bullets(items, { ...BUL, h: bh, valign: 'middle' });
    console.log(`  p${s.page} 箇条書き枠 ${BUL.y}〜${r3(BUL.y + bh)}in（必要 ${need}in、上下中央）、強調ボックス ${e.y}〜${r3(e.y + e.h)}in`);
  }

  // ===== S11 まとめ（箇条書き5行） =====
  {
    const id = 'S11';
    const s = slide(id, 4);
    const items = [
      mark(q(id, 'ACE 阻害薬・ARB は休薬でも継続でも可。止めたら再開日を書く'), ['止めたら再開日を書く', VER]),
      mark(q(id, 'β遮断薬・スタチン・吸入薬・抗てんかん薬・レボドパは当日も飲む'), ['当日も飲む', BLUE]),
      mark(q(id, 'ステロイドカバー：5 mg/日以上を 1 か月以上。普段の量に上乗せ'), ['普段の量に上乗せ', VER]),
      mark(q(id, 'MAO-B 阻害薬にペチジン・トラマドールは禁忌。SSRI でも慎重に'), ['禁忌', RED]),
      mark(q(id, 'ピル 4 週前、サプリ 1〜2 週前、リチウムは大出血手術で 72 時間前'), ['72 時間前', VER]),
    ];
    const size = 24, gap = 28;
    const hh = heightFor(items.length, size, gap);
    if (hh > BOTTOM - TOP + 1e-6) throw new Error(`${id}: 箇条書きの高さ ${hh} > ${r3(BOTTOM - TOP)}`);
    s.bullets(items, { y: TOP, h: r3(BOTTOM - TOP), size, gap, valign: 'middle' });
  }

  DEVIATIONS.push(`S5〜S8 表: セルの余白を tightTable 既定の [3,6,3,6]pt から [${TABLE_MARGIN.join(',')}]pt にした` +
    '（左を 0.125in にして 1 文字目を罫線から離す。左右の合計は同じで、列幅・改行・行の高さは変わらない）');

  // spec の on-slide（と表の列見出し）がすべてスライドに置かれたか
  for (const [id, sp] of Object.entries(SPEC)) {
    for (const str of sp.quoted) {
      if (!(USED[id] && USED[id].has(str))) throw new Error(`${id}: spec の「${str}」をスライドに置いていません`);
    }
  }

  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  fs.mkdirSync(RENDER_DIR, { recursive: true });
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.render) console.log('[render] ' + JSON.stringify(res.render));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
