'use strict';
/**
 * 4-11 高齢者とフレイル・認知機能
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk()/q() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。spec から文言を変えた箇所は alt() で照合・記録し、ビルド時に一覧を出す。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら scratchpad の render/<NAME>）。
 *
 * 密度への対応:
 *  - 2 列カード（S3〜S6）の本文は 22pt のまま、1 行 1 枠の s.text を一定の行送りで置く（4-13 と同じ方式。
 *    spcPts は slide_lint の高さ推定を壊し、spcPct は代替フォントで行送りが変わるため）。
 *    項目の頭は小さな丸（ネイティブ図形）で示す。カードの内側余白は上下とも文字枠から測り、
 *    項目の間隔はカードごとに余りを配る（左右のカードで下端の余白がそろう）。
 *  - S3・S4 は行送り 1.2 em。
 *  - S5・S6 は spec の文言のままだと 22pt・1.2 em では各カード 10〜12 行になり、全幅の強調ボックスの上に
 *    入らない（1.2 em で入るのは各カード 8 行まで）。文言の短縮を最小にするため行送りを 1.1 em に詰め、
 *    それでも入らない分だけ語句を短くした（alt()。数値・単位・薬剤名・閾値は落としていない）。
 *    左右のカード幅は、語句の切れ目で割ったときの行数が最小になる幅（全角換算の字数で指定）。
 *    2 列の合計幅で入る字数は全角約 33 字なので、spec の文言のままでは S5 は左 11・右 10 行、
 *    S6 は左 10・右 11 行になり、入るのは S5 で 10 行、S6 で 9 行まで（出典 2 行の上、強調ボックスの下に
 *    0.24in を空け、カードの下余白を 0.065in［文字枠から］とったとき）。
 */
const fs = require('fs');
const path = require('path');
const { Deck, C, LAYOUT, fits, heightFor } = require('../_lib/decklib');

const NAME = '4-11_elderly_frailty_cognition';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR ||
  path.join('/tmp/claude-0/-home-user-Anes-lecture/0b714c0d-c11e-535a-a97c-0e6ce9d53e16/scratchpad/render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（見出し・ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
const SPEC_MD = fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8');
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const head = part.split('\n')[0];
    const title = (head.match(/title=「([^」]*)」/) || [])[1];
    const ch = (head.match(/\bch=(\d+)/) || [])[1];
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
      if (inOn) for (const qq of ln.matchAll(/「([^」]*)」/g)) quoted.add(qq[1]);
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
/** spec の文言そのまま（照合のみ） */
function q(id, str) {
  if (!SPEC[id].quoted.has(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
  return str;
}
/** spec の文言 str を、breaks の各語の直前で改行して行配列にする（文言は変えない。改行位置の半角スペースは落とす） */
function brk(id, str, ...breaks) {
  q(id, str);
  return split(str, breaks, id);
}
function split(str, breaks, id) {
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
/** spec の文言を収まりのため短くした箇所（spec 側の元の文言を照合し、ビルド時に一覧を出す） */
const DEVIATIONS = [];
function alt(id, orig, shown, why) {
  q(id, orig);
  DEVIATIONS.push(`${id}: 「${orig}」→「${shown}」（${why}）`);
  return shown;
}
/** alt() の結果を breaks の語の直前で改行する */
function altBrk(id, orig, shown, why, ...breaks) { return split(alt(id, orig, shown, why), breaks, id); }

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 区切りだけでは入らないとき（この講義の出典は 3 件とも 1 件 5.4〜10.3in で、どの 2 件を並べても
//     11.6in を超える）は、文献の途中で改行するが、著者名・誌名の途中では割らない。改行してよい位置は
//     区切り（rank 0）＞ 著者と誌名の間（「et al.」「Theou O.」など著者の末尾の「.」の後。rank 1）＞
//     誌名と年の間（rank 2）の順。
//  2) 1) の位置で入らなければ区切りの前後の空白を詰める（' / ' → '/'）。
//  3) それでも入らなければ（）内の補足を後ろの文献から外す（deviations に記録）。
//  4) どの段階でも 1) の位置で割れないときだけ、文言はそのままで残りの空白で改行する（ログに出す）。
// 文献そのものは落とさない（落とさないと入らないときは Error）。
// ---------------------------------------------------------------------------
const SRC = LAYOUT.source;
function libLines(items) {
  const lines = [];
  for (const it of items) {
    const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
    if (cur && fits(cur, SRC.size, SRC.w)) lines[lines.length - 1] = cur;
    else lines.push(it);
  }
  return lines;
}
const SPLIT_KIND = ['区切り', '著者と誌名の間', '誌名と年の間', '語の切れ目'];
function split2(str, sep, maxRank) {
  const fitsLine = l => fits(l, SRC.size, SRC.w);
  const cands = [];
  for (let i = 1; i < str.length; i++) {
    let a, b, rank;
    if (str.startsWith(sep, i)) { a = str.slice(0, i); b = str.slice(i + sep.length); rank = 0; }
    else if (str[i] === ' ') {
      a = str.slice(0, i); b = str.slice(i + 1);
      if (a.endsWith('/') || b.startsWith('/')) continue;
      if (/ et$/.test(a)) continue;                   // 「et」と「al.」の間では切らない
      if (a.endsWith('.')) rank = 1;                   // 著者の末尾（et al. / Theou O. / Society. など）の後
      else if (/^\d/.test(b)) rank = 2;                // 誌名と年の間
      else rank = 3;
    } else continue;
    a = a.trim(); b = b.trim();
    if (rank <= maxRank && a && b && fitsLine(a) && fitsLine(b)) cands.push({ a, b, rank });
  }
  if (!cands.length) return null;
  cands.sort((x, y) => x.rank - y.rank || y.a.length - x.a.length);
  return cands[0];
}
function sourcesOf(id) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  const items = spec.split(' / ');
  if (libLines(items).length <= 2) return items;
  const stages = [
    { items, sep: ' / ', dev: null },
    { items, sep: '/', dev: "区切りの前後の空白を詰めた（' / ' → '/'）" },
  ];
  const cut = items.slice(), removed = [];
  for (let k = cut.length - 1; k >= 0; k--) {
    if (!/（[^）]*）$/.test(cut[k])) continue;
    const was = cut[k];
    cut[k] = cut[k].replace(/（[^）]*）$/, '');
    removed.push(`「${was}」→「${cut[k]}」`);
    const its = cut.slice(), rm = removed.join('、');
    stages.push({ items: its, sep: ' / ', dev: `${rm}（（）内の補足を外した）` });
    stages.push({ items: its, sep: '/', dev: `${rm}（（）内の補足を外し、区切りの前後の空白を詰めた）` });
  }
  for (const st of stages) {
    const r = split2(st.items.join(st.sep), st.sep, 2);
    if (!r) continue;
    if (st.dev) {
      DEVIATIONS.push(`${id} 出典: ${st.dev}。文言のままでは著者名・誌名を割らずに 16pt 2 行に入らないため。` +
        `改行は${SPLIT_KIND[r.rank]}: ${r.a} ⏎ ${r.b}`);
    } else console.log(`  ${id} 出典: ${SPLIT_KIND[r.rank]}で改行（文言はそのまま）: ${r.a} ⏎ ${r.b}`);
    return [r.a, r.b];
  }
  const r = split2(spec, ' / ', 3);
  if (r) {
    console.log(`  ${id} 出典: 区切りの空白を詰めても（）の補足を外しても、著者名・誌名を割らない位置では 2 行に入らない。` +
      `文言はそのままで語の切れ目で改行: ${r.a} ⏎ ${r.b}`);
    return [r.a, r.b];
  }
  throw new Error(`${id}: 出典が 2 行に収まりません`);
}
/**
 * 1 行の文字列に部分装飾を付けて run 配列にする。marks: [[語, 色, 追加の書式?], ...]（太字）。語は行の中に 1 回だけ現れること。
 * 警告（レッド）は色だけに頼らないよう下線も付ける（RED_U）。
 */
function mark(line, ...marks) {
  let runs = [line];
  for (const [w, color, extra] of marks) {
    const next = [];
    let found = 0;
    for (const r of runs) {
      if (typeof r !== 'string') { next.push(r); continue; }
      const k = r.indexOf(w);
      if (k < 0) { next.push(r); continue; }
      found++;
      if (k > 0) next.push(r.slice(0, k));
      next.push({ text: w, color, bold: true, ...(extra || {}) });
      if (k + w.length < r.length) next.push(r.slice(k + w.length));
    }
    if (found !== 1) throw new Error(`mark: 「${line}」に「${w}」が ${found} 回`);
    runs = next;
  }
  return runs;
}
const runText = l => (Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(l));

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const Y0 = LAYOUT.body.y;        // 1.85
const Y0_T = 1.82;               // 詰めるスライドの上端（金下線の下端 1.576in から 0.244in）
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const r3 = v => Math.round(v * 1000) / 1000;
/**
 * 本文の下端。全スライドで出典が 2 行（枠 y6.5・下揃え。PowerPoint のメイリオ 16pt 2 行で文字の上端 約6.68in）
 * なので、その 0.24in 上。
 */
const BOTTOM = 6.44;

// ---------------------------------------------------------------------------
// 行送りを詰めたテキスト（カードと強調ボックス）。22pt 以上のまま、1 行ずつ decklib の s.text（上揃え）を
// 行送り em おきに置く。枠の高さは decklib / lint の 1 行の下限（22pt 0.41in、24pt 0.46in）で、
// 次の行の枠と少し重なる（文字どうしの枠の重なりは lint・decklib とも対象外）。
// ---------------------------------------------------------------------------
const lineIn = (size, em) => size * em / 72;
const BOX_H = { 22: 0.41, 24: 0.46 };
function tightText(s, lines, o) {
  const size = o.size || 22;
  const em = o.em || 1.2;
  if (!BOX_H[size]) throw new Error(`tightText: size ${size} は 22 / 24 のみ`);
  lines.forEach((line, i) => {
    s.text([line], {
      x: r3(o.x), y: r3(o.y + i * lineIn(size, em)), w: r3(o.w), h: BOX_H[size],
      size, bold: !!o.bold, color: o.color || C.dark, align: o.align || 'left',
    });
  });
  return lines.length * lineIn(size, em);
}

// ---------------------------------------------------------------------------
// カード（S3〜S6）: 角丸の塗り箱＋見出し＋項目。項目の頭は見出しと同じ色の小さな丸（ネイティブ図形）。
// 項目は語句の切れ目で 2〜4 行に割り、2 行目以降も丸の右にそろえる。文字・丸はすべて箱の内側。
// ---------------------------------------------------------------------------
const CARD = { padX: 0.07, ind: 0.2, dot: 0.1, dotX: 0.03, size: 22, titleSize: 24 };
/** 1 文字（全角）あたりの推定幅（22pt）。カード幅の計算用 */
const CH22 = 22 / 72 * 1.12;
/** 本文の全角 n 字が入るカード幅 */
const cardWFor = n => r3(n * CH22 + 2 * CARD.padX + CARD.ind + 0.002);
/** 項目間を除いたカードの高さ。padTop は箱の上端から見出しの文字枠まで、padBottom は最終行の文字枠の下端から箱の下端まで */
function cardFixedH(c, k) {
  const nLines = c.items.reduce((a, it) => a + it.length, 0);
  return k.padTop + lineIn(CARD.titleSize, k.em) + k.titleGap + (nLines - 1) * lineIn(CARD.size, k.em) + BOX_H[CARD.size] + k.padBottom;
}
function drawCard(s, c, x, y, w, h, itemGap, k) {
  s.box({ x: r3(x), y: r3(y), w: r3(w), h: r3(h), fill: C.bg, radius: 0.08 });
  const ix = x + CARD.padX, iw = w - 2 * CARD.padX;
  tightText(s, [c.title], { x: ix, y: y + k.padTop, w: iw, size: CARD.titleSize, bold: true, color: c.titleColor, em: k.em });
  let yy = y + k.padTop + lineIn(CARD.titleSize, k.em) + k.titleGap;
  const L = lineIn(CARD.size, k.em);
  c.items.forEach(lines => {
    // 丸は 1 行目の文字の高さの中ほど（メイリオ 22pt で枠上端から約 0.2in）に置く
    s.circle({ cx: r3(ix + CARD.dotX + CARD.dot / 2), cy: r3(yy + k.dotY), d: CARD.dot, fill: c.titleColor });
    tightText(s, lines, { x: ix + CARD.ind, y: yy, w: iw - CARD.ind, size: CARD.size, em: k.em });
    yy += lines.length * L + itemGap;
  });
  const end = yy - itemGap - L + BOX_H[CARD.size];   // 最終行の文字枠の下端
  if (end > y + h - k.padBottom + 1e-6) {
    throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さ不足（本文の下端 ${r3(end)} > ${r3(y + h - k.padBottom)}）`);
  }
  return end;
}
/**
 * 2 列のカード。widths: 各カードの幅（in）。高さは y〜bottom。
 * 項目の間隔はカードごとに余りを配る（上限 maxGap）。左右のカードとも最終行の下の余白が padBottom にそろう。
 */
function cardRow(s, cards, o) {
  const k = { em: o.em, padTop: o.padTop, padBottom: o.padBottom, titleGap: o.titleGap, dotY: o.dotY };
  const sumW = o.widths.reduce((a, b) => a + b, 0) + GAP * (cards.length - 1);
  if (sumW > BW + 1e-6) throw new Error(`p${s.page}: カード幅の合計 ${r3(sumW)} > ${r3(BW)}`);
  const h = r3(o.bottom - o.y);
  const gaps = cards.map(c => {
    const free = h - cardFixedH(c, k);
    if (free < -1e-9) throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さが足りません（h=${h}、必要 ${r3(cardFixedH(c, k))}）`);
    const g = c.items.length > 1 ? Math.min(o.maxGap, free / (c.items.length - 1)) : 0;
    return Math.max(0, Math.floor(g * 1000) / 1000);
  });
  let x = X0;
  const ends = cards.map((c, i) => { const e = drawCard(s, c, x, o.y, o.widths[i], h, gaps[i], k); x += o.widths[i] + GAP; return e; });
  console.log(`  p${s.page} カード: 幅 ${o.widths.join(' / ')}in、${o.y}〜${o.bottom}in（高さ ${h}）、行送り ${o.em} em、` +
    `行数 ${cards.map(c => c.items.reduce((a, it) => a + it.length, 0)).join(' / ')}、項目間 ${gaps.join(' / ')}in、` +
    `下余白（文字枠から） ${ends.map(e => r3(o.bottom - e)).join(' / ')}in`);
  return { h, gaps };
}

/** 強調ボックス（白地＋ティールの枠。文字は太字・中央）。decklib の box（通常の行送り）。 */
const EMPH = { radius: 0.1, lineW: 3 };
function emphBox(s, o) {
  s.box({
    x: r3(o.x != null ? o.x : X0), y: r3(o.y), w: r3(o.w != null ? o.w : BW), h: o.h, fill: C.white, radius: EMPH.radius,
    line: { color: o.lineColor || C.teal, width: EMPH.lineW },
    text: o.lines, size: o.size, bold: true, color: C.dark, margin: o.margin,
  });
}
/**
 * 行送りを詰めた強調ボックス（22pt・2 行）。下端 yBottom に置き、{y, h} を返す（h を渡せばその高さ）。
 * 字面は行の枠の中で下寄り（メイリオで枠上端から約 0.18〜1.18 em、LibreOffice の代替フォントの描画で
 * 約 0.2〜1.11 em）。両者の平均（GLYPH）で字面の上下が箱の中央に来るように文字の枠を置く。
 */
const GLYPH = { top: 0.19, bottom: 1.145 };
const EMPH_T = { padY: 0.05, padX: 0.15 };
function emphTight(s, o) {
  const size = o.size || 22;
  const n = o.lines.length, L = lineIn(size, o.em), emIn = size / 72;
  const h = r3(o.h != null ? o.h : 2 * EMPH_T.padY + n * L);
  const y = r3(o.yBottom - h);
  s.box({ x: X0, y, w: BW, h, fill: C.white, radius: EMPH.radius, line: { color: o.lineColor || C.teal, width: EMPH.lineW } });
  const glyphH = (n - 1) * L + (GLYPH.bottom - GLYPH.top) * emIn;   // 1 行目の字面の上端から最終行の字面の下端まで
  const ty = y + (h - glyphH) / 2 - GLYPH.top * emIn;
  if (ty < y - 1e-6 || ty + (n - 1) * L + BOX_H[size] > y + h + 1e-6) throw new Error(`p${s.page} 強調ボックス: 文字枠が箱からはみ出します（h=${h}）`);
  tightText(s, o.lines, { x: X0 + EMPH_T.padX, y: ty, w: BW - 2 * EMPH_T.padX, size, bold: true, align: 'center', em: o.em });
  return { y, h };
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['フレイル', '認知\n機能', 'せん妄', '薬と\n麻酔', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-11 高齢者とフレイル・認知機能',
  });
  /** 本文スライド。タイトルと章は spec の見出し行（## Sn  ch=…  title=「…」）から取る */
  const slide = id => {
    const sp = SPEC[id];
    if (!sp.title || sp.ch == null) throw new Error(`${id}: spec の見出し行に title / ch がありません`);
    return deck.slide({ title: sp.title, chapter: sp.ch, notes: notesOf(id), sources: sourcesOf(id) });
  };
  const VER = C.vermilion, BLUE = C.blue, RED = C.red;
  const RED_U = { underline: true };   // 警告のレッドは下線も付け、近くのバーミリオンと色だけで区別させない

  // ===== S1 表紙 =====
  {
    const subtitle = '4-11 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: SPEC.S1.title, subtitle, meta, notes: notesOf('S1') });
  }

  // 箇条書き＋強調ボックスのスライド（S2・S7）: 24pt・段落後 10pt。強調ボックスは 28pt 太字 2 行
  const BUL = { y: 1.9, size: 24, gap: 10 };
  const bulH = n => heightFor(n, BUL.size, BUL.gap);
  const EMPH_H = 1.2;
  const EMPH_W = { x: X0, w: BW };   // 強調ボックスは全スライド全幅（S3〜S6 と同じ）

  // ===== S2 フレイルが予後を決める（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S2';
    const s = slide(id);
    s.bullets([
      mark(q(id, 'CFS でフレイル → 全死亡のオッズ比 4.0（メタ解析）'), ['オッズ比 4.0', VER]),
      mark(q(id, '1 年死亡 13.6% 対 4.8%（診断コード定義）、術直後ほど差が大きい'), ['13.6% 対 4.8%', VER]),
      mark(q(id, '合併症 OR 2.5、在院 1.7 倍、施設退院 OR 20（Makary 2010）'), ['施設退院 OR 20', VER]),
      mark(q(id, '術後せん妄 OR 2.1（メタ解析）'), ['OR 2.1', VER]),
      mark(q(id, 'CFS 4 以上で注意、5 以上をフレイル（閾値は研究で違う）'), ['4 以上で注意', BLUE], ['5 以上をフレイル', BLUE]),
    ], BUL);
    const lines = brk(id, 'ASA-PS や心リスク指数にフレイルを足すと、予測が良くなる', '予測');
    const by = r3(BOTTOM - 0.04 - EMPH_H);
    if (by < BUL.y + bulH(5) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(5))}in`);
    emphBox(s, { ...EMPH_W, y: by, h: EMPH_H, size: 28, lines: [mark(lines[0], ['フレイルを足す', VER]), lines[1]] });
  }

  // 2 列カード＋強調ボックス（1 行・24pt）の共通寸法（S3・S4。行送り 1.2 em）
  // padBottom は最終行の文字枠の下端から（メイリオの字面の下端は文字枠の下端の約 0.05in 上）
  const K12 = { em: 1.2, padTop: 0.06, padBottom: 0.08, titleGap: 0.04, dotY: 0.2, maxGap: 0.36 };
  const EMPH1_24 = 0.58;
  const halfW = r3((BW - GAP) / 2);   // 5.995

  // ===== S3 フレイルの測り方（2列カード + 強調ボックス） =====
  {
    const id = 'S3';
    const s = slide(id);
    const ey = r3(BOTTOM - EMPH1_24);
    emphBox(s, { y: ey, h: EMPH1_24, size: 24,
      lines: [mark(q(id, '転倒歴、眼鏡・補聴器、独居・介護者、本人の治療の目標も聞く'), ['本人の治療の目標', VER])] });
    cardRow(s, [
      {
        title: q(id, 'CFS（9 段階）'), titleColor: BLUE,
        items: [
          [mark(q(id, '約 2 週間前の病前の状態で付ける'), ['病前', VER])],
          brk(id, '移動、生活機能（IADL・ADL）、思考、健康の自己評価を聞く', '思考'),
          [mark(q(id, '迷ったら重い方'), ['重い方', VER])],
          brk(id, '若年者・固定した単一障害には使わない', '使わない'),
        ],
      },
      {
        title: q(id, 'J-CHS 基準（2020）と周辺'), titleColor: VER,
        items: [
          [q(id, '握力 男 28 kg 未満・女 18 kg 未満')],
          [q(id, '歩行速度 1.0 m/秒未満')],
          (() => { const l = brk(id, '体重減少、疲労感、身体活動の低下（3 項目以上でフレイル）', '（3'); return [l[0], mark(l[1], ['3 項目以上でフレイル', VER])]; })(),
          [q(id, '基本チェックリスト 8 点以上')],
          brk(id, '栄養：BMI 18.5 未満、Alb 3.0 g/dL 未満、6 か月で 10〜15% 超の体重減少', 'Alb', '6 か月'),
        ],
      },
    ], { ...K12, y: Y0_T, bottom: r3(ey - GAP), widths: [halfW, halfW] });
  }

  // ===== S4 Mini-Cog を全員に（2列カード + 強調ボックス） =====
  // 右のカードは「せん妄 21% 対 7%、自宅以外への退院 67% 対 34%（Culley 2017）」を「自宅以外への」の後で割り、
  // 2 行に収める（3 行に割ると右のカードが 9 行になり、1.2 em では強調ボックスの上に入らない）。
  // 左のカードは「（ASA Brain Health の専門家合意）」（全角換算 16.74 字）が 1 行に入る幅、右は残り（16.3 字）。
  {
    const id = 'S4';
    const s = slide(id);
    const w1 = cardWFor(16.74);
    const ey = r3(BOTTOM - EMPH1_24);
    // 24pt で全幅 1 行（推定 12.10in）に入れるため、箱の内側余白を 0.05in にする
    emphBox(s, { y: ey, h: EMPH1_24, size: 24, margin: 0.05,
      lines: [mark(q(id, '65 歳を超える全員に『術後に混乱・記憶障害が起こりうる』と説明する'), ['『術後に混乱・記憶障害が起こりうる』', VER])] });
    cardRow(s, [
      {
        title: q(id, 'どう測るか'), titleColor: BLUE,
        items: [
          brk(id, '65 歳超の全員にベースラインの簡易スクリーニング（ASA Brain Health の専門家合意）', '簡易', '（ASA'),
          (() => { const l = brk(id, 'Mini-Cog：3 語の遅延再生＋時計描画、2〜4 分、2 点以下で陽性', '時計描画'); return [mark(l[0], ['Mini-Cog：', BLUE]), mark(l[1], ['2 点以下で陽性', VER])]; })(),
          brk(id, 'MMSE 7〜10 分（23 点以下）、HDS-R（20 点以下）', 'HDS-R'),
          [q(id, '眼鏡・補聴器をつけ、前投薬の前に')],
        ],
      },
      {
        title: q(id, '陽性なら'), titleColor: VER,
        items: [
          [mark(q(id, '診断ではない。詳しい評価へつなぐ'), ['診断ではない', VER])],
          brk(id, '認知症の診断がない人工関節患者の 24% が陽性', '人工関節'),
          brk(id, 'せん妄 21% 対 7%、自宅以外への退院 67% 対 34%（Culley 2017）', '退院'),
          brk(id, '家族に周術期神経認知障害のリスクを説明', 'リスク'),
          [q(id, '意思決定能力を確かめる（→ 4-5）')],
        ],
      },
    ], { ...K12, y: Y0_T, bottom: r3(ey - GAP), widths: [w1, r3(BW - GAP - w1)] });
  }

  // 詰める 2 列カード（S5・S6。行送り 1.1 em）。下余白は文字枠から 0.065in（メイリオの字面から約 0.11in）
  const K11 = { em: 1.1, padTop: 0.04, padBottom: 0.065, titleGap: 0.01, dotY: 0.19, maxGap: 0.12 };

  // ===== S5 術後せん妄：リスクと予防（2列カード + 強調ボックス） =====
  // spec の文言のままでは 22pt で左 11・右 10 行（1.2 em で入るのは 8 行）。行送りを 1.1 em にし、
  // 強調ボックスを 1 行（22pt）にして各カード 10 行まで入れる。そのために 2 か所を短くした（alt()）。
  // 左のカード 16 字・右のカード 17 字（全角換算）の幅（右は「多因子の非薬物的介入（強い推奨）：」が 17 字、
  // 左を 17 字以上にすると右の項目が 2 行ずつ増える）。
  // 左のカードで外すのは「（ESAIC）」ではなく「（高齢者評価）」にした（MMSE 25 未満は ESAIC の閾値で、
  // S4 の MMSE 23 点以下と目的が違うことをスライド面で示すため。高齢者評価の項目も ESAIC の推奨にある）。
  {
    const id = 'S5';
    const s = slide(id);
    const EH = 0.46;
    const ey = r3(BOTTOM - EH);
    emphBox(s, { y: ey, h: EH, size: 22,
      lines: [mark(alt(id, '半数以上が見逃される（低活動型）。回復室を出る前と病棟でスクリーニングする',
        '半数以上が見逃される（低活動型）。回復室を出る前と病棟でスクリーニング',
        '22pt・全幅で 1 行に収め、カードに 1 行ぶんの高さを空けるため、末尾の「する」を外して体言止めにした。' +
        '元の文言は全角換算 37 字で、22pt・全幅（12.23in）の 1 行に入るのは 35 字まで。2 行にするとカードが各 9 行になり項目が入らない'),
      ['半数以上が見逃される', RED, RED_U])] });
    const w1 = cardWFor(16.0);
    cardRow(s, [
      {
        title: q(id, 'リスク因子'), titleColor: BLUE,
        items: [
          brk(id, '高齢、ASA-PS 3 以上、Charlson 2 超、MMSE 25 未満（ESAIC）', 'Charlson', 'MMSE'),
          altBrk(id, 'フレイル、感覚障害、低栄養、多剤、貧血、社会的因子（高齢者評価）', 'フレイル、感覚障害、低栄養、多剤、貧血、社会的因子',
            '左のカードが 1 行あふれる（11 行）ため、付記「（高齢者評価）」を外した。「多剤、貧血、社会的因子（高齢者評価）」は 18 字で' +
            '左のカード（16 字）の 1 行に入らない。ノートに「老年医学的評価（スライドの「高齢者評価」）」とある', '多剤'),
          brk(id, 'せん妄の既往、アルコール多飲、ベンゾジアゼピン、全身麻酔の手術（国内の加算）', 'ベンゾ', '（国内'),
          (() => { const l = brk(id, '大腿骨近位部骨折 35〜65%、大手術 20〜40%', '大手術'); return [mark(l[0], ['35〜65%', VER]), mark(l[1], ['20〜40%', VER])]; })(),
        ],
      },
      {
        title: q(id, '予防（AGS 2015・ESAIC 2024）'), titleColor: VER,
        items: [
          (() => { const l = brk(id, '多因子の非薬物的介入（強い推奨）：見当識、睡眠、早期離床、眼鏡・補聴器、水分・栄養、鎮痛、薬の見直し、酸素化、便秘', '見当識', '眼鏡', '薬の見直し'); return [mark(l[0], ['多因子の非薬物的介入', VER]), l[1], l[2], l[3]]; })(),
          [mark(q(id, 'せん妄のオッズ 44% 減（メタ解析）'), ['44% 減', VER])],
          [mark(q(id, '老年科医の関与 32% 対 50%（骨折）'), ['32% 対 50%', VER])],
          (() => { const l = brk(id, '予防薬は勧めない（ESAIC 弱）。抗精神病薬は根拠不十分、コリンエステラーゼ阻害薬は始めない（AGS 強）', '抗精神病薬', 'コリン', '始めない'); return [mark(l[0], ['予防薬は勧めない', RED, RED_U]), l[1], l[2], l[3]]; })(),
        ],
      },
    ], { ...K11, y: Y0_T, bottom: r3(ey - GAP), widths: [w1, r3(BW - GAP - w1)] });
  }

  // ===== S6 避ける薬と麻酔の調節（2列カード + 強調ボックス） =====
  // spec の文言のままでは 22pt・下の幅で左 11 行・右 11 行（1.2 em で入るのは 8 行）。
  // 行送りを 1.1 em にして各カード 9 行まで入れ、薬剤名・数値を落とさずに 4 か所を短くした（alt()）。
  // 左のカード 18 字・右のカード 15 字（全角換算）の幅（左は「抗コリン：スコポラミン、アトロピン、」と
  // 「（ファモチジンも）、メトクロプラミド」が 18 字。右を 16 字にすると左の項目が 3 行に割れる）。
  // 強調ボックス（2 行）の高さは、カード 9 行と下余白をとった残り（約 0.8in）。
  {
    const id = 'S6';
    const s = slide(id);
    const emLines = brk(id, '常用薬 5〜6 剤以上は転倒・有害事象が増える。お薬手帳で睡眠薬・抗コリン薬・市販薬まで確認', 'お薬手帳');
    const EH6 = 0.799;
    const em = emphTight(s, { yBottom: BOTTOM, em: K11.em, lineColor: C.teal, h: EH6,
      lines: [mark(emLines[0], ['5〜6 剤以上', VER]), mark(emLines[1], ['お薬手帳', BLUE])] });
    const w1 = cardWFor(18.0);
    cardRow(s, [
      {
        title: q(id, '避ける・慎重に（AGS・Beers）'), titleColor: VER,
        items: [
          brk(id, 'ベンゾジアゼピン（ミダゾラム、ジアゼパム）', '（ミダゾラム'),
          altBrk(id, '抗コリン薬：スコポラミン、アトロピン、プロメタジン、ヒドロキシジン', '抗コリン：スコポラミン、アトロピン、プロメタジン、ヒドロキシジン',
            '左のカードが 1 行あふれるため、見出し語「抗コリン薬：」を「抗コリン：」にした（薬剤名はすべて残す）', 'プロメタジン'),
          brk(id, 'ジフェンヒドラミン、H2 遮断薬（ファモチジンも）、メトクロプラミド', '（ファモチジン'),
          [alt(id, 'ペチジン（シバリングも加温が先）、ハロペリドールの予防投与', 'ペチジン、ハロペリドールの予防投与',
            '左のカードが 1 行あふれるため、補足「（シバリングも加温が先）」を外した。ノートに「シバリングにはまず加温し」とある')],
          (() => { const l = brk(id, '例外：離脱予防のベンゾジアゼピン、徐脈へのアトロピン', '徐脈'); return [mark(l[0], ['例外：', BLUE]), l[1]]; })(),
        ],
      },
      {
        title: q(id, '麻酔の調節'), titleColor: BLUE,
        items: [
          (() => { const l = brk(id, 'MAC は 10 歳ごとに約 6% 低下（80 歳のセボフルラン約 1.4%）', '（80'); return [mark(l[0], ['約 6% 低下', VER]), l[1]]; })(),
          brk(id, '年齢補正 MAC 比を見る。BIS は高く出ることがある', 'BIS'),
          brk(id, 'プロポフォール・ミダゾラムは減量し、ゆっくり', '減量'),
          [mark(alt(id, 'ベースラインに対する相対的低血圧を避ける。正常体温', '相対的低血圧を避ける。正常体温',
            '右のカードが 1 行あふれるため「ベースラインに対する」を外した（まとめ S8 と同じ言い方。ノートに「ベースラインからみた相対的な低血圧」とある）'),
          ['相対的低血圧', VER])],
          altBrk(id, '鎮静の少ない多角的鎮痛。区域麻酔そのものではせん妄は減らない', '鎮静の少ない多角的鎮痛。区域麻酔自体でせん妄は減らない',
            '右のカードが 1 行あふれるため「そのものでは」を同じ意味の「自体で」にした。「区域麻酔そのものではせん妄は減らない」は' +
            ' 18 字で右のカード（15 字）の 1 行に入らない。「そのもの」を外すと主張が強くなる（区域麻酔ではせん妄は減らない）ので、' +
            '区域麻酔という選択だけでは減らないという意味は残した', '区域麻酔'),
        ],
      },
    ], { ...K11, y: Y0_T, bottom: r3(em.y - GAP), widths: [w1, r3(BW - GAP - w1)] });
  }

  // ===== S7 相談先と離島での工夫（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S7';
    const s = slide(id);
    s.bullets([
      mark(q(id, '老年科・認知症ケアチーム・精神科リエゾンに術前から（遠隔でも）'), ['術前から', VER]),
      mark(q(id, '家族に：眼鏡・補聴器・時計・慣れた物を持参し、付き添いを'), ['付き添い', BLUE]),
      mark(q(id, '看護師・薬剤師・リハビリと多因子介入を入院初日から'), ['入院初日から', VER]),
      mark(q(id, '術前から栄養と運動（プレハビリテーション）、飲酒・睡眠薬の整理'), ['栄養と運動', BLUE]),
      mark(q(id, '退院先と治療の目標を本人・家族と前もって話す（→ 4-6）'), ['前もって話す', VER]),
    ], BUL);
    const lines = brk(id, '老年科医がいない島では、チェックリストと多職種で『老年科の目』を代わりに入れる', '『老年科の目』');
    const by = r3(BOTTOM - 0.04 - EMPH_H);
    if (by < BUL.y + bulH(5) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(5))}in`);
    emphBox(s, { ...EMPH_W, y: by, h: EMPH_H, size: 28, lines: [lines[0], mark(lines[1], ['『老年科の目』', VER])] });
  }

  // ===== S8 まとめ（箇条書き5行） =====
  {
    const id = 'S8';
    const s = slide(id);
    s.bullets([
      mark(q(id, '年齢だけでなく CFS を病前で。4 以上は注意、5 以上はフレイル'), ['CFS を病前で', BLUE]),
      mark(q(id, '65 歳超は全員 Mini-Cog。陽性なら家族に説明し能力を確かめる'), ['全員 Mini-Cog', BLUE]),
      mark(q(id, 'せん妄は大手術で 20〜40%、半数以上を見逃す。多因子介入で予防'), ['多因子介入で予防', VER]),
      mark(q(id, '薬で予防しない。ベンゾジアゼピン・抗コリン薬・ペチジンを避ける'), ['予防しない', RED, RED_U]),
      mark(q(id, 'MAC は 10 歳で 6% 低下。相対的低血圧を避け、鎮静の少ない鎮痛'), ['10 歳で 6% 低下', VER]),
    ], { y: 2.0, size: 24, gap: 32 });
  }

  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
