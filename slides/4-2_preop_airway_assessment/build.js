'use strict';
/**
 * 4-2 術前診察での気道評価
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。q()/brk() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 図はすべてネイティブ図形（箱・freeform）。図のラベルは図の外に置く。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fits, fitInfo, heightFor, measure } = require('../_lib/decklib');

const NAME = '4-2_preop_airway_assessment';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

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
      // on-slide の「…」（入れ子の『』はそのまま）と、layout 行の「…」（表の列名など）
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

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 文言はそのままで、文献の途中の語の切れ目で改行する（著者「et al.」の後 > 誌名と年の間 > その他の空白）。
//  2) 入らなければ区切りの前後の空白を詰める（' / ' → '/'）。
//  3) それでも入らなければ（）内の補足を後ろの文献から外す（外した数が最小の形。空白は詰めない形を優先）。
//  文献そのものは落とさない。元と違う形にしたものは DEVIATIONS に記録してビルド時に表示する。
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
      if (a.endsWith('.')) rank = 1;                   // 著者名（et al. / 頭文字.）の後
      else if (/^\d/.test(b)) rank = 2;                // 誌名と年の間
      else rank = 3;
      if (a.endsWith('/') || b.startsWith('/')) continue;
    } else continue;
    a = a.trim(); b = b.trim();
    if (a && b && fitsLine(a) && fitsLine(b)) cands.push({ a, b, rank });
  }
  if (!cands.length) return null;
  cands.sort((x, y) => x.rank - y.rank || y.a.length - x.a.length);
  return [cands[0].a, cands[0].b];
}
/** override: spec の sources 行の代わりに使う形（指定時は DEVIATIONS に記録） */
function sourcesFor(id, override) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  const base = override || spec;
  const note = (form, lines) => {
    if (form !== spec) DEVIATIONS.push(`${id} 出典: ${spec} → ${lines.join(' ⏎ ')}`);
    return lines;
  };
  const items = base.split(' / ');
  if (items.every(it => fits(it, SRC.size, SRC.w)) && libLines(items).length <= 2) {
    if (override) note(base, libLines(items));
    return items;
  }
  let r = split2(base, ' / ');
  if (r) return note(base, r);
  r = split2(items.join('/'), '/');
  if (r) return note(items.join('/'), r);
  const cut = items.slice();
  for (let k = cut.length - 1; k >= 0; k--) {
    if (!/（[^）]*）$/.test(cut[k])) continue;
    cut[k] = cut[k].replace(/（[^）]*）$/, '');
    for (const sep of [' / ', '/']) {
      const s = cut.join(sep);
      r = split2(s, sep);
      if (r) return note(s, r);
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
const BOT = 6.45;                // 出典 2 行（枠 y6.5）のときの本文下端（全スライド出典 2 行）
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;

/** 強調ボックス（白地＋色枠・26pt 太字）。S2・S6・S8 で位置・幅をそろえる */
const EMPH = { x: X0 + 0.6, w: r3(BW - 1.2), radius: 0.1, lineW: 3.5, size: 26, margin: 0.08 };
const emphH = n => r2(heightFor(n, EMPH.size, 0) + 2 * EMPH.margin + 0.02);
function emphBox(s, y, lines, lineColor) {
  const h = emphH(lines.length);
  s.box({
    x: EMPH.x, y: r3(y), w: EMPH.w, h, fill: C.white, radius: EMPH.radius,
    line: { color: lineColor, width: EMPH.lineW }, margin: EMPH.margin,
    text: lines, size: EMPH.size, bold: true, color: C.dark,
  });
  return { y: r3(y), h, bottom: r3(y + h) };
}

/**
 * 箇条書きの列（手組み）。行頭記号「•」を別の枠に置き、記号は常に本文の黒にする
 * （decklib の bullets は行頭の run の色が記号に移るため、行頭から色を付ける行で記号まで色が付く）。
 * items: 各項目 = 1 行（文字列 or run 配列）。y を渡すと上端、yc を渡すと縦中央。戻り値 {top, bottom}
 */
function bulletCol(s, items, o) {
  const size = o.size || 24, ind = o.indent || 0.45, gap = o.gap != null ? o.gap : 0.3;
  const lh = heightFor(1, size, 0);
  const total = lh * items.length + gap * (items.length - 1);
  const top = o.yc != null ? o.yc - total / 2 : o.y;
  items.forEach((line, i) => {
    const y = r3(top + i * (lh + gap));
    s.text('•', { x: o.x, y, w: ind, h: lh, size, color: C.dark });
    s.text([line], { x: r3(o.x + ind), y, w: r3(o.w - ind), h: lh, size, color: C.dark });
  });
  return { top: r3(top), bottom: r3(top + total) };
}

// ---------------------------------------------------------------------------
// 見出し＋項目のカード（行頭記号つき）を横に並べる（S3・S6・S8）
//  - 幅: 各カードの最長行（ライブラリ推定）＋行頭記号ぶんを確保し、余りを等分する。
//  - 1 項目を 2 行に割ったときは 2 行目を記号の右にぶら下げる（項目の区切りが見える）。
//  - alignRows: 項目 i の上端を全カードでそろえる（項目数が同じとき）。
// ---------------------------------------------------------------------------
const CARD = { padX: 0.2, padY: 0.12, titleSize: 24, size: 22, indent: 0.36, titleGap: 0.08, radius: 0.08 };
function cardRow(s, cards, o) {
  const n = cards.length;
  const need = cards.map(c => Math.max(
    measure(runText(c.title), CARD.titleSize),
    ...c.items.flat().map(l => measure(runText(l), CARD.size) + CARD.indent)));
  const avail = BW - GAP * (n - 1) - 2 * CARD.padX * n;
  const extra = avail - need.reduce((a, b) => a + b, 0);
  if (extra < -1e-9) throw new Error(`p${s.page}: カード幅が ${r3(-extra)}in 足りません`);
  const widths = need.map(v => Math.floor((v + extra / n + 2 * CARD.padX) * 1000) / 1000);
  const h = r3(o.bottom - o.y);
  const titleH = heightFor(1, CARD.titleSize, 0);
  const bodyAvail = h - 2 * CARD.padY - titleH - CARD.titleGap;
  const itemH = lines => heightFor(lines.length, CARD.size, 0);
  const nItems = cards[0].items.length;
  if (o.alignRows && cards.some(c => c.items.length !== nItems)) throw new Error(`p${s.page}: alignRows は項目数が同じカードのみ`);
  const slotsOf = c => (o.alignRows
    ? c.items.map((_, r) => Math.max(...cards.map(d => itemH(d.items[r]))))
    : c.items.map(itemH));
  let x = X0;
  const info = [];
  cards.forEach((c, i) => {
    const w = widths[i];
    const slots = slotsOf(c);
    const sum = slots.reduce((a, b) => a + b, 0);
    let g = c.items.length > 1 ? (bodyAvail - sum) / (c.items.length - 1) : 0;
    if (o.alignRows) g = Math.min(...cards.map(d => (bodyAvail - slotsOf(d).reduce((a, b) => a + b, 0)) / (d.items.length - 1)));
    g = Math.min(g, o.maxItemGap != null ? o.maxItemGap : 0.36);
    g = Math.floor(g * 1000) / 1000;
    if (g < -1e-9) throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さが ${r3(-g * (c.items.length - 1))}in 足りません`);
    s.box({ x: r3(x), y: o.y, w, h, fill: C.bg, radius: CARD.radius });
    const tx = r3(x + CARD.padX), iw = r3(w - 2 * CARD.padX);
    s.text([c.title], { x: tx, y: r3(o.y + CARD.padY), w: iw, h: titleH, size: CARD.titleSize, bold: true, color: c.titleColor });
    let y = o.y + CARD.padY + titleH + CARD.titleGap;
    c.items.forEach((lines, k) => {
      const ih = itemH(lines);
      const iy = r3(y + (slots[k] - ih) / 2);
      s.text('•', { x: tx, y: iy, w: CARD.indent, h: heightFor(1, CARD.size, 0), size: CARD.size, color: C.dark });
      s.text(lines, { x: r3(tx + CARD.indent), y: iy, w: r3(iw - CARD.indent), h: ih, size: CARD.size, color: C.dark });
      y += slots[k] + g;
    });
    const end = y - g;
    if (end > o.y + h - CARD.padY + 1e-6) throw new Error(`p${s.page} カード「${runText(c.title)}」: 本文の下端 ${r3(end)}in > ${r3(o.y + h - CARD.padY)}in`);
    info.push(`w ${w}in・項目間 ${r2(g)}in`);
    x += w + GAP;
  });
  console.log(`  p${s.page} カード: ${info.join(' / ')}、高さ ${h}in（${o.y}〜${r3(o.y + h)}in）`);
}

/**
 * 行の多い表（S4 用。3-11 の tightTable を写したもの）。decklib の s.table と同じ見た目
 * （見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、上下中央）で、行間だけを固定値 LS pt（= 1.2 em）に詰める。
 * 見出し行も本文と同じ左揃え（列の頭をそろえる）。
 * s.table は行高を実機メイリオの 1.5 em で取るため、7 行の表が本文ゾーン（出典 2 行時 1.85〜6.45in）に入らない。
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

// ---------------------------------------------------------------------------
// 口の模式図（S5）。正面から見た開口: 口腔（淡色）・咽頭の入口（暗色のアーチ）・口蓋垂・舌。
//  アーチの縁（軟口蓋の縁と左右の口蓋弓）と口蓋垂は青の線、舌はバーミリオンの淡色。
//  class が上がるほど舌の上縁が上がり、口蓋弓 → 口蓋垂 → 軟口蓋の順に隠れる。
//  上下の位置は口の高さに対する割合 f（上端 0・下端 1）。
// ---------------------------------------------------------------------------
const MOUTH = {
  hard: 'FBE9DF', soft: 'F2C4B0', dark: '404040', tongue: 'E8A673', lip: '595959',
  fSoft: 0.2,      // 軟口蓋の帯の上端（ここより上が硬口蓋）
  fTop: 0.33,      // アーチの頂点（軟口蓋の縁の中央。口蓋垂の付け根）
  fUvula: 0.5,     // 口蓋垂の先端
  fShoulder: 0.62, // アーチの肩（ここから下が口蓋弓の柱）
  fBottom: 0.86,   // 柱の下端（舌に隠れる）
  halfW: 0.36,     // アーチの半幅（口の半幅に対する割合）
  softW: 0.13,     // 軟口蓋・口蓋弓の帯の幅（アーチの外側へ。in）
  levels: { 1: 0.8, 2: 0.6, 3: 0.42, 4: 0.17 },   // 舌の上縁（両端の高さ）
  dome: 0.04,      // 舌の上縁の中央の盛り上がり（in）
  uvTopHalfW: 0.1, uvTipR: 0.075,                  // 口蓋垂の付け根の半幅・先端の丸みの半径（in）
};
/** アーチ（下端 fBottom の左 → 肩 → 頂点 fTop の弧 → 肩 → 右下）。half: 半幅（in） */
function archPts(cx, half, yTop, yShoulder, yBottom) {
  const pts = [[r3(cx - half), r3(yBottom)], [r3(cx - half), r3(yShoulder)]];
  for (let k = 1; k < 24; k++) {
    const t = Math.PI - (Math.PI * k) / 24;
    pts.push([r3(cx + half * Math.cos(t)), r3(yShoulder - (yShoulder - yTop) * Math.sin(t))]);
  }
  pts.push([r3(cx + half), r3(yShoulder)], [r3(cx + half), r3(yBottom)]);
  return pts;
}
function mouthFig(s, o) {
  const { cx, top, w, h, cls } = o;
  const rx = w / 2, ry = h / 2, cy = top + ry;
  const Y = f => top + f * h;
  const N = 72;
  const ell = [];
  for (let k = 0; k < N; k++) {
    const t = (2 * Math.PI * k) / N;
    ell.push([r3(cx + rx * Math.cos(t)), r3(cy + ry * Math.sin(t))]);
  }
  // 口腔の天井（硬口蓋。淡色）
  s.polyline(ell, { color: MOUTH.hard, width: 1, closed: true, fill: MOUTH.hard });
  // 軟口蓋と左右の口蓋弓（アーチの外側の帯。やや濃い色）
  const a = MOUTH.halfW * rx, ys = Y(MOUTH.fShoulder), yB = Y(MOUTH.fBottom);
  s.polyline(archPts(cx, a + MOUTH.softW, Y(MOUTH.fSoft), ys, yB), { color: MOUTH.soft, width: 1, closed: true, fill: MOUTH.soft });
  // 咽頭の入口（暗色のアーチ）。縁＝軟口蓋の縁＋左右の口蓋弓（青）
  s.polyline(archPts(cx, a, Y(MOUTH.fTop), ys, yB), { color: C.blue, width: 2.5, closed: true, fill: MOUTH.dark });
  // 口蓋垂（アーチの頂点から垂れる）
  const uw = MOUTH.uvTopHalfW, ur = MOUTH.uvTipR, yt = Y(MOUTH.fTop) - 0.01, yb = Y(MOUTH.fUvula) - ur;
  const uv = [[r3(cx - uw), r3(yt)], [r3(cx - ur), r3(yb)]];
  for (let k = 1; k < 12; k++) {
    const t = Math.PI - (Math.PI * k) / 12;
    uv.push([r3(cx + ur * Math.cos(t)), r3(yb + ur * Math.sin(t))]);
  }
  uv.push([r3(cx + ur), r3(yb)], [r3(cx + uw), r3(yt)]);
  s.polyline(uv, { color: C.blue, width: 2, closed: true, fill: MOUTH.soft });
  // 舌（上縁は両端が口の輪郭に接し、中央がわずかに盛り上がる）
  const yl = Y(MOUTH.levels[cls]);
  const dy = (yl - cy) / ry;
  const X = rx * Math.sqrt(1 - dy * dy);
  const tg = [];
  for (let k = 0; k <= 24; k++) {
    const x = -X + (2 * X * k) / 24;
    tg.push([r3(cx + x), r3(yl - MOUTH.dome * (1 - (x / X) ** 2))]);
  }
  const p0 = Math.asin(dy);
  for (let k = 1; k < 36; k++) {
    const t = p0 + ((Math.PI - 2 * p0) * k) / 36;
    tg.push([r3(cx + rx * Math.cos(t)), r3(cy + ry * Math.sin(t))]);
  }
  s.polyline(tg, { color: C.vermilion, width: 2, closed: true, fill: MOUTH.tongue });
  // 口唇（輪郭）
  s.polyline(ell, { color: MOUTH.lip, width: 2.5, closed: true });
}

async function main() {
  const CHAPTERS = ['目的と\n問診', '診察', 'マスク\n換気', '限界と\n計画', 'まとめ'];
  const chipSpec = '「目的と\\n問診」「診察」「マスク\\n換気」「限界と\\n計画」「まとめ」';
  if (!SPEC_MD.includes(chipSpec)) throw new Error('章チップのラベルが spec と違います');
  const deck = new Deck({
    file: OUT,
    chapters: CHAPTERS,
    author: 'Anes-lecture',
    deckTitle: '4-2 術前診察での気道評価',
  });

  // ===== S1 表紙 =====
  {
    const subtitle = '4-2 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    if (!SPEC_MD.includes('title=「術前診察での気道評価」')) throw new Error('S1: 表紙タイトルが spec と違います');
    deck.cover({ title: '術前診察での気道評価', subtitle, meta, notes: notesOf('S1') });
  }

  // 箇条書きのスライド（S2・S9 は 26pt、S7 は 1 行目が 26pt に入らないので 24pt）。行頭記号は bulletCol で黒に
  const BUL = { x: X0, w: 12.2 };

  // ===== S2 気道評価は何のため（箇条書き4行 + 強調ボックス） =====
  {
    const id = 'S2';
    const s = deck.slide({ title: '気道評価は何のため', chapter: 0, notes: notesOf(id), sources: sourcesFor(id) });
    const bl = bulletCol(s, [
      mark(q(id, '困難気道を『当てる』ためではなく『備える』ため'), ['『備える』', C.vermilion]),
      q(id, '予測できない困難気道は、必ず一定の割合である'),
      mark(q(id, '見つけたら：道具、人、方法（覚醒下）を前もって用意'), ['前もって用意', C.blue]),
      mark(q(id, '評価しない理由はない。全例で、同じ手順で'), ['全例で、同じ手順で', C.blue]),
    ], { ...BUL, y: 2.0, size: 26, gap: 0.24 });
    const bulBottom = bl.bottom;
    const lines = brk(id, '前回の挿管困難の記録は強い予測因子（ただし拾えるのは約 3 割）', '（ただし');
    const eh = emphH(2);
    const ey = BOT - eh;
    if (ey < bulBottom + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(ey - bulBottom)}in`);
    emphBox(s, ey, [mark(lines[0], ['強い予測因子', C.blue]), mark(lines[1], ['約 3 割', C.vermilion])], C.teal);
  }

  // ===== S3 問診で聞く・読む（2列カード） =====
  {
    const id = 'S3';
    const s = deck.slide({ title: '問診で聞く・読む', chapter: 0, notes: notesOf(id), sources: sourcesFor(id) });
    q(id, '記録を読む'); q(id, '患者に聞く');
    cardRow(s, [
      {
        title: '記録を読む', titleColor: C.blue,
        items: [
          brk(id, '前回の麻酔記録：喉頭展開の見え方、道具', '喉頭展開'),
          [q(id, '困難気道のカード・アラート')],
          [q(id, '頭頸部の手術・放射線治療')],
          brk(id, '気道の病気（腫瘍、狭窄、声帯麻痺）', '（腫瘍'),
        ],
      },
      {
        title: '患者に聞く', titleColor: C.vermilion,
        items: [
          [q(id, 'いびき、無呼吸、日中の眠気（OSA）')],
          brk(id, '首や顎が動きにくい病気（リウマチ、強直性脊椎炎）', '（リウマチ'),
          [q(id, '動揺歯、義歯、差し歯、前歯の突出')],
          brk(id, '『前の麻酔で喉を傷めた』『管が入りにくかった』', '『管が'),
        ],
      },
    ], { y: TOP, bottom: BOT, maxItemGap: 0.36 });
  }

  // ===== S4 診察：6 つの所見（表 3列 × 6行＋見出し行） =====
  {
    const id = 'S4';
    const s = deck.slide({ title: '診察：6 つの所見', chapter: 1, notes: notesOf(id), sources: sourcesFor(id) });
    const B = text => ({ text, bold: true });
    const cell = (str, ...breaks) => brk(id, str, ...breaks).join('\n');
    // 列幅は spec の目安（2.8 / 5.2 / 4.2）から、2 列目の「座位、開口、舌を出す、発声なし」と
    // 3 列目の「class III・IV（→ S5）」が 1 行に入るように配分しなおした（合計 12.2in）。
    // 表の高さ（7 行で 4.63in）が 1.85〜6.45in に 0.03in 入らないので、上端を 1.82in にする
    // （金下線の下端 1.58in との間は 0.24in）。
    const TY = 1.82;
    const t = tightTable(s, [
      [q(id, '所見'), q(id, 'どう見るか'), q(id, '困難を疑う目安')],
      [B(q(id, '開口')), q(id, '最大開口で上下の切歯の間'), cell('3 cm 未満（資料により 3〜4 cm）', '（資料')],
      [B(q(id, 'Mallampati')), q(id, '座位、開口、舌を出す、発声なし'), q(id, 'class III・IV（→ S5）')],
      [B(cell('甲状オトガイ間距離', '間距離')), cell('頸部を伸展し、甲状軟骨上縁〜オトガイ', '甲状軟骨'), cell('6〜6.5 cm 未満（指 3 本未満）', '（指')],
      [B(q(id, '上唇咬合テスト')), q(id, '下の歯で上唇を噛めるか'), q(id, '噛めない（class III）')],
      [B(q(id, '頸部の動き')), q(id, '顎を胸につけ、次に上を向く'), cell('伸展が小さい、痛みで動かない', '痛み')],
      [B(q(id, '首・顎・歯')), cell('首の太さ、小さい下顎、前歯の突出、ひげ', '前歯'), cell('太い首（特に肥満）、小顎、前突', '小顎')],
    ], { x: X0, y: TY, colW: [2.72, 5.36, 4.12] });
    console.log(`  p${s.page} 表: ${t.y}〜${r3(t.y + t.h)}in（行高 ${t.rowH.join(', ')}）`);
    if (t.y + t.h > BOT + 1e-6) throw new Error(`${id} 表の下端 ${r3(t.y + t.h)} > ${BOT}`);
  }

  // ===== S5 Mallampati 分類（4列カード＋各カードの上に口の模式図＋下に1行） =====
  {
    const id = 'S5';
    const s = deck.slide({ title: 'Mallampati 分類', chapter: 1, notes: notesOf(id), sources: sourcesFor(id) });
    // カード幅: 本文を各 3 行にそろえるため、class I・II（1 行に全角 8 字）を広く、III・IV を残りの等分にする
    // （4 枚とも 4 行にすると模式図の高さが 1in を切る）。文字は左右 0.15in の内側余白の中。
    const padX = 0.15, pad = 0.12, tgap = 0.04;
    const cards = [
      { cls: 1, title: q(id, 'class I'), color: C.blue, body: brk(id, '軟口蓋、口蓋弓、口蓋垂が全部見える', '口蓋垂', '全部') },
      { cls: 2, title: q(id, 'class II'), color: C.blue, body: brk(id, '軟口蓋、口峡、口蓋垂（口蓋弓は隠れる）', '口蓋垂', '隠れる') },
      { cls: 3, title: q(id, 'class III'), color: C.vermilion, body: brk(id, '軟口蓋と口蓋垂の根元だけ', '口蓋垂', '根元') },
      { cls: 4, title: q(id, 'class IV'), color: C.vermilion, body: brk(id, '軟口蓋が見えない（硬口蓋のみ）', '見えない', '（硬口蓋') },
    ];
    const needIn = c => Math.max(measure(c.title, 24), ...c.body.map(l => measure(l, 22)));
    const innerAll = BW - 3 * GAP - 8 * padX;
    const wide = Math.ceil(Math.max(needIn(cards[0]), needIn(cards[1])) * 1000) / 1000;
    const narrow = Math.floor(((innerAll - 2 * wide) / 2) * 1000) / 1000;
    if (narrow < Math.max(needIn(cards[2]), needIn(cards[3])) - 1e-9) throw new Error(`${id}: class III・IV のカード幅が足りません`);
    const inner = [wide, wide, narrow, narrow];
    const nLines = Math.max(...cards.map(c => c.body.length));
    const lineH = heightFor(1, 24, 0);
    const lineY = r3(BOT - lineH);
    const bodyH = heightFor(nLines, 22, 1);
    const titleH = heightFor(1, 24, 0);
    const cardH = r3(pad + titleH + tgap + bodyH + pad);
    const cardY = r3(lineY - GAP - cardH);
    const figTop = TOP, figH = r3(cardY - GAP - figTop), figW = 1.95;
    if (figH < 1.2) throw new Error(`${id}: 模式図の高さが ${figH}in しかありません`);
    let x = X0;
    cards.forEach((c, i) => {
      const cw = r3(inner[i] + 2 * padX);
      x = r3(x);
      mouthFig(s, { cx: r3(x + cw / 2), top: figTop, w: figW, h: figH, cls: c.cls });
      s.box({ x, y: cardY, w: cw, h: cardH, fill: C.bg, radius: 0.08 });
      s.text([c.title], { x: r3(x + padX), y: r3(cardY + pad), w: inner[i], h: titleH, size: 24, bold: true, color: c.color, align: 'center' });
      s.text(c.body, { x: r3(x + padX), y: r3(cardY + pad + titleH + tgap), w: inner[i], h: bodyH, size: 22, gap: 1, color: C.dark });
      x += cw + GAP;
    });
    s.text([mark(q(id, '座位・最大開口・舌を出す・発声なし。III・IV で困難を疑う'), ['III・IV で困難を疑う', C.vermilion])],
      { x: X0, y: lineY, w: BW, h: lineH, size: 24, align: 'center' });
    console.log(`  p${s.page} 模式図 ${figTop}〜${r3(figTop + figH)}in / カード ${cardY}〜${r3(cardY + cardH)}in（内側 ${inner.join(', ')}） / 1行 ${lineY}〜${r3(lineY + lineH)}in`);
  }

  // ===== S6 マスク換気困難の予測（2列カード + 強調ボックス） =====
  {
    const id = 'S6';
    // 出典 4 件（各 6.8in 前後）は、空白を詰めても 16pt で 2 行に入らない（約 27.8in ＞ 2 × 11.6in）。
    // 同じ著者・同じ誌の Kheterpal 3 件を「Kheterpal S, et al. Anesthesiology 2006;…, 2009;…, 2013;…」にまとめる
    // （文献は落とさない。deviation に記載）。
    const spec = SPEC[id].sources;
    const K = 'Kheterpal S, et al. Anesthesiology ';
    const ks = ['2006;105:885-891', '2009;110:891-897', '2013;119:1360-1369'];
    ks.forEach(k => { if (!spec.includes(K + k)) throw new Error(`${id} 出典に「${K + k}」がありません`); });
    const merged = spec.split(' / ').filter(it => !it.startsWith(K)).concat([K + ks.join(', ')]).join(' / ');
    const s = deck.slide({ title: 'マスク換気困難の予測', chapter: 2, notes: notesOf(id), sources: sourcesFor(id, merged) });
    // 語呂合わせ OBESE の頭文字を青の太字に（文言はそのまま）
    const ini = str => [{ text: str[0], color: C.blue, bold: true }, str.slice(1)];
    const obese = brk(id, 'Obese：BMI 26 超（Kheterpal は 30 以上）', '（Kheterpal');
    q(id, 'Langeron の 5 因子（OBESE）'); q(id, '『換気不能』の予測因子');
    const eh = emphH(1);
    const ey = r3(BOT - eh);
    cardRow(s, [
      {
        title: 'Langeron の 5 因子（OBESE）', titleColor: C.blue,
        items: [
          [ini(obese[0]), obese[1]],
          [ini(q(id, 'Bearded：ひげ'))],
          [ini(q(id, 'Elderly：55 歳超（同 57 歳以上）'))],
          [ini(q(id, 'Snorer：いびき'))],
          [ini(q(id, 'Edentulous：歯がない'))],
        ],
      },
      {
        title: '『換気不能』の予測因子', titleColor: C.vermilion,
        items: [
          [q(id, '頸部への放射線治療')],
          [q(id, '男性')],
          [q(id, '閉塞性睡眠時無呼吸')],
          [q(id, 'Mallampati III・IV')],
          [q(id, 'ひげ')],
        ],
      },
    ], { y: TOP, bottom: r3(ey - GAP) });
    emphBox(s, ey, [mark(q(id, '挿管困難とマスク換気困難が重なる患者が、最も危ない'), ['最も危ない', C.red])], C.red);
  }

  // ===== S7 予測の限界を知る（箇条書き5行） =====
  {
    const id = 'S7';
    const s = deck.slide({ title: '予測の限界を知る', chapter: 3, notes: notesOf(id), sources: sourcesFor(id) });
    bulletCol(s, [
      mark(q(id, '挿管困難は定義により約 2〜10%。9 割以上は予測されずに起きる'), ['9 割以上は予測されずに起きる', C.vermilion]),
      mark(q(id, '所見 1 つの感度は低い（Mallampati 約 5 割、TMD 約 2 割）'), ['感度は低い', C.blue]),
      mark(q(id, '特異度は高く、『異常あり』なら困難の確率が上がる'), ['特異度は高く', C.blue]),
      mark(q(id, '異常なし＝安全、ではない。全例で次の手を用意'), ['異常なし＝安全、ではない', C.red]),
      q(id, '評価の価値は『見つけたときに備えられる』こと'),
    ], { ...BUL, yc: r2((TOP + BOT) / 2), size: 24, gap: 0.4 });
  }

  // ===== S8 評価を計画につなぐ（2列カード + 強調ボックス） =====
  {
    const id = 'S8';
    const s = deck.slide({ title: '評価を計画につなぐ', chapter: 3, notes: notesOf(id), sources: sourcesFor(id) });
    q(id, '困難が予想されるとき'); q(id, '患者と記録');
    const eh = emphH(1);
    const ey = r3(BOT - eh);
    cardRow(s, [
      {
        title: '困難が予想されるとき', titleColor: C.vermilion,
        items: [
          [q(id, '上級医に相談、応援を頼む（遠隔でも）')],
          [q(id, 'ビデオ喉頭鏡を最初から')],
          [q(id, '声門上器具、ブジーを開けて置く')],
          [q(id, '覚醒下挿管（気管支鏡）を考える')],
          [q(id, '輪状甲状靱帯を触って印をつける')],
        ],
      },
      {
        title: '患者と記録', titleColor: C.blue,
        items: [
          brk(id, '患者に説明：歯の損傷、覚醒下の処置', '歯の'),
          [q(id, '前酸素化を十分に')],
          [q(id, '所見と計画を記録')],
          brk(id, '困難だったら：カード・アラート・紹介状', 'カード'),
        ],
      },
    ], { y: TOP, bottom: r3(ey - GAP) });
    emphBox(s, ey, [mark(q(id, '一人で対応できないと判断したら、転院・延期も選択肢'), ['転院・延期も選択肢', C.vermilion])], C.teal);
  }

  // ===== S9 まとめ（箇条書き5行） =====
  {
    const id = 'S9';
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: notesOf(id), sources: sourcesFor(id) });
    bulletCol(s, [
      mark(q(id, '気道評価は『備える』ため。全例で同じ手順で'), ['『備える』', C.vermilion]),
      q(id, '前回の記録は強い予測因子。問診でいびき・歯・頸部'),
      q(id, '開口、Mallampati III・IV、TMD 6 cm、上唇咬合、頸部'),
      mark(q(id, 'マスク換気困難：OBESE、放射線、OSA。重なると危険'), ['重なると危険', C.red]),
      mark(q(id, '異常なし≠安全。見つけたら道具・人・覚醒下・転院'), ['異常なし≠安全', C.red]),
    ], { ...BUL, yc: r2((TOP + BOT) / 2), size: 26, gap: 0.36 });
  }

  if (DEVIATIONS.length) console.log('[build] 出典の略記:\n  ' + DEVIATIONS.join('\n  '));
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.stack || e.message || e); process.exit(1); });
