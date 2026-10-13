'use strict';
/**
 * 2-4 パルスオキシメータ
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, fits, fitInfo, units, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '2-4_pulse_oximeter';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
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
      if (inOn || /^(figure|layout):/.test(ln)) {
        for (const q of ln.matchAll(/「([^」]*)」/g)) quoted.add(q[1]);
      }
    }
    out[id] = { notes, sources: sources && sources !== '—' ? sources : null, quoted };
  }
  return out;
}
const SPEC = parseSpec(fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8'));
function notesOf(id) {
  const n = SPEC[id] && SPEC[id].notes;
  if (!n) throw new Error(`spec.md の ${id} に notes がありません`);
  return n;
}

/**
 * spec の on-slide（または figure/layout の「」）の文言 str を、breaks の各語の直前で改行して行配列にする。
 * str が spec に無い場合は Error（文言を変えていないことの確認）。改行位置の半角スペースは落とす。
 */
function brk(id, str, ...breaks) {
  if (!SPEC[id].quoted.has(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
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
/** spec の文言そのまま（照合のみ） */
function q(id, str) { return brk(id, str)[0]; }

/**
 * 出典。既定は ' / ' 区切りの項目（ライブラリが詰めて最大2行に）。
 * breakBefore を渡すと、その語の直前（空白）で2行に割る（文言は変えない）。
 * text を渡すと spec の代わりにその文字列を使う（短縮版。deviations に記載）。
 */
function sourcesOf(id, breakBefore, text) {
  const s = text || (SPEC[id] && SPEC[id].sources);
  if (!s) return undefined;
  if (!breakBefore) return s.split(' / ');
  const k = s.indexOf(breakBefore);
  if (k <= 0) throw new Error(`${id} 出典: 改行位置「${breakBefore}」が見つかりません`);
  let l1 = s.slice(0, k), l2 = s.slice(k);
  if (l1.endsWith(' / ')) l1 = l1.slice(0, -3);
  else if (l1.endsWith(' ')) l1 = l1.slice(0, -1);
  else throw new Error(`${id} 出典: 改行位置は空白の直後にしてください`);
  const S = LAYOUT.source;
  if (!fits(l1, S.size, S.w) || !fits(l2, S.size, S.w)) throw new Error(`${id} 出典: 2行に収まりません\n  ${l1}\n  ${l2}`);
  return [l1, l2];
}
/** 出典を短縮したスライドは、ノートの末尾に spec の出典全文を残す */
function notesWithFullSources(id) {
  return notesOf(id) + '\n出典（全文）: ' + SPEC[id].sources;
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const em = (text, color, bold = true) => ({ text, color, bold });
const r3 = v => Math.round(v * 1000) / 1000;
const AXIS_C = C.gray;
const LEAD_C = C.gray;
// ティールの文字用。C.teal（00A8AA）は白地で約 2.9:1 と薄いので、文字には濃い 00807F（白地で約 4.8:1）を使う。
// 図形の塗り・線（S3 の点、S5 の「表示」チップの枠）は C.teal のまま。
const TEAL_TXT = C.teal;   // ライブラリ既定（全講義で共通。変えるなら decklib 側で）

/**
 * 箇条書き（1 項目が 2〜3 行になる場合用）。行頭記号と本文を別の枠に置き、折り返し行を本文の頭にそろえる。
 * items: [[行, 行...], ...]（行は文字列 or run 配列）。戻り値: 全体の高さ。
 */
function bulletBlockHeight(items, size, igap) {
  return items.reduce((a, lines) => a + heightFor(lines.length, size, 0), 0) + igap * (items.length - 1);
}
function bulletBlock(s, items, o) {
  const size = o.size || 24, igap = o.igap != null ? o.igap : 0.3, ind = o.indent || 0.42;
  let y = o.y;
  items.forEach(lines => {
    const h = heightFor(lines.length, size, 0);
    s.text('•', { x: o.x, y, w: ind, h: heightFor(1, size, 0), size, color: o.color });
    s.text(lines, { x: o.x + ind, y, w: o.w - ind, h, size, color: o.color });
    y += h + igap;
  });
  return y - igap - o.y;
}

/** 右向きの太い矢印（ネイティブ freeform の塗り図形）。x0→x1、中心の高さ yc。 */
function blockArrow(s, x0, x1, yc, color = '595959') {
  const shaft = 0.09, head = 0.24, hl = Math.min(0.22, (x1 - x0) * 0.55);
  s.polyline([
    [x0, yc - shaft], [x1 - hl, yc - shaft], [x1 - hl, yc - head], [x1, yc],
    [x1 - hl, yc + head], [x1 - hl, yc + shaft], [x0, yc + shaft],
  ], { closed: true, fill: color, color, width: 0.75 });
}

/**
 * 表（S7 用）。decklib の s.table と同じ見た目（見出し行 D9D9D9・太字・中央、罫線 BFBFBF 1pt、上下中央）で、
 * セル左右余白 mIn（インチ）を取り、文字幅は実機メイリオの実寸（全角 = 1em）で検査する。
 * s.table はセル幅を lint の割増式（全角 ×1.12）で検査するため、「静脈の拍動（強い締め付け、三尖弁逆流）」
 * （全角19字・22pt = 実寸 5.81in）が左列の内幅 6.0in に入るのに 2 行に割らされる。lint（slide_lint.py）は
 * 表のセルを検査しないので、ここでは実寸 + 余裕 SAFE で検査し、重なり検査用の登録は decklib と同じく行う。
 */
function realTable(s, rows, o) {
  const size = o.size || 22, m = o.mIn, colW = o.colW;
  const LINE = (size / 72) * 1.5;   // 実機メイリオの 1 行高（decklib の MEIRYO_LINE と同じ）
  const SAFE = 0.12;                // 実寸幅に対する余裕（インチ）
  const w = colW.reduce((a, b) => a + b, 0);
  const rowH = [];
  const out = rows.map((row, ri) => {
    const head = ri === 0;
    let maxLines = 1;
    const cells = row.map((cell, ci) => {
      const c = typeof cell === 'string' ? { text: cell } : cell;
      const lines = c.text.split('\n');
      const inner = colW[ci] - 2 * m;
      lines.forEach(l => {
        const real = units(l) * size / 72;
        if (real + SAFE > inner + 1e-9) {
          throw new Error(`${s.page ? 'p' + s.page : ''} 表 行${ri + 1}列${ci + 1}「${l}」がセル幅に収まりません（実寸 ${r3(real)}in + 余裕 ${SAFE} > 内幅 ${r3(inner)}in）`);
        }
      });
      maxLines = Math.max(maxLines, lines.length);
      const color = c.color || C.dark, bold = c.bold != null ? c.bold : head;
      return {
        text: lines.map((l, li) => ({
          text: l,
          options: { fontFace: FONT, lang: 'ja-JP', fontSize: size, color, bold: bold || undefined, breakLine: li < lines.length - 1 || undefined },
        })),
        options: {
          fill: head ? { color: C.lightGray } : undefined, align: head ? 'center' : 'left', valign: 'middle',
          fontFace: FONT, fontSize: size, color, bold: bold || undefined,
        },
      };
    });
    rowH.push(Math.ceil(maxLines * LINE * 100) / 100);
    return cells;
  });
  const h = rowH.reduce((a, b) => a + b, 0);
  s._bounds(`p${s.page} 表`, o.x, o.y, w, h);
  // pptxgenjs は margin[0] < 1 のとき値をインチとして扱う（上下 0・左右 mIn インチ）
  s.raw.addTable(out, {
    x: o.x, y: o.y, w, colW, rowH, fontFace: FONT, fontSize: size, margin: [0, m, 0, m],
    border: { type: 'solid', pt: 1, color: C.border }, valign: 'middle',
  });
  s._reg('box', o.x, o.y, w, h, 'table');
  return { x: o.x, y: o.y, w, h, rowH };
}

/** 決定的な疑似乱数（悪い脈波のゆらぎ用） */
function rng(seed) {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6D2B79F5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['原理', '数値', '遅れ', '誤差', '使い方'],
    author: 'Anes-lecture',
    deckTitle: '2-4 パルスオキシメータ',
  });

  // ===== S1 表紙 =====
  deck.cover({
    title: 'パルスオキシメータ',
    subtitle: '2-4 モニタリング',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 何を測っているか（左: 指先センサーの模式図 / 右: 箇条書き4行） =====
  {
    const id = 'S2';
    const s = deck.slide({ title: '何を測っているか', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    // --- 模式図（ラベルはすべて図の外: 発光部の上と左の列） ---
    const labH = heightFor(1, 22, 0);
    const eH = 0.45, eW = 0.55;           // 発光部の四角
    const fH = 1.6, rH = 0.42;            // 指（横長の角丸）・受光部の高さ
    const figH = labH + 0.26 + eH + fH + rH;
    const yL = r3((1.85 + 6.75) / 2 - figH / 2);   // 発光部ラベル（赤色光・赤外光）の上端（本文ゾーンの上下中央）
    const eY = yL + labH + 0.26;          // 発光部の上端
    const fY = eY + eH;
    const fX = 1.9, fW = 3.0;
    const rY = fY + fH;                   // 受光部
    const cxR = 2.7, cxI = 4.1;           // 赤色・赤外の中心 x
    s.box({ x: fX, y: fY, w: fW, h: fH, fill: C.bg, line: { color: C.gray, width: 1.5 }, shape: 'roundRect', radius: 0.7 });
    s.box({ x: cxR - eW / 2, y: eY, w: eW, h: eH, fill: C.vermilion });
    s.box({ x: cxI - eW / 2, y: eY, w: eW, h: eH, fill: C.gray });
    s.box({ x: cxR - 0.45, y: rY, w: cxI - cxR + 0.9, h: rH, fill: '404040' });
    s.arrow({ x1: cxR, y1: eY + eH + 0.06, x2: cxR, y2: rY - 0.04, color: C.vermilion, width: 3.5 });
    s.arrow({ x1: cxI, y1: eY + eH + 0.06, x2: cxI, y2: rY - 0.04, color: C.gray, width: 3.5 });
    const lw = 1.1;
    s.label('赤色光', { x: cxR - lw / 2, y: yL, w: lw, h: labH, color: C.vermilion, bold: true });
    s.label('赤外光', { x: cxI - lw / 2, y: yL, w: lw, h: labH, color: '595959', bold: true });
    const lx = X0, lW = 1.1;              // 左の列（部位名）
    s.label('発光部', { x: lx, y: eY + eH / 2 - labH / 2, w: lW, h: labH, align: 'right' });
    s.label('指', { x: lx, y: fY + fH / 2 - labH / 2, w: lW, h: labH, align: 'right' });
    s.label('受光部', { x: lx, y: rY + rH / 2 - labH / 2, w: lW, h: labH, align: 'right' });
    // --- 箇条書き ---
    const items = [
      brk(id, '赤色光（660 nm）と赤外光（940 nm）の吸収の比', '赤外光'),
      [q(id, '脈で変わる成分（動脈血）だけを取り出す')],
      brk(id, 'SpO2 ＝ 酸素化ヘモグロビンの割合（機能的飽和度）', '（機能的'),
      [q(id, '脈波（プレチスモグラム）も表示する')],
    ];
    const igap = 0.4, bx = 5.25;
    const bh = bulletBlockHeight(items, 24, igap);
    const figMid = (yL + rY + rH) / 2;
    bulletBlock(s, items, { x: bx, y: r3(figMid - bh / 2), w: XR - bx, size: 24, igap });
  }

  // ===== S3 SpO2 と PaO2 の関係（左: 酸素解離曲線 / 右: 箇条書き3行） =====
  {
    const id = 'S3';
    const s = deck.slide({ title: 'SpO2 と PaO2 の関係', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    const sat = P => (P <= 0 ? 0 : 1 / (23400 / (P * P * P + 150 * P) + 1));   // Severinghaus 1979
    const labH = heightFor(1, 22, 0);
    const row1 = 1.88, row2 = row1 + labH + 0.14;       // 図の上の2段（点のラベル）
    const ax0 = 1.35, ax1 = 6.65, y0 = 5.58;            // 軸: PaO2 0→ax0, 100→ax1 / SaO2 0%→y0
    const pA = 60, pB = 90;
    const yTopA = row2 + labH + 0.42;                   // 点 A（PaO2 60）の高さ
    const H = (y0 - yTopA) / sat(pA);                   // SaO2 100% ぶんの高さ
    const X = P => ax0 + (ax1 - ax0) * P / 100;
    const Y = S => y0 - H * S;
    // 軸
    s.arrow({ x1: ax0, y1: y0, x2: ax0, y2: Y(1) - 0.3, color: AXIS_C, width: 2 });
    s.arrow({ x1: ax0, y1: y0, x2: ax1 + 0.3, y2: y0, color: AXIS_C, width: 2 });
    // 補助の点線（PaO2 60 / SaO2 90）
    const xA = X(pA), yA = Y(sat(pA)), xB = X(pB), yB = Y(sat(pB)), d = 0.2;
    s.vline({ x: xA, y: yA + d / 2 + 0.04, h: y0 - yA - d / 2 - 0.04, color: C.vermilion, width: 1.75, dash: 'dash' });
    s.hline({ x: ax0, y: yA, w: xA - ax0 - d / 2 - 0.04, color: C.vermilion, width: 1.75, dash: 'dash' });
    // 曲線（S 字）
    const pts = [];
    for (let P = 0; P <= 100; P += 1) pts.push([X(P), Y(sat(P))]);
    s.polyline(pts, { color: C.blue, width: 3.5 });
    // 2点
    s.circle({ cx: xA, cy: yA, d, fill: C.vermilion });
    s.circle({ cx: xB, cy: yB, d, fill: C.teal });
    // 目盛りのラベル（軸の外）
    const tw = 0.7;
    s.label('90', { x: ax0 - 0.12 - tw, y: yA - labH / 2, w: tw, h: labH, align: 'right', color: C.vermilion, bold: true });
    s.label('0', { x: ax0 - 0.12 - tw, y: y0 - labH / 2, w: tw, h: labH, align: 'right' });
    const tickY = y0 + 0.1;
    s.label('60', { x: xA - tw / 2, y: tickY, w: tw, h: labH, color: C.vermilion, bold: true });
    s.label('100', { x: ax1 - tw / 2, y: tickY, w: tw, h: labH });
    // 軸の名前
    s.text('SaO2（%）', { x: X0, y: row1, w: 2.0, h: labH, size: 22 });
    const xtW = 3.0;
    // 目盛り（60・100）との間を 0.18in あける（下端 6.78 < 出典 6.85）
    s.text('PaO2（mmHg）', { x: ax1 + 0.3 - xtW, y: tickY + labH + 0.18, w: xtW, h: labH, size: 22, align: 'right' });
    // 点のラベル（図の上。引き出し線で点へ）
    const lw = 3.0;
    const lA = { x: xA + 0.45 - lw, y: row2 }, lB = { x: Math.min(xB - lw / 2, ax1 + 0.3 - lw), y: row1 };
    s.text('PaO2 60 / SpO2 90', { x: lA.x, y: lA.y, w: lw, h: labH, size: 22, bold: true, color: C.vermilion, align: 'right' });
    s.text('PaO2 90 / SpO2 97', { x: lB.x, y: lB.y, w: lw, h: labH, size: 22, bold: true, color: TEAL_TXT, align: 'center' });
    s.polyline([[xA, row2 + labH + 0.05], [xA, yA - d / 2 - 0.04]], { color: C.vermilion, width: 1.5 });
    s.polyline([[xB, row1 + labH + 0.05], [xB, yB - d / 2 - 0.04]], { color: TEAL_TXT, width: 1.5 });
    // --- 箇条書き（右 1/3） ---
    const items = [
      [[em(q(id, 'SpO2 90% ≈ PaO2 60 mmHg'), C.vermilion)]],
      brk(id, '100% では PaO2 は分からない（200 でも 500 でも 100%）', '（200'),   // 1・2・3 行にして間隔の差を小さく
      brk(id, '曲線の肩より上では、PaO2 が下がっても SpO2 は動かない', 'PaO2', 'SpO2'),
    ];
    const igap = 0.42, bx = 7.3;
    const bh = bulletBlockHeight(items, 24, igap);
    const mid = (row1 + tickY + labH) / 2;
    bulletBlock(s, items, { x: bx, y: r3(mid - bh / 2), w: XR - bx, size: 24, igap });
  }

  // ===== S4 精度（左: 大きな数値 / 右: 箇条書き4行） =====
  {
    const id = 'S4';
    // 出典: spec の3件は 16pt・2行に収まらない（102.6字相当 > 2行 93.2字相当）。
    // FDA ガイダンスの題名から「– Premarket Notification Submissions」を省き、全文はノート末尾に残す。
    const short = SPEC[id].sources.replace('Pulse Oximeters – Premarket Notification Submissions [510(k)s]', 'Pulse Oximeters [510(k)s]');
    if (short === SPEC[id].sources) throw new Error('S4 出典の短縮に失敗');
    const s = deck.slide({ title: '精度', chapter: 1, notes: notesWithFullSources(id), sources: sourcesOf(id, 'N Engl J Med', short) });
    const by = 2.15, bw = 4.13, bh = 3.95, tm = 0.1;
    s.box({ x: X0, y: by, w: bw, h: bh, fill: C.bg, radius: 0.1 });
    const num = brk(id, 'Arms 2〜3 %', '2〜3');      // 1 行だと箱の幅ぎりぎりなので「Arms」と数値で 2 行に
    const hNum = heightFor(num.length, 44, 0);
    const cap = brk(id, 'SaO2 70〜100% で、測定の約 2/3 がこの幅に入る', '測定の', 'この幅');
    const hCap = heightFor(cap.length, 22, 0), gNum = 0.3;
    // 上下中央。44pt の行は字面の上に余白が大きいぶん、0.07in 上げて見た目の上下の余白をそろえる
    const ny = by + (bh - hNum - gNum - hCap) / 2 - 0.07;
    // 各枠の中で上下中央（PowerPoint の行高でも代替フォントの描画でも、箱の中で上下の余白がそろう）
    s.text(num, { x: X0 + tm, y: ny, w: bw - 2 * tm, h: hNum, size: 44, bold: true, color: TEAL_TXT, align: 'center', valign: 'middle' });
    s.text(cap, { x: X0 + tm, y: ny + hNum + gNum, w: bw - 2 * tm, h: hCap, size: 22, align: 'center', valign: 'middle' });
    const lx = X0 + bw + 0.24, gap = 22, lh = heightFor(4, 24, gap);
    s.bullets([
      [em(q(id, '70% 未満は保証されていない'), C.vermilion)],
      q(id, '低い値ほど誤差が大きい'),
      q(id, '皮膚の色が濃いと、高めに出ることがある'),
      q(id, '動脈血ガスと合わないときは血ガスを信じる'),
    ], { x: lx, y: r3(by + bh / 2 - lh / 2), w: XR - lx, h: lh, size: 24, gap });
  }

  // ===== S5 表示の遅れ（上: フロー4チップ / 下: 箇条書き4行） =====
  {
    const id = 'S5';
    const s = deck.slide({ title: '表示の遅れ', chapter: 2, notes: notesOf(id), sources: sourcesOf(id, 'Anaesthesia 2005') });
    const chips = [
      brk(id, '肺の酸素化が落ちる', '落ちる'),
      brk(id, '血液が指に届く（循環時間）', '（循環'),
      brk(id, '数秒〜十数秒の平均化', '平均化'),
      [q(id, '表示')],
    ];
    const g = 0.76, wLast = 1.3, w = (BW - wLast - 3 * g) / 3, cy = 1.95, ch = 1.25;
    let x = X0;
    chips.forEach((lines, i) => {
      const cw = i === 3 ? wLast : w;
      const last = i === 3;
      s.box({ x, y: cy, w: cw, h: ch, fill: last ? 'E6F6F6' : C.bg, line: { color: last ? C.teal : C.blue, width: 2 },
        radius: 0.1, text: lines.join('\n'), size: 24, bold: last, color: C.dark });
      if (i < 3) blockArrow(s, x + cw + 0.24, x + cw + g - 0.24, cy + ch / 2);
      x += cw + g;
    });
    const gap = 16, lh = heightFor(4, 24, gap);
    s.bullets([
      q(id, '指は耳や額より遅れる（循環時間が長い）'),
      q(id, '平均化の時間を短くすると早いが、ノイズも増える'),
      q(id, '低灌流・冷え・昇圧薬で遅れが伸びる'),
      ['無呼吸は、SpO2 が下がる前に ', em('EtCO2 で気づく', TEAL_TXT)],
    ], { x: X0, y: 3.75, w: 12.2, h: lh, size: 24, gap });
    q(id, '無呼吸は、SpO2 が下がる前に EtCO2 で気づく');
  }

  // ===== S6 酸素予備と低下の速さ（左: 模式図 / 右: 箇条書き3行） =====
  {
    const id = 'S6';
    // 出典: 論文題名込みで 16pt・2行に収まらない（94.2字相当 > 93.2）。
    // 題名中の「intravenous」を「IV」に略し、全文はノート末尾に残す。
    const short = SPEC[id].sources.replace('1 mg/kg intravenous succinylcholine', '1 mg/kg IV succinylcholine');
    if (short === SPEC[id].sources) throw new Error('S6 出典の短縮に失敗');
    const s = deck.slide({ title: '酸素予備と低下の速さ', chapter: 2, notes: notesWithFullSources(id), sources: sourcesOf(id, 'unparalyzed', short) });
    const labH = heightFor(1, 22, 0);
    const ax0 = 1.2, ax1 = 7.2, y0 = 5.6, yTop = 2.85, yEnd = 4.55;
    s.text('SpO2', { x: X0, y: 1.88, w: 1.4, h: labH, size: 22 });
    s.arrow({ x1: ax0, y1: y0, x2: ax0, y2: 2.58, color: AXIS_C, width: 2 });
    s.arrow({ x1: ax0, y1: y0, x2: ax1, y2: y0, color: AXIS_C, width: 2 });
    s.text('時間', { x: ax1 - 1.0, y: y0 + 0.12, w: 1.0, h: labH, size: 22, align: 'right' });
    // 曲線: 高い値で平坦 → 肩 → 急に下がる
    const curve = (y1, xFall, L) => {
      const pts = [[ax0 + 0.06, y1], [xFall, y1]];
      for (let k = 1; k <= 24; k++) {
        const u = k / 24;
        pts.push([xFall + L * u, y1 + (yEnd - y1) * Math.pow(u, 2.2)]);
      }
      return pts;
    };
    const L = 1.45;
    const xNo = ax0 + 0.7, xYes = ax0 + 3.35;
    s.polyline(curve(yTop + 0.24, xNo, L), { color: C.vermilion, width: 3.5 });   // 始点の高さを 0.24in ずらして重ならないように
    s.polyline(curve(yTop, xYes, L), { color: C.blue, width: 3.5 });
    // ラベル（曲線の下端のすぐ下。図の線にかけない）
    const lw = 2.3, ly = yEnd + 0.24;
    s.text('前酸素化なし', { x: xNo + L - lw / 2, y: ly, w: lw, h: labH, size: 22, bold: true, color: C.vermilion, align: 'center' });
    s.text('前酸素化あり', { x: xYes + L - lw / 2, y: ly, w: lw, h: labH, size: 22, bold: true, color: C.blue, align: 'center' });
    // --- 箇条書き ---
    const items = [
      brk(id, '前酸素化で、無呼吸に耐えられる時間が延びる', '耐えられる'),
      brk(id, '肥満・小児・妊婦・病気の肺は早く落ちる', '病気の'),
      brk(id, '落ち始めると急に下がる（曲線の肩を越える）', '（曲線'),
    ];
    const igap = 0.42, bx = 7.6;
    const bh = bulletBlockHeight(items, 24, igap);
    const mid = (1.88 + y0 + 0.12 + labH) / 2;
    bulletBlock(s, items, { x: bx, y: r3(mid - bh / 2), w: XR - bx, size: 24, igap });
  }

  // ===== S7 誤差の要因（表 2列 × 8行＋見出し行） =====
  {
    const id = 'S7';
    // 出典: spec の5件は 16pt で 3 行分（134.0字相当）。表（見出し＋8行、22pt）は高さ 4.60in が要り、
    // 金下線・出典との間を 0.24in 以上あけるには出典を 1 行（y6.85）に収める必要がある。
    // 「第一著者 イニシャル 年」に短縮し、全文はノート末尾に残す。
    const short = 'Barker SJ 1987 / Barker SJ 1989 / Coté CJ 1988 / Hinkelbein J 2007 / Scheller MS 1986';
    if (!fits(short, LAYOUT.source.size, LAYOUT.source.w)) throw new Error('S7 出典: 1行に収まりません');
    const s = deck.slide({ title: '誤差の要因', chapter: 3, notes: notesWithFullSources(id), sources: [short] });
    const c = t => q(id, t);
    const rows = [
      [c('要因'), c('表示への影響')],
      [c('体動・シバリング'), c('値が乱れる。脈拍数が心電図と合わない')],
      [c('低灌流（冷え・昇圧薬・カフの加圧）'), c('測れない、遅れる')],
      [c('周囲の強い光・電気メス'), c('ノイズ')],
      [c('濃い色のマニキュア・人工爪'), c('やや低めに出る、測れないことがある')],
      [brk(id, '色素（メチレンブルー、インドシアニングリーン）', 'インドシアニン').join('\n'), c('一過性に低く出る')],
      [c('一酸化炭素ヘモグロビン'), { text: c('高く出る（実際より良く見える）'), color: C.vermilion, bold: true }],
      [c('メトヘモグロビン'), c('85% 付近に近づく')],
      [c('静脈の拍動（強い締め付け、三尖弁逆流）'), c('低く出る')],
    ];
    // 列幅: 左 6.2in（全角19字 = 実寸 5.81in が内幅 6.0in に入る）・右 6.03in。セル左右余白 0.1in、上下 0。
    // 2 行になるのは「色素」の行だけ（全角21字は左列に入らない）→ 高さ 4.60in。
    // 上端 1.90（金下線 1.52 から 0.38）・下端 6.50（出典 6.85 まで 0.35）。
    const c1 = 6.2;
    const t = realTable(s, rows, { x: X0, y: 1.9, colW: [c1, r3(BW - c1)], size: 22, mIn: 0.1 });
    if (t.y + t.h > LAYOUT.source.y - 0.24) throw new Error(`S7 表の下端 ${r3(t.y + t.h)}in が出典に近すぎます`);
  }

  // ===== S8 脈波と脈拍数も見る（上: 波形2つ / 下: 箇条書き4行） =====
  {
    const id = 'S8';
    const s = deck.slide({ title: '脈波と脈拍数も見る', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    const gap = 0.5, pw = (BW - gap) / 2, py = 1.85, ph = 1.45;
    const pxL = X0, pxR = X0 + pw + gap;
    s.box({ x: pxL, y: py, w: pw, h: ph, fill: C.bg, radius: 0.08 });
    s.box({ x: pxR, y: py, w: pw, h: ph, fill: C.bg, radius: 0.08 });
    // 良い脈波: 規則的 4 拍、立ち上がりが急、下降脚に小さな切痕
    {
      const mx = 0.3, top = py + 0.2, yb = py + ph - 0.2, A = yb - top;
      const beat = (pw - 2 * mx) / 4;
      const shape = [[0, 0], [0.05, 0.3], [0.09, 0.75], [0.12, 0.95], [0.15, 1.0], [0.19, 0.96], [0.27, 0.8],
        [0.35, 0.62], [0.4, 0.52], [0.43, 0.49], [0.47, 0.55], [0.51, 0.56], [0.57, 0.5], [0.67, 0.37],
        [0.79, 0.23], [0.9, 0.12], [1.0, 0.04]];
      const pts = [];
      for (let b = 0; b < 4; b++) shape.forEach(([t, v], k) => { if (b === 0 || k > 0) pts.push([pxL + mx + (b + t) * beat, yb - A * v]); });
      s.polyline(pts, { color: C.blue, width: 3.5 });
    }
    // 悪い脈波: 不規則で振幅が小さい（ゆらぎ＋ノイズ）
    {
      const rand = rng(24);
      const mx = 0.3, yMid = py + ph / 2 + 0.05, A = ph - 0.4;
      const x0 = pxR + mx, x1 = pxR + pw - mx;
      const pts = [];
      let x = x0;
      const lens = [0.95, 0.6, 1.25, 0.75, 1.05, 0.55];
      const amps = [0.32, 0.18, 0.4, 0.22, 0.3, 0.15];
      let bi = 0;
      while (x < x1 - 1e-6 && bi < lens.length) {
        const Lb = Math.min(lens[bi], x1 - x), amp = amps[bi];
        const n = Math.max(6, Math.round(Lb / 0.05));
        for (let k = (bi === 0 ? 0 : 1); k <= n; k++) {
          const t = k / n;
          const xx = x + Lb * t;
          const wander = 0.07 * Math.sin((xx - x0) * 1.3);
          const v = amp * Math.pow(Math.sin(Math.PI * t), 2) + wander + (rand() - 0.5) * 0.09;
          pts.push([xx, yMid + 0.2 * A - A * v]);
        }
        x += Lb; bi++;
      }
      s.polyline(pts, { color: C.vermilion, width: 3 });
    }
    // 図の下のラベル
    const labH = heightFor(1, 24, 0), ly = py + ph + 0.2;
    s.text(q(id, '良い脈波：数値を信じてよい'), { x: pxL, y: ly, w: pw, h: labH, size: 24, bold: true, color: C.blue, align: 'center' });
    s.text(q(id, '悪い脈波：数値を信じない'), { x: pxR, y: ly, w: pw, h: labH, size: 24, bold: true, color: C.vermilion, align: 'center' });
    const gapB = 10, lh = heightFor(4, 24, gapB);
    s.bullets([
      q(id, '脈波の形がきれいなときだけ、数値を信じる'),
      q(id, '脈拍数が心電図の心拍数と合っているか'),
      q(id, '灌流指標（PI）が低ければ、部位を変える・温める'),
      ['脈波の呼吸性変動は輸液反応性のヒント', em('（→ 9-8）', TEAL_TXT)],
    ], { x: X0, y: r3(ly + labH + 0.32), w: 12.2, h: lh, size: 24, gap: gapB });
    q(id, '脈波の呼吸性変動は輸液反応性のヒント（→ 9-8）');
  }

  // ===== S9 現場での使い方（箇条書き5行） =====
  {
    const id = 'S9';
    const s = deck.slide({ title: '現場での使い方', chapter: 4, notes: notesOf(id), sources: sourcesOf(id, 'Basic Anesthetic') });
    s.bullets([
      q(id, 'センサーは指・耳・額。血圧計と同じ腕は避ける'),
      q(id, 'アラームの下限は患者ごとに決める'),
      q(id, '下がったら：センサー → 患者 → 回路の順に確かめる'),
      ['酸素投与中は、', em('SpO2 が保たれたまま低換気が進む', C.vermilion)],
      q(id, '長時間は部位を変える（圧迫・熱傷を防ぐ）'),
    ], { y: 2.05, size: 28, gap: 24 });
    q(id, '酸素投与中は、SpO2 が保たれたまま低換気が進む');
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const id = 'S10';
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      q(id, '2つの光の吸収の比で、動脈血の酸素飽和度を推定する'),
      q(id, 'SpO2 90% ≈ PaO2 60 mmHg。100% では余裕は分からない'),
      q(id, '誤差（Arms）は 2〜3%。低い値と濃い皮膚で増える'),
      ['表示は数十秒遅れる。無呼吸は ', em('EtCO2 で気づく', TEAL_TXT)],
      ['脈波の形と脈拍数が合うときだけ、', em('数値を信じる', C.vermilion)],
    ], { y: 2.05, size: 28, gap: 24 });
    q(id, '表示は数十秒遅れる。無呼吸は EtCO2 で気づく');
    q(id, '脈波の形と脈拍数が合うときだけ、数値を信じる');
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
