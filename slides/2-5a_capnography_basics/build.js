'use strict';
/**
 * 2-5a カプノグラフィ①：原理と正常波形
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言はここに転記（改行位置のみ調整）、ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, fits, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '2-5a_capnography_basics';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

// ---------------------------------------------------------------------------
// spec.md からノートと出典を読む（原文のまま）
// ---------------------------------------------------------------------------
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const lines = part.split('\n');
    let notes = null, sources = null;
    for (let i = 0; i < lines.length; i++) {
      if (/^notes:\s*\|\s*$/.test(lines[i])) {
        const buf = [];
        for (let j = i + 1; j < lines.length && (lines[j].startsWith('  ') || lines[j] === ''); j++) {
          buf.push(lines[j].replace(/^ {2}/, ''));
        }
        while (buf.length && buf[buf.length - 1] === '') buf.pop();
        notes = buf.join('\n');
      }
      const m = lines[i].match(/^sources:\s*(.*)$/);
      if (m) sources = m[1].trim();
    }
    out[id] = { notes, sources: sources && sources !== '—' ? sources : null };
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
 * 出典: spec の sources 行を decklib に渡す項目に変換する。
 * 既定は ' / ' 区切りの項目（ライブラリが詰めて最大2行に）。
 * 1項目が1行に収まらない場合は breakBefore の位置（空白か ' / '）で2行に割る（文言は変えない）。
 */
function sourcesOf(id, breakBefore) {
  const s = SPEC[id] && SPEC[id].sources;
  if (!s) return undefined;
  if (!breakBefore) return s.split(' / ');
  const k = s.indexOf(breakBefore);
  if (k <= 0) throw new Error(`${id} 出典: 改行位置「${breakBefore}」が見つかりません`);
  let l1 = s.slice(0, k), l2 = s.slice(k);
  if (l1.endsWith(' / ')) l1 = l1.slice(0, -3);
  else if (l1.endsWith(' ')) l1 = l1.slice(0, -1);
  else throw new Error(`${id} 出典: 改行位置は空白の直後にしてください`);
  const S = LAYOUT.source;
  if (!fits(l1, S.size, S.w) || !fits(l2, S.size, S.w)) throw new Error(`${id} 出典: 2行に収まりません`);
  return [l1, l2];
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const em = (text, color, bold = true) => ({ text, color, bold });
const r3 = v => Math.round(v * 1000) / 1000;

/**
 * 文字に使うティール。C.teal(00A8AA) は白地で約 2.9:1、F2F2F2 地で約 2.6:1 と投影で読みにくいので、
 * 文字・現在章チップには濃いティール（白地 約 4.8:1、F2F2F2 地 約 4.3:1）を使う。
 */
const TEAL = C.teal;   // 書式ルールのティール（シリーズ共通）

/**
 * 章チップの文字を大きく・現在章を濃いティールにする（この講義だけの上書き）。
 * decklib は 11pt・幅 0.68in で置く。slide_lint の小サイズ許容は 10.5/12/16pt なので 16pt にし、
 * 3字ラベル（使い方・まとめ）が実機メイリオでも 1 行に収まるよう幅を 0.76in に広げる（右端・間隔は同じ）。
 * pptxgenjs のスライドオブジェクト（書き出し前）を直接書き換える。
 */
const CHIP_PT = 16, CHIP_W = 0.76;
function restyleChips(s) {
  const ch = LAYOUT.chip;
  const chips = s.raw._slideObjects.filter(o => o._type === 'text' && o.options && o.options.shape === 'roundRect'
    && Math.abs(o.options.y - ch.y) < 1e-6 && Math.abs(o.options.h - ch.h) < 1e-6);
  if (chips.length !== 5) throw new Error(`p${s.page}: 章チップが見つかりません（${chips.length} 個）`);
  let x = ch.right - chips.length * CHIP_W - (chips.length - 1) * ch.gap;
  chips.forEach(o => {
    const cur = o.options.fill && String(o.options.fill.color).toUpperCase() === C.teal;
    o.options.x = r3(x);
    o.options.w = CHIP_W;
    o.options.fontSize = CHIP_PT;
    if (cur) o.options.fill = { color: TEAL };
    o.text.forEach(t => { if (t.options) t.options.fontSize = CHIP_PT; });
    x += CHIP_W + ch.gap;
  });
  return s;
}

/**
 * 見出し＋本文の項目（各項目は1〜2行）を縦に並べたカード。文字はすべて箱の内側。
 * center: true なら本文の項目群を見出しの下の残りの高さの中で上下中央に置く。
 */
function itemCard(s, o) {
  const pad = 0.15, itemGap = o.itemGap != null ? o.itemGap : 0.2;
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.bg, radius: 0.08 });
  const iw = o.w - 2 * pad;
  let cy = o.y + pad;
  s.text(o.title, { x: o.x + pad, y: cy, w: iw, h: heightFor(1, 24, 0), size: 24, bold: true, color: o.titleColor });
  cy += heightFor(1, 24, 0) + 0.14;
  const hs = o.items.map(lines => heightFor(lines.length, 22, 0));
  const need = hs.reduce((a, b) => a + b, 0) + itemGap * (hs.length - 1);
  const room = o.y + o.h - pad - cy;
  if (need > room + 1e-6) throw new Error(`itemCard「${o.title}」: 高さ不足（必要 ${(o.h + need - room).toFixed(2)}in）`);
  if (o.center) cy += (room - need) / 2;
  o.items.forEach((lines, i) => {
    s.text(lines, { x: o.x + pad, y: cy, w: iw, h: hs[i], size: 22 });
    cy += hs[i] + itemGap;
  });
}

/** 右向きの太い矢印（ネイティブ freeform の塗り図形）。x0→x1、中心の高さ yc。 */
function blockArrow(s, x0, x1, yc, color = '595959') {
  const shaft = 0.09, head = 0.25, hl = Math.min(0.22, (x1 - x0) * 0.55);
  s.polyline([
    [x0, yc - shaft], [x1 - hl, yc - shaft], [x1 - hl, yc - head], [x1, yc],
    [x1 - hl, yc + head], [x1 - hl, yc + shaft], [x0, yc + shaft],
  ], { closed: true, fill: color, color, width: 0.75 });
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['原理', '波形', '数値', '使い方', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '2-5a カプノグラフィ①：原理と正常波形',
  });
  const newSlide = o => deck.slide(o);   // 章チップはライブラリ既定（シリーズ共通）

  // ===== S1 表紙 =====
  deck.cover({
    title: 'カプノグラフィ① 原理と正常波形',
    subtitle: '2-5 モニタリング',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 なぜ CO2 を測るのか（産生→運搬→排出） =====
  {
    const s = newSlide({ title: 'なぜ CO2 を測るのか', chapter: 0, notes: notesOf('S2'), sources: sourcesOf('S2') });
    // カード間 0.9in に太い矢印（長さ 0.42in、左右に 0.24in）。カードは内側余白 0.12 で本文1行が収まる幅。
    const gapA = 0.9, am = 0.24, cw = (BW - 2 * gapA) / 3, cy = 2.6, ch = 2.15;
    const up = em('↑', C.vermilion), dn = em('↓', C.blue);
    const cards = [
      { title: '産生（代謝）', lines: [[up, ' 発熱・悪性高熱症'], [dn, ' 低体温']] },
      { title: '運搬（循環）', lines: [[dn, ' 心拍出量の低下'], [dn, ' 肺塞栓']] },
      { title: '排出（換気）', lines: [[up, ' 低換気'], [dn, ' 過換気']] },
    ];
    cards.forEach((c, i) => {
      const x = X0 + i * (cw + gapA);
      s.card({ x, y: cy, w: cw, h: ch, pad: 0.12, title: c.title, lines: c.lines, titleColor: C.dark, gap: 14 });
      if (i < 2) blockArrow(s, x + cw + am, x + cw + gapA - am, cy + ch / 2);
    });
    s.text('EtCO2 は代謝・循環・換気の3つを同時に映す', { x: X0, y: cy + ch + 0.7, w: BW, size: 28, bold: true, color: TEAL, align: 'center' });
  }

  // ===== S3 測定の原理 =====
  {
    const s = newSlide({ title: '測定の原理', chapter: 0, notes: notesOf('S3'), sources: sourcesOf('S3') });
    // 右: 縦長カード2枚（高さは本文3行に合わせ、本文は見出しの下で上下中央）。全体を本文ゾーンの上下中央付近に置く。
    const cw = 3.45, gap = 0.28, x1 = XR - 2 * cw - gap, cy = 2.8, ch = 2.7;
    // 左: 箇条書き（1項目目は語の切れ目で2行）。カードの上下中央にそろえる。
    const bx = X0, tx = X0 + 0.5, tw = 5.3 - tx, size = 28, igap = 0.45;
    const items = [['CO2 は特定の波長の', '赤外線を吸収する'], ['吸収の量 → CO2 分圧']];
    const lh = items.map(lines => heightFor(lines.length, size, 0));
    let y = cy + ch / 2 - (lh.reduce((a, b) => a + b, 0) + igap) / 2;
    items.forEach((lines, i) => {
      s.text('•', { x: bx, y, w: 0.45, size });
      s.text(lines, { x: tx, y, w: tw, size });
      y += lh[i] + igap;
    });
    itemCard(s, { x: x1, y: cy, w: cw, h: ch, title: 'メインストリーム', titleColor: C.blue, center: true, itemGap: 0.14,
      items: [['回路の途中で測る'], ['遅れがない'], ['センサーが重い']] });
    itemCard(s, { x: x1 + cw + gap, y: cy, w: cw, h: ch, title: 'サイドストリーム', titleColor: C.vermilion, center: true, itemGap: 0.14,
      items: [['ガスを吸引して測る'], ['数秒の遅れ'], ['水で詰まる']] });
  }

  // ===== S4 正常波形の4つの相 =====
  {
    const s = newSlide({ title: '正常波形の4つの相', chapter: 1, notes: notesOf('S4'),
      sources: sourcesOf('S4', 'Anesthesiology 2013') });

    // --- 各相の意味（右: 小カード4枚。高さをそろえ、本文はすべて1行） ---
    // 相 III の本文が1行に収まる幅（本文枠 4.6in）を取り、残りを波形に使う。
    const cpad = 0.13, ctw = 1.15, cgap = 0.1, cbw = 4.6, cpr = 0.12;
    const rw = cpad + ctw + cgap + cbw + cpr, rx = XR - rw;
    const rh = 0.8, rgap = 0.3, ry0 = 1.95;
    const rows = [
      { t: '相 I', b: '死腔のガス（CO2 = 0）' },
      { t: '相 II', b: '死腔と肺胞ガスの混合' },
      { t: '相 III', b: '肺胞プラトー、終点が EtCO2' },
      { t: '相 0', b: '吸気、基線は 0 に戻る' },
    ];
    rows.forEach((r, i) => {
      const y = ry0 + i * (rh + rgap);
      s.box({ x: rx, y, w: rw, h: rh, fill: C.bg, radius: 0.08 });
      s.text(r.t, { x: rx + cpad, y: y + 0.1, w: ctw, h: rh - 0.2, size: 24, bold: true, color: C.blue, valign: 'middle' });
      s.text(r.b, { x: rx + cpad + ctw + cgap, y: y + 0.1, w: cbw, h: rh - 0.2, size: 22, valign: 'middle' });
    });

    // --- 時間カプノグラム（左） ---
    // ラベルはすべて波形の外: 上段（y=tierB）に α角・相 III・β角・EtCO2、基線の下に 相 II・相 0・相 I・時間。
    const figR = rx - 0.24;                     // 図の右端（カードとの間隔 0.24in）
    const xa = 0.65, yb = 5.5, xEnd = figR - 0.02; // 軸の原点と横軸の右端
    const top0 = 3.5, top1 = 3.35;              // プラトーの始点・終点の高さ（わずかに右上がり）
    const tierB = 2.6, labH = heightFor(1, 24, 0), below = yb + 0.15;
    const wA = 0.76, wIII = 1.15, wII = 0.95, wI = 0.75;
    const RISE = 0.32, PL = 1.25, FALL = 0.17, TAIL = 0.04;
    // 1呼吸目: α角ラベルが y 軸から 0.24in 離れる位置に立ち上がりの角を置く
    const p0 = xa + 0.24 + wA + 0.2, u1 = p0 - RISE, p1 = p0 + PL, d1 = p1 + FALL;
    // 相 0 は下降脚の真下。相 I（2呼吸目の立ち上がりの直前）との間を 0.24in あける
    const l0 = [(p1 + d1) / 2 - wI / 2, (p1 + d1) / 2 + wI / 2];
    const u2 = l0[1] + 0.24 + wI + 0.02;
    const p0b = u2 + RISE, p1b = p0b + PL, d2 = p1b + FALL;
    if (d2 + TAIL + 0.06 > xEnd) throw new Error(`S4: 波形が図の幅に収まりません（${r3(d2)} > ${r3(xEnd)}）`);
    const breath = u => {
      const p = u + RISE, q = p + PL;
      return [
        [u, yb], [u + 0.06, yb - 0.12], [u + 0.12, yb - 0.75], [u + 0.18, yb - 1.55],
        [u + 0.23, top0 + 0.25], [u + 0.28, top0 + 0.06], [p, top0],     // 相 II（急峻な立ち上がり）
        [q, top1],                                                       // 相 III（肺胞プラトー）
        [q + 0.07, top1 + 0.5], [q + 0.13, yb - 0.3], [q + FALL, yb],     // 相 0（急な下降）
      ];
    };
    // 軸
    s.arrow({ x1: xa, y1: yb, x2: xa, y2: 2.55, color: C.gray, width: 2 });
    s.arrow({ x1: xa, y1: yb, x2: xEnd, y2: yb, color: C.gray, width: 2 });
    s.text('CO2 分圧', { x: X0, y: 1.85, w: 1.45, size: 22, color: C.dark });
    s.text('時間', { x: xEnd - 0.75, y: below, w: 0.75, size: 22, color: C.dark, align: 'right' });
    // 波形（2呼吸）
    s.polyline([[xa + 0.05, yb], ...breath(u1), ...breath(u2), [d2 + TAIL, yb]], { color: C.blue, width: 3.5 });

    // 相の番号ラベル（波形の外: 上か下）
    const lab = (t, x, y, w) => s.text(t, { x, y, w, h: labH, size: 24, bold: true, color: C.dark, align: 'center' });
    lab('相 II', u1 + RISE / 2 - wII / 2, below, wII);      // 立ち上がりの真下（基線の下）
    lab('相 III', (p0 + p1) / 2 - wIII / 2, tierB, wIII);    // プラトーの上
    lab('相 0', l0[0], below, wI);                           // 下降脚の下（基線の下）
    lab('相 I', u2 - 0.02 - wI, below, wI);                  // 次の立ち上がり直前の平坦部の下

    // α角・β角（1呼吸目の角。弧は角の内側、ラベルは波形の上に出して引き出し線でつなぐ）
    const arc = (cx, cy, r, a0, a1) => {
      const pts = [];
      for (let k = 0; k <= 12; k++) {
        const t = (a0 + (a1 - a0) * k / 12) * Math.PI / 180;
        pts.push([cx + r * Math.cos(t), cy + r * Math.sin(t)]);
      }
      return pts;
    };
    s.polyline(arc(p0, top0, 0.26, -12, 112), { color: C.vermilion, width: 2 });
    s.polyline(arc(p1, top1, 0.26, 84, 167), { color: C.vermilion, width: 2 });
    const xA = p0 - 0.2 - wA, xB = p1 + 0.2;
    s.text('α角', { x: xA, y: tierB, w: wA, h: labH, size: 24, bold: true, color: C.vermilion, align: 'center' });
    s.polyline([[xA + wA - 0.06, tierB + labH + 0.03], [p0 - 0.04, top0 - 0.06]], { color: C.vermilion, width: 1.5 });
    s.text('β角', { x: xB, y: tierB, w: wA, h: labH, size: 24, bold: true, color: C.vermilion, align: 'center' });
    s.polyline([[xB + 0.06, tierB + labH + 0.03], [p1 + 0.04, top1 - 0.06]], { color: C.vermilion, width: 1.5 });

    // EtCO2（2呼吸目のプラトー終点）
    const dE = 0.17, wE = 1.1;
    s.circle({ cx: p1b, cy: top1, d: dE, fill: TEAL });
    const xE = Math.min(p1b - wE / 2, figR - wE);
    s.text('EtCO2', { x: xE, y: tierB, w: wE, h: labH, size: 24, bold: true, color: TEAL, align: 'center' });
    s.polyline([[p1b, tierB + labH + 0.03], [p1b, top1 - dE / 2 - 0.04]], { color: TEAL, width: 1.5 });

    if (process.env.DEBUG_S4) {
      console.log('S4 geometry', { rx: r3(rx), figR: r3(figR), xa, u1: r3(u1), p0: r3(p0), p1: r3(p1), d1: r3(d1), u2: r3(u2), p1b: r3(p1b), d2: r3(d2),
        alpha: [r3(xA), r3(xA + wA)], III: [r3((p0 + p1) / 2 - wIII / 2), r3((p0 + p1) / 2 + wIII / 2)], beta: [r3(xB), r3(xB + wA)], EtCO2: [r3(xE), r3(xE + wE)],
        II: [r3(u1 + RISE / 2 - wII / 2), r3(u1 + RISE / 2 + wII / 2)], ph0: l0.map(r3), phI: [r3(u2 - 0.02 - wI), r3(u2 - 0.02)], time: [r3(xEnd - 0.75), r3(xEnd)] });
    }
  }

  // ===== S5 EtCO2 の正常値 =====
  {
    const s = newSlide({ title: 'EtCO2 の正常値', chapter: 2, notes: notesOf('S5'), sources: sourcesOf('S5') });
    // 左カードと右の箇条書きを本文ゾーン（下線〜出典行）の上下中央付近にそろえる
    const bx = X0, by = 2.55, bw = 4.3, bh = 3.0, tm = 0.06;
    s.box({ x: bx, y: by, w: bw, h: bh, fill: C.bg, radius: 0.1 });
    const hNum = heightFor(1, 44, 0), hCap = heightFor(1, 22, 0), gNum = 0.25;
    const ny = by + (bh - hNum - gNum - hCap) / 2;
    s.text('35–45 mmHg', { x: bx + tm, y: ny, w: bw - 2 * tm, h: hNum, size: 44, bold: true, color: TEAL, align: 'center' });
    s.text('EtCO2 の一般的な正常範囲', { x: bx + tm, y: ny + hNum + gNum, w: bw - 2 * tm, h: hCap, size: 22, align: 'center' });
    // 右: 3行の箇条書き（spec どおり同じ階層）。カードとの間 0.24in、右端まで使って幅に余裕を持たせる。
    const lx = bx + bw + 0.24, lgap = 20, lh = heightFor(3, 24, lgap);
    s.bullets([
      ['PaCO2 より ', em('2〜5 mmHg', C.vermilion), ' 低い（健康な肺）'],
      '差が広がる ＝ 死腔が増えた',
      '肺塞栓・心拍出量の低下・COPD',
    ], { x: lx, y: by + bh / 2 - lh / 2, w: XR - lx, h: lh, size: 24, gap: lgap });
  }

  // ===== S6 変化の三つの理由 =====
  {
    const s = newSlide({ title: '変化の三つの理由', chapter: 2, notes: notesOf('S6'), sources: sourcesOf('S6') });
    const c1 = 1.6, cr = (BW - c1) / 3;
    s.table([
      ['', '産生（代謝）', '運搬（循環）', '排出（換気）'],
      [{ text: '上がる', bold: true, color: C.vermilion, align: 'center' },
        '発熱・シバリング\n悪性高熱症', '駆血帯の解除\nCO2 気腹', '低換気\n吸収剤の劣化'],
      [{ text: '下がる', bold: true, color: C.blue, align: 'center' },
        '低体温', '低血圧・出血\n肺塞栓・心停止', '過換気'],
    ], { x: X0, y: 2.0, colW: [c1, cr, cr, cr], size: 24, align: 'center', rowH: [0.75, 1.35, 1.35] });
    s.text([['波形が消える（回路外れ・閉塞・食道挿管）は別に扱う ', em('→ 2-5b', TEAL)]],
      { x: X0, y: 5.85, w: BW, size: 24 });
  }

  // ===== S7 チューブ位置の確認 =====
  {
    const s = newSlide({ title: 'チューブ位置の確認', chapter: 3, notes: notesOf('S7'),
      sources: sourcesOf('S7', 'from PUMA') });
    s.box({ x: X0, y: 1.9, w: BW, h: 0.8, fill: C.bg, line: { color: C.red, width: 2.25 }, radius: 0.08,
      text: '持続する呼気 CO2 がなければ、チューブを抜くのが原則', size: 28, bold: true, color: C.red });
    // 箇条書きは1行ずつ一定の行送りで置き、「→ 7-5 で詳しく」を5行目と同じ行の右端にそろえる。
    const items = [
      ['呼気で上がる波形が ', em('7 呼吸以上', C.vermilion), '、小さくならずに続く'],
      ['ピークが基線より ', em('7.5 mmHg（1 kPa）', C.vermilion), 'を超える'],
      'SpO2 は遅れて下がる。待たない',
      '胃に入った CO2 は低く、数呼吸で消える',
      '平坦なら、回路外れ・閉塞も同時に確認する',
    ];
    const by0 = 3.0, pitch = 0.62, bw = 9.45;
    const gw = 2.8, gx = XR - gw;                      // 案内の枠（右端そろえ）
    items.forEach((it, i) => {
      const w = i === items.length - 1 ? gx - 0.24 - X0 : bw;
      s.bullets([it], { x: X0, y: by0 + i * pitch, w, size: 24, gap: 0 });
    });
    s.text('→ 7-5 で詳しく', { x: gx, y: by0 + (items.length - 1) * pitch, w: gw, size: 24, bold: true, color: TEAL, align: 'right' });
  }

  // ===== S8 手術室の外での使い方 =====
  {
    const s = newSlide({ title: '手術室の外での使い方', chapter: 3, notes: notesOf('S8'),
      sources: sourcesOf('S8', '2025 (doi:') });
    // 3枚とも同じ高さ（カード1の本文量に合わせる）。カード2・3は本文を見出しの下で上下中央に置く。
    const gap = 0.3, cw = (BW - 2 * gap) / 3, cy = 2.15, ch = 3.7;
    const cards = [
      { title: '心肺蘇生', titleColor: C.vermilion, items: [
        ['胸骨圧迫の質を映す'],
        ['低い・下がる', [em('→ 圧迫を見直す', C.vermilion)]],
        ['急に上がって続く', [em('→ 自己心拍再開', TEAL)]],
      ] },
      { title: '鎮静・搬送', titleColor: C.blue, center: true, itemGap: 0.4, items: [['無呼吸をすぐ検出'], ['SpO2 より早い']] },
      { title: '挿管患者', titleColor: TEAL, center: true, itemGap: 0.4, items: [['搬送中もつける'], ['事故抜管の検出']] },
    ];
    cards.forEach((c, i) => itemCard(s, { x: X0 + i * (cw + gap), y: cy, w: cw, h: ch, ...c }));
  }

  // ===== S9 まとめ =====
  {
    const s = newSlide({ title: 'まとめ', chapter: 4, notes: notesOf('S9'),
      sources: sourcesOf('S9', '2022;77:1395-1415 / 日本') });
    s.bullets([
      'CO2 は代謝→循環→換気の流れを映す',
      '4つの相：死腔 → 混合 → 肺胞プラトー → 吸気',
      ['正常は ', em('35–45 mmHg', TEAL), '、PaCO2 より 2〜5 mmHg 低い'],
      ['持続する呼気 CO2 がなければ、', em('気管内と見なさない', C.red)],
      '蘇生でも鎮静でも搬送でも使う',
    ], { y: 2.0, size: 28, gap: 16 });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
