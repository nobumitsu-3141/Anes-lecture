'use strict';
/**
 * 2-7 観血的動脈圧
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言はここに転記（改行位置のみ調整）、ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 波形はすべて s.polyline（ネイティブ図形）。ラベルは図の外に置き、引き出し線でつなぐ。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, fits, measure, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '2-7_arterial_line';
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
/** 1 行に収まる出典は ' / ' で項目に分けて渡す（ライブラリが 1〜2 行に詰める） */
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  return s ? s.split(' / ') : undefined;
}
/**
 * 1 項目が 16pt・1 行に収まらない出典は、語の切れ目で 2 行に割って渡す。
 * lines を半角スペース 1 つで連結すると ref（既定は spec の sources 行）と一致することを確認する。
 */
function sourcesSplit(id, lines, ref) {
  const S = LAYOUT.source;
  const want = ref || SPEC[id].sources;
  if (lines.join(' ') !== want) throw new Error(`${id} 出典の分割が spec の文言と一致しません:\n  ${lines.join(' ')}\n  ${want}`);
  lines.forEach(l => { if (!fits(l, S.size, S.w)) throw new Error(`${id} 出典「${l}」が 1 行に収まりません`); });
  if (fits(lines.join(' / '), S.size, S.w)) throw new Error(`${id} 出典: 1 行に収まるので分割不要`);
  return lines;
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const em = (text, color, bold = true) => ({ text, color, bold });
const WAVE_W = 3.5;              // 波形の線幅(pt)
const LEAD_C = C.gray;           // 引き出し線
const LEAD_W = 1.5;
// 文字用のティール。C.teal(00A8AA) は白地 2.9:1・F2F2F2 地 2.6:1 で投影に弱いので、文字だけ濃くする
// （白地 5.2:1、F2F2F2 地 4.6:1。図の線・塗りは C.teal のまま）
const TEAL_T = '007A7D';
const ARROW_C = '404040';

/**
 * 見出し付きカード（手組み）。項目ごとに間隔をあけ、1 項目を 2 行に割ったときは行間を詰める。
 * items: 文字列 | {lines:[行...], color, bold}（行は文字列 or run 配列）
 */
function cardBody(o) {
  const ts = o.titleSize || 26, bs = o.size || 22;
  const pad = o.pad != null ? o.pad : 0.18;
  const tgap = o.titleGap != null ? o.titleGap : 0.12;
  const igap = o.itemGap != null ? o.itemGap : 0.16;
  const items = o.items.map(it => (typeof it === 'string' ? { lines: [it] } : it));
  const hs = items.map(it => heightFor(it.lines.length, bs, 0));
  const need = 2 * pad + heightFor(1, ts, 0) + tgap + hs.reduce((a, b) => a + b, 0) + igap * (items.length - 1);
  return { ts, bs, pad, tgap, igap, items, hs, need };
}
function listCardH(o) { return Math.ceil(cardBody(o).need * 100) / 100; }
function listCard(s, o) {
  const b = cardBody(o);
  if (b.need > o.h + 1e-6) throw new Error(`カード「${o.title}」: 高さ不足（必要 ${b.need.toFixed(2)}in > h ${o.h}in）`);
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill || C.bg, radius: 0.08, line: o.line });
  const iw = o.w - 2 * b.pad;
  let y = o.y + b.pad;
  s.text(o.title, { x: o.x + b.pad, y, w: iw, size: b.ts, bold: true, color: o.titleColor || C.dark });
  y += heightFor(1, b.ts, 0) + b.tgap;
  b.items.forEach((it, i) => {
    s.text(it.lines, { x: o.x + b.pad, y, w: iw, h: b.hs[i], size: b.bs, color: it.color || C.dark, bold: !!it.bold });
    y += b.hs[i] + b.igap;
  });
}

/**
 * 列の箇条書き（手組み）。items: 各項目 = 行の配列（行は文字列 or run 配列）。
 * 1 項目が長いときだけ語句の切れ目で 2 行にする（文言は変えない）。
 * y を渡すと上端、yc を渡すと縦中央に合わせる。戻り値 {top, bottom}
 */
function bulletCol(s, items, o) {
  const size = o.size || 24, ind = o.indent || 0.42, gap = o.gap != null ? o.gap : 0.3;
  const hs = items.map(it => heightFor(it.length, size, 0));
  const total = hs.reduce((a, b) => a + b, 0) + gap * (items.length - 1);
  const top = o.yc != null ? o.yc - total / 2 : o.y;
  let y = top;
  items.forEach((lines, i) => {
    s.text('•', { x: o.x, y, w: ind, size, color: C.dark });
    s.text(lines, { x: o.x + ind, y, w: o.w - ind, size, color: C.dark });
    y += hs[i] + gap;
  });
  return { top, bottom: top + total };
}

/** 引き出し線（折れ線。lint の重なり判定の対象外なので、文字や波形に掛けないよう座標で管理する） */
function leader(s, pts) { s.polyline(pts, { color: LEAD_C, width: LEAD_W }); }

/**
 * 水平・垂直の矢印（軸＋塗りの三角の矢じり）。pptxgenjs の矢じりは線幅に比例して小さいので、
 * 矢じりを自前の三角形（freeform）で描いて大きさを決める。hl: 矢じりの長さ、hw: 矢じりの幅(in)
 */
function flowArrow(s, x1, y1, x2, y2, o = {}) {
  const color = o.color || ARROW_C, width = o.width || 3, hl = o.hl || 0.22, hw = o.hw || 0.26;
  const L = Math.hypot(x2 - x1, y2 - y1), ux = (x2 - x1) / L, uy = (y2 - y1) / L;
  if (Math.abs(ux) > 1e-9 && Math.abs(uy) > 1e-9) throw new Error('flowArrow は水平・垂直のみ');
  const bx = x2 - ux * hl, by = y2 - uy * hl;
  if (o.from) s.polyline(o.from.concat([[bx + ux * 0.02, by + uy * 0.02]]), { color, width });
  else if (uy === 0) s.hline({ x: Math.min(x1, bx), y: y1, w: Math.abs(bx - x1) + 0.02, color, width });
  else s.vline({ x: x1, y: Math.min(y1, by), h: Math.abs(by - y1) + 0.02, color, width });
  const px = -uy * hw / 2, py = ux * hw / 2;
  s.polyline([[x2, y2], [bx + px, by + py], [bx - px, by - py]], { color, width: 1, closed: true, fill: color });
}

// ---------------------------------------------------------------------------
// 動脈圧波形のモデル（1 拍、t∈[0,1] → 拡張期圧を 0・収縮期圧を 1 とした高さ）
//   tp: ピークの時刻、tn: 重複切痕の時刻、hn: 切痕の高さ、d/wd: 切痕後の小さな山、
//   k: 拡張期の下降の速さ、e: 収縮期の下降の形（>1 で丸いピーク・鋭い切痕、<1 で鋭いピーク・不明瞭な切痕）
// ---------------------------------------------------------------------------
const PROF = {
  std: { tp: 0.13, tn: 0.37, hn: 0.52, d: 0.07, wd: 0.05, k: 2.4, e: 1.4 },
  aorta: { tp: 0.17, tn: 0.36, hn: 0.66, d: 0.06, wd: 0.04, k: 2.0, e: 1.9 },
  radial: { tp: 0.10, tn: 0.47, hn: 0.36, d: 0.025, wd: 0.07, k: 2.6, e: 0.62 },
};
function pulse(t, p) {
  const { tp, tn, hn, d, wd, k, e } = p;
  if (t <= tp) return 0.5 * (1 - Math.cos(Math.PI * t / tp));
  if (t <= tn) { const u = (t - tp) / (tn - tp); return 1 - (1 - hn) * Math.pow(u, e); }
  const t3 = tn + wd;
  if (t <= t3) return hn + d * Math.sin((Math.PI / 2) * (t - tn) / wd);
  const s = t - t3, S = 1 - t3, h0 = hn + d;
  return h0 * (Math.exp(-k * s) - Math.exp(-k * S)) / (1 - Math.exp(-k * S));
}
/** 1 拍の点列。x: 立ち上がりの足、T: 幅、y0/y1: 拍の始め/終わりの拡張期レベル、P: 脈圧の高さ(in) */
function beatPts(x, T, y0, y1, P, prof, n = 90) {
  const pts = [];
  for (let i = 0; i <= n; i++) {
    const t = i / n;
    pts.push([x + t * T, y0 + (y1 - y0) * t - P * pulse(t, prof)]);
  }
  return pts;
}
/** 拍の中の時刻 t の座標 */
function beatAt(x, T, y0, y1, P, prof, t) { return [x + t * T, y0 + (y1 - y0) * t - P * pulse(t, prof)]; }
/** 立ち上がり上の高さ h（0〜1）の時刻 */
function upT(prof, h) { return prof.tp * Math.acos(1 - 2 * h) / Math.PI; }

// ---------------------------------------------------------------------------
// 急速フラッシュ試験の応答（2 次系: 固有振動数と減衰係数で波形が変わる）
//   入力 = 動脈圧波形 + 矩形波（フラッシュ中は入力どおりに張り付く）。出力の点列を返す。
// ---------------------------------------------------------------------------
function flushTrace(fx, yDia, scale, kind) {
  const zeta = { ok: 0.3, under: 0.12, over: 2.6 }[kind];
  const lam = 0.24, w = 2 * Math.PI / lam;        // 固有振動の波長(in)
  const T = 0.62, P = 0.42, plate = 0.8;           // 拍の幅・脈圧・フラッシュの高さ（拡張期=0 の相対値）
  // 拍の足: フラッシュ前に 1 拍、離したあと少し間をあけて 2 拍（振動を平らな基線の上で見せる）
  const feet = [-3 * T, -2 * T, -T, 0.1, 2.2, 2.2 + T];
  const tRise = 0.82, tFall = 1.38, W = 3.66;
  const inp = u => {
    if (u >= tRise && u < tFall) return plate;
    for (const f of feet) if (u >= f && u < f + T) return P * pulse((u - f) / T, PROF.std);
    return 0;
  };
  const dx = 0.002;
  let y = 0, v = 0;
  for (let u = -3 * T; u < 0; u += dx) { const a = w * w * (inp(u) - y) - 2 * zeta * w * v; v += a * dx; y += v * dx; }
  const pts = [];
  let k = 0;
  for (let u = 0; u <= W + 1e-9; u += dx, k++) {
    if (u >= tRise && u < tFall) { y = plate; v = 0; }
    else { const a = w * w * (inp(u) - y) - 2 * zeta * w * v; v += a * dx; y += v * dx; }
    if (k % 3 === 0 || Math.abs(u - tRise) < dx || Math.abs(u - tFall) < dx) pts.push([fx + u, yDia - scale * y]);
  }
  return pts;
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['適応', '準備', '波形', '減衰', '注意'],
    author: 'Anes-lecture',
    deckTitle: '2-7 観血的動脈圧',
  });

  // ===== S1 表紙 =====
  deck.cover({
    title: '観血的動脈圧',
    subtitle: '2-7 モニタリング',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 いつ入れるか（3 列カード） =====
  {
    // 出典: Scheer 2002 の 1 件が論文題名込みで 16pt・2 行に収まらないため、共著者を "et al." に略す。
    // 全文はノートの末尾に残す（deviations に記載）。
    const full = SPEC.S2.sources;
    const short = full.replace('Scheer B, Perel A, Pfeiffer UJ.', 'Scheer B, et al.');
    const src = sourcesSplit('S2', [
      'Scheer B, et al. Clinical review: complications and risk factors of peripheral',
      'arterial catheters. Crit Care 2002;6:199-204 / Saugel B, et al. Crit Care 2020;24:172',
    ], short);
    const s = deck.slide({ title: 'いつ入れるか', chapter: 0, notes: notesOf('S2') + '\n出典（全文）: ' + full, sources: src });
    // 列幅: 22pt で 1 項目 1 行に収める（適応 10 字、部位・避ける 11 字）
    const pad = 0.12, gap = 0.245;
    const ws = [3.70, 4.02, 4.02];
    const cards = [
      { title: '適応', titleColor: C.blue, items: ['一拍ごとの血圧が要る', '頻回の血液ガス', '昇圧薬の持続投与', '変動が大きい手術'] },
      { title: '部位', titleColor: C.blue, items: ['橈骨動脈が第一', 'ほかに大腿・上腕・足背'] },
      // spec の 4 項目をそのまま 4 行に（行の位置をカード1とそろえる）
      { title: '避ける・注意', titleColor: C.vermilion, items: [
        '穿刺部の感染・循環障害', 'シャント側は避ける',
        'Allen テストで', { lines: [[em('虚血は予測できない', C.vermilion)]] }] },
    ];
    const common = { pad, titleGap: 0.2, itemGap: 0.36 };
    const h = Math.max(...cards.map(c => listCardH({ ...common, ...c })));
    let x = X0;
    cards.forEach((c, i) => {
      listCard(s, { ...common, ...c, x, y: 1.98, w: ws[i], h });
      x += ws[i] + gap;
    });
  }

  // ===== S3 システムの構成（流れ図 + 箇条書き） =====
  {
    const src = sourcesSplit('S3', [
      'Saugel B, et al. How to measure blood pressure using an arterial catheter:',
      'a systematic 5-step approach. Crit Care 2020;24:172',
    ]);
    const s = deck.slide({ title: 'システムの構成', chapter: 1, notes: notesOf('S3'), sources: src });
    const fy = 3.1, fh = 1.1;
    const boxes = [
      { text: 'カテーテル', w: 2.25 },
      { text: ['硬く短い', 'チューブ'], w: 2.25 },
      { text: 'トランスデューサ', w: 3.35 },
      { text: 'モニター', w: 2.25 },
    ];
    const g = (BW - boxes.reduce((a, b) => a + b.w, 0)) / (boxes.length - 1);
    let x = X0, tdx = 0;
    boxes.forEach((b, i) => {
      s.box({ x, y: fy, w: b.w, h: fh, fill: C.bg, line: { color: C.blue, width: 2.25 }, radius: 0.1,
        text: b.text, size: 24, bold: true, color: C.dark });
      if (i === 2) tdx = x + b.w / 2;
      if (i < boxes.length - 1) flowArrow(s, x + b.w + 0.08, fy + fh / 2, x + b.w + g - 0.08, fy + fh / 2);
      x += b.w + g;
    });
    // トランスデューサの上に加圧バッグ
    const bw = 4.4, bh = 0.62, by = 1.92;
    s.box({ x: tdx - bw / 2, y: by, w: bw, h: bh, fill: C.white, line: { color: C.vermilion, width: 2.25 }, radius: 0.1,
      text: '加圧バッグ 300 mmHg', size: 24, bold: true, color: C.vermilion });
    flowArrow(s, tdx, by + bh + 0.06, tdx, fy - 0.06);
    s.bullets([
      '加圧バッグ 300 mmHg、持続フラッシュ約 3 mL/時',
      'チューブは短く硬く、気泡を残さない',
      '三方活栓は最小限にする',
    ], { y: 4.5, size: 24, gap: 12 });
  }

  // ===== S4 ゼロ点と高さ合わせ（2 枚のカード + 右下に小さな図） =====
  {
    const s = deck.slide({ title: 'ゼロ点と高さ合わせ', chapter: 1, notes: notesOf('S4'), sources: sourcesOf('S4') });
    const cw = (BW - 0.29) / 2, y = 1.95;
    const common = { size: 24, titleGap: 0.12, itemGap: 0.14 };
    const L = { title: 'ゼロ点', titleColor: C.blue, items: ['三方活栓を大気に開放', 'モニターの「ゼロ」を押す', '大気圧と回路のずれを消す'] };
    const R = { title: '高さ', titleColor: C.vermilion, items: [
      '右房の高さに置く', '第 4 肋間・中腋窩線',
      { lines: [['10 cm ずれると約 ', em('7.4 mmHg', C.vermilion)]] },
      '体位を変えたら合わせ直す'] };
    const h = Math.max(listCardH({ ...common, ...L }), listCardH({ ...common, ...R }));
    const rx = X0 + cw + 0.29;
    listCard(s, { ...common, ...L, x: X0, y, w: cw, h });
    listCard(s, { ...common, ...R, x: rx, y, w: cw, h });
    // カードの下（全幅・中央寄せ）: 胸郭（横から）・右房・右房の高さの点線・トランスデューサ。
    // 文字は 22pt、図形と 0.24in 以上あける。「右房」は胸郭の輪郭の内側（線に掛けない）に置く。
    const fh = 0.9, fy = y + h + 0.28, ray = fy + fh / 2;           // 図の帯: 下端は出典(6.85)の 0.24in 以上上
    if (fy + fh > LAYOUT.source.y - 0.24) throw new Error('S4 図が出典に近すぎます');
    const lh = heightFor(1, 22, 0), ly = ray - lh / 2;
    const g = 0.26, labW = t => measure(t, 22) + 0.06;
    const chw = 3.4, dash = 2.3, tdw = 0.62, tdh = 0.38, dot = 0.3;
    const total = labW('胸郭') + g + chw + dash + tdw + g + labW('トランスデューサ');
    let x = X0 + (BW - total) / 2;
    s.text('胸郭', { x, y: ly, w: labW('胸郭'), size: 22, color: '595959', bold: true });
    x += labW('胸郭') + g;
    const chx = x;
    s.box({ x: chx, y: fy, w: chw, h: fh, shape: 'roundRect', radius: 0.42, line: { color: C.gray, width: 2.25 } });
    const rax = chx + 0.55;                                         // 「右房」ラベル（輪郭の内側）
    s.text('右房', { x: rax, y: ly, w: labW('右房'), size: 22, color: C.vermilion, bold: true });
    const dcx = rax + labW('右房') + g + dot / 2;
    s.circle({ cx: dcx, cy: ray, d: dot, fill: C.vermilion });
    const tdx = chx + chw + dash;
    s.hline({ x: dcx + dot / 2 + 0.06, y: ray, w: tdx - 0.08 - (dcx + dot / 2 + 0.06), color: C.vermilion, width: 2.25, dash: 'dash' });
    s.box({ x: tdx, y: ray - tdh / 2, w: tdw, h: tdh, fill: C.teal, radius: 0.05 });
    s.text('トランスデューサ', { x: tdx + tdw + g, y: ly, w: labW('トランスデューサ'), size: 22, color: TEAL_T, bold: true });
  }

  // ===== S5 波形の読み方（左 2/3 に 2 拍、右 1/3 に箇条書き） =====
  {
    const s = deck.slide({ title: '波形の読み方', chapter: 2, notes: notesOf('S5'), sources: sourcesOf('S5') });
    const p = PROF.std;
    const x0 = 0.8, f1 = 1.05, T = 3.3, f2 = f1 + T, yD = 5.2, P = 2.4;
    let pts = [[x0, yD]];
    pts = pts.concat(beatPts(f1, T, yD, yD, P, p), beatPts(f2, T, yD, yD, P, p).slice(1));
    s.polyline(pts, { color: C.blue, width: WAVE_W });
    const at = (f, t) => beatAt(f, T, yD, yD, P, p, t);
    const lab = (t, o) => s.text(t, { size: 22, color: C.dark, ...o });
    // 収縮期圧（ピーク）: 1 拍目のピークの上
    const pk = at(f1, p.tp);
    lab('収縮期圧（ピーク）', { x: 0.7, y: 1.9, w: 3.2 });
    leader(s, [[pk[0], 2.42], [pk[0], pk[1] - 0.1]]);
    // 重複切痕: 1 拍目の切痕の右上
    const nt = at(f1, p.tn);
    lab('重複切痕', { x: nt[0] + 0.1, y: 2.95, w: 1.45 });
    leader(s, [[nt[0] + 0.3, 3.45], [nt[0] + 0.02, nt[1] - 0.1]]);
    // 拡張期の下降: 2 拍目の下降の右上
    const dd = at(f2, 0.77);
    lab('拡張期の下降', { x: 5.85, y: 3.15, w: 2.2 });
    leader(s, [[dd[0], 3.68], [dd[0], dd[1] - 0.1]]);
    // 拡張期圧: 1 拍目の足（最低点）の下
    lab('拡張期圧', { x: f1 - 0.5, y: 5.62, w: 1.45 });
    leader(s, [[f1, yD + 0.1], [f1, 5.58]]);
    // 収縮期の立ち上がり: 2 拍目の立ち上がりの中点へ、右下（拍の下の空き）から斜めに当てる。
    // 立ち上がりの線とは平行にせず、終点は線の右 0.07in で止める（波形の線に重ねない）
    const us = at(f2, upT(p, 0.5));
    const lcx = f2 + 0.8;
    lab('収縮期の立ち上がり', { x: lcx - 1.6, y: 5.62, w: 3.2, align: 'center' });
    leader(s, [[lcx + 0.45, 5.56], [us[0] + 0.07, us[1] + 0.04]]);
    bulletCol(s, [
      ['平均血圧 ≒', '拡張期圧 ＋ 脈圧 ÷ 3'],
      ['立ち上がりの速さ →', '収縮力の手がかり'],
      ['重複切痕 ≒', '大動脈弁の閉鎖'],
    ], { x: 8.45, w: XR - 8.45, yc: 3.85, gap: 0.42 });
  }

  // ===== S6 中枢と末梢で違う（大動脈 / 橈骨動脈の 1 拍ずつ + 箇条書き 4 行） =====
  {
    const s = deck.slide({ title: '中枢と末梢で違う', chapter: 2, notes: notesOf('S6'), sources: sourcesOf('S6') });
    // 圧 → y（両図で同じ縦横スケール）。橈骨のピーク(141)が見出しの下端から 0.24in 以上離れるように
    const yp = mm => 4.08 - (mm - 72) * 0.021;
    const T = 3.6, lead = 0.25, tail = 0.3, fw = lead + T + tail;
    const half = BW / 2;
    const figs = [
      { label: '大動脈', prof: PROF.aorta, dia: 80, sys: 122, color: C.blue },
      { label: '橈骨動脈', prof: PROF.radial, dia: 73, sys: 141, color: C.vermilion },
    ];
    const aSys = yp(122), aDia = yp(80);
    const fxs = [];
    figs.forEach((f, i) => {
      const cx = X0 + half * i + half / 2, fx = cx - fw / 2;
      s.text(f.label, { x: cx - 0.95, y: 1.86, w: 1.9, size: 24, bold: true, color: f.color, align: 'center' });
      // 大動脈の収縮期圧・拡張期圧のレベル（比較用の点線）
      s.hline({ x: fx, y: aSys, w: fw, color: C.lightGray, width: 1.5, dash: 'dash' });
      s.hline({ x: fx, y: aDia, w: fw, color: C.lightGray, width: 1.5, dash: 'dash' });
      const yD = yp(f.dia), P = yD - yp(f.sys);
      const pts = [[fx, yD]].concat(beatPts(fx + lead, T, yD, yD, P, f.prof, 120));
      const ye = pts[pts.length - 1][1];
      pts.push([fx + fw, ye]);
      s.polyline(pts, { color: f.color, width: WAVE_W });
      fxs.push([fx, fx + fw]);
    });
    // 点線の注記: 2 つの図の間に「収縮期」「拡張期」（大動脈の収縮期圧・拡張期圧の高さ。線と同じ高さに置く）
    const gx0 = fxs[0][1], gx1 = fxs[1][0], lw = measure('収縮期', 22) + 0.06, lh = heightFor(1, 22, 0);
    if (gx1 - gx0 < lw + 2 * 0.24) throw new Error('S6 点線の注記が入りません');
    [['収縮期', aSys], ['拡張期', aDia]].forEach(([t, yy]) =>
      s.text(t, { x: (gx0 + gx1 - lw) / 2, y: yy - lh / 2, w: lw, size: 22, color: '595959', align: 'center' }));
    if (aDia + lh / 2 > 4.4 - 0.24) throw new Error('S6 注記と箇条書きが近すぎます');
    s.bullets([
      '末梢ほど収縮期圧は高く、拡張期圧は低く出る',
      ['平均血圧はほぼ同じ → ', em('管理は平均血圧で', TEAL_T)],
      '重複切痕は遅れて、不明瞭になる',
      '足背動脈は橈骨よりさらに収縮期圧が高い',
    ], { y: 4.4, size: 24, gap: 8 });
  }

  // ===== S7 減衰と共鳴（フラッシュ試験 3 図 + カード 2 枚 + 一行） =====
  {
    const src = sourcesSplit('S7', [
      'Saugel B, et al. Crit Care 2020;24:172 / Gardner RM. Direct blood pressure measurement:',
      'dynamic response requirements. Anesthesiology 1981;54:227-236',
    ]);
    const s = deck.slide({ title: '減衰と共鳴', chapter: 3, notes: notesOf('S7'), sources: src });
    const cw = (BW - 2 * 0.265) / 3;
    const scale = 0.9, yDia = 1.98 + scale * 0.8;   // フラッシュの上端を y 1.98 に
    const panels = [
      { kind: 'ok', label: '適正：振動 1–2 回', color: C.blue },
      { kind: 'under', label: '減衰不足：振動が多い', color: C.vermilion },
      { kind: 'over', label: '過減衰：振動なし', color: C.teal, text: TEAL_T },
    ];
    let lo = 0;
    panels.forEach((p, i) => {
      const x = X0 + i * (cw + 0.265);
      const pts = flushTrace(x + (cw - 3.66) / 2, yDia, scale, p.kind);
      lo = Math.max(lo, ...pts.map(q => q[1]));
      s.polyline(pts, { color: p.color, width: 3 });
      p.x = x;
    });
    const ly = lo + 0.2;
    panels.forEach(p => s.text(p.label, { x: p.x, y: ly, w: cw, size: 22, bold: true, color: p.text || p.color, align: 'center' }));
    // 下段: カード 2 枚 + 一行（平均血圧）
    const cy = ly + heightFor(1, 22, 0) + 0.3;
    const cardW = 4.45, g = 0.27, mw = BW - 2 * cardW - 2 * g;
    const common = { size: 22, titleGap: 0.1, itemGap: 0.12, pad: 0.15 };
    const A = { title: '減衰不足', titleColor: C.vermilion, items: ['収縮期↑ 拡張期↓', '長いチューブ・接続の多さ'] };
    const B = { title: '過減衰', titleColor: TEAL_T, items: ['収縮期↓ 拡張期↑', '気泡・血栓・屈曲・緩み'] };
    const h = Math.max(listCardH({ ...common, ...A }), listCardH({ ...common, ...B }));
    listCard(s, { ...common, ...A, x: X0, y: cy, w: cardW, h });
    listCard(s, { ...common, ...B, x: X0 + cardW + g, y: cy, w: cardW, h });
    s.box({ x: XR - mw, y: cy, w: mw, h, fill: C.white, line: { color: '404040', width: 2.25 }, radius: 0.08,
      text: ['平均血圧は', '比較的保たれる'], size: 24, bold: true, color: C.dark, margin: 0.06 });
  }

  // ===== S8 波形がおかしいとき（強調ボックス + 5 段のフロー、2 段に折り返し） =====
  {
    const s = deck.slide({ title: '波形がおかしいとき', chapter: 3, notes: notesOf('S8'), sources: sourcesOf('S8') });
    s.box({ x: X0, y: 1.95, w: BW, h: 0.9, fill: C.bg, line: { color: C.red, width: 2.5 }, radius: 0.08,
      text: '波形が消えた ＝ まず患者（脈・SpO2 の脈波）', size: 28, bold: true, color: C.red });
    // 1 段目 ①②③（全幅）、2 段目 ④⑤（中央寄せ、箱の間隔は 1 段目と同じ）。③ から ④ へ折り返す
    const ch = 0.88, r1 = 3.42, r2 = 5.3;
    const row1 = [{ w: 2.75, t: '① 患者を見る', c: C.red }, { w: 3.45, t: '② 三方活栓の向き', c: C.blue }, { w: 4.55, t: '③ 屈曲・気泡・バッグ圧', c: C.blue }];
    const row2 = [{ w: 3.45, t: '④ ゼロ点・高さ', c: C.blue }, { w: 3.45, t: '⑤ フラッシュ試験', c: C.blue }];
    const g = (BW - row1.reduce((a, b) => a + b.w, 0)) / (row1.length - 1);
    const place = (row, y, x0) => {
      let x = x0;
      row.forEach((b, i) => {
        b.x = x;
        s.box({ x, y, w: b.w, h: ch, fill: C.bg, line: { color: b.c, width: 2.25 }, radius: 0.12, text: b.t, size: 24, bold: true, color: C.dark });
        if (i < row.length - 1) flowArrow(s, x + b.w + 0.08, y + ch / 2, x + b.w + g - 0.08, y + ch / 2);
        x += b.w + g;
      });
    };
    place(row1, r1, X0);
    const w2 = row2.reduce((a, b) => a + b.w, 0) + g * (row2.length - 1);
    place(row2, r2, X0 + (BW - w2) / 2);
    // ③ → ④ の折り返し（下 → 左 → 下）
    const ym = (r1 + ch + r2) / 2, c3 = row1[2].x + row1[2].w / 2, c4 = row2[0].x + row2[0].w / 2;
    flowArrow(s, c4, ym, c4, r2 - 0.08, { from: [[c3, r1 + ch + 0.08], [c3, ym], [c4, ym]] });
  }

  // ===== S9 安全：やってはいけない（2 × 2 カード） =====
  {
    const s = deck.slide({ title: '安全：やってはいけない', chapter: 4, notes: notesOf('S9') });
    const cw = (BW - 0.3) / 2;
    const common = { size: 24, titleGap: 0.14, itemGap: 0.16 };
    const cards = [
      { title: '薬を入れない', titleColor: C.red, items: ['動脈ラインに注射しない', '赤いラベル、注入口を作らない'] },
      { title: '離断 ＝ 出血', titleColor: C.red, items: ['接続と固定を確認', 'アラームを切らない'] },
      { title: '虚血を見る', titleColor: C.vermilion, items: ['指の色・温度・動き', '不要になったら抜く'] },
      { title: '採血の手順', titleColor: C.blue, items: ['捨て血 → 採血 → フラッシュ', '活栓を元に戻す'] },
    ];
    const need = Math.max(...cards.map(c => listCardH({ ...common, ...c })));
    const vg = 0.3, h = Math.max(need, (6.7 - 1.95 - vg) / 2);
    cards.forEach((c, i) => {
      const x = X0 + (i % 2) * (cw + 0.3), y = 1.95 + Math.floor(i / 2) * (h + vg);
      listCard(s, { ...common, ...c, x, y, w: cw, h });
    });
  }

  // ===== S10 呼吸性変動への入口（連続波形 6 拍 + 箇条書き 2 行） =====
  {
    const s = deck.slide({ title: '呼吸性変動への入口', chapter: 4, notes: notesOf('S10') });
    const p = PROF.std;
    const xs = 0.85, x1 = 1.1, n = 6, T = (12.45 - x1) / n, Lr = 3 * T;   // 呼吸 1 回 = 3 拍
    const y0 = 4.4, P0 = 1.6;
    const ph = x => 2 * Math.PI * (x - x1 - p.tp * T) / Lr + Math.PI / 2;
    const D = x => y0 - 0.1 * Math.sin(ph(x));
    const Pp = x => P0 * (1 + 0.2 * Math.sin(ph(x)));
    let pts = [[xs, D(x1)]];
    for (let k = 0; k < n; k++) {
      const a = x1 + k * T;
      pts = pts.concat(beatPts(a, T, D(a), D(a + T), Pp(a + p.tp * T), p, 90).slice(k === 0 ? 0 : 1));
    }
    s.polyline(pts, { color: C.blue, width: WAVE_W });
    // 収縮期のピークを結ぶ包絡線（点線、ピークの 0.32in 上。波形の左端から右端まで）
    const env = [];
    for (let x = xs; x <= 12.45 + 1e-9; x += 0.05) env.push([x, D(x) - Pp(x) - 0.32]);
    s.polyline(env, { color: C.vermilion, width: 2.25, dash: 'dash' });
    s.bullets([
      ['陽圧換気で波形が大きく揺れる → ', em('輸液反応性のヒント', TEAL_T)],
      ['条件と判断の仕方は ', em('9-8', TEAL_T), ' で'],
    ], { y: 5.0, size: 28, gap: 16 });
  }

  // ===== S11 まとめ =====
  {
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: notesOf('S11'), sources: sourcesOf('S11') });
    // 行頭の語に色を付けると pptx の行頭記号がその色を引き継ぐため、記号は別の文字枠で黒(262626)に固定する
    bulletCol(s, [
      ['一拍ごとの血圧と波形が要るときに入れる'],
      [['ゼロ点は', em('大気開放', TEAL_T), '、高さは', em('右房', TEAL_T), 'に合わせる']],
      [['管理は', em('平均血圧', TEAL_T), 'で。末梢ほど収縮期は高く出る']],
      [[em('フラッシュ試験', TEAL_T), 'で減衰を確認する']],
      [[em('薬を入れない、離断させない、指を見る', C.red)]],
    ], { x: X0, w: BW, size: 32, indent: 0.55, gap: 0.3, yc: (LAYOUT.body.top + 0.25 + LAYOUT.source.y - 0.3) / 2 });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
