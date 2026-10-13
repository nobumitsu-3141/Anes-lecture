'use strict';
/**
 * 2-6 筋弛緩モニター
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言はここに転記（改行位置のみ調整）、ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, fits, heightFor, minCardH, LAYOUT } = require('../_lib/decklib');

const NAME = '2-6_neuromuscular_monitor';
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
/** 出典: spec の sources 行を ' / ' 区切りの項目にしてライブラリへ（詰めて最大2行） */
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  return s ? s.split(' / ') : undefined;
}
/**
 * 出典が2行（16pt・11.6in）に収まらないスライド用: 2行を明示する。
 * spec の sources 行に含まれない語が混じっていないか（短縮のみか）を確認する。
 */
function sourcesLines(id, l1, l2) {
  const s = SPEC[id] && SPEC[id].sources;
  if (!s) throw new Error(`${id}: spec に sources がありません`);
  const S = LAYOUT.source;
  for (const l of [l1, l2]) {
    if (!fits(l, S.size, S.w)) throw new Error(`${id} 出典: 1行に収まりません「${l}」`);
    for (const tok of l.split(/\s+|\s\/\s/).filter(t => t && t !== '/')) {
      if (!s.includes(tok)) throw new Error(`${id} 出典: spec にない語「${tok}」`);
    }
  }
  return [l1, l2];
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const em = (text, color, bold = true) => ({ text, color, bold });

/**
 * 文字・塗りに使う濃いティール。C.teal(00A8AA) は白地で約 2.9:1、F2F2F2 地で約 2.6:1 と
 * 投影で薄く見えるので、強調文字・白文字の地色にはこの色を使う（白地 約 4.8:1、F2F2F2 地 約 4.3:1）。
 * 2-5a と同じ値。
 */
const TEAL = C.teal;   // 書式ルールのティール（シリーズ共通）

/**
 * 現在章のチップの地色を濃いティールにする（白の太字 11pt が約 4.8:1 になる）。
 * decklib は C.teal で置くので、書き出し前の pptxgenjs のスライドオブジェクトを直接書き換える。
 * 位置・寸法・文字は変えない。
 */
function darkenActiveChip(s) {
  const ch = LAYOUT.chip;
  const chips = s.raw._slideObjects.filter(o => o._type === 'text' && o.options && o.options.shape === 'roundRect'
    && Math.abs(o.options.y - ch.y) < 1e-6 && Math.abs(o.options.h - ch.h) < 1e-6);
  if (chips.length !== 5) throw new Error(`p${s.page}: 章チップが見つかりません（${chips.length} 個）`);
  const cur = chips.filter(o => o.options.fill && String(o.options.fill.color).toUpperCase() === C.teal);
  if (cur.length !== 1) throw new Error(`p${s.page}: 現在章のチップが 1 個ではありません（${cur.length} 個）`);
  cur[0].options.fill = { color: TEAL };
  return s;
}

/** 見出し＋本文の項目（各項目は1〜2行）を縦に並べたカード。文字はすべて箱の内側。 */
function itemCard(s, o) {
  const pad = o.pad != null ? o.pad : 0.15;
  const itemGap = o.itemGap != null ? o.itemGap : 0.2;
  const size = o.size || 22, tsize = o.titleSize || 24;
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.bg, radius: 0.08 });
  const iw = o.w - 2 * pad;
  let cy = o.y + (o.top != null ? o.top : pad);
  const th = heightFor(1, tsize, 0);
  s.text(o.title, { x: o.x + pad, y: cy, w: iw, h: th, size: tsize, bold: true, color: o.titleColor });
  cy += th + (o.titleGap != null ? o.titleGap : 0.14);
  const hs = o.items.map(lines => heightFor(lines.length, size, 0));
  const need = hs.reduce((a, b) => a + b, 0) + itemGap * (hs.length - 1);
  const room = o.y + o.h - pad - cy;
  if (need > room + 1e-6) throw new Error(`itemCard「${o.title}」: 高さ不足（必要 ${(o.h + need - room).toFixed(2)}in）`);
  // center: 本文の項目群を見出しの下の残りの高さの中で上下中央に置く（項目の少ないカードが上に偏らない）
  if (o.center) cy += (room - need) / 2;
  o.items.forEach((lines, i) => {
    s.text(lines, { x: o.x + pad, y: cy, w: iw, h: hs[i], size });
    cy += hs[i] + itemGap;
  });
}

/** 塗りつぶしの三角形の矢じり（polyline）。tip=先端、dir=向き（単位ベクトルでなくてよい） */
function arrowHead(s, tip, dir, color, len = 0.2, half = 0.11) {
  const n = Math.hypot(dir[0], dir[1]);
  const ux = dir[0] / n, uy = dir[1] / n, px = -uy, py = ux;
  const bx = tip[0] - ux * len, by = tip[1] - uy * len;
  s.polyline([tip, [bx + px * half, by + py * half], [bx - px * half, by - py * half]],
    { color, width: 1, closed: true, fill: color });
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['必要性', '原理', '刺激', '判定', '実践'],
    author: 'Anes-lecture',
    deckTitle: '2-6 筋弛緩モニター',
  });
  const newSlide = o => deck.slide(o);   // 章チップはライブラリ既定（シリーズ共通）

  // ===== S1 表紙 =====
  deck.cover({
    title: '筋弛緩モニター',
    subtitle: '2-6 モニタリング',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 なぜ必ず付けるのか =====
  {
    // 出典は全体で 2 行（16pt）に入らないため、ASA ガイドラインの表題を略記（他スライドと同じ形）にする
    const s = newSlide({
      title: 'なぜ必ず付けるのか', chapter: 0, notes: notesOf('S2'),
      sources: sourcesLines('S2',
        '日本麻酔科学会「安全な麻酔のためのモニター指針」（2025年8月改訂） / Suzuki T, et al.',
        'J Anesth 2026;40:173-181 / Thilen SR, et al. Anesthesiology 2023;138:13-41'),
    });
    const gap = 0.5, cw = (BW - gap) / 2, xL = X0, xR = X0 + cw + gap;
    // 上: 数値の呼び出し（左右）。左の添え書きは「残存筋弛緩（TOF比 <0.9）」を割らないよう 3 行にする
    const yN = 1.75, hN = heightFor(1, 44, 0), yS = yN + hN + 0.08;
    s.text('約 30%', { x: xL, y: yN, w: cw, h: hN, size: 44, bold: true, color: C.vermilion, align: 'center', valign: 'middle' });
    const subL = ['回復室到着時に', '残存筋弛緩（TOF比 <0.9）が', 'ある割合の中央値'];
    s.text(subL, { x: xL, y: yS, w: cw, size: 22, align: 'center' });
    const subBot = yS + heightFor(subL.length, 22, 0);
    s.vline({ x: X0 + cw + gap / 2, y: yN + 0.1, h: subBot - 0.1 - (yN + 0.1), color: C.lightGray, width: 1.5 });
    s.text('見た目では分からない', { x: xR, y: yN, w: cw, h: hN, size: 36, bold: true, color: C.vermilion, align: 'center', valign: 'middle' });
    s.text(['頭を 5 秒持ち上げられても', 'TOF比 0.9 未満はある'], { x: xR, y: yS, w: cw, size: 22, align: 'center' });
    // 下: 横長カード2枚（強調の TOF比 0.9 以上 は濃いティール）
    const cy = subBot + 0.24, ch = minCardH(2, { gap: 6 });
    s.card({ x: xL, y: cy, w: cw, h: ch, title: 'JSA 2025', titleColor: C.blue, gap: 6,
      lines: ['深さを定量的に評価し、', ['抜管前に ', em('TOF比 0.9 以上', TEAL), 'を確認']] });
    s.card({ x: xR, y: cy, w: cw, h: ch, title: 'ASA 2023', titleColor: C.blue, gap: 6,
      lines: ['定量モニターを使い、', [em('TOF比 0.9 以上', TEAL), 'を確認して抜管']] });
  }

  // ===== S3 どこを刺激し何を測るか =====
  {
    const s = newSlide({ title: 'どこを刺激し何を測るか', chapter: 1, notes: notesOf('S3'), sources: sourcesOf('S3') });
    // --- 左: 前腕〜手の模式図（掌側、手は右向き、母指は上＝橈側、尺側は下） ---
    // 右の箇条書き（24pt・最長行）が x 5.45 から始まるので図の右端は 5.2 まで。縦は本文ゾーンいっぱいに使う。
    const line = { color: C.gray, width: 1.5 };
    const tipY = 2.8;                                                // 母指の先端
    const th = { w: 0.6, h: 0.9 };                                  // 母指
    const hd = { x: 3.6, y: tipY + th.h, w: 1.6, h: 1.85 };         // 手
    const fa = { x: 0.6, w: hd.x - 0.6, h: 1.5 };                   // 前腕（手の高さの中央）
    fa.y = hd.y + (hd.h - fa.h) / 2;
    th.x = hd.x + 0.25; th.y = tipY;
    s.box({ ...fa, fill: C.bg, line, radius: 0.3, shape: 'roundRect' });
    s.box({ ...hd, fill: C.bg, line, radius: 0.3, shape: 'roundRect' });
    s.box({ ...th, fill: C.bg, line, radius: 0.2, shape: 'roundRect' });
    // 電極（手首の尺骨側＝下縁寄り。遠位が陰極）。間隔はラベル幅（22pt 太字 5 字）に合わせる
    const labW = 1.75, d = 0.42;
    const ey = fa.y + fa.h - 0.36;
    const eNeg = fa.x + fa.w - 0.42, ePos = eNeg - labW;
    s.circle({ cx: ePos, cy: ey, d, fill: C.blue });
    s.circle({ cx: eNeg, cy: ey, d, fill: C.vermilion });
    // 電極のラベル（図の外・下）。引き出し線の真下に中央揃え
    const ly = fa.y + fa.h + 0.38, lTop = ey + d / 2 + 0.06;
    s.vline({ x: ePos, y: lTop, h: ly - 0.06 - lTop, color: C.blue, width: 1.5 });
    s.vline({ x: eNeg, y: lTop, h: ly - 0.06 - lTop, color: C.vermilion, width: 1.5 });
    s.label('＋（陽極）', { x: ePos - labW / 2, y: ly, w: labW, h: 0.5, bold: true, color: C.blue });
    s.label('−（陰極）', { x: eNeg - labW / 2, y: ly, w: labW, h: 0.5, bold: true, color: C.vermilion });
    // 母指の内転方向の矢印（弧）。始点は母指の先端のすぐ上
    const ac = [th.x + th.w / 2, hd.y], r = th.h + 0.07;
    const pt = deg => [ac[0] + r * Math.cos(deg * Math.PI / 180), ac[1] + r * Math.sin(deg * Math.PI / 180)];
    const arc = [];
    for (let a = -88; a <= -38; a += 2) arc.push(pt(a));
    s.polyline(arc, { color: TEAL, width: 3.5 });
    const tipA = -30, tip = pt(tipA), tr = tipA * Math.PI / 180;
    arrowHead(s, tip, [-Math.sin(tr), Math.cos(tr)], TEAL, 0.24, 0.13);
    // ラベル（図の外・上）と引き出し線
    const lab = { x: hd.x + hd.w - 2.85, y: 1.85, w: 2.85, h: 0.5 };
    s.label('母指内転筋の反応', { ...lab, bold: true, color: TEAL });
    const mid = pt(-60);
    s.vline({ x: mid[0], y: lab.y + lab.h + 0.06, h: mid[1] - 0.08 - (lab.y + lab.h + 0.06), color: TEAL, width: 1.5 });
    if (ly + 0.5 > 6.7) throw new Error(`S3: 電極ラベルが下にはみ出します（${(ly + 0.5).toFixed(2)}）`);
    // --- 右: 箇条書き ---
    s.bullets([
      '尺骨神経を最大上刺激する',
      '母指内転筋の動きを測る',
      ['電極は手首の尺骨側、遠位に', em('陰極', C.vermilion)],
      '加速度（AMG）か筋電図（EMG）で数値化',
    ], { x: 5.45, y: 2.6, w: XR - 5.45, size: 24, gap: 24 });
  }

  // ===== S4 刺激パターン =====
  {
    const s = newSlide({ title: '刺激パターン', chapter: 2, notes: notesOf('S4'), sources: sourcesOf('S4') });
    // カード幅: 間隔 0.30・内側余白 0.15。PTC の見出し「（テタヌス後カウント）」（22pt で 3.76in）が
    // 割らずに入る最小幅を中央のカードにとり、左右の 2 枚は同幅（左右対称）。
    // TOF の「10 秒以上あけて繰り返す」は語句の切れ目で 2 行にする。
    const gap = 0.3, pad = 0.15, wMid = 4.07, wSide = (BW - 2 * gap - wMid) / 2;
    const ws = [wSide, wMid, wSide];
    const cy = 1.85;
    const figTop = cy + 0.2, figBot = cy + 0.75;                // 刺激の縦線の上端・下端
    const cards = [
      {
        color: C.blue, abbr: 'TOF', jp: '（四連刺激）',
        items: [['2 Hz で 4 発'], ['カウント 0–4 と', '比 T4/T1'], ['10 秒以上あけて', '繰り返す']],
        // 等間隔の縦線 4 本
        pulses: w => { const sp = 0.72, st = (w - 3 * sp) / 2; return [0, 1, 2, 3].map(k => [st + k * sp, 3]); },
      },
      {
        color: C.vermilion, abbr: 'PTC', jp: '（テタヌス後カウント）',
        items: [['50 Hz 5 秒 →', '3 秒後に 1 Hz'], ['深い筋弛緩を数える']],
        // 密な縦線の束（テタヌス）→ 空白 → 等間隔の縦線 10 本
        pulses: w => {
          const tet = 12, tsp = 0.045, gapT = 0.55, n = 10, sp = 0.22;
          const span = (tet - 1) * tsp + gapT + (n - 1) * sp, st = (w - span) / 2;
          const out = [];
          for (let k = 0; k < tet; k++) out.push([st + k * tsp, 2]);
          for (let k = 0; k < n; k++) out.push([st + (tet - 1) * tsp + gapT + k * sp, 3]);
          return out;
        },
      },
      {
        color: TEAL, abbr: 'DBS', jp: '（二連バースト）',
        items: [['短い刺激を 2 回'], ['触覚で減衰を見やすい']],
        // 密な縦線 3 本 → 短い空白 → 密な縦線 3 本
        pulses: w => {
          const bsp = 0.08, gapB = 1.1, span = 4 * bsp + gapB, st = (w - span) / 2, out = [];
          for (let k = 0; k < 3; k++) out.push([st + k * bsp, 2.5]);
          for (let k = 0; k < 3; k++) out.push([st + 2 * bsp + gapB + k * bsp, 2.5]);
          return out;
        },
      },
    ];
    // 高さ: 最も行数の多い TOF の本文に合わせる（本文は上揃えで 3 枚の見出し・本文の始まりを揃える）
    const ty = figBot + 0.18, thh = 1.05, itemGap = 0.1;
    const bodyH = c => c.items.reduce((a, l) => a + heightFor(l.length, 22, 0), 0) + itemGap * (c.items.length - 1);
    const ch = Math.ceil((ty + thh + 0.14 + Math.max(...cards.map(bodyH)) + pad - cy) * 100) / 100;
    let x = X0;
    cards.forEach((c, i) => {
      const w = ws[i], iw = w - 2 * pad;
      s.box({ x, y: cy, w, h: ch, fill: C.bg, radius: 0.08 });
      // 刺激の模式図（ベースライン＋縦線）
      s.hline({ x: x + 0.25, y: figBot, w: w - 0.5, color: C.gray, width: 1 });
      c.pulses(w).forEach(([dx, pw]) => s.vline({ x: x + dx, y: figTop, h: figBot - figTop, color: c.color, width: pw }));
      // 見出し（略語＋日本語）
      s.text([[{ text: c.abbr, size: 28, bold: true, color: c.color }], [{ text: c.jp, size: 22, bold: true }]],
        { x: x + pad, y: ty, w: iw, h: thh, size: 22 });
      // 本文
      let by = ty + thh + 0.14;
      c.items.forEach((lines, k) => {
        const h = heightFor(lines.length, 22, 0);
        s.text(lines, { x: x + pad, y: by, w: iw, h, size: 22 });
        by += h + (k < c.items.length - 1 ? itemGap : 0);
      });
      if (by > cy + ch - pad + 1e-6) throw new Error(`S4 カード${i + 1}: 高さ不足（${by.toFixed(2)}）`);
      x += w + gap;
    });
  }

  // ===== S5 深さの段階と拮抗 =====
  {
    const s = newSlide({ title: '深さの段階と拮抗', chapter: 3, notes: notesOf('S5'), sources: sourcesOf('S5') });
    const st = t => ({ text: t, bold: true });
    const hl = t => ({ text: t, bold: true, color: C.white, fill: TEAL });   // 白太字が約 4.8:1
    s.table([
      ['段階', '指標', 'スガマデクスの目安'],
      [st('完全'), 'PTC 0', '待つ\n（挿管直後の緊急時は 16 mg/kg）'],
      [st('深い'), 'PTC ≥1、TOF カウント 0', 'PTC 1–2 が出たら 4 mg/kg'],
      [st('中等度'), 'TOF カウント 1–3', 'T2 再出現後に 2 mg/kg'],
      [st('浅い'), 'TOF カウント 4、TOF比 <0.4', '2 mg/kg'],
      [st('最小'), 'TOF比 0.4〜0.9 未満', '2 mg/kg（ネオスチグミンも可）'],
      [hl('回復'), hl('TOF比 ≥0.9（定量）'), hl('不要。抜管できる')],
    ], { x: X0, y: 1.85, colW: [1.7, 4.98, 5.55], size: 22, margin: [2, 7, 2, 7] });
    // 列幅は spec の 2.0/4.2/6.0 から 1.7/4.98/5.55 に（「浅い」の指標 TOF カウント 4、TOF比 <0.4 を余裕をもって
    // 1 行に収める。3 列目は「（挿管直後の緊急時は 16 mg/kg）」が 1 行で入る幅を残す。1 列目は最長「中等度」3 字）
  }

  // ===== S6 定性と定量の違い =====
  {
    // 出典は全体で 2 行に入らないため、JSA FAQ の表題を略記にする
    const s = newSlide({
      title: '定性と定量の違い', chapter: 3, notes: notesOf('S6'),
      sources: sourcesLines('S6',
        'Thilen SR, et al. Anesthesiology 2023;138:13-41 / Fuchs-Buder T, et al.',
        'Eur J Anaesthesiol 2023;40:82-94 / 日本麻酔科学会 モニター指針 FAQ（2025年10月更新）'),
    });
    // 左右 2 枚の大きなカード。本文 24pt は「TOF比 0.4 以上の減衰は分からない」が半幅に入らないため 22pt のまま、
    // 行間を広げてカードを高くする。左の結論は色（レッド）ではなく太字＋枠線で示す（見出しのバーミリオンと紛れない）。
    const gap = 0.3, cw = (BW - gap) / 2, cy = 1.9, pad = 0.15, pitch = 0.78, frm = 0.12;
    const tS = 26, tH = heightFor(1, tS, 0), lH = heightFor(1, 22, 0);
    const l0 = cy + 0.18 + tH + 0.22;                       // 本文1行目の上端
    const ch = Math.ceil((l0 + 2 * pitch + lH + frm + 0.2 - cy) * 100) / 100;
    const col = (x, title, titleColor, lines) => {
      s.box({ x, y: cy, w: cw, h: ch, fill: C.bg, radius: 0.08 });
      s.text(title, { x: x + pad, y: cy + 0.18, w: cw - 2 * pad, h: tH, size: tS, bold: true, color: titleColor });
      lines.forEach((l, i) => {
        const y = l0 + i * pitch;
        if (l.frame) {
          // 結論: 白地・枠線つきの箱（文字は図形内）。文字の左端を他の行にそろえる
          s.box({ x: x + pad - frm, y: y - frm, w: cw - 2 * pad + 2 * frm, h: lH + 2 * frm, fill: C.white,
            line: { color: C.vermilion, width: 2.25 }, radius: 0.06, margin: frm,
            text: l.frame, size: 22, bold: true, color: C.dark, align: 'left' });
        } else {
          s.text([l], { x: x + pad, y, w: cw - 2 * pad, h: lH, size: 22 });
        }
      });
    };
    col(X0, '定性（目・指で見る）', C.vermilion,
      ['カウント（0–4）は分かる', 'TOF比 0.4 以上の減衰は分からない', { frame: '0.9 の確認には使えない' }]);
    col(X0 + cw + gap, '定量（AMG・EMG）', TEAL,
      ['TOF比が数字で出る', [em('0.9 を確認できる唯一の方法', TEAL)], '記録に残せる']);
    s.text([['AMG は投与前から TOF比 >1.0 が多い → 投与前の値で割って', em('正規化', TEAL)]],
      { x: X0, y: cy + ch + 0.45, w: BW, size: 24 });
  }

  // ===== S7 部位の選び方 =====
  {
    const s = newSlide({ title: '部位の選び方', chapter: 3, notes: notesOf('S7'), sources: sourcesOf('S7') });
    // 見出しの色: 母指は濃いティール（F2F2F2 地で約 4.3:1）。カードの高さは項目の多い中央に合わせて詰め、
    // 本文は見出しの下で上下中央に置く（項目の少ない左右のカードが上に偏らない）。
    const gap = 0.3, cw = (BW - 2 * gap) / 3, cy = 1.95, ch = 4.02;
    const cards = [
      { title: '母指内転筋（推奨）', titleColor: TEAL, items: [['抜管の判断はここで'], ['尺骨神経を刺激']] },
      { title: '眼の周りの筋', titleColor: C.vermilion,
        items: [['早く回復して見える'], ['モニタリングに', '使わない（ASA）'], ['使ったら', '拮抗前に母指へ']] },
      { title: '片麻痺の患者', titleColor: C.blue, items: [['麻痺側は', '効きにくく見える'], ['健側で測る']] },
    ];
    cards.forEach((c, i) => itemCard(s, { x: X0 + i * (cw + gap), y: cy, w: cw, h: ch, size: 24, titleSize: 24,
      top: 0.18, titleGap: 0.2, itemGap: 0.24, center: true, ...c }));
  }

  // ===== S8 手順：装着から抜管まで =====
  {
    const s = newSlide({ title: '手順：装着から抜管まで', chapter: 4, notes: notesOf('S8'), sources: sourcesOf('S8') });
    // 上段 3 つ・下段 2 つ（下段は中央寄せ）。チップと矢印の間は 0.26in あける
    const cw = 3.4, ag = (BW - 3 * cw) / 2, chh = 1.15, ins = 0.26;
    const y1 = 1.9, y2 = y1 + chh + 0.66;
    const xs = [0, 1, 2].map(k => X0 + k * (cw + ag));
    const x2 = X0 + (BW - (2 * cw + ag)) / 2, xs2 = [x2, x2 + cw + ag];
    const chip = (x, y, text, last) => s.box({
      x, y, w: cw, h: chh, radius: 0.1, text, size: 24,
      fill: last ? TEAL : C.bg, color: last ? C.white : C.dark, bold: !!last,
    });
    chip(xs[0], y1, '① 導入前に装着');
    chip(xs[1], y1, '② 筋弛緩薬の前に\n基準値');
    chip(xs[2], y1, '③ 術中は\nTOF を記録');
    chip(xs2[0], y2, '④ 深さに応じて\n拮抗');
    chip(xs2[1], y2, '⑤ TOF比 ≥0.9\nで抜管', true);
    const ar = (xa, xb, y) => s.arrow({ x1: xa + ins, y1: y, x2: xb - ins, y2: y, color: '404040', width: 3 });
    ar(xs[0] + cw, xs[1], y1 + chh / 2);
    ar(xs[1] + cw, xs[2], y1 + chh / 2);
    ar(xs2[0] + cw, xs2[1], y2 + chh / 2);
    // ③ → ④ の折り返し（③の下端中央 → 左へ → ④の上端中央）
    const c3 = xs[2] + cw / 2, c4 = xs2[0] + cw / 2, yA = y1 + chh + 0.1, ym = y1 + chh + 0.33;
    s.vline({ x: c3, y: yA, h: ym - yA, color: '404040', width: 3 });
    s.hline({ x: c4, y: ym, w: c3 - c4, color: '404040', width: 3 });
    s.arrow({ x1: c4, y1: ym, x2: c4, y2: y2 - 0.1, color: '404040', width: 3 });
    s.bullets([
      ['拮抗後も', em('再クラーレ化', C.vermilion), 'に注意（量はモニターで決める）'],
      '値と時刻を麻酔記録に残す',
    ], { y: y2 + chh + 0.4, size: 24, gap: 12 });
  }

  // ===== S9 落とし穴 =====
  {
    const s = newSlide({ title: '落とし穴', chapter: 4, notes: notesOf('S9'), sources: sourcesOf('S9') });
    const items = [
      '電極の乾き・冷えた手で反応が小さく出る',
      '腕が体側に固定され、母指が動けない（AMG）',
      '基準値を取らずに TOF比を読む',
      '目の周りの筋で回復と判断する',
      '臨床所見だけで抜管する',
    ];
    const size = 28, lh = heightFor(1, size, 0), pitch = 0.92, d = 0.26;
    const tx = X0 + 0.6;
    items.forEach((t, i) => {
      const y = 2.05 + i * pitch;
      s.circle({ cx: X0 + 0.25, cy: y + lh / 2 + 0.02, d, fill: C.red });
      s.text(t, { x: tx, y, w: XR - tx, h: lh, size });
    });
  }

  // ===== S10 まとめ =====
  {
    const s = newSlide({ title: 'まとめ', chapter: 4, notes: notesOf('S10'), sources: sourcesOf('S10') });
    s.bullets([
      ['筋弛緩薬を使うなら、', em('定量モニター', TEAL), 'を必ず付ける'],
      '尺骨神経を刺激し、母指内転筋で測る',
      '深さは PTC と TOF カウント、回復は TOF比',
      '拮抗薬の量は深さで決める（2・4 mg/kg）',
      ['抜管前に ', em('TOF比 0.9 以上', TEAL), 'を確認する（JSA 2025）'],
    ], { y: 2.0, size: 28, gap: 22 });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
