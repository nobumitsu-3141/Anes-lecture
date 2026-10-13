'use strict';
/**
 * 2-8 脳波モニターと術中覚醒
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言はここに転記（改行位置のみ調整）、ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, fits, measure, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '2-8_eeg_awareness';
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
/** ノート（spec の notes 全文） */
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
 * 出典が2行（16pt・11.6in）に収まらないスライド用: 略記した2行を明示する。
 * spec の sources 行に含まれない語が混じっていないか（短縮のみか）を確認する。
 */
function sourcesLines(id, l1, l2) {
  const s = SPEC[id] && SPEC[id].sources;
  if (!s) throw new Error(`${id}: spec に sources がありません`);
  const S = LAYOUT.source;
  const strip = t => t.replace(/[.,:;]+$/, '');
  for (const l of [l1, l2]) {
    if (!fits(l, S.size, S.w)) throw new Error(`${id} 出典: 1行に収まりません「${l}」`);
    for (const tok of l.split(/\s+/).filter(t => t && t !== '/')) {
      if (!s.includes(tok) && !s.includes(strip(tok))) throw new Error(`${id} 出典: spec にない語「${tok}」`);
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
const r3 = v => Math.round(v * 1000) / 1000;

/**
 * 文字用のティール。C.teal(00A8AA) は白地 約 2.9:1・F2F2F2 地 約 2.6:1 と淡いので、
 * 文字と現在章チップの地には濃いティール（白地 5.3:1・F2F2F2 地 4.7:1）を使う。図形の塗り・線は C.teal のまま。
 */
const TEAL_T = C.teal;   // 書式ルールのティール（シリーズ共通）
/** 覚醒（浅い側の基準）の色。S5 の目盛り「100–80 覚醒」の灰と対応させる（白地 7:1） */
const AWAKE = '595959';

/**
 * 章チップの上書き（この講義だけ）。decklib は 11pt・幅 0.68in・上端 0.06in・右端 13.13in で置く。
 * slide_lint の小サイズ許容は 10.5/12/16pt なので 16pt にし、幅 0.76in、上端をタイトル枠の上端 0.35in に、
 * 右端を本文の右端 12.78in に揃える。現在章の地は濃いティール（白太字 16pt で 5.3:1）。
 * lint のパンくず判定（上端<1.9cm・高さ<1.9cm・幅<2.2cm・左>17cm）の範囲に収める。
 * pptxgenjs のスライドオブジェクト（書き出し前）を直接書き換える。
 */
const CHIP = { pt: 16, w: 0.76, h: 0.6, y: 0.35, gap: 0.06, right: XR };
function restyleChips(s) {
  const ch = LAYOUT.chip;
  const chips = s.raw._slideObjects.filter(o => o._type === 'text' && o.options && o.options.shape === 'roundRect'
    && Math.abs(o.options.y - ch.y) < 1e-6 && Math.abs(o.options.h - ch.h) < 1e-6);
  if (chips.length !== 5) throw new Error(`p${s.page}: 章チップが見つかりません（${chips.length} 個）`);
  const cur = chips.filter(o => o.options.fill && String(o.options.fill.color).toUpperCase() === C.teal);
  if (cur.length !== 1) throw new Error(`p${s.page}: 現在章のチップが 1 個ではありません（${cur.length} 個）`);
  let x = CHIP.right - chips.length * CHIP.w - (chips.length - 1) * CHIP.gap;
  chips.forEach(o => {
    o.options.x = r3(x);
    o.options.y = CHIP.y;
    o.options.w = CHIP.w;
    o.options.h = CHIP.h;
    o.options.fontSize = CHIP.pt;
    if (o === cur[0]) o.options.fill = { color: TEAL_T };
    o.text.forEach(t => { if (t.options) t.options.fontSize = CHIP.pt; });
    x += CHIP.w + CHIP.gap;
  });
  return s;
}

const runText = l => (Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : (typeof l === 'string' ? l : l.text));

/**
 * 見出し＋本文の項目（各項目は1〜2行）を縦に並べたカード。文字はすべて箱の内側。
 * items: [[行, 行], [行], ...]（1項目 = 行の配列。2行目は同じ項目の続き）
 */
function itemCard(s, o) {
  const pad = o.pad != null ? o.pad : 0.15;
  const itemGap = o.itemGap != null ? o.itemGap : 0.2;
  const size = o.size || 22, tsize = o.titleSize || 24;
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill || C.bg, radius: 0.08 });
  const iw = o.w - 2 * pad;
  const tl = Array.isArray(o.title) ? o.title : [o.title];
  const th = heightFor(tl.length, tsize, 0);
  const tg = o.titleGap != null ? o.titleGap : 0.14;
  const ihs = o.items.map(lines => heightFor(lines.length, size, 0));
  const itemsH = ihs.reduce((a, b) => a + b, 0) + itemGap * (o.items.length - 1);
  const contentH = th + tg + itemsH;
  let cy = o.y + (o.center ? Math.max(pad, (o.h - contentH) / 2) : (o.top != null ? o.top : pad));
  s.text(tl, { x: o.x + pad, y: cy, w: iw, h: th, size: tsize, bold: true, color: o.titleColor });
  cy += th + tg;
  // itemsCenter: 見出しは上に置いたまま、本文の項目群を見出しの下の残りの高さの中で上下中央に置く
  if (o.itemsCenter) cy += Math.max(0, (o.y + o.h - (o.bottom != null ? o.bottom : pad) - cy - itemsH) / 2);
  o.items.forEach((lines, i) => {
    const h = ihs[i];
    s.text(lines, { x: o.x + pad, y: cy, w: iw, h, size });
    cy += h + (i < o.items.length - 1 ? itemGap : 0);
  });
  if (cy > o.y + o.h - pad + 1e-6) throw new Error(`itemCard「${tl.join('')}」: 高さ不足（必要 ${(cy + pad - o.y).toFixed(2)}in）`);
}
/** itemCard の内容に必要な最小幅（最長行 + 左右余白） */
function itemCardMinW(o) {
  const pad = o.pad != null ? o.pad : 0.15;
  const size = o.size || 22, tsize = o.titleSize || 24;
  const tl = Array.isArray(o.title) ? o.title : [o.title];
  let w = Math.max(...tl.map(t => measure(runText(t), tsize)));
  o.items.forEach(lines => lines.forEach(l => { w = Math.max(w, measure(runText(l), size)); }));
  return w + 2 * pad + 0.02;
}
/** itemCard の内容に必要な高さ（上余白 top + 見出し + 本文。下余白は含まない） */
function itemCardMinH(o) {
  const size = o.size || 22, tsize = o.titleSize || 24;
  const tl = Array.isArray(o.title) ? o.title : [o.title];
  const tg = o.titleGap != null ? o.titleGap : 0.14, itemGap = o.itemGap != null ? o.itemGap : 0.2;
  const top = o.top != null ? o.top : (o.pad != null ? o.pad : 0.15);
  return top + heightFor(tl.length, tsize, 0) + tg
    + o.items.reduce((a, lines) => a + heightFor(lines.length, size, 0), 0) + itemGap * (o.items.length - 1);
}
/**
 * 横に並べるカードの幅: できるだけ等幅にする（各カードは最小幅以上。足りないカードだけ広げ、残りを等分）。
 */
function balancedWidths(minWs, total, gap) {
  const T = total - gap * (minWs.length - 1);
  if (minWs.reduce((a, b) => a + b, 0) > T + 1e-9) throw new Error('カード幅が足りません');
  let fixed = new Set();
  for (;;) {
    const free = minWs.map((w, i) => i).filter(i => !fixed.has(i));
    const c = (T - [...fixed].reduce((a, i) => a + minWs[i], 0)) / free.length;
    const grow = free.filter(i => minWs[i] > c + 1e-9);
    if (!grow.length) return minWs.map((w, i) => (fixed.has(i) ? w : c));
    grow.forEach(i => fixed.add(i));
  }
}

// ---------------------------------------------------------------------------
// 脳波の模式波形（決定的な擬似乱数）
// ---------------------------------------------------------------------------
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
/** なめらかな雑音（値ノイズ）: 周波数 f [Hz]、0〜T 秒、振幅 ±1 */
function valueNoise(seed, f, T) {
  const r = rng(seed), n = Math.ceil(f * T) + 3, v = Array.from({ length: n }, () => r() * 2 - 1);
  return t => {
    const u = t * f, i = Math.floor(u), k = u - i, s = k * k * (3 - 2 * k);
    return v[i] * (1 - s) + v[i + 1] * s;
  };
}
/** 波形の点列: 横 x0〜x1 を T 秒とみなし、fn(t)（インチ、上向き正）を中心 cy に描く */
function wavePts(fn, x0, x1, cy, T = 2, step = 0.012) {
  const pts = [];
  for (let x = x0; x <= x1 + 1e-9; x += step) {
    const t = (x - x0) / (x1 - x0) * T;
    pts.push([+x.toFixed(3), +(cy - fn(t)).toFixed(3)]);
  }
  return pts;
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['覚醒', '脳波', 'BIS', '限界', '予防'],
    author: 'Anes-lecture',
    deckTitle: '2-8 脳波モニターと術中覚醒',
  });
  const newSlide = o => deck.slide(o);   // 章チップはライブラリ既定（シリーズ共通）

  // ===== S1 表紙 =====
  deck.cover({
    title: '脳波モニターと術中覚醒',
    subtitle: '2-8 モニタリング（発展）',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 術中覚醒の頻度 =====
  {
    // 出典は全体で 2 行（16pt）に入らないため、NAP5 論文の表題を略記
    const s = newSlide({
      title: '術中覚醒の頻度', chapter: 0, notes: notesOf('S2'),
      sources: sourcesLines('S2',
        'Pandit JJ, et al. 5th National Audit Project (NAP5). Br J Anaesth 2014;113:549-559',
        'Sebel PS, et al. Anesth Analg 2004;99:833-839'),
    });
    // 色の意味: 青 = 基準（全体・筋弛緩薬なし）、バーミリオン = 多くなる場面（筋弛緩薬・帝王切開）
    const gap = 0.33, cw = (BW - gap) / 2, ch = 1.65, rg = 0.28;
    const cells = [
      { num: '1/19,000', cap: '自己申告（NAP5、全体）', color: C.blue },
      { num: '1/8,200', cap: '筋弛緩薬を使った麻酔', color: C.vermilion },
      { num: '1/135,900', cap: '筋弛緩薬を使わない麻酔', color: C.blue },
      { num: '1/670', cap: '帝王切開（全身麻酔）', color: C.vermilion },
    ];
    const nh = heightFor(1, 44, 0), caph = heightFor(1, 22, 0);
    cells.forEach((c, i) => {
      const x = X0 + (i % 2) * (cw + gap), y = 1.85 + Math.floor(i / 2) * (ch + rg);
      s.box({ x, y, w: cw, h: ch, fill: C.bg, radius: 0.08 });
      s.text([[{ text: '約 ', size: 28, color: C.dark }, { text: c.num, size: 44, bold: true, color: c.color }]],
        { x: x + 0.15, y: y + 0.1, w: cw - 0.3, h: nh, size: 44, align: 'center', valign: 'middle' });
      s.text(c.cap, { x: x + 0.15, y: y + 0.1 + nh + 0.04, w: cw - 0.3, h: caph, size: 22, align: 'center' });
    });
    const ny = 1.85 + 2 * ch + rg + 0.27;
    s.text([['面接で聞き出すと ', em('1–2 / 1,000', TEAL_T), ' の報告もある（Brice 法）']],
      { x: X0, y: ny, w: BW, size: 24, align: 'center' });
  }

  // ===== S3 危険因子 =====
  {
    const s = newSlide({ title: '危険因子', chapter: 0, notes: notesOf('S3'), sources: sourcesOf('S3') });
    // 3 枚をできるだけ等幅に（麻酔の「若手の担当・時間外・緊急」は「・」の位置で 2 行）。
    // 本文は上揃え（3 枚で行の高さをそろえる）。カードの高さは項目のいちばん多い「麻酔」に合わせて詰める。
    const common = { size: 22, titleSize: 26, top: 0.25, titleGap: 0.26, itemGap: 0.26 };
    const cards = [
      { title: '患者', titleColor: C.blue, items: [['術中覚醒の既往'], ['若年成人・女性・肥満']] },
      { title: '麻酔', titleColor: C.vermilion,
        items: [[[em('筋弛緩薬', C.vermilion)]], ['TIVA（濃度が見えない）'], ['導入薬が切れる空白'], ['若手の担当・', '時間外・緊急']] },
      { title: '手術', titleColor: TEAL_T, items: [['帝王切開'], ['心臓・胸部手術'], ['循環が不安定で', '浅くする']] },
    ].map(c => ({ ...common, ...c }));
    const gap = 0.3, ws = balancedWidths(cards.map(itemCardMinW), BW, gap);
    const h = Math.max(...cards.map(itemCardMinH)) + 0.25;
    let x = X0;
    cards.forEach((c, i) => { itemCard(s, { x, y: 1.95, w: ws[i], h, ...c }); x += ws[i] + gap; });
  }

  // ===== S4 脳波は麻酔で変わる =====
  {
    // 出典は全体で 1 行に入らないため 1 行 1 文献。Purdon 2015 は表題を略し著者と誌名を残す
    const s = newSlide({
      title: '脳波は麻酔で変わる', chapter: 1, notes: notesOf('S4'),
      sources: sourcesLines('S4',
        'Purdon PL, Sampson A, Pavone KJ, Brown EN. Anesthesiology 2015;123:937-960',
        'Fritz BA, et al. Anesth Analg 2016;122:234-242'),
    });
    const wx0 = 0.75, wx1 = 6.25, lx = 6.7, lw = XR - lx, T = 2;
    const TAU = 2 * Math.PI;
    // 振幅は覚醒 → 鎮静 → 全身麻酔で大きくなる（1 行目と 2 行目の差をはっきりさせる）
    // 1) 覚醒: 速く小さい不規則な波（β）
    const n1 = valueNoise(11, 34, T), n1b = valueNoise(12, 21, T);
    const f1 = t => 0.09 * (0.55 * n1(t) + 0.45 * n1b(t) + 0.25 * Math.sin(TAU * 26 * t));
    // 2) 浅い鎮静: 少し大きく、規則的な波（α〜β）が混ざる
    const n2 = valueNoise(21, 3, T), n2b = valueNoise(22, 30, T);
    const f2 = t => 0.28 * (0.55 * Math.sin(TAU * 11 * t) + 0.3 * Math.sin(TAU * 17 * t + 1) + 0.25 * n2(t) + 0.12 * n2b(t));
    // 3) 全身麻酔: 大きくゆっくりした δ に、規則的な α が重なる
    const n3 = valueNoise(31, 2.2, T);
    const f3 = t => 0.27 * Math.sin(TAU * 1.25 * t + 0.4) + 0.08 * n3(t) + 0.08 * Math.sin(TAU * 10 * t);
    // 4) 深すぎ: 平坦な区間と、ときどき出る大きな波（バーストサプレッション）
    const n4 = valueNoise(41, 40, T), n4b = valueNoise(42, 7, T);
    const bursts = [[0.2, 0.62], [1.18, 1.62]];
    const f4 = t => {
      let v = 0.012 * n4(t);
      for (const [a, b] of bursts) {
        if (t > a && t < b) {
          const env = Math.sin(Math.PI * (t - a) / (b - a));
          v += env * (0.19 * Math.sin(TAU * 9 * (t - a)) + 0.16 * Math.sin(TAU * 3 * (t - a)) + 0.07 * n4b(t));
        }
      }
      return v;
    };
    // 4 行を等間隔（1.1in）に。色は状態ごと: 覚醒=灰（S5 の目盛りの 100–80 と対応）、鎮静=青、全身麻酔=ティール、深すぎ=バーミリオン
    const pitch = 1.1, cy0 = 2.25;
    const rows = [
      { fn: f1, color: AWAKE, label: [[em('覚醒：', AWAKE), '速く小さい波（β）']] },
      { fn: f2, color: C.blue, label: [[em('浅い鎮静：', C.blue), 'α と β']] },
      { fn: f3, color: C.teal, label: [[em('全身麻酔：', TEAL_T), '前頭部の α'], ['＋ ゆっくり大きい δ']] },
      { fn: f4, color: C.vermilion, label: [[em('深すぎ：', C.vermilion), 'バーストサプレッション']] },
    ];
    let prevBottom = -1;
    rows.forEach((r, i) => {
      const cy = cy0 + i * pitch;
      const pts = wavePts(r.fn, wx0, wx1, cy, T);
      const top = Math.min(...pts.map(p => p[1])), bottom = Math.max(...pts.map(p => p[1]));
      if (prevBottom > 0 && top - prevBottom < 0.24) throw new Error(`S4: 波形 ${i} と ${i + 1} の間隔が ${r3(top - prevBottom)}in`);
      prevBottom = bottom;
      s.polyline(pts, { color: r.color, width: 2.75 });
      const h = heightFor(r.label.length, 24, 0);
      s.text(r.label, { x: lx, y: cy - h / 2, w: lw, h, size: 24, valign: 'middle' });
    });
  }

  // ===== S5 BIS の仕組みと数字 =====
  {
    // 出典は 5 件で 2 行に入らないため、各文献を「著者, 誌名 年」に略記（1 文献を行の途中で切らない）
    const s = newSlide({
      title: 'BIS の仕組みと数字', chapter: 2, notes: notesOf('S5'),
      sources: sourcesLines('S5',
        'Avidan, N Engl J Med 2008 / Pilge, Anesthesiology 2006 / Zanner, Br J Anaesth 2009',
        'Messner, Anesth Analg 2003 / Bruhn, J Clin Monit Comput 2000'),
    });
    // --- 右: 縦の目盛り（0〜100 を 5 段階に色分け）＋各段の右にラベル
    const bw = 0.55, lw = 4.2, lgap = 0.25;
    const bx = XR - lw - lgap - bw, lx = bx + bw + lgap;
    const top = 1.95, bot = 6.2, seg = (bot - top) / 5;
    const bands = [
      { fill: 'D9D9D9', label: [[em('100–80', C.dark), ' 覚醒']] },
      { fill: '9DC3E6', label: [[em('80–60', C.dark), ' 鎮静']] },
      { fill: C.teal, label: [[em('60–40 全身麻酔の目安', TEAL_T)]] },
      { fill: 'F4B183', label: [[em('40–20', C.dark), ' 深い（SR を見る）']] },
      { fill: C.vermilion, label: [[em('20 未満', C.dark), ' 平坦が多い']] },
    ];
    const lh = heightFor(1, 22, 0);
    bands.forEach((b, i) => {
      const y = top + i * seg;
      s.box({ x: bx, y, w: bw, h: seg, fill: b.fill, line: { color: C.white, width: 1.5 } });
      s.text(b.label, { x: lx, y: y + (seg - lh) / 2, w: lw, h: lh, size: 22 });
    });
    // --- 左上: 額に貼るセンサーの図（顔の輪郭の角丸四角＋額の帯）
    const fw = 2.7, fh = 2.2, fx = X0 + (7.05 - fw) / 2, fy = 1.95;
    s.box({ x: fx, y: fy, w: fw, h: fh, shape: 'roundRect', radius: 0.75, fill: C.bg, line: { color: C.gray, width: 2 } });
    const sbx = fx + 0.3, sby = fy + 0.32, sbw = fw - 0.6, sbh = 0.36;
    s.box({ x: sbx, y: sby, w: sbw, h: sbh, shape: 'roundRect', radius: 0.12, fill: C.blue });
    [0.22, 0.5, 0.78].forEach(k => s.circle({ cx: sbx + sbw * k, cy: sby + sbh / 2, d: 0.2, fill: C.white }));
    [0.33, 0.67].forEach(k => s.circle({ cx: fx + fw * k, cy: fy + 1.15, d: 0.18, fill: C.gray }));
    s.hline({ x: fx + fw * 0.38, y: fy + 1.7, w: fw * 0.24, color: C.gray, width: 2 });
    // --- 左下: 箇条書き3行
    s.bullets([
      ['0–100 の指標、全身麻酔の目安は ', em('40–60', TEAL_T)],
      '表示は数十秒〜2 分遅れる',
      '同時に見る：SQI・EMG・SR（抑制比）',
    ], { x: X0, y: 4.5, w: 7.0, size: 22, gap: 12 });
  }

  // ===== S6 読むときの3点セット =====
  {
    const s = newSlide({ title: '読むときの3点セット', chapter: 2, notes: notesOf('S6'), sources: sourcesOf('S6') });
    // 3 枚をできるだけ等幅に（本文 22pt。各項目 1 行、SR の補足だけ別行）。
    // 本文は上揃え（3 枚で行の高さをそろえる）。カードの高さは「SR と EMG」に合わせて詰める。
    const common = { size: 22, titleSize: 28, top: 0.3, titleGap: 0.4, itemGap: 0.45 };
    const cards = [
      { title: '数値と傾向', titleColor: C.blue, items: [['一点の値で動かない'], ['トレンドの向きを見る']] },
      { title: '生波形・DSA', titleColor: TEAL_T, items: [['α 帯 ＋ 徐波があるか'], ['平坦な区間がないか']] },
      { title: 'SR と EMG', titleColor: C.vermilion, items: [['SR > 0 ＝ 平坦あり', '（多くは深すぎ）'], ['EMG ↑ ＝ 値が高く出る']] },
    ].map(c => ({ ...common, ...c }));
    const gap = 0.3, ws = balancedWidths(cards.map(itemCardMinW), BW, gap);
    const h = Math.max(...cards.map(itemCardMinH)) + 0.3;
    let x = X0;
    cards.forEach((c, i) => { itemCard(s, { x, y: 1.95, w: ws[i], h, ...c }); x += ws[i] + gap; });
  }

  // ===== S7 限界 =====
  {
    // 出典は全体で 2 行に入らないため、Avidan 2 編をまとめ巻・頁を略記（1 文献を行の途中で切らない）
    const s = newSlide({
      title: '限界', chapter: 3, notes: notesOf('S7'),
      sources: sourcesLines('S7',
        'Avidan MS, et al. N Engl J Med 2008, 2011 / Myles PS, et al. B-Aware. Lancet 2004',
        'Purdon PL, et al. The Ageing Brain. Br J Anaesth 2015;115 Suppl 1:i46-i57'),
    });
    const items = [
      'ケタミン・亜酸化窒素では高めに出る',
      '筋電図などの雑音で乱れる（SQI を見る）',
      '薬ごとに脳波の形が違う',
      '強い低体温で低く出る',
      '小児・高齢者では解釈が変わる',
    ];
    const bwid = 7.5;
    s.bullets(items, { x: X0, y: 2.0, w: bwid, size: 24, gap: 22 });
    const cx = X0 + bwid + 0.3, cw = XR - cx;
    itemCard(s, {
      x: cx, y: 1.95, w: cw, h: 3.8, size: 24, titleSize: 24, center: true, titleGap: 0.3,
      title: ['B-Unaware／', 'BAG-RECALL'], titleColor: TEAL_T,
      items: [['吸入麻酔では BIS は', '呼気麻酔薬濃度の', ['アラームに', em('勝らなかった', C.vermilion)]]],
    });
  }

  // ===== S8 誰に付けるか =====
  {
    // 出典は全体で 2 行に入らないため、NICE DG6 の副題（対象機種名）を略記（1 文献を行の途中で切らない）
    const s = newSlide({
      title: '誰に付けるか', chapter: 4, notes: notesOf('S8'),
      sources: sourcesLines('S8',
        'NICE Diagnostics guidance DG6. Depth of anaesthesia monitors (2012)',
        'Pandit JJ, et al. Br J Anaesth 2014;113:549-559'),
    });
    // 見出しの色: 推奨（付ける場面）= 濃いティール、補足の 2 枚 = 黒
    // 右カードの幅は「NICE DG6：高リスク例と TIVA で選択肢」（22pt）が 1 行に入る最小にして、左を広げる
    const gap = 0.3, y0 = 1.85, rg = 0.26;
    const right = [
      { title: '吸入麻酔なら', items: [['呼気麻酔薬濃度のアラーム'], ['下限 0.7 MAC（試験での設定）']] },
      { title: '根拠', items: [['NICE DG6：高リスク例と TIVA で選択肢'], ['NAP5：TIVA＋筋弛緩薬で推奨']] },
    ].map(c => ({ ...c, size: 22, titleSize: 24, titleGap: 0.12, itemGap: 0.1, center: true, titleColor: C.dark }));
    const rw = Math.ceil(Math.max(...right.map(itemCardMinW)) * 100) / 100;
    const lw = BW - gap - rw, rx = X0 + lw + gap;
    const bottom = s.sourceTop - 0.24;
    const rh = Math.floor(((bottom - y0 - rg) / 2) * 100) / 100;
    itemCard(s, {
      x: X0, y: y0, w: lw, h: 2 * rh + rg, size: 24, titleSize: 28, top: 0.25, titleGap: 0.35, itemGap: 0.42,
      title: '付ける場面', titleColor: TEAL_T,
      items: [[[em('TIVA ＋ 筋弛緩薬', TEAL_T)]], ['覚醒の危険因子がある'], ['深すぎを避けたい', '高齢者・フレイル']],
    });
    right.forEach((c, i) => itemCard(s, { x: rx, y: y0 + i * (rh + rg), w: rw, h: rh, ...c }));
  }

  // ===== S9 予防の実務 =====
  {
    const s = newSlide({ title: '予防の実務', chapter: 4, notes: notesOf('S9'), sources: sourcesOf('S9') });
    // 上段 3 つ、下段 2 つ（中央寄せ）。① は「いちばんの予防」（S10）なので濃いティールで強調
    const cw = 3.55, ag = (BW - 3 * cw) / 2, chh = 1.1;
    const lg = 0.07, leg = 0.32;                        // 折り返し線: チップとの隙間・縦の足
    const y1 = 1.85, y2 = y1 + chh + 2 * lg + 2 * leg;
    const xs1 = [0, 1, 2].map(k => X0 + k * (cw + ag));
    const x2 = X0 + (BW - (2 * cw + ag)) / 2;
    const xs2 = [x2, x2 + cw + ag];
    const chip = (x, y, text, hl) => s.box({
      x, y, w: cw, h: chh, radius: 0.1, text, size: 24, margin: 0.05,
      fill: hl ? TEAL_T : C.bg, color: hl ? C.white : C.dark, bold: !!hl,
    });
    chip(xs1[0], y1, '① 薬が\n届いているか', true);
    chip(xs1[1], y1, '② 導入〜挿管の\n空白をなくす');
    chip(xs1[2], y1, '③ 筋弛緩は\n必要なときだけ');
    chip(xs2[0], y2, '④ 覚醒は\n筋弛緩が戻ってから');
    chip(xs2[1], y2, '⑤ 疑ったら\n認めて説明');
    const AR = { color: '404040', width: 3 };
    const ar = (xa, xb, y) => s.arrow({ x1: xa + 0.15, y1: y, x2: xb - 0.15, y2: y, ...AR });
    ar(xs1[0] + cw, xs1[1], y1 + chh / 2);
    ar(xs1[1] + cw, xs1[2], y1 + chh / 2);
    ar(xs2[0] + cw, xs2[1], y2 + chh / 2);
    // ③ → ④ の折り返し（縦の足 0.32in ずつ。角は線幅の半分だけ重ねて継ぎ目を消す）
    const c3 = xs1[2] + cw / 2, c4 = xs2[0] + cw / 2, ya = y1 + chh + lg, ym = ya + leg, ov = 0.02;
    s.vline({ x: c3, y: ya, h: leg + ov, ...AR });
    s.hline({ x: c4 - ov, y: ym, w: c3 - c4 + 2 * ov, ...AR });
    s.arrow({ x1: c4, y1: ym - ov, x2: c4, y2: y2 - lg, ...AR });
    s.bullets([
      [em('TIVA：', C.dark), '静脈路を目で確認、ポンプの動作と残量、ラインの外れ'],
      [em('覚醒を訴えたら：', C.dark), '否定せず聞く、記録し、フォローにつなぐ'],
    ], { y: y2 + chh + 0.28, size: 24, gap: 10 });   // 下端 6.25in（出典行 6.5in との間 0.25in）
  }

  // ===== S10 まとめ =====
  {
    const s = newSlide({ title: 'まとめ', chapter: 4, notes: notesOf('S10'), sources: sourcesOf('S10') });
    // 強調は締めの 2 行（付ける対象と、いちばんの予防）に絞る。行間を広げて本文域に均す
    // 行間 24pt: 実機メイリオ（1 行 = 1.5em）で 1.95〜6.2in、出典行（6.5in）の上に 0.24in 以上残る
    const y = 1.95, gap = 24;
    s.bullets([
      ['術中覚醒は約 1/19,000、筋弛緩薬で約 1/8,200'],
      ['全身麻酔の脳波は前頭部の α ＋ δ、平坦は深すぎ'],
      ['BIS 40–60 は目安。値と傾向と生波形を一緒に見る'],
      ['TIVA ＋ 筋弛緩薬では', em('脳波モニターを付ける', TEAL_T)],
      ['薬が届いているかの確認が、', em('いちばんの予防', TEAL_T)],
    ], { y, size: 28, gap });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.render && res.render.images) console.log('[build] render: ' + (res.render.images[0] ? path.dirname(res.render.images[0]) : '(none)'));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
