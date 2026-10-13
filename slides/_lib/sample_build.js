'use strict';
/**
 * decklib の動作見本（文言はすべてダミー）。
 *   NODE_PATH=<pptxbuild>/node_modules node slides/_lib/sample_build.js [out.pptx] [--render]
 */
const path = require('path');
const { Deck, C } = require('./decklib');

const OUT = process.argv.find(a => a.endsWith('.pptx')) ||
  path.join(process.env.PPTXBUILD_DIR || process.cwd(), 'sample.pptx');

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['指針', '7項目', '心電図', '血圧', 'まとめ'],
    author: 'decklib sample',
    deckTitle: 'decklib 見本デッキ',
  });

  deck.cover({
    title: '見本デッキのタイトル',
    subtitle: '2-3 見本サブタイトル',
    meta: '初期研修医向け麻酔講義（見本）',
    notes: '表紙のノート（見本）。\n2行目のノート。',
  });

  // --- p1: 箇条書き + 出典 ---
  const s1 = deck.slide({
    title: '箇条書きの見本', chapter: 0,
    notes: 'ここにスピーカーノートの全文を書く（見本）。\n改行は段落として保存される。',
    sources: ['見本学会 見本ガイドライン 2019', '見本著者 他. 見本誌 2020'],
  });
  s1.bullets([
    '一つ目の要点（見本テキスト）',
    ['二つ目は ', { text: '強調語', color: C.vermilion, bold: true }, ' を含む行'],
    '三つ目の要点 SpO2 と EtCO2',
    { text: '下位項目の例（level 1）', level: 1 },
    '五つ目の要点（体言止め）',
  ], { y: 1.95 });

  // --- p2: カード3枚 + 波形 + ラベル + 矢印 ---
  const s2 = deck.slide({ title: 'カードと波形の見本', chapter: 2, notes: '波形はネイティブ図形（custGeom）。' });
  const cw = 3.9, gap = (12.23 - 3 * cw) / 2;
  const cards = [
    { title: '見出しA', lines: ['本文の一行目', '本文の二行目'], titleColor: C.blue },
    { title: '見出しB', lines: ['本文の一行目', '本文の二行目'], titleColor: C.vermilion },
    { title: '見出しC', lines: ['本文の一行目', '本文の二行目'], titleColor: C.teal },
  ];
  cards.forEach((c, i) => s2.card({ x: 0.55 + i * (cw + gap), y: 1.85, w: cw, h: 1.9, ...c }));

  // 心電図風の波形（見本）: 2拍
  const base = 6.3, x0 = 0.7, beat = 4.1;
  const pts = [];
  for (let b = 0; b < 2; b++) {
    const o = x0 + b * beat;
    pts.push([o, base], [o + 0.5, base]);
    for (let k = 0; k <= 8; k++) pts.push([o + 0.5 + k * 0.06, base - 0.22 * Math.sin(Math.PI * k / 8)]); // P
    pts.push([o + 1.3, base], [o + 1.42, base + 0.15], [o + 1.55, base - 1.35], [o + 1.68, base + 0.3], [o + 1.8, base]); // QRS
    pts.push([o + 2.2, base]);
    for (let k = 0; k <= 12; k++) pts.push([o + 2.2 + k * 0.07, base - 0.4 * Math.sin(Math.PI * k / 12)]); // T
    pts.push([o + beat, base]);
  }
  s2.polyline(pts, { color: C.blue, width: 3 });
  // ラベルは波形の上の余白へ（重ねない）
  s2.label('P波', { x: 0.75 + 0.2, y: 5.25, w: 0.9, h: 0.5, color: C.blue });
  s2.label('QRS', { x: 0.7 + 1.55 - 0.5, y: 4.3, w: 1.0, h: 0.5, color: C.blue, bold: true });
  s2.label('T波', { x: 0.7 + 2.62 - 0.45, y: 5.1, w: 0.9, h: 0.5, color: C.blue });
  // 要点ボックス + 矢印
  s2.arrow({ x1: 9.05, y1: 5.6, x2: 9.85, y2: 5.6, color: C.dark, width: 3 });
  s2.box({ x: 10.0, y: 5.05, w: 2.78, h: 1.1, fill: C.bg, radius: 0.08, text: '要点ボックス', size: 24, color: C.red, bold: true });

  // --- p3: 表 ---
  const s3 = deck.slide({ title: '表の見本', chapter: 3, notes: '表はネイティブ table（22pt 以上）。', sources: ['見本出典 2021'] });
  s3.table([
    ['項目', '使うもの', '頻度'],
    ['見本項目A', '見本機器A', '毎回'],
    ['見本項目B', '見本機器B\n（二行目）', '必要時'],
    ['見本項目C', { text: '強調セル', color: C.vermilion, bold: true }, '毎回'],
  ], { x: 0.55, y: 1.85, colW: [3.6, 4.6, 4.03] });
  s3.label('表の下の一文（見本）', { x: 0.55, y: 5.6, w: 12.23, h: 0.6, size: 24, align: 'left' });

  const res = await deck.save({ render: process.argv.includes('--render') });
  if (res.violations !== 0) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
