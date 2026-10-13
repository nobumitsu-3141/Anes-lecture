'use strict';
/**
 * 3-11 局所麻酔薬
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言はここに転記（改行位置のみ調整）、ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 図はすべてネイティブ図形（箱・円・線・freeform）。ラベルは図の外に置く。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fits, fitInfo, heightFor, measure } = require('../_lib/decklib');

const NAME = '3-11_local_anesthetics';
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
/** 出典は ' / ' で項目に分けて渡す（ライブラリが 1〜2 行に詰める） */
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  return s ? s.split(' / ') : undefined;
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const TOP = 1.85;                // 本文の上端
const BOT1 = 6.75;               // 出典 1 行のときの本文下端
const BOT2 = 6.45;               // 出典 2 行（y 6.5）のときの本文下端
const em = (text, color, bold = true) => ({ text, color, bold });
const ARROW_C = '404040';
const MEMBRANE = 'D9D9D9';
const MEMBRANE_LINE = C.gray;    // 膜の 2 本線（帯の上端・下端）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;
const CM = 2.54;

// ---------------------------------------------------------------------------
// 章チップとタイトル枠（この講義だけの上書き）
// ---------------------------------------------------------------------------
/** 章ラベル。4 字の「最大用量」は語の切れ目で 2 行にする（spec: 4 字以内で 2 行） */
const CHAPTERS = ['仕組み', '種類', '最大\n用量', '計算', '限界'];

/**
 * 章チップ。decklib は 11pt・0.68×0.64in・現在章は C.teal 地に白太字で置く。
 * - 文字を 16pt に（slide_lint の小サイズ許容は 10.5/12/16pt で、14pt は違反になる）。
 *   2 行の「最大／用量」は行間を 20pt 固定にして箱（高さ 0.70in）に収める。
 * - 3 字の「仕組み」が 16pt で 1 行に入るよう幅を 0.80in に（右端 13.13in・間隔 0.06in は decklib と同じ）。
 *   左端は 8.89in。最長タイトル（S8、lint 推定の右端 8.76in）とは lint の最小間隔 0.2cm 以上あく。
 * - 現在章の地は書式ルールのティール C.teal のまま、文字を C.dark の太字に
 *   （白は C.teal 地で約 2.9:1、C.dark は約 5.2:1）。
 * lint のパンくず判定（上端 < 1.9cm・高さ < 1.9cm・幅 < 2.2cm・左 > 17cm）の範囲に収める。
 * 書き出し前の pptxgenjs のスライドオブジェクトを直接書き換える。
 */
const CHIP = { pt: 16, ls: 20, w: 0.80, h: 0.70, y: LAYOUT.chip.y, gap: LAYOUT.chip.gap, right: LAYOUT.chip.right };
const CHIP_LEFT = r3(CHIP.right - CHAPTERS.length * CHIP.w - (CHAPTERS.length - 1) * CHIP.gap);
/** タイトル枠の幅。チップ列の左端の 0.06in 手前で止める（decklib 既定 9.3in だとチップ列に 0.36in 以上かかる） */
const TITLE_W = r2(CHIP_LEFT - 0.06 - LAYOUT.title.x);

function restyleChips(s) {
  const ch = LAYOUT.chip;
  const chips = s.raw._slideObjects.filter(o => o._type === 'text' && o.options && o.options.shape === 'roundRect'
    && Math.abs(o.options.y - ch.y) < 1e-6 && Math.abs(o.options.h - ch.h) < 1e-6);
  if (chips.length !== CHAPTERS.length) throw new Error(`p${s.page}: 章チップが見つかりません（${chips.length} 個）`);
  const cur = chips.filter(o => o.options.fill && String(o.options.fill.color).toUpperCase() === C.teal);
  if (cur.length !== 1) throw new Error(`p${s.page}: 現在章のチップが 1 個ではありません（${cur.length} 個）`);
  let x = CHIP_LEFT;
  chips.forEach((o, k) => {
    const ls = CHAPTERS[k].split('\n');
    ls.forEach(l => {
      if (measure(l, CHIP.pt) > CHIP.w - 0.02) throw new Error(`章チップ「${l}」が ${CHIP.pt}pt・幅 ${CHIP.w}in に収まりません`);
    });
    if (ls.length * CHIP.ls / 72 > CHIP.h - 0.1) throw new Error(`章チップ「${CHAPTERS[k]}」が高さ ${CHIP.h}in に収まりません`);
    o.options.x = r3(x);
    o.options.y = CHIP.y;
    o.options.w = CHIP.w;
    o.options.h = CHIP.h;
    o.options.fontSize = CHIP.pt;
    o.text.forEach(t => {
      if (!t.options) return;
      t.options.fontSize = CHIP.pt;
      t.options.lineSpacing = CHIP.ls;
      if (o === cur[0]) t.options.color = C.dark;
    });
    x += CHIP.w + CHIP.gap;
  });
  return s;
}

/** レイアウト KAWAZOE のタイトル枠を TITLE_W に縮める（スライド側の位置・寸法はこの値で決まる） */
function shrinkTitleBox(deck) {
  const lay = deck.pptx.slideLayouts.find(l => l._name === 'KAWAZOE');
  const ph = lay && lay._slideObjects.find(ob => ob.options && ob.options.placeholder === 'title');
  if (!ph) throw new Error('レイアウト KAWAZOE のタイトル枠が見つかりません');
  ph.options.w = TITLE_W;
}

/** 本文スライド: decklib の slide() に、章チップの上書きと、縮めたタイトル枠・チップ列に対する長さの検査を足す */
function contentSlide(deck, o) {
  const s = deck.slide(o);   // 章チップ・タイトル枠はライブラリ既定（全講義で共通）
  const right = LAYOUT.title.x + measure(o.title, LAYOUT.title.size);
  const limit = Math.min(LAYOUT.title.x + TITLE_W, CHIP_LEFT - 0.2 / CM);
  if (right > limit + 1e-9) throw new Error(`p${s.page}「${o.title}」: タイトルの推定右端 ${r3(right)}in > 上限 ${r3(limit)}in`);
  return s;
}

/**
 * 見出し付きカード（手組み）。項目ごとに間隔をあけ、1 項目を 2 行に割ったときは行間を詰める。
 * items: 文字列 | {lines:[行...], color, bold}（行は文字列 or run 配列）
 * marker: true なら各項目の頭に「•」を置き、本文をぶら下げる。
 * pad: 上下の内側余白、padX: 左右の内側余白（省略時は pad）。
 * slots: 項目ごとの枠の高さ（並べたカードどうしで i 番目の項目の高さをそろえる用）。項目は枠の上下中央に置く。
 */
function cardBody(o) {
  const ts = o.titleSize || 26, bs = o.size || 24;
  const pad = o.pad != null ? o.pad : 0.18;
  const padX = o.padX != null ? o.padX : pad;
  const tgap = o.titleGap != null ? o.titleGap : 0.14;
  const igap = o.itemGap != null ? o.itemGap : 0.18;
  const items = o.items.map(it => (typeof it === 'string' ? { lines: [it] } : it));
  const hs = items.map(it => heightFor(it.lines.length, bs, 0));
  const slots = o.slots ? hs.map((h, i) => Math.max(h, o.slots[i])) : hs;
  const need = 2 * pad + heightFor(1, ts, 0) + tgap + slots.reduce((a, b) => a + b, 0) + igap * (items.length - 1);
  return { ts, bs, pad, padX, tgap, igap, items, hs, slots, need };
}
function listCardH(o) { return Math.ceil(cardBody(o).need * 100) / 100; }
function listCard(s, o) {
  const b = cardBody(o);
  if (b.need > o.h + 1e-6) throw new Error(`カード「${o.title}」: 高さ不足（必要 ${b.need.toFixed(2)}in > h ${o.h}in）`);
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill || C.bg, radius: 0.08, line: o.line });
  const iw = o.w - 2 * b.padX;
  let y = o.y + b.pad;
  s.text(o.title, { x: o.x + b.padX, y, w: iw, size: b.ts, bold: true, color: o.titleColor || C.dark });
  y += heightFor(1, b.ts, 0) + b.tgap;
  const ind = o.marker ? (o.indent || 0.34) : 0;
  b.items.forEach((it, i) => {
    const iy = r2(y + (b.slots[i] - b.hs[i]) / 2);
    if (o.marker) s.text('•', { x: o.x + b.padX, y: iy, w: ind, h: heightFor(1, b.bs, 0), size: b.bs, color: C.dark });
    s.text(it.lines, { x: o.x + b.padX + ind, y: iy, w: iw - ind, h: b.hs[i], size: b.bs, color: it.color || C.dark, bold: !!it.bold });
    y += b.slots[i] + b.igap;
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
    s.text('•', { x: o.x, y, w: ind, h: heightFor(1, size, 0), size, color: C.dark });
    s.text(lines, { x: o.x + ind, y, w: o.w - ind, h: hs[i], size, color: C.dark });
    y += hs[i] + gap;
  });
  return { top: r2(top), bottom: r2(top + total) };
}

/**
 * 水平・垂直の矢印（軸＋塗りの三角の矢じり）。pptxgenjs の矢じりは線幅に比例して小さいので、
 * 矢じりを自前の三角形（freeform）で描いて大きさを決める。hl: 矢じりの長さ、hw: 矢じりの幅(in)
 * from: 折れ線の途中点（[[x,y],...]）。最後の区間は (x1,y1)→(x2,y2) と同じ向きであること。
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

/**
 * 行の多い表（S4・S5 用）。decklib の s.table と同じ見た目（見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、
 * 上下中央）で、行間だけを固定値 LS pt（= 1.2 em）に詰める。見出し行も本文と同じ左揃え（列の頭をそろえる）。
 * s.table は行高を実機メイリオの 1.5 em で取るため、11 行ある表が本文ゾーン（出典 2 行時 1.85〜6.45in）に
 * 入らない。セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、重なり検査用の登録も同じく行う。
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

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: CHAPTERS,
    author: 'Anes-lecture',
    deckTitle: '3-11 局所麻酔薬',
  });

  // ===== S1 表紙 =====
  deck.cover({
    title: '局所麻酔薬',
    subtitle: '3-11 麻酔薬理',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 どうやって効くか =====
  {
    const s = contentSlide(deck, { title: 'どうやって効くか', chapter: 0, notes: notesOf('S2'), sources: sourcesOf('S2') });
    // --- 図: 横長の膜（灰の帯＋上端・下端の 2 本線）に Na チャネル（青の 2 つの箱。間が孔）
    const MT = 3.75, MB = 4.45;                 // 膜の上端・下端
    const CH_T = 3.5, CH_B = 4.7;               // チャネルの上端・下端（膜より少し出す）
    const CHL = 3.95, CHW = 0.55, PORE = 0.5;   // 左半分の左端・幅、孔の幅
    const CX = CHL + CHW + PORE / 2;            // 孔の中心 x（4.75）
    const CHR = CHL + 2 * CHW + PORE;           // 右半分の右端（5.55）
    const FX = 2.9;                              // 分子が膜を通る x
    const FR = 6.05;                             // 図の右端
    s.box({ x: X0, y: MT, w: CHL - X0, h: MB - MT, fill: MEMBRANE });
    s.box({ x: CHR, y: MT, w: FR - CHR, h: MB - MT, fill: MEMBRANE });
    for (const yl of [MT, MB]) {
      s.hline({ x: X0, y: yl, w: r2(CHL - X0), color: MEMBRANE_LINE, width: 2 });
      s.hline({ x: CHR, y: yl, w: r2(FR - CHR), color: MEMBRANE_LINE, width: 2 });
    }
    s.box({ x: CHL, y: CH_T, w: CHW, h: CH_B - CH_T, fill: C.blue, radius: 0.08 });
    s.box({ x: CHL + CHW + PORE, y: CH_T, w: CHW, h: CH_B - CH_T, fill: C.blue, radius: 0.08 });
    // 膜の外の非イオン型（白抜き）→ 膜を通過 → 内側のイオン型（塗り）→ 孔を内側から塞ぐ
    s.circle({ cx: FX, cy: 2.95, d: 0.55, fill: C.white, line: { color: C.vermilion, width: 3 } });
    flowArrow(s, FX, 3.35, FX, 4.95);
    s.circle({ cx: FX, cy: 5.3, d: 0.55, fill: C.vermilion });
    flowArrow(s, CX, 5.3, CX, 4.42, { from: [[FX + 0.4, 5.3], [CX, 5.3]] });
    s.circle({ cx: CX, cy: 4.1, d: 0.4, fill: C.vermilion });
    // ラベル（図の外の余白へ）
    s.label('非イオン型', { x: FX - 1.0, y: 1.95, w: 2.0, h: 0.46, color: C.vermilion, bold: true });
    s.label('Na チャネル', { x: CX - 1.0, y: 2.8, w: 2.0, h: 0.46, color: C.blue, bold: true });
    s.label('細胞の外', { x: X0, y: 3.05, w: 1.6, h: 0.46, align: 'left', color: C.source });
    s.label('細胞の内側', { x: X0, y: 4.69, w: 1.9, h: 0.46, align: 'left', color: C.source });
    s.label('イオン型', { x: FX - 0.8, y: 5.82, w: 1.6, h: 0.46, color: C.vermilion, bold: true });
    // --- 右: 箇条書き 3 項目
    bulletCol(s, [
      ['Na チャネルを細胞の', '内側から塞ぎ、活動電位を止める'],
      ['非イオン型が膜を通り、', 'イオン型がチャネルに効く'],
      ['炎症（酸性）の組織では効きにくい'],
    ], { x: 6.35, w: XR - 6.35, size: 24, gap: 0.36, yc: 4.12 });
  }

  // ===== S3 性質を決める 3 つ =====
  {
    const s = contentSlide(deck, { title: '性質を決める 3 つ', chapter: 0, notes: notesOf('S3'), sources: sourcesOf('S3') });
    const cw = 3.9, gap = (BW - 3 * cw) / 2;
    // 3 枚の色分けは枠線（ブルー／バーミリオン／ティール 2.5pt）で示し、見出しは C.dark の太字にする。
    // ティール 00A8AA の文字は地に対して約 2.6〜2.9:1 で読みにくく、ティールは書式ルールの値をシリーズ共通で使うため。
    const cards = [
      { title: 'pKa → 発現の速さ', edge: C.blue, items: ['低いほど速い', 'リドカイン 7.9', 'ブピバカイン 8.1'] },
      {
        title: '脂溶性 → 強さ', edge: C.vermilion, items: [
          '高いほど少量で効く',
          { lines: ['ブピバカインは', 'リドカインの約 4 倍'] },
          { lines: ['中毒も強い'], color: C.red, bold: true },
        ],
      },
      { title: '蛋白結合 → 持続', edge: C.teal, items: ['高いほど長い', 'ブピバカイン 約 95%', 'リドカイン 約 65%'] },
    ];
    const common = { titleSize: 26, size: 24, pad: 0.18, titleGap: 0.16, itemGap: 0.22, fill: C.white, titleColor: C.dark };
    // i 番目の項目の枠の高さを 3 枚でそろえる（2 行の項目に合わせ、1 行の項目は枠の上下中央）。
    // カード 1・3 の下だけが空くのを防ぎ、横に並んだ項目の高さもそろう。
    const nLines = it => (typeof it === 'string' ? 1 : it.lines.length);
    const slots = cards[0].items.map((_, i) => heightFor(Math.max(...cards.map(c => nLines(c.items[i]))), common.size, 0));
    const h = Math.max(...cards.map(c => listCardH({ ...common, ...c, slots })));
    const y = r2((TOP + BOT2 - h) / 2);
    cards.forEach((c, i) => listCard(s, {
      ...common, ...c, slots, line: { color: c.edge, width: 2.5 }, x: r2(X0 + i * (cw + gap)), y, w: cw, h,
    }));
  }

  // ===== S4 種類と作用時間 =====
  {
    const s = contentSlide(deck, { title: '種類と作用時間', chapter: 1, notes: notesOf('S4'), sources: sourcesOf('S4') });
    const LONG = '3〜8 時間\n（伝達は 12 時間前後）';
    const t = tightTable(s, [
      ['薬', '発現', '持続（目安）', '主な用途'],
      ['リドカイン', '速い', '1〜2 時間', '浸潤、硬膜外、伝達、\n気道の表面麻酔'],
      ['メピバカイン', '速い', '1.5〜3 時間', '浸潤、伝達'],
      ['ブピバカイン', 'やや遅い', LONG, '脊髄くも膜下、\n硬膜外、伝達'],
      ['レボブピバカイン', 'やや遅い', LONG, '硬膜外、伝達\n（ブピバカインより\n心毒性が少ない）'],
      ['ロピバカイン', 'やや遅い', LONG, '硬膜外、伝達、\n術後の持続'],
    ], { x: X0, y: TOP, colW: [2.92, 1.56, 3.86, BW - 2.92 - 1.56 - 3.86] });
    if (t.y + t.h > BOT2 + 1e-6) throw new Error(`S4 表の下端 ${t.y + t.h} > ${BOT2}`);
  }

  // ===== S5 添付文書の最大用量 =====
  {
    const s = contentSlide(deck, { title: '添付文書の最大用量', chapter: 2, notes: notesOf('S5'), sources: sourcesOf('S5') });
    const dose = text => ({ text, bold: true });
    const t = tightTable(s, [
      ['薬', '添付文書の\n基準最高用量（1 回）', '注記'],
      ['リドカイン', dose('200 mg'), 'アドレナリン含有は 500 mg\n（1% で 50 mL）'],
      ['メピバカイン', dose('500 mg'), '硬膜外・伝達・浸潤とも\n（0.5〜2% 共通）'],
      ['ブピバカイン', dose('2 mg/kg'), '脊髄くも膜下は 20 mg まで'],
      ['レボブピバカイン', dose('150 mg'), '伝達・硬膜外とも。\n0.25% 伝達は 1 回 100 mg'],
      ['ロピバカイン', dose('300 mg（伝達）、\n150〜200 mg（硬膜外）'), '伝達は 0.75% 40 mL'],
    ], { x: X0, y: TOP, colW: [2.92, 4.25, BW - 2.92 - 4.25] });
    if (t.y + t.h > BOT2 + 1e-6) throw new Error(`S5 表の下端 ${t.y + t.h} > ${BOT2}`);
  }

  // ===== S6 体重あたりの目安 =====
  {
    const s = contentSlide(deck, { title: '体重あたりの目安', chapter: 2, notes: notesOf('S6'), sources: sourcesOf('S6') });
    const ctr = text => ({ text, align: 'center' });
    const t = s.table([
      ['薬', '教科書の目安（mg/kg）', 'アドレナリン添加'],
      ['リドカイン', ctr('3〜4.5'), ctr('7')],
      ['メピバカイン', ctr('4.5〜5'), ctr('7')],
      ['ブピバカイン', ctr('2〜2.5'), ctr('3')],
      ['ロピバカイン', ctr('3'), ctr('—')],
      ['レボブピバカイン', ctr('2'), ctr('—')],
    ], { x: X0, y: TOP, colW: [3.6, 4.6, BW - 3.6 - 4.6] });
    const y = r2(t.y + t.h + 0.34);
    s.box({
      x: X0, y, w: BW, h: 0.82, fill: C.white, line: { color: C.teal, width: 2.5 }, radius: 0.08, size: 28,
      text: [['添付文書の 1 回量と、教科書の目安の', em('『低い方』', C.vermilion), 'を守る']],
    });
    if (y + 0.82 > BOT2 + 1e-6) throw new Error(`S6 一文の下端 ${y + 0.82} > ${BOT2}`);
  }

  // ===== S7 mg/kg → mL の計算 =====
  {
    const s = contentSlide(deck, { title: 'mg/kg → mL の計算', chapter: 3, notes: notesOf('S7'), sources: sourcesOf('S7') });
    // 右: 例のカード 2 枚（上下）。左右の内側余白は 0.2in（S3・S9 のカードと同程度）、上下は 0.15in
    const cw = 5.78, cx = XR - cw, cgap = 0.26;
    const card = { titleSize: 24, size: 22, pad: 0.15, padX: 0.2, titleGap: 0.07, itemGap: 0.03, titleColor: C.blue };
    const ex1 = {
      ...card, title: '60 kg、0.25% ブピバカイン',
      items: ['上限 2 mg/kg × 60 = 120 mg', '0.25% = 2.5 mg/mL', { lines: [['120 ÷ 2.5 = ', em('48 mL', C.vermilion)]] }],
    };
    const ex2 = {
      ...card, title: '50 kg、1% リドカイン',
      items: [
        '体重の目安 3 mg/kg × 50 = 150 mg',
        { lines: [['添付文書 200 mg → ', em('低い方 150 mg', C.dark)]] },
        { lines: [['1% = 10 mg/mL → ', em('15 mL', C.vermilion)]] },
      ],
    };
    const ch = Math.max(listCardH(ex1), listCardH(ex2));
    if (TOP + 2 * ch + cgap > BOT1 + 1e-6) throw new Error(`S7 カードの下端 ${r2(TOP + 2 * ch + cgap)} > ${BOT1}`);
    listCard(s, { ...ex1, x: cx, y: TOP, w: cw, h: ch });
    listCard(s, { ...ex2, x: cx, y: r2(TOP + ch + cgap), w: cw, h: ch });
    // 左: 換算の式（2 段。上の式で % を mg/mL にしてから下の式へ）。26pt で箱の幅に余裕を残す
    const FS = 26;
    const lw = cx - 0.4 - X0;
    const yc = TOP + (2 * ch + cgap) / 2;
    const fh = heightFor(1, FS, 0);
    s.text([['X % = ', em('X × 10 mg/mL', C.blue)]], { x: X0, y: r2(yc - 1.25 - fh / 2), w: lw, size: FS, bold: true, align: 'center' });
    flowArrow(s, X0 + lw / 2, r2(yc - 1.25 + fh / 2 + 0.26), X0 + lw / 2, r2(yc + 1.25 - fh / 2 - 0.26));
    s.text('mL = 上限(mg) ÷ 濃度(mg/mL)', { x: X0, y: r2(yc + 1.25 - fh / 2), w: lw, size: FS, bold: true, align: 'center' });
  }

  // ===== S8 用量以下でも中毒は起こる =====
  {
    const s = contentSlide(deck, { title: '用量以下でも中毒は起こる', chapter: 4, notes: notesOf('S8'), sources: sourcesOf('S8') });
    s.label('吸収が多い順', { x: X0, y: 1.98, w: 2.4, h: 0.46, align: 'left', color: C.source, bold: true });
    // 横フロー（多い → 少ない）。長い語は語の切れ目で 2 行に
    const chips = [
      { t: '血管内', fill: C.red, color: C.white },
      { t: '気管\n（表面麻酔）' },
      { t: '肋間' },
      { t: '仙骨\n硬膜外' },
      { t: '腰部\n硬膜外' },
      { t: '腕神経叢' },
      { t: '皮下' },
    ];
    const SZ = 22, PADX = 0.06, CHIP_H = 0.95, CY = 2.7;
    const ws = chips.map(c => r2(Math.max(...c.t.split('\n').map(l => fitInfo(l, SZ, 100).needIn)) + 2 * PADX + 0.16));
    const g = (BW - ws.reduce((a, b) => a + b, 0)) / (chips.length - 1);
    if (g < 0.36) throw new Error(`S8 フローの間隔が狭すぎます (${g.toFixed(2)}in)`);
    let x = X0;
    chips.forEach((c, i) => {
      s.box({
        x: r2(x), y: CY, w: ws[i], h: CHIP_H, fill: c.fill || C.bg, radius: 0.08,
        text: c.t, size: SZ, color: c.color || C.dark, bold: !!c.fill,
      });
      if (i < chips.length - 1) flowArrow(s, r2(x + ws[i] + 0.07), CY + CHIP_H / 2, r2(x + ws[i] + g - 0.07), CY + CHIP_H / 2, { hl: 0.16, hw: 0.22, width: 2.5 });
      x += ws[i] + g;
    });
    s.bullets([
      '同じ量でも、血管の多い部位ほど血中濃度が上がる',
      ['血管内なら注入中〜直後に痙攣。量が多ければ', em('心停止', C.red)],
      '年齢の両極端・小柄、心・肝疾患、妊娠、アシドーシスで出やすい',
    ], { y: 4.3, size: 24, gap: 20 });
  }

  // ===== S9 アドレナリン添加 =====
  {
    const s = contentSlide(deck, { title: 'アドレナリン添加', chapter: 4, notes: notesOf('S9'), sources: sourcesOf('S9') });
    const gap = 0.4, cw = (BW - gap) / 2;
    const common = { titleSize: 28, size: 22, pad: 0.18, titleGap: 0.14, itemGap: 0.16, marker: true, indent: 0.36 };
    const L = {
      ...common, title: '利点', titleColor: C.blue, items: [
        { lines: ['吸収が遅くなり、', '血中濃度のピークが下がる'] },
        '作用時間が延びる',
        '出血が減る',
        { lines: ['血管内誤注入のテスト', '（心拍数の上昇）'] },
      ],
    };
    const R = {
      ...common, title: '注意', titleColor: C.vermilion, items: [
        '1:200,000（5 μg/mL）で混ぜる',
        '市販品は 1:100,000 と濃い',
        { lines: ['陰茎は添付文書で禁忌'], color: C.red, bold: true },
        '耳・指趾は慎重に',
        '不整脈、高血圧、虚血性心疾患',
      ],
    };
    const h = Math.max(listCardH(L), listCardH(R));
    const y = r2((TOP + BOT2 - h) / 2);
    listCard(s, { ...L, x: X0, y, w: cw, h });
    listCard(s, { ...R, x: X0 + cw + gap, y, w: cw, h });
  }

  // ===== S10 まとめ =====
  {
    const s = contentSlide(deck, { title: 'まとめ', chapter: 4, notes: notesOf('S10'), sources: sourcesOf('S10') });
    const items = [
      'Na チャネルを内側から塞ぐ。炎症の組織では効きにくい',
      'リドカインは速く短い、ブピバカイン系は遅く長い',
      ['添付文書の 1 回量と mg/kg の目安の', em('低い方', C.vermilion), 'を守る'],
      ['針を持つ前に、', em('上限を mL で計算する', C.vermilion)],
      ['血管内に入れば', em('用量以下でも中毒', C.red), '。吸引・分割・会話'],
    ];
    // 段落後 26pt で項目の間を少し広げ、枠を本文ゾーン全体（TOP〜BOT1）にして上下中央揃えにする
    // （行の高さが書体で変わっても、文字の塊が上に寄らない）
    const size = 28, gap = 26;
    const hh = heightFor(items.length, size, gap);
    if (hh > BOT1 - TOP + 1e-6) throw new Error(`S10 箇条書きの高さ ${hh} > ${r2(BOT1 - TOP)}`);
    s.bullets(items, { y: TOP, h: r2(BOT1 - TOP), size, gap, valign: 'middle' });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
