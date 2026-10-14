'use strict';
/**
 * 3-6 筋弛緩の拮抗と残存筋弛緩 — 組版スクリプト
 *
 *   cd slides/3-6_reversal_residual_block
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *     書き写しの検査: 各スライドに置いた文字列を順につなげたものが、spec の on-slide の「」の中身を
 *     順につなげたものと一字一句同じでなければ Error（checkOnSlide）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 *     3 件が 2 行に入らないもの（S2・S3・S8）は、1 件を語の切れ目で 2 つに分けて渡す（文言は同じ）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, LAYOUT, measure, heightFor } = require('../_lib/decklib');

const NAME = '3-6_reversal_residual_block';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.RENDER_DIR || process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);
const RENDER = !process.argv.includes('--no-render');

// ---------------------------------------------------------------------------
// spec.md を読む（タイトル・on-slide の「」・ノート・出典）
// ---------------------------------------------------------------------------
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const lines = part.split('\n');
    const title = (lines[0].match(/title=「([^」]*)」/) || [])[1] || null;
    let notes = null, sources = null, subtitle = null, meta = null;
    const quoted = [];
    for (let i = 1; i < lines.length; i++) {
      const ln = lines[i];
      if (/^on-slide:\s*$/.test(ln)) {
        for (let j = i + 1; j < lines.length && /^ {2}- /.test(lines[j]); j++) {
          for (const m of lines[j].matchAll(/「([^」]*)」/g)) quoted.push(m[1]);
        }
      }
      if (/^notes:\s*\|\s*$/.test(ln)) {
        const buf = [];
        for (let j = i + 1; j < lines.length && (lines[j].startsWith('  ') || lines[j] === ''); j++) {
          buf.push(lines[j].replace(/^ {2}/, ''));
        }
        while (buf.length && buf[buf.length - 1] === '') buf.pop();
        notes = buf.join('\n');
      }
      let m;
      if ((m = ln.match(/^sources:\s*(.*)$/))) sources = m[1].trim();
      if ((m = ln.match(/^subtitle:\s*(.*)$/))) subtitle = m[1].trim();
      if ((m = ln.match(/^meta:\s*(.*)$/))) meta = m[1].trim();
    }
    out[id] = { title, quoted, notes, subtitle, meta, sources: sources && sources !== '—' ? sources : null };
  }
  return out;
}
const SPEC = parseSpec(fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8'));
function spec(id) {
  if (!SPEC[id]) throw new Error(`spec.md に ${id} がありません`);
  if (!SPEC[id].notes) throw new Error(`spec.md の ${id} に notes がありません`);
  return SPEC[id];
}

/**
 * 出典の項目（spec の sources 行を ' / ' で分けたもの）。
 * split: {項目の先頭一致文字列: 2行目の書き出し} — 2行に収めるために、その語の直前で項目を2つに分ける。
 * 分けた2つは文字列としてつなげると元の項目と同じ（つなぎ目の半角スペースだけ落ちる）。
 */
function sourcesOf(id, split = {}) {
  if (!SPEC[id].sources) return undefined;
  const items = [];
  for (const it of SPEC[id].sources.split(' / ')) {
    const key = Object.keys(split).find(k => it.startsWith(k));
    if (!key) { items.push(it); continue; }
    const k = it.indexOf(split[key]);
    if (k <= 0) throw new Error(`${id}: 出典「${it}」に分割位置「${split[key]}」がありません`);
    items.push(it.slice(0, k).replace(/ $/, ''), it.slice(k));
  }
  return items;
}

// ---------------------------------------------------------------------------
// 書き写しの検査: スライドに置いた文字列（置いた順）を記録し、spec の「」と突き合わせる
// ---------------------------------------------------------------------------
const PLACED = {};
const textOf = line => (Array.isArray(line) ? line.map(r => (typeof r === 'string' ? r : r.text)).join('') : (typeof line === 'object' && line ? String(line.text) : String(line)));
function rec(id, line) { (PLACED[id] = PLACED[id] || []).push(textOf(line)); }
function checkOnSlide() {
  for (const [id, sp] of Object.entries(SPEC)) {
    if (!sp.quoted.length) continue;
    // 改行位置の空白は落ちるので、つなぎ目の比較は空白を除いて行い、
    // 各行が spec の文字列の一部であること（行内の空白・文字が同じこと）は別に確かめる
    const placed = PLACED[id] || [];
    const whole = sp.quoted.join('\n');
    placed.forEach(l => {
      if (!whole.includes(l)) throw new Error(`${id}: 「${l}」が spec の on-slide にありません`);
    });
    const want = sp.quoted.join('').replace(/\s+/g, '');
    const got = placed.join('').replace(/\s+/g, '');
    if (want !== got) {
      let k = 0;
      while (k < want.length && want[k] === got[k]) k++;
      throw new Error(`${id}: スライド面の文言が spec の on-slide と違います（${k} 字目から）\n` +
        `  spec: …${want.slice(Math.max(0, k - 10), k + 30)}\n  got : …${got.slice(Math.max(0, k - 10), k + 30)}`);
    }
  }
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;          // 0.55
const BW = LAYOUT.body.w;          // 12.23
const RIGHT = LAYOUT.body.right;   // 12.78
const TOP = 1.85;                  // 本文の上端
const GAP = 0.24;                  // 要素間の最小間隔
const SRC2_TOP = LAYOUT.source.y2; // 出典2行のときの出典枠の上端 6.5
const BOTTOM = SRC2_TOP - GAP;     // 出典2行のスライドの本文下端 6.26
const ARROW = '404040';
const b = t => ({ text: t, bold: true });                         // 数値の強調（太字）
const em = (t, color) => ({ text: t, bold: true, color });        // 色付きの強調
const drug = t => ({ text: t, bold: true });                       // 行頭の薬剤名（色を付けると行頭記号の色も変わるので太字だけ）

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す） */
function contentSlide(deck, id, chapter, split) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id, split) });
  if (s.hasSources && s.sourceTop < SRC2_TOP - 1e-6) throw new Error(`${id}: 出典が3行以上`);
  return s;
}

/**
 * 見出し＋項目のカード。項目が2行以上になるときは、2行目以降を字下げする（ぶら下げ）。
 * 項目間は gap（既定は高さに応じて minGap〜maxGap）。各行は文字列か run 配列。
 * items: [[1行目, 2行目...], ...]
 */
function itemCardNeedW(o) {
  const ts = o.titleSize, bs = o.bodySize;
  let need = measure(textOf(o.title), ts);
  for (const lines of o.items) {
    lines.forEach((l, k) => { need = Math.max(need, measure(textOf(l), bs) + (k ? o.indent : 0)); });
  }
  return need + 2 * o.pad + 0.001;
}
function itemCard(s, id, o) {
  const { x, y, w, h, pad, padY, titleSize: ts, bodySize: bs, titleGap: tg, indent } = o;
  s.box({ x, y, w, h, fill: o.fill || C.bg, line: o.line, radius: 0.08 });
  const iw = w - 2 * pad;
  const th = heightFor(1, ts, 0);
  s.text(o.title, { x: x + pad, y: y + padY, w: iw, h: th, size: ts, bold: true, color: o.titleColor });
  rec(id, o.title);
  const top = y + padY + th + tg;
  const avail = y + h - padY - top;
  const lh = heightFor(1, bs, 0);
  const itemH = o.items.map(lines => lh + (lines.length > 1 ? heightFor(lines.length - 1, bs, 0) : 0));
  const sum = itemH.reduce((a, v) => a + v, 0);
  const n = o.items.length;
  let g = o.gap != null ? o.gap : (n > 1 ? Math.min(o.maxGap, (avail - sum) / (n - 1)) : 0);
  if (sum + g * (n - 1) > avail + 1e-6 || (o.gap == null && n > 1 && g < o.minGap - 1e-6)) {
    throw new Error(`${id} カード「${textOf(o.title)}」: 高さ不足（項目 ${sum.toFixed(2)}in + 間隔 > ${avail.toFixed(2)}in）`);
  }
  let cy = top;
  o.items.forEach((lines, k) => {
    s.text([lines[0]], { x: x + pad, y: cy, w: iw, h: lh, size: bs, color: C.dark });
    rec(id, lines[0]);
    if (lines.length > 1) {
      const rest = lines.slice(1);
      s.text(rest, { x: x + pad + indent, y: cy + lh, w: iw - indent, h: heightFor(rest.length, bs, 0), size: bs, color: C.dark });
      rest.forEach(l => rec(id, l));
    }
    cy += itemH[k] + (k < n - 1 ? g : 0);
  });
  return { gap: g, used: cy - y + padY };
}

/** カードの共通寸法（見出し 24pt・本文 22pt。項目の2行目以降は 0.3in 字下げ） */
const CARD = { pad: 0.14, padY: 0.16, titleSize: 24, bodySize: 22, titleGap: 0.16, indent: 0.3, minGap: 0.08, maxGap: 0.34 };

/**
 * 横並びのカード。列幅は中身の最長行から決め、余りは狭い列から順に配る（水位合わせ）。
 * cards: [{title, titleColor, items}], y=TOP〜bottom
 */
function cardRow(s, id, cards, bottom, geo = {}) {
  const n = cards.length;
  const total = BW - (n - 1) * GAP;
  const need = cards.map(c => itemCardNeedW({ ...CARD, ...geo, ...c }));
  if (need.reduce((a, v) => a + v, 0) > total + 1e-9) {
    throw new Error(`${id}: カードの必要幅 ${need.map(v => v.toFixed(2)).join(' + ')} > ${total.toFixed(2)}in`);
  }
  const fixed = new Set();
  let W = 0;
  for (;;) {
    const free = need.map((v, i) => i).filter(i => !fixed.has(i));
    const rest = total - [...fixed].reduce((a, i) => a + need[i], 0);
    W = rest / free.length;
    const over = free.filter(i => need[i] > W + 1e-9);
    if (!over.length) break;
    over.forEach(i => fixed.add(i));
  }
  const widths = need.map((v, i) => (fixed.has(i) ? v : W));
  let x = X0;
  const h = bottom - TOP;
  cards.forEach((c, i) => {
    itemCard(s, id, { ...CARD, ...geo, ...c, x, y: TOP, w: widths[i], h });
    x += widths[i] + GAP;
  });
  return widths;
}

/** 強調ボックス（白地・ティールの枠・太字）。lines は1行（文字列/run 配列）または行の配列（o.multi か run 配列の配列） */
function keyBox(s, id, lines, o) {
  const arr = (Array.isArray(lines) && lines.length && (Array.isArray(lines[0]) || o.multi)) ? lines : [lines];
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.white, line: { color: C.teal, width: 2.5 }, radius: 0.08,
    text: arr, size: o.size || 28, bold: true, color: C.dark, align: o.align || 'center' });
  arr.forEach(l => rec(id, l));
}

/** フローのチップ（灰色の角丸・枠・太字）。lines は行の配列 */
function chip(s, id, lines, o) {
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.bg, line: o.line || { color: C.blue, width: 2 }, radius: 0.1,
    text: lines, size: o.size, bold: true, color: C.dark, align: o.align || 'center', margin: o.margin });
  lines.forEach(l => rec(id, l));
}

/** 箇条書きの下端が出典に近すぎないか */
function checkBottom(id, what, bottom) {
  if (bottom > BOTTOM + 1e-6) throw new Error(`${id}: ${what}が出典に近すぎます（下端 ${bottom.toFixed(2)}in > ${BOTTOM.toFixed(2)}in）`);
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['残存\n筋弛緩', 'スガマ\nデクス', 'ネオスチ\nグミン', '再クラ\nーレ化', '手順'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '3-6 筋弛緩の拮抗と残存筋弛緩',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 残存筋弛緩とは（箇条書き4行＋下に強調ボックス） =====
  {
    const id = 'S2';
    // 出典: 3 件は 1 件ずつでは 2 行に入らないので、Suzuki の項を「J Anesth」の前で分ける
    const s = contentSlide(deck, id, 0, { 'Suzuki': 'J Anesth' });
    const items = [
      [b('TOF比 0.9 未満'), 'で手術室を出る状態'],
      '咽頭の筋が弱く、誤嚥・上気道閉塞・低酸素を起こす',
      '頭を 5 秒上げられても TOF比 0.9 とは限らない',
      ['拮抗しないと、挿管量 1 回でも回復室で ', em('45%', C.vermilion), ' に残る'],
    ];
    items.forEach(t => rec(id, t));
    const by = 1.92, gap = 16;
    s.bullets(items, { y: by, size: 24, gap });
    const bh = heightFor(items.length, 24, gap);
    const kh = 1.26;
    const ky = BOTTOM - kh;
    if (ky - (by + bh) < 0.3 - 1e-6) throw new Error(`S2: 箇条書きと強調ボックスの間が狭すぎます（${(ky - by - bh).toFixed(2)}in）`);
    keyBox(s, id, [
      ['定量モニターで ', em('TOF比 0.9 以上', C.vermilion), 'を確認して抜管'],
      [{ text: '（JSA 2025 推奨）', size: 24, bold: false }],
    ], { x: X0, y: ky, w: BW, h: kh, size: 28 });
  }

  // ===== S3 なぜ 0.9 なのか（左に縦の帯の目盛り、右に各段の説明） =====
  {
    const id = 'S3';
    // 出典: Kopman の項を「Anesthesiology」の前で分ける
    const s = contentSlide(deck, id, 0, { 'Kopman': 'Anesthesiology' });
    // 段は上から 0.4 → 1.0（spec の並び＝回復の進む順）
    const rows = [
      { label: 'TOF比 0.4', desc: '目や手で減衰がわかる限界', fill: C.vermilion },
      { label: '0.7', desc: '頭を上げられる。しかし嚥下は障害', fill: C.vermilion },
      { label: '0.9', desc: '咽頭機能がほぼ回復。抜管の基準', fill: C.teal, key: true },
      { label: '1.0', desc: '完全回復', fill: C.blue },
    ];
    const segH = (BOTTOM - TOP) / rows.length;
    const labW = 2.3, bandX = X0 + labW + 0.3, bandW = 0.8;
    const cardX = bandX + bandW + 0.36, cardW = RIGHT - cardX;
    const lsz = 28, lh = heightFor(1, lsz, 0);
    // 目盛りのラベル（帯の左の外側）
    rows.forEach((r, i) => {
      const y = TOP + i * segH;
      s.text([r.label], { x: X0, y: y + (segH - lh) / 2, w: labW, h: lh, size: lsz, bold: true,
        color: r.key ? C.teal : C.dark, align: 'right' });
      rec(id, r.label);
    });
    // 縦の帯（4 段。段の境は白線）
    rows.forEach((r, i) => {
      s.box({ x: bandX, y: TOP + i * segH, w: bandW, h: segH, fill: r.fill, line: { color: C.white, width: 2 } });
    });
    // 各段の説明（帯の右）。0.9 の段はティールの枠で強調
    const ch = segH - GAP;
    rows.forEach((r, i) => {
      const y = TOP + i * segH + GAP / 2;
      s.box({ x: cardX, y, w: cardW, h: ch, radius: 0.08,
        fill: r.key ? C.white : C.bg, line: r.key ? { color: C.teal, width: 3 } : undefined,
        text: [r.desc], size: 28, bold: !!r.key, color: C.dark, align: 'left', margin: 0.25 });
      rec(id, r.desc);
    });
  }

  // ===== S4 スガマデクス（3列カード） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 1);
    cardRow(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          ['T2 再出現後：', [b('2 mg/kg')]],
          ['PTC 1〜2 出現後：', [b('4 mg/kg')]],
          ['緊急時（ロクロニウム', ['3 分後）：', b('16 mg/kg')]],
        ],
      },
      {
        title: '仕組み', titleColor: C.blue,
        items: [
          ['ロクロニウムを', '1:1 で包み込む'],
          ['深い遮断からも', '数分で回復'],
          ['ベクロニウムにも有効'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['用量不足で', '再クラーレ化'],
          ['重度腎障害：', 'CLcr 30 mL/分未満'],
          ['アナフィラキシー、', '徐脈'],
        ],
      },
    ], BOTTOM);
  }

  // ===== S5 ネオスチグミン（3列カード） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 2);
    cardRow(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          [['1 回 ', b('0.5〜2.0 mg'), ' を緩徐に静注']],
          ['モニターの回復か', '自発呼吸を確認後'],
          [[b('5 mg'), ' を超えない']],
          [['アトロピン ', b('0.25〜1.0 mg'), ' を併用']],
        ],
      },
      {
        title: '仕組み', titleColor: C.blue,
        items: [
          ['コリンエステラーゼ', 'を阻害'],
          ['アセチルコリンを', '増やして競り勝つ'],
          ['効果に天井がある'],
        ],
      },
      {
        title: '限界', titleColor: C.vermilion,
        items: [
          ['深い遮断は', '拮抗できない'],
          ['最大効果まで', '7〜10 分'],
          ['徐脈・分泌・', '気管支攣縮'],
        ],
      },
    ], BOTTOM);
  }

  // ===== S6 どちらを使うか（表 4 列 × 4 行） =====
  {
    const id = 'S6';
    // 出典は 1 件ずつの折り返しで 2 行に入る
    const s = contentSlide(deck, id, 2);
    const blue = t => ({ text: t, bold: true, color: C.blue });
    const verm = t => ({ text: t, bold: true, color: C.vermilion });
    const rows = [
      ['深さ', 'モニター', 'スガマデクス', 'ネオスチグミン'],
      [{ text: '深い', bold: true }, 'PTC 1〜2', { text: [[blue('4 mg/kg')]] }, { text: [[verm('効かない')]] }],
      [{ text: '中等度', bold: true }, 'TOF 1〜3 回', { text: [[blue('2 mg/kg')], '（TOF 1 回は 4 mg/kg）'] },
        { text: [[verm('使わない')], '（遅く不確実）'] }],
      [{ text: '浅い', bold: true }, 'TOF 4 回、比 0.4 未満', { text: [[blue('2 mg/kg')]] }, { text: '不十分なことが多い', color: C.vermilion }],
      [{ text: '最小', bold: true }, 'TOF比 0.4〜0.9 未満', { text: [[blue('2 mg/kg')]] }, { text: [[blue('使える')], '（天井に注意）'] }],
    ];
    rows.forEach(r => r.forEach(c => {
      const t = (c && typeof c === 'object' && !Array.isArray(c)) ? c.text : c;
      (Array.isArray(t) ? t : String(t).split('\n')).forEach(l => rec(id, l));
    }));
    const rowH = [0.62, 0.78, 1.1, 0.78, 1.1];
    const tb = s.table(rows, { x: X0, y: TOP, colW: [1.27, 3.7, 3.94, 3.32], size: 22, rowH });
    checkBottom(id, '表', tb.y + tb.h);
  }

  // ===== S7 再クラーレ化（左右2枚のカード＋下に一文） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 3);
    const kh = 0.7, ky = BOTTOM - kh;
    const cardBottom = ky - 0.26;
    cardRow(s, id, [
      {
        title: '起こる状況', titleColor: C.vermilion,
        items: [
          [['深い遮断にスガマデクス ', b('2 mg/kg')]],
          ['ロクロニウムの総量が多い'],
          ['ネオスチグミンで深い遮断を無理に拮抗'],
          ['腎機能低下＋少なめの量'],
        ],
      },
      {
        title: '気づき方と対応', titleColor: C.blue,
        items: [
          ['抜管後〜病棟で呼吸が浅い、', 'SpO2 低下'],
          ['TOF を測り直す'],
          ['スガマデクスを追加', '（深さに応じて）'],
        ],
      },
    ], cardBottom, { padY: 0.12, titleGap: 0.1 });
    keyBox(s, id, ['いったん 0.9 まで戻っても、', em('薬の量が足りなければ戻る', C.vermilion)],
      { x: X0, y: ky, w: BW, h: kh, size: 28 });
  }

  // ===== S8 抜管前の手順（横フロー 5 段） =====
  {
    const id = 'S8';
    // 出典: Thilen の項を「Anesthesiology」の前で分ける
    const s = contentSlide(deck, id, 4, { 'Thilen': 'Anesthesiology' });
    const steps = [
      { lines: ['深さを測る', '（TOF・PTC）'], wt: 6 },
      { lines: ['深さで', '用量を', '決める'], wt: 1 },
      { lines: ['拮抗薬を', '投与'], wt: 1 },
      { lines: ['定量 TOF比', '0.9 以上を', '確認'], key: true, wt: 3 },
      { lines: ['意識・呼吸と', '合わせて抜管'], wt: 2 },
    ];
    const size = 24, m = 0.04, ag = 0.5, ins = 0.1;
    const need = steps.map(st => Math.max(...st.lines.map(l => measure(textOf(l), size))) + 2 * m + 0.001);
    const total = BW - (steps.length - 1) * ag;
    const slack = total - need.reduce((a, v) => a + v, 0);
    // 英大文字（TOF・PTC）と和欧間の空き（PowerPoint 既定で入る）は推定より幅を取るので、
    // 余りは英字を含むチップに厚く配る（wt の比）
    const wsum = steps.reduce((a, st) => a + st.wt, 0);
    const extra = steps.map(st => slack * st.wt / wsum);
    if (slack < 0.5) throw new Error(`S8: チップの余裕が足りません（${need.map(v => v.toFixed(2)).join(' + ')}、余り ${slack.toFixed(2)}in）`);
    const widths = need.map((v, i) => v + extra[i]);
    const chh = 2.2;
    const y = TOP + (BOTTOM - TOP - chh) / 2;
    let x = X0;
    steps.forEach((st, i) => {
      chip(s, id, st.lines, { x, y, w: widths[i], h: chh, size, margin: m,
        line: st.key ? { color: C.teal, width: 3.5 } : { color: C.blue, width: 2 } });
      if (i < steps.length - 1) {
        const yc = y + chh / 2;
        s.arrow({ x1: x + widths[i] + ins, y1: yc, x2: x + widths[i] + ag - ins, y2: yc, color: ARROW, width: 3 });
      }
      x += widths[i] + ag;
    });
  }

  // ===== S9 残りやすい患者（箇条書き5行） =====
  {
    const id = 'S9';
    const s = contentSlide(deck, id, 4);
    const items = [
      '高齢者、肝・腎機能障害',
      '低体温（回復が遅い。冷えた手では反応が小さく出る）',
      '揮発性麻酔薬、マグネシウム、アミノグリコシド',
      '追加投与が多かった、持続投与していた',
      '神経筋疾患（重症筋無力症など）',
    ];
    items.forEach(t => rec(id, t));
    const gap = 22, by = 1.92, size = 28;
    s.bullets(items, { y: by, size, gap });
    checkBottom(id, '箇条書き', by + heightFor(items.length, size, gap));
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const id = 'S10';
    const s = contentSlide(deck, id, 4);
    const items = [
      ['残存筋弛緩＝', b('TOF比 0.9 未満'), '。誤嚥・閉塞・低酸素の原因'],
      [drug('スガマデクス'), '：T2 で 2 mg/kg、PTC 1〜2 で 4 mg/kg'],
      [drug('ネオスチグミン'), '：TOF比 0.4 以上だけ。アトロピン併用。5 mg まで'],
      '用量不足は再クラーレ化。深さを測ってから量を決める',
      '定量的モニターで 0.9 以上を確認してから抜管する',
    ];
    items.forEach(t => rec(id, t));
    const gap = 28, by = 1.95, size = 24;
    s.bullets(items, { y: by, size, gap });
    checkBottom(id, '箇条書き', by + heightFor(items.length, size, gap));
  }

  checkOnSlide();
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.render && res.render.images) console.log('[build] render: ' + (res.render.images[0] ? path.dirname(res.render.images[0]) : '(none)'));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
