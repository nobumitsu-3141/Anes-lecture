'use strict';
/**
 * 3-5 筋弛緩薬 — 組版スクリプト
 *
 *   cd slides/3-5_neuromuscular_blockers
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *     書き写しの検査: 各スライドに置いた文字列を順につなげたものが、spec の on-slide の「」の中身を
 *     順につなげたものと一字一句同じでなければ Error（checkOnSlide）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 *     2行に収まらないものは、項目の途中（引用の区切り）で改行させる。それでも収まらない S8・S9 だけ
 *     「日本麻酔科学会」→「JSA」などの最小限の略記にする（build の出力 deviations に記載）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, LAYOUT, LINT, measure, fits, heightFor, estHeightIn, realHeightIn } = require('../_lib/decklib');

const NAME = '3-5_neuromuscular_blockers';
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

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を ' / ' で項目に分ける。
//   split:   {項目の先頭: 分割位置の文字列} — 項目の途中で改行させる（文言は変えない）
//   replace: [[元, 後], ...] — 2行に収まらないときだけの略記（deviations に記載）
// 分割した項目の前後が同じ行に詰められる（＝余計な ' / ' が入る）と Error。
// ---------------------------------------------------------------------------
const SRC = LAYOUT.source;
function sourcesOf(id, split = {}, replace = null) {
  if (!SPEC[id].sources) return undefined;
  let src = SPEC[id].sources;
  if (replace) {
    for (const [from, to] of replace) {
      if (!src.includes(from)) throw new Error(`${id}: 出典に「${from}」がありません`);
      src = src.replace(from, to);
    }
  }
  const items = [];
  const joints = new Set();   // items[k] と items[k+1] が1項目の分割（同じ行に入ってはいけない）
  for (const it of src.split(' / ')) {
    const key = Object.keys(split).find(k => it.startsWith(k));
    if (!key) { items.push(it); continue; }
    const k = it.indexOf(split[key]);
    if (k <= 0) throw new Error(`${id}: 出典「${it}」に分割位置「${split[key]}」がありません`);
    joints.add(items.length);
    items.push(it.slice(0, k).replace(/ $/, ''), it.slice(k));
  }
  // ライブラリと同じ詰め方で行を作り、分割点が行の切れ目に来ることを確かめる
  const lineOf = [];
  let cur = null, ln = -1;
  items.forEach(it => {
    const cand = cur != null ? cur + ' / ' + it : null;
    if (cand && fits(cand, SRC.size, SRC.w)) cur = cand;
    else { cur = it; ln++; }
    lineOf.push(ln);
  });
  for (const k of joints) {
    if (lineOf[k] === lineOf[k + 1]) throw new Error(`${id}: 出典の分割点「${items[k + 1].slice(0, 10)}」が行の途中に来ます`);
  }
  if (ln + 1 > 2) throw new Error(`${id}: 出典が ${ln + 1} 行になります`);
  return items;
}

// ---------------------------------------------------------------------------
// 書き写しの検査: スライドに置いた文字列（置いた順）を記録し、spec の「」と突き合わせる
// ---------------------------------------------------------------------------
const PLACED = {};
const textOf = line => {
  if (Array.isArray(line)) return line.map(r => (typeof r === 'string' ? r : r.text)).join('');
  if (line && typeof line === 'object') return String(line.text);
  return String(line);
};
function rec(id, line) { (PLACED[id] = PLACED[id] || []).push(textOf(line)); }
function checkOnSlide() {
  for (const [id, sp] of Object.entries(SPEC)) {
    if (!sp.quoted.length) continue;
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
const TOP = 1.85;                  // 本文の上端
const GAP = 0.24;                  // 要素間の最小間隔
const BOTTOM1 = 6.6;               // 出典1行のときの本文の下端（出典 y6.85）
const BOTTOM2 = SRC.y2 - GAP;      // 出典2行のときの本文の下端（出典 y6.5 → 6.26）
const ARROW_C = '404040';
const b = t => ({ text: t, bold: true });                         // 数値の強調（太字）
const em = (t, color) => ({ text: t, bold: true, color });         // 色の強調
const danger = t => ({ text: t, bold: true, color: C.red });       // 危険（レッド）

/** 混在サイズの行の高さ（lint 式と実機推定の大きい方。decklib の heightFor と同じ考え方） */
function hNeed(sizes) {
  const paras = sizes.map(size => ({ text: 'あ', size, spaceAfter: 0 }));
  const lintNeed = estHeightIn(paras, 100) - LINT.overflowTolIn;
  return Math.ceil(Math.max(lintNeed, realHeightIn(paras)) * 100) / 100;
}

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す） */
function contentSlide(deck, id, chapter, split, replace) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id, split, replace) });
  if (s.hasSources && s.sourceTop < SRC.y2 - 1e-6) throw new Error(`${id}: 出典が3行以上`);
  return s;
}

/**
 * 見出し＋項目のカード。項目が2行以上になるときは、2行目以降を字下げする（ぶら下げ）。
 * title: 文字列 | 行の配列（各行は文字列 or {text, size}）
 * items: [[1行目, 2行目...], ...]（各行は文字列 or run 配列 or run）
 * 項目間は gap 固定、または minGap〜maxGap で高さに合わせて配る。
 */
function titleLinesOf(o) {
  const tl = Array.isArray(o.title) ? o.title : [o.title];
  return { lines: tl, sizes: tl.map(l => (l && typeof l === 'object' && !Array.isArray(l) && l.size) || o.titleSize) };
}
function itemCardNeedW(o) {
  const { lines, sizes } = titleLinesOf(o);
  let need = Math.max(...lines.map((l, i) => measure(textOf(l), sizes[i])));
  for (const ls of o.items) {
    ls.forEach((l, k) => { need = Math.max(need, measure(textOf(l), o.bodySize) + (k ? o.indent : 0)); });
  }
  return need + 2 * o.pad + 0.001;
}
function itemCardNeedH(o, gap) {
  const { sizes } = titleLinesOf(o);
  const h1 = heightFor(1, o.bodySize, 0);
  const itemH = o.items.map(ls => h1 * ls.length);
  return 2 * o.padY + hNeed(sizes) + o.titleGap + itemH.reduce((a, v) => a + v, 0) + gap * (o.items.length - 1);
}
function itemCard(s, id, o) {
  const { x, y, w, h, pad, padY, bodySize: bs, titleGap: tg, indent } = o;
  s.box({ x, y, w, h, fill: o.fill || C.bg, line: o.line, radius: 0.08 });
  const iw = w - 2 * pad;
  const { lines: tl, sizes } = titleLinesOf(o);
  const th = hNeed(sizes);
  s.text(tl, { x: x + pad, y: y + padY, w: iw, h: th, size: o.titleSize, bold: true, color: o.titleColor });
  tl.forEach(l => rec(id, l));
  const top = y + padY + th + tg;
  const avail = y + h - padY - top;
  const h1 = heightFor(1, bs, 0);
  const itemH = o.items.map(ls => h1 * ls.length);
  const sum = itemH.reduce((a, v) => a + v, 0);
  const n = o.items.length;
  const g = o.gap != null ? o.gap : (n > 1 ? Math.min(o.maxGap, (avail - sum) / (n - 1)) : 0);
  if (sum + g * (n - 1) > avail + 1e-6 || (o.gap == null && n > 1 && g < o.minGap - 1e-6)) {
    throw new Error(`${id} カード「${textOf(tl[0])}」: 高さ不足（項目 ${sum.toFixed(2)}in + 間隔 ${g.toFixed(2)}in×${n - 1} > ${avail.toFixed(2)}in）`);
  }
  // 行ごとに別の枠に置く（複数行の枠は描画系で行送りが変わるので、行送りを h1 にそろえる）
  let cy = top;
  o.items.forEach((ls, k) => {
    ls.forEach((l, j) => {
      const ind = j ? indent : 0;
      s.text([l], { x: x + pad + ind, y: cy + j * h1, w: iw - ind, h: h1, size: bs, color: C.dark });
      rec(id, l);
    });
    cy += itemH[k] + (k < n - 1 ? g : 0);
  });
  return { gap: g, bottom: cy };
}

/** 左右に並べるカードの項目間隔を共通にする（行の高さをそろえる）。各カードで取れる間隔の最小値 */
function commonGap(cards, geo) {
  const gs = cards.map(c => {
    const o = { ...geo, ...c };
    const n = o.items.length;
    if (n < 2) return o.maxGap;
    const free = o.h - itemCardNeedH(o, 0);
    return Math.min(o.maxGap, free / (n - 1));
  });
  const g = Math.min(...gs);
  if (g < geo.minGap - 1e-6) throw new Error(`カードの項目間隔が狭すぎます (${g.toFixed(2)}in)`);
  return g;
}

/** 3列カードの共通寸法（本文 22pt。項目の2行目は 0.3in 字下げ） */
const CARD3 = { pad: 0.12, padY: 0.16, titleSize: 24, bodySize: 22, titleGap: 0.14, indent: 0.3, minGap: 0.08, maxGap: 0.42 };

/**
 * 3列カード。列幅は中身の最長行から決め、余りは狭い列から順に配る（水位合わせ）。
 * 高さは top〜bottom いっぱい。cards: [{title, titleColor, items}]
 */
function threeCards(s, id, cards, top, bottom) {
  const total = BW - 2 * GAP;
  const need = cards.map(c => itemCardNeedW({ ...CARD3, ...c }));
  if (need.reduce((a, v) => a + v, 0) > total + 1e-9) {
    throw new Error(`${id}: 3列カードの必要幅 ${need.map(v => v.toFixed(2)).join(' + ')} > ${total.toFixed(2)}in`);
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
  const geo = { ...CARD3, y: top, h: bottom - top };
  const gap = commonGap(cards, geo);   // 3枚で行の間隔をそろえる
  let x = X0;
  cards.forEach((c, i) => {
    itemCard(s, id, { ...geo, ...c, x, w: widths[i], gap });
    x += widths[i] + GAP;
  });
}

/** 強調ボックス（白地・ティールの枠・太字）。line は1行（文字列 or run 配列） */
function keyBox(s, id, line, o) {
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.white, line: { color: C.teal, width: 2.5 }, radius: 0.08,
    text: [line], size: o.size || 28, bold: true, color: C.dark, align: o.align || 'center' });
  rec(id, line);
}

/** フローのチップ（灰色の角丸・色枠・太字）。lines は行の配列 */
function chip(s, id, lines, o) {
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.bg, line: { color: o.lineColor || C.blue, width: 2.25 }, radius: 0.1,
    text: lines, size: o.size, bold: true, color: C.dark, align: 'center' });
  lines.forEach(l => rec(id, l));
}

/**
 * 水平・垂直の矢印（軸＋塗りの三角の矢じり）。pptxgenjs の矢じりは線幅に比例して小さいので、
 * 矢じりを自前の三角形（freeform）で描く。hl: 矢じりの長さ、hw: 矢じりの幅(in)
 */
function flowArrow(s, x1, y1, x2, y2, o = {}) {
  const color = o.color || ARROW_C, width = o.width || 3, hl = o.hl || 0.24, hw = o.hw || 0.3;
  const L = Math.hypot(x2 - x1, y2 - y1), ux = (x2 - x1) / L, uy = (y2 - y1) / L;
  if (Math.abs(ux) > 1e-9 && Math.abs(uy) > 1e-9) throw new Error('flowArrow は水平・垂直のみ');
  const bx = x2 - ux * hl, by = y2 - uy * hl;
  if (uy === 0) s.hline({ x: Math.min(x1, bx), y: y1, w: Math.abs(bx - x1) + 0.02, color, width });
  else s.vline({ x: x1, y: Math.min(y1, by), h: Math.abs(by - y1) + 0.02, color, width });
  const px = -uy * hw / 2, py = ux * hw / 2;
  s.polyline([[x2, y2], [bx + px, by + py], [bx - px, by - py]], { color, width: 1, closed: true, fill: color });
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['仕組み', 'ロクロ\nニウム', 'スキサメ\nトニウム', '重い\n副作用', '使い方'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '3-5 筋弛緩薬',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 筋弛緩薬は何をする薬か（左右2枚の大きなカード＋下に強調ボックス） =====
  {
    const id = 'S2';
    const s = contentSlide(deck, id, 0);
    const cg = 0.4, cw = (BW - cg) / 2;
    const keyH = 0.82, keyGap = 0.3;
    const ch = BOTTOM2 - keyH - keyGap - TOP;
    const geo = { y: TOP, w: cw, h: ch, pad: 0.22, padY: 0.16, titleSize: 28, bodySize: 24, titleGap: 0.12, indent: 0.36, minGap: 0.06, maxGap: 0.3 };
    const cards = [
      { x: X0, title: 'できること', titleColor: C.blue,
        items: [['骨格筋を動かなくする'], ['挿管しやすくする'], ['術野を静かにする'], ['人工呼吸に同調させる']] },
      { x: X0 + cw + cg, title: 'できないこと', titleColor: C.vermilion,
        items: [['眠らせない（意識は残る）'], ['痛みを取らない'], [[danger('自分で呼吸できなくなる')]]] },
    ];
    const gap = commonGap(cards, geo);
    cards.forEach(c => itemCard(s, id, { ...geo, ...c, gap }));
    keyBox(s, id, ['筋弛緩薬を入れたら、', em('換気は自分の責任', C.vermilion)],
      { x: X0, y: TOP + ch + keyGap, w: BW, h: keyH });
  }

  // ===== S3 非脱分極性と脱分極性（上に一文＋左右2枚のカード） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 0);
    const lead = ['受容体を', em('『ふさぐ』', C.blue), 'か', em('『刺激し続ける』', C.vermilion), 'か'];
    const leadH = heightFor(1, 30, 0);
    s.text([lead], { x: X0, y: TOP, w: BW, h: leadH, size: 30, bold: true, align: 'center' });
    rec(id, lead);
    const cg = 0.34, cw = (BW - cg) / 2;
    const cy = TOP + leadH + 0.3;
    const geo = { y: cy, w: cw, h: BOTTOM1 - cy, pad: 0.15, padY: 0.16, titleSize: 28, bodySize: 24, titleGap: 0.12, indent: 0.36, minGap: 0.1, maxGap: 0.3 };
    const cards = [
      { x: X0, title: ['非脱分極性', { text: '（ロクロニウム）', size: 24 }], titleColor: C.blue,
        items: [['アセチルコリンと', '受容体を取り合う'], ['筋は静かに動かなくなる'], ['スガマデクスで拮抗できる']] },
      { x: X0 + cw + cg, title: ['脱分極性', { text: '（スキサメトニウム）', size: 24 }], titleColor: C.vermilion,
        items: [['受容体を刺激し続けて不応にする'], ['最初に筋線維束攣縮'], [[b('拮抗薬はない')]]] },
    ];
    const gap = commonGap(cards, geo);
    cards.forEach(c => itemCard(s, id, { ...geo, ...c, gap }));
  }

  // ===== S4 ロクロニウム（3列カード） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 1, { '日本麻酔科学会': 'Ⅵ（2025）' });
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          [['挿管：', b('0.6 mg/kg')], ['（上限 ', b('0.9 mg/kg'), '）']],
          [['追加：', b('0.1〜0.2 mg/kg')]],
          [['持続：', b('7 μg/kg/分'), 'で開始']],
        ],
      },
      {
        title: '時間', titleColor: C.blue,
        items: [
          [['効くまで ', b('1〜2 分')]],
          [['高齢者は', b('約 3 分')]],
          [['持続 ', b('30〜60 分')]],
          ['高齢・吸入麻酔で延長'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['注射時の痛み'],
          ['未変化体で胆汁・尿へ'],
          ['肝腎障害・高齢で遷延'],
          ['アナフィラキシー'],
        ],
      },
    ], TOP, BOTTOM2);
  }

  // ===== S5 迅速導入でのロクロニウム（横フロー3段＋下に注意文） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 1);
    const steps = [
      { lines: ['大量', ['（', b('0.9 mg/kg'), '、'], '海外は 1.2）'], color: C.blue },
      { lines: [[b('60〜90 秒'), 'で'], '挿管'], color: C.blue },
      { lines: ['失敗なら', 'スガマデクス', [em('16 mg/kg', C.vermilion)]], color: C.vermilion },
    ];
    const size = 28;
    const need = steps.map(st => Math.max(...st.lines.map(l => measure(textOf(l), size))) + 0.16);
    const cw = Math.max(...need) + 0.16;                // 3枚とも同じ幅
    const ag = (BW - 3 * cw) / 2;                        // 矢印の区間
    if (ag < 0.26 * 2 + 0.6) throw new Error(`S5: 矢印の区間が狭すぎます (${ag.toFixed(2)}in)`);
    const chh = 2.05;
    const keyH = 0.95, between = 0.6;
    const y1 = TOP + (BOTTOM2 - TOP - (chh + between + keyH)) / 2;
    steps.forEach((st, i) => {
      const x = X0 + i * (cw + ag);
      chip(s, id, st.lines, { x, y: y1, w: cw, h: chh, size, lineColor: st.color });
      if (i < steps.length - 1) {
        const yc = y1 + chh / 2, ins = 0.26;
        flowArrow(s, x + cw + ins, yc, x + cw + ag - ins, yc);
      }
    });
    keyBox(s, id, ['量を増やすほど速く効くが、', em('長く残る', C.vermilion)],
      { x: X0, y: y1 + chh + between, w: BW, h: keyH, size: 30 });
  }

  // ===== S6 スキサメトニウム（3列カード） =====
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 2, { '日本麻酔科学会': 'Ⅵ（2025）' });
    threeCards(s, id, [
      {
        title: '用量', titleColor: C.blue,
        items: [
          [['1 回 ', b('10〜60 mg'), ' 静注'], '（添付文書・成人）'],
          [['挿管の慣用量 ', b('1 mg/kg')]],
          [[b('拮抗薬はない')]],
        ],
      },
      {
        title: '時間', titleColor: C.blue,
        items: [
          [['効くまで', b('約 1 分')]],
          [[b('約 10 分'), 'で戻る']],
          ['血漿コリンエステラーゼで', '分解'],
        ],
      },
      {
        title: '副作用', titleColor: C.vermilion,
        items: [
          [[danger('高カリウム血症')]],
          [[danger('悪性高熱症'), 'の誘因']],
          ['徐脈、筋痛、', '眼圧・胃内圧上昇'],
          ['アナフィラキシー'],
        ],
      },
    ], TOP, BOTTOM2);
  }

  // ===== S7 高カリウム血症・悪性高熱（左右2枚のカード） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 3);
    const cg = 0.34, cw = (BW - cg) / 2;
    const geo = { y: TOP, w: cw, h: BOTTOM2 - TOP, pad: 0.15, padY: 0.16, titleSize: 28, bodySize: 24, titleGap: 0.14, indent: 0.36, minGap: 0.1, maxGap: 0.32 };
    const cards = [
      { x: X0, title: ['高カリウム血症', { text: '（スキサメトニウム）', size: 24 }], titleColor: C.vermilion,
        items: [
          ['脱分極で筋から K が出る'],
          [['通常は', b('約 0.5 mEq/L'), ' 上昇']],
          ['熱傷・広範外傷・脱神経・', '長期臥床・神経筋疾患では', ['急上昇し', danger('心停止')]],
        ] },
      { x: X0 + cw + cg, title: '悪性高熱症', titleColor: C.vermilion,
        items: [
          ['誘因：スキサメトニウムと', '揮発性麻酔薬'],
          [[b('EtCO2 上昇'), '、頻脈、'], '筋硬直、体温上昇'],
          [[b('ダントロレン'), 'の場所を知っておく']],
        ] },
    ];
    const gap = commonGap(cards, geo);
    cards.forEach(c => itemCard(s, id, { ...geo, ...c, gap }));
  }

  // ===== S8 アナフィラキシー（箇条書き4行＋強調ボックス） =====
  {
    const id = 'S8';
    // 出典は3件（全角換算 101 字）で 16pt・2行（46.6 字×2）にどこで割っても入らない。
    // 3件目を spec の notes にある略記「JSA「…」2021」にし、2件目は誌名の後で改行させる。
    const s = contentSlide(deck, id, 3, { 'Takazawa T': '2023;131' },
      [['日本麻酔科学会「アナフィラキシーに対する対応プラクティカルガイド」（2021年2月）',
        'JSA「アナフィラキシーに対する対応プラクティカルガイド」2021']]);
    const items = [
      '筋弛緩薬は周術期アナフィラキシーの主な原因の一つ',
      ['英国 NAP6：抗菌薬 ', b('47%'), '、筋弛緩薬 ', b('33%')],
      '国内（JESPA）：ロクロニウム、スガマデクス、セファゾリンが上位',
      '徴候：低血圧、頻脈、気管支攣縮、皮膚症状（ないこともある）',
    ];
    items.forEach(t => rec(id, t));
    const by = 1.98, gap = 18;
    s.bullets(items, { y: by, size: 24, gap });
    const bh = heightFor(items.length, 24, gap);
    const keyH = 0.95;
    const ky = by + bh + 0.4;
    keyBox(s, id, ['第一選択は', em('アドレナリン', C.vermilion), '。原因の推定は後でよい'],
      { x: X0, y: ky, w: BW, h: keyH });
    if (ky + keyH > BOTTOM2 + 1e-6) throw new Error('S8: 強調ボックスが出典に近すぎます');
  }

  // ===== S9 使い方の原則（箇条書き5行） =====
  {
    const id = 'S9';
    // 出典は 16pt・2行に入らない（Suzuki の項の途中で割っても 0.05 字超える）。
    // 2-6 と同じく「Suzuki T, et al.」の後で改行し、3件目の「日本麻酔科学会」を「JSA」に略す。
    const s = contentSlide(deck, id, 4, { 'Suzuki T': 'J Anesth' },
      [['日本麻酔科学会「麻酔薬および麻酔関連薬使用ガイドライン 第4版」Ⅵ（2025）',
        'JSA「麻酔薬および麻酔関連薬使用ガイドライン 第4版」Ⅵ（2025）']]);
    const items = [
      '換気と挿管の準備を整えてから入れる',
      ['効き目はモニターで確かめる（挿管は ', b('TOF 0'), ' を目安に）'],
      '追加は時間ではなく、モニターを見て',
      '高齢、肝腎障害、低体温、吸入麻酔、Mg で長く残る',
      ['終わりに拮抗して、', b('TOF比 0.9 以上'), 'を確認する'],
    ];
    items.forEach(t => rec(id, t));
    const gap = 22, by = 2.0;
    const bh = heightFor(items.length, 26, gap);
    if (by + bh > BOTTOM2 + 1e-6) throw new Error('S9: 箇条書きが出典に近すぎます');
    s.bullets(items, { y: by, size: 26, gap });
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const id = 'S10';
    const s = contentSlide(deck, id, 4);
    const items = [
      ['筋弛緩薬は意識も痛みも取らない。', b('換気は自分の責任')],
      [b('ロクロニウム'), '：0.6 mg/kg、持続 30〜60 分、スガマデクスで拮抗'],
      [b('スキサメトニウム'), '：速く短い。高 K・悪性高熱症に注意'],
      '筋弛緩薬と抗菌薬はアナフィラキシーの主な原因',
      ['深さはモニターで決め、抜管前に ', b('TOF比 0.9 以上')],
    ];
    items.forEach(t => rec(id, t));
    const gap = 26, by = 2.05;
    const bh = heightFor(items.length, 24, gap);
    if (by + bh > BOTTOM1 + 1e-6) throw new Error('S10: 箇条書きが出典に近すぎます');
    s.bullets(items, { y: by, size: 24, gap });
  }

  checkOnSlide();
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
