'use strict';
/**
 * 4-1 術前診察の流れと ASA-PS — 組版スクリプト
 *
 *   cd slides/4-1_preop_visit_asa_ps
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide（S6 は layout の列見出しも）をそのまま書き写している
 *     （変えているのは改行位置だけ）。例外は DEVIATIONS に挙げたものだけ。
 *     書き写しの検査: 各スライドに置いた文字列を順につなげたものが、spec の「」の中身を
 *     （DEVIATIONS を当てたうえで）順につなげたものと一字一句同じでなければ Error（checkOnSlide）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 *     16pt で 2 行に入らない S4・S8 だけ SOURCE_DEVIATIONS で縮める（文献・年は落とさない）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 図はすべてネイティブ図形（箱・線・freeform の矢じり・表）。ラベルは図の外に置く。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, measure, fitInfo, heightFor } = require('../_lib/decklib');

const NAME = '4-1_preop_visit_asa_ps';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.RENDER_DIR || process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);
const RENDER = !process.argv.includes('--no-render');

// ---------------------------------------------------------------------------
// spec.md を読む（タイトル・layout と on-slide の「」・ノート・出典・表紙の subtitle/meta）
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
      // layout 行の「」（S6 の列見出し）も、スライド面に載せる文言として扱う
      if (/^layout:/.test(ln)) for (const m of ln.matchAll(/「([^」]*)」/g)) quoted.push(m[1]);
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
// 出典。16pt・幅 11.6in で 2 行に入らないものだけ縮める。文献そのものと版の年は落とさない。
// items はライブラリに渡す項目（1 項目 = 1 行に入る塊。行内の連結は ' / '）。ライブラリの出典は 2 行まで（3 行は Error）。
//  - S4: ASA の 1 文献だけで 11.12in（1 行がほぼ埋まる）、日本麻酔科学会の 2 文書を ' / ' で並べると 13.00in
//        （空白を詰めても 12.85in）で、2 行に入らない。（2021）（2025）は版を示すので外さず（QA 指摘）、
//        2 文書に共通の発行者名「日本麻酔科学会」を 1 回にまとめて「・」でつなぐ（2 行目 11.15in）。
//  - S8: どの 2 文献を並べても 1 行（11.6in）を超える（最短の組でも 14.07in）ので、' / ' で分けると 3 行になる。
//        2 行に入る改行位置は、spec の順のままでは文献の途中の「…Circulation」の後だけ（1 行目 10.02in）。
//        著者名の後（「Thompson A, et al.」の後）で分けると 2 行目が 12.53in になり入らない。
//        さらに 2 行目が 13.03in になるため（2024 AHA/ACC）を外す（2 行目 10.93in）。
// ---------------------------------------------------------------------------
const SOURCE_DEVIATIONS = {
  S4: {
    replace: [['（2021） / 日本麻酔科学会 悪性高熱症管理ガイドライン', '（2021）・悪性高熱症管理ガイドライン']],
    items: [
      'ASA Practice Advisory for Preanesthesia Evaluation. Anesthesiology 2012;116:522-538',
      '日本麻酔科学会 周術期禁煙プラクティカルガイド（2021）・悪性高熱症管理ガイドライン（2025）',
    ],
  },
  S8: {
    drop: ['（2024 AHA/ACC）'],
    // 1 行目の末尾「Circulation」と 2 行目の頭「2024;…」の間の空白で改行（文献の途中で折り返す）
    items: [
      'Hartle A, et al. Anaesthesia 2016;71:326-337 / Thompson A, et al. Circulation',
      '2024;150:e351-e442 / Kothari P, et al. J Cardiothorac Vasc Anesth 2025;39:2408-2420',
    ],
    wrapInside: true,
  },
};
function sourcesOf(id) {
  const raw = SPEC[id].sources;
  if (!raw) return undefined;
  const dev = SOURCE_DEVIATIONS[id];
  if (!dev) return raw.split(' / ');
  // 検査: 外した（）・置き換えを当てれば spec の sources 行と同じであること（改行位置の空白 1 個を除く）
  let want = raw;
  for (const d of dev.drop || []) {
    if (!want.includes(d)) throw new Error(`SOURCE_DEVIATIONS ${id}: spec の出典に「${d}」がありません`);
    want = want.replace(d, '');
  }
  for (const [from, to] of dev.replace || []) {
    if (!want.includes(from)) throw new Error(`SOURCE_DEVIATIONS ${id}: spec の出典に「${from}」がありません`);
    want = want.replace(from, to);
  }
  const got = dev.items.join(dev.wrapInside ? ' ' : ' / ');
  if (got !== want) throw new Error(`SOURCE_DEVIATIONS ${id}: 出典が spec と違います\n  spec: ${want}\n  got : ${got}`);
  return dev.items;
}

// ---------------------------------------------------------------------------
// 書き写しの検査: スライドに置いた文字列（置いた順）を記録し、spec の「」と突き合わせる
// ---------------------------------------------------------------------------
/**
 * spec からのずれ（deviations）。[元, 後]。build の最後に一覧を出力する。
 *  - S6 III の例（語は削らず、「COPD」の位置だけ 2 番目 → 4 番目に移す）:
 *    表（見出し行＋6 行）と表の下の注記（E）を 1.85〜6.75in に入れるには、表を 4.26in
 *    （1 行 0.367in × 10 行＋余白）以下、つまり III の行を 2 行にする必要がある。
 *    spec の順のまま 2 行にすると、どこで分けても長いほうの行が 25.86 字分以上
 *    （「定期透析中、3 か月より前の心筋梗塞・脳卒中・ステント」8.85in）になり、定義列（「（機能制限あり）」が
 *    入る 2.93in）と分類列を引いた例の列（8.27in、使える幅 8.10in）に入らない。3 行にすると表が 4.63in になり
 *    注記が出典行に重なる。「定期」を外すと意味が変わる（定期透析をしていない末期腎不全は IV）ので外さない。
 *    COPD を「定期透析中、」の後に移すと 2 行（8.01in / 7.91in）に入る。例の並び（順序に意味はない）だけの変更。
 */
const DEVIATIONS = {
  S6: [[
    '不良な糖尿病・高血圧、COPD、BMI 40 以上、定期透析中、3 か月より前の心筋梗塞・脳卒中・ステント',
    '不良な糖尿病・高血圧、BMI 40 以上、定期透析中、COPD、3 か月より前の心筋梗塞・脳卒中・ステント',
  ]],
};
function quotedOf(id) {
  const q = SPEC[id].quoted.slice();
  for (const [from, to] of DEVIATIONS[id] || []) {
    const k = q.indexOf(from);
    if (k < 0) throw new Error(`DEVIATIONS: ${id} の on-slide に「${from}」がありません`);
    q[k] = to;
  }
  return q;
}
const PLACED = {};
const textOf = line => (Array.isArray(line) ? line.map(textOf).join('')
  : (line && typeof line === 'object') ? textOf(line.text) : String(line));
function rec(id, line) { (PLACED[id] = PLACED[id] || []).push(textOf(line)); }
function checkOnSlide() {
  for (const [id, sp] of Object.entries(SPEC)) {
    if (!sp.quoted.length) continue;
    const quoted = quotedOf(id);
    const placed = PLACED[id] || [];
    const whole = quoted.join('\n');
    placed.forEach(l => {
      if (!whole.includes(l)) throw new Error(`${id}: 「${l}」が spec の on-slide にありません`);
    });
    // 改行位置の半角スペースは落とすので、つなぎ目の比較は空白を除いて行う
    const want = quoted.join('').replace(/\s+/g, '');
    const got = placed.join('').replace(/\s+/g, '');
    if (want !== got) {
      let k = 0;
      while (k < want.length && want[k] === got[k]) k++;
      throw new Error(`${id}: スライド面の文言が spec の on-slide と違います（${k} 字目から）\n` +
        `  spec: …${want.slice(Math.max(0, k - 10), k + 30)}\n  got : …${got.slice(Math.max(0, k - 10), k + 30)}`);
    }
  }
  // spec 側で「」の付いていない文言が無いこと（各スライドで1つ以上を置いた）
  for (const id of Object.keys(SPEC)) {
    if (id !== 'S1' && !(PLACED[id] && PLACED[id].length)) throw new Error(`${id}: スライド面に何も置いていません`);
  }
}

/** spec の文言 str を、breaks の各語の直前で改行して行配列にする（文言は変えない。改行位置の半角スペースは落とす） */
function brk(str, ...breaks) {
  const lines = [];
  let from = 0;
  for (const b of breaks) {
    const k = str.indexOf(b, from + 1);
    if (k < 0) throw new Error(`「${str}」に改行位置「${b}」が見つかりません`);
    lines.push(str.slice(from, k).replace(/ $/, ''));
    from = k;
  }
  lines.push(str.slice(from));
  return lines;
}
/** 1行の文字列に部分装飾（太字＋色）を付けて run 配列にする。marks: [[語, 色], ...]。語は行の中に1回だけ */
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
      next.push(color ? { text: w, color, bold: true } : { text: w, bold: true });
      if (k + w.length < r.length) next.push(r.slice(k + w.length));
    }
    if (found !== 1) throw new Error(`mark: 「${line}」に「${w}」が ${found} 回`);
    runs = next;
  }
  return runs;
}

// ---------------------------------------------------------------------------
// 共通の寸法
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;          // 0.55
const BW = LAYOUT.body.w;          // 12.23
const TOP = 1.85;                  // 本文の上端
const GAP = 0.24;                  // 要素間の最小間隔（0.6cm）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;
/**
 * 本文の下端。出典1行（枠 y6.85・h0.5・下揃え。文字の上端は約 7.02in）: 6.75in（指示どおり）。
 * 出典2行（枠 y6.5・h0.85・下揃え）: 2行の文字は枠の下 0.67in に収まり上端は約 6.68in なので、
 * 本文は 6.42in まで（文字との間隔 0.26in。3-8・3-9 と同じ値）。
 */
const BOTTOM_1SRC = LAYOUT.body.bottom;   // 6.75
const BOTTOM_2SRC = 6.42;
const bottomOf = s => (!s.hasSources ? LAYOUT.body.bottom : (s.sourceTop < LAYOUT.source.y - 1e-6 ? BOTTOM_2SRC : BOTTOM_1SRC));
const ACCENT = C.vermilion;        // 注意・要点の語
const ARROW_C = '404040';

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す） */
function contentSlide(deck, id, chapter) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id) });
  if (s.hasSources && s.sourceTop < LAYOUT.source.y2 - 1e-6) throw new Error(`${id}: 出典が3行以上`);
  return s;
}
function checkBottom(s, id, what, bottom) {
  if (bottom > bottomOf(s) + 1e-6) throw new Error(`${id} ${what}: 下端 ${r3(bottom)}in > ${r3(bottomOf(s))}in`);
  console.log(`  ${id} ${what}: bottom ${r3(bottom)} (limit ${r3(bottomOf(s))})`);
}

/**
 * 箇条書きの段落後の間隔(pt)を、本文の下端まで使い切る最大の値にする（minGap 以上）。
 * 高さは lint 式と実機メイリオ推定（1 行 1.5em）の大きい方（heightFor）で測る。
 * LibreOffice の描画は代替フォントで行の高さが低く、下に空きが残って見えるが、PowerPoint では下端近くまで埋まる。
 */
function fillGap(s, id, y, n, size, minGap) {
  let gap = minGap;
  while (y + heightFor(n, size, gap + 1) <= bottomOf(s) + 1e-6) gap++;
  if (y + heightFor(n, size, gap) > bottomOf(s) + 1e-6) throw new Error(`${id}: 箇条書きが下端に入りません（gap ${gap}pt）`);
  console.log(`  ${id}: bullet gap ${gap}pt`);
  return gap;
}

/** 箇条書き（1項目1行）。戻り値は下端 */
function bulletList(s, id, items, o) {
  items.forEach(t => rec(id, t));
  s.bullets(items, { y: o.y, size: o.size, gap: o.gap });
  const bh = heightFor(items.length, o.size, o.gap);
  return r3(o.y + bh);
}

/** 強調ボックス（白地にティールの枠、文字は太字の濃い灰。1行） */
function keyBox(s, id, str, o) {
  rec(id, str);
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.white, line: { color: C.teal, width: 3 }, radius: 0.1,
    text: str, size: o.size || 28, bold: true, color: C.dark, margin: 0.12 });
}

/**
 * 水平・垂直の矢印（軸＋塗りの三角の矢じり）。pptxgenjs の矢じりは線幅に比例して小さいので、
 * 矢じりを自前の三角形（freeform）で描いて大きさを決める。hl: 矢じりの長さ、hw: 矢じりの幅(in)
 */
function flowArrow(s, x1, y1, x2, y2, o = {}) {
  const color = o.color || ARROW_C, width = o.width || 3, hl = o.hl || 0.2, hw = o.hw || 0.26;
  const L = Math.hypot(x2 - x1, y2 - y1), ux = (x2 - x1) / L, uy = (y2 - y1) / L;
  if (Math.abs(ux) > 1e-9 && Math.abs(uy) > 1e-9) throw new Error('flowArrow は水平・垂直のみ');
  const bx = x2 - ux * hl, by = y2 - uy * hl;
  if (uy === 0) s.hline({ x: r3(Math.min(x1, bx)), y: r3(y1), w: r3(Math.abs(bx - x1) + 0.02), color, width });
  else s.vline({ x: r3(x1), y: r3(Math.min(y1, by)), h: r3(Math.abs(by - y1) + 0.02), color, width });
  const px = -uy * hw / 2, py = ux * hw / 2;
  s.polyline([[r3(x2), r3(y2)], [r3(bx + px), r3(by + py)], [r3(bx - px), r3(by - py)]], { color, width: 1, closed: true, fill: color });
}

/**
 * 見出し＋項目のカード（手組み）。文字はすべて箱の内側。
 *  - items: 各項目 = 行配列（行は文字列 or run 配列）
 *  - sep: true なら項目の間に細い区切り線（項目は枠の中で均等に配る）
 *  - marker: true なら項目を箇条書き（•）にして上から gap(pt) 間隔で並べる
 *  - padR / padB: 文字枠の右・下の余白（既定は pad）。文字枠の幅の推定（lint 式）だけに効き、箱は変わらない
 */
const CARD = { pad: 0.18, titleSize: 26, titleGap: 0.12 };
/** 区切り線つきカードの高さ（項目間を g にするとき） */
function sepCardH(c, g) {
  const pad = c.pad != null ? c.pad : CARD.pad, size = c.size || 24, tg = c.titleGap != null ? c.titleGap : CARD.titleGap;
  const hs = c.items.map(lines => heightFor(lines.length, size, 1));
  return r2(2 * pad + heightFor(1, c.titleSize || CARD.titleSize, 0) + tg + hs.reduce((a, b) => a + b, 0) + g * (c.items.length - 1));
}
function itemCard(s, id, c) {
  const pad = c.pad != null ? c.pad : CARD.pad, ts = c.titleSize || CARD.titleSize, size = c.size || 24;
  const padR = c.padR != null ? c.padR : pad, padB = c.padB != null ? c.padB : pad;
  s.box({ x: c.x, y: c.y, w: c.w, h: c.h, fill: C.bg, radius: 0.08 });
  const iw = r3(c.w - pad - padR);
  const th = heightFor(1, ts, 0);
  s.text([c.title], { x: r3(c.x + pad), y: r3(c.y + pad), w: iw, h: th, size: ts, bold: true, color: c.titleColor || C.dark });
  rec(id, c.title);
  const top = c.y + pad + th + (c.titleGap != null ? c.titleGap : CARD.titleGap);
  const bottom = c.y + c.h - padB;
  if (c.marker) {
    const flat = c.items.map(lines => {
      if (lines.length !== 1) throw new Error(`${id} カード「${c.title}」: 箇条書きの項目は1行`);
      return lines[0];
    });
    flat.forEach(l => rec(id, l));
    const bh = heightFor(flat.length, size, c.gap);
    if (top + bh > bottom + 1e-6) throw new Error(`${id} カード「${c.title}」: 高さ不足（必要 ${r3(top + bh + pad - c.y)}in > ${c.h}in）`);
    s.bullets(flat, { x: r3(c.x + pad), y: r3(top), w: iw, h: bh, size, gap: c.gap });
    return;
  }
  // 区切り線つき: 項目の高さの合計を引いた残りを項目間に等分する
  const lineGap = 1; // 項目内の段落間（pt）
  const hs = c.items.map(lines => heightFor(lines.length, size, lineGap));
  const n = c.items.length;
  const free = bottom - top - hs.reduce((a, b) => a + b, 0);
  const g = n > 1 ? free / (n - 1) : 0;
  if (n > 1 && g < 0.2) throw new Error(`${id} カード「${c.title}」: 項目間隔 ${r3(g)}in < 0.2in`);
  let cy = top;
  c.items.forEach((lines, i) => {
    s.text(lines, { x: r3(c.x + pad), y: r3(cy), w: iw, h: hs[i], size, gap: lineGap, color: C.dark });
    lines.forEach(l => rec(id, l));
    cy += hs[i];
    if (i < n - 1) {
      s.hline({ x: r3(c.x + pad), y: r3(cy + g / 2), w: iw, color: C.border, width: 1 });
      cy += g;
    }
  });
}

/**
 * 行の多い表（S6 用。3-11 の tightTable を写したもの）。decklib の s.table と同じ見た目（見出し行 D9D9D9・太字、
 * 罫線 BFBFBF 1pt、上下中央）で、行間を固定値 1.2 em、余白 [3,6,3,6]pt に詰める（22pt のまま）。
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
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['流れ', '問診と\n診察', 'ASA-PS', '延期と\n相談', 'まとめ'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '4-1 術前診察の流れと ASA-PS',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 術前診察は何のため（箇条書き4行＋強調ボックス） =====
  {
    const id = 'S2';
    const s = contentSlide(deck, id, 0);
    const bb = bulletList(s, id, [
      'リスクを見つけ、減らせるものは減らす',
      'その患者に合う麻酔の計画を立てる',
      '説明して同意を得る。不安を減らす',
      '検査を『出す』ための時間ではない',
    ], { y: 1.98, size: 28, gap: 18 });
    const by = r2(bb + 0.34), bh = 0.86;
    keyBox(s, id, '問診と診察で大筋が決まる。検査は確認のため', { x: X0, y: by, w: BW, h: bh });
    checkBottom(s, id, 'box', by + bh);
  }

  // ===== S3 診察の流れ（横並びのフロー5箱＋各箱の下に補足） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 0);
    // 箱は同じ幅 2.06in。22pt 太字の「カルテを読む」（推定 2.053in、実寸約 1.83in）が 1 行に入る最小の幅。
    // 補足は 22pt・中央揃えで、箱より広い 2.40in の枠に入れる（7 字の「必要なものだけ」「申し送り、記録」が 1 行に入る幅）。
    // 補足の枠どうしが本文幅 12.23in に収まるよう、箱の間隔（ピッチ）は (12.23 − 2.40) / 4 = 2.4575in、
    // 箱の間 0.3975in に矢印（箱の端から左右 0.1in あける）。
    // 補足を 1 行にするには「主科の記録、既往、検査、内服」で 4.79in 要り、ピッチ 2.46in では入らない（22pt 未満は不可）。
    // 語句の切れ目（「、」の後）で分け、箱1 は 3 行（6 字を超える 2 行の組がない）、箱4 は 1 行、ほかは 2 行。
    const steps = [
      { label: 'カルテを読む', sub: brk('主科の記録、既往、検査、内服', '既往', '検査') },
      { label: '問診', sub: brk('麻酔歴、薬、運動耐容能', '運動') },
      { label: '診察', sub: brk('気道、心肺、全身', '全身') },
      { label: '検査を選ぶ', sub: ['必要なものだけ'] },
      { label: '計画と説明', sub: brk('同意、申し送り、記録', '申し送り') },
    ];
    const n = steps.length, bw = 2.06, sw = 2.40;
    const pitch = (BW - sw) / (n - 1), g = pitch - bw, off = (sw - bw) / 2;
    const bh = 1.0, subGap = 0.3, subPt = 22, subLine = 4, arrowPad = 0.1;
    const subH = Math.max(...steps.map(st => heightFor(st.sub.length, subPt, subLine)));
    const groupH = bh + subGap + subH;
    const by = r2(TOP + (bottomOf(s) - TOP - groupH) / 2 - 0.05);   // 縦はほぼ中央
    steps.forEach((st, i) => {
      const sx = r3(X0 + i * pitch), x = r3(sx + off);
      s.box({ x, y: by, w: bw, h: bh, fill: C.blue, radius: 0.1, text: st.label, size: 22, bold: true, color: C.white, margin: 0 });
      rec(id, st.label);
      s.text(st.sub, { x: sx, y: r2(by + bh + subGap), w: sw, h: heightFor(st.sub.length, subPt, subLine), size: subPt, gap: subLine, align: 'center', color: C.dark });
      st.sub.forEach(l => rec(id, l));
      if (i < n - 1) flowArrow(s, x + bw + arrowPad, by + bh / 2, x + bw + g - arrowPad, by + bh / 2, { hl: 0.12, hw: 0.22, width: 3 });
    });
    console.log(`  ${id}: flow box w=${bw} pitch=${r3(pitch)} gap=${r3(g)} arrow=${r3(g - 2 * arrowPad)} sub w=${sw} y=${by}`);
    checkBottom(s, id, 'flow', by + groupH);
  }

  // ===== S4 問診で必ず聞くこと（箇条書き6行） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 1);
    const b = bulletList(s, id, [
      '麻酔歴：本人と血縁者の『麻酔で困ったこと』',
      'アレルギー：何で、何が起きたか、いつ',
      '内服：抗血栓薬、糖尿病薬、ステロイド、市販薬・漢方',
      '喫煙・飲酒、運動耐容能（階段 2 フロア分 ≈ 4 METs 超）',
      mark('併存疾患と『最近の変化』（胸痛、息切れ、失神）', ['『最近の変化』', ACCENT]),
      '妊娠の可能性、歯、いびき・無呼吸、逆流',
    ], { y: 1.92, size: 28, gap: fillGap(s, id, 1.92, 6, 28, 12) });
    checkBottom(s, id, 'bullets', b);
  }

  // ===== S5 診察で必ず見ること（3列カード） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 1);
    const cw = r3((BW - 2 * GAP) / 3);
    const cards = [
      { title: '気道（→ 4-2）', items: [['開口、Mallampati'], ['頸部の動き'], ['歯、あご、首の太さ'], ['困難気道の既往']] },
      { title: '心肺', items: [['血圧、脈、SpO2'], ['心雑音、ラ音、喘鳴'], ['頸静脈、下肢の浮腫'], ['会話での息切れ']] },
      { title: '全身', items: [['身長・体重・BMI'], ['静脈路の取りやすさ'], ['脊椎（区域麻酔）'], ['麻痺・感覚の左右差']] },
    ];
    // 見出しの下 0.3in、項目の間 0.4in（区切り線は項目間の中央）。カードは本文ゾーンの縦中央
    const geo = { titleColor: C.blue, size: 24, titleGap: 0.3 };
    const ch = Math.max(...cards.map(c => sepCardH({ ...geo, ...c }, 0.4)));
    const cy = r2(TOP + (bottomOf(s) - TOP - ch) / 2);
    cards.forEach((c, i) => itemCard(s, id, { ...c, ...geo, x: r3(X0 + i * (cw + GAP)), y: cy, w: cw, h: ch }));
    checkBottom(s, id, 'cards', cy + ch);
  }

  // ===== S6 ASA-PS 分類（2020）（表 3列 × 見出し＋6行、E は表の下に注記1行） =====
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 2);
    const head = ['分類', '定義', '例（ASA 2020 の例から）'];
    const cls = t => ({ text: t, bold: true, align: 'center' });
    // [分類, 定義の行, 例の行]
    const body = [
      ['I', ['健康な患者'], ['非喫煙、飲酒なしかごく少量']],
      ['II', ['軽い全身疾患'], ['喫煙、妊娠、BMI 30〜40、良好な糖尿病・高血圧']],
      ['III', brk('重い全身疾患（機能制限あり）', '（機能'),
        brk('不良な糖尿病・高血圧、BMI 40 以上、定期透析中、COPD、3 か月より前の心筋梗塞・脳卒中・ステント', 'COPD')],   // deviation（COPD の位置。DEVIATIONS を参照）
      ['IV', brk('常に生命を脅かす重い疾患', '重い'),
        brk('3 か月以内の心筋梗塞・脳卒中・ステント、進行中の虚血・重症弁膜症、敗血症、ショック', '進行中')],
      ['V', brk('手術なしでは生存が見込めない', '生存'), brk('破裂した大動脈瘤、大量外傷、mass effect のある頭蓋内出血', 'mass')],
      ['VI', ['脳死の臓器提供者'], ['']],
    ];
    head.forEach(t => rec(id, t));
    const rows = [head];
    body.forEach(([c, def, ex]) => {
      rec(id, c);
      def.forEach(l => rec(id, l));
      ex.forEach(l => rec(id, l));
      rows.push([cls(c), def.join('\n'), ex.join('\n')]);
    });
    // 列幅（spec の目安 1.4 / 4.0 / 6.8in から見直し。合計 12.2in）:
    //  spec の列幅だと定義が 2 行、例が II・V で 2 行、III・IV で 3 行になり、表が 5.37in（注記を含め出典行の下まで）になる。
    //  定義 2.93in = 8 字の行「（機能制限あり）」(2.738in)＋左右余白 12pt。分類 1.0in（「分類」0.684in の左右に約 0.08in）。
    //  残りを例の列へ（8.27in、使える幅 8.103in）。II の例が 1 行、III・IV・V の例が 2 行に入る。
    const t = tightTable(s, rows, { x: X0, y: TOP, colW: [1.0, 2.93, 8.27] });
    console.log(`  ${id}: table rows ${t.rowH.join('/')} h=${t.h}`);
    // 表の下の注記（E）。表の下端から 0.24in あけ、本文の下端 6.75in で止める。
    // 枠の高さは lint 式（22pt 1 行 0.40in）。実機メイリオの行送り（0.46in）より低いが、上揃えで下に何もないため見た目は変わらない。
    const note = 'E：緊急。遅れると生命・身体の危険が大きく増すとき（例 III E）';
    rec(id, note);
    const ny = r2(t.y + t.h + GAP), nh = r2(bottomOf(s) - ny);
    s.text([mark(note, ['E：', null])], { x: X0, y: ny, w: BW, h: nh, size: 22, color: C.dark });
    checkBottom(s, id, 'note', ny + nh);
  }

  // ===== S7 ASA-PS の使い方と限界（箇条書き5行） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 2);
    const b = bulletList(s, id, [
      mark('患者の全身状態の分類。手術の危険度は含まない', ['手術の危険度は含まない', ACCENT]),
      '年齢、手術部位、術者の腕、緊急度（E 以外）は入らない',
      '評価者でばらつく。公式の『例』を見て合わせる',
      mark('危機的偶発症の死亡 I 0.32 → IV 70.20／1 万麻酔（JSA 2000 年）', ['I 0.32', C.blue], ['IV 70.20', ACCENT]),
      '分類が目的ではない。計画と説明につなぐ',
    ], { y: 1.98, size: 24, gap: fillGap(s, id, 1.98, 5, 24, 28) });
    checkBottom(s, id, 'bullets', b);
  }

  // ===== S8 延期を考える所見と相談先（2列カード） =====
  {
    const id = 'S8';
    const s = contentSlide(deck, id, 3);
    // 22pt の箇条書き。2 枚は同じ幅 (12.23 − 0.24) / 2 = 5.995in。
    // 最長の行は右の「主科：延期の不利益を一緒に考える」（推定 5.475in、実寸約 4.89in）＋行頭記号 0.35in。
    // 左と上の余白は 0.15in、文字枠の右の余白は 0.02in（推定幅を入れるため。実寸では右に約 0.6in あく）。
    // 行間（段落後の間隔）は 2 枚とも同じ値にし、左の 7 項目が入る最大の値をとる。2 枚とも上詰め。
    const pad = 0.15, padR = 0.02, padB = 0.1, titleGap = 0.08, w = r3((BW - GAP) / 2);
    const cy = TOP, ch = r2(bottomOf(s) - cy);
    const room = ch - pad - padB - heightFor(1, CARD.titleSize, 0) - titleGap;
    let gap = 0;
    while (heightFor(7, 22, gap + 1) <= room + 1e-6) gap++;
    if (gap < 4) throw new Error(`${id}: 箇条書きの行間 ${gap}pt が小さすぎます`);
    console.log(`  ${id}: cards w=${w} h=${ch} bullet gap=${gap}pt`);
    const geo = { y: cy, w, h: ch, pad, padR, padB, titleGap, size: 22, gap, marker: true };
    itemCard(s, id, {
      ...geo, x: X0, title: '延期して評価・治療する', titleColor: ACCENT,
      items: [
        ['新しい胸痛・失神・息切れの悪化'],
        ['未評価の心雑音、新しい不整脈'],
        ['脳卒中・TIA から 3 か月以内'],
        ['再測定で血圧 180/110 mmHg 以上'],
        ['小児の上気道炎（→ 4-9）'],
        ['血糖・カリウムの大きな異常'],
        ['抗血栓薬・糖尿病薬が未調整'],
      ],
    });
    itemCard(s, id, {
      ...geo, x: r3(X0 + w + GAP), title: '相談する相手', titleColor: C.blue,
      items: [
        ['上級医（遠隔でも相談できる）'],
        ['主科：延期の不利益を一緒に考える'],
        ['内科・循環器内科：評価と治療'],
        ['患者・家族：利益と危険を共有'],
      ],
    });
    checkBottom(s, id, 'cards', cy + ch);
  }

  // ===== S9 記録と申し送り（箇条書き5行＋強調ボックス） =====
  {
    const id = 'S9';
    const s = contentSlide(deck, id, 3);
    const bb = bulletList(s, id, [
      'ASA-PS と、その根拠になった所見',
      '気道所見と気道の計画（道具、人）',
      '麻酔方法、ルート・モニター、術後鎮痛と行き先',
      '説明した内容と同意。未解決の問題',
      '当日朝：絶飲食、内服、バイタル、『変わったこと』',
    ], { y: 1.95, size: 26, gap: 12 });
    const by = r2(bb + 0.28), bh = r2(bottomOf(s) - by);
    if (bh < 0.76) throw new Error(`${id}: 強調ボックスの高さ ${bh}in が足りません`);
    keyBox(s, id, '自分がいなくても、読めばわかる記録に', { x: X0, y: by, w: BW, h: Math.min(bh, 0.86) });
    checkBottom(s, id, 'box', by + Math.min(bh, 0.86));
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const id = 'S10';
    const s = contentSlide(deck, id, 4);
    const b = bulletList(s, id, [
      '術前診察＝リスクを見つけ減らし、計画し、説明する',
      '問診：麻酔歴、アレルギー、内服、METs、最近の変化',
      '診察：気道、心肺、全身。自分の目と手で確かめる',
      'ASA-PS は全身状態の分類。手術リスクは別に評価',
      '新しい症状・未評価の異常は延期して相談。記録を残す',
    ], { y: 1.98, size: 28, gap: fillGap(s, id, 1.98, 5, 28, 22) });
    checkBottom(s, id, 'bullets', b);
  }

  checkOnSlide();
  console.log('  deviations（spec の on-slide から変えた文言）:');
  for (const [id, list] of Object.entries(DEVIATIONS)) {
    list.forEach(([from, to]) => console.log(`    ${id}: 「${from}」→「${to}」`));
  }
  for (const [id, d] of Object.entries(SOURCE_DEVIATIONS)) {
    for (const x of d.drop || []) console.log(`    ${id} 出典: 「${x}」を外す`);
    for (const [from, to] of d.replace || []) console.log(`    ${id} 出典: 「${from}」→「${to}」`);
    if (d.wrapInside) console.log(`    ${id} 出典: 文献の途中（${d.items[0].split(' ').pop()} の後）で改行`);
  }
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
