'use strict';
/**
 * 4-4 術前検査の選び方 — 組版スクリプト
 *
 *   cd slides/4-4_preop_testing
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide（S3〜S5 は layout の列見出しも）をそのまま書き写している
 *     （変えているのは改行位置だけ。DEVIATIONS は空）。
 *     書き写しの検査: 各スライドに置いた文字列を順につなげたものが、spec の「」の中身を
 *     順につなげたものと一字一句同じでなければ Error（checkOnSlide）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 *     16pt で 2 行に入らない S2・S6・S8 だけ SOURCE_LAYOUT で行の分け方を決める（文献は落とさない）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 図はすべてネイティブ図形（箱・線・表）。ラベルは図の外に置く。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, measure, fitInfo, heightFor } = require('../_lib/decklib');

const NAME = '4-4_preop_testing';
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
      // layout 行の「」（S3〜S5 の列見出し）も、スライド面に載せる文言として扱う
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
// 出典。16pt・幅 11.6in で 2 行に入らないものだけ、行の分け方をここで決める（文献は落とさない）。
// items はライブラリに渡す項目（1 項目 = 1 行）。
//  - S2・S6・S8: どの 2 文献も ' / ' でつなぐと 1 行に並ばない（S2 最短の組 12.89in、S6 は 11.72in、
//        S8 は ASA の 1 文献だけで 11.12in）。文献を落とさず、文字は 1 字も変えずに、文献の途中の空白で
//        改行する（mode 'wrapInside'。4-1 S8 と同じ扱い）。区切りは全行 ' / '（前後に空白）のまま。
//        S2: 「…Cochrane Database Syst Rev」｜「2019;CD007293 / …」（誌名と年の間。11.58in / 8.96in）
//        S6: 「…Yonekura H, et al. BMJ Open」｜「2019;9:e032306 / Hanai …」（誌名と年の間。10.31in / 8.20in）
//            （以前は区切りの空白を詰めて「e032306/Hanai」としていたが、他スライドの ' / ' と書式がそろわない）
//        S8: 「…for Preanesthesia Evaluation.」｜「Anesthesiology 2012;… / …」（題名と誌名の間。9.32in / 10.58in）
// ---------------------------------------------------------------------------
const SOURCE_LAYOUT = {
  S2: {
    mode: 'wrapInside',
    items: [
      'Schein OD, et al. N Engl J Med 2000;342:168-175 / Keay L, et al. Cochrane Database Syst Rev',
      '2019;CD007293 / Benarroch-Gampel J, et al. Ann Surg 2012;256:518-528',
    ],
  },
  S6: {
    mode: 'wrapInside',
    items: [
      'Yonekura H, et al. Anesth Analg 2018;126:1633-1640 / Yonekura H, et al. BMJ Open',
      '2019;9:e032306 / Hanai S, et al. Sex Transm Dis 2025;52:402-406',
    ],
  },
  S8: {
    mode: 'wrapInside',
    items: [
      'NICE NG45（2016） / ASA Practice Advisory for Preanesthesia Evaluation.',
      'Anesthesiology 2012;116:522-538 / Schein OD, et al. N Engl J Med 2000;342:168-175',
    ],
  },
};
function sourcesOf(id) {
  const raw = SPEC[id].sources;
  if (!raw) return undefined;
  const lay = SOURCE_LAYOUT[id];
  if (!lay) return raw.split(' / ');
  // 検査: 行をつなげ直すと spec の sources 行と同じであること
  let got;
  if (lay.mode === 'wrapInside') got = lay.items.join(' ');
  else if (lay.mode === 'compact') got = lay.items.join(' / ').replace(/\s*\/\s*/g, ' / ');
  else throw new Error(`SOURCE_LAYOUT ${id}: mode 不明`);
  if (got !== raw) throw new Error(`SOURCE_LAYOUT ${id}: 出典が spec と違います\n  spec: ${raw}\n  got : ${got}`);
  return lay.items;
}

// ---------------------------------------------------------------------------
// 書き写しの検査: スライドに置いた文字列（置いた順）を記録し、spec の「」と突き合わせる
// ---------------------------------------------------------------------------
/** spec からのずれ（deviations）。[元, 後]。なし。 */
const DEVIATIONS = {};
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
 * 本文は 6.42in まで（文字との間隔 0.26in。4-1 と同じ値）。
 */
const BOTTOM_2SRC = 6.42;
const bottomOf = s => (!s.hasSources ? LAYOUT.body.bottom : (s.sourceTop < LAYOUT.source.y - 1e-6 ? BOTTOM_2SRC : LAYOUT.body.bottom));
const ACCENT = C.vermilion;        // 要点の数値・語
/**
 * 強調ボックス（S2・S6・S7 で文字 24pt・下端 = 本文の下端にそろえる）。1 行のボックス（S2・S6）は高さ 0.80in・余白 0.1in。
 * S7 は 2 行なので高さだけ 1.12in（実機メイリオ 2 行 1.0in ＋余白 0.06in）。
 * 上の箇条書き（5 行・24pt）は行間 16pt で配り、箇条書きの下端（実機メイリオ推定）とボックスの間を 0.24in 以上あける。
 */
const KEY = { size: 24, h1: 0.8, margin1: 0.1, h2: 1.12, margin2: 0.06 };
const KEY_BULLET_GAP = 16;

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

/** 箇条書き（1項目1行）。戻り値は下端 */
function bulletList(s, id, items, o) {
  items.forEach(t => rec(id, t));
  const bh = heightFor(items.length, o.size, o.gap);
  s.bullets(items, { y: o.y, size: o.size, gap: o.gap, h: bh });
  return r3(o.y + bh);
}

/** 強調ボックス（白地にティールの枠、文字は太字の濃い灰）。str は1行 or 行配列 */
function keyBox(s, id, str, o) {
  const lines = Array.isArray(str) ? str : [str];
  lines.forEach(l => rec(id, l));
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.white, line: { color: C.teal, width: 3 }, radius: 0.1,
    text: lines.join('\n'), size: o.size || 28, bold: true, color: C.dark, margin: o.margin != null ? o.margin : 0.12 });
}

/**
 * 見出し＋項目のカード（手組み。4-1 の itemCard と同じ）。文字はすべて箱の内側。
 * items: 各項目 = 行配列（行は文字列 or run 配列）。項目の間に細い区切り線（項目は枠の中で均等に配る）。
 */
const CARD = { pad: 0.18, titleSize: 26, titleGap: 0.12 };
function sepCardH(c, g) {
  const pad = c.pad != null ? c.pad : CARD.pad, size = c.size || 24, tg = c.titleGap != null ? c.titleGap : CARD.titleGap;
  const hs = c.items.map(lines => heightFor(lines.length, size, 1));
  return r2(2 * pad + heightFor(1, c.titleSize || CARD.titleSize, 0) + tg + hs.reduce((a, b) => a + b, 0) + g * (c.items.length - 1));
}
function itemCard(s, id, c) {
  const pad = c.pad != null ? c.pad : CARD.pad, ts = c.titleSize || CARD.titleSize, size = c.size || 24;
  const padX = c.padX != null ? c.padX : pad;
  s.box({ x: c.x, y: c.y, w: c.w, h: c.h, fill: C.bg, radius: 0.08 });
  const iw = r3(c.w - 2 * padX);
  const th = heightFor(1, ts, 0);
  s.text([c.title], { x: r3(c.x + padX), y: r3(c.y + pad), w: iw, h: th, size: ts, bold: true, color: c.titleColor || C.dark });
  rec(id, c.title);
  const top = c.y + pad + th + (c.titleGap != null ? c.titleGap : CARD.titleGap);
  const bottom = c.y + c.h - pad;
  const lineGap = 1; // 項目内の段落間（pt）
  const hs = c.items.map(lines => heightFor(lines.length, size, lineGap));
  const n = c.items.length;
  const free = bottom - top - hs.reduce((a, b) => a + b, 0);
  const g = n > 1 ? free / (n - 1) : 0;
  const minG = c.minItemGap != null ? c.minItemGap : 0.2;
  if (n > 1 && g < minG - 1e-6) throw new Error(`${id} カード「${c.title}」: 項目間隔 ${r3(g)}in < ${minG}in`);
  let cy = top;
  c.items.forEach((lines, i) => {
    s.text(lines, { x: r3(c.x + padX), y: r3(cy), w: iw, h: hs[i], size, gap: lineGap, color: C.dark });
    lines.forEach(l => rec(id, l));
    cy += hs[i];
    if (i < n - 1) {
      s.hline({ x: r3(c.x + padX), y: r3(cy + g / 2), w: iw, color: C.border, width: 1 });
      cy += g;
    }
  });
}

/**
 * 行の多い表（3-11 の tightTable を写したもの）。decklib の s.table と同じ見た目（見出し行 D9D9D9・太字、
 * 罫線 BFBFBF 1pt、上下中央）で、行間を固定値 1.2 em、余白 [3,6,3,6]pt に詰める（22pt のまま）。
 * s.table は行高を実機メイリオの 1.5 em で取るため、S3・S4・S5 の表が本文ゾーンに入らない
 * （S3 は表の下の1行と合わせて 6.88in、S4 は 6.97in、S5 は 6.54in まで下がり、出典 2 行の上端 6.42in を超える）。
 * セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、重なり検査用の登録も同じく行う。
 * cell: 文字列（'\n' で改行）| {text, color, bold, align}
 * o.extra: 本文の行（見出し行を除く）の高さに足すゆとり（in）。余白 [3,6,3,6]pt は変えず、文字は上下中央のまま。
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
    rowH.push(r2(Math.ceil((maxLines * LS / 72 + (m[0] + m[2]) / 72) * 100 - 1e-6) / 100 + (head ? 0 : (o.extra || 0))));
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

/** 表の組み立て: head（列見出し）と body（各セル = 行配列）を rec しながら tightTable の rows にする */
function buildRows(id, head, body) {
  head.forEach(t => rec(id, t));
  const rows = [head];
  body.forEach(r => {
    rows.push(r.map(cell => {
      const c = Array.isArray(cell) ? { lines: cell } : cell;
      c.lines.forEach(l => rec(id, l));
      return { text: c.lines.join('\n'), color: c.color, bold: c.bold, align: c.align };
    }));
  });
  return rows;
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['考え方', '検査別', '国内\n事情', 'まとめ'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '4-4 術前検査の選び方',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 ルーチン検査はなぜ害か（箇条書き5行＋強調ボックス） =====
  {
    const id = 'S2';
    const s = contentSlide(deck, id, 0);
    const bb = bulletList(s, id, [
      mark('1 項目で健常者の 5% が範囲外。20 項目なら最大 64%（独立と仮定）', ['最大 64%', ACCENT]),
      mark('管理に影響しうる異常は 0.22%、胸部 X 線で管理を変えたのは 0.1%', ['0.22%', ACCENT], ['0.1%', ACCENT]),
      mark('白内障 19,557 件の RCT：合併症は両群とも 31.3/1,000 手術', ['両群とも', ACCENT]),
      mark('当日の検査で異常が出ても 61.6% はそのまま手術（ヘルニア）', ['61.6%', ACCENT]),
      mark('費用 2.55 倍（白内障）、7 日以内の再受診が増加（日帰り手術）', ['2.55 倍', ACCENT], ['再受診が増加', ACCENT]),
    ], { y: 1.95, size: 24, gap: KEY_BULLET_GAP });
    const bh = KEY.h1, by = r2(bottomOf(s) - bh);
    if (by - bb < GAP - 1e-6) throw new Error(`${id}: 箇条書きと強調ボックスの間 ${r3(by - bb)}in < ${GAP}in`);
    keyBox(s, id, '異常が出たとき何をするか決められない検査は、出さない', { x: X0, y: by, w: BW, h: bh, size: KEY.size, margin: KEY.margin1 });
    console.log(`  ${id}: bullets bottom ${bb}, box y ${by}`);
    checkBottom(s, id, 'box', by + bh);
  }

  // ===== S3 患者の状態 × 手術の侵襲（表 4列 × 見出し＋3行、下に1行） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 0);
    // 列幅は spec どおり 4.4 / 2.6 / 2.6 / 2.6in（ASA の 3 列は同じ幅。合計 12.2in）。
    //  ASA 列のセル内の幅は 2.433in で、22pt 全角 7 字（2.396in）まで。8 字の「腎機能・心電図を」「腎機能・心電図。」
    //  「血算・凝固は考慮」は 1 行に入らないので、7 字以内の語の切れ目で分ける:
    //  小 ASA 3〜4「腎機能・｜心電図を考慮」、中 ASA 2「腎機能・｜心電図を｜併存症で考慮」、
    //  中 ASA 3〜4「腎機能・｜心電図。血算・｜凝固は考慮」、大 ASA 2「血算、腎機能、｜心電図」。
    //  行の高さは 1 行目 2 行・2〜3 行目 3 行（ASA 1 の大手術も 3 行）で、表の高さは変わらない。
    //  手術の侵襲 4.4in は「中（初発の鼠径ヘルニア、」4.107in・「大（腹式子宮全摘、TURP、」が 1 行に入る幅。
    const head = ['手術の侵襲（NICE NG45）', 'ASA 1', 'ASA 2', 'ASA 3〜4'];
    const rows = buildRows(id, head, [
      [brk('小（皮膚病変の切除、乳腺膿瘍ドレナージ）', '乳腺'), ['検査なし'], ['検査なし'], brk('腎機能・心電図を考慮', '心電図')],
      [brk('中（初発の鼠径ヘルニア、静脈瘤、扁桃、膝関節鏡）', '静脈瘤'), ['検査なし'],
        brk('腎機能・心電図を併存症で考慮', '心電図', '併存症'), brk('腎機能・心電図。血算・凝固は考慮', '心電図', '凝固')],
      [brk('大（腹式子宮全摘、TURP、甲状腺、人工関節、肺）', '甲状腺'), brk('血算。腎機能と 65 歳超の心電図は考慮', '65', '心電図'),
        brk('血算、腎機能、心電図', '心電図'), brk('血算、腎機能、心電図、凝固は肝疾患で', '心電図', '凝固')],
    ]);
    const t = tightTable(s, rows, { x: X0, y: TOP, colW: [4.4, 2.6, 2.6, 2.6] });
    console.log(`  ${id}: table rows ${t.rowH.join('/')} h=${t.h}`);
    const line = '凝固・呼吸機能・胸部 X 線・尿検査は、ルーチンには出さない';
    rec(id, line);
    const ly = r2(t.y + t.h + 0.3), lh = heightFor(1, 24, 0);
    s.text([mark(line, ['ルーチンには出さない', ACCENT])], { x: X0, y: ly, w: BW, h: lh, size: 24, color: C.dark });
    checkBottom(s, id, 'line', ly + lh);
  }

  // ===== S4 血液検査：誰に、いつ（表 3列 × 見出し＋5行） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 1);
    // 列幅（spec の目安 2.2 / 4.6 / 5.4in から見直し。合計 12.2in）:
    //  見出し＋5 行を 2 行ずつに収めないと出典 2 行の上（6.42in）に入らない（腎機能の行が 3 行になると 6.77in）。
    //  「出す相手」の AKI リスクを 2 行に入れるには 5.87in 要る（1 行目「AKI リスク（CKD、心不全、肝疾患、」5.701in）。
    //  検査の列は「妊娠反応」「腎機能・」（4 字 1.369in）が入る 1.54in に。残りをポイントの列へ（4.79in）。
    const head = ['検査', '出す相手', 'ポイント'];
    const rows = buildRows(id, head, [
      [['血算'], brk('大手術は全員、中手術は併存症で（NICE）', '中手術'), brk('貧血の補正に 2〜4 週かかる。外来で早く出す', 'かかる')],
      [brk('腎機能・電解質', '電解質'), brk('AKI リスク（CKD、心不全、肝疾患、糖尿病、65 歳以上、腎毒性薬）', '糖尿病'),
        brk('緊急・腹腔内手術、周術期の NSAIDs もリスク', '周術期')],
      [['凝固'], ['出血歴・肝疾患・抗凝固薬'], brk('出血歴の問診が先。正常でも出血は否定できない', '正常')],
      [['HbA1c'], ['糖尿病で 3 か月以内の値がなければ'], brk('英国 AAGBI 2015：8.5% 以上で待機手術を延期', '8.5%')],
      [['妊娠反応'], brk('可能性を聞き、疑えば同意を得て検査', '疑えば'), brk('未認識の妊娠 0.3%。一律検査は勧めない', '一律')],
    ]);
    const t = tightTable(s, rows, { x: X0, y: TOP, colW: [1.54, 5.87, 4.79] });
    console.log(`  ${id}: table rows ${t.rowH.join('/')} h=${t.h}`);
    checkBottom(s, id, 'table', t.y + t.h);
  }

  // ===== S5 心電図・画像・肺機能（表 3列 × 見出し＋4行） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 1);
    // 列幅（spec の目安 2.2 / 5.6 / 4.4in から見直し。合計 12.2in）:
    //  検査の列は「呼吸機能検査」（6 字 2.053in）を 1 行に入れるため 2.22in（セル内 2.053in）。
    //  「12 誘導心電図」（2.197in）は入らないので「12 誘導｜心電図」の 2 行のまま。
    //  「心疾患・症状・リスク因子があり、」（5.476in）と「新しい呼吸器症状・所見があるとき」を 1 行に入れるため
    //  出す相手を 5.65in に。出さない相手は「無症状で安定した心疾患の」（4.107in）が入る 4.33in。
    //  胸部 X 線の「ルーチンには出さない（NICE）。」は 5.215in で、どの配分でも入らないため 3 行に分ける。
    const head = ['検査', '出す相手', '出さない相手'];
    const rows = buildRows(id, head, [
      [brk('12 誘導心電図', '心電図'), brk('心疾患・症状・リスク因子があり、中〜高リスク手術', '中〜高'),
        brk('低リスク手術の無症状者（AHA/ACC 3）', '（AHA')],
      [['胸部 X 線'], ['新しい呼吸器症状・所見があるとき'], brk('ルーチンには出さない（NICE）。年齢も理由にならない', '出さない', '年齢')],
      [['心エコー'], brk('心不全の徴候、新しい息切れ、心雑音+症状、中等度以上の弁膜症', '心雑音'),
        brk('無症状で安定した心疾患の低〜中リスク手術', '低〜中')],
      [['呼吸機能検査'], ['肺切除、原因不明の呼吸困難'], brk('それ以外はルーチンに出さない（ACP）', '出さない')],
    ]);
    // 本文ゾーンに 0.47in 余るので、本文の 4 行に 0.08in ずつゆとりを足す（表の下端 6.27in）
    const t = tightTable(s, rows, { x: X0, y: TOP, colW: [2.22, 5.65, 4.33], extra: 0.08 });
    console.log(`  ${id}: table rows ${t.rowH.join('/')} h=${t.h}`);
    checkBottom(s, id, 'table', t.y + t.h);
  }

  // ===== S6 国内の慣行と保険（箇条書き5行＋強調ボックス） =====
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 2);
    const bb = bulletList(s, id, [
      mark('低リスク手術でも約 6 割が術前血液検査（全国レセプト、成人）', ['約 6 割', ACCENT]),
      mark('小児は半数に血液型、8 割に血小板。輸血は 0.008%', ['0.008%', ACCENT]),
      mark('1 施設：手術患者の 94.8% に梅毒検査。要精査例の精査は 52.9%', ['94.8%', ACCENT], ['52.9%', ACCENT]),
      mark('手術前 1 週間以内の主な検査は『手術前医学管理料』に包括される', ['『手術前医学管理料』', C.blue]),
      mark('療担規則：検査は『診療上必要があると認められる場合』に', ['『診療上必要があると認められる場合』', C.blue]),
    ], { y: 1.95, size: 24, gap: KEY_BULLET_GAP });
    const bh = KEY.h1, by = r2(bottomOf(s) - bh);
    if (by - bb < GAP - 1e-6) throw new Error(`${id}: 箇条書きと強調ボックスの間 ${r3(by - bb)}in < ${GAP}in`);
    keyBox(s, id, '『病院のセットだから』は検査の根拠にならない。セットを見直す', { x: X0, y: by, w: BW, h: bh, size: KEY.size, margin: KEY.margin1 });
    console.log(`  ${id}: bullets bottom ${bb}, box y ${by}`);
    checkBottom(s, id, 'box', by + bh);
  }

  // ===== S7 輸血の準備と T&S（2列カード＋強調ボックス） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 2);
    // カード幅: 準備量のカードを 7.4in にして「MSBOS：術式別の準備量（C/T 比 1.5 以下）」（22pt 6.865in）を
    // 内側 7.1in に余裕をもって 1 行に入れる（メイリオは英数字が広いため）。検査のカードは残りの 4.59in
    // （内側 4.29in。「交差適合の検体は 3 日以内」4.161in が 1 行、「ABO は別の時点の 2 検体で確認」4.839in は 2 行）。
    // 両カードとも本文 4 行・3 項目で高さがそろう（区切り線の位置は、2 行の項目が左は 2 番目・右は 1 番目なので左右で違う）。
    // 下の強調ボックスは 24pt・2 行（KEY.h2）。カードとの間を 0.36in あけるため、カードの上下余白 0.15in・
    // 見出し 24pt・見出しと本文の間 0.08in・項目間 0.17in 以上に詰める。
    const pad = 0.15, padX = 0.15, size = 22, titleSize = 24, titleGap = 0.08, minItemGap = 0.17;
    const wR = 7.4, wL = r3(BW - GAP - wR);
    const L = {
      title: '検査', titleColor: C.blue,
      items: [['ABO・RhD・不規則抗体'], brk('ABO は別の時点の 2 検体で確認', '2 検体'), ['交差適合の検体は 3 日以内']],
    };
    const R = {
      title: '準備量', titleColor: ACCENT,
      items: [
        brk('T&S：輸血の可能性が低い手術。事前の交差試験を省く', '事前').map((l, i) => (i === 0 ? mark(l, ['T&S：', null]) : l)),
        [mark('MSBOS：術式別の準備量（C/T 比 1.5 以下）', ['MSBOS：', null])],
        [mark('SBOE：出血量・Hb から患者ごとに', ['SBOE：', null])],
      ],
    };
    const boxLines = brk('離島は血液の取り寄せに時間がかかる。在庫と搬送時間を前提に早めに決める', '在庫');
    const bh = KEY.h2, by = r2(bottomOf(s) - bh);
    const ch = r2(by - 0.36 - TOP);   // カードと強調ボックスの間 0.36in（最小 0.24in）
    const cc = { pad, padX, size, titleSize, titleGap, minItemGap };
    const need = Math.max(sepCardH({ ...L, ...cc }, minItemGap), sepCardH({ ...R, ...cc }, minItemGap));
    if (need > ch + 1e-6) throw new Error(`${id}: カードの高さ不足（必要 ${need}in > ${ch}in）`);
    itemCard(s, id, { ...L, ...cc, x: X0, y: TOP, w: wL, h: ch });
    itemCard(s, id, { ...R, ...cc, x: r3(X0 + wL + GAP), y: TOP, w: wR, h: ch });
    keyBox(s, id, boxLines, { x: X0, y: by, w: BW, h: bh, size: KEY.size, margin: KEY.margin2 });
    console.log(`  ${id}: cards h ${ch} (need ${need}), box y ${by}`);
    checkBottom(s, id, 'box', by + bh);
  }

  // ===== S8 まとめ（箇条書き5行） =====
  {
    const id = 'S8';
    const s = contentSlide(deck, id, 3);
    const b = bulletList(s, id, [
      mark('ルーチン検査は偽陽性を生み、管理に影響しうる異常は 0.1〜0.22%', ['0.1〜0.22%', ACCENT]),
      '患者の状態（ASA）× 手術の侵襲で選ぶ（NICE の表）',
      '血算は大手術の外来で早く。凝固は出血歴の問診が先',
      '心電図・X 線は症状と疾患で。低リスク手術では年齢だけで出さない',
      mark('院内セットは慣行。『異常が出たらどうするか』で決める', ['『異常が出たらどうするか』', ACCENT]),
    ], { y: 2.0, size: 24, gap: 31 });   // 行間 31pt で本文ゾーン（出典 2 行の上 6.42in）まで配る
    checkBottom(s, id, 'bullets', b);
  }

  checkOnSlide();
  console.log('  deviations（spec の on-slide から変えた文言）:', Object.keys(DEVIATIONS).length ? '' : 'なし');
  for (const [id, list] of Object.entries(DEVIATIONS)) {
    list.forEach(([from, to]) => console.log(`    ${id}: 「${from}」→「${to}」`));
  }
  for (const [id, d] of Object.entries(SOURCE_LAYOUT)) {
    console.log(`    ${id} 出典: ${d.mode === 'compact' ? '区切りの前後の空白を詰めた' : '文献の途中で改行（文字は変えない）'}`);
  }
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
