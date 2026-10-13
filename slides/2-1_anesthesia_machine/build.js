'use strict';
/**
 * 2-1 麻酔器の構造とガスの流れ — 組版スクリプト
 *
 *   cd slides/2-1_anesthesia_machine
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をここに転記（変えているのは改行位置だけ）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文）。【要確認】の注記はノートにだけ残る。
 *   - 出典は spec の sources 行を項目に分けて渡す（2行への折り返しはライブラリ）。
 *     S3・S4・S5・S7 は 16pt・2行に収まらないため最小限の略記にし、全文をノート末尾に残す。
 *
 * デッキ内の組版ルール（QA 指摘を受けて統一）
 *   - 3 列カード: 幅 3.91in で等幅（S9 だけ 3.85/3.85/4.05）、見出し 28pt・本文 22pt。同じ行番号の項目は全カードで同じ y（行グリッド）。
 *   - 箇条書き: 横並び（図＋箇条書き）は 24pt・2 行の項目で 1 行目（見出し行）を太字。全幅は 26pt。
 *     行頭記号はすべて濃灰 404040（強調色に染めない）。
 *   - ティールの文字は 008083（白・F2F2F2 上でコントラスト 4.2:1 以上）。図形の線・塗りは 00A8AA のまま。
 *   - 2 行の項目は 1 行ずつテキスト枠を置く（実機メイリオの行送り 1.5em と同じ位置になる）。
 * PDF / slide-*.jpg はリポジトリの外（RENDER_DIR または DECK_RENDER_DIR。既定は OS の一時ディレクトリ）に出す。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, heightFor, fits, measure } = require('../_lib/decklib');

const NAME = '2-1_anesthesia_machine';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.RENDER_DIR || process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);
const RENDER = !process.argv.includes('--no-render');

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
/** 略記した出典のスライド用: ノートの末尾に出典の全文を足す（ノート本文は原文のまま） */
function notesWithFullSource(id) {
  return notesOf(id) + '\n出典（全文）: ' + SPEC[id].sources;
}
/** 1 行に収まる項目は ' / ' で分けて渡す（ライブラリが 1〜2 行に詰める） */
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  if (!s) throw new Error(`spec.md の ${id} に sources がありません`);
  return s.split(' / ');
}
/**
 * 2 行を明示して渡す出典。各行が 16pt・11.6in に収まること、
 * ref（既定は spec の sources 行）を空白1つで連結した形と一致することを確認する。
 */
function sourcesLines(id, lines, ref) {
  const S = LAYOUT.source;
  const want = ref != null ? ref : SPEC[id].sources;
  if (lines.join(' ') !== want) throw new Error(`${id} 出典の行分けが元の文言と一致しません:\n  ${lines.join(' ')}\n  ${want}`);
  lines.forEach(l => { if (!fits(l, S.size, S.w)) throw new Error(`${id} 出典「${l}」が 1 行に収まりません`); });
  return lines;
}
/** spec の sources 行の一部を置き換えた略記（置換元が見つからなければ止める） */
function abbreviate(id, pairs) {
  let s = SPEC[id].sources;
  for (const [from, to] of pairs) {
    if (!s.includes(from)) throw new Error(`${id} 出典に「${from}」がありません`);
    s = s.replace(from, to);
  }
  return s;
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const GAP = 0.24;                // 要素間の最小間隔
const CARD_W = 3.91;             // 3 列カードの幅（等幅）
const TEAL_T = C.teal;           // ティールの文字色はライブラリ既定（全講義で共通。変えるなら decklib 側で）
const MARK = C.chipText;         // 行頭記号の色（全スライド共通）
// ⇄ だけに当てる記号フォント。Windows と Office for Mac に同梱され、矢印ブロック（U+2190–21FF）を持つ。
// （メイリオの ⇄ は環境により代替フォントの小さな字形になるため。LibreOffice では FreeSerif に置換される）
const SYM_FONT = 'Cambria Math';
const SYM_RE = /([⇄])/;
const r2 = v => Math.round(v * 100) / 100;

/** 1 行の指定 → {runs, dx}。line は 文字列 | run 配列 | {t: 文字列|run配列, dx} */
function normL(line, bold) {
  let t = line, dx = 0;
  if (line && typeof line === 'object' && !Array.isArray(line) && 't' in line) { t = line.t; dx = line.dx || 0; }
  let runs = Array.isArray(t) ? t.map(r => (typeof r === 'string' ? { text: r } : { ...r })) : [{ text: String(t) }];
  if (bold) runs = runs.map(r => ({ ...r, bold: r.bold != null ? r.bold : true }));
  return { runs, dx };
}

/**
 * 1 行のテキスト枠。⇄ を含む行は記号だけ SYM_FONT にした run で直接置く
 * （decklib の run はフォント指定を持たないため。幅は fits で検査し、重なり検査にも登録する）。
 */
function line1(s, line, o) {
  const { runs, dx } = normL(line, o.bold);
  const x = r2(o.x + dx), w = r2(o.w - dx), size = o.size;
  const h = heightFor(1, size, 0);
  const text = runs.map(r => r.text).join('');
  if (!SYM_RE.test(text)) {
    s.text([runs], { x, y: r2(o.y), w, h, size, color: o.color || C.dark, align: o.align || 'left' });
    return h;
  }
  if (!fits(text, size, w)) throw new Error(`p${s.page} 「${text}」が枠幅 ${w}in に収まりません`);
  const parts = [];
  runs.forEach(r => String(r.text).split(SYM_RE).filter(Boolean).forEach(t => parts.push({ ...r, text: t, sym: SYM_RE.test(t) })));
  s.raw.addText(parts.map(r => ({
    text: r.text,
    options: { fontFace: r.sym ? SYM_FONT : FONT, lang: 'ja-JP', fontSize: size, color: r.color || o.color || C.dark, bold: r.bold || undefined },
  })), { x, y: r2(o.y), w, h, margin: 0, isTextBox: true, fontFace: FONT, lang: 'ja-JP', fontSize: size, align: o.align || 'left', valign: 'top' });
  s._reg('txt', x, r2(o.y), w, h, `text「${text.slice(0, 12)}」`);
  return h;
}

/**
 * 箇条書き（1 項目 = 1〜2 行）。items: [[行1, 行2], ...]。行頭記号は 1 行目だけ、色は全項目 MARK。
 * o.leadBold で各項目の 1 行目を太字。戻り値は下端 y。
 */
function itemList(s, items, o) {
  const size = o.size || 24, ig = o.itemGap != null ? o.itemGap : 0.25;
  const ind = o.indent != null ? o.indent : r2(size * 1.15 / 72 + 0.02);
  const lh = heightFor(1, size, 0);
  let y = o.y;
  items.forEach((lines, k) => {
    s.text('•', { x: o.x, y: r2(y), w: r2(Math.max(ind, measure('•', size) + 0.01)), h: lh, size, color: MARK });
    lines.forEach((l, i) => line1(s, l, { x: o.x + ind, y: y + i * lh, w: o.w - ind, size, bold: o.leadBold && i === 0, color: o.color }));
    y += lines.length * lh + (k < items.length - 1 ? ig : 0);
  });
  return y;
}
/** itemList の高さ */
function itemListH(items, size, ig) {
  const lh = heightFor(1, size, 0);
  return items.reduce((a, l) => a + l.length * lh, 0) + ig * (items.length - 1);
}

/**
 * 3 列カード（等幅・全カード同じ高さ）。見出し＋項目。同じ行番号の項目は全カードで同じ y に置く
 * （行の高さ = その行で最も行数の多い項目に合わせる）。文字はすべて箱の内側。
 */
function cardRow(s, cards, o) {
  const n = cards.length;
  const ws = o.ws || Array(n).fill(CARD_W);
  const gap = (BW - ws.reduce((a, b) => a + b, 0)) / (n - 1);
  if (gap < GAP - 1e-6) throw new Error(`cardRow: カードの間隔 ${gap.toFixed(3)}in < ${GAP}in`);
  const ts = o.titleSize || 28, bs = o.bodySize || 22;
  const pad = o.pad != null ? o.pad : 0.12, padY = o.padY != null ? o.padY : 0.2;
  const ig = o.itemGap != null ? o.itemGap : 0.3, tg = o.titleGap != null ? o.titleGap : 0.18;
  const th = heightFor(1, ts, 0), lh = heightFor(1, bs, 0);
  const nRows = Math.max(...cards.map(c => c.items.length));
  const rowH = Array.from({ length: nRows }, (_, k) => Math.max(...cards.map(c => (c.items[k] ? c.items[k].length : 0))) * lh);
  const h = r2(2 * padY + th + tg + rowH.reduce((a, b) => a + b, 0) + ig * (nRows - 1) + 0.005);
  const rowY = [];
  let yy = o.y + padY + th + tg;
  rowH.forEach(hh => { rowY.push(yy); yy += hh + ig; });
  let x = X0;
  cards.forEach((c, i) => {
    const w = ws[i], iw = w - 2 * pad;
    s.box({ x: r2(x), y: o.y, w, h, fill: C.bg, radius: 0.08 });
    s.text(c.title, { x: r2(x + pad), y: r2(o.y + padY), w: r2(iw), h: th, size: ts, bold: true, color: c.titleColor || C.blue });
    c.items.forEach((lines, k) => {
      lines.forEach((l, j) => line1(s, l, { x: x + pad, y: rowY[k] + j * lh, w: iw, size: bs }));
    });
    x += w + gap;
  });
  return { h, bottom: o.y + h };
}

/** 流れの向きを示す山形（ネイティブ図形）。dir: 'r' | 'l' | 'u'。幅 2a・奥行き b */
function chevron(s, cx, cy, dir, color) {
  const a = 0.14, b = 0.08;
  const pts = dir === 'r' ? [[cx - b, cy - a], [cx + b, cy], [cx - b, cy + a]]
    : dir === 'l' ? [[cx + b, cy - a], [cx - b, cy], [cx + b, cy + a]]
      : [[cx - a, cy + b], [cx, cy - b], [cx + a, cy + b]];
  s.polyline(pts.map(p => [r2(p[0]), r2(p[1])]), { color, width: 4.5 });
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['供給', '流量', '気化器', '回路', '安全'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '2-1 麻酔器の構造とガスの流れ',
  });

  // ===== S1 表紙 =====
  deck.cover({
    title: '麻酔器の構造とガスの流れ',
    subtitle: '2-1 モニタリング・麻酔器',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 ガスの流れの全体像（7 チップのフロー＋3 区分の横線とラベル＋箇条書き2行） =====
  {
    const s = deck.slide({ title: 'ガスの流れの全体像', chapter: 0, notes: notesOf('S2'), sources: sourcesOf('S2') });
    // 境目になる装置（圧力調整器・流量計）は幅を広げ、中圧部の区間を確保する
    const chips = [
      { t: '配管・\nボンベ', w: 1.2 },
      { t: '圧力\n調整器', w: 1.5 },
      { t: '流量計', w: 1.5 },
      { t: '気化器', w: 1.2 },
      { t: '共通ガス\n出口', w: 1.54 },
      { t: '呼吸回路', w: 1.54 },
      { t: '患者', w: 1.0 },
    ];
    const cy = 1.95, ch = 1.05;
    const az = (BW - chips.reduce((a, c) => a + c.w, 0)) / (chips.length - 1);   // チップ間（矢印の区間）
    let x = X0;
    chips.forEach((c, i) => {
      c.x = x;
      s.box({ x: r2(x), y: cy, w: c.w, h: ch, fill: C.blue, radius: 0.1, text: c.t, size: 22, bold: true, color: C.white, margin: 0.08 });
      if (i < chips.length - 1) s.arrow({ x1: r2(x + c.w + 0.08), y1: cy + ch / 2, x2: r2(x + c.w + az - 0.08), y2: cy + ch / 2, color: C.chipText, width: 3 });
      x += c.w + az;
    });
    // 3 区分の横線。境目 = 圧力調整器と流量計（チップの中心）。境目の下は 0.5in あけ、
    // チップの中心から破線を下ろして「この装置が境目」と示す。
    const ly = cy + ch + 0.42, tick = 0.13, half = 0.25;
    const b1 = chips[1].x + chips[1].w / 2, b2 = chips[2].x + chips[2].w / 2, e = chips[4].x + chips[4].w;
    const segs = [
      { x1: X0, x2: b1 - half, color: C.vermilion, tcol: C.vermilion, name: '高圧部', sub: '（ボンベ）' },
      { x1: b1 + half, x2: b2 - half, color: C.blue, tcol: C.blue, name: '中圧部', sub: '（配管〜流量計）' },
      { x1: b2 + half, x2: e, color: C.teal, tcol: TEAL_T, name: '低圧部', sub: '（流量計〜出口）' },
    ];
    segs.forEach(g => {
      s.hline({ x: r2(g.x1), y: ly, w: r2(g.x2 - g.x1), color: g.color, width: 5 });
      s.vline({ x: r2(g.x1), y: r2(ly - tick), h: 2 * tick, color: g.color, width: 3 });
      s.vline({ x: r2(g.x2), y: r2(ly - tick), h: 2 * tick, color: g.color, width: 3 });
    });
    [b1, b2].forEach(b => s.vline({ x: r2(b), y: cy + ch, h: r2(ly + tick - cy - ch), color: C.gray, width: 1.5, dash: 'dash' }));
    // ラベル（22pt・2 行）は線の下。区分の中央に置き、隣と 0.24in 以上あける（重なるときは右へ）
    const lyT = ly + tick + GAP;
    let prevR = -Infinity;
    segs.forEach(g => {
      const tw = Math.max(measure(g.name, 22), measure(g.sub, 22));
      let lx = Math.max(X0, (g.x1 + g.x2) / 2 - tw / 2);
      if (lx < prevR + GAP) lx = prevR + GAP;
      if (lx > g.x2 - 0.3) throw new Error(`S2 ラベル「${g.name}」が区分から離れすぎます`);
      s.text([[{ text: g.name, bold: true, color: g.tcol }], g.sub], { x: r2(lx - 0.02), y: r2(lyT), w: r2(tw + 0.04), size: 22, align: 'center' });
      prevR = lx + tw;
    });
    const items = [
      ['酸素フラッシュは流量計と気化器を通らず、出口に直接入る'],
      [['低圧部は圧が低く、漏れが見つけにくい → ', { text: 'リークテスト', bold: true, color: TEAL_T }]],
    ];
    itemList(s, items, { x: X0, y: r2(lyT + heightFor(2, 22, 0) + 0.5), w: BW, size: 26, itemGap: 0.22 });
  }

  // ===== S3 医療ガスの供給（3 列カード＋下に一行の注意） =====
  {
    // 出典: 4 件で 16pt・2 行に収まらない（約107字相当 > 93字相当）ため、厚生労働省通知の日付を略す。全文はノートへ。
    const short = abbreviate('S3', [['（医政発0817第6号、令和2年8月17日、令和3年12月16日一部改正）', '（医政発0817第6号）']]);
    const s = deck.slide({ title: '医療ガスの供給', chapter: 0, notes: notesWithFullSource('S3'), sources: short.split(' / ') });
    const row = cardRow(s, [
      { title: '配管（中央供給）', titleColor: C.blue, items: [['約 0.4 MPa で供給'], ['端末の色：酸素 緑'], ['亜酸化窒素 青・空気 黄']] },
      { title: 'ボンベ（予備）', titleColor: C.vermilion, items: [['酸素ボンベは黒色'], ['ピンインデックスで', '取り違え防止'], ['残量は圧力計で見る']] },
      { title: '吸引・排除', titleColor: TEAL_T, items: [['吸引の端末は黒'], ['余剰ガス排除装置に', 'つなぐ']] },
    ], { y: 1.92, itemGap: 0.3, titleGap: 0.14, padY: 0.16 });
    s.text('ボンベの色と配管の色は別の規格。酸素ボンベは黒、配管は緑', { x: X0, y: r2(row.bottom + 0.3), w: BW, size: 24, bold: true, color: C.red });
  }

  // ===== S4 流量計と酸素フラッシュ（左: 流量計の模式図 / 右: 箇条書き4項目） =====
  {
    // 出典: 3 件の合計（約90字相当）は 2 行に入るが、行の切れ目が文献の途中になり 2 行目が 0.15 字分あふれる。
    //       Subrahmanyam の頁を省略形（472-480 → 472-80）にして 2 行に収める。全文はノートへ。
    const short = abbreviate('S4', [['2013;57:472-480', '2013;57:472-80']]);
    const s = deck.slide({
      title: '流量計と酸素フラッシュ', chapter: 1, notes: notesWithFullSource('S4'),
      sources: sourcesLines('S4', [
        '日本麻酔科学会「麻酔器の始業点検」（2022年10月改訂第6版3） / Subrahmanyam M, Mohan S.',
        'Indian J Anaesth 2013;57:472-80 / De Cooman S, et al. J Clin Monit Comput 2015;29:491-497',
      ], short),
    });
    // 図: 管 3 本（左から 亜酸化窒素・空気・酸素。酸素が最も下流）。上端で合流して右へ「気化器へ」。
    const tubes = [
      { cx: 1.45, name: '亜酸化窒素', fy: 3.95, col: C.chipText },
      { cx: 2.9, name: '空気', fy: 4.45, col: C.chipText },
      { cx: 4.35, name: '酸素', fy: 3.35, col: C.teal },
    ];
    const tw = 0.42, ty = 2.55, th = 2.7, my = 2.1, fd = 0.3;
    tubes.forEach(t => {
      const o2 = t.name === '酸素';
      s.vline({ x: t.cx, y: my, h: ty - my, color: C.chipText, width: 3 });
      s.box({ x: t.cx - tw / 2, y: ty, w: tw, h: th, shape: 'roundRect', radius: 0.18, fill: C.bg, line: { color: o2 ? C.teal : C.gray, width: o2 ? 3 : 2 } });
      s.circle({ cx: t.cx, cy: t.fy, d: fd, fill: t.col });
    });
    const ax2 = 4.8;
    s.arrow({ x1: tubes[0].cx, y1: my, x2: ax2, y2: my, color: C.chipText, width: 3 });
    s.text('気化器へ', { x: ax2 + 0.15, y: r2(my - heightFor(1, 22, 0) / 2), w: 1.45, size: 22, valign: 'middle' });
    // 管のラベルは図の下
    const lyT = ty + th + 0.24;
    tubes.forEach(t => {
      const w = measure(t.name, 22) + 0.04, o2 = t.name === '酸素';
      s.text(t.name, { x: r2(t.cx - w / 2), y: lyT, w: r2(w), size: 22, align: 'center', bold: o2, color: o2 ? TEAL_T : C.dark });
    });
    // 右: 4 項目（各 2 行。1 行目を太字）。全角括弧で始まる行は括弧の左の空きぶん左へ寄せる
    const items = [
      ['酸素は最も下流で合流する', { t: '（漏れても酸素が残る）', dx: -0.15 }],
      ['低酸素防止装置：', '酸素の割合を一定以上に保つ'],
      ['酸素供給圧が下がると', '亜酸化窒素が止まる'],
      ['酸素フラッシュ：', '35〜75 L/分、気化器を通らない'],
    ];
    // 「気化器へ」（右端 約6.3in）と行頭記号の間を 0.5in 以上あける
    const ig = 0.16, bx = 6.9, h = itemListH(items, 24, ig);
    itemList(s, items, { x: bx, y: r2(1.85 + (6.35 - 1.85 - h) / 2), w: XR - bx, size: 24, itemGap: ig, indent: 0.36, leadBold: true });
  }

  // ===== S5 気化器（3 列カード） =====
  {
    // 出典: 3 件で 16pt・2 行に収まらない（約96字相当 > 93字相当）ため、学会資料の版・月を略す（2022年10月改訂第6版3 → 2022）。全文はノートへ。
    const short = abbreviate('S5', [['（2022年10月改訂第6版3）', '（2022）']]);
    const s = deck.slide({
      title: '気化器', chapter: 2, notes: notesWithFullSource('S5'),
      sources: sourcesLines('S5', [
        'Chakravarti S, Basu S. Indian J Anaesth 2013;57:464-471 / Kapoor MC, Vakamudi M.',
        'J Anaesthesiol Clin Pharmacol 2012;28:92-100 / 日本麻酔科学会「麻酔器の始業点検」（2022）',
      ], short),
    });
    cardRow(s, [
      { title: '可変バイパス型', titleColor: C.blue, items: [['ガスの一部を', '液の上に通す'], ['薬ごとに専用'], ['温度で補正']] },
      { title: 'デスフルラン用', titleColor: C.vermilion, items: [['沸点が低い', { t: [{ text: '（約 22.8 ℃）', bold: true, color: C.vermilion }], dx: -0.1 }], ['加温・加圧して気化'], ['電源が要る']] },
      { title: '取り扱い', titleColor: TEAL_T, items: [['傾けない'], ['同時に2つ開かない'], ['充填口は薬ごとの色']] },
    ], { y: 1.95, itemGap: 0.36, titleGap: 0.22, padY: 0.24 });
  }

  // ===== S6 呼吸回路（循環式）（左: 回路の模式図 / 右: 箇条書き4項目） =====
  {
    const s = deck.slide({ title: '呼吸回路（循環式）', chapter: 3, notes: notesOf('S6'), sources: sourcesOf('S6') });
    const LC = C.chipText, LW = 3, IN = C.blue, EX = C.vermilion;
    const x0 = 2.3, xR = 5.2, xY = 5.6, yT = 2.65, yB = 4.95, yM = (yT + yB) / 2;
    const lh = heightFor(1, 22, 0), lh2 = heightFor(2, 22, 0);
    // CO2 吸収缶（左の縦線の上。下から入って上へ抜ける）
    const abW = 0.7, abY = 3.9, abH = 0.7;
    // 環: 吸気側（吸収缶の出口 → 左上 → 上側 → Y ピース）をブルー、呼気側（Y ピース → 下側 → 吸収缶の入口）をバーミリオン
    s.polyline([[x0, abY], [x0, yT], [xR, yT], [xY, yM]], { color: IN, width: LW });
    s.polyline([[xY, yM], [xR, yB], [x0, yB], [x0, abY + abH]], { color: EX, width: LW });
    s.box({ x: x0 - abW / 2, y: abY, w: abW, h: abH, fill: C.bg, line: { color: C.teal, width: 3 } });
    s.text(['CO2', '吸収缶'], { x: r2(x0 - abW / 2 - GAP - 1.06), y: r2(abY + abH / 2 - lh2 / 2), w: 1.06, h: lh2, size: 22, align: 'right', bold: true, color: TEAL_T });
    // 患者（Y ピース）: 環の右端から右へ
    const pcx = 6.55, pd = 0.42;
    s.hline({ x: xY, y: yM, w: r2(pcx - pd / 2 - xY), color: LC, width: LW });
    s.circle({ cx: pcx, cy: yM, d: pd, fill: C.bg, line: { color: LC, width: 2.5 } });
    const pw = r2(measure('（Y ピース）', 22) + 0.02);
    s.text(['患者', '（Y ピース）'], { x: r2(pcx - pw / 2), y: r2(yM + pd / 2 + GAP), w: pw, h: lh2, size: 22, align: 'center' });
    // 吸気弁（上側、右向きの三角）。ラベルは上
    const iv = 3.4, tv = 0.22;
    s.polyline([[iv - tv, yT - tv], [iv + tv, yT], [iv - tv, yT + tv]], { color: IN, width: 1, closed: true, fill: IN });
    s.text('吸気弁', { x: iv - 0.6, y: r2(yT - tv - GAP - lh), w: 1.2, h: lh, size: 22, align: 'center', bold: true, color: IN });
    // 呼気弁（下側、左向きの三角）。ラベルは三角の上（環の内側）
    const ev = 4.45;
    s.polyline([[ev + tv, yB - tv], [ev - tv, yB], [ev + tv, yB + tv]], { color: EX, width: 1, closed: true, fill: EX });
    s.text('呼気弁', { x: ev - 0.6, y: r2(yB - tv - GAP - lh), w: 1.2, h: lh, size: 22, align: 'center', bold: true, color: EX });
    // APL 弁（下側から枝を出した小さな弁）。ラベルは四角の右
    const apl = 3.5, aplY = 5.5, aplS = 0.3;
    s.vline({ x: apl, y: yB, h: r2(aplY - yB), color: LC, width: LW });
    s.box({ x: apl - aplS / 2, y: aplY, w: aplS, h: aplS, fill: C.white, line: { color: LC, width: 2.5 } });
    const aw = r2(measure('APL 弁', 22) + 0.04);
    s.text('APL 弁', { x: r2(apl + aplS / 2 + GAP), y: r2(aplY + aplS / 2 - lh / 2), w: aw, h: lh, size: 22 });
    // バッグ（下側から吊るす）。ラベルは左
    const bag = 2.75, bagY = yB + 0.3, bagW = 0.56, bagH = 0.8;
    s.vline({ x: bag, y: yB, h: r2(bagY - yB), color: LC, width: LW });
    s.box({ x: bag - bagW / 2, y: bagY, w: bagW, h: bagH, shape: 'ellipse', fill: C.bg, line: { color: LC, width: 2.5 } });
    s.text('バッグ', { x: r2(bag - bagW / 2 - GAP - 1.06), y: r2(bagY + bagH / 2 - lh / 2), w: 1.06, h: lh, size: 22, align: 'right' });
    // 新鮮ガスの入口（左から環の吸気側へ）
    const fy = 3.05;
    s.arrow({ x1: 0.7, y1: fy, x2: x0, y2: fy, color: IN, width: 3.5 });
    s.text('新鮮ガス', { x: X0, y: r2(fy - GAP - lh), w: r2(measure('新鮮ガス', 22) + 0.04), h: lh, size: 22, bold: true, color: IN });
    // 流れの向き（山形 0.28in。吸気側ブルー・呼気側バーミリオン）
    chevron(s, 4.4, yT, 'r', IN);                       // 上側: 右へ
    chevron(s, 3.87, yB, 'l', EX);                      // 下側: 左へ（APL 弁と呼気弁の間）
    chevron(s, x0, r2((fy + abY) / 2), 'u', IN);        // 左側: 上へ（吸収缶と新鮮ガスの合流点の間）
    // 右: 4 項目（各 2 行・24pt。1 行目を太字）
    const items = [
      ['一方弁が', '吸気と呼気を分ける'],
      ['呼気は CO2 吸収缶を', '通って再利用される'],
      ['APL 弁：用手換気のときに', '圧を逃がす'],
      ['バッグ ↔ 人工呼吸器を', '切り替えて使う'],
    ];
    const bx = 7.85, ig = 0.22, h = itemListH(items, 24, ig);
    itemList(s, items, { x: bx, y: r2(1.85 + (6.6 - 1.85 - h) / 2), w: XR - bx, size: 24, itemGap: ig, leadBold: true });
  }

  // ===== S7 二酸化炭素吸収剤（3 列カード） =====
  {
    // 出典: 4 件で 16pt・2 行に収まらない（頁を略しても約25in > 2行 23.2in）ため、
    //       各文献を「著者, 誌名 年」の形に揃えて略す（巻・頁を略す）。全文はノートへ。
    const short = abbreviate('S7', [[';80:1187-1193', ''], [';96:173-182', ''], [';99:769-774', ''], [';57:516-524', '']]);
    const s = deck.slide({ title: '二酸化炭素吸収剤', chapter: 3, notes: notesWithFullSource('S7'), sources: short.split(' / ') });
    cardRow(s, [
      { title: '働き', titleColor: C.blue, items: [['呼気の CO2 を', '化学反応で吸収'], ['色の変化で消耗を知る']] },
      { title: '落とし穴', titleColor: C.red, items: [[[{ text: '乾燥すると', bold: true, color: C.red }], [{ text: 'CO・異常発熱', bold: true, color: C.red }]], ['色が戻ることがある'], ['時間でも管理する']] },
      { title: '交換', titleColor: TEAL_T, items: [['基線が上がったら交換'], ['交換したら回路の', 'リークテスト'], ['在庫と記録']] },
    ], { y: 1.95, itemGap: 0.32, titleGap: 0.22, padY: 0.22 });
  }

  // ===== S8 新鮮ガス流量と低流量麻酔（左: 2 段の横棒 / 右: 箇条書き4項目） =====
  {
    const s = deck.slide({ title: '新鮮ガス流量と低流量麻酔', chapter: 3, notes: notesOf('S8'), sources: sourcesOf('S8') });
    // 図: 上段「高流量」は設定と吸入がほぼ同じ、下段「低流量」は吸入が設定より短い。ラベルは棒の左。
    const lh = heightFor(1, 22, 0);
    const gx = X0, gw = 1.1;               // 段のラベル
    const bx = 1.88, bw = 0.72;            // 棒のラベル
    const x0 = bx + bw + GAP, full = 5.6 - x0, barH = 0.44, rowGap = 0.26;
    const groups = [
      { name: '高流量', y: 2.45, ratio: 0.95 },
      { name: '低流量', y: 4.45, ratio: 0.6 },
    ];
    groups.forEach(g => {
      const bars = [
        { lab: '設定', len: full, col: C.blue },
        { lab: '吸入', len: full * g.ratio, col: C.vermilion },
      ];
      bars.forEach((b, i) => {
        const y = g.y + i * (barH + rowGap);
        s.box({ x: r2(x0), y, w: r2(b.len), h: barH, fill: b.col });
        s.text(b.lab, { x: bx, y: r2(y + barH / 2 - lh / 2), w: bw, h: lh, size: 22, align: 'right', bold: true, color: b.col });
      });
      // 設定値の位置を示す破線（設定の棒の右端から少し離し、棒の上下にも少し伸ばす）
      s.vline({ x: r2(x0 + full + 0.07), y: r2(g.y - 0.18), h: r2(2 * barH + rowGap + 0.36), color: C.chipText, width: 1.5, dash: 'dash' });
      const gh = 2 * barH + rowGap;
      s.text(g.name, { x: gx, y: r2(g.y + gh / 2 - lh / 2), w: gw, h: lh, size: 22, bold: true });
    });
    const items = [
      ['流量が多いほど、', '吸入濃度は設定値に近づく'],
      ['低流量では再呼吸が増え、', '設定と吸入濃度がずれる'],
      ['酸素濃度計と呼気ガス濃度で', '実際の値を見る'],
      ['低流量の利点：', '薬の節約・加湿・環境負荷の低減'],
    ];
    const ig = 0.22, h = itemListH(items, 24, ig);
    itemList(s, items, { x: 6.35, y: r2(1.9 + (6.6 - 1.9 - h) / 2), w: XR - 6.35, size: 24, itemGap: ig, leadBold: true });
  }

  // ===== S9 安全装置と余剰ガス（3 列カード） =====
  {
    const s = deck.slide({ title: '安全装置と余剰ガス', chapter: 4, notes: notesOf('S9'), sources: sourcesOf('S9') });
    // 「用手換気はできる機種か」（22pt で 3.76in）を 1 行に入れるため 3 枚目だけ 4.05in（内幅 3.81in）、1・2 枚目は 3.85in。
    // 等幅 3.91in では 3 行に割れ、行グリッドで 2 行目が 3 行分の高さになる。
    cardRow(s, [
      { title: '必ず見るもの', titleColor: TEAL_T, items: [['吸気酸素濃度計'], ['気道内圧・', '回路外れアラーム'], ['呼気ガス濃度']] },
      { title: '余剰ガス排除', titleColor: C.blue, items: [['吸引圧 1 kPa 以上', '2 kPa 未満'], ['または 25〜50 L/分'], ['接続を毎朝確認']] },
      { title: '電源', titleColor: C.vermilion, items: [['停電時の', 'バッテリー時間を知る'], ['バッテリーが切れても', '用手換気はできる機種か']] },
    ], { ws: [3.85, 3.85, 4.05], y: 1.95, itemGap: 0.3, titleGap: 0.18, padY: 0.22 });
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const s = deck.slide({
      title: 'まとめ', chapter: 4, notes: notesOf('S10'),
      // 1 項目（学会資料＋URL）が 16pt・1 行に入らないため、間の空白で 2 行に分ける
      sources: sourcesLines('S10', [
        '日本麻酔科学会「麻酔器の始業点検」（2022年10月改訂第6版3）',
        'https://anesth.or.jp/files/pdf/guideline_checkout_20221117.pdf',
      ]),
    });
    itemList(s, [
      ['配管・ボンベ → 流量計 → 気化器 → 回路 → 患者'],
      ['酸素は最も下流で合流。フラッシュは気化器を通らない'],
      ['一方弁と吸収剤で呼気を再利用する循環式'],
      [[{ text: '設定より実測', bold: true, color: TEAL_T }, '：酸素濃度計と呼気ガス濃度']],
      [[{ text: '乾いた吸収剤は使わない', bold: true, color: C.red }, '。基線が上がったら交換']],
    ], { x: X0, y: 1.98, w: BW, size: 26, itemGap: 0.32 });
  }

  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.render) console.log('[build] render: ' + res.render.images.length + ' images in ' + RENDER_DIR);
  if (res.violations !== 0) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
