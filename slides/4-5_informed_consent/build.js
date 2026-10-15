'use strict';
/**
 * 4-5 麻酔の説明と同意
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk()/q() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。収まらずに短くした箇所は cut()/srcFix() を通し、最後に一覧を出す。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fitInfo, heightFor, measure } = require('../_lib/decklib');

const NAME = '4-5_informed_consent';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
const SPEC_MD = fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8');
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const lines = part.split('\n');
    let notes = null, sources = null;
    const quoted = new Set();
    let inOn = false;
    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i];
      if (/^notes:\s*\|\s*$/.test(ln)) {
        const buf = [];
        for (let j = i + 1; j < lines.length && (lines[j].startsWith('  ') || lines[j] === ''); j++) {
          buf.push(lines[j].replace(/^ {2}/, ''));
        }
        while (buf.length && buf[buf.length - 1] === '') buf.pop();
        notes = buf.join('\n');
      }
      const m = ln.match(/^sources:\s*(.*)$/);
      if (m) sources = m[1].trim();
      if (/^on-slide:/.test(ln)) { inOn = true; continue; }
      if (inOn && !/^ {2}- /.test(ln)) inOn = false;
      if (inOn) {
        // 「…」の中身（入れ子の『』はそのまま）
        for (const qq of ln.matchAll(/「([^」]*)」/g)) quoted.add(qq[1]);
      }
    }
    out[id] = { notes, sources: sources && sources !== '—' ? sources : null, quoted };
  }
  return out;
}
const SPEC = parseSpec(SPEC_MD);
function notesOf(id) {
  const n = SPEC[id] && SPEC[id].notes;
  if (!n) throw new Error(`spec.md の ${id} に notes がありません`);
  return n;
}
/** 出典: spec の sources 行を ' / ' 区切りの項目にしてライブラリへ（ライブラリが詰めて最大2行） */
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  return s ? s.split(' / ') : undefined;
}
/** spec の文言そのまま（照合のみ） */
function q(id, str) {
  if (!SPEC[id].quoted.has(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
  return str;
}
/** str を breaks の各語の直前で改行して行配列にする（文言は変えない。改行位置の半角スペースは落とす） */
function split(id, str, breaks) {
  const lines = [];
  let from = 0;
  for (const b of breaks) {
    const k = str.indexOf(b, from + 1);
    if (k < 0) throw new Error(`${id}: 「${str}」に改行位置「${b}」が見つかりません`);
    lines.push(str.slice(from, k).replace(/ $/, ''));
    from = k;
  }
  lines.push(str.slice(from));
  return lines;
}
function brk(id, str, ...breaks) { return split(id, q(id, str), breaks); }

// ---------------------------------------------------------------------------
// spec から変えた箇所（どうしても収まらないときだけ）。最後に一覧を出す。
// ---------------------------------------------------------------------------
const DEVIATIONS = [];
/** スライド面の文言を短くする（元の文言が spec にあることを確かめ、記録する） */
function cut(id, orig, after, why) {
  q(id, orig);
  DEVIATIONS.push(`${id} 本文: 「${orig}」→「${after}」（${why}）`);
  return after;
}
/** 出典の項目を差し替える。fixes: [[元の項目, 後の項目 | null(落とす)], ...] */
function srcFix(id, fixes, why) {
  const items = sourcesOf(id).slice();
  for (const [orig, after] of fixes) {
    const k = items.indexOf(orig);
    if (k < 0) throw new Error(`${id}: 出典「${orig}」が spec の sources にありません`);
    if (after == null) items.splice(k, 1);
    else items[k] = after;
    DEVIATIONS.push(`${id} 出典: 「${orig}」→${after == null ? '（スライド面から外す。ノートの参考文献には残る）' : '「' + after + '」'}（${why}）`);
  }
  return items;
}

/**
 * 1行の文字列に部分装飾を付けて run 配列にする。marks: [[語, 色], ...]（太字）。
 * 語は行の中に1回だけ現れること。
 */
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
      next.push({ text: w, color, bold: true });
      if (k + w.length < r.length) next.push(r.slice(k + w.length));
    }
    if (found !== 1) throw new Error(`mark: 「${line}」に「${w}」が ${found} 回`);
    runs = next;
  }
  return runs;
}
const runText = l => (Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(l));

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const Y0 = LAYOUT.body.y;        // 1.85
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
/** 出典 2 行（枠 y6.5）のスライドの本文下端。decklib の警告線（6.47）より上 */
const BOT2 = 6.46;
const r3 = v => Math.round(v * 1000) / 1000;
const r2 = v => Math.round(v * 100) / 100;
const CM = 2.54;
/** 強調ボックス（位置・幅・線をスライド間で揃える） */
const EMPH = { x: X0 + 0.6, w: BW - 1.2, radius: 0.1, lineW: 3.5 };

// ---------------------------------------------------------------------------
// 行送りを詰めたテキスト（カード・強調ボックス用）
// 本文 22pt のまま、カード 2 枚＋強調ボックスを本文ゾーンに入れるため、行送りを 1.2 em に詰める
// （tightTable と同じ行送り）。PowerPoint のメイリオの 1 行（約 1.5 em）× 0.8 = 1.2 em。
// 行間は倍数（spcPct 80%）で入れる。固定値（spcPts）だと slide_lint の高さ推定が壊れるため。
// slide_lint の推定（行 = pt × 0.8 × 1.18、段落後 6pt）と同じ式で枠の高さを確かめる。
// ---------------------------------------------------------------------------
const TIGHT = 0.8;
const lineIn = size => size * 1.2 / 72;
const tightH = (n, size) => Math.ceil(n * lineIn(size) * 1000) / 1000;
const lintEst = (n, size) => n * (size / 72 * TIGHT * 1.18 + 6 / 72);

function tightText(s, lines, o) {
  const size = o.size || 22;
  if (size < 22) throw new Error(`p${s.page}: 本文は 22pt 以上（${size}pt）`);
  const norm = lines.map(l => (Array.isArray(l) ? l : [l]).map(r => (typeof r === 'string' ? { text: r } : r)));
  norm.forEach(runs => {
    const t = runs.map(r => r.text).join('');
    const fi = fitInfo(t, size, o.w);
    if (!fi.ok) throw new Error(`p${s.page}「${t}」が幅 ${r3(o.w)}in に収まりません（推定 ${fi.needIn}in、約${fi.overChars}字オーバー）`);
  });
  const h = o.h != null ? o.h : tightH(norm.length, size);
  if (h + 1e-9 < norm.length * lineIn(size)) throw new Error(`p${s.page}: 枠高 ${h}in に ${norm.length} 行が入りません`);
  if (lintEst(norm.length, size) > h + 0.3 / CM + 1e-9) throw new Error(`p${s.page}: lint 推定の高さ ${r3(lintEst(norm.length, size))}in > 枠高 ${h}in + 0.12in`);
  const where = `p${s.page} tight「${runText(norm[0].map(r => r.text))}」`;
  s._bounds(where, o.x, o.y, o.w, h);
  const out = [];
  norm.forEach((runs, li) => {
    runs.forEach((r, ri) => {
      const op = { fontFace: FONT, lang: 'ja-JP', fontSize: size, color: r.color || o.color || C.dark };
      if (r.bold != null ? r.bold : o.bold) op.bold = true;
      if (ri === runs.length - 1 && li < norm.length - 1) op.breakLine = true;
      out.push({ text: r.text, options: op });
    });
  });
  s.raw.addText(out, {
    x: o.x, y: o.y, w: o.w, h, margin: 0, isTextBox: true,
    fontFace: FONT, lang: 'ja-JP', fontSize: size,
    align: o.align || 'left', valign: o.valign || 'top', lineSpacingMultiple: TIGHT,
  });
  s._reg('txt', o.x, o.y, o.w, h, where);
  return h;
}

// ---------------------------------------------------------------------------
// カード（S3・S4・S6）: 角丸の塗り箱＋見出し（24pt 太字）＋項目（22pt）。文字はすべて箱の内側。
// 幅: 各カードの最長行に左右の余白を足し、余りを等分。高さ: y から bottom まで（左右で揃える）。
// 項目の間隔: 最も詰まったカードに合わせて余りを配る（上限 maxItemGap）。
// ---------------------------------------------------------------------------
const CARD = { padY: 0.1, titleSize: 24, size: 22, titleGap: 0.05, radius: 0.08, maxPadX: 0.2, minPadX: 0.08 };

function cardRow(s, cards, o) {
  const n = cards.length;
  const need = cards.map(c => Math.max(measure(c.title, CARD.titleSize), ...c.items.flat().map(l => measure(runText(l), CARD.size))));
  const extra = BW - o.gap * (n - 1) - need.reduce((a, b) => a + b, 0);
  const padX = Math.min(CARD.maxPadX, extra / (2 * n));
  if (padX < CARD.minPadX - 1e-9) throw new Error(`p${s.page}: カードの左右余白が ${r3(padX)}in（${CARD.minPadX}in 未満）`);
  const slack = (extra - 2 * n * padX) / n;   // 余りは各カードに等分（内側の幅が広がる）
  const h = r3(o.bottom - o.y);
  const fixed = 2 * CARD.padY + tightH(1, CARD.titleSize) + CARD.titleGap;
  const linesH = c => c.items.reduce((a, it) => a + tightH(it.length, CARD.size), 0);
  let g = Math.min(o.maxItemGap, ...cards.map(c => (h - fixed - linesH(c)) / (c.items.length - 1)));
  g = Math.floor(g * 1000) / 1000;
  if (g < 0.06 - 1e-9) throw new Error(`p${s.page}: 項目の間隔が ${r3(g)}in（0.06in 未満）。高さ不足`);
  let x = X0;
  const info = [];
  cards.forEach((c, i) => {
    const w = r3(need[i] + 2 * padX + slack);
    const iw = r3(w - 2 * padX), ix = r3(x + padX);
    s.box({ x: r3(x), y: o.y, w, h, fill: C.bg, radius: CARD.radius });
    let y = o.y + CARD.padY;
    tightText(s, [c.title], { x: ix, y: r3(y), w: iw, size: CARD.titleSize, bold: true, color: c.titleColor });
    y += tightH(1, CARD.titleSize) + CARD.titleGap;
    c.items.forEach(lines => {
      tightText(s, lines, { x: ix, y: r3(y), w: iw, size: CARD.size });
      y += tightH(lines.length, CARD.size) + g;
    });
    const end = y - g;
    if (end > o.y + h - CARD.padY + 1e-6) throw new Error(`p${s.page} カード「${c.title}」: 本文の下端 ${r3(end)}in > ${r3(o.y + h - CARD.padY)}in`);
    info.push(`${w}in（内側 ${iw}in・最長行 ${r3(need[i])}in）`);
    x += w + o.gap;
  });
  console.log(`  p${s.page} カード: 幅 ${info.join(' / ')}、左右余白 ${r3(padX)}in、高さ ${h}in（下端 ${r3(o.y + h)}in）、項目間 ${r3(g)}in`);
}

/** 強調ボックス（行送りを詰めた版。22pt 太字・白地＋色枠・中央揃え）。下端 bottom に合わせて置き、上端を返す */
function emphTight(s, lines, o) {
  const size = o.size || 22, padY = 0.05;
  const th = tightH(lines.length, size);
  const h = r3(th + 2 * padY);
  const y = r3(o.bottom - h);
  s.box({ x: EMPH.x, y, w: EMPH.w, h, fill: C.white, radius: EMPH.radius, line: { color: o.lineColor || C.teal, width: EMPH.lineW } });
  tightText(s, lines, { x: EMPH.x + 0.15, y: r3(y + padY), w: EMPH.w - 0.3, h: th, size, bold: true, align: 'center' });
  return y;
}

/** 強調ボックス（decklib の box。28pt 太字・白地＋色枠）。y は上端 */
function emphBox(s, y, h, lines, lineColor) {
  s.box({
    x: EMPH.x, y: r3(y), w: EMPH.w, h, fill: C.white, radius: EMPH.radius,
    line: { color: lineColor || C.teal, width: EMPH.lineW },
    text: lines, size: 28, bold: true, color: C.dark, margin: 0.05,
  });
}

// ---------------------------------------------------------------------------
// 行の多い表（S5）。3-11 の tightTable と同じ（行間 1.2 em・22pt。見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、上下中央）。
// s.table は行高を実機メイリオの 1.5 em で取るため、見出し＋7 行の表が本文ゾーンに入らない。
// セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、重なり検査用の登録も同じく行う。
// cell: 文字列（'\n' で改行）| {text, color, bold, align}
// ---------------------------------------------------------------------------
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
    chapters: ['法と\n判例', '何を\n話すか', 'リスク\nの数', '能力と\n代諾', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-5 麻酔の説明と同意',
  });
  const slide = o => deck.slide(o);

  // ===== S1 表紙 =====
  {
    const subtitle = '4-5 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: '麻酔の説明と同意', subtitle, meta, notes: notesOf('S1') });
  }

  // 箇条書き＋強調ボックスのスライド（S2・S7）: 24pt・段落後 10pt・y1.85、強調ボックスは 28pt 2 行で出典 2 行の上
  const BUL = { y: Y0, size: 24, gap: 10 };
  const bulH = n => heightFor(n, BUL.size, BUL.gap);
  const EMPH28_H = 1.25;

  // ===== S2 説明義務：法律と判例（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S2';
    const s = slide({ title: '説明義務：法律と判例', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, '医療法：適切な説明を行い、理解を得るよう努める（1997 年〜）'), ['医療法：', C.blue]),
      mark(q(id, '民法（判例）：診療契約上の義務。怠れば債務不履行・不法行為'), ['民法（判例）：', C.blue]),
      mark(q(id, '最高裁 2001 年：診断、内容、危険性、代替と利害得失、予後'), ['最高裁 2001 年：', C.blue]),
      mark(q(id, '最高裁 2000 年：輸血方針を告げず手術→意思決定の権利を奪い賠償'), ['最高裁 2000 年：', C.blue]),
      mark(q(id, '麻酔科医の説明義務違反で賠償を命じた下級審の例もある'), ['麻酔科医', C.blue]),
    ], BUL);
    const by = BOT2 - EMPH28_H;
    if (by < BUL.y + bulH(5) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(5))}in`);
    const lines = brk(id, '悪い結果が起きなくても、説明しなかったこと自体が責任を問われる', '説明しなかった');
    emphBox(s, by, EMPH28_H, [lines[0], mark(lines[1], ['説明しなかったこと自体', C.vermilion])], C.teal);
  }

  // ===== S3 誰が、いつ説明するか（2列カード + 強調ボックス） =====
  {
    const id = 'S3';
    const s = slide({ title: '誰が、いつ説明するか', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    const ey = emphTight(s, brk(id, '一人麻酔科医でも、外科医の説明だけで済ませない。術前に診察して補足説明する手順を決めておく', '術前に'),
      { bottom: BOT2, lineColor: C.teal });
    cardRow(s, [
      {
        title: q(id, '誰が'), titleColor: C.blue,
        items: [
          brk(id, 'JSA 認定施設の 96.7% で麻酔科医が説明（2005 年調査）', '麻酔科医が'),
          [cut(id, '説明の責任は担当する麻酔科医（AAGBI 2017）', '説明の責任は担当する麻酔科医',
            '2 列カード＋強調ボックス＋出典 2 行が 22pt で本文ゾーンに入らないため。出典表記は同じスライドの出典行（Yentis 2017, AAGBI）に残る')],
          brk(id, '外科医の説明だけでは麻酔固有のリスクが抜けやすい', '麻酔固有'),
          [q(id, '動画は補助。責任は麻酔科医に残る')],
          [q(id, '研修医は自分の立場を伝える')],
        ],
      },
      {
        title: q(id, 'いつ'), titleColor: C.vermilion,
        items: [
          brk(id, '待機手術は前日まで。当日・導入直前に初めて伝えない', '当日'),
          brk(id, '麻酔管理料（Ⅰ）の要件：術前診察は当日以外（緊急を除く）', '術前診察'),
          brk(id, '手元に残る資料を渡し、質問の時間をとる', '質問'),
          [q(id, '同意は 1 回の出来事ではなく過程')],
        ],
      },
    ], { y: Y0, bottom: r3(ey - GAP), gap: GAP, maxItemGap: 0.22 });
  }

  // ===== S4 何を話すか（2列カード + 強調ボックス） =====
  {
    const id = 'S4';
    const s = slide({ title: '何を話すか', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    const ey = emphTight(s, [q(id, 'チェックリストを読み上げない。その患者の状態でのリスクに合わせて話す')],
      { bottom: BOT2, lineColor: C.teal });
    const risk1 = brk(id, 'よくある：吐き気、咽頭痛、歯・唇の損傷、しびれ、頭痛（脊麻）', '歯・唇');
    const risk2 = brk(id, '重大：神経・眼の障害、術中覚醒、アナフィラキシー、死亡', 'アナフィラキシー');
    cardRow(s, [
      {
        title: q(id, '麻酔の流れ'), titleColor: C.blue,
        items: [
          [q(id, '絶飲食、前投薬、ルート、モニター')],
          [q(id, '導入（マスク、点滴、挿管）')],
          [q(id, '術後の鎮痛と吐き気の対策')],
          brk(id, '代替手段と、うまくいかないときの次の手（脊麻が効かなければ全身麻酔）', 'うまくいかない', '（脊麻'),
        ],
      },
      {
        title: q(id, 'リスク'), titleColor: C.vermilion,
        items: [
          [mark(risk1[0], ['よくある：', C.dark]), risk1[1]],
          [mark(risk2[0], ['重大：', C.red]), risk2[1]],
          brk(id, 'この患者が重視するリスク（歯、声、記憶）を聞く', '（歯'),
          [q(id, '質問と、それへの答えも記録')],
        ],
      },
    ], { y: Y0, bottom: r3(ey - GAP), gap: GAP, maxItemGap: 0.22 });
  }

  // ===== S5 リスクを数字で伝える（表 3列 × 7行＋見出し行） =====
  // 22pt・行間 1.2 em のまま。2 行になる行が 4 つ（歯・硬膜穿刺後頭痛・術中覚醒・死亡）あり、
  // tightTable 既定の上下余白 3pt では 5.08in と本文ゾーン（1.85〜6.46in）を超えるため、上下余白だけ 0.5pt にする。
  // 列幅は spec の目安（3.6/4.6/4.0）から、1 列目を最長の 13 字、3 列目を「Higgins 2002」が 1 行に入る幅にし、残りを 2 列目へ。
  {
    const id = 'S5';
    const sources = srcFix(id, [['Cook TM, et al. Br J Anaesth 2009;102:179-190（NAP3）', null]],
      '4 件は空白を詰め（）の補足を外しても 16pt で 2 行に収まらない。表の出典列に「NAP3」、ノートに全文が残る');
    const s = slide({ title: 'リスクを数字で伝える', chapter: 2, notes: notesOf(id), sources });
    const cell = str => q(id, str);
    const cellB = (str, ...breaks) => brk(id, str, ...breaks).join('\n');
    // 見出し行は spec の layout 行（列「事象」「頻度の目安」「出典」）
    const HEAD = ['事象', '頻度の目安', '出典'];
    if (!SPEC_MD.includes('列' + HEAD.map(h => `「${h}」`).join(''))) throw new Error(`${id}: 表の見出しが spec と違います`);
    const t = tightTable(s, [
      HEAD,
      [cell('術後の咽頭痛'), cell('挿管 45.4%、LMA 17.5%'), cell('Higgins 2002')],
      [cell('吐き気・嘔吐'), cell('10〜79%（危険因子 0〜4 個で）'), cell('Apfel 1999')],
      [cell('歯の損傷（処置が要るもの）'), cellB('約 1/4,500。歯が悪いとオッズ比 50', '歯が'), cell('Warner 1999')],
      [cell('硬膜穿刺後頭痛（産婦）'), cellB('硬膜外の偶発穿刺 1.5%、うち 52.1%。27G 脊麻 1.7%', 'うち'), cell('Choi 2003')],
      [cell('術中覚醒'), cellB('約 1/19,600（筋弛緩薬使用 1/8,200、帝王切開 1/670）', '1/8,200'), cell('NAP5')],
      [cell('脊麻・硬膜外の永続的な障害'), cell('10 万件に 2.0〜4.2 件'), cell('NAP3')],
      [cell('麻酔が単独の原因の死亡'), cellB('1 万麻酔に 0.07〜0.21（国内 1994〜2005 年）', '（国内'), cellB('JSA 偶発症調査', '偶発')],
    ], { x: X0, y: Y0, colW: [4.62, 5.2, 2.38], margin: [0.5, 6, 0.5, 6] });
    console.log(`  p${s.page} 表: 下端 ${r3(t.y + t.h)}in、行高 ${t.rowH.join(' / ')}`);
    if (t.y + t.h > BOT2 + 1e-6) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)}in > ${BOT2}in`);
  }

  // ===== S6 意思決定能力と代諾（2列カード + 強調ボックス） =====
  {
    const id = 'S6';
    const sources = srcFix(id, [
      ['宗教的輸血拒否に関する合同委員会「宗教的輸血拒否に関するガイドライン」（2008）', '宗教的輸血拒否に関する合同委員会 GL（2008）'],
      ['厚生労働省「身寄りがない人の入院及び医療に係る意思決定が困難な人への支援に関するガイドライン」（2019）',
        '厚生労働省「身寄りがない人の入院及び医療に係る意思決定が困難な人への支援に関するGL」（2019）'],
    ], '厚労省 GL の項目だけで 16pt の 1 行（全角 46 字）を超え、3 件が 2 行に収まらない。「ガイドライン」を本文と同じ略記「GL」にした');
    const s = slide({ title: '意思決定能力と代諾', chapter: 3, notes: notesOf(id), sources });
    const ey = emphTight(s, brk(id, '緊急で同意がとれない：推定的同意で行い、緊急性・省いた理由・事後の説明を記録する', '緊急性'),
      { bottom: BOT2, lineColor: C.teal });
    const why6 = '2 列カード＋強調ボックス＋出典 2 行が 22pt で本文ゾーンに入らないため';
    cardRow(s, [
      {
        title: q(id, '能力の 4 要素'), titleColor: C.blue,
        items: [
          [q(id, '選択を表明できる')],
          [q(id, '情報を理解できる')],
          [q(id, '自分のこととして認識できる')],
          [q(id, '比べて論理的に考えられる')],
          brk(id, '成人は能力があると推定。認知症・高齢・前投薬だけで無能力と決めない', '認知症', '無能力'),
        ],
      },
      {
        title: q(id, '代わりに決める人'), titleColor: C.vermilion,
        items: [
          [q(id, '18 歳以上：本人（2022 年〜成年）')],
          brk(id, '18 歳未満：親権者の同意（離婚後の共同親権 2026 年 4 月〜）', '（離婚後'),
          [cut(id, '15〜17 歳は本人の意思も確認（輸血拒否 GL の区分）', '15〜17 歳は本人の意思も確認（輸血拒否 GL）', why6)],
          brk(id, '成年後見人に医療同意権はない。家族の同意の法的根拠も明確でない', '家族の'),
          [cut(id, '身寄りがない人：チームと倫理委員会で（厚労省 GL 2019）', '身寄りがない人：チームと倫理委員会で',
            why6 + '。出典表記は同じスライドの出典行（厚生労働省 …GL（2019））に残る')],
        ],
      },
    ], { y: Y0, bottom: r3(ey - GAP), gap: GAP, maxItemGap: 0.22 });
  }

  // ===== S7 記録と、離島での工夫（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S7';
    const sources = srcFix(id, [['最高裁判所 平成 12 年 2 月 29 日 第三小法廷判決', '最高裁判所 平成12年2月29日 第三小法廷判決']],
      '3 件が 16pt で 2 行に収まらない（2 行目が全角 0.95 字超過）。日付の中の空白だけを詰めた');
    const s = slide({ title: '記録と、離島での工夫', chapter: 3, notes: notesOf(id), sources });
    s.bullets([
      mark(q(id, '記録：説明したリスク・代替、質問と回答、同意、同席者'), ['記録：', C.blue]),
      mark(q(id, '同意書の署名は『同意の過程があった証拠』にすぎない'), ['同意書の署名', C.blue]),
      mark(q(id, '輸血拒否：施設の方針を前もって伝える。対応できなければ転院'), ['輸血拒否：', C.blue]),
      mark(q(id, '外科医に任せるときも、麻酔科医が術前に診察して補足する'), ['外科医に任せるとき', C.blue]),
      mark(q(id, '電話・オンラインは補助。当日前に対面で理解を確かめる（私見）'), ['電話・オンライン', C.blue]),
    ], BUL);
    const by = BOT2 - EMPH28_H;
    if (by < BUL.y + bulH(5) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(5))}in`);
    const lines = brk(id, '患者の不安（『前の麻酔で…』）は、根拠が薄くても説明の対象', '根拠が');
    emphBox(s, by, EMPH28_H, [lines[0], mark(lines[1], ['根拠が薄くても', C.vermilion])], C.teal);
  }

  // ===== S8 まとめ（箇条書き5行） =====
  // 最長行（説明するのは…）が 26pt では 1 行に入らないため 24pt（1 バレット 1 行を優先）
  {
    const id = 'S8';
    const sources = srcFix(id, [
      ['最高裁判所 平成 13 年 11 月 27 日 第三小法廷判決', '最高裁判所 平成13年11月27日 第三小法廷判決'],
      ['Yentis SM, et al. Anaesthesia 2017;72:93-105（AAGBI）', 'Yentis SM, et al. Anaesthesia 2017;72:93-105'],
    ], '3 件が 16pt で 2 行に収まらない。日付の中の空白を詰め、（）の補足を外した');
    const s = slide({ title: 'まとめ', chapter: 4, notes: notesOf(id), sources });
    s.bullets([
      mark(q(id, '説明義務は法律と判例の要請。診断・内容・危険・代替・予後'), ['診断・内容・危険・代替・予後', C.blue]),
      mark(q(id, '説明するのは担当の麻酔科医。待機手術は前日までに、資料を渡して'), ['担当の麻酔科医', C.blue]),
      mark(q(id, '覚醒 1/8,200（筋弛緩薬あり）、麻酔単独死亡 1 万に 0.07〜0.21'), ['1/8,200', C.vermilion], ['0.07〜0.21', C.vermilion]),
      mark(q(id, '能力は 4 要素で決定ごと。18 歳で成人。未成年にも年齢相応に説明'), ['4 要素', C.blue]),
      mark(q(id, '署名は証拠にすぎない。説明の中身と質問への答えを記録する'), ['説明の中身と質問への答え', C.blue]),
    ], { y: 1.95, size: 24, gap: 24 });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  console.log('\n[deviations]');
  DEVIATIONS.forEach(d => console.log('  - ' + d));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
