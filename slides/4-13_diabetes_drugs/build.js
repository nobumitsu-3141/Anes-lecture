'use strict';
/**
 * 4-13 常用薬②：糖尿病薬
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk()/q() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。spec から文言を変えた箇所は alt() で照合・記録し、ビルド時に一覧を出す。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら scratchpad の render/<NAME>）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fitInfo, heightFor } = require('../_lib/decklib');

const NAME = '4-13_diabetes_drugs';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR ||
  path.join('/tmp/claude-0/-home-user-Anes-lecture/0b714c0d-c11e-535a-a97c-0e6ce9d53e16/scratchpad/render', NAME);

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
      if (inOn) for (const qq of ln.matchAll(/「([^」]*)」/g)) quoted.add(qq[1]);
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
/**
 * spec の文言 str を、breaks の各語の直前で改行して行配列にする（文言は変えない）。
 * 改行位置の半角スペースは落とす。
 */
function brk(id, str, ...breaks) {
  q(id, str);
  return split(str, breaks, id);
}
function split(str, breaks, id) {
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
/** spec の文言を収まりのため短くした箇所（spec 側の元の文言を照合し、ビルド時に一覧を出す） */
const DEVIATIONS = [];
function alt(id, orig, shown, why) {
  q(id, orig);
  DEVIATIONS.push(`${id}: 「${orig}」→「${shown}」（${why}）`);
  return shown;
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
const r3 = v => Math.round(v * 1000) / 1000;
/**
 * 本文の下端。全スライドで出典が 2 行（枠 y6.5・下揃え。PowerPoint のメイリオ 16pt 2 行で文字の上端 約6.68in）
 * なので、その 0.24in 上。
 */
const BOTTOM = 6.44;
/** 箱で始まるスライド（S3・S6 のカード、S4 の表、S7 の指示例）の上端。金下線（8pt、下端 約1.576in）から 0.24in */
const TOP_BOX = 1.82;

// ---------------------------------------------------------------------------
// 行送りを詰めたテキスト（S3・S6 のカードと強調ボックス、S7 の指示例）
// 3-11 の tightTable と同じ考え方で、22pt 以上のまま行送りを 1.2 em にする。
// 段落の行間（spcPct / spcPts）は使わず、1 行ずつ decklib の s.text（上揃え）を 1.2 em おきに置く。
//  - spcPts の固定値は slide_lint の高さ推定を壊し、spcPct は代替フォント（LibreOffice の Inter）で
//    行送りが変わる。1 行 1 枠なら行の位置が枠の y で決まり、PowerPoint でも描画確認でも同じ行送りになる。
//  - 枠の高さは decklib / lint の 1 行の下限（22pt 0.41in、24pt 0.46in）で、次の行の枠と少し重なる
//    （文字どうしの枠の重なりは lint・decklib とも対象外）。
// ---------------------------------------------------------------------------
const LS_EM = 1.2;
const lineIn = size => size * LS_EM / 72;            // 行送り（22pt 0.367in、24pt 0.4in）
const BOX_H = { 22: 0.41, 24: 0.46 };                // 1 行の枠の高さ
const overhang = size => BOX_H[size] - lineIn(size); // 最終行の枠が行送りからはみ出す分

function tightText(s, lines, o) {
  const size = o.size || 22;
  if (!BOX_H[size]) throw new Error(`tightText: size ${size} は 22 / 24 のみ`);
  lines.forEach((line, i) => {
    s.text([line], {
      x: r3(o.x), y: r3(o.y + i * lineIn(size)), w: r3(o.w), h: BOX_H[size],
      size, bold: !!o.bold, color: o.color || C.dark, align: o.align || 'left',
    });
  });
  return lines.length * lineIn(size);
}

// ---------------------------------------------------------------------------
// カード（S3・S6）: 角丸の塗り箱＋見出し＋「•」付きの項目。文字はすべて箱の内側。
// 項目は語句の切れ目で 2 行に割り、2 行目は記号の右にぶら下げる。
// 余白は上 padT・下 padB。見出し（24pt）は枠の上に行間の余りがあるので、見た目の上下余白が揃うよう下を厚くする。
// ---------------------------------------------------------------------------
const CARD = { padX: 0.1, padT: 0.06, padB: 0.11, titleSize: 24, size: 22, titleGap: 0.04, ind: 0.28 };

function cardNeedW(c) {
  const t = fitInfo(runText(c.title), CARD.titleSize, 100).needIn;
  const b = Math.max(...c.items.flat().map(l => fitInfo(runText(l), CARD.size, 100).needIn)) + CARD.ind;
  return Math.max(t, b) + 2 * CARD.padX;
}
function cardFixedH(c) {
  const nLines = c.items.reduce((a, it) => a + it.length, 0);
  return CARD.padT + CARD.padB + lineIn(CARD.titleSize) + CARD.titleGap + nLines * lineIn(CARD.size);
}
function drawCard(s, c, x, y, w, h, itemGap) {
  s.box({ x: r3(x), y: r3(y), w: r3(w), h: r3(h), fill: C.bg, radius: 0.08 });
  const ix = x + CARD.padX, iw = w - 2 * CARD.padX;
  tightText(s, [c.title], { x: ix, y: y + CARD.padT, w: iw, size: CARD.titleSize, bold: true, color: c.titleColor });
  let yy = y + CARD.padT + lineIn(CARD.titleSize) + CARD.titleGap;
  c.items.forEach(lines => {
    // 記号の枠は推定器（全角 1 字 = 0.342in）に合わせて広めに取り、本文の枠と重ねる（文字どうしの枠の重なりは無害）
    tightText(s, ['•'], { x: ix, y: yy, w: 0.35, size: CARD.size });
    tightText(s, lines, { x: ix + CARD.ind, y: yy, w: iw - CARD.ind, size: CARD.size });
    yy += lines.length * lineIn(CARD.size) + itemGap;
  });
  const end = yy - itemGap;
  if (end > y + h - CARD.padB + 1e-6 || end + overhang(CARD.size) > y + h + 1e-6) {
    throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さ不足（本文の下端 ${r3(end)} > ${r3(y + h - CARD.padB)}）`);
  }
  return r3(y + h - end);   // 最終行の行送りの下端からカード下端まで
}
/**
 * 2 列のカード。幅は各カードの最長行に必要な幅を確保し、余りを等分。高さは y〜bottom。
 * 項目の間隔は、詰まった方のカードに合わせて余りを配る（上限 maxGap）。
 */
function cardRow(s, cards, o) {
  const need = cards.map(cardNeedW);
  const avail = BW - GAP * (cards.length - 1);
  const extra = avail - need.reduce((a, b) => a + b, 0);
  if (extra < -1e-9) throw new Error(`p${s.page}: カード幅が ${r3(-extra)}in 足りません`);
  const ws = need.map(v => Math.floor((v + extra / cards.length) * 1000) / 1000);
  const h = r3(o.bottom - o.y);
  let gap = Math.min(o.maxGap, ...cards.map(c => (c.items.length > 1 ? (h - cardFixedH(c)) / (c.items.length - 1) : Infinity)));
  gap = Math.floor(gap * 1000) / 1000;
  if (gap < -1e-9) throw new Error(`p${s.page}: カードの高さが ${r3(-gap * (cards[0].items.length - 1))}in ほど足りません（h=${h}）`);
  gap = Math.max(0, gap);
  let x = X0;
  const padB = cards.map((c, i) => { const p = drawCard(s, c, x, o.y, ws[i], h, gap); x += ws[i] + GAP; return p; });
  console.log(`  p${s.page} カード: 幅 ${ws.join(' / ')}in（必要 ${need.map(r3).join(' / ')}）、${o.y}〜${r3(o.y + h)}in（高さ ${h}in）、` +
    `項目間 ${gap}in、下余白 ${padB.join(' / ')}in（上 ${CARD.padT}in）`);
}

/**
 * 強調ボックス（白地＋色枠。文字は太字・中央）。S2・S5 は 28pt の通常の行送り（decklib の box）。
 * 枠の色はデッキ内のルール: 通常の要点はティール、危険の警告（S3 の正常血糖ケトアシドーシス）は赤（C.red = 危険）。
 */
const EMPH = { x: X0 + 0.6, w: BW - 1.2, radius: 0.1, lineW: 3.5 };
function emphBox(s, y, h, lines, lineColor, size = 28) {
  s.box({
    x: EMPH.x, y: r3(y), w: EMPH.w, h, fill: C.white, radius: EMPH.radius,
    line: { color: lineColor, width: EMPH.lineW },
    text: lines, size, bold: true, color: C.dark,
  });
}
/** 行送りを詰めた強調ボックス（S3・S6。22pt・2 行）。戻り値は高さ */
const EMPH_T = { padY: 0.05, padX: 0.15 };
function emphTightH(n, size) { return r3(2 * EMPH_T.padY + n * lineIn(size)); }
function emphTight(s, yBottom, lines, lineColor, size = 22) {
  const h = emphTightH(lines.length, size);
  const y = yBottom - h;
  s.box({ x: X0, y: r3(y), w: BW, h, fill: C.white, radius: EMPH.radius, line: { color: lineColor, width: EMPH.lineW } });
  tightText(s, lines, { x: X0 + EMPH_T.padX, y: y + EMPH_T.padY, w: BW - 2 * EMPH_T.padX, size, bold: true, align: 'center' });
  return { y, h };
}

/**
 * 行の多い表（S4 用、3-11 の tightTable を写したもの）。decklib の s.table と同じ見た目（見出し行 D9D9D9・太字、
 * 罫線 BFBFBF 1pt、上下中央）で、行間だけを固定値 LS pt（= 1.2 em）に詰める。見出し行も本文と同じ左揃え。
 * s.table は行高を実機メイリオの 1.5 em で取るため、6 行すべて 2 行になる表が本文ゾーンに入らない。
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
  return { x: o.x, y: o.y, w, h: r3(h), rowH };
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['SGLT2', 'ケトアシ\nドーシス', '経口薬', 'インス\nリン', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-13 常用薬②：糖尿病薬',
  });
  const slide = o => deck.slide(o);   // 章チップはライブラリ既定（4 字ラベルの 10.5pt 調整も decklib 側）

  // ===== S1 表紙 =====
  {
    const subtitle = '4-13 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: '常用薬②：糖尿病薬', subtitle, meta, notes: notesOf('S1') });
  }

  // 箇条書き＋強調ボックスのスライド（S2・S5）: 24pt・段落後 10pt。強調ボックスは 28pt 太字 2 行
  const BUL = { y: 1.9, size: 24, gap: 10 };
  const bulH = n => heightFor(n, BUL.size, BUL.gap);
  const EMPH_H = 1.2;

  // ===== S2 SGLT2 阻害薬の休薬（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S2';
    const s = slide({ title: 'SGLT2 阻害薬の休薬', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, '日本糖尿病学会：3 日前から休薬、食事が十分とれてから再開'), ['3 日前から休薬', C.vermilion]),
      q(id, '米国も 3 日前（エルツグリフロジンは 4 日）。英国は前日と当日'),
      mark(q(id, '心不全・CKD の目的で処方されている例も同じ薬。処方医と相談'), ['処方医と相談', C.blue]),
      mark(q(id, '配合剤（DPP-4 阻害薬などとの合剤）の中身を見落とさない'), ['中身', C.blue]),
      mark(q(id, '緊急：すぐ中止。ケトン体と血液ガスを反復測定、上昇なら治療'), ['すぐ中止', C.red]),
    ], BUL);
    const lines = brk(id, '止めたら再開日まで指示（止めっぱなしでは心不全・腎を守れない）', '（止めっぱなし');
    const by = BOTTOM - 0.04 - EMPH_H;
    if (by < BUL.y + bulH(5) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(5))}in`);
    emphBox(s, by, EMPH_H, [mark(lines[0], ['再開日まで指示', C.vermilion]), lines[1]], C.teal);
  }

  // ===== S3 正常血糖ケトアシドーシス（2列カード + 強調ボックス） =====
  // 本文 17 行ぶんの項目が 2 列カード＋強調ボックスに入らないため、カード本文と強調ボックスは行送り 1.2 em。
  // それでも左のカードが 1 行あふれるので、項目 3 だけ語を詰めた（alt()）。
  // 項目 3 を spec どおり 2 行（「手術ストレスと絶食が重なると」/「起きやすい」）にすると左カードは 9 行で、
  // 0.37in 足りない（カードは上 TOP_BOX〜強調ボックスの 0.24in 上まで、項目間 0）。1 行で入れるには幅 約7.0in が要る。
  // 強調ボックスは危険の警告なので赤枠（EMPH のルール）。
  {
    const id = 'S3';
    const s = slide({ title: '正常血糖ケトアシドーシス', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    const emLines = brk(id, 'SGLT2 阻害薬を飲んでいた・最近止めた患者の術後の嘔気・腹痛・頻呼吸 → 血液ガスとケトン', '術後の');
    const em = emphTight(s, BOTTOM, [emLines[0], mark(emLines[1], ['血液ガスとケトン', C.red])], C.red);
    cardRow(s, [
      {
        title: q(id, 'なぜ起きる・どう気づく'), titleColor: C.blue,
        items: [
          brk(id, '尿糖で血糖が下がり、インスリン↓グルカゴン↑→ケトン体', 'グルカゴン'),
          brk(id, '血糖は正常〜軽度高値（周術期例の中央値 179 mg/dL）', '（周術期'),
          [alt(id, '手術ストレスと絶食が重なると起きやすい', '手術ストレスと絶食で起きやすい', '左カードが 1 行あふれるため「が重なると」→「で」')],
          brk(id, '倦怠感、嘔気・嘔吐、腹痛、頻呼吸、意識の変化', '頻呼吸'),
          [q(id, 'PONV や術後痛と紛らわしい')],
        ],
      },
      {
        title: q(id, '診断と治療'), titleColor: C.vermilion,
        items: [
          brk(id, '血中 β-ヒドロキシ酪酸 3.0 mmol/L 以上', '3.0'),
          [q(id, 'pH 7.3 未満 または HCO3− 18 未満')],
          brk(id, '尿ケトンは補助（早期は軽く出る。術後は陽性も多い）', '術後は'),
          [q(id, 'アニオンギャップ 12 超')],
          brk(id, '輸液＋インスリン＋最初から 5〜10% ブドウ糖（K を先に測る）', '5〜10%'),
        ],
      },
    ], { y: TOP_BOX, bottom: em.y - GAP, maxGap: 0.12 });
  }

  // ===== S4 経口薬と GLP-1 の扱い（表 3列 × 6行＋見出し行） =====
  // spec の列幅（3.2 / 3.4 / 5.6）では 3 列目が狭く 6 行すべてが 2 行になり、表の高さ 5.37in が
  // 本文ゾーン（TOP_BOX〜BOTTOM の 4.62in）に入らない（入るのは 2 行の行が 4 つまで）。
  // メトホルミン・DPP-4・GLP-1・α-GI の 4 行は 3 列目がどの幅でも 2 行になるので、SGLT2 と SU 薬の行を
  // 1 行に保つ必要があり、3 列目は「正常血糖ケトアシドーシス（→ S2、S3）」が 1 行に入る 6.52in が下限。
  // 残り 5.68in（合計 12.2in）は、2 列目を「当日朝は飲まない」が 1 行に入る 2.91in、1 列目を 2.77in に配る
  // （「GLP-1 受容体作動薬」を 1 行にするには 1 列目 3.28in が要り、2 列目が「3 日前から中止」の入る 2.52in を割る）。
  // 「GLP-1 受容体作動薬」は「GLP-1 受容体」/「作動薬」、「α-GI、チアゾリジン」は「、」の後で折る。
  {
    const id = 'S4';
    const s = slide({ title: '経口薬と GLP-1 の扱い', chapter: 2, notes: notesOf(id), sources: sourcesOf(id) });
    const drug = text => ({ text, bold: true });
    const t = [
      ['薬', '当日', '理由・注意'],
      [drug(q(id, 'SGLT2 阻害薬')), { text: q(id, '3 日前から中止'), color: C.vermilion, bold: true }, q(id, '正常血糖ケトアシドーシス（→ S2、S3）')],
      [drug(q(id, 'メトホルミン')), brk(id, '2 日前から休薬（国内の目安）', '（国内').join('\n'), brk(id, '添付文書は手術前後を禁忌。造影剤では 48 時間', '造影剤').join('\n')],
      [drug(q(id, 'SU 薬、グリニド')), q(id, '当日休薬'), q(id, '絶食中の遷延性低血糖。ブドウ糖を準備')],
      [drug(q(id, 'DPP-4 阻害薬')), brk(id, '当日朝は休薬（英国は継続）', '（英国').join('\n'), brk(id, '低血糖はまれ。国内添付文書は手術前後を禁忌', '国内').join('\n')],
      [drug(brk(id, 'GLP-1 受容体作動薬', '作動薬').join('\n')), brk(id, '原則続ける（→ 4-3）', '（→').join('\n'), brk(id, '高リスクは流動食・胃エコー。休薬なら週 1 回製剤は 1 週前', '休薬').join('\n')],
      [drug(brk(id, 'α-GI、チアゾリジン', 'チアゾ').join('\n')), q(id, '当日朝は飲まない'), brk(id, 'α-GI は食事とともに効く。チアゾリジンは体液貯留に注意', 'チアゾ').join('\n')],
    ];
    const res = tightTable(s, t, { x: X0, y: TOP_BOX, colW: [2.77, 2.91, 6.52] });
    console.log(`  p${s.page} 表: 行高 ${res.rowH.join(' / ')}in、下端 ${r3(res.y + res.h)}in`);
    if (res.y + res.h > BOTTOM + 0.011) throw new Error(`${id}: 表の下端 ${r3(res.y + res.h)} > ${BOTTOM}`);
  }

  // ===== S5 インスリンの術前調整（箇条書き5行 + 強調ボックス） =====
  // 色の使い分け: 青＝項目の頭、バーミリオン＝数値（75〜80%、50%）、赤＝危険（絶対に抜かない）。
  // 赤とバーミリオンは色覚で見分けにくいので、数値以外にバーミリオンを使わない（「静注インスリン持続」は太字の黒）。
  {
    const id = 'S5';
    const s = slide({ title: 'インスリンの術前調整', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, '持効型：前夜（朝打ちなら当日朝）の分を普段の 75〜80%'), ['持効型：', C.blue], ['75〜80%', C.vermilion]),
      mark(q(id, '中間型（NPH）：当日朝は 50%'), ['中間型（NPH）：', C.blue], ['50%', C.vermilion]),
      mark(q(id, '速効型・超速効型：絶食中は打たない（補正分だけ）'), ['速効型・超速効型：', C.blue]),
      mark(q(id, '1 型糖尿病：基礎インスリンは絶対に抜かない（抜くと DKA）'), ['1 型糖尿病：', C.blue], ['絶対に抜かない', C.red]),
      mark(q(id, '1 食を超える絶食、ポンプが使えない：静注インスリン持続へ'), ['静注インスリン持続', C.dark]),
    ], BUL);
    const lines = brk(id, 'ポンプ（CSII）は条件を満たせば継続。術野から離し、対極板と電気メスの間に置かない', '術野');
    const by = BOTTOM - 0.04 - EMPH_H;
    if (by < BUL.y + bulH(5) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(5))}in`);
    emphBox(s, by, EMPH_H, lines, C.teal);
  }

  // ===== S6 HbA1c と血糖目標（2列カード + 強調ボックス） =====
  // 各カード 8 行ぶんの項目と 2 行の強調ボックスが入るよう、カード本文と強調ボックスは行送り 1.2 em。
  {
    const id = 'S6';
    const s = slide({ title: 'HbA1c と血糖目標', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    const emLines = brk(id, '国内の目安：空腹時 100〜140 または食後 160〜200 mg/dL、尿ケトン陰性。満たさなければ糖尿病内科へ（術前入院も）', '尿ケトン');
    const em = emphTight(s, BOTTOM, [mark(emLines[0], ['国内の目安：', C.blue]), emLines[1]], C.teal);
    cardRow(s, [
      {
        title: q(id, 'HbA1c（3 か月以内の値を）'), titleColor: C.blue,
        items: [
          brk(id, '英国：8.5% 超は手術を進めるか、かかりつけに戻すかを臨床判断', 'かかりつけ'),
          brk(id, 'ESC 2022：8.5% 以上の待機手術は安全・実際的なら延期（I）', '安全'),
          brk(id, '理想は 6.5〜7%。下げすぎの害にも注意', '下げすぎ'),
          brk(id, '国内に明確な延期の閾値は見当たらない', '閾値'),
        ],
      },
      {
        title: q(id, '周術期の血糖'), titleColor: C.vermilion,
        items: [
          brk(id, '目標 100〜180 mg/dL（上限 180。ICU は 140〜180）', '（上限'),
          brk(id, '静注インスリン中は 1 時間ごとに測る', '1 時間'),
          brk(id, '235 mg/dL 超が 2 回続けば、ケトン体を測る', 'ケトン体'),
          brk(id, 'デキサメタゾン（PONV 予防）で上がる', '（PONV'),
        ],
      },
    ], { y: TOP_BOX, bottom: em.y - GAP, maxGap: 0.12 });
  }

  // ===== S7 当日の指示の例（大きなボックス1つに指示例6行 + 下に1行） =====
  // 指示例 6 項目は 22pt で 10 行になる。行送り 1.2 em で箱に入れ、項目の頭（「：」まで）を青の太字にして区切る。
  // 2 行目は 1 字ぶん下げる。
  // 箱の下余白は上より厚く取る（最終行が箱の下端に寄って見えないように）。10 行＋下の 1 行＋2 行の出典で
  // 縦の余りは 0.2in 弱しかなく、項目間に見える大きさの余白（0.05in × 5）は入らないので、余りは下余白に充てる。
  // 「◯/◯」は spec の文言（4 個とも U+25EF）のまま。描画確認で大きさが揃わないのは代替フォントのため。
  {
    const id = 'S7';
    const s = slide({ title: '当日の指示の例', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    const lab = C.blue;
    const items = [
      (() => { const l = brk(id, 'SGLT2 阻害薬：◯/◯（手術 3 日前）から飲まない、最終内服 ◯/◯。食事が戻れば再開（処方医に連絡済み）', '食事が'); return [mark(l[0], ['SGLT2 阻害薬：', lab]), l[1]]; })(),
      (() => { const l = brk(id, 'メトホルミン：手術 2 日前から中止（院内ルールに従う）。造影 CT なら 48 時間後に腎機能を見て再開', '造影'); return [mark(l[0], ['メトホルミン：', lab]), l[1]]; })(),
      [mark(q(id, 'SU 薬・DPP-4 阻害薬：当日朝は飲まない'), ['SU 薬・DPP-4 阻害薬：', lab])],
      (() => { const l = brk(id, '持効型インスリン：前夜（朝打ちなら当日朝）は普段の 75%（1 型は必ず打つ）。速効型：当日朝は打たない', '（1 型'); return [mark(l[0], ['持効型インスリン：', lab]), mark(l[1], ['速効型：', lab])]; })(),
      (() => { const l = brk(id, '血糖：入室時と 1〜2 時間ごと。100 未満で対策、70 未満はブドウ糖、180 以上が続けばインスリン', '70 未満'); return [mark(l[0], ['血糖：', lab]), l[1]]; })(),
      [mark(q(id, '嘔気・腹痛・頻呼吸、または血糖 235 超が 2 回：血液ガスとケトン体'), ['嘔気・腹痛・頻呼吸、または血糖 235 超が 2 回：', lab])],
    ];
    const PAD = { x: 0.14, top: 0.08, bottom: 0.14 }, IND = 0.34;
    const nLines = items.reduce((a, it) => a + it.length, 0);
    const boxH = r3(PAD.top + PAD.bottom + nLines * lineIn(22));
    const footH = BOX_H[24];   // 24pt 1 行の枠（decklib / lint の下限。実機メイリオ 1.5 em でも警告なし）
    const by = TOP_BOX;
    if (by + boxH + GAP + footH > BOTTOM + 1e-9) throw new Error(`${id}: 箱 ${boxH} + 下の 1 行が ${BOTTOM} を超えます`);
    s.box({ x: X0, y: by, w: BW, h: boxH, fill: C.bg, radius: 0.08 });
    let y = by + PAD.top;
    const iw = BW - 2 * PAD.x;
    items.forEach(lines => {
      tightText(s, [lines[0]], { x: X0 + PAD.x, y, w: iw, size: 22 });
      if (lines.length > 1) tightText(s, lines.slice(1), { x: X0 + PAD.x + IND, y: y + lineIn(22), w: iw - IND, size: 22 });
      y += lines.length * lineIn(22);
    });
    const fy = r3(by + boxH + GAP + (BOTTOM - (by + boxH + GAP) - footH) / 2);
    s.text([mark(q(id, '配合剤は成分名で確認。止めた薬は再開日まで書く'), ['止めた薬は再開日まで書く', C.vermilion])],
      { x: X0, y: fy, w: BW, h: footH, size: 24, bold: true, align: 'center' });
    console.log(`  p${s.page} 指示例: 箱 ${by}〜${r3(by + boxH)}in（${nLines} 行、余白 上 ${PAD.top} / 下 ${PAD.bottom}in）、` +
      `下の 1 行 ${fy}〜${r3(fy + footH)}in`);
  }

  // ===== S8 まとめ（箇条書き5行） =====
  {
    const id = 'S8';
    const s = slide({ title: 'まとめ', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, 'SGLT2 阻害薬は 3 日前から休薬。食事が戻ったら再開日を決める'), ['3 日前から休薬', C.vermilion]),
      mark(q(id, '正常血糖ケトアシドーシス：血糖が正常でもケトンと血液ガス'), ['ケトンと血液ガス', C.red]),
      mark(q(id, 'メトホルミンは 2 日前、SU 薬・DPP-4 は当日朝に休薬'), ['2 日前', C.blue], ['当日朝', C.blue]),
      mark(q(id, '持効型は 75〜80%、速効型は打たない。1 型は基礎を抜かない'), ['75〜80%', C.blue], ['1 型は基礎を抜かない', C.red]),
      mark(q(id, 'HbA1c は 3 か月以内の値。周術期の血糖は上限 180 mg/dL'), ['上限 180 mg/dL', C.blue]),
    ], { y: 2.0, size: 24, gap: 28 });
  }

  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
