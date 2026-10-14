'use strict';
/**
 * 3-10 γ計算の演習 — 組版スクリプト
 *
 *   cd slides/3-10_gamma_calculation
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *     書き写しの検査: 各スライドに置いた文字列を順につなげたものが、spec の on-slide の「」の中身を
 *     順につなげたものと一字一句同じでなければ Error（checkOnSlide）。
 *     「」の外の構造ラベル（定義:、式:、問 1: など）は文言ではないので検査の対象外。
 *     S5 の「問 1」〜「問 3」は layout の「番号つきの 3 つのカード」の番号として置く。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, LAYOUT, measure, heightFor } = require('../_lib/decklib');

const NAME = '3-10_gamma_calculation';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.RENDER_DIR || process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);
const RENDER = !process.argv.includes('--no-render');

// ---------------------------------------------------------------------------
// spec.md を読む（タイトル・on-slide の「」・ノート・出典・表紙の subtitle/meta）
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
/** 出典の項目（spec の sources 行を ' / ' で分けたもの） */
function sourcesOf(id) {
  if (!SPEC[id].sources) return undefined;
  return SPEC[id].sources.split(' / ');
}

// ---------------------------------------------------------------------------
// 書き写しの検査: スライドに置いた文字列（置いた順）を記録し、spec の「」と突き合わせる
// ---------------------------------------------------------------------------
const PLACED = {};
const textOf = line => (Array.isArray(line) ? line.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(line));
function rec(id, line) { (PLACED[id] = PLACED[id] || []).push(textOf(line)); }
function checkOnSlide() {
  for (const [id, sp] of Object.entries(SPEC)) {
    if (!sp.quoted.length) continue;
    // 改行位置の空白は落とすので、つなぎ目の比較は空白を除いて行い、
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

/** spec の文言 str を、breaks の各語の直前で改行して行配列にする（文言は変えない。改行位置の空白は落とす） */
function brk(str, ...breaks) {
  const lines = [];
  let from = 0;
  for (const b of breaks) {
    const k = str.indexOf(b, from + 1);
    if (k < 0) throw new Error(`「${str}」に改行位置「${b}」が見つかりません`);
    lines.push(str.slice(from, k).replace(/[ 　]+$/, ''));
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
// 共通の寸法・色
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;          // 0.55
const BW = LAYOUT.body.w;          // 12.23
const XR = LAYOUT.body.right;      // 12.78
const TOP = 1.85;                  // 本文の上端
const GAP = 0.24;                  // 要素間の最小間隔（0.6cm）
const r3 = v => Math.round(v * 1000) / 1000;
/** 本文の下端（出典の枠の上端から 0.24in 空ける） */
const bottomOf = s => (s.hasSources ? s.sourceTop : LAYOUT.body.bottom + GAP) - GAP;
const KEY = C.blue;                // 式・定義・途中式の見出し
const ANS = C.vermilion;           // 答え
const DANGER = C.red;              // 危険の強調

/**
 * 実機の余裕を見た幅の検査。メイリオでは × ÷ が全角幅で描かれることがあるので、
 * 推定器（0.56 字）に 0.44 字を足して数え、枠幅の 98% を超える行は Error（改行位置で直す）。
 */
function safeFit(id, line, size, widthIn) {
  const t = textOf(line);
  const extra = (t.match(/[×÷]/g) || []).length * 0.44 * (size / 72) * 1.12;
  const need = measure(t, size) + extra;
  if (need > widthIn * 0.98) throw new Error(`${id}: 「${t}」が実機で折り返すおそれ（${r3(need)}in > ${r3(widthIn)}in × 0.98）。改行位置を足してください`);
  return need;
}

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す） */
function contentSlide(deck, id, chapter) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id) });
  if (s.hasSources && s.sourceTop < LAYOUT.source.y2 - 1e-6) throw new Error(`${id}: 出典が3行以上`);
  return s;
}

/** 箇条書き（1項目1行）。下端が本文の下端を越えたら Error */
function bulletList(s, id, items, o) {
  items.forEach(t => rec(id, typeof t === 'object' && !Array.isArray(t) ? t.text : t));
  const w = o.w || 12.2;
  items.forEach(t => safeFit(id, typeof t === 'object' && !Array.isArray(t) ? t.text : t, o.size, w - (Math.round(o.size * 1.15)) / 72));
  s.bullets(items, { x: o.x, y: o.y, w, size: o.size, gap: o.gap });
  const bh = heightFor(items.length, o.size, o.gap);
  if (o.y + bh > bottomOf(s) + 1e-6) throw new Error(`${id}: 箇条書きが出典に近すぎます（下端 ${r3(o.y + bh)}in > ${r3(bottomOf(s))}in）`);
  console.log(`  ${id}: bullets ${o.size}pt gap ${o.gap} → bottom ${r3(o.y + bh)} (limit ${r3(bottomOf(s))})`);
  return o.y + bh;
}

/**
 * 箇条書き（1項目1行）を、行頭記号と本文を別の枠にして置く。見た目は s.bullets と同じ（記号＋字下げ 1.15em）。
 * 全角の開き括弧（『「（ など）は字面が字送りの右半分にあり、行頭が他の行より右にずれて見えるので、
 * その行だけ本文の枠を 0.45em 左に寄せて行頭をそろえる。pitch は行の間隔（in）。
 */
function alignedBullets(s, id, items, o) {
  const size = o.size;
  const ind = Math.round(size * 1.15) / 72;
  const lh = heightFor(1, size, 0);
  items.forEach((t, i) => {
    const y = r3(o.y + i * o.pitch);
    const txt = textOf(t);
    rec(id, t);
    const shift = /^[「『（【〔［]/.test(txt) ? (size / 72) * 0.45 : 0;
    const x = r3(o.x + ind - shift);
    const w = r3(o.x + o.w - x);
    safeFit(id, t, size, w);
    s.text(['•'], { x: o.x, y, w: r3(ind), h: lh, size, color: C.dark });
    s.text([t], { x, y, w, h: lh, size, color: C.dark });
  });
  const bottom = o.y + (items.length - 1) * o.pitch + lh;
  if (bottom > bottomOf(s) + 1e-6) throw new Error(`${id}: 箇条書きが出典に近すぎます（下端 ${r3(bottom)}in > ${r3(bottomOf(s))}in）`);
  console.log(`  ${id}: aligned bullets ${size}pt pitch ${o.pitch} → bottom ${r3(bottom)} (limit ${r3(bottomOf(s))})`);
  return bottom;
}

/** 文字入りの箱（1図形）。文言を記録し、実機の余裕を検査する */
function textBox(s, id, line, o) {
  rec(id, line);
  safeFit(id, line, o.size, o.w - 2 * (o.margin != null ? o.margin : 0.08));
  const text = Array.isArray(line) ? [line] : line;
  s.box({ ...o, text });
}

// ---------------------------------------------------------------------------
// 見出し＋項目のカード（各項目は1〜数行。項目の間は細い区切り線）。文字はすべて箱の内側。
// 左右2枚は「段」をそろえる: i 番目の項目の枠の高さは、2枚のうち行数の多い方に合わせ、
// 行数の少ない方はその枠の上下中央に置く（区切り線の位置が左右で一致する）。
// ---------------------------------------------------------------------------
const CARD = { pad: 0.16, titleSize: 28, size: 26, titleGap: 0.14, lineGap: 1, minGap: 0.2, maxGap: 0.5 };
function cardHeadH() {
  return CARD.pad + heightFor(1, CARD.titleSize, 0) + CARD.titleGap;
}
function itemCard(s, id, c, x, y, w, h, gap, rowH) {
  const pad = CARD.pad;
  s.box({ x, y, w, h, fill: C.bg, radius: 0.08 });
  const iw = r3(w - 2 * pad);
  const th = heightFor(1, CARD.titleSize, 0);
  let cy = y + pad;
  safeFit(id, c.title, CARD.titleSize, iw);
  s.text([c.title], { x: r3(x + pad), y: r3(cy), w: iw, h: th, size: CARD.titleSize, bold: true, color: c.titleColor });
  rec(id, c.title);
  cy += th + CARD.titleGap;
  c.items.forEach((lines, i) => {
    const ih = heightFor(lines.length, CARD.size, CARD.lineGap);
    lines.forEach(l => safeFit(id, l, CARD.size, iw));
    s.text(lines, { x: r3(x + pad), y: r3(cy + (rowH[i] - ih) / 2), w: iw, h: ih, size: CARD.size, gap: CARD.lineGap, color: C.dark });
    lines.forEach(l => rec(id, l));
    cy += rowH[i];
    if (i < c.items.length - 1) {
      s.hline({ x: r3(x + pad), y: r3(cy + gap / 2), w: iw, color: C.border, width: 1 });
      cy += gap;
    }
  });
  if (cy > y + h - pad + 1e-6) throw new Error(`${id} カード「${c.title}」: 高さ不足（必要 ${r3(cy + pad - y)}in > ${h}in）`);
}
/** 左右2枚のカード（等幅）。高さは本文の下端まで。項目の段（枠の高さ・間隔）は2枚で共通 */
function twoCards(s, id, cards) {
  const w = r3((BW - GAP) / 2);
  const h = r3(bottomOf(s) - TOP);
  const n = cards[0].items.length;
  if (cards.some(c => c.items.length !== n)) throw new Error(`${id}: 左右のカードの項目数が違います（段をそろえられません）`);
  const rowH = Array.from({ length: n }, (_, i) =>
    Math.max(...cards.map(c => heightFor(c.items[i].length, CARD.size, CARD.lineGap))));
  const itemsH = rowH.reduce((a, b) => a + b, 0);
  let gap = n < 2 ? 0 : Math.min(CARD.maxGap, (h - cardHeadH() - CARD.pad - itemsH) / (n - 1));
  if (n >= 2 && gap < CARD.minGap - 1e-6) throw new Error(`${id}: カードの高さが足りません（項目間隔 ${r3(gap)}in < ${CARD.minGap}in）`);
  gap = Math.floor(gap * 100) / 100;
  cards.forEach((c, i) => itemCard(s, id, c, r3(X0 + i * (w + GAP)), TOP, w, h, gap, rowH));
  console.log(`  ${id}: 2 cards w=${w} h=${h} gap=${gap} rows=${rowH.join('/')}`);
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['定義', '式', '演習', '検算', '落とし穴'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '3-10 γ計算の演習',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 γ とは何か（定義を大きく＋箇条書き3行） =====
  {
    const id = 'S2';
    const s = contentSlide(deck, id, 0);
    // 定義（44pt・中央）
    const dw = 8.0;
    textBox(s, id, '1 γ = 1 μg/kg/分', {
      x: r3(X0 + (BW - dw) / 2), y: 2.05, w: dw, h: 1.35, fill: C.bg, radius: 0.1,
      size: 44, bold: true, color: KEY,
    });
    bulletList(s, id, [
      '体重 1 kg あたり、1 分あたり、1 μg',
      '体重で割っているので、患者が違っても同じ数で指示できる',
      mark('ポンプに入れるのは mL/時。橋渡しが『希釈濃度』', ['『希釈濃度』', ANS]),
    ], { y: 3.88, size: 26, gap: 26 });
  }

  // ===== S3 換算の式（式を大きく、その下に逆の式と一文） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 1);
    textBox(s, id, mark('mL/時 = γ × 体重(kg) × 60 ÷ 濃度(μg/mL)', ['60', ANS]), {
      x: X0, y: 1.95, w: BW, h: 1.35, fill: C.bg, radius: 0.1, size: 34, bold: true, color: C.dark,
    });
    const iw = 10.2;
    textBox(s, id, mark('γ = mL/時 × 濃度(μg/mL) ÷ (体重(kg) × 60)', ['60', ANS]), {
      x: r3(X0 + (BW - iw) / 2), y: 3.75, w: iw, h: 1.05, line: { color: C.border, width: 1.5 }, radius: 0.1,
      size: 28, color: C.dark,
    });
    const line = mark('60 は『分』を『時』に直す数', ['60', ANS]);
    rec(id, line);
    safeFit(id, line, 28, BW);
    s.text([line], { x: X0, y: 5.3, w: BW, size: 28, align: 'center', color: C.dark });
  }

  // ===== S4 希釈濃度の出し方（式＋3例の表） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 1);
    textBox(s, id, mark('濃度(μg/mL) = 薬の量(mg) × 1,000 ÷ 総量(mL)', ['× 1,000', ANS]), {
      x: X0, y: 1.95, w: BW, h: 1.1, fill: C.bg, radius: 0.1, size: 32, bold: true, color: C.dark,
    });
    const head = ['薬', '希釈（例）', '濃度'];
    const body = [
      ['レミフェンタニル', '2 mg を 40 mL に', '50 μg/mL'],
      ['ノルアドレナリン', '5 mg を 50 mL に', '100 μg/mL'],
      ['ニカルジピン', '10 mg/10 mL 原液（施設の方法）', '1,000 μg/mL'],
    ];
    head.forEach(t => rec(id, t));
    body.forEach(r => r.forEach(t => rec(id, t)));
    const size = 24;
    const colW = [3.45, 6.0, 2.78];
    const m = [4, 10, 4, 12];
    body.forEach(r => r.forEach((t, ci) => safeFit(id, t, size, colW[ci] - (m[1] + m[3]) / 72)));
    // 見出しの揃えは本文の列に合わせる（薬・希釈は左、濃度は中央）
    const headRow = head.map((t, ci) => ({ text: t, align: ci < 2 ? 'left' : 'center' }));
    const rows = [headRow].concat(body.map(([drug, dil, conc]) => [
      { text: drug, bold: true },
      dil,
      { text: conc, bold: true, color: KEY, align: 'center' },
    ]));
    const t = s.table(rows, { x: X0, y: 3.5, colW, size, rowH: [0.66, 0.72, 0.72, 0.72], margin: m });
    if (t.y + t.h > bottomOf(s) + 1e-6) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)}in > ${r3(bottomOf(s))}in`);
    console.log(`  ${id}: table bottom ${r3(t.y + t.h)} (limit ${r3(bottomOf(s))})`);
  }

  // ===== S5 演習（番号つきの 3 つのカード・縦） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 2);
    const qs = [
      brk('体重 60 kg。レミフェンタニル 2 mg を 40 mL に希釈。0.25 γ は何 mL/時？', '0.25 γ'),
      brk('体重 50 kg。ノルアドレナリン 5 mg を 50 mL に希釈。0.05 γ は何 mL/時？　また、3 mL/時 なら何 γ？', '0.05 γ'),
      brk('体重 70 kg。ニカルジピン 10 mg/10 mL 原液。2 γ は何 mL/時？　また、10 μg/kg のボーラスは何 mL？', '2 γ'),
    ];
    const size = 24, lgap = 6;
    const bottom = bottomOf(s);
    // カードの間隔は 0.3in（下限 0.24in より広く取る）。高さは 1.35in を上限に、本文の下端までに収める
    const cgap = 0.3;
    const ch = r3(Math.min(1.35, (bottom - TOP - 2 * cgap) / 3));
    if (ch < heightFor(2, size, lgap) + 0.2) throw new Error(`${id}: カードの高さが足りません（${ch}in）`);
    const chipW = 1.05, chipH = 0.72;
    const tx = r3(X0 + 0.18 + chipW + 0.24);
    const tw = r3(XR - 0.18 - tx);
    qs.forEach((lines, i) => {
      const y = r3(TOP + i * (ch + cgap));
      s.box({ x: X0, y, w: BW, h: ch, fill: C.bg, radius: 0.08 });
      s.box({
        x: r3(X0 + 0.18), y: r3(y + (ch - chipH) / 2), w: chipW, h: chipH, fill: KEY, radius: 0.08,
        text: `問 ${i + 1}`, size: 24, bold: true, color: C.white,
      });
      const th = heightFor(lines.length, size, lgap);
      lines.forEach(l => { rec(id, l); safeFit(id, l, size, tw); });
      s.text(lines, { x: tx, y: r3(y + (ch - th) / 2), w: tw, h: th, size, gap: lgap, color: C.dark });
    });
    console.log(`  ${id}: card h=${ch} gap=${cgap} bottom=${r3(TOP + 3 * ch + 2 * cgap)} text w=${tw}`);
  }

  // ===== S6 解答 1：レミフェンタニル（途中式 4 行＋答えの強調） =====
  // 4 行を右の角括弧でひとまとめにし、括弧の中央から答えの箱へ矢印。答えの箱は 4 行の上下中央にそろえる。
  // 答えの強調は箱だけにする（4 行目の「18 mL/時」は本文色のまま。箱と二重に見せない）。
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 2);
    // 26pt（S7・S8 のカードの途中式と同じ）。28pt では最長の行の右端（× ÷ が全角幅で描かれた場合）が
    // 9.0in 近くまで来て、括弧・矢印・答えの箱が右に入らない
    const size = 26;
    const steps = [
      mark('濃度：2 mg × 1,000 ÷ 40 mL = 50 μg/mL', ['濃度：', KEY]),
      mark('1 分あたり：0.25 × 60 kg = 15 μg/分', ['1 分あたり：', KEY]),
      mark('1 時間あたり：15 × 60 = 900 μg/時', ['1 時間あたり：', KEY]),
      mark('ポンプ：900 ÷ 50 = 18 mL/時', ['ポンプ：', KEY]),
    ];
    const lh = heightFor(1, size, 0), pitch = 1.02;
    const y0 = 2.2;
    // 行の幅は最長の行（実機の余裕込み）に合わせる
    const need = Math.max(...steps.map(l => safeFit(id, l, size, 9.0)));
    const lw = r3(need / 0.98 + 0.02);
    steps.forEach((l, i) => {
      rec(id, l);
      s.text([l], { x: X0, y: r3(y0 + i * pitch), w: lw, h: lh, size, color: C.dark });
    });
    // 角括弧（4 行の文字の上端〜下端）。最長の行の右端（実機の余裕込み）から 0.24in 以上、文字枠の外に置く
    const top = y0 + 0.06, bot = y0 + 3 * pitch + lh - 0.06, mid = (top + bot) / 2;
    const tick = 0.12;
    const kx = r3(Math.max(X0 + need + GAP, X0 + lw + 0.06) + tick);
    s.polyline([[kx - tick, r3(top)], [kx, r3(top)], [kx, r3(bot)], [kx - tick, r3(bot)]], { color: C.chipText, width: 2.5 });
    // 答え（強調）
    const bw = 3.3, bh = 1.5, bx = r3(XR - bw);
    if (bx - kx < 0.45) throw new Error(`${id}: 括弧と答えの箱の間が狭すぎます（${r3(bx - kx)}in）`);
    s.arrow({ x1: kx, y1: r3(mid), x2: r3(bx - 0.1), y2: r3(mid), color: C.chipText, width: 3 });
    textBox(s, id, '18 mL/時', {
      x: bx, y: r3(mid - bh / 2), w: bw, h: bh, fill: C.white, line: { color: ANS, width: 3 }, radius: 0.1,
      size: 48, bold: true, color: ANS,
    });
    console.log(`  ${id}: lines w=${lw} bracket x=${kx} y=${r3(top)}..${r3(bot)} answer box x=${bx} y=${r3(mid - bh / 2)}..${r3(mid + bh / 2)}`);
  }

  // ===== S7 解答 2：ノルアドレナリン（左右2枚のカード） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 2);
    twoCards(s, id, [
      {
        title: '0.05 γ → mL/時', titleColor: KEY,
        items: [
          brk('濃度：5 mg × 1,000 ÷ 50 = 100 μg/mL', '= 100'),
          brk('0.05 × 50 kg × 60 = 150 μg/時', '= 150'),
          [mark('150 ÷ 100 = 1.5 mL/時', ['1.5 mL/時', ANS])],
        ],
      },
      {
        title: '3 mL/時 → γ', titleColor: KEY,
        items: [
          brk('3 mL/時 × 100 μg/mL = 300 μg/時', '= 300'),
          ['300 ÷ 60 = 5 μg/分'],
          [mark('5 ÷ 50 kg = 0.1 γ', ['0.1 γ', ANS])],
        ],
      },
    ]);
  }

  // ===== S8 解答 3：ニカルジピン（左右2枚のカード） =====
  {
    const id = 'S8';
    const s = contentSlide(deck, id, 2);
    twoCards(s, id, [
      {
        title: '2 γ → mL/時', titleColor: KEY,
        items: [
          brk('濃度：1 mg/mL = 1,000 μg/mL', '= 1,000'),
          brk('2 × 70 kg × 60 = 8,400 μg/時', '= 8,400'),
          [mark('8,400 ÷ 1,000 = 8.4 mL/時', ['8.4 mL/時', ANS])],
        ],
      },
      {
        title: '10 μg/kg のボーラス', titleColor: KEY,
        items: [
          ['10 × 70 kg = 700 μg'],
          [mark('700 ÷ 1,000 = 0.7 mL', ['0.7 mL', ANS])],
          // 注意点（投与のしかた）なので、危険色の太字にして下段（左の答えの段）に置く
          [[{ text: 'ゆっくり静注', color: DANGER, bold: true }]],
        ],
      },
    ]);
  }

  // ===== S9 検算の習慣（箇条書き4行＋強調ボックス） =====
  {
    const id = 'S9';
    const s = contentSlide(deck, id, 3);
    // 3 行目は全角の『 で始まるので、行頭記号と本文を別の枠にして本文の行頭をそろえる（alignedBullets）。
    // 行の間隔は s.bullets（26pt・段落後 22pt）の実機の行送りと同じ 0.85in
    const bb = alignedBullets(s, id, [
      '答えの桁を疑う：1.5 と 15 と 150 のどれが『普通』か',
      '単位を書いて追う：μg/分 → μg/時 → mL/時',
      mark('『体重 × 3 mg を 50 mL に』なら 1 mL/時 = 1 γ', ['1 mL/時 = 1 γ', KEY]),
      '院内の標準希釈表と、2 人での確認',
    ], { x: X0, y: 1.98, w: 12.2, size: 26, pitch: 0.85 });
    const bh = 1.1;
    const by = r3(bottomOf(s) - bh);
    if (by < bb + GAP - 1e-6) throw new Error(`${id}: 強調ボックスが箇条書きに近すぎます`);
    textBox(s, id, '計算が合っていても、濃度の前提が違えば全部違う', {
      x: X0, y: by, w: BW, h: bh, fill: C.bg, radius: 0.1, size: 30, bold: true, color: DANGER,
    });
  }

  // ===== S10 落とし穴（箇条書き5行） =====
  {
    const id = 'S10';
    const s = contentSlide(deck, id, 4);
    bulletList(s, id, [
      // 見出し語（取り違える 2 組の単位）を太字にして、2〜5 行目と形をそろえる
      mark('mg と μg（1,000 倍）、分と時（60 倍）', ['mg と μg', null], ['1,000 倍', ANS], ['分と時', null], ['60 倍', ANS]),
      mark('総量：薬液の体積を入れ忘れる', ['総量：', null]),
      mark('ポンプの設定単位（mL/時か、γか、mg/時か）', ['ポンプの設定単位', null]),
      mark('ライン内の死腔：流量が少ないと届くまで時間がかかる', ['ライン内の死腔：', null]),
      mark('体重：実測か推定か、肥満では何を使うか（→ 3-1）', ['体重：', null]),
    ], { y: 2.0, size: 28, gap: 22 });
  }

  // ===== S11 まとめ（箇条書き5行） =====
  {
    const id = 'S11';
    const s = contentSlide(deck, id, 4);
    bulletList(s, id, [
      { text: '1 γ = 1 μg/kg/分', color: KEY, bold: true },
      { text: 'mL/時 = γ × 体重 × 60 ÷ 濃度(μg/mL)', color: KEY, bold: true },
      { text: '濃度 = 薬(mg) × 1,000 ÷ 総量(mL)', color: KEY, bold: true },
      '答えの桁を疑い、単位を書いて追う',
      '院内の標準希釈表に従い、2 人で確認する',
    ], { y: 1.98, size: 28, gap: 24 });
  }

  checkOnSlide();
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
