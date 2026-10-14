'use strict';
/**
 * 3-2 静脈麻酔薬 — 組版スクリプト
 *
 *   cd slides/3-2_iv_anesthetics
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *     書き写しの検査: 各スライドに置いた文字列を順につなげたものが、spec の on-slide の「」の中身を
 *     順につなげたものと一字一句同じでなければ Error（checkOnSlide）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * この講義だけの上書き（decklib の既定はほかの講義と共通なので触らない）:
 *   - 章チップ: 12pt・折り返しなし・幅 0.75in・右端を本文の右端 12.78in に・タイトルの行の高さの中央に置く。
 *     現在章の地は濃いティール 007C80（白太字で約 5:1）。
 *   - タイトル枠: チップ列の左端の手前で止める（既定の 9.3in 幅だとチップ列に重なる）。
 *   - 出典とページ番号: 下端を 7.20in にそろえる（スライド下端から 0.30in）。出典枠の幅は 11.4in（入る範囲で）。
 *   書き出し前の pptxgenjs のスライドオブジェクトを直接書き換える（3-11・2-8 の restyleChips と同じ方法）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, LAYOUT, measure, heightFor } = require('../_lib/decklib');

const NAME = '3-2_iv_anesthetics';
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
const sourcesOf = id => (SPEC[id].sources ? SPEC[id].sources.split(' / ') : undefined);

// ---------------------------------------------------------------------------
// 書き写しの検査: スライドに置いた文字列（置いた順）を記録し、spec の「」と突き合わせる
// ---------------------------------------------------------------------------
const PLACED = {};
const textOf = line => (Array.isArray(line) ? line.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(line));
function rec(id, line) { (PLACED[id] = PLACED[id] || []).push(textOf(line)); }
function checkOnSlide() {
  for (const [id, sp] of Object.entries(SPEC)) {
    if (!sp.quoted.length) continue;
    // 改行位置の空白（「小児の ICU」⏎「人工呼吸中の…」）は落ちるので、つなぎ目の比較は空白を除いて行い、
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
const TOP = 1.85;                  // 本文の上端
const GAP = 0.24;                  // 要素間の最小間隔
const SRC2_TOP = LAYOUT.source.y2; // 出典2行のときの出典枠の上端 6.5
const b = t => ({ text: t, bold: true });                       // 数値の強調（太字）
const danger = t => ({ text: t, bold: true, color: C.red });    // 禁忌（レッド）
const drug = t => ({ text: t, bold: true });                     // まとめの薬剤名（色を付けると行頭記号の色も変わるので太字だけ）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;
const CM = 2.54;

// ---------------------------------------------------------------------------
// 章チップ・タイトル枠・出典・ページ番号（この講義だけの上書き）
// ---------------------------------------------------------------------------
const CHAPTERS = ['共通', 'プロポ\nフォール', 'レミマ\nゾラム', 'ミダ・\nケタ', '使い分け'];

/**
 * 章チップ。decklib は 11pt・0.68×0.64in・上端 0.06in・右端 13.13in で置き、
 * 「フォール」（4 字）が実機でも描画でも折り返して 3 行になる。
 * - 文字 12pt（slide_lint の小サイズ許容は 10.5/12/16pt）。bodyPr を wrap="none" にして、ラベルの改行（spec の 2 行）だけで組む。
 * - 幅 0.75in・間隔 0.04in。右端は本文の右端 12.78in。左端は 8.87in。
 *   一番長いタイトル（S2「静脈麻酔薬に共通すること」、lint 推定の右端 8.76in）から lint の最小間隔 0.2cm 以上あく
 *   （チップを 0.80in にすると左端が 8.62in になり、lint の TITLEパンくず侵入になる）。
 * - 上端 0.50in: チップ列の上下の中央（0.82in）をタイトル枠の上下の中央（0.825in）にそろえる。下端 1.14in は金下線の上 0.3in。
 * - 現在章の地は濃いティール 007C80（白太字 12pt で約 5.0:1。C.teal 00A8AA だと約 2.9:1）。
 * lint のパンくず判定（上端 < 1.9cm・高さ < 1.9cm・幅 < 2.2cm・左 > 17cm）の範囲に収める。
 */
const CHIP = { pt: 12, w: 0.75, h: 0.64, y: 0.5, gap: 0.04, right: LAYOUT.body.right, radius: 0.08, cur: '007C80' };
const CHIP_LEFT = r3(CHIP.right - CHAPTERS.length * CHIP.w - (CHAPTERS.length - 1) * CHIP.gap);
/** タイトル枠の幅。チップ列の左端の 0.05in 手前で止める（lint の TITLE2行化を避けるため、S2 の推定幅 8.21in 以上は要る） */
const TITLE_W = r2(CHIP_LEFT - 0.05 - LAYOUT.title.x);

function restyleChips(s) {
  const ch = LAYOUT.chip;
  const chips = s.raw._slideObjects.filter(o => o._type === 'text' && o.options && o.options.shape === 'roundRect'
    && Math.abs(o.options.y - ch.y) < 1e-6 && Math.abs(o.options.h - ch.h) < 1e-6);
  if (chips.length !== CHAPTERS.length) throw new Error(`p${s.page}: 章チップが見つかりません（${chips.length} 個）`);
  const cur = chips.filter(o => o.options.fill && String(o.options.fill.color).toUpperCase() === C.teal);
  if (cur.length !== 1) throw new Error(`p${s.page}: 現在章のチップが 1 個ではありません（${cur.length} 個）`);
  const inner = CHIP.w - 2 * CHIP.radius * 0.29289;   // 角丸の文字領域（角の丸みぶん内側）
  let x = CHIP_LEFT;
  chips.forEach((o, k) => {
    const ls = CHAPTERS[k].split('\n');
    // 実寸（メイリオ全角 = 1em）で文字領域に入ること（折り返しはしない設定だが、箱からはみ出さないことを確かめる）
    ls.forEach(l => {
      const realW = [...l].length * CHIP.pt / 72;
      if (realW > inner) throw new Error(`章チップ「${l}」が ${CHIP.pt}pt・幅 ${CHIP.w}in に収まりません（${r3(realW)} > ${r3(inner)}in）`);
    });
    if (ls.length * CHIP.pt * 1.5 / 72 > CHIP.h - 0.05) throw new Error(`章チップ「${CHAPTERS[k]}」が高さ ${CHIP.h}in に収まりません`);
    o.options.x = r3(x);
    o.options.y = CHIP.y;
    o.options.w = CHIP.w;
    o.options.h = CHIP.h;
    o.options.rectRadius = CHIP.radius;
    o.options.fontSize = CHIP.pt;
    o.options._bodyProp = { ...(o.options._bodyProp || {}), wrap: false };
    if (o === cur[0]) o.options.fill = { color: CHIP.cur };
    o.text.forEach(t => { if (t.options) t.options.fontSize = CHIP.pt; });
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

/**
 * 出典とページ番号。decklib は出典の枠の下端を 7.35in、ページ番号の下端を 7.40in に置く（下の余白 0.1〜0.15in）。
 * 枠の上端（出典 1 行 6.85in・2 行 6.5in）は変えずに高さを詰め、下端を 7.20in（下の余白 0.30in）にそろえる。
 * 下揃えなので文字は 0.15in 上がる。枠の高さは lint の推定（1 行 0.40in、2 行 0.80in）から許容 0.3cm を引いた値以上・
 * 実機メイリオの高さ（1 行 0.33in、2 行 0.67in）以上。
 * 出典の枠の幅は 11.4in（ページ番号の枠 12.30in との間を 0.24in 以上あける）。ライブラリは 11.6in で行を組むので、
 * 推定幅が 11.4in を超える行があるときはその行が入る幅にする。
 */
const FOOT_BOTTOM = 7.2;
const SRC_W = 11.4;
function restyleFooter(s) {
  const S = LAYOUT.source, pg = LAYOUT.page;
  const objs = s.raw._slideObjects.filter(o => o._type === 'text' && o.options);
  const page = objs.filter(o => o.options.x === pg.x && o.options.y === pg.y && o.options.w === pg.w);
  if (page.length !== 1) throw new Error(`p${s.page}: ページ番号が見つかりません（${page.length} 個）`);
  page[0].options.y = r3(FOOT_BOTTOM - pg.h);
  const src = objs.filter(o => o.options.x === S.x && o.options.w === S.w && (o.options.y === S.y || o.options.y === S.y2));
  if (src.length !== 1) throw new Error(`p${s.page}: 出典が見つかりません（${src.length} 個）`);
  const o = src[0];
  const lines = o.text.map(t => t.text);
  const maxW = Math.max(...lines.map(l => measure(l, S.size)));
  const w = Math.min(S.w, Math.max(SRC_W, Math.ceil(maxW * 100) / 100));
  const h = r2(FOOT_BOTTOM - o.options.y);
  const lintH = lines.length * (S.size / 72 * 1.2 * 1.18 + 6 / 72) - 0.3 / CM;
  if (h + 1e-9 < Math.max(lintH, lines.length * S.size * 1.5 / 72)) throw new Error(`p${s.page}: 出典の枠 ${h}in が低すぎます`);
  if (o.options.x + w > pg.x - GAP + 1e-9) throw new Error(`p${s.page}: 出典の枠の右端 ${r3(o.options.x + w)}in がページ番号に近すぎます`);
  o.options.h = h;
  o.options.w = w;
  const it = s._items.find(i => i.desc === '出典');
  if (it) { it.h = h; it.w = w; }
  return s;
}

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す）＋チップ・出典・ページ番号の上書きと、縮めたタイトル枠に対する長さの検査 */
function contentSlide(deck, id, chapter) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id) });
  const right = LAYOUT.title.x + measure(sp.title, LAYOUT.title.size);
  const limit = Math.min(LAYOUT.title.x + TITLE_W, CHIP_LEFT - 0.2 / CM);
  if (right > limit + 1e-9) throw new Error(`${id}「${sp.title}」: タイトルの推定右端 ${r3(right)}in > 上限 ${r3(limit)}in`);
  restyleChips(s);
  restyleFooter(s);
  return s;
}

/**
 * 見出し＋項目のカード。項目が2行になるときは、2行目以降を1字ぶん字下げする（ぶら下げ）。
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
  s.box({ x, y, w, h, fill: C.bg, radius: 0.08 });
  const iw = w - 2 * pad;
  const th = heightFor(1, ts, 0);
  s.text(o.title, { x: x + pad, y: y + padY, w: iw, h: th, size: ts, bold: true, color: o.titleColor });
  rec(id, o.title);
  const top = y + padY + th + tg;
  const avail = y + h - padY - top;
  const itemH = o.items.map(lines => heightFor(1, bs, 0) + (lines.length > 1 ? heightFor(lines.length - 1, bs, 0) : 0));
  const sum = itemH.reduce((a, v) => a + v, 0);
  const n = o.items.length;
  let g = o.gap != null ? o.gap : (n > 1 ? Math.min(o.maxGap, (avail - sum) / (n - 1)) : 0);
  if (sum + g * (n - 1) > avail + 1e-6 || (o.gap == null && n > 1 && g < o.minGap - 1e-6)) {
    throw new Error(`${id} カード「${textOf(o.title)}」: 高さ不足（項目 ${sum.toFixed(2)}in + 間隔 > ${avail.toFixed(2)}in）`);
  }
  // center: 項目の間隔を maxGap までに抑え、余った高さは見出しの下の領域で上下に等分する（項目の少ない列の下が空かないように）
  let cy = top + (o.center ? Math.max(0, (avail - sum - g * (n - 1)) / 2) : 0);
  o.items.forEach((lines, k) => {
    s.text([lines[0]], { x: x + pad, y: cy, w: iw, h: heightFor(1, bs, 0), size: bs, color: C.dark });
    rec(id, lines[0]);
    if (lines.length > 1) {
      const rest = lines.slice(1);
      s.text(rest, { x: x + pad + indent, y: cy + heightFor(1, bs, 0), w: iw - indent, h: heightFor(rest.length, bs, 0), size: bs, color: C.dark });
      rest.forEach(l => rec(id, l));
    }
    cy += itemH[k] + (k < n - 1 ? g : 0);
  });
  return { gap: g };
}

/**
 * 3列カードの共通寸法（本文 22pt。項目の2行目は 0.3in 字下げ）。左右の内側余白は S2 のカードと同じ 0.18in。
 * 項目の間隔は列ごとに、カードの高さに合わせて 0.08〜0.45in で広げ、余りは上下に等分する（center）。
 */
const CARD3 = { pad: 0.18, padY: 0.14, titleSize: 24, bodySize: 22, titleGap: 0.12, indent: 0.3, minGap: 0.08, maxGap: 0.45, center: true };

/**
 * 3列カード。列幅は中身の最長行から決め、余りは狭い列から順に配る（水位合わせ）。
 * cards: [{title, titleColor, items}], 下端 bottom
 */
function threeCards(s, id, cards, bottom) {
  const total = BW - 2 * GAP;
  const need = cards.map(c => itemCardNeedW({ ...CARD3, ...c }));
  if (need.reduce((a, v) => a + v, 0) > total + 1e-9) {
    throw new Error(`${id}: 3列カードの必要幅 ${need.map(v => v.toFixed(2)).join(' + ')} > ${total.toFixed(2)}in`);
  }
  // 水位合わせ: 幅 = max(need, W)、Σ幅 = total
  let fixed = new Set();
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
    itemCard(s, id, { ...CARD3, ...c, x, y: TOP, w: widths[i], h });
    x += widths[i] + GAP;
  });
}

/** 強調ボックス（白地・ティールの枠・太字） */
function keyBox(s, id, line, o) {
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.white, line: { color: C.teal, width: 2.5 }, radius: 0.08,
    text: [line], size: o.size || 28, bold: true, color: C.dark, align: o.align || 'center' });
  rec(id, line);
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: CHAPTERS,
    author: '初期研修医向け麻酔講義',
    deckTitle: '3-2 静脈麻酔薬',
  });
  shrinkTitleBox(deck);

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 静脈麻酔薬に共通すること（左右2枚の大きなカード＋下に一文） =====
  {
    const id = 'S2';
    const s = contentSlide(deck, id, 0);
    const cw = (BW - GAP) / 2;
    const ch = 3.5;
    const geo = { y: TOP, w: cw, h: ch, pad: 0.18, padY: 0.2, titleSize: 28, bodySize: 24, titleGap: 0.2, indent: 0.3, gap: 0.3 };
    itemCard(s, id, {
      ...geo, x: X0, title: '作用', titleColor: C.blue,
      items: [
        ['GABA 受容体を介して意識を消す', '（ケタミンは NMDA 受容体）'],
        ['鎮痛作用はない', '（ケタミンを除く）'],
      ],
    });
    itemCard(s, id, {
      ...geo, x: X0 + cw + GAP, title: '共通の副作用', titleColor: C.vermilion,
      items: [
        ['血圧低下（血管拡張・心抑制）'],
        ['呼吸抑制・無呼吸'],
        ['上気道の閉塞（舌根沈下）'],
      ],
    });
    keyBox(s, id, '入れたら効果のピークまで待ち、反応を見て追加する', { x: X0, y: TOP + ch + 0.32, w: BW, h: 0.85 });
  }

  // ===== S3 プロポフォール（3列カード） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 1);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          [['導入：', b('0.5 mg/kg/10 秒'), 'の'], '速度で、就眠まで'],
          [['維持：', b('4〜10 mg/kg/時')]],
          ['ASA III・IV は', 'より緩徐に'],
        ],
      },
      {
        title: '特徴', titleColor: C.blue,
        items: [
          ['覚醒が速く、', 'すっきり'],
          ['PONV が少ない'],
          ['血圧低下・', '呼吸抑制が強い'],
          ['注射痛'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['脂肪乳剤：清潔操作、', '開封後すぐ使う'],
          ['小児の ICU 人工呼吸中の', ['鎮静は', danger('禁忌')]],
          ['本剤・成分への', ['過敏症の既往（', danger('禁忌'), '）']],
        ],
      },
    ], SRC2_TOP - GAP);
  }

  // ===== S4 レミマゾラム（3列カード） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 2);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          [['導入：', b('12 mg/kg/時'), 'で、'], '意識消失まで'],
          [['維持：', b('1 mg/kg/時'), 'から、'], ['上限 ', b('2 mg/kg/時')]],
          ['年齢・状態に応じて減速'],
        ],
      },
      {
        title: '特徴', titleColor: C.blue,
        items: [
          ['超短時間作用型', 'ベンゾジアゼピン'],
          ['主に肝のエステラーゼ', '（CES1）で分解'],
          ['プロポフォールより', '血圧低下が少ない'],
          ['フルマゼニルで拮抗'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['リンゲル液で白濁・', '閉塞（配合変化）'],
          ['覚醒が遅れる', 'ことがある'],
          ['拮抗後の再鎮静'],
        ],
      },
    ], SRC2_TOP - GAP);
  }

  // ===== S5 ミダゾラム（3列カード） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 3);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          ['前投薬：', [b('0.08〜0.10 mg/kg'), ' 筋注']],
          ['導入：', [b('0.15〜0.30 mg/kg'), ' 静注']],
          ['鎮静：少量ずつ', '（ICU・歯科が適応）'],
        ],
      },
      {
        title: '特徴', titleColor: C.blue,
        items: [
          ['健忘・抗不安'],
          ['作用がやや長い'],
          ['肝代謝', '（CYP3A4）'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['オピオイド併用で呼吸抑制'],
          ['高齢者は少量から'],
          ['フルマゼニルで拮抗'],
          [[danger('併用禁忌'), '：'], 'リトナビル含有製剤など'],
        ],
      },
    ], SRC2_TOP - GAP);
  }

  // ===== S6 ケタミン（3列カード） =====
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 3);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          [['初回 ', b('1〜2 mg/kg'), ' を'], '1 分以上かけて静注'],
          ['追加：初回量と', '同量または半量'],
        ],
      },
      {
        title: '特徴', titleColor: C.blue,
        items: [
          ['鎮痛作用がある'],
          ['血圧・呼吸が', '保たれやすい'],
          ['気道反射が', '比較的保たれる'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['分泌物の増加'],
          ['覚醒時の夢・興奮', '（ベンゾジアゼピン併用）'],
          ['麻薬として管理'],
          [['脳圧亢進・高血圧は', danger('禁忌')]],
        ],
      },
    ], SRC2_TOP - GAP);
  }

  // ===== S7 使い分け（表 2列 × 6行＋見出し行） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 4);
    const body = [
      ['通常の導入', 'プロポフォール、またはレミマゾラム'],
      ['循環が不安定・出血', 'ケタミン減量、またはレミマゾラムを減速'],
      ['高齢者', 'どの薬も減量し、ゆっくり'],
      ['PONV のリスクが高い', 'プロポフォールで維持（TIVA）'],
      ['拮抗薬が欲しい', 'レミマゾラム（フルマゼニル）'],
      ['喘息', 'プロポフォール、ケタミン'],
    ];
    const head = ['状況', '選び方'];
    head.forEach(t => rec(id, t));
    body.forEach(r => r.forEach(t => rec(id, t)));
    s.table([head, ...body.map(([a, c]) => [{ text: a, bold: true, fill: C.bg }, c])], {
      x: X0, y: TOP, colW: [4.0, BW - 4.0], size: 24, rowH: 0.62, headAlign: 'left',
    });
  }

  // ===== S8 導入時の低血圧に備える（箇条書き4行＋強調ボックス） =====
  {
    const id = 'S8';
    const s = contentSlide(deck, id, 4);
    const items = [
      '静脈麻酔薬は血管拡張と心抑制で血圧を下げる',
      '脱水、高齢、ARB・ACE 阻害薬の内服、出血で強く出る',
      'ゆっくり分割投与し、輸液を進め、昇圧薬を手元に',
      '挿管の刺激で一時的に上がり、そのあとまた下がる',
    ];
    items.forEach(t => rec(id, t));
    // 箇条書き（実機メイリオで下端 約 5.07in）と強調ボックスの間、ボックスと出典の間をどちらも約 0.3in にする
    const by = 1.9, gap = 20, kh = 0.85;
    s.bullets(items, { y: by, size: 28, gap });
    const bh = heightFor(items.length, 28, gap);
    const ky = r2(by + bh + GAP);
    // 強調は数値と単位だけ（バーミリオン＋下線。色だけに頼らない）
    keyBox(s, id, ['導入薬を入れたら、血圧は ', { text: '1〜2 分', color: C.vermilion, underline: true }, 'ごとに測る'],
      { x: X0, y: ky, w: BW, h: kh });
    if (ky + kh > SRC2_TOP - GAP + 1e-6) throw new Error('S8: 強調ボックスが出典に近すぎます');
  }

  // ===== S9 まとめ（箇条書き5行） =====
  {
    const id = 'S9';
    const s = contentSlide(deck, id, 4);
    const items = [
      [drug('プロポフォール'), '：0.5 mg/kg/10 秒で導入、維持 4〜10 mg/kg/時'],
      [drug('レミマゾラム'), '：12 mg/kg/時で導入、維持 1 mg/kg/時'],
      [drug('ケタミン'), '：1〜2 mg/kg、鎮痛あり、循環が保たれやすい'],
      'どの薬も、少量から、待って、反応を見て追加',
      '導入後は血圧を 1〜2 分ごとに、昇圧薬を手元に',
    ];
    items.forEach(t => rec(id, t));
    // 1 項目目（プロポフォール）は 28pt・26pt では 1 行に入らない（推定 12.8in・11.9in > 使える幅 11.8in）ので、
    // 5 項目とも 25pt（推定 11.4in）にそろえる。行間を広げて、出典の上（実機メイリオで下端 約 6.17in）まで使う
    s.bullets(items, { y: 1.9, w: BW, size: 25, gap: 30 });
  }

  checkOnSlide();
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
