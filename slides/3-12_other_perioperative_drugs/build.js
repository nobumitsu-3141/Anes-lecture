'use strict';
/**
 * 3-12 周術期のその他の薬
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk()/q() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, measure, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '3-12_other_perioperative_drugs';
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
        for (const q of ln.matchAll(/「([^」]*)」/g)) quoted.add(q[1]);
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
/**
 * spec の文言 str を、breaks の各語の直前で改行して行配列にする（文言は変えない）。
 * 改行位置の半角スペースは落とす。
 */
function brk(id, str, ...breaks) {
  q(id, str);
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
const r2 = v => Math.round(v * 100) / 100;
/** 箇条書きスライドの強調ボックスの下端の上限（出典2行・枠 y6.5 の上） */
const BOTTOM_2SRC = 6.32;
/** 強調ボックス（S2・S4・S8 で位置・幅・線を揃える） */
const EMPH = { x: X0 + 0.6, w: BW - 1.2, radius: 0.1, lineW: 3.5, size: 28 };

// ---------------------------------------------------------------------------
// 章チップ
// decklib の章チップ（0.68in 幅・角丸 0.1in・11pt・余白 0）は、角丸のぶん文字領域が約 44.7pt しかなく、
// 「抗コリン」（全角 4 字 = 44pt）が「抗コリ / ン」に割れる。このデッキでは位置・寸法・色は変えずに、
// 角丸を 0.05in（adj 約 7800）、文字を 10.5pt にする（文字領域 約 46.9pt に対し 4 字 = 42pt）。
// 書き出し前の pptxgenjs のスライドオブジェクトを直接書き換える（2-6 の darkenActiveChip と同じ方法）。
// ---------------------------------------------------------------------------
const CHIP = { size: 10.5, radius: 0.05 };
function tuneChips(s, n) {
  const ch = LAYOUT.chip;
  const chips = s.raw._slideObjects.filter(o => o._type === 'text' && o.options && o.options.shape === 'roundRect'
    && Math.abs(o.options.y - ch.y) < 1e-6 && Math.abs(o.options.h - ch.h) < 1e-6 && Math.abs(o.options.w - ch.w) < 1e-6);
  if (chips.length !== n) throw new Error(`p${s.page}: 章チップが ${chips.length} 個（${n} 個のはず）`);
  for (const o of chips) {
    o.options.rectRadius = CHIP.radius;
    o.options.fontSize = CHIP.size;
    for (const r of o.text) if (r.options) r.options.fontSize = CHIP.size;
    // 4 字の行が 1 行に収まること（角丸の内接ぶんを引いた文字領域で判定。全角 1em）
    const inner = (ch.w - 2 * CHIP.radius * (1 - Math.SQRT1_2)) * 72;
    for (const r of o.text) {
      const need = [...String(r.text)].length * CHIP.size;
      if (need > inner - 2) throw new Error(`p${s.page}: 章チップ「${r.text}」${need}pt > 文字領域 ${inner.toFixed(1)}pt`);
    }
  }
  return s;
}

// ---------------------------------------------------------------------------
// 実機メイリオの字幅の目安（警告・幅の配分用）。
// ライブラリの推定（和文 1.12em・欧文 0.627em）は和文に厚く、欧文・数字の多い行は余裕が小さい。
// メイリオの欧文は Verdana 系で幅が広いので、その字幅で別に見積もり、カード幅の配分に使う。
// ---------------------------------------------------------------------------
const LATIN_EM = {
  ' ': 0.352, '.': 0.364, ',': 0.364, ':': 0.454, ';': 0.454, '-': 0.454, '/': 0.454, '%': 1.076, '(': 0.454, ')': 0.454,
  0: 0.636, 1: 0.636, 2: 0.636, 3: 0.636, 4: 0.636, 5: 0.636, 6: 0.636, 7: 0.636, 8: 0.636, 9: 0.636,
  a: 0.601, b: 0.623, c: 0.521, d: 0.623, e: 0.596, f: 0.352, g: 0.623, h: 0.633, i: 0.274, j: 0.344, k: 0.592, l: 0.274,
  m: 0.973, n: 0.633, o: 0.607, p: 0.623, q: 0.623, r: 0.427, s: 0.521, t: 0.394, u: 0.633, v: 0.592, w: 0.818,
  x: 0.592, y: 0.592, z: 0.525,
  A: 0.684, B: 0.686, C: 0.698, D: 0.771, E: 0.632, F: 0.575, G: 0.775, H: 0.751, I: 0.421, J: 0.455, K: 0.693,
  L: 0.557, M: 0.843, N: 0.748, O: 0.787, P: 0.603, Q: 0.787, R: 0.695, S: 0.684, T: 0.616, U: 0.732, V: 0.684,
  W: 0.989, X: 0.685, Y: 0.615, Z: 0.685,
};
function realW(text, size) {
  let em = 0;
  for (const ch of String(text)) {
    const o = ch.codePointAt(0);
    em += o <= 0x24F ? (LATIN_EM[ch] != null ? LATIN_EM[ch] : 0.65) : 1.0;
  }
  return em * size / 72;
}

// ---------------------------------------------------------------------------
// カード（S3・S5・S6・S7 で共通）
// ---------------------------------------------------------------------------
const CARD = { padX: 0.2, padY: 0.15, radius: 0.08, titleSize: 24, size: 22, realSlack: 0.95 };
const CARD_BOTTOM_2SRC = 6.4;    // 出典2行（枠 y6.5、文字の上端 約6.68in）のときのカード下端
const CARD_BOTTOM_1SRC = 6.75;   // 出典1行（枠 y6.85、文字の上端 約7.02in）のときのカード下端
const titleH = n => heightFor(n, CARD.titleSize, 0);
const itemH = lines => heightFor(lines.length, CARD.size, 0);

/** カード1枚: 箱＋見出し＋項目。文字はすべて箱の内側。c.itemYs は本文の開始 y からの各項目の上端 */
function itemCard(s, c) {
  s.box({ x: c.x, y: c.y, w: c.w, h: c.h, fill: C.bg, radius: CARD.radius });
  const iw = r3(c.w - 2 * CARD.padX), tx = r3(c.x + CARD.padX);
  s.text(c.title, { x: tx, y: r3(c.y + CARD.padY), w: iw, h: titleH(c.title.length), size: CARD.titleSize, bold: true, color: c.titleColor });
  const by = c.y + CARD.padY + titleH(c.titleBand) + c.titleGap;
  c.items.forEach((lines, i) => s.text(lines, { x: tx, y: r3(by + c.itemYs[i]), w: iw, h: itemH(lines), size: CARD.size }));
  const last = c.items.length - 1;
  const end = by + c.itemYs[last] + itemH(c.items[last]);
  if (end > c.y + c.h - CARD.padY + 1e-6) throw new Error(`p${s.page} カード「${runText(c.title[0])}」: 高さ不足（本文の下端 ${r3(end)}in > ${r3(c.y + c.h - CARD.padY)}in）`);
}

/**
 * 横に並べるカード。
 *  - 幅: 各カードの最長行（ライブラリ推定と、実機推定 ÷ 0.95 の大きい方）に比例して配る（どのカードも同じ割合の余裕）。
 *  - 高さ: y から bottom まで（同じスライド種別で下端を揃える）。
 *  - 本文の開始 y: 見出しの行数が違っても全カードで揃える。
 *  - itemGap: 'auto' なら最も詰まったカードに合わせて余りを配る（上限 maxItemGap）。
 *  - alignRows: 項目 i の上端を全カードで揃える（項目数が同じとき）。
 */
function cardRow(s, cards, o) {
  const n = cards.length;
  const lines = c => [...c.title.map(t => ({ t: runText(t), size: CARD.titleSize })), ...c.items.flat().map(l => ({ t: runText(l), size: CARD.size }))];
  const estNeed = cards.map(c => Math.max(...lines(c).map(l => measure(l.t, l.size))));
  const realNeed = cards.map(c => Math.max(...lines(c).map(l => realW(l.t, l.size))));
  const availInner = BW - o.gap * (n - 1) - 2 * CARD.padX * n;
  let need = cards.map((c, i) => Math.max(estNeed[i], realNeed[i] / CARD.realSlack));
  if (need.reduce((a, b) => a + b, 0) > availInner) need = estNeed.slice();
  const k = availInner / need.reduce((a, b) => a + b, 0);
  if (k < 1 - 1e-9) throw new Error(`p${s.page}: カード幅が足りません（${r3((1 - k) * need.reduce((a, b) => a + b, 0))}in 不足）`);
  const inner = need.map(v => Math.floor(v * k * 1000) / 1000);
  cards.forEach((c, i) => {
    if (estNeed[i] > inner[i] + 1e-9) throw new Error(`p${s.page} カード${i + 1}: 推定幅 ${r3(estNeed[i])} > 内側 ${inner[i]}`);
  });
  const band = Math.max(...cards.map(c => c.title.length));
  const h = r3(o.bottom - o.y);
  // 項目の相対位置（itemGap は後で決める）
  const nItems = cards[0].items.length;
  if (o.alignRows && cards.some(c => c.items.length !== nItems)) throw new Error(`p${s.page}: alignRows は項目数が同じカードのみ`);
  const rowH = o.alignRows ? Array.from({ length: nItems }, (_, r) => Math.max(...cards.map(c => itemH(c.items[r])))) : null;
  const heights = c => (o.alignRows ? rowH : c.items.map(itemH));
  const fixed = c => 2 * CARD.padY + titleH(band) + o.titleGap + heights(c).reduce((a, b) => a + b, 0);
  let gapItems = o.itemGap;
  if (gapItems === 'auto') {
    gapItems = Math.min(o.maxItemGap, ...cards.map(c => (c.items.length > 1 ? (h - fixed(c)) / (c.items.length - 1) : Infinity)));
    gapItems = Math.floor(gapItems * 100) / 100;
  }
  if (gapItems < 0.08 - 1e-9) throw new Error(`p${s.page}: 項目の間隔が ${r3(gapItems)}in（0.08in 未満）`);
  let x = X0;
  const info = [];
  cards.forEach((c, i) => {
    const w = r3(inner[i] + 2 * CARD.padX);
    const hs = heights(c);
    const itemYs = hs.map((_, r) => hs.slice(0, r).reduce((a, b) => a + b, 0) + r * gapItems);
    itemCard(s, { ...c, x: r3(x), y: o.y, w, h, titleBand: band, titleGap: o.titleGap, itemYs });
    info.push(`${w}in（欧文込み実機推定 ${Math.round(100 * realNeed[i] / inner[i])}%）`);
    x += w + o.gap;
  });
  console.log(`  p${s.page} カード: 幅 ${info.join(' / ')}、高さ ${h}in（下端 ${r3(o.y + h)}in）、項目間 ${r2(gapItems)}in`);
}

/** 強調ボックス（1〜2行・28pt 太字・白地＋色枠）。y は上端 */
function emphBox(s, y, h, lines, lineColor) {
  s.box({
    x: EMPH.x, y: r3(y), w: EMPH.w, h, fill: C.white, radius: EMPH.radius,
    line: { color: lineColor, width: EMPH.lineW },
    text: lines, size: EMPH.size, bold: true, color: C.dark,
  });
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['予防的\n抗菌薬', '抗コリン\n薬', 'ステロ\nイド', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '3-12 周術期のその他の薬',
  });
  const NCH = 4;
  const slide = o => deck.slide(o);   // 章チップはライブラリ既定（4 字ラベルの調整も decklib 側）

  // ===== S1 表紙 =====
  {
    const subtitle = '3-12 麻酔薬理', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: '周術期のその他の薬', subtitle, meta, notes: notesOf('S1') });
  }

  // 箇条書きスライド（S2・S4・S8）は 26pt・段落後 14pt・y1.95 で揃える
  const BUL = { y: 1.95, size: 26, gap: 14 };
  const bulH = n => heightFor(n, BUL.size, BUL.gap);

  // ===== S2 予防的抗菌薬の目的（箇条書き4行 + 強調ボックス） =====
  {
    const id = 'S2';
    const s = slide({ title: '予防的抗菌薬の目的', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, '手術部位感染（SSI）を減らすため。治療ではない'), ['治療ではない', C.vermilion]),
      q(id, '切開の瞬間に、組織に十分な濃度があること'),
      mark(q(id, '狙うのは皮膚の常在菌（黄色ブドウ球菌など）→ セファゾリン'), ['セファゾリン', C.blue]),
      mark(q(id, '閉創後は 24 時間以内に終える（手術によっては 48 時間）'), ['24 時間以内', C.blue]),
    ], BUL);
    // 強調ボックス（2行・28pt）。枠はティール（強調）
    const lines = brk(id, 'タイムアウト：執刀前 60 分以内に開始（バンコマイシンは 120 分）', '（バンコマイシン');
    const boxH = 1.25;
    const by = BOTTOM_2SRC - 0.06 - boxH;
    if (by < BUL.y + bulH(4) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(4))}in`);
    emphBox(s, by, boxH, [mark(lines[0], ['執刀前 60 分以内に開始', C.vermilion]), mark(lines[1], ['120 分', C.vermilion])], C.teal);
  }

  // ===== S3 用量・再投与・アレルギー（3列カード） =====
  {
    const id = 'S3';
    const s = slide({ title: '用量・再投与・アレルギー', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: brk(id, '用量（セファゾリン）', '（'), titleColor: C.blue,
        items: [
          brk(id, '成人 1 g（体重が重ければ 2 g）', '重ければ'),
          [q(id, '小児は体重あたり')],
          brk(id, '腎機能で間隔を調節', '間隔'),
        ],
      },
      {
        title: [q(id, '術中の再投与')], titleColor: C.vermilion,
        items: [
          brk(id, '長い手術では半減期の約 2 倍ごと', '約 2'),
          brk(id, 'セファゾリンは 3〜4 時間ごと', '3〜4').map((l, i) => (i === 1 ? mark(l, ['3〜4 時間ごと', C.blue]) : l)),
          brk(id, '大量出血（1,500 mL 以上）でも', '以上'),
        ],
      },
      {
        title: [q(id, 'β-ラクタムアレルギー')], titleColor: C.red,
        items: [
          brk(id, '重いアレルギー歴ならクリンダマイシンかバンコマイシン（腹部は陰性菌用を追加）', 'クリンダ', 'バンコ', '（腹部'),
          brk(id, '『ペニシリンアレルギー』の中身を聞く', 'の中身'),
        ],
      },
    ], { y: Y0, bottom: CARD_BOTTOM_2SRC, gap: GAP, titleGap: 0.12, itemGap: 'auto', maxItemGap: 0.3 });
  }

  // ===== S4 抗菌薬のアナフィラキシー（箇条書き4行 + 強調ボックス） =====
  {
    const id = 'S4';
    const s = slide({ title: '抗菌薬のアナフィラキシー', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, '周術期アナフィラキシーの原因で最も多い（英国 NAP6：47%）'), ['47%', C.vermilion]),
      q(id, '麻酔導入と同じ時間帯に入るので、原因がまぎれる'),
      mark(q(id, '投与後 20 分ほどは血圧・気道内圧・皮膚を見る'), ['20 分ほど', C.blue]),
      q(id, '抗菌薬は、できれば導入の数分前に（反応を見分けやすい）'),
    ], BUL);
    // 強調ボックス（1行・28pt）。危険なので枠と「アドレナリン」はレッド
    const boxH = 0.95;
    const by = BOTTOM_2SRC - 0.2 - boxH;
    if (by < BUL.y + bulH(4) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(4))}in`);
    emphBox(s, by, boxH, [mark(q(id, '原因不明の低血圧＋気管支攣縮 → アドレナリン'), ['アドレナリン', C.red])], C.red);
  }

  // ===== S5 アトロピン（3列カード） =====
  {
    const id = 'S5';
    const s = slide({ title: 'アトロピン', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: [q(id, '用量')], titleColor: C.blue,
        items: [
          brk(id, '添付文書：成人 0.5 mg 皮下・筋注（場合により静注）', '皮下', '（場合'),
          brk(id, '徐脈（蘇生 GL）：0.5 mg 静注を 3〜5 分ごと、総量 3 mg まで', '静注を', '総量'),
          brk(id, '小児 0.01〜0.02 mg/kg（添付文書外）', '（添付'),
        ],
      },
      {
        title: [q(id, '作用')], titleColor: C.vermilion,
        items: [
          brk(id, 'ムスカリン受容体の遮断', '受容体'),
          brk(id, '心拍数↑、分泌↓、気管支拡張', '分泌', '気管支'),
          brk(id, '静注で 1〜2 分で効き始める', '効き'),
        ],
      },
      {
        title: [q(id, '注意')], titleColor: C.red,
        items: [
          brk(id, '成人で 0.5 mg 未満だと逆に徐脈になりうる', '逆に').map((l, i) => (i === 0 ? mark(l, ['0.5 mg 未満', C.red]) : l)),
          brk(id, '閉塞隅角緑内障、前立腺肥大による排尿障害、麻痺性イレウスは禁忌', '前立腺', '排尿', '麻痺性')
            .map((l, i) => (i === 3 ? mark(l, ['禁忌', C.red]) : l)),
          brk(id, '高齢者のせん妄、小児の発熱', '小児'),
        ],
      },
    ], { y: Y0, bottom: CARD_BOTTOM_1SRC, gap: GAP, titleGap: 0.1, itemGap: 'auto', maxItemGap: 0.3 });
  }

  // ===== S6 抗コリン薬の使いどころ（左右2枚のカード。項目の上端を左右で揃える） =====
  {
    const id = 'S6';
    const s = slide({ title: '抗コリン薬の使いどころ', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: [q(id, '使う場面')], titleColor: C.blue,
        items: [
          brk(id, '迷走神経反射の徐脈（眼球圧迫、腹膜牽引、気腹）', '（眼球'),
          [q(id, 'ネオスチグミンの併用（→ 3-6）')],
          [q(id, '気道操作・ケタミン前の分泌抑制')],
        ],
      },
      {
        title: [q(id, '考えること')], titleColor: C.vermilion,
        items: [
          brk(id, '徐脈の原因（低酸素、麻酔薬、刺激）を外すのが先', '（低酸素', '外す').map((l, i) => (i === 2 ? mark(l, ['外すのが先', C.vermilion]) : l)),
          brk(id, '心拍数だけ上げても血圧が戻らないことがある', '血圧'),
          [q(id, 'エフェドリンとの使い分け（→ 3-8）')],
        ],
      },
    ], { y: Y0, bottom: CARD_BOTTOM_2SRC, gap: 0.32, titleGap: 0.16, itemGap: 'auto', maxItemGap: 0.45, alignRows: true });
  }

  // ===== S7 ステロイドの 3 つの場面（3列カード） =====
  {
    const id = 'S7';
    const s = slide({ title: 'ステロイドの 3 つの場面', chapter: 2, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: [q(id, 'PONV の予防')], titleColor: C.blue,
        items: [
          brk(id, 'デキサメタゾン 3.3〜6.6 mg（海外の 4〜8 mg）を導入時に', '3.3', '（海外', '導入時'),
          brk(id, '効くまで時間がかかるので早めに', '早めに'),
          [q(id, '国内では適応外（→ 5-4）')],
        ],
      },
      {
        title: [q(id, 'ステロイドカバー')], titleColor: C.vermilion,
        items: [
          brk(id, '長期内服で副腎が抑制された患者', '抑制'),
          brk(id, 'ヒドロコルチゾン 100 mg を導入時', '100 mg').map((l, i) => (i === 1 ? mark(l, ['100 mg', C.blue]) : l)),
          brk(id, '以後 200 mg/24 時間の持続（倍量を内服できるまで）', 'の持続', '内服'),
        ],
      },
      {
        title: [q(id, '気道浮腫・補助')], titleColor: C.red,
        items: [
          brk(id, '抜管後の気道浮腫の予防・治療', '予防'),
          brk(id, 'アナフィラキシーの二次治療（効果は遅い）', '二次', '（効果'),
          [q(id, '第一選択ではない')],
        ],
      },
    ], { y: Y0, bottom: CARD_BOTTOM_2SRC, gap: GAP, titleGap: 0.12, itemGap: 'auto', maxItemGap: 0.3 });
  }

  // ===== S8 ステロイドの注意（箇条書き4行 + 結論は強調ボックス） =====
  {
    const id = 'S8';
    const s = slide({ title: 'ステロイドの注意', chapter: 2, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, '高血糖：糖尿病患者では術後の血糖を見る'), ['高血糖：', C.blue]),
      mark(q(id, '創感染：単回投与で明らかな増加はないとされる'), ['創感染：', C.blue]),
      mark(q(id, '急速静注で会陰部の灼熱感（デキサメタゾン）'), ['急速静注で', C.blue]),
      mark(q(id, 'カバー不足は副腎クリーゼ、過剰は高血糖・感染'), ['副腎クリーゼ', C.red]),
    ], BUL);
    // 結論（5 項目め）は黒の太字。ティールは枠線だけに使う（文字にすると白地で約 2.9:1 と薄い）
    const boxH = 0.95;
    const by = BOTTOM_2SRC - 0.2 - boxH;
    if (by < BUL.y + bulH(4) + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - BUL.y - bulH(4))}in`);
    emphBox(s, by, boxH, [q(id, '迷ったら、カバーする方が安全')], C.teal);
  }

  // ===== S9 まとめ（箇条書き5行） =====
  // 最長行（デキサメタゾン…）が 26pt では 1 行に入らないため 24pt（1 バレット 1 行を優先）
  {
    const id = 'S9';
    const s = slide({ title: 'まとめ', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, 'セファゾリン：執刀前 60 分以内開始、3〜4 時間ごと、24 時間以内'), ['セファゾリン：', C.blue]),
      mark(q(id, '抗菌薬は周術期アナフィラキシーの主な原因。投与後 20 分ほど見る'), ['抗菌薬', C.blue]),
      mark(q(id, 'アトロピン 0.5 mg：少量で逆に徐脈、禁忌を覚える'), ['アトロピン 0.5 mg：', C.blue]),
      mark(q(id, 'デキサメタゾン 3.3〜6.6 mg（海外の 4〜8 mg）を導入時に。適応外'), ['デキサメタゾン 3.3〜6.6 mg', C.blue]),
      mark(q(id, 'ステロイドカバー：ヒドロコルチゾン 100 mg、迷ったらカバー'), ['ステロイドカバー：', C.blue]),
    ], { y: 2.0, size: 24, gap: 28 });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
