'use strict';
/**
 * 3-8 昇圧薬の使い分け
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
const { Deck, C, measure, fits, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '3-8_vasopressors';
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

// 出典（16pt・幅 11.6in・最大2行）。ライブラリは ' / ' の区切りでしか改行しないので、
// 区切りどおりだと3行になる場合だけ、文言はそのままで改行位置（半角スペース）を選んで2行にする。
const SRC = LAYOUT.source;
const SOURCE_SPLITS = {};
function libLines(items) {
  const lines = [];
  for (const it of items) {
    const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
    if (cur && fits(cur, SRC.size, SRC.w)) lines[lines.length - 1] = cur;
    else lines.push(it);
  }
  return lines;
}
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  if (!s) return undefined;
  const items = s.split(' / ');
  if (items.every(it => fits(it, SRC.size, SRC.w)) && libLines(items).length <= 2) return items;
  // 2行に詰める改行位置を探す: (0) ' / ' の区切り (1) 「et al.」の直後 (2) その他の半角スペース
  let best = null;
  for (let p = 0; p < s.length; p++) {
    if (s[p] !== ' ') continue;
    if (s[p + 1] === '/') continue;                    // '/' を行頭に置かない
    let l1, pri;
    if (s.slice(p - 2, p + 1) === ' / ') { l1 = s.slice(0, p - 2); pri = 0; }
    else { l1 = s.slice(0, p); pri = /et al\.$/.test(l1) ? 1 : 2; }
    const l2 = s.slice(p + 1);
    if (!fits(l1, SRC.size, SRC.w) || !fits(l2, SRC.size, SRC.w)) continue;
    if (!best || pri < best.pri || (pri === best.pri && l1.length > best.l1.length)) best = { l1, l2, pri };
  }
  if (!best) throw new Error(`${id}: 出典が2行に収まりません`);
  SOURCE_SPLITS[id] = [best.l1, best.l2];
  return [best.l1, best.l2];
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
/** 行配列のうち、語 w を含む行にだけ mark を掛ける */
function markIn(lines, w, color) {
  let n = 0;
  const out = lines.map(l => {
    if (typeof l === 'string' && l.includes(w)) { n++; return mark(l, [w, color]); }
    return l;
  });
  if (n !== 1) throw new Error(`markIn: 「${w}」が ${n} 行に`);
  return out;
}
const runText = l => (Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(l));

// ---------------------------------------------------------------------------
// 共通の寸法
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const Y0 = LAYOUT.body.y;        // 1.85
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const r3 = v => Math.round(v * 1000) / 1000;
const BOTTOM_2SRC = 6.42;        // 出典2行（枠 y6.5）のときの本文の下端
const BOTTOM_1SRC = 6.75;        // 出典1行（枠 y6.85）のときの本文の下端
const bottomOf = s => (s.hasSources && s.sourceTop < 6.8 ? BOTTOM_2SRC : BOTTOM_1SRC);

// ---------------------------------------------------------------------------
// 見出し＋項目のカード（各項目は1〜数行。項目の間は細い区切り線）。文字はすべて箱の内側。
// ---------------------------------------------------------------------------
// lineGap: 項目内の段落間（pt）。0 だと lint が 6pt と仮定して高さを多めに見積もるので 1pt を明示する
const CARD = { pad: 0.14, titleSize: 24, size: 22, titleGap: 0.08, itemGap: 0.2, lineGap: 1 };
function itemCardMinW(c) {
  let w = measure(runText(c.title), CARD.titleSize);
  c.items.forEach(lines => lines.forEach(l => { w = Math.max(w, measure(runText(l), CARD.size)); }));
  return w + 2 * CARD.pad + 0.005;
}
function itemCardMinH(c) {
  return 2 * CARD.pad + heightFor(1, CARD.titleSize, 0) + CARD.titleGap
    + c.items.reduce((a, lines) => a + heightFor(lines.length, CARD.size, CARD.lineGap), 0) + CARD.itemGap * (c.items.length - 1);
}
function itemCard(s, c, x, y, w, h) {
  const pad = CARD.pad;
  s.box({ x, y, w, h, fill: C.bg, radius: 0.08 });
  const iw = r3(w - 2 * pad);
  const th = heightFor(1, CARD.titleSize, 0);
  let cy = y + pad;
  s.text(c.title, { x: x + pad, y: r3(cy), w: iw, h: th, size: CARD.titleSize, bold: true, color: c.titleColor });
  cy += th + CARD.titleGap;
  c.items.forEach((lines, i) => {
    const ih = heightFor(lines.length, CARD.size, CARD.lineGap);
    s.text(lines, { x: x + pad, y: r3(cy), w: iw, h: ih, size: CARD.size, gap: CARD.lineGap });
    cy += ih;
    if (i < c.items.length - 1) {
      s.hline({ x: x + pad, y: r3(cy + CARD.itemGap / 2), w: iw, color: C.border, width: 1 });
      cy += CARD.itemGap;
    }
  });
  if (cy > y + h - pad + 1e-6) throw new Error(`itemCard「${c.title}」: 高さ不足（必要 ${r3(cy + pad - y)}in > ${h}in）`);
}
/** 横に並べるカード。各カードの最小幅（最長行）を満たし、余りを均等に配る。高さは共通。 */
function cardRow(s, cards, o = {}) {
  const gap = o.gap != null ? o.gap : GAP;
  const minW = cards.map(itemCardMinW);
  const free = BW - gap * (cards.length - 1) - minW.reduce((a, b) => a + b, 0);
  if (free < -1e-9) throw new Error(`p${s.page}: カード幅が足りません（${r3(-free)}in 不足。最小幅 ${minW.map(r3).join(', ')}）`);
  const needH = Math.max(...cards.map(itemCardMinH));
  const h = Math.ceil(needH * 100) / 100;
  const y = o.y != null ? o.y : Y0;
  const bottom = bottomOf(s);
  if (y + h > bottom + 1e-9) throw new Error(`p${s.page}: カード下端 ${r3(y + h)}in > 下限 ${bottom}in`);
  let x = X0;
  cards.forEach((c, i) => {
    const w = minW[i] + free / cards.length;
    itemCard(s, c, r3(x), y, r3(w), h);
    x += w + gap;
  });
  console.log(`  p${s.page}: cards w=${minW.map((m) => r3(m + free / cards.length)).join('/')} (free ${r3(free)}) h=${h} → bottom ${r3(y + h)}`);
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['4つの\n要素', 'フェニレ\nフリン', 'エフェ\nドリン', 'ノルアド\nレナリン', '使い分け'],
    author: 'Anes-lecture',
    deckTitle: '3-8 昇圧薬の使い分け',
  });

  // ===== S1 表紙 =====
  {
    const subtitle = '3-8 麻酔薬理', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    if (!SPEC_MD.includes('title=「昇圧薬の使い分け」')) throw new Error('S1: 表紙タイトルが spec と違います');
    deck.cover({ title: '昇圧薬の使い分け', subtitle, meta, notes: notesOf('S1') });
  }

  // ===== S2 血圧は何で決まるか（式2段 + 4要素チップ + 一文） =====
  {
    const id = 'S2';
    const s = deck.slide({ title: '血圧は何で決まるか', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    const eqSize = 40;
    const eqH = heightFor(1, eqSize, 0);
    s.text([mark(q(id, '血圧 ≈ 心拍出量 × 体血管抵抗'), ['心拍出量', C.blue])],
      { x: X0, y: 2.0, w: BW, h: eqH, size: eqSize, bold: true, align: 'center' });
    s.text([mark(q(id, '心拍出量 = 1 回拍出量 × 心拍数'), ['心拍出量', C.blue])],
      { x: X0, y: r3(2.0 + eqH + 0.2), w: BW, h: eqH, size: eqSize, bold: true, align: 'center' });
    const eqBottom = 2.0 + 2 * eqH + 0.2;
    // 4 要素のチップ（ティール枠・横並び）
    const chips = ['前負荷', '後負荷', '収縮力', '心拍数'].map(t => q(id, t));
    const cw = 2.5, ch = 1.0, cg = 0.5;
    const cx0 = X0 + (BW - (4 * cw + 3 * cg)) / 2;
    const cy = 4.3;
    if (cy < eqBottom + GAP) throw new Error(`${id}: 式とチップの間隔 ${r3(cy - eqBottom)}in`);
    chips.forEach((t, i) => s.box({
      x: r3(cx0 + i * (cw + cg)), y: cy, w: cw, h: ch, fill: C.white, radius: 0.12,
      line: { color: C.teal, width: 3 }, text: t, size: 28, bold: true, color: C.dark,
    }));
    // 一文
    const sy = cy + ch + 0.4;
    s.text([mark(q(id, '昇圧薬は、この 4 つのどれかを動かす'), ['4 つのどれか', C.teal])],
      { x: X0, y: sy, w: BW, size: 28, bold: true, align: 'center' });
  }

  // ===== S3 受容体で整理する（表 4列・4行） =====
  {
    const id = 'S3';
    const s = deck.slide({ title: '受容体で整理する', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    const drug = t => ({ text: q(id, t), bold: true });
    const sym = t => {
      q(id, t);
      if (t === '◎') return { text: t, color: C.teal, bold: true, align: 'center' };
      if (t === '−') return { text: t, color: C.gray, align: 'center' };
      // 「◎（β2 も）」: 記号だけティール太字（他の ◎ とそろえる）
      if (t.startsWith('◎')) return { text: [[{ text: '◎', color: C.teal, bold: true }, t.slice(1)]], align: 'center' };
      return { text: t, align: 'center' };
    };
    const rows = [
      [q(id, '薬'), brk(id, 'α1：血管収縮', '血管'), brk(id, 'β1：心拍数・収縮力', '心拍数'), q(id, '主な効き方')],
      [drug('フェニレフリン'), sym('◎'), sym('−'), brk(id, '後負荷↑ 心拍数↓（反射） 前負荷↑は状態次第', '前負荷')],
      [drug('エフェドリン'), sym('○（間接）'), sym('○（間接＋直接）'), q(id, '心拍数↑ 収縮力↑ 血管収縮')],
      [drug('ノルアドレナリン'), sym('◎'), sym('○'), q(id, '後負荷↑ 前負荷↑ 収縮力↑')],
      [drug('アドレナリン'), sym('◎'), sym('◎（β2 も）'), q(id, '全部↑（→ 3-9）')],
    ].map(r => r.map(c => (Array.isArray(c) ? c.join('\n') : c)));
    const colW = [2.92, 1.9, 2.92, 4.49];
    const t = s.table(rows, {
      x: X0, y: Y0, colW, size: 22, margin: [4, 6, 4, 6],
      rowH: [1.06, 1.1, 0.74, 0.74, 0.74],
    });
    if (t.y + t.h > bottomOf(s) + 1e-9) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)}in`);
  }

  // ===== S4 フェニレフリン（3列カード） =====
  {
    const id = 'S4';
    const s = deck.slide({ title: 'フェニレフリン', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: q(id, '用量'), titleColor: C.blue,
        items: [
          markIn(brk(id, '添付文書：1 回 0.2 mg 静注（0.1〜0.5 mg）', '（0.1'), '1 回 0.2 mg', C.blue),
          [q(id, '反復は 10〜15 分おき')],
          markIn(markIn(brk(id, '麻酔中の慣用（添付文書外）：0.1 mg/mL に希釈し 0.05〜0.1 mg ずつ', '（添付', '0.1 mg/mL', '0.05'),
            '（添付文書外）', C.vermilion), '0.05〜0.1 mg', C.blue),
        ],
      },
      {
        title: q(id, '向く場面'), titleColor: C.teal,
        items: [
          [q(id, '頻脈ぎみの低血圧')],
          brk(id, '脊麻・麻酔薬による血管拡張', '血管'),
          brk(id, '心拍数を上げたくない患者', '上げ'),
        ],
      },
      {
        title: q(id, '注意'), titleColor: C.vermilion,
        items: [
          [[{ text: q(id, '反射性徐脈'), color: C.vermilion, bold: true }]],
          brk(id, '心機能が悪いと心拍出量が落ちる', '心拍出量'),
          brk(id, '少量ボーラスの効果は数分、繰り返す', '効果は'),
        ],
      },
    ]);
  }

  // ===== S5 エフェドリン（3列カード） =====
  {
    const id = 'S5';
    const s = deck.slide({ title: 'エフェドリン', chapter: 2, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: q(id, '用量'), titleColor: C.blue,
        items: [
          markIn(brk(id, '添付文書：麻酔時の血圧降下に 1 回 4〜8 mg 静注', '麻酔時', '1 回'), '4〜8 mg', C.blue),
          brk(id, '希釈（慣用）：40 mg/1 mL を 10 mL に（4 mg/mL）', '40 mg', '（4 mg'),
          [q(id, '1〜2 mL ずつ')],
        ],
      },
      {
        title: q(id, '向く場面'), titleColor: C.teal,
        items: [
          [q(id, '徐脈ぎみの低血圧')],
          brk(id, '麻酔導入後の最初の一手', '最初'),
          [q(id, '心拍数も上げたいとき')],
        ],
      },
      {
        title: q(id, '注意'), titleColor: C.vermilion,
        items: [
          brk(id, '繰り返すと効かなくなる（タキフィラキシー）', '効かなく', '（タキ'),
          brk(id, '効き始めは少し遅く、長く残る', '長く'),
          markIn(brk(id, 'カテコールアミン投与中は禁忌', '投与中'), '禁忌', C.red),
        ],
      },
    ]);
  }

  // ===== S6 ノルアドレナリン（3列カード） =====
  {
    const id = 'S6';
    const s = deck.slide({ title: 'ノルアドレナリン', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: q(id, '用量'), titleColor: C.blue,
        items: [
          brk(id, '添付文書：1 mg/250 mL を 0.5〜1.0 mL/分', '0.5〜'),
          markIn(markIn(brk(id, '慣用（添付文書外）：0.03〜0.1 μg/kg/分', '0.03'), '（添付文書外）', C.vermilion),
            '0.03〜0.1 μg/kg/分', C.blue),
          markIn(brk(id, '少量ボーラスは施設の方法で（添付文書外）', '施設', '（添付'), '（添付文書外）', C.vermilion),
        ],
      },
      {
        title: q(id, '向く場面'), titleColor: C.teal,
        items: [
          brk(id, '続く低血圧、血管拡張（敗血症など）', '血管', '（敗血'),
          brk(id, 'ボーラスで追いつかないとき', '追い'),
          brk(id, '収縮力を落としたくないとき', '落とし'),
        ],
      },
      {
        title: q(id, '注意'), titleColor: C.vermilion,
        items: [
          markIn(brk(id, '血管外漏出（末梢なら太い近位の静脈・短時間・観察）', '太い', '短時間'), '血管外漏出', C.red),
          brk(id, '過剰で臓器虚血、不整脈', '不整脈'),
          brk(id, '側管の押し込みで一気に入る', '一気'),
        ],
      },
    ]);
  }

  // ===== S7 使い分け（表 3列・5行） =====
  {
    const id = 'S7';
    const s = deck.slide({ title: '使い分け', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    const first = (str, ...breaks) => {
      // 第一選択: 先頭の薬名（または処置）を青太字、括弧書きは通常
      const lines = breaks.length ? brk(id, str, ...breaks) : [q(id, str)];
      return lines.map(l => {
        const k = l.indexOf('（');
        if (k === 0) return l;
        if (k < 0) return [{ text: l, color: C.blue, bold: true }];
        return [{ text: l.slice(0, k), color: C.blue, bold: true }, l.slice(k)];
      });
    };
    const rows = [
      [q(id, '状況'), q(id, '足りない要素'), q(id, '第一選択')],
      [q(id, '低血圧＋徐脈'), q(id, '心拍数・収縮力'), { text: first('エフェドリン（徐脈が強ければアトロピン）', '（徐脈') }],
      [q(id, '低血圧＋頻脈'), q(id, '後負荷（血管拡張）'), { text: first('フェニレフリン') }],
      [{ text: brk(id, '下がり続ける、ボーラスが効かない', 'ボーラス') }, q(id, '後負荷＋収縮力'), { text: first('ノルアドレナリン持続') }],
      [{ text: brk(id, '心機能が悪い、低心拍出', '低心拍出') }, q(id, '収縮力'), { text: first('ドブタミン、アドレナリン（→ 3-9）', '（→') }],
      [q(id, '出血・脱水'), q(id, '前負荷'), { text: first('輸液・輸血（昇圧薬は時間稼ぎ）') }],
    ];
    const t = s.table(rows, { x: X0, y: Y0, colW: [3.31, 3.31, 5.61], size: 22, margin: [2, 8, 2, 8] });
    if (t.y + t.h > bottomOf(s) + 1e-9) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)}in > ${bottomOf(s)}`);
    console.log(`  p${s.page}: table h=${r3(t.h)} → bottom ${r3(t.y + t.h)}`);
  }

  // ===== S8 昇圧薬の前に（箇条書き5行） =====
  {
    const id = 'S8';
    const s = deck.slide({ title: '昇圧薬の前に', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, '原因を探す：出血、深い麻酔、体位、圧迫、気胸、アナフィラキシー'), ['原因を探す：', C.blue]),
      mark(q(id, '前負荷を戻す：輸液、下肢挙上、圧迫の解除'), ['前負荷を戻す：', C.blue]),
      mark(q(id, '麻酔薬を減らす（吸入濃度、持続投与）'), ['麻酔薬を減らす', C.blue]),
      [{ text: q(id, '昇圧薬は原因治療までのつなぎ'), color: C.teal, bold: true }],
      mark(q(id, '目安：平均血圧 65 mmHg 以上を保つ'), ['目安：', C.blue], ['65 mmHg 以上', C.teal]),
    ], { y: 2.0, size: 24, gap: 32 });
  }

  // ===== S9 落とし穴（箇条書き5行） =====
  {
    const id = 'S9';
    const s = deck.slide({ title: '落とし穴', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, 'フェニレフリンで徐脈になり、心拍出量が落ちる'), ['心拍出量が落ちる', C.red]),
      mark(q(id, 'エフェドリンを連発して効かなくなる'), ['効かなくなる', C.red]),
      mark(q(id, 'ノルアドレナリンのラインを押し込んで急上昇'), ['急上昇', C.red]),
      mark(q(id, '希釈を間違えて 10 倍量（mg と μg）'), ['10 倍量', C.red]),
      mark(q(id, '昇圧薬で血圧だけ戻して、出血に気づかない'), ['出血に気づかない', C.red]),
    ], { y: 2.0, size: 28, gap: 26 });
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const id = 'S10';
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      q(id, '血圧 ≈ 心拍出量 × 体血管抵抗。4 要素のどれが足りないか'),
      mark(q(id, '徐脈ぎみ → エフェドリン 4〜8 mg'), ['エフェドリン 4〜8 mg', C.blue]),
      mark(q(id, '頻脈ぎみ → フェニレフリン 0.05〜0.1 mg ずつ（慣用）'), ['フェニレフリン 0.05〜0.1 mg ずつ', C.blue]),
      mark(q(id, '続く低血圧 → ノルアドレナリン持続'), ['ノルアドレナリン持続', C.blue]),
      mark(q(id, '昇圧薬はつなぎ。原因（出血・麻酔深度・前負荷）を探す'), ['昇圧薬はつなぎ。', C.teal]),
    ], { y: 2.0, size: 28, gap: 24 });
  }

  for (const [id, ls] of Object.entries(SOURCE_SPLITS)) console.log(`  ${id} 出典の改行位置: 「${ls[0]}」⏎「${ls[1]}」`);
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
