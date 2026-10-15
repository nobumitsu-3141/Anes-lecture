'use strict';
/**
 * 4-6 周術期の DNAR
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
const { Deck, C, LAYOUT, fits, heightFor } = require('../_lib/decklib');

const NAME = '4-6_perioperative_dnar';
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
  if (!s) throw new Error(`spec.md の ${id} に sources がありません`);
  return s.split(' / ');
}
/**
 * 2 行に分けて渡す出典（文献の途中の空白で改行する。文言は変えない）。
 * spec の sources 行を、before の直前の半角スペースで 2 行に割る。各行が 16pt・11.6in に収まること、
 * 2 行がライブラリで 1 行に詰められないことを確認する。
 * （項目単位では 3 行になるが、全体の長さは 2 行に収まる出典に使う）
 */
function sourcesBreak(id, before) {
  const S = LAYOUT.source;
  const want = SPEC[id].sources;
  const k = want.indexOf(' ' + before);
  if (k < 0) throw new Error(`${id} 出典: 改行位置「${before}」が見つかりません`);
  const lines = [want.slice(0, k), want.slice(k + 1)];
  if (lines.join(' ') !== want) throw new Error(`${id} 出典の行分けが元の文言と一致しません`);
  lines.forEach(l => { if (!fits(l, S.size, S.w)) throw new Error(`${id} 出典「${l}」が 1 行に収まりません`); });
  if (fits(lines.join(' / '), S.size, S.w)) throw new Error(`${id} 出典: 2 行がライブラリで 1 行に詰められてしまいます`);
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
/** 略記した出典のスライド用: ノートの末尾に出典の全文を足す（ノート本文は原文のまま） */
function notesWithFullSource(id) {
  return notesOf(id) + '\n出典（全文）: ' + SPEC[id].sources;
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
/** 項目（行配列）のうち、指定した行に部分装飾を付ける。marks: {行番号: [[語, 色], ...]} */
function markLines(lines, marks) {
  return lines.map((l, i) => (marks[i] ? mark(l, ...marks[i]) : l));
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const Y0 = LAYOUT.body.y;        // 1.85
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const r3 = v => Math.round(v * 1000) / 1000;
/** 出典 2 行（枠 y6.5）のときの本文の下端の上限 */
const BOTTOM_2SRC = 6.45;
/** 強調ボックス（白地＋ティールの枠、黒の太字。中央寄せ） */
const EMPH = { radius: 0.1, lineW: 3.5 };

/** 強調ボックス。lines は行の配列（各行は文字列 or run 配列） */
function emphBox(s, o) {
  s.box({
    x: r3(o.x != null ? o.x : X0), y: r3(o.y), w: r3(o.w != null ? o.w : BW), h: o.h,
    fill: C.white, radius: EMPH.radius, line: { color: C.teal, width: EMPH.lineW },
    text: o.lines, size: o.size, bold: true, color: C.dark,
  });
}

/**
 * カード: 角丸の塗り箱 + 見出し + 項目（各項目は行配列、項目間に itemGap）。文字はすべて箱の内側。
 * c: {x,y,w,h, title:[行], titleColor, titleSize, items:[[行,...],...], pad, titleGap, itemGap}
 */
function itemCard(s, c) {
  const pad = c.pad, ts = c.titleSize, bs = 22;
  s.box({ x: r3(c.x), y: r3(c.y), w: r3(c.w), h: r3(c.h), fill: C.bg, radius: 0.08 });
  const iw = r3(c.w - 2 * pad), tx = r3(c.x + pad);
  const th = heightFor(c.title.length, ts, 0);
  s.text(c.title, { x: tx, y: r3(c.y + pad), w: iw, h: th, size: ts, bold: true, color: c.titleColor });
  let y = c.y + pad + th + c.titleGap;
  c.items.forEach((lines, i) => {
    const h = heightFor(lines.length, bs, 0);
    s.text(lines, { x: tx, y: r3(y), w: iw, h, size: bs });
    y += h + (i < c.items.length - 1 ? c.itemGap : 0);
  });
  if (y > c.y + c.h - pad + 1e-6) throw new Error(`p${s.page} カード「${runText(c.title[0])}」: 高さ不足（本文の下端 ${r3(y)}in > ${r3(c.y + c.h - pad)}in）`);
  return y;
}
/** カードに必要な高さ */
function cardNeedH(c) {
  const th = heightFor(c.title.length, c.titleSize, 0);
  const body = c.items.reduce((a, l) => a + heightFor(l.length, 22, 0), 0) + c.itemGap * (c.items.length - 1);
  return 2 * c.pad + th + c.titleGap + body;
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['何が\n問題か', 'ASAの\n3択', '国内の\n勧告', '術前の\n手順', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-6 周術期の DNAR',
  });

  // ===== S1 表紙 =====
  {
    const subtitle = '4-6 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    if (!SPEC_MD.includes('title=「周術期の DNAR」')) throw new Error('S1: 表紙タイトルが spec と違います');
    deck.cover({ title: '周術期の DNAR', subtitle, meta, notes: notesOf('S1') });
  }

  // 箇条書き＋強調ボックスのスライド（S2・S5）: 段落後 18pt・y1.95。強調ボックスは 28pt・h0.9
  const BUL_Y = 1.95, BUL_GAP = 18;

  // ===== S2 DNAR の定義を確かめる（箇条書き4行 + 強調ボックス） =====
  {
    const id = 'S2';
    // 出典: 3 件を項目単位で詰めると 3 行になるため、Hiraoka の著者名と誌名の間で 2 行に分ける（文言は変えない）
    const s = deck.slide({ title: 'DNAR の定義を確かめる', chapter: 0, notes: notesOf(id), sources: sourcesBreak(id, 'Int J Gen Med') });
    const size = 26;
    s.bullets([
      mark(q(id, 'DNAR＝心停止のときに心肺蘇生を始めない、という指示'), ['心停止のとき', C.blue]),
      mark(q(id, 'それ以外の治療（昇圧薬、人工呼吸、透析、手術）は別の話'), ['別の話', C.vermilion]),
      mark(q(id, '『急変時』というあいまいな言葉に置き換えない'), ['置き換えない', C.vermilion]),
      mark(q(id, 'DNAR と終末期医療は同じものではない'), ['同じものではない', C.vermilion]),
    ], { y: BUL_Y, size, gap: BUL_GAP });
    const bulBottom = BUL_Y + heightFor(4, size, BUL_GAP);
    const boxH = 0.9, by = 5.3;
    if (by < bulBottom + GAP) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - bulBottom)}in`);
    if (by + boxH > BOTTOM_2SRC) throw new Error(`${id}: 強調ボックスが出典に掛かります`);
    emphBox(s, { x: X0 + 1.2, w: BW - 2.4, y: by, h: boxH, size: 28,
      lines: [mark(q(id, 'DNAR＝何もしない、ではない'), ['ではない', C.vermilion])] });
  }

  // ===== S3 手術室で何が問題か（2列カード + 強調ボックス） =====
  {
    const id = 'S3';
    // 出典: 3 件を項目単位で詰めると 3 行になるため、Ramachandran の著者名と誌名の間で 2 行に分ける
    const s = deck.slide({ title: '手術室で何が問題か', chapter: 0, notes: notesOf(id), sources: sourcesBreak(id, 'Anesthesiology 2013') });
    const common = { pad: 0.12, titleSize: 24, titleGap: 0.06, itemGap: 0.1 };
    const cards = [
      {
        ...common, title: [q(id, '麻酔は『蘇生』を含む')], titleColor: C.blue,
        items: [
          brk(id, '気管挿管、人工呼吸、昇圧薬は麻酔の一部', '麻酔の一部'),
          markLines(brk(id, 'CPR に固有なのは胸骨圧迫と除細動だけ', '胸骨圧迫'), { 1: [['胸骨圧迫と除細動だけ', C.blue]] }),
          brk(id, '導入後の低血圧への対応は蘇生ではない', '蘇生ではない'),
        ],
      },
      {
        ...common, title: [q(id, '手術室の心停止は違う')], titleColor: C.vermilion,
        items: [
          brk(id, '医原性・可逆的な原因もある（原因で見込みが違う）', '（原因で'),
          markLines(brk(id, '周術期心停止の退院生存 32%（院内心停止全体は 15〜17%）', '（院内'), { 0: [['32%', C.blue]] }),
          markLines(brk(id, '麻酔が原因なら約 9 割が救命（1967〜84 年）', '（1967'), { 0: [['約 9 割', C.blue]] }),
        ],
      },
    ];
    const cy = 1.82;
    const ch = Math.ceil(Math.max(...cards.map(cardNeedH)) * 100) / 100;
    const cw = (BW - GAP) / 2;
    cards.forEach((c, i) => itemCard(s, { ...c, x: X0 + i * (cw + GAP), y: cy, w: cw, h: ch }));
    const boxH = 0.6, by = cy + ch + GAP;
    if (by + boxH > BOTTOM_2SRC) throw new Error(`${id}: 強調ボックスの下端 ${r3(by + boxH)}in が出典に掛かります`);
    emphBox(s, { y: by, h: boxH, size: 24,
      lines: [mark(q(id, '黙って停止するのも、確認せず適用するのも誤り。術前に話し合う'), ['誤り', C.red])] });
    console.log(`  p${s.page} カード h=${ch}、強調ボックス y=${r3(by)}〜${r3(by + boxH)}`);
  }

  // ===== S4 ASA の 3 つの選択肢（3列カード + 下に1行） =====
  {
    const id = 'S4';
    // 出典: ASA の文献名（（2023 再承認）まで）が 16pt・1 行に入らないため、文献名の途中の空白（Patients / with）で 2 行に分ける
    const src = sourcesBreak(id, 'with DNR Orders');
    const s = deck.slide({ title: 'ASA の 3 つの選択肢', chapter: 1, notes: notesOf(id), sources: src });
    const common = { pad: 0.1, titleSize: 24, titleGap: 0.1, itemGap: 0.24, titleColor: C.blue };
    const cards = [
      {
        ...common, title: [q(id, 'A：全面的に蘇生する')],
        items: [
          brk(id, '手術中と術直後は DNAR を停止', 'DNAR'),
          brk(id, '適切な蘇生処置はすべて行う', 'すべて'),
        ],
      },
      {
        ...common, title: [q(id, 'B：処置を限定する')],
        items: [
          brk(id, '拒否する処置を指定（例：胸骨圧迫、除細動、気管挿管）', '（例', '除細動'),
          markLines(brk(id, '麻酔に不可欠な処置は拒否できない', '処置は'), { 1: [['拒否できない', C.vermilion]] }),
        ],
      },
      {
        ...common, title: [q(id, 'C：目標・価値観で決める')],
        items: [
          [q(id, 'すぐ戻せる事象には蘇生')],
          brk(id, '永続的な後遺症が残りそうなら蘇生しない', '残りそう'),
          [q(id, '判断は医療者に委ねる')],
        ],
      },
    ];
    // 幅: 各カードの最長行（推定幅）に比例して配る
    const { measure } = require('../_lib/decklib');
    const need = cards.map(c => Math.max(...c.title.map(t => measure(runText(t), c.titleSize)),
      ...c.items.flat().map(l => measure(runText(l), 22))));
    const availInner = BW - 2 * GAP - cards.reduce((a, c) => a + 2 * c.pad, 0);
    const slack = availInner - need.reduce((a, b) => a + b, 0);
    if (slack < 0) throw new Error(`${id}: カード幅が ${r3(-slack)}in 足りません`);
    const ch = Math.ceil(Math.max(...cards.map(cardNeedH)) * 100) / 100 + 0.1;
    let x = X0;
    cards.forEach((c, i) => {
      const w = Math.floor((need[i] + slack / 3 + 2 * c.pad) * 1000) / 1000;
      itemCard(s, { ...c, x, y: Y0, w, h: ch });
      x += w + GAP;
    });
    const ly = r3(Y0 + ch + 0.36);
    s.text([mark(q(id, 'どれを選んでも『いつ元に戻すか』を一緒に決める（→ S5）'), ['『いつ元に戻すか』', C.vermilion])],
      { x: X0, y: ly, w: BW, size: 26, bold: true });
    console.log(`  p${s.page} カード幅 ${cards.map((c, i) => r3(need[i] + slack / 3 + 2 * c.pad)).join(' / ')}、h=${ch}、下の1行 y=${ly}`);
  }

  // ===== S5 術後、いつ元に戻すか（箇条書き4行 + 強調ボックス） =====
  // 1 行目が 26pt では行頭記号ぶんで 1 行に入らないため 24pt（1 バレット 1 行を優先）
  {
    const id = 'S5';
    const s = deck.slide({ title: '術後、いつ元に戻すか', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    const size = 24;
    s.bullets([
      mark(q(id, 'ASA：PACU 退室時か、麻酔・手術の急性の影響から回復した時点'), ['PACU 退室時', C.blue]),
      mark(q(id, '手術患者の心停止の 86% は術後に起きる（半数は 5 日以内）'), ['86%', C.blue]),
      q(id, '『術後 ◯ 日までは治療、改善なければ戻す』という区切り方も'),
      mark(q(id, '決めていないと、術後のコードが誰にもわからなくなる'), ['誰にもわからなくなる', C.vermilion]),
    ], { y: BUL_Y, size, gap: BUL_GAP });
    const bulBottom = BUL_Y + heightFor(4, size, BUL_GAP);
    const boxH = 0.9, by = 5.3;
    if (by < bulBottom + GAP) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - bulBottom)}in`);
    if (by + boxH > BOTTOM_2SRC) throw new Error(`${id}: 強調ボックスが出典に掛かります`);
    emphBox(s, { x: X0 + 1.2, w: BW - 2.4, y: by, h: boxH, size: 28,
      lines: [mark(q(id, '停止の有効期間を、術前に決めて記録する'), ['術前に決めて記録する', C.vermilion])] });
  }

  // ===== S6 国内の勧告と制度（2列カード + 強調ボックス） =====
  // 見出しは 24pt では 2 列に 1 行で入らず、2 行にすると強調ボックスが出典に掛かるため、見出しを 22pt 太字・1 行にする
  {
    const id = 'S6';
    const s = deck.slide({ title: '国内の勧告と制度', chapter: 2, notes: notesOf(id), sources: sourcesOf(id) });
    const common = { pad: 0.1, titleSize: 22, titleGap: 0.08 };
    const cards = [
      {
        ...common, title: [q(id, '日本集中治療医学会の勧告（2016）')], titleColor: C.blue, itemGap: 0.08,
        items: [
          markLines(brk(id, '心停止時にのみ有効。ほかの治療は別に議論', 'ほかの'), { 0: [['心停止時にのみ有効', C.blue]] }),
          brk(id, '妥当性は繰り返し話し合って評価する', '話し合って'),
          [mark(q(id, 'Partial DNAR 指示は行わない'), ['行わない', C.vermilion])],
          [q(id, '日本版 POLST には準拠しない')],
        ],
      },
      {
        ...common, title: [q(id, '厚生労働省ガイドライン（2018 改訂）')], titleColor: C.vermilion, itemGap: 0.08,
        items: [
          brk(id, '本人の意思 → 推定意思 → 本人にとっての最善', '→ 本人に'),
          brk(id, '繰り返し話し合い、文書に残す（ACP）', '文書に'),
          brk(id, '決定が困難・合意できなければ複数の専門家の話し合いの場', '複数の'),
        ],
      },
    ];
    const { measure } = require('../_lib/decklib');
    const need = cards.map(c => Math.max(...c.title.map(t => measure(runText(t), c.titleSize)),
      ...c.items.flat().map(l => measure(runText(l), 22))));
    const availInner = BW - GAP - cards.reduce((a, c) => a + 2 * c.pad, 0);
    const slack = availInner - need.reduce((a, b) => a + b, 0);
    if (slack < 0) throw new Error(`${id}: カード幅が ${r3(-slack)}in 足りません`);
    const ch = Math.ceil(Math.max(...cards.map(cardNeedH)) * 100) / 100;
    // 項目の少ないカードは、余った高さを項目間に配る
    const cards2 = cards.map(c => {
      const free = ch - cardNeedH(c);
      return free > 0 && c.items.length > 1 ? { ...c, itemGap: c.itemGap + Math.floor(free / (c.items.length - 1) * 100) / 100 } : c;
    });
    let x = X0;
    cards2.forEach((c, i) => {
      const w = Math.floor((need[i] + slack / 2 + 2 * c.pad) * 1000) / 1000;
      itemCard(s, { ...c, x, y: Y0, w, h: ch });
      x += w + GAP;
    });
    const boxH = 0.6, by = Y0 + ch + GAP;
    if (by + boxH > BOTTOM_2SRC) throw new Error(`${id}: 強調ボックスの下端 ${r3(by + boxH)}in が出典に掛かります`);
    emphBox(s, { y: by, h: boxH, size: 24,
      lines: [mark(q(id, '他院・前回入院の DNAR は、経緯と本人の意思を確かめ直す'), ['確かめ直す', C.vermilion])] });
    console.log(`  p${s.page} カード h=${ch}、slack=${r3(slack)}、強調ボックス y=${r3(by)}〜${r3(by + boxH)}`);
  }

  // ===== S7 術前に行う 6 つのこと（番号つき箇条書き6行） =====
  // 番号は文言に含まれるので行頭記号は付けない（番号をブルーの太字に）。最長行が 26pt では 1 行に入らないため 24pt
  {
    const id = 'S7';
    // 出典: 3 件を項目単位で詰めると 3 行になるため、Nolan の著者名と誌名の間で 2 行に分ける
    const s = deck.slide({ title: '術前に行う 6 つのこと', chapter: 3, notes: notesOf(id), sources: sourcesBreak(id, 'Anaesthesia 2024') });
    const lines = [
      '1. DNAR の有無と経緯を確かめる（誰が、いつ、本人は参加したか）',
      '2. 手術の目的と、麻酔に含まれる処置（挿管・昇圧薬）を説明する',
      '3. 選択肢（A・B・C）を示し、本人または代理人と決める',
      '4. 元に戻す時点（PACU 退室、術後 ◯ 日）を決める',
      '5. 記録と指示変更。タイムアウトで外科医・看護師・病棟に伝える',
      '6. 外科医と方針を前もってそろえる（緊急手術になる前に）',
    ].map((l, i) => mark(q(id, l), [`${i + 1}.`, C.blue]));
    const size = 24, gap = 20, y = 1.95;
    const h = heightFor(6, size, gap);
    if (y + h > BOTTOM_2SRC) throw new Error(`${id}: 下端 ${r3(y + h)}in が出典に掛かります`);
    s.text(lines, { x: X0, y, w: BW, h, size, gap });
  }

  // ===== S8 記録の例（大きなボックスに記録例4行 + 下に1行） =====
  // 2 行目が 24pt ではボックスの内側に 1 行で入らないため、記録例は 22pt（4 行のまま）
  {
    const id = 'S8';
    const s = deck.slide({ title: '記録の例', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    const pad = 0.25, size = 22, gap = 16;
    const rec = [
      mark(q(id, '術前 DNAR あり（3 月 2 日、病棟で家族と決定。本人は不参加）'), ['術前 DNAR あり', C.blue]),
      mark(q(id, '本人と再確認：手術中と PACU 退室までは蘇生を全面的に行う（選択肢 A）'), ['本人と再確認：', C.blue]),
      mark(q(id, 'PACU 退室後は元の DNAR に戻す。術後の合併症の治療は通常どおり行う'), ['PACU 退室後', C.blue]),
      mark(q(id, '説明者：麻酔科 ◯◯、同席：本人、長女、病棟看護師 ◯◯'), ['説明者：', C.blue], ['同席：', C.blue]),
    ];
    const th = heightFor(4, size, gap);
    const by = 1.95, bh = r3(th + 2 * pad + 0.1);
    s.box({ x: X0, y: by, w: BW, h: bh, fill: C.bg, radius: 0.08, line: { color: C.border, width: 1.5 } });
    s.text(rec, { x: X0 + pad, y: r3(by + pad + 0.05), w: r3(BW - 2 * pad), h: th, size, gap });
    const ly = r3(by + bh + 0.42);
    s.text([mark(q(id, '『急変時 DNAR』『様子を見る』のようなあいまいな記録はしない'), ['あいまいな記録はしない', C.red])],
      { x: X0, y: ly, w: BW, size: 26, bold: true });
    console.log(`  p${s.page} 記録ボックス y=${by}〜${r3(by + bh)}、下の1行 y=${ly}`);
  }

  // ===== S9 まとめ（箇条書き5行） =====
  // 出典: 3 件（ASA・日本集中治療医学会・厚生労働省）は 16pt で 3 行になる。区切りの空白を詰め（）の補足を外しても
  //       2 行に入らないため、ASA の文献名を略記する（文献は落とさない）。全文はノートの末尾へ。
  {
    const id = 'S9';
    const short = abbreviate(id, [['ASA Statement on Ethical Guidelines for the Anesthesia Care of Patients with DNR Orders', 'ASA Statement on DNR Orders']]);
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: notesWithFullSource(id), sources: short.split(' / ') });
    s.bullets([
      mark(q(id, 'DNAR＝心停止時に CPR を始めない。ほかの治療は別'), ['CPR を始めない', C.blue]),
      mark(q(id, '麻酔は挿管・人工呼吸・昇圧薬を含む。蘇生ではない'), ['蘇生ではない', C.blue]),
      mark(q(id, '黙って停止も、確認せず適用も誤り。術前に再検討する'), ['術前に再検討する', C.blue]),
      mark(q(id, '選択肢 A・B・C と『いつ戻すか』を決めて記録する'), ['『いつ戻すか』', C.blue]),
      mark(q(id, '国内の勧告：Partial DNAR は不可、繰り返し話し合う'), ['Partial DNAR は不可', C.blue]),
    ], { y: 2.0, size: 26, gap: 22 });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
