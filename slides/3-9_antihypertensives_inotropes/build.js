'use strict';
/**
 * 3-9 降圧薬・強心薬 — 組版スクリプト
 *
 *   cd slides/3-9_antihypertensives_inotropes
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *     書き写しの検査: 各スライドに置いた文字列を順につなげたものが、spec の on-slide の「」の中身を
 *     順につなげたものと一字一句同じでなければ Error（checkOnSlide）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, LAYOUT, measure, heightFor } = require('../_lib/decklib');

const NAME = '3-9_antihypertensives_inotropes';
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
const textOf = line => (Array.isArray(line) ? line.map(textOf).join('')
  : (line && typeof line === 'object') ? textOf(line.text) : String(line));
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
/** 行配列のうち、語を含む行にだけ mark を掛ける。marks: [[語, 色], ...] */
function marks(lines, ...ms) {
  let out = lines.slice();
  for (const [w, color] of ms) {
    let n = 0;
    out = out.map(l => {
      if (textOf(l).includes(w)) {
        n++;
        if (typeof l === 'string') return mark(l, [w, color]);
        // すでに run 配列: 文字列の run の中で探す
        const next = [];
        let found = 0;
        for (const r of l) {
          if (typeof r !== 'string' || !r.includes(w) || found) { next.push(r); continue; }
          found++;
          next.push(...mark(r, [w, color]));
        }
        if (!found) throw new Error(`marks: 「${textOf(l)}」の「${w}」が装飾済みの run の中にあります`);
        return next;
      }
      return l;
    });
    if (n !== 1) throw new Error(`marks: 「${w}」が ${n} 行に`);
  }
  return out;
}

// ---------------------------------------------------------------------------
// 共通の寸法
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;          // 0.55
const BW = LAYOUT.body.w;          // 12.23
const TOP = 1.85;                  // 本文の上端
const GAP = 0.24;                  // 要素間の最小間隔（0.6cm）
const r3 = v => Math.round(v * 1000) / 1000;
/**
 * 本文の下端。出典1行（枠 y6.85）: 枠から 0.24in 空けて 6.61。
 * 出典2行（枠 y6.5・h0.85・下揃え）: 2行の文字は枠の下 0.67in（16pt×1.5×2）に収まり上端は約 6.68in なので、
 * 本文は 6.42in まで（文字との間隔 0.26in。3-8 と同じ値）。
 */
const BOTTOM_1SRC = LAYOUT.source.y - GAP;   // 6.61
const BOTTOM_2SRC = 6.42;
const bottomOf = s => (!s.hasSources ? LAYOUT.body.bottom : (s.sourceTop < LAYOUT.source.y - 1e-6 ? BOTTOM_2SRC : BOTTOM_1SRC));
const DOSE = C.blue;               // 用量の数値（太字ブルー）
const NOTE = C.vermilion;          // （目安）など書き分けの注記
const DANGER = C.red;              // 禁忌・危険

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す） */
function contentSlide(deck, id, chapter) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id) });
  if (s.hasSources && s.sourceTop < LAYOUT.source.y2 - 1e-6) throw new Error(`${id}: 出典が3行以上`);
  return s;
}

// ---------------------------------------------------------------------------
// 見出し＋項目のカード（各項目は1〜数行。項目の間は細い区切り線）。文字はすべて箱の内側。
// 見出しは1行または行配列（2行まで）。
// ---------------------------------------------------------------------------
// lineGap: 項目内の段落間（pt）。0 だと lint が 6pt と仮定して高さを多めに見積もるので 1pt を明示する
const CARD = { pad: 0.14, titleSize: 24, size: 22, titleGap: 0.1, lineGap: 1, minGap: 0.2, maxGap: 0.42 };
const titleLines = c => (c.titleMulti ? c.title : [c.title]);   // titleMulti: title は行配列
function cardMinW(c) {
  let w = 0;
  titleLines(c).forEach(l => { w = Math.max(w, measure(textOf(l), CARD.titleSize)); });
  c.items.forEach(lines => lines.forEach(l => { w = Math.max(w, measure(textOf(l), CARD.size)); }));
  return w + 2 * CARD.pad + 0.01;
}
function cardItemsH(c) {
  return c.items.reduce((a, lines) => a + heightFor(lines.length, CARD.size, CARD.lineGap), 0);
}
function cardHeadH(c) {
  return CARD.pad + heightFor(titleLines(c).length, CARD.titleSize, 0) + CARD.titleGap;
}
function itemCard(s, id, c, x, y, w, h, gap) {
  const pad = CARD.pad;
  s.box({ x, y, w, h, fill: C.bg, radius: 0.08 });
  const iw = r3(w - 2 * pad);
  const tl = titleLines(c);
  const th = heightFor(tl.length, CARD.titleSize, 0);
  let cy = y + pad;
  s.text(tl, { x: r3(x + pad), y: r3(cy), w: iw, h: th, size: CARD.titleSize, bold: true, color: c.titleColor });
  tl.forEach(l => rec(id, l));
  cy += th + CARD.titleGap;
  c.items.forEach((lines, i) => {
    const ih = heightFor(lines.length, CARD.size, CARD.lineGap);
    s.text(lines, { x: r3(x + pad), y: r3(cy), w: iw, h: ih, size: CARD.size, gap: CARD.lineGap, color: C.dark });
    lines.forEach(l => rec(id, l));
    cy += ih;
    if (i < c.items.length - 1) {
      s.hline({ x: r3(x + pad), y: r3(cy + gap / 2), w: iw, color: C.border, width: 1 });
      cy += gap;
    }
  });
  if (cy > y + h - pad + 1e-6) throw new Error(`${id} カード「${textOf(tl[0])}」: 高さ不足（必要 ${r3(cy + pad - y)}in > ${h}in）`);
}
/**
 * 3列カード。列幅は中身の最長行から決め、余りは狭い列から順に配る（水位合わせ）。
 * 高さは本文の下端まで。項目の間隔はカードごとに、余白を項目間に配って maxGap まで（minGap 未満は Error）。
 */
function threeCards(s, id, cards) {
  const total = BW - GAP * (cards.length - 1);
  const need = cards.map(cardMinW);
  if (need.reduce((a, v) => a + v, 0) > total + 1e-9) {
    throw new Error(`${id}: 3列カードの必要幅 ${need.map(r3).join(' + ')} > ${r3(total)}in`);
  }
  const fixed = new Set();
  let W = 0;
  for (;;) {
    const free = need.map((v, i) => i).filter(i => !fixed.has(i));
    const rest = total - [...fixed].reduce((a, i) => a + need[i], 0);
    W = rest / free.length;
    const over = free.filter(i => need[i] > W + 1e-9);
    if (!over.length) break;
    over.forEach(i => fixed.add(i));
  }
  const widths = need.map((v, i) => r3(fixed.has(i) ? v : W));
  const bottom = bottomOf(s);
  const h = r3(bottom - TOP);
  const gaps = cards.map(c => {
    const n = c.items.length;
    if (n < 2) return 0;
    const avail = h - cardHeadH(c) - CARD.pad - cardItemsH(c);
    const g = Math.min(CARD.maxGap, avail / (n - 1));
    if (g < CARD.minGap - 1e-6) throw new Error(`${id} カード「${textOf(titleLines(c)[0])}」: 高さが足りません（項目間隔 ${r3(g)}in < ${CARD.minGap}in）`);
    return Math.floor(g * 100) / 100;
  });
  let x = X0;
  cards.forEach((c, i) => {
    itemCard(s, id, c, r3(x), TOP, widths[i], h, gaps[i]);
    x += widths[i] + GAP;
  });
  console.log(`  ${id}: cards w=${widths.join('/')} need=${need.map(r3).join('/')} h=${h} gap=${gaps.join('/')}`);
}

/** 箇条書き（1項目1行）。下端が本文の下端を越えたら Error */
function bulletList(s, id, items, o) {
  items.forEach(t => rec(id, t));
  s.bullets(items, { y: o.y, size: o.size, gap: o.gap });
  const bh = heightFor(items.length, o.size, o.gap);
  if (o.y + bh > bottomOf(s) + 1e-6) throw new Error(`${id}: 箇条書きが出典に近すぎます（下端 ${r3(o.y + bh)}in > ${r3(bottomOf(s))}in）`);
  console.log(`  ${id}: bullets ${o.size}pt gap ${o.gap} → bottom ${r3(o.y + bh)} (limit ${r3(bottomOf(s))})`);
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['整理', '降圧薬', '心拍数', '強心薬', 'アドレ\nナリン'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '3-9 降圧薬・強心薬',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 4 要素で整理する（表 3列 × 5行） =====
  {
    const id = 'S2';
    const s = contentSlide(deck, id, 0);
    const head = ['動かす要素', '方向', '薬'];
    const dir = (t, color) => ({ text: t, color, bold: true });
    // [要素, 方向（行配列）, 薬名, 括弧]
    const body = [
      ['後負荷', [dir('下げる', C.blue)], 'ニカルジピン', '（動脈を広げる）'],
      ['前負荷', [dir('下げる', C.blue)], 'ニトログリセリン', '（静脈を広げる）'],
      ['心拍数', [dir('下げる', C.blue)], 'ランジオロール', '（β1 遮断）'],
      ['収縮力', [dir('上げる', C.vermilion)], 'ドブタミン', '（β1 刺激）'],
      ['全部', [dir('上げる', C.vermilion), '（用量で変わる）'], 'アドレナリン', '（α＋β）'],
    ];
    head.forEach(t => rec(id, t));
    const rows = [head];
    body.forEach(([el, dirLines, drug, paren]) => {
      rec(id, el);
      dirLines.forEach(l => rec(id, l));
      rec(id, drug + paren);
      rows.push([
        { text: el, bold: true, align: 'center' },
        { text: dirLines.map(l => (typeof l === 'string' ? [{ text: l }] : [l])), align: 'center' },
        { text: [[{ text: drug, bold: true }, paren]] },
      ]);
    });
    const size = 24;
    const colW = [2.5, 3.4, 6.33];
    const rowH = [0.68, 0.68, 0.68, 0.68, 0.68, 1.18];
    const t = s.table(rows, { x: X0, y: TOP, colW, size, rowH, margin: [4, 10, 4, 12] });
    if (t.y + t.h > bottomOf(s) + 1e-6) throw new Error(`${id}: 表の下端 ${r3(t.y + t.h)}in > ${r3(bottomOf(s))}in`);
    console.log(`  ${id}: table bottom ${r3(t.y + t.h)} (limit ${r3(bottomOf(s))})`);
  }

  // ===== S3 術中高血圧の前に（箇条書き5行） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 1);
    bulletList(s, id, [
      '麻酔が浅い、痛み刺激（切開、牽引、止血帯）',
      '高二酸化炭素血症、低酸素',
      '膀胱の充満、頭蓋内圧の上昇',
      '薬：昇圧薬の入れすぎ、ライン内の薬の押し込み',
      { text: 'まず原因。降圧薬はそのあと', bold: true },
    ], { y: 1.98, size: 28, gap: 26 });
  }

  // ===== S4 ニカルジピン（3列カード） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 1);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          marks(brk('持続：2〜10 μg/kg/分で開始', 'で開始'), ['2〜10 μg/kg/分', DOSE]),
          marks(brk('急速降圧：10〜30 μg/kg を静注', '10〜30'), ['10〜30 μg/kg', DOSE]),
          ['高齢者は低用量から'],
        ],
      },
      {
        title: '作用', titleColor: C.blue,
        items: [
          brk('Ca 拮抗薬、動脈を広げる（後負荷↓）', '（後負荷'),
          ['反射性に心拍数が上がる'],
          ['冠動脈も広がる'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['効果がやや長く残る'],
          ['末梢静脈で血管痛'],
          brk('下げすぎたら輸液と昇圧薬', '輸液'),
        ],
      },
    ]);
  }

  // ===== S5 ニトログリセリン（3列カード） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 1);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          marks(brk('異常高血圧：0.5〜5 μg/kg/分で開始', '0.5〜5'), ['0.5〜5 μg/kg/分', DOSE]),
          marks(brk('低血圧維持：1〜5 μg/kg/分で開始', '1〜5'), ['1〜5 μg/kg/分', DOSE]),
          ['心不全・狭心症は別の用量'],
        ],
      },
      {
        title: '作用', titleColor: C.blue,
        items: [
          brk('NO で静脈を広げる（前負荷↓）', '（前負荷'),
          ['高用量で動脈にも'],
          ['冠動脈を広げる'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['頭痛、反射性頻脈、耐性'],
          brk('塩化ビニル製の回路に吸着', '回路'),
          marks(brk('PDE5 阻害薬・リオシグアトとの併用は禁忌', 'リオシグアト', '併用'), ['禁忌', DANGER]),
        ],
      },
    ]);
  }

  // ===== S6 ランジオロール（3列カード） =====
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 2);
    threeCards(s, id, [
      {
        title: brk('用量（添付文書・手術時）', '（添付'), titleMulti: true, titleColor: C.blue,
        items: [
          marks(brk('0.125 mg/kg/分を 1 分間', '1 分間'), ['0.125 mg/kg/分', DOSE], ['1 分間', DOSE]),
          marks(['その後 0.04 mg/kg/分'], ['0.04 mg/kg/分', DOSE]),
          marks(brk('0.01〜0.04 mg/kg/分で調節', 'で調節'), ['0.01〜0.04 mg/kg/分', DOSE]),
        ],
      },
      {
        title: '作用', titleColor: C.blue,
        items: [
          brk('超短時間作用型 β1 遮断薬', 'β1'),
          brk('心拍数を下げる、血圧はあまり下げない', '血圧', 'あまり'),
          ['止めれば数分で戻る'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          marks(brk('禁忌：心原性ショック、II 度以上の房室ブロック、洞不全', 'II 度', '房室'), ['禁忌：', DANGER]),
          marks(brk('心機能低下例は低用量で（1 μg/kg/分から）', '（1'), ['1 μg/kg/分', DOSE]),
          brk('喘息では気管支攣縮、術後・ICU の用量は別', '術後'),
        ],
      },
    ]);
  }

  // ===== S7 ドブタミン（3列カード） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 3);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          marks(['通常 1〜5 μg/kg/分'], ['1〜5 μg/kg/分', DOSE]),
          marks(brk('必要なら 20 μg/kg/分まで', '20'), ['20 μg/kg/分', DOSE]),
          brk('アンプルは 5% ブドウ糖か生食で希釈', '5%', '生食'),
        ],
      },
      {
        title: '作用', titleColor: C.blue,
        items: [
          brk('β1 刺激で収縮力↑、心拍数↑', '心拍数'),
          ['β2 で血管はやや広がる'],
          brk('血圧は上がるとは限らない', '限らない'),
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['頻脈、不整脈'],
          brk('心筋の酸素消費が増える', '増える'),
          brk('前負荷が足りないと効かない', '効かない'),
        ],
      },
    ]);
  }

  // ===== S8 アドレナリン（3列カード） =====
  {
    const id = 'S8';
    const s = contentSlide(deck, id, 4);
    threeCards(s, id, [
      {
        title: '場面と用量（ガイドライン）', titleColor: C.blue,
        items: [
          marks(['心停止：1 mg 静注、3〜5 分ごと'], ['1 mg', DOSE], ['3〜5 分', DOSE]),
          brk('アナフィラキシー（麻酔中）：少量ずつ静注、静脈路なしは筋注', '少量'),
          marks(brk('低心拍出：持続 0.01〜0.1 μg/kg/分程度（目安）', '0.01', '（目安）'), ['0.01〜0.1 μg/kg/分', DOSE], ['（目安）', NOTE]),
        ],
      },
      {
        title: '作用', titleColor: C.blue,
        items: [
          brk('α と β のすべてに効く', 'すべて'),
          brk('低用量で β、高用量で α が前に出る', '高用量', '前に'),
          ['気管支も広げる'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          brk('不整脈、心筋虚血、高血糖、乳酸上昇', '高血糖'),
          marks(brk('濃度の取り違え（1 mg/mL）', '（1'), ['濃度の取り違え', DANGER]),
          brk('静注は希釈して少量ずつ', '少量'),
        ],
      },
    ]);
  }

  // ===== S9 落とし穴（箇条書き5行） =====
  {
    const id = 'S9';
    const s = contentSlide(deck, id, 4);
    bulletList(s, id, [
      mark('ニカルジピンで下げすぎ、しばらく戻らない', ['ニカルジピン', null]),
      mark('ニトログリセリンと PDE5 阻害薬の併用', ['ニトログリセリン', null], ['PDE5 阻害薬', null]),
      mark('β 遮断で心不全・徐脈が悪化', ['β 遮断', null]),
      mark('アドレナリンの濃度・桁の間違い', ['アドレナリン', null]),
      mark('原因（浅い麻酔、高 CO2、膀胱）を忘れて薬で押さえる', ['原因', null]),
    ], { y: 1.98, size: 28, gap: 26 });
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const id = 'S10';
    const s = contentSlide(deck, id, 4);
    bulletList(s, id, [
      mark('後負荷↓：ニカルジピン 2〜10 μg/kg/分で開始', ['後負荷↓：', null], ['2〜10 μg/kg/分', DOSE]),
      mark('前負荷↓：ニトログリセリン 0.5〜5 μg/kg/分で開始', ['前負荷↓：', null], ['0.5〜5 μg/kg/分', DOSE]),
      mark('心拍数↓：ランジオロール 0.125（1 分）→ 0.04 mg/kg/分', ['心拍数↓：', null], ['0.125', DOSE], ['0.04 mg/kg/分', DOSE]),
      mark('収縮力↑：ドブタミン 1〜5、最大 20 μg/kg/分', ['収縮力↑：', null], ['1〜5', DOSE], ['20 μg/kg/分', DOSE]),
      mark('アドレナリンは全部上げる（用量で変わる）。濃度と桁に注意', ['濃度と桁に注意', C.vermilion]),
    ], { y: 1.98, size: 26, gap: 26 });
  }

  checkOnSlide();
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
