'use strict';
/**
 * 3-4 オピオイド — 組版スクリプト
 *
 *   cd slides/3-4_opioids
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

const NAME = '3-4_opioids';
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
/** 出典の項目（spec の sources 行を ' / ' で分けたもの。2行への折り返しはライブラリ） */
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
    // 改行位置の空白（「維持：0.25 µg/kg/分」⏎「から調節…」）は落ちるので、つなぎ目の比較は空白を除いて行い、
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
const BOTTOM = SRC2_TOP - GAP;     // 出典2行のスライドの本文下端 6.26
const b = t => ({ text: t, bold: true });                        // 数値の強調（太字）
const danger = t => ({ text: t, bold: true, color: C.red });     // 危険・禁止（レッド）
const drug = t => ({ text: t, bold: true });                      // 行頭の薬剤名（色を付けると行頭記号の色も変わるので太字だけ）

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す） */
function contentSlide(deck, id, chapter) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id) });
  if (s.hasSources && s.sourceTop < SRC2_TOP - 1e-6) throw new Error(`${id}: 出典が3行以上`);
  return s;
}
/** そのスライドの本文の下端（出典の上 0.24in） */
const bottomOf = s => (s.hasSources ? s.sourceTop - GAP : LAYOUT.body.bottom);

/**
 * 見出し＋項目のカード。項目が2行以上になるときは、2行目以降を字下げする（ぶら下げ）。
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
  s.box({ x, y, w, h, fill: o.fill || C.bg, line: o.line, radius: 0.08 });
  const iw = w - 2 * pad;
  const th = heightFor(1, ts, 0);
  s.text(o.title, { x: x + pad, y: y + padY, w: iw, h: th, size: ts, bold: true, color: o.titleColor });
  rec(id, o.title);
  const top = y + padY + th + tg;
  const avail = y + h - padY - top;
  const lh = heightFor(1, bs, 0);
  const itemH = o.items.map(lines => lh + (lines.length > 1 ? heightFor(lines.length - 1, bs, 0) : 0));
  const sum = itemH.reduce((a, v) => a + v, 0);
  const n = o.items.length;
  let g = o.gap != null ? o.gap : (n > 1 ? Math.min(o.maxGap, (avail - sum) / (n - 1)) : 0);
  if (sum + g * (n - 1) > avail + 1e-6 || (o.gap == null && n > 1 && g < o.minGap - 1e-6)) {
    throw new Error(`${id} カード「${textOf(o.title)}」: 高さ不足（項目 ${sum.toFixed(2)}in + 間隔 > ${avail.toFixed(2)}in）`);
  }
  let cy = top;
  o.items.forEach((lines, k) => {
    s.text([lines[0]], { x: x + pad, y: cy, w: iw, h: lh, size: bs, color: C.dark });
    rec(id, lines[0]);
    if (lines.length > 1) {
      const rest = lines.slice(1);
      s.text(rest, { x: x + pad + indent, y: cy + lh, w: iw - indent, h: heightFor(rest.length, bs, 0), size: bs, color: C.dark });
      rest.forEach(l => rec(id, l));
    }
    cy += itemH[k] + (k < n - 1 ? g : 0);
  });
  return { gap: g, used: cy - y + padY };
}

/** 3列カードの共通寸法（本文 22pt。項目の2行目以降は 0.3in 字下げ） */
const CARD3 = { pad: 0.14, padY: 0.16, titleSize: 24, bodySize: 22, titleGap: 0.16, indent: 0.3, minGap: 0.08, maxGap: 0.34 };

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
  const widths = need.map((v, i) => (fixed.has(i) ? v : W));
  let x = X0;
  const h = bottom - TOP;
  cards.forEach((c, i) => {
    itemCard(s, id, { ...CARD3, ...c, x, y: TOP, w: widths[i], h });
    x += widths[i] + GAP;
  });
}

/** 強調ボックス（白地・ティールの枠・太字）。lines は1行（文字列/run 配列）または行の配列 */
function keyBox(s, id, lines, o) {
  const arr = (Array.isArray(lines) && lines.length && (Array.isArray(lines[0]) || o.multi)) ? lines : [lines];
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.white, line: { color: C.teal, width: 2.5 }, radius: 0.08,
    text: arr, size: o.size || 28, bold: true, color: C.dark, align: o.align || 'center' });
  arr.forEach(l => rec(id, l));
}

/** 箇条書き（置いた文字列を記録し、下端を確かめる） */
function bulletList(s, id, items, o) {
  items.forEach(t => rec(id, t));
  s.bullets(items, { y: o.y, size: o.size, gap: o.gap });
  const bh = heightFor(items.length, o.size, o.gap);
  if (o.y + bh > o.bottom + 1e-6) throw new Error(`${id}: 箇条書きが出典に近すぎます（下端 ${(o.y + bh).toFixed(2)}in > ${o.bottom.toFixed(2)}in）`);
  return o.y + bh;
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['共通', 'フェンタ\nニル', 'レミフェ\nンタニル', 'モルヒネ', '安全'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '3-4 オピオイド',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 作用と副作用（左右2枚の大きなカード） =====
  {
    const id = 'S2';
    const s = contentSlide(deck, id, 0);
    const cw = (BW - GAP) / 2;
    const geo = { y: TOP, w: cw, h: bottomOf(s) - TOP, pad: 0.18, padY: 0.22, titleSize: 28, bodySize: 24, titleGap: 0.26, indent: 0.36, minGap: 0.2, maxGap: 0.4 };
    const L = itemCard(s, id, {
      ...geo, x: X0, title: '作用（μ 受容体）', titleColor: C.blue,
      items: [
        ['鎮痛'],
        ['鎮静薬の必要量を減らす（相乗）'],
        ['刺激への血圧・心拍の', '反応を抑える'],
        ['咳の抑制'],
      ],
    });
    itemCard(s, id, {
      ...geo, x: X0 + cw + GAP, title: '副作用', titleColor: C.vermilion, gap: L.gap,
      items: [
        [[danger('呼吸抑制'), '：呼吸数↓、CO2 応答↓']],
        ['徐脈・血圧低下'],
        ['悪心嘔吐、掻痒、尿閉、便秘'],
        ['筋硬直・声門閉鎖（急速投与）'],
      ],
    });
  }

  // ===== S3 フェンタニル（3列カード） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 1);
    threeCards(s, id, [
      {
        title: '用量（目安）', titleColor: C.blue,
        items: [
          [['導入：', b('1〜3 µg/kg'), ' 静注']],
          ['術中の追加：', [b('25〜50 µg'), ' ずつ']],
          [['術後：', b('25〜50 µg'), ' ずつ、'], '反応を見て'],
        ],
      },
      {
        title: '動態', titleColor: C.blue,
        items: [
          ['効果のピークまで', [b('3〜6 分')]],
          ['単回では再分布で切れる'],
          ['長く入れると', '蓄積して切れにくい'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['術後の呼吸抑制が', '遅れて出る'],
          ['腎不全でも使える'],
          ['硬膜外・くも膜下', '投与もある'],
        ],
      },
    ], bottomOf(s));
  }

  // ===== S4 レミフェンタニル（3列カード） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 2);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          [['導入：', b('0.5 µg/kg/分')], ['（強い刺激なら ', b('1.0'), '）']],
          [['維持：', b('0.25 µg/kg/分')], ['から調節（最大 ', b('2.0'), '）']],
          ['浅麻酔時の追加：', [b('0.5〜1.0 µg/kg')], ['（', b('2〜5 分'), '間隔）']],
        ],
      },
      {
        title: '動態', titleColor: C.blue,
        items: [
          ['血液・組織の', 'エステラーゼで分解'],
          ['何時間入れても', [b('数分'), 'で切れる']],
          ['肝腎機能によらない'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['切れると痛みが急に戻る'],
          ['ボーラスで徐脈・筋硬直'],
          ['単独では麻酔にならない'],
          ['硬膜外・くも膜下', ['には', danger('禁止')]],
        ],
      },
    ], bottomOf(s));
  }

  // ===== S5 モルヒネ（3列カード） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 3);
    threeCards(s, id, [
      {
        title: '用量（目安）', titleColor: C.blue,
        items: [
          [['静注：', b('2〜5 mg'), ' ずつ、'], '反応を見て'],
          ['くも膜下：', [b('0.1〜0.2 mg')]],
          ['PCA の設定は', '施設の方式'],
        ],
      },
      {
        title: '動態', titleColor: C.blue,
        items: [
          ['効果のピークが遅く', ['（', b('15〜30 分'), '）、長い']],
          ['肝代謝、', '代謝物は腎排泄'],
          ['腎不全で蓄積（M6G）'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['ヒスタミン遊離', '（血管拡張・低血圧）'],
          ['くも膜下投与後の', '遅発性呼吸抑制'],
          ['腎不全では避ける'],
        ],
      },
    ], bottomOf(s));
  }

  // ===== S6 使い分け（表 2列 × 6行＋見出し行） =====
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 4);
    const head = ['場面', '選び方'];
    const body = [
      ['導入〜挿管', ['フェンタニル 1〜3 µg/kg（目安）、', 'またはレミフェンタニル持続']],
      ['術中', ['レミフェンタニル持続を刺激に合わせて調節']],
      ['覚醒前〜術後', ['フェンタニル少量、モルヒネ、区域麻酔']],
      ['腎不全', ['モルヒネを避け、フェンタニルを']],
      ['高齢者・OSA', ['減量し、呼吸を観察']],
      ['肥満', ['実体重で計算しない（除脂肪体重・標準体重）']],
    ];
    head.forEach(t => rec(id, t));
    body.forEach(([a, c]) => { rec(id, a); c.forEach(t => rec(id, t)); });
    // 行の高さ: 1行 0.556in、2行 1.06in（24pt・上下余白 2pt）。出典2行の上 0.24in まで
    const c1 = 2.9;
    const rowH = [0.556, 1.06, 0.556, 0.556, 0.556, 0.556, 0.556];
    const t = s.table([head, ...body.map(([a, c]) => [{ text: a, bold: true, fill: C.bg }, c.join('\n')])], {
      x: X0, y: TOP, colW: [c1, BW - c1], size: 24, rowH, margin: [2, 10, 2, 10],
    });
    if (t.y + t.h > bottomOf(s) + 1e-6) throw new Error(`S6: 表が出典に近すぎます（下端 ${(t.y + t.h).toFixed(2)}in）`);
  }

  // ===== S7 呼吸抑制を防ぐ（箇条書き5行＋下に強調ボックス） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 4);
    const bottom = bottomOf(s);
    const end = bulletList(s, id, [
      '量より反応：呼吸数、呼吸の深さ、SpO2、EtCO2',
      '鎮静薬・ベンゾジアゼピンとの併用で増強する',
      '術後は呼吸数を数える。酸素投与中は SpO2 だけでは分からない',
      '高齢者、OSA、肥満、心肺疾患、腎不全（モルヒネ）はハイリスク',
      'ナロキソンの場所と使い方を知っておく（→ 3-7）',
    ], { y: 1.9, size: 24, gap: 14, bottom });
    const kh = 0.8;
    const ky = bottom - kh;
    if (ky - end < GAP - 1e-6) throw new Error(`S7: 箇条書きと強調ボックスの間が狭すぎます（${(ky - end).toFixed(2)}in）`);
    keyBox(s, id, ['呼吸数が少ない（例：10 回/分未満）、反応が鈍い → ', danger('すぐ報告')],
      { x: X0, y: ky, w: BW, h: kh, size: 26 });
  }

  // ===== S8 麻薬としての取り扱い（箇条書き5行） =====
  {
    const id = 'S8';
    const s = contentSlide(deck, id, 4);
    bulletList(s, id, [
      '麻薬施用者の指示で、麻薬処方箋・施用記録に残す',
      '残液は捨てずに返納する。アンプルの空も返す',
      '紛失・破損は、すぐに報告して届け出る',
      'ケタミンも麻薬。ミダゾラム・レミマゾラムは向精神薬',
      '施錠した保管庫から、必要な分だけ',
    ], { y: 1.95, size: 28, gap: 26, bottom: bottomOf(s) });
  }

  // ===== S9 まとめ（箇条書き5行） =====
  {
    const id = 'S9';
    const s = contentSlide(deck, id, 4);
    bulletList(s, id, [
      '鎮痛と相乗作用。副作用の中心は呼吸抑制',
      [drug('フェンタニル'), '：ピークまで数分、繰り返すと蓄積'],
      [drug('レミフェンタニル'), '：数分で切れる。止める前に次の鎮痛を'],
      [drug('モルヒネ'), '：遅く長い。腎不全では避ける'],
      '術後は呼吸数を数える。ナロキソンの場所を知る',
    ], { y: 1.95, size: 28, gap: 20, bottom: bottomOf(s) });
  }

  checkOnSlide();
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
