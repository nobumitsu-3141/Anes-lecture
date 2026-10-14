'use strict';
/**
 * 3-7 オピオイドとベンゾジアゼピンの拮抗 — 組版スクリプト
 *
 *   cd slides/3-7_naloxone_flumazenil
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *     書き写しの検査: 各スライドに置いた文字列を順につなげたものが、spec の on-slide の「」の中身を
 *     順につなげたものと一字一句同じでなければ Error（checkOnSlide）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 *     1項目で1行（16pt・全角46字）に入らないもの（S2）は、語の切れ目で2行に分けて渡す（文言は同じ）。
 * spec からのずれ（deviations）:
 *   - S4 出典: 「ASA Task Force on Neuraxial Opioids, et al. Anesthesiology」→「ASA Task Force on Neuraxial Opioids. Anesthesiology」
 *     （3 件が 16pt・2 行に入らないため。理由の計算は S4 の箇所）。
 *   - S8 layout: 「右に各段の一言」は、spec に一言の文言が段2・3の括弧書きしか無いため採らず、
 *     括弧書きを箱の中に入れた中央寄せの縦フローにした（文言は spec のまま）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, LAYOUT, measure, heightFor } = require('../_lib/decklib');

const NAME = '3-7_naloxone_flumazenil';
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

/**
 * 出典の項目（spec の sources 行を ' / ' で分けたもの）。
 * split: {項目の先頭一致文字列: 2行目の書き出し} — 1項目が1行に入らないときだけ、その語の直前で2行に分ける。
 * 分けた2行は文字列としてつなげると元の項目と同じ（つなぎ目の半角スペースだけ落ちる）。
 */
function sourcesOf(id, split = {}, replace = null) {
  if (!SPEC[id].sources) return undefined;
  let src = SPEC[id].sources;
  if (replace) {
    // 表記の調整（2行に収まらないときだけ。deviations に記載）
    for (const [from, to] of replace) {
      if (!src.includes(from)) throw new Error(`${id}: 出典に「${from}」がありません`);
      src = src.replace(from, to);
    }
  }
  const items = [];
  for (const it of src.split(' / ')) {
    const key = Object.keys(split).find(k => it.startsWith(k));
    if (!key) { items.push(it); continue; }
    const k = it.indexOf(split[key]);
    if (k <= 0) throw new Error(`${id}: 出典「${it}」に分割位置「${split[key]}」がありません`);
    items.push(it.slice(0, k).replace(/ $/, ''), it.slice(k));
  }
  return items;
}

// ---------------------------------------------------------------------------
// 書き写しの検査: スライドに置いた文字列（置いた順）を記録し、spec の「」と突き合わせる
// ---------------------------------------------------------------------------
const PLACED = {};
/** 行は文字列か run 配列。字下げを足す行は {line, dx}（dx: 追加の字下げ in） */
const isDx = l => l != null && !Array.isArray(l) && typeof l === 'object' && 'line' in l;
const lineOf = l => (isDx(l) ? l.line : l);
const dxOf = l => (isDx(l) ? (l.dx || 0) : 0);
const textOf = line => {
  const l = lineOf(line);
  return Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(l);
};
function rec(id, line) { (PLACED[id] = PLACED[id] || []).push(textOf(line)); }
function checkOnSlide() {
  for (const [id, sp] of Object.entries(SPEC)) {
    if (!sp.quoted.length) continue;
    // 改行位置の空白は落ちるので、つなぎ目の比較は空白を除いて行い、
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
const ARROW = '404040';
const b = t => ({ text: t, bold: true });                         // 数値の強調（太字）
const em = (t, color) => ({ text: t, bold: true, color });        // 色付きの強調
const drug = t => ({ text: t, bold: true });                       // 行頭の薬剤名（色を付けると行頭記号の色も変わるので太字だけ）

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す） */
function contentSlide(deck, id, chapter, split, replace) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id, split, replace) });
  if (s.hasSources && s.sourceTop < SRC2_TOP - 1e-6) throw new Error(`${id}: 出典が3行以上`);
  return s;
}

/**
 * 見出し＋項目のカード。項目が2行以上になるときは、2行目以降を字下げする（ぶら下げ）。
 * 項目間は gap（既定は高さに応じて minGap〜maxGap）。各行は文字列か run 配列。
 * items: [[1行目, 2行目...], ...]
 */
function itemCardNeedW(o) {
  const ts = o.titleSize, bs = o.bodySize;
  let need = measure(textOf(o.title), ts);
  for (const lines of o.items) {
    lines.forEach((l, k) => { need = Math.max(need, measure(textOf(l), bs) + (k ? o.indent + dxOf(l) : 0)); });
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
      if (rest.some(l => dxOf(l))) {
        // 行ごとに字下げが違う（全角の開き括弧の後に揃える等）ときは1行ずつ置く（行送りは lh）
        rest.forEach((l, k) => {
          const dx = indent + dxOf(l);
          s.text([lineOf(l)], { x: x + pad + dx, y: cy + lh * (k + 1), w: iw - dx, h: lh, size: bs, color: C.dark });
          rec(id, l);
        });
      } else {
        s.text(rest, { x: x + pad + indent, y: cy + lh, w: iw - indent, h: heightFor(rest.length, bs, 0), size: bs, color: C.dark });
        rest.forEach(l => rec(id, l));
      }
    }
    cy += itemH[k] + (k < n - 1 ? g : 0);
  });
  return { gap: g, used: cy - y + padY };
}

/** 3列カードの共通寸法（本文 22pt。左右の内側余白 0.18in。項目の2行目以降は 0.26in 字下げ） */
const CARD3 = { pad: 0.18, padY: 0.16, titleSize: 24, bodySize: 22, titleGap: 0.16, indent: 0.26, minGap: 0.08, maxGap: 0.34 };

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

/** フローのチップ（灰色の角丸・ブルーの枠・太字）。lines は行の配列 */
function chip(s, id, lines, o) {
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.bg, line: { color: C.blue, width: 2 }, radius: 0.1,
    text: lines, size: o.size, bold: true, color: C.dark, align: o.align || 'center', margin: o.margin });
  lines.forEach(l => rec(id, l));
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['考え方', 'ナロ\nキソン', 'フルマ\nゼニル', '落とし穴', '手順'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '3-7 オピオイドとベンゾジアゼピンの拮抗',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 拮抗薬を使う前に（横フロー 4 段＋下に一文） =====
  {
    const id = 'S2';
    // 出典は1項目で1行に入らないので「第4版」」の後で2行に分ける
    const s = contentSlide(deck, id, 0, { '日本麻酔科学会': 'Ⅱ 鎮痛薬' });
    const steps = [
      ['気道・', '換気を確保'],
      ['原因を', '見積もる'],
      ['拮抗薬は', '少量ずつ'],
      ['切れるまで', '観察'],
    ];
    // チップ幅 2.3in（28pt・全角5字＝2.18in が内側余白 0.06in で入る）。矢印は前後に 0.24in 空けて太く
    const cw = 2.3, ag = (BW - 4 * cw) / 3, chh = 1.5, ins = GAP;
    const keyH = 1.1, between = 0.75;
    const y1 = TOP + (BOTTOM - TOP - (chh + between + keyH)) / 2;
    steps.forEach((lines, i) => {
      const x = X0 + i * (cw + ag);
      chip(s, id, lines, { x, y: y1, w: cw, h: chh, size: 28, margin: 0.06 });
      if (i < steps.length - 1) {
        const yc = y1 + chh / 2;
        s.arrow({ x1: x + cw + ins, y1: yc, x2: x + cw + ag - ins, y2: yc, color: ARROW, width: 4.5 });
      }
    });
    keyBox(s, id, ['拮抗薬は作動薬より', em('早く切れる', C.vermilion)],
      { x: X0, y: y1 + chh + between, w: BW, h: keyH, size: 32 });
  }

  // ===== S3 ナロキソン（3列カード） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 1);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          [['1 回 ', b('0.2 mg'), ' 静注']],
          ['効果不十分なら', [b('2〜3 分'), '間隔で'], [b('0.2 mg'), ' を'], [b('1〜2 回'), '追加']],
          ['状態に応じて適宜増減'],
        ],
      },
      {
        title: '特徴', titleColor: C.blue,
        items: [
          ['μ受容体の', '競合的拮抗薬'],
          [[b('1〜2 分'), 'で効き始める']],
          [[b('30〜45 分'), 'で切れる'], '（量による）'],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['急な覚醒：激痛、', '高血圧、頻脈、肺水腫'],
          ['切れたあとの', '再呼吸抑制'],
          ['慢性使用者では', '離脱症状'],
        ],
      },
    ], BOTTOM);
  }

  // ===== S4 ナロキソンの使いどころ（左右2枚のカード） =====
  {
    const id = 'S4';
    // 出典: spec の 3 件（全角換算 94.9 字）は 16pt・2 行（11.6in×2 = 93.2 字）にどこで割っても収まらない。
    // ASA の項の「, et al」を落とし（団体著者名だけにする）、「Anesthesiology」の後で2行に割る。
    //   各項目は 17.4 / 39.1 / 36.1 字で、どの2項目も1行（46.6 字）に並ばない（添付文書＋学会指針 = 54.7 字）。
    //   学会指針の行に残る幅は 9.3 字だけなので、ASA の項を割れるのは年・巻・頁（8.4 字）の前だけ。
    const s = contentSlide(deck, id, 1, { 'ASA Task Force': '2016;124' },
      [['ASA Task Force on Neuraxial Opioids, et al. Anesthesiology', 'ASA Task Force on Neuraxial Opioids. Anesthesiology']]);
    const lw = 6.3, rw = BW - GAP - lw;
    const geo = { y: TOP, h: BOTTOM - TOP, pad: 0.2, padY: 0.2, titleSize: 28, bodySize: 24, titleGap: 0.22, indent: 0.36, minGap: 0.2, maxGap: 0.36 };
    const L = itemCard(s, id, {
      ...geo, x: X0, w: lw, title: '使う', titleColor: C.blue,
      items: [
        // 3行目は2行目の「（」の後（「8」の位置）に揃える: 全角1字 = 本文 24pt の 1em
        ['呼吸数が少ない', ['（', b('8〜10 回/分未満'), 'など。'], { line: '施設基準で）', dx: 24 / 72 }],
        ['刺激しても呼吸が増えない'],
        ['SpO2 が下がり、換気で支えている'],
      ],
    });
    itemCard(s, id, {
      ...geo, x: X0 + lw + GAP, w: rw, title: '使わない・待つ', titleColor: C.vermilion, gap: L.gap,
      items: [
        ['刺激すれば呼吸する', '（呼びかけ、肩たたき）'],
        ['痛みが強く鎮痛を残したい'],
        ['原因が別にありそう'],
      ],
    });
  }

  // ===== S5 フルマゼニル（3列カード） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 2);
    threeCards(s, id, [
      {
        title: '用量（添付文書）', titleColor: C.blue,
        items: [
          // 半角英数の多い行（「0.1 mg、総量 1 mg まで」）は実機メイリオで幅が出るので、短い行に分ける
          [['初回 ', b('0.2 mg'), ' を'], '緩徐に静注'],
          [[b('4 分'), '以内に覚醒なければ'], [b('0.1 mg'), ' 追加']],
          [['以後 ', b('1 分'), '間隔で'], [b('0.1 mg'), '、'], ['総量 ', b('1 mg'), ' まで']],
        ],
      },
      {
        title: '特徴', titleColor: C.blue,
        items: [
          ['ベンゾジアゼピン', '受容体の', '競合的拮抗薬'],
          [[b('1〜2 分'), 'で効き始める']],
          [[b('30〜60 分'), 'で切れる']],
        ],
      },
      {
        title: '注意', titleColor: C.vermilion,
        items: [
          ['再鎮静（ミダゾラムの', '長時間投与後）'],
          ['痙攣（依存者、', '三環系抗うつ薬中毒）'],
          ['効かなければ', '別の原因'],
        ],
      },
    ], BOTTOM);
  }

  // ===== S6 レミマゾラムの拮抗（箇条書き4行＋強調ボックス） =====
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 2);
    const items = [
      [drug('レミマゾラム'), 'は', drug('フルマゼニル'), 'で拮抗できる'],
      '消失は速いが、拮抗後の再鎮静は起こりうる（長時間投与で多い）',
      '拮抗したら観察を続ける',
      '使うのは、覚醒遅延や呼吸抑制で本当に必要なときだけ',
    ];
    items.forEach(t => rec(id, t));
    const by = 1.92, gap = 16;
    s.bullets(items, { y: by, size: 24, gap });
    const bh = heightFor(items.length, 24, gap);
    const kh = 1.22;
    const ky = by + bh + 0.36;
    keyBox(s, id, [
      ['拮抗で『覚めた』ではなく、'],
      [em('『自分で呼吸して覚めている』', C.vermilion), 'を確認'],
    ], { x: X0, y: ky, w: BW, h: kh, size: 28 });
    if (ky + kh > BOTTOM + 1e-6) throw new Error(`S6: 強調ボックスが出典に近すぎます（下端 ${(ky + kh).toFixed(2)}in）`);
  }

  // ===== S7 落とし穴（箇条書き5行） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 3);
    const items = [
      [b('切れたあとに戻る'), '：病棟に帰してから呼吸が止まる'],
      [b('一気に入れる'), '：激痛、高血圧、肺水腫、痙攣'],
      [b('拮抗薬で診断する'), '：効かないのに他の原因を探さない'],
      'オピオイドとベンゾジアゼピンの両方が効いている',
      '拮抗した事実を申し送らない',
    ];
    items.forEach(t => rec(id, t));
    const gap = 24, by = 1.92;
    s.bullets(items, { y: by, size: 28, gap });
    const bh = heightFor(items.length, 28, gap);
    if (by + bh > BOTTOM + 1e-6) throw new Error(`S7: 箇条書きが出典に近すぎます（下端 ${(by + bh).toFixed(2)}in）`);
  }

  // ===== S8 手順（縦フロー 5 段） =====
  {
    const id = 'S8';
    const s = contentSlide(deck, id, 4);
    // spec の layout は「右に各段の一言」だが、on-slide に一言の文言があるのは段2・3の括弧書きだけ。
    // 右列が2段分しか埋まらず左に偏るので、括弧書きは箱の中（本文の後ろ・ブルー）に戻し、
    // 5段を同じ幅の箱にしてスライド中央に置く（矢印も中央）。
    const rows = [
      ['気道を開け、酸素、必要なら換気', null],
      ['原因を並べる', '（薬・筋弛緩・CO2・体温・糖）'],
      ['拮抗薬を少量ずつ', '（反応を見て追加）'],
      ['呼吸数・反応を再評価', null],
      ['作動薬が切れるまで観察し、申し送る', null],
    ];
    const size = 26, chh = 0.56, cw = 9.4, m = 0.2;
    const n = rows.length;
    const bottom = s.sourceTop - GAP;
    const vg = (bottom - TOP - n * chh) / (n - 1);   // 段の間隔（矢印の入る隙間）
    const cx = X0 + (BW - cw) / 2, ax = X0 + BW / 2, ains = 0.06;
    if (vg < 0.36) throw new Error(`S8: 段の間隔 ${vg.toFixed(2)}in が狭すぎます`);
    rows.forEach(([main, side], i) => {
      const y = TOP + i * (chh + vg);
      const line = side ? [main, { text: side, color: C.blue }] : main;
      chip(s, id, [line], { x: cx, y, w: cw, h: chh, size, margin: m });
      if (i < n - 1) s.arrow({ x1: ax, y1: y + chh + ains, x2: ax, y2: y + chh + vg - ains, color: ARROW, width: 4.5 });
    });
  }

  // ===== S9 まとめ（箇条書き5行） =====
  {
    const id = 'S9';
    const s = contentSlide(deck, id, 4);
    const items = [
      '使う前に換気と原因の見積もり',
      [drug('ナロキソン'), '：0.2 mg/A を希釈し分割', em('（添付文書外）', C.vermilion), '。鎮痛は残す'],
      [drug('フルマゼニル'), '：0.2 mg → 0.1 mg ずつ、総量 1 mg まで'],
      'どちらも作動薬より先に切れる。観察と申し送り',
      '効かなければ、薬以外の原因を探す',
    ];
    items.forEach(t => rec(id, t));
    const gap = 28, by = 1.95;
    s.bullets(items, { y: by, size: 24, gap });
    const bh = heightFor(items.length, 24, gap);
    if (by + bh > BOTTOM + 1e-6) throw new Error(`S9: 箇条書きが出典に近すぎます（下端 ${(by + bh).toFixed(2)}in）`);
  }

  checkOnSlide();
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
