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
 * 章チップ・タイトル枠・表紙・色は decklib の既定のまま（この講義では上書きしない）。
 * この講義だけの調整は出典の枠だけ: 文字の位置（下揃え・下端 7.35in）は既定のまま、枠の上の空きを詰めて
 * 枠の上端を文字の上端に合わせる（tightenSource）。3列カードの下端は、この枠の上端から 0.24in 空けて決める。
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
const b = t => ({ text: t, bold: true });                       // 数値の強調（太字）
const danger = t => ({ text: t, bold: true, color: C.red });    // 禁忌（レッド）
const drug = t => ({ text: t, bold: true });                     // まとめの薬剤名（色を付けると行頭記号の色も変わるので太字だけ）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;
const CM = 2.54;

// ---------------------------------------------------------------------------
// 出典の枠（この講義だけの調整）
// ---------------------------------------------------------------------------
const CHAPTERS = ['共通', 'プロポ\nフォール', 'レミマ\nゾラム', 'ミダ・\nケタ', '使い分け'];

/**
 * 出典。decklib は 1 行なら y6.85・h0.5、2 行なら y6.5・h0.85 の枠に下揃えで置く（どちらも下端 7.35in）。
 * 文字の位置は変えずに（下端 7.35in・下揃えのまま）、枠の高さを文字に必要な高さまで詰め、枠の上端を文字の上端に合わせる。
 * 必要な高さは lint の推定（1 行 0.40in、2 行 0.80in）から許容 0.3cm を引いた値と、実機メイリオの高さ（1 行 0.33in、
 * 2 行 0.67in）の大きいほう → 1 行 0.34in（上端 7.01in）、2 行 0.68in（上端 6.67in）。
 * 枠の幅は 11.4in（ページ番号の枠 12.30in との間を 0.24in 以上あける）。ライブラリは 11.6in で行を組むので、
 * 推定幅が 11.4in を超える行があるときはその行が入る幅にする。ページ番号はライブラリの既定のまま。
 */
const SRC_W = 11.4;
function srcBoxH(nLines) {
  const S = LAYOUT.source;
  const lintH = nLines * (S.size / 72 * 1.2 * 1.18 + 6 / 72) - 0.3 / CM;
  const realH = nLines * S.size * 1.5 / 72;
  return Math.ceil(Math.max(lintH, realH) * 100 - 1e-9) / 100;
}
const SRC_BOTTOM = r2(LAYOUT.source.y2 + LAYOUT.source.h2);   // 7.35（1 行の枠 y + h も同じ）
const SRC2_TOP = r2(SRC_BOTTOM - srcBoxH(2));                  // 出典 2 行の枠の上端 6.67
function tightenSource(s) {
  const S = LAYOUT.source, pg = LAYOUT.page;
  const objs = s.raw._slideObjects.filter(o => o._type === 'text' && o.options);
  const src = objs.filter(o => o.options.x === S.x && o.options.w === S.w && (o.options.y === S.y || o.options.y === S.y2));
  if (src.length !== 1) throw new Error(`p${s.page}: 出典が見つかりません（${src.length} 個）`);
  const o = src[0];
  if (Math.abs(o.options.y + o.options.h - SRC_BOTTOM) > 1e-6) throw new Error(`p${s.page}: 出典の枠の下端が ${SRC_BOTTOM}in ではありません`);
  if (o.options.valign !== 'bottom') throw new Error(`p${s.page}: 出典が下揃えではありません`);
  const lines = o.text.map(t => t.text);
  const maxW = Math.max(...lines.map(l => measure(l, S.size)));
  const w = Math.min(S.w, Math.max(SRC_W, Math.ceil(maxW * 100) / 100));
  const h = srcBoxH(lines.length);
  const y = r2(SRC_BOTTOM - h);
  if (o.options.x + w > pg.x - GAP + 1e-9) throw new Error(`p${s.page}: 出典の枠の右端 ${r3(o.options.x + w)}in がページ番号に近すぎます`);
  o.options.y = y;
  o.options.h = h;
  o.options.w = w;
  s.sourceTop = y;
  const it = s._items.find(i => i.desc === '出典');
  if (it) { it.y = y; it.h = h; it.w = w; }
  return s;
}

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す）＋出典の枠を詰める */
function contentSlide(deck, id, chapter) {
  const sp = spec(id);
  const s = deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id) });
  if (s.hasSources) tightenSource(s);
  return s;
}

/**
 * 見出し＋項目のカード。項目が2行になるときは、2行目以降を字下げする（ぶら下げ）。各行は文字列か run 配列。
 * 本文は見出しのすぐ下から上寄せで並べ、項目の間は gap（in）。項目の少ないカードは下を余らせる。
 * items: [[1行目, 2行目...], ...]
 */
const FILL = [];   // 各行の推定幅 / 枠幅（ログ用）
function itemCardNeedW(o) {
  const ts = o.titleSize, bs = o.bodySize;
  let need = measure(textOf(o.title), ts);
  for (const lines of o.items) {
    lines.forEach((l, k) => { need = Math.max(need, measure(textOf(l), bs) + (k ? o.indent : 0)); });
  }
  return need + 2 * o.pad + 0.001;
}
function itemCard(s, id, o) {
  const { x, y, w, h, pad, padY, titleSize: ts, bodySize: bs, titleGap: tg, indent, gap: g } = o;
  const padB = o.padB != null ? o.padB : padY;
  s.box({ x, y, w, h, fill: C.bg, radius: 0.08 });
  const iw = w - 2 * pad;
  const th = heightFor(1, ts, 0);
  const lh = heightFor(1, bs, 0);
  s.text(o.title, { x: x + pad, y: y + padY, w: iw, h: th, size: ts, bold: true, color: o.titleColor });
  rec(id, o.title);
  const top = y + padY + th + tg;
  const avail = y + h - padB - top;
  const itemH = o.items.map(lines => lh + (lines.length > 1 ? heightFor(lines.length - 1, bs, 0) : 0));
  const n = o.items.length;
  const used = itemH.reduce((a, v) => a + v, 0) + g * (n - 1);
  if (used > avail + 1e-6) {
    throw new Error(`${id} カード「${textOf(o.title)}」: 高さ不足（項目 ${used.toFixed(3)}in > ${avail.toFixed(3)}in）`);
  }
  let cy = top;
  o.items.forEach((lines, k) => {
    s.text([lines[0]], { x: x + pad, y: cy, w: iw, h: lh, size: bs, color: C.dark });
    rec(id, lines[0]);
    FILL.push({ id, text: textOf(lines[0]), r: measure(textOf(lines[0]), bs) / iw });
    if (lines.length > 1) {
      const rest = lines.slice(1);
      s.text(rest, { x: x + pad + indent, y: cy + lh, w: iw - indent, h: heightFor(rest.length, bs, 0), size: bs, color: C.dark });
      rest.forEach(l => { rec(id, l); FILL.push({ id, text: textOf(l), r: measure(textOf(l), bs) / (iw - indent) }); });
    }
    cy += itemH[k] + (k < n - 1 ? g : 0);
  });
  return { top, bottom: cy, spare: avail - used };
}

/**
 * 3列カード（S3〜S6）の共通寸法。本文 22pt、項目の2行目は 0.3in 字下げ。左右の内側余白は S2 のカードと同じ 0.18in。
 * - 本文は上寄せ。どのカードも本文の1行目の上端は CARD_TOP + 0.10 + 0.50 + 0.08 = 2.51in にそろう。
 * - 項目の間は 0.20in（全カード共通）。一番詰まる S4「特徴」（4 項目・7 行）が 7 × 0.46 + 3 × 0.20 = 3.82in で入る高さにする。
 * - 上下の内側余白 0.10in（行の枠の上下の空きと合わせて、見た目で左右の余白 0.18in と同じくらい）。
 * - カードの上端 1.83in（金下線の太さ 8pt の下の縁から 0.25in）、下端は出典 2 行の枠の上端 6.67in から 0.24in 上の 6.43in。
 */
const CARD3 = { pad: 0.18, padY: 0.10, padB: 0.10, titleSize: 24, bodySize: 22, titleGap: 0.08, indent: 0.3, gap: 0.20 };
const CARD_TOP = 1.83;
const CARD_BOTTOM = r2(SRC2_TOP - GAP);

/**
 * 3列カード。列幅は中身の最長行から決め、余りは狭い列から順に配る（水位合わせ）。
 * cards: [{title, titleColor, items}]。上端 CARD_TOP・下端 CARD_BOTTOM（S3〜S6 で共通）
 */
function threeCards(s, id, cards) {
  if (s.sourceTop < CARD_BOTTOM + GAP - 1e-6) throw new Error(`${id}: 出典（上端 ${s.sourceTop}in）がカードに近すぎます`);
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
  const h = r2(CARD_BOTTOM - CARD_TOP);
  const res = cards.map((c, i) => {
    const r = itemCard(s, id, { ...CARD3, ...c, x, y: CARD_TOP, w: widths[i], h });
    x += widths[i] + GAP;
    return r;
  });
  console.log(`[3-2] ${id}: 幅 ${widths.map(v => v.toFixed(2)).join(' / ')}in、本文の上端 ${res.map(r => r.top.toFixed(2)).join(' / ')}in、` +
    `下の余り ${res.map(r => r.spare.toFixed(2)).join(' / ')}in`);
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
    ]);
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
    ]);
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
    ]);
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
    ]);
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
    // 箇条書き（実機メイリオで下端 約 5.15in）と強調ボックス（5.45〜6.30in）の間を約 0.3in、
    // ボックスと出典（2 行、上端 6.67in）の間を約 0.37in にする
    const by = 1.9, gap = 22, kh = 0.85;
    s.bullets(items, { y: by, size: 28, gap });
    const bh = heightFor(items.length, 28, gap);
    const ky = r2(by + bh + GAP);
    // 強調は数値と単位だけ（バーミリオン＋下線。色だけに頼らない）
    keyBox(s, id, ['導入薬を入れたら、血圧は ', { text: '1〜2 分', color: C.vermilion, underline: true }, 'ごとに測る'],
      { x: X0, y: ky, w: BW, h: kh });
    if (ky + kh > s.sourceTop - GAP + 1e-6) throw new Error('S8: 強調ボックスが出典に近すぎます');
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
    // 5 項目とも 25pt（推定 11.4in）にそろえる（S8 の 28pt にそろえるには文言を削る必要があり、しない）。
    // 行間を広げて、出典（1 行、上端 7.01in）の上まで使う: 実機メイリオで下端 約 6.44in、枠（lint 推定）の下端 6.65in
    s.bullets(items, { y: 1.95, w: BW, size: 25, gap: 34 });
  }

  checkOnSlide();
  const tight = FILL.slice().sort((a, b) => b.r - a.r).slice(0, 5);
  console.log('[3-2] カード本文の行の推定幅 / 枠幅（上位 5）: ' + tight.map(f => `${f.id}「${f.text}」${(f.r * 100).toFixed(0)}%`).join('、'));
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
