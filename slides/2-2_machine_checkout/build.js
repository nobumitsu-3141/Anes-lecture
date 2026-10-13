'use strict';
/**
 * 2-2 始業点検の実際 — 組版スクリプト
 *
 *   cd slides/2-2_machine_checkout
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *   - 章チップ・タイトル枠・表紙は decklib の既定（全講義で共通）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 *     S10 は1項目が1行に入らないため、URL の前の空白で2項目に分けた（文言は同じ）。
 * render（PDF/JPG）はリポジトリの外に出す: RENDER_DIR / DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LINT, fits, measure, units, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '2-2_machine_checkout';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.RENDER_DIR || process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);
const RENDER = !process.argv.includes('--no-render');

// ---------------------------------------------------------------------------
// spec.md からノートと出典を読む（原文のまま）
// ---------------------------------------------------------------------------
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const lines = part.split('\n');
    let notes = null, sources = null;
    for (let i = 0; i < lines.length; i++) {
      if (/^notes:\s*\|\s*$/.test(lines[i])) {
        const buf = [];
        for (let j = i + 1; j < lines.length && (lines[j].startsWith('  ') || lines[j] === ''); j++) {
          buf.push(lines[j].replace(/^ {2}/, ''));
        }
        while (buf.length && buf[buf.length - 1] === '') buf.pop();
        notes = buf.join('\n');
      }
      const m = lines[i].match(/^sources:\s*(.*)$/);
      if (m) sources = m[1].trim();
    }
    out[id] = { notes, sources: sources && sources !== '—' ? sources : null };
  }
  return out;
}
const SPEC = parseSpec(fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8'));
function notesOf(id) {
  const n = SPEC[id] && SPEC[id].notes;
  if (!n) throw new Error(`spec.md の ${id} に notes がありません`);
  return n;
}
/**
 * 出典: spec の sources 行を ' / ' 区切りの項目にしてライブラリへ（詰めて最大2行）。
 * 1項目が16pt・11.6in の1行に入らないときは、URL の前の空白で分ける（文言は変えない）。
 */
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  if (!s) return undefined;
  const S = LAYOUT.source;
  const items = [];
  for (const it of s.split(' / ')) {
    if (fits(it, S.size, S.w)) { items.push(it); continue; }
    const k = it.lastIndexOf(' http');
    if (k < 0) throw new Error(`${id} 出典「${it}」が1行に入らず、分ける位置がありません`);
    items.push(it.slice(0, k), it.slice(k + 1));
  }
  return items;
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const ARROW = C.chipText;        // 矢印の色 404040
const em = (text, color, bold = true) => ({ text, color, bold });
/** S4・S5・S8 の3列カードの共通寸法（内側の左右余白 0.24in・見出し 24pt・本文 22pt） */
const CARD3 = { y: 1.85, pad: 0.24, padY: 0.2, titleSize: 24, bodySize: 22, titleGap: 0.22, itemGap: 0.4 };
/** S9・S10 の箇条書きの文字サイズ（S10 の矢印の行が1行に入る最大。同じサイズなら字下げも同じ） */
const BULLET_PT = 25;

const CHAPTERS = ['準備', 'ガス', '回路', '呼吸器', '記録'];   // 章チップは decklib の既定（全講義で共通）

/** 本文スライド: decklib の slide()（章チップ・タイトルの検査はライブラリ） */
function contentSlide(deck, o) {
  return deck.slide({ title: o.title, chapter: o.chapter, notes: o.notes, sources: o.sources });
}

/** レイアウト（pptxgenjs のスライドマスタ）のタイトル用プレースホルダ。スライド側の位置はこの値で決まる */
function titlePlaceholder(deck, name) {
  const lay = deck.pptx.slideLayouts.find(l => l._name === name);
  const ph = lay && lay._slideObjects.find(ob => ob.options && ob.options.placeholder === 'title');
  if (!ph) throw new Error(`レイアウト ${name} のタイトル枠が見つかりません`);
  return ph;
}

/** 指定した文字の run だけ書体を変える（メイリオに字形のない記号用） */
function setRunFont(s, ch, face) {
  let n = 0;
  for (const ob of s.raw._slideObjects) {
    if (!Array.isArray(ob.text)) continue;
    for (const r of ob.text) if (r && r.text === ch && r.options) { r.options.fontFace = face; n++; }
  }
  if (n === 0) throw new Error(`「${ch}」の run が見つかりません`);
}

/**
 * 見出し＋項目のカード。1項目を2行に割るときは項目内の行間を詰め、項目間だけを空ける
 * （card() は全段落に同じ段落後スペースを付けるため、割った行が別項目に見える）。
 * items: [[行, 行], [行], ...]（各行は文字列 or run 配列）
 */
function itemCard(s, o) {
  const pad = o.pad != null ? o.pad : 0.12;
  const padY = o.padY != null ? o.padY : pad;
  const ts = o.titleSize || 26, bs = o.bodySize || 24;
  const ig = o.itemGap != null ? o.itemGap : 0.3;
  const tg = o.titleGap != null ? o.titleGap : 0.2;
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill || C.bg, radius: 0.08, line: o.line });
  const iw = o.w - 2 * pad;
  let y = o.y + padY;
  const th = heightFor(1, ts, 0);
  s.text(o.title, { x: o.x + pad, y, w: iw, h: th, size: ts, bold: true, color: o.titleColor || C.blue });
  y += th + tg;
  if (o.bodyY != null) {   // 列どうしで1項目目の高さをそろえる
    if (o.bodyY < y - 1e-6) throw new Error(`itemCard「${o.title}」: bodyY ${o.bodyY} が見出しに掛かります`);
    y = o.bodyY;
  }
  o.items.forEach((lines, k) => {
    if (o.rowY) {   // 列どうしで各項目の行位置をそろえる
      if (o.rowY[k] < y - 1e-6) throw new Error(`itemCard「${o.title}」: rowY[${k}] が前の項目に掛かります`);
      y = o.rowY[k];
    }
    const h = heightFor(lines.length, bs, 0);
    s.text(lines, { x: o.x + pad, y, w: iw, h, size: bs, color: o.bodyColor || C.dark });
    y += h + (k < o.items.length - 1 ? ig : 0);
  });
  if (y > o.y + o.h - padY + 1e-6) throw new Error(`itemCard「${o.title}」: 高さ不足（必要 ${(y + padY - o.y).toFixed(2)}in > ${o.h}in）`);
}
const itemsH = (items, bs, ig) => items.reduce((a, l) => a + heightFor(l.length, bs, 0), 0) + ig * (items.length - 1);
/** 見出しの下の残りの高さに、列のうち最も高い項目群を置いたときの開始位置（余白の share を上に） */
function bodyStart(cards, o) {
  const top = o.y + o.padY + heightFor(1, o.titleSize, 0) + o.titleGap;
  const avail = o.y + o.h - o.padY - top;
  const need = Math.max(...cards.map(c => itemsH(c.items, o.bodySize, o.itemGap)));
  if (need > avail + 1e-6) throw new Error(`bodyStart: 高さ不足（必要 ${need.toFixed(2)} > ${avail.toFixed(2)}in）`);
  return top + (avail - need) * (o.share != null ? o.share : 0.4);
}
/**
 * 3列カード（同じ幅・同じ高さ）。各項目の行位置を3枚でそろえる（行の高さ = その行で最も行数の多い項目）。
 * geo.h を省くと内容に合わせた高さにする。
 */
function threeCards(s, cards, geo) {
  const gap = 0.24, cw = (BW - 2 * gap) / 3;
  const n = cards[0].items.length;
  if (cards.some(c => c.items.length !== n)) throw new Error('threeCards: 3枚の項目数が違います');
  const rowH = Array.from({ length: n }, (_, k) => Math.max(...cards.map(c => heightFor(c.items[k].length, geo.bodySize, 0))));
  const need = rowH.reduce((a, b) => a + b, 0) + geo.itemGap * (n - 1);
  const top = geo.y + geo.padY + heightFor(1, geo.titleSize, 0) + geo.titleGap;
  const h = geo.h != null ? geo.h : Math.ceil((top - geo.y + need + geo.padY) * 100) / 100;
  const avail = geo.y + h - geo.padY - top;
  if (need > avail + 1e-6) throw new Error(`threeCards: 高さ不足（必要 ${need.toFixed(2)} > ${avail.toFixed(2)}in）`);
  let y = top + (avail - need) * (geo.share != null ? geo.share : 0.4);
  const rowY = rowH.map(rh => { const v = y; y += rh + geo.itemGap; return v; });
  cards.forEach((c, i) => itemCard(s, { ...geo, h, x: X0 + i * (cw + gap), w: cw, rowY, ...c }));
  return { top: geo.y, bottom: geo.y + h };
}

/**
 * 縦に並べた番号つき手順チップ。steps の各要素は1行の文字列、または行の配列（2行に割る手順）。
 * チップ（塗り箱）の内側に、丸数字と本文を別のテキスト枠で置く（2行目の頭を本文の1行目にそろえるため）。
 * 1行のチップは高さ h、n 行のチップは本文の高さ＋上下の余白。戻り値: {top, bottom}
 */
function stepChips(s, steps, o) {
  const { x, w, h, gap } = o;
  const size = o.size || 24;
  const padL = 0.08, numW = 0.44, padR = 0.05, tx = x + padL + numW + 0.08, tw = x + w - padR - tx;
  const lh1 = heightFor(1, size, 0);
  let y = o.y;
  steps.forEach((st, i) => {
    const lines = Array.isArray(st) ? st : [st];
    const m = lines[0].match(/^([\u2460-\u2473])\s(.*)$/);
    if (!m) throw new Error(`手順「${lines[0]}」: 先頭に丸数字が必要です`);
    const th = heightFor(lines.length, size, 0);
    const ch = lines.length === 1 ? h : th + (h - lh1);
    const ty = y + (ch - th) / 2;
    s.box({ x, y, w, h: ch, fill: C.bg, radius: 0.08 });
    s.text(m[1], { x: x + padL, y: ty, w: numW, h: lh1, size, bold: true, color: C.blue });
    s.text([m[2], ...lines.slice(1)], { x: tx, y: ty, w: tw, h: th, size, color: C.dark });
    y += ch + (i < steps.length - 1 ? gap : 0);
  });
  return { top: o.y, bottom: y };
}

// ---------------------------------------------------------------------------
async function main() {
  // 章チップは contentSlide() が置く（decklib には chapters を渡さない）
  const deck = new Deck({
    file: OUT,
    chapters: CHAPTERS,
    author: '初期研修医向け麻酔講義',
    deckTitle: '2-2 始業点検の実際',
  });

  // ===== S1 表紙 =====
  {
    const cv = deck.cover({
      title: '始業点検の実際',
      subtitle: '2-2 モニタリング・麻酔器（OJT）',
      meta: '初期研修医向け麻酔講義｜日本麻酔科学会\n「麻酔器の始業点検」（2022年10月改訂）に沿って',
      notes: notesOf('S1'),
    });
  }

  // ===== S2 なぜ毎朝点検するのか（左に大きな原則カード、右に縦2枚のカード） =====
  {
    const s = contentSlide(deck, { title: 'なぜ毎朝点検するのか', chapter: 0, notes: notesOf('S2'), sources: sourcesOf('S2') });
    // 出典が2行（y6.5）なので本文は y6.26 まで
    const top = 1.85, gap = 0.24, lw = 6.55, rx = X0 + lw + 0.3, rw = XR - rx;
    // 右: 見出し24pt・本文22pt（2枚を縦に積んで y6.26 に収めるため本文は22pt）
    const rgeo = { pad: 0.15, padY: 0.1, titleSize: 24, bodySize: 22, titleGap: 0.1, itemGap: 0.08 };
    const right = [
      { title: '見つけたい異常', titleColor: C.blue, items: [['漏れ・外れ'], ['ガスの取り違え'], ['弁の固着・吸収剤の消耗']] },
      { title: '離島では', titleColor: C.vermilion, items: [['予備機がない'], ['代わりの換気手段を毎朝確認']] },
    ];
    const cardH = c => 2 * rgeo.padY + heightFor(1, rgeo.titleSize, 0) + rgeo.titleGap + itemsH(c.items, rgeo.bodySize, rgeo.itemGap);
    let y = top;
    right.forEach(c => {
      const h = Math.ceil(cardH(c) * 100) / 100;
      itemCard(s, { ...rgeo, ...c, x: rx, y, w: rw, h });
      y += h + gap;
    });
    const bottom = y - gap;
    if (bottom > 6.26 + 1e-6) throw new Error(`S2: 右のカードが出典行に近すぎます（下端 ${bottom.toFixed(2)}in）`);
    // 左: 原則（28pt。2行目は語の切れ目で2行に）
    const lgeo = { y: top, h: bottom - top, pad: 0.22, padY: 0.2, titleSize: 28, bodySize: 28, titleGap: 0.3, itemGap: 0.35 };
    const left = {
      title: '原則', titleColor: C.teal, fill: C.white, line: { color: C.teal, width: 3 },
      items: [
        ['使う人が、使う前に、毎回'],
        ['セルフチェック機能があっても', '手で確かめる項目がある'],
        [[em('直らない異常があれば使わない', C.red)]],
      ],
    };
    const bodyY = bodyStart([left], { ...lgeo, share: 0.35 });
    itemCard(s, { ...lgeo, ...left, x: X0, w: lw, bodyY });
  }

  // ===== S3 点検の順番（6つのチップのフロー。2段に折り返す＋下に一文） =====
  {
    const s = contentSlide(deck, { title: '点検の順番', chapter: 0, notes: notesOf('S3'), sources: sourcesOf('S3') });
    const steps = [
      ['① ボンベ・', '流量計'],
      ['② 遮断機構・', '配管'],
      ['③ 気化器・', '酸素濃度計'],
      ['④ 吸収剤・', '回路のリーク'],
      ['⑤ 人工呼吸器・', 'アラーム'],
      ['⑥ 吸引・予備・', '記録'],
    ];
    const ag = 0.85, cw = (BW - 2 * ag) / 3, chh = 1.3, ins = 0.24;
    const y1 = 1.95, y2 = y1 + chh + 0.72;
    const xs = [0, 1, 2].map(k => X0 + k * (cw + ag));
    steps.forEach((lines, i) => {
      const x = xs[i % 3], y = i < 3 ? y1 : y2;
      const [num, rest] = [lines[0].slice(0, 1), lines[0].slice(1)];
      s.box({
        x, y, w: cw, h: chh, fill: C.bg, radius: 0.1, line: { color: C.blue, width: 2 }, margin: 0.08,
        text: [[em(num, C.blue), rest], lines[1]], size: 26, bold: true, color: C.dark,
      });
      if (i % 3 < 2) {
        const yc = y + chh / 2;
        s.arrow({ x1: x + cw + ins, y1: yc, x2: x + cw + ag - ins, y2: yc, color: ARROW, width: 3 });
      }
    });
    // ③ → ④ の折り返し（③の下端中央 → 左へ → ④の上端中央）
    const c3 = xs[2] + cw / 2, c4 = xs[0] + cw / 2, yA = y1 + chh + 0.1, ym = y1 + chh + 0.36;
    s.vline({ x: c3, y: yA, h: ym - yA, color: ARROW, width: 3 });
    s.hline({ x: c4, y: ym, w: c3 - c4, color: ARROW, width: 3 });
    s.arrow({ x1: c4, y1: ym, x2: c4, y2: y2 - 0.1, color: ARROW, width: 3 });
    // 一文
    s.text([['上流（ガスの供給）から下流（患者側）へ、', em('同じ順番で毎回', C.teal)]], {
      x: X0, y: y2 + chh + 0.36, w: BW, size: 28, color: C.dark, align: 'center',
    });
  }

  // ===== S4 ガス供給の点検（3列カード） =====
  {
    const s = contentSlide(deck, { title: 'ガス供給の点検', chapter: 1, notes: notesOf('S4'), sources: sourcesOf('S4') });
    // 3列とも同じ幅。1行に入らない項目は語の切れ目で2行に（数値と単位は割らない）
    const cards = [
      { title: '補助ボンベ・流量計', items: [['酸素ボンベを開けて', [em('5 MPa 以上', C.dark)]], ['酸素 5 L/分 が流れる'], ['流量計が全範囲で動く']] },
      { title: '遮断機構', items: [['酸素と亜酸化窒素を', '5 L/分'], ['酸素ボンベを閉じる'], ['アラームが鳴り、', '亜酸化窒素が止まる']] },
      { title: '配管', items: [['ホースの目視と接続'], [['供給圧 ', em('392 ± 49 kPa', C.dark)]], ['排除装置もつなぐ']] },
    ];
    // カード内の左右の余白 0.24in（内側幅 3.44in）。見出し 24pt・本文 22pt
    // （22pt「流量計が全範囲で動く」3.42in が入る。24pt では 3.73in で入らない）。各項目の行位置は3枚でそろえる
    threeCards(s, cards, CARD3);
  }

  // ===== S5 気化器・酸素濃度計・吸収剤（3列カード） =====
  {
    // タイトルは spec どおり（全角中黒・全角13字）。章チップ列を組み直してチップに掛からない
    const s = contentSlide(deck, { title: '気化器・濃度計・吸収剤', chapter: 1, notes: notesOf('S5'), sources: sourcesOf('S5') });
    const cards = [
      { title: '気化器', items: [['電源・残量・注入栓'], ['OFF で酸素を流し、', 'においがしない'], ['ダイヤルが円滑に動く']] },
      { title: '酸素濃度計', items: [['室内気で 21% に校正'], ['酸素を流して', '100% 近く'], ['下限アラームを設定']] },
      { title: '吸収剤・回路', items: [['色・量・詰まり方を', '目で見る'], ['水抜きのあとは閉じる'], ['回路を組み、', '接続を確かめる']] },
    ];
    // 出典が2行（y6.5）なので y6.26 まで（3行とも2行の項目があるため、項目間を詰める）
    const r = threeCards(s, cards, { ...CARD3, itemGap: 0.26 });
    if (r.bottom > 6.26 + 1e-6) throw new Error(`S5: カードが出典行に近すぎます（下端 ${r.bottom.toFixed(2)}in）`);
  }

  // ===== S6 回路のリークテスト（縦5つの手順チップ＋右に赤の注意ボックス） =====
  {
    const s = contentSlide(deck, { title: '回路のリークテスト', chapter: 2, notes: notesOf('S6'), sources: sourcesOf('S6') });
    // チップ幅は S7 とそろえる（24pt。最長の⑤の本文 8.21in が1行に入る幅）。
    // ④は 24pt で 9.07in あり1行に入らないため「。」の後で2行に割る
    const cw = 8.88;
    const r = stepChips(s, [
      '① 新鮮ガス 0、APL 弁を閉じ、Y ピースを塞ぐ',
      '② 酸素 5〜10 L/分 で 30 cmH2O まで膨らます',
      '③ バッグを押して 40〜50 cmH2O、漏れがない',
      ['④ APL 弁を開いて圧が下がる。', '酸素フラッシュが出る'],
      '⑤ テスト肺をつけて用手換気、一方弁の動きを見る',
    ], { x: X0, y: 1.85, w: cw, h: 0.66, gap: 0.24 });
    // 注意ボックス: 26pt「回路や吸収剤を」(2.83in) が内側幅に入る
    const bx = X0 + cw + 0.28, bw = XR - bx, bh = 2.75;
    s.box({
      x: bx, y: (r.top + r.bottom) / 2 - bh / 2, w: bw, h: bh, fill: C.white, radius: 0.1,
      line: { color: C.red, width: 3 }, margin: 0.1,
      text: ['回路や吸収剤を', '交換したら、', 'もう一度', 'リークテスト'], size: 26, bold: true, color: C.red,
    });
  }

  // ===== S7 人工呼吸器とアラーム（縦5つの手順チップ＋右に補足カード） =====
  {
    const s = contentSlide(deck, { title: '人工呼吸器とアラーム', chapter: 3, notes: notesOf('S7'), sources: sourcesOf('S7') });
    // 最長の①の本文（24pt 8.21in）が1行に入る幅
    const cw = 8.88;
    const r = stepChips(s, [
      '① テスト肺をつけ、用手から人工呼吸に切り替える',
      '② 従量式で換気量・気道内圧の表示を見る',
      '③ 従量式：テスト肺を外して低圧アラーム',
      '④ 従量式：Y ピースを塞いで高圧アラーム',
      '⑤ 従圧式：回路を塞いで低換気量アラーム',
    ], { x: X0, y: 1.9, w: cw, h: 0.72, gap: 0.24 });
    // 補足カード: 左右の余白 0.24in（S4/S5/S8 のカードと同じ）。見出し 24pt・本文 22pt
    // （22pt「バッグ ↔ 呼吸器」が内側幅に入る）
    const kx = X0 + cw + 0.26, kw = XR - kx;
    const card = { title: '切り替え', titleColor: C.blue, items: [[['バッグ ', em('↔', C.blue), ' 呼吸器']], ['切り替えたら', '胸と EtCO2']] };
    const geo = { pad: 0.24, padY: 0.22, titleSize: 24, bodySize: 22, titleGap: 0.3, itemGap: 0.4 };
    const kh = Math.ceil((2 * geo.padY + heightFor(1, geo.titleSize, 0) + geo.titleGap + itemsH(card.items, geo.bodySize, geo.itemGap)) * 100) / 100;
    itemCard(s, { ...geo, ...card, x: kx, y: (r.top + r.bottom) / 2 - kh / 2, w: kw, h: kh });
  }

  // ===== S8 排除装置・吸引・予備（3列カード） =====
  {
    const s = contentSlide(deck, { title: '排除装置・吸引・予備', chapter: 3, notes: notesOf('S8'), sources: sourcesOf('S8') });
    const cards = [
      { title: '余剰ガス排除', items: [['吸引圧 1 kPa 以上', '2 kPa 未満'], ['または 25〜50 L/分'], ['接続が外れていない']] },
      { title: '吸引器', items: [['吸引圧が上がる'], ['カテーテルと予備'], ['手の届く所に']] },
      { title: '予備の換気', items: [['自己膨張式バッグ'], ['予備の酸素ボンベ'], ['マスクと', '経口エアウェイ']] },
    ];
    // 出典が2行（y6.5）なので y6.26 まで
    const r = threeCards(s, cards, CARD3);
    if (r.bottom > 6.26 + 1e-6) throw new Error(`S8: カードが出典行に近すぎます（下端 ${r.bottom.toFixed(2)}in）`);
  }

  // ===== S9 記録と異常時（箇条書き4行＋下に赤の強調ボックス） =====
  {
    const s = contentSlide(deck, { title: '記録と異常時', chapter: 4, notes: notesOf('S9'), sources: sourcesOf('S9') });
    s.bullets([
      '点検表にチェックし、日付と名前を書く',
      '直らない異常があれば使わない。臨床工学技士に連絡する',
      '別の麻酔器に替えるか、修理まで手術を延期する',
      '症例の合間に回路を替えたら、リークテストをやり直す',
    ], { y: 1.9, size: BULLET_PT, gap: 22 });
    s.box({
      x: X0, y: 5.3, w: BW, h: 0.9, fill: C.red, radius: 0.1, margin: 0.1,
      text: '麻酔器が止まっても、換気は止めない：自己膨張式バッグへ', size: 28, bold: true, color: C.white,
    });
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const s = contentSlide(deck, { title: 'まとめ', chapter: 4, notes: notesOf('S10'), sources: sourcesOf('S10') });
    // 2行目（矢印の流れ）が1行に入る最大の 25pt に全行をそろえる（S9 の箇条書きも同じ 25pt・同じ字下げ）
    s.bullets([
      '使う人が、使う前に、毎回、同じ順番で',
      'ボンベ → 配管 → 気化器・酸素濃度計 → 吸収剤・回路 → 呼吸器',
      '回路や吸収剤を替えたら、リークテストをやり直す',
      '酸素濃度計と予備の換気手段は毎朝確かめる',
      '直らない異常があれば使わない。記録に残す',
    ], { y: 1.95, size: BULLET_PT, gap: 30 });
  }

  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.render) console.log('[build] render: ' + res.render.images.length + ' images in ' + RENDER_DIR);
  if (res.violations !== 0) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
