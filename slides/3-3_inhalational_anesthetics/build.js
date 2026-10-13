'use strict';
/**
 * 3-3 吸入麻酔薬 — 組版スクリプト
 *
 *   cd slides/3-3_inhalational_anesthetics
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *     L(id, 文字列, 改行位置...) は文字列が spec の on-slide の「」にあることを確かめてから行に割る。
 *     最後に、置いた行を順につなげたものが spec の「」を順につなげたものと一字一句同じか検査する（checkOnSlide）。
 *   - 図の軸名・曲線名は spec の figure/layout 行にある語だけを使う（fig）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文。【要確認】もノートには残す）。
 *   - 出典は spec の sources 行を ' / ' で項目に分けて渡す（2行への折り返しはライブラリ）。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, LAYOUT, measure, heightFor } = require('../_lib/decklib');

const NAME = '3-3_inhalational_anesthetics';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || process.env.RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);
const RENDER = !process.argv.includes('--no-render');

// ---------------------------------------------------------------------------
// spec.md を読む（タイトル・on-slide の「」・図の指示・ノート・出典）
// ---------------------------------------------------------------------------
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const lines = part.split('\n');
    const title = (lines[0].match(/title=「([^」]*)」/) || [])[1] || null;
    let notes = null, sources = null, subtitle = null, meta = null, figText = '';
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
      if (/^(figure|layout):/.test(ln)) figText += ln + '\n';
    }
    out[id] = { title, quoted, notes, subtitle, meta, figText, sources: sources && sources !== '—' ? sources : null };
  }
  return out;
}
const SPEC = parseSpec(fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8'));
function spec(id) {
  if (!SPEC[id]) throw new Error(`spec.md に ${id} がありません`);
  if (!SPEC[id].notes) throw new Error(`spec.md の ${id} に notes がありません`);
  if (!SPEC[id].title) throw new Error(`spec.md の ${id} に title がありません`);
  return SPEC[id];
}
const sourcesOf = id => (SPEC[id].sources ? SPEC[id].sources.split(' / ') : undefined);

// ---------------------------------------------------------------------------
// 書き写しの検査
// ---------------------------------------------------------------------------
const PLACED = {};
const textOf = line => (Array.isArray(line) ? line.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(line));

/**
 * spec の on-slide の文言 str を、breaks の各語の直前で改行して行配列にする（置いた順に記録）。
 * str が spec に無ければ Error。改行位置の半角スペースは落とす。
 */
function L(id, str, ...breaks) {
  if (!SPEC[id].quoted.includes(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
  const lines = [];
  let from = 0;
  for (const b of breaks) {
    const k = str.indexOf(b, from + 1);
    if (k < 0) throw new Error(`${id}: 「${str}」に改行位置「${b}」が見つかりません`);
    lines.push(str.slice(from, k).replace(/ $/, ''));
    from = k;
  }
  lines.push(str.slice(from));
  (PLACED[id] = PLACED[id] || []).push(...lines);
  return lines;
}
/** 1行を run に分けて一部を装飾する。parts を連結すると元の行に一致すること */
function R(line, ...parts) {
  const joined = parts.map(p => (typeof p === 'string' ? p : p.text)).join('');
  if (joined !== line) throw new Error(`装飾の分割が元の行と一致しません「${joined}」≠「${line}」`);
  return parts.map(p => (typeof p === 'string' ? { text: p } : p));
}
/** 行配列 lines の k 行目を run に分ける */
function E(lines, k, ...parts) {
  const out = lines.slice();
  out[k] = R(lines[k], ...parts);
  return out;
}
/** 図の指示（figure / layout 行）にある語。照合のみ */
function fig(id, str) {
  if (!SPEC[id].figText.includes(str)) throw new Error(`${id}: 「${str}」は spec の figure/layout にありません`);
  return str;
}
function checkOnSlide() {
  for (const [id, sp] of Object.entries(SPEC)) {
    if (!sp.quoted.length) continue;
    const placed = PLACED[id] || [];
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
const XR = LAYOUT.body.right;      // 12.78
const BW = LAYOUT.body.w;          // 12.23
const TOP = 1.85;                  // 本文の上端
const GAP = 0.24;                  // 要素間の最小間隔
const BOT2 = LAYOUT.source.y2 - GAP; // 出典2行（y 6.5）のスライドの本文下端 6.26
const BOT1 = LAYOUT.source.y - GAP;  // 出典1行（y 6.85）のスライドの本文下端 6.61
const r3 = v => Math.round(v * 1000) / 1000;
const b = t => ({ text: t, bold: true });                       // 数値の強調（太字）
const danger = t => ({ text: t, bold: true, color: C.red });    // 危険（レッド）
const em = (t, color) => ({ text: t, bold: true, color });      // 色の強調
const L22 = heightFor(1, 22, 0);
const AXIS_C = C.gray;

/** 本文スライド（spec のタイトル・ノート・出典をそのまま渡す） */
function contentSlide(deck, id, chapter) {
  const sp = spec(id);
  return deck.slide({ title: sp.title, chapter, notes: sp.notes, sources: sourcesOf(id) });
}

/**
 * 見出し＋項目のカード。項目が2行になるときは、2行目以降を字下げする（ぶら下げ）。
 * items: [[1行目, 2行目...], ...]（各行は文字列か run 配列）
 */
function itemCardNeedW(o) {
  let need = measure(textOf(o.title), o.titleSize);
  for (const lines of o.items) {
    lines.forEach((l, k) => { need = Math.max(need, measure(textOf(l), o.bodySize) + (k ? o.indent : 0)); });
  }
  return need + 2 * o.pad + 0.001;
}
function itemCardNeedH(o) {
  const lines = o.items.reduce((a, it) => a + it.length, 0);
  return 2 * o.padY + heightFor(1, o.titleSize, 0) + o.titleGap + heightFor(1, o.bodySize, 0) * lines + o.minGap * (o.items.length - 1);
}
function itemCard(s, id, o) {
  const { x, y, w, h, pad, padY, titleSize: ts, bodySize: bs, titleGap: tg, indent } = o;
  s.box({ x, y, w, h, fill: C.bg, radius: 0.08 });
  const iw = w - 2 * pad;
  const th = heightFor(1, ts, 0);
  s.text(o.title, { x: x + pad, y: y + padY, w: iw, h: th, size: ts, bold: true, color: o.titleColor });
  const top = y + padY + th + tg;
  const avail = y + h - padY - top;
  const lh = heightFor(1, bs, 0);
  const itemH = o.items.map(lines => lh * lines.length);
  const sum = itemH.reduce((a, v) => a + v, 0);
  const n = o.items.length;
  const g = n > 1 ? Math.min(o.maxGap, (avail - sum) / (n - 1)) : 0;
  if (sum + g * (n - 1) > avail + 1e-6 || (n > 1 && g < o.minGap - 1e-6)) {
    throw new Error(`${id} カード「${textOf(o.title)}」: 高さ不足（項目 ${sum.toFixed(2)}in + 間隔 > ${avail.toFixed(2)}in）`);
  }
  let cy = top;
  o.items.forEach((lines, k) => {
    s.text([lines[0]], { x: x + pad, y: r3(cy), w: iw, h: lh, size: bs, color: C.dark });
    if (lines.length > 1) {
      const rest = lines.slice(1);
      s.text(rest, { x: x + pad + indent, y: r3(cy + lh), w: iw - indent, h: lh * rest.length, size: bs, color: C.dark });
    }
    cy += itemH[k] + (k < n - 1 ? g : 0);
  });
}

/**
 * 3列カード。列幅は中身の最長行から決め、余りは狭い列から順に配る（水位合わせ）。
 * cards: [{title, titleColor, items}], 下端 bottom, geo: 寸法（本文サイズなど）
 */
function threeCards(s, id, cards, bottom, geo) {
  const total = BW - 2 * GAP;
  const need = cards.map(c => itemCardNeedW({ ...geo, ...c }));
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
    itemCard(s, id, { ...geo, ...c, x: r3(x), y: TOP, w: r3(widths[i]), h: r3(h) });
    x += widths[i] + GAP;
  });
}

/**
 * 箇条書き（1 項目が 2 行以上になる場合用）。行頭記号と本文を別の枠に置き、折り返し行を本文の頭にそろえる。
 * items: [[行, 行...], ...]（行は文字列 or run 配列）
 */
function bulletBlockHeight(items, size, igap) {
  return items.reduce((a, lines) => a + heightFor(lines.length, size, 0), 0) + igap * (items.length - 1);
}
function bulletBlock(s, items, o) {
  const size = o.size || 24, igap = o.igap != null ? o.igap : 0.3, ind = o.indent || (size >= 24 ? 0.42 : 0.38);
  let y = o.y;
  items.forEach(lines => {
    const h = heightFor(lines.length, size, 0);
    s.text('•', { x: o.x, y: r3(y), w: ind, h: heightFor(1, size, 0), size, color: C.dark });
    s.text(lines, { x: o.x + ind, y: r3(y), w: r3(o.w - ind), h, size, color: C.dark });
    y += h + igap;
  });
  return y - igap - o.y;
}

/** 矢印（ネイティブ freeform）。pts の折れ線を軸に、終点に塗りの三角の矢じり */
function polyArrow(s, pts, o = {}) {
  const color = o.color || C.dark, width = o.width || 3;
  const wIn = width / 72;
  const len = o.head || Math.max(0.2, wIn * 2.6);
  const half = o.half || Math.max(0.11, wIn * 1.5);
  const tip = pts[pts.length - 1], prev = pts[pts.length - 2];
  const dx = tip[0] - prev[0], dy = tip[1] - prev[1], n = Math.hypot(dx, dy);
  const ux = dx / n, uy = dy / n, px = -uy, py = ux;
  const bx = tip[0] - ux * len, by = tip[1] - uy * len;
  const shaftEnd = [tip[0] - ux * len * 0.7, tip[1] - uy * len * 0.7];
  s.polyline([...pts.slice(0, -1), shaftEnd].map(p => [r3(p[0]), r3(p[1])]), { color, width });
  s.polyline([tip, [bx + px * half, by + py * half], [bx - px * half, by - py * half]].map(p => [r3(p[0]), r3(p[1])]),
    { color, width: 0.75, closed: true, fill: color });
}
/** 軸（線＋塗りの矢じり。始点→終点） */
function axis(s, x1, y1, x2, y2) {
  polyArrow(s, [[x1, y1], [x2, y2]], { color: AXIS_C, width: 2, head: 0.16, half: 0.08 });
}
/** 凡例 1 行: 短い線（曲線と同じ色・太さ）＋ラベル（同色・太字） */
function legend(s, o) {
  const yc = o.y + L22 / 2;
  s.polyline([[o.x, r3(yc)], [o.x + o.sw, r3(yc)]], { color: o.color, width: o.lw || 3.5 });
  s.text(o.text, { x: r3(o.x + o.sw + 0.15), y: o.y, w: o.w, h: L22, size: 22, bold: true, color: o.color });
}

/** フローの矢印（水平・垂直のみ）。o.from を渡すとその折れ線を経由して終点へ */
function flowArrow(s, x1, y1, x2, y2, o = {}) {
  const color = o.color || '404040', width = o.width || 3, hl = o.hl || 0.2, hw = o.hw || 0.24;
  const Ln = Math.hypot(x2 - x1, y2 - y1), ux = (x2 - x1) / Ln, uy = (y2 - y1) / Ln;
  if (Math.abs(ux) > 1e-9 && Math.abs(uy) > 1e-9) throw new Error('flowArrow は水平・垂直のみ');
  const bx = x2 - ux * hl, by = y2 - uy * hl;
  if (o.from) s.polyline(o.from.concat([[bx + ux * 0.02, by + uy * 0.02]]).map(p => [r3(p[0]), r3(p[1])]), { color, width });
  else if (uy === 0) s.hline({ x: r3(Math.min(x1, bx)), y: r3(y1), w: r3(Math.abs(bx - x1) + 0.02), color, width });
  else s.vline({ x: r3(x1), y: r3(Math.min(y1, by)), h: r3(Math.abs(by - y1) + 0.02), color, width });
  const px = -uy * hw / 2, py = ux * hw / 2;
  s.polyline([[x2, y2], [bx + px, by + py], [bx - px, by - py]].map(p => [r3(p[0]), r3(p[1])]),
    { color, width: 1, closed: true, fill: color });
}

/** 3列カードの寸法 */
const CARD24 = { pad: 0.14, padY: 0.16, titleSize: 26, bodySize: 24, titleGap: 0.14, indent: 0.3, minGap: 0.1, maxGap: 0.32 };
const CARD22 = { pad: 0.13, padY: 0.15, titleSize: 24, bodySize: 22, titleGap: 0.12, indent: 0.3, minGap: 0.1, maxGap: 0.3 };

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['仕組み', 'MAC', 'セボ・\nデス', '亜酸化\n窒素', '使い方'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '3-3 吸入麻酔薬',
  });

  // ===== S1 表紙 =====
  {
    const sp = spec('S1');
    deck.cover({ title: sp.title, subtitle: sp.subtitle, meta: sp.meta, notes: sp.notes });
  }

  // ===== S2 取り込みと排泄（左: 肺胞濃度／吸入濃度の立ち上がりの模式図 / 右: 箇条書き3行） =====
  {
    const id = 'S2';
    const s = contentSlide(deck, id, 0);
    // --- 図（模式。横軸 時間、縦軸 肺胞濃度 ÷ 吸入濃度 0〜1）---
    const ax0 = 1.3, ay0 = 5.5, axTop = 3.0, axR = 7.15;
    const xs = ax0 + 0.03, xe = 6.9, T = 30;            // 0〜30 分（模式）
    const y1 = 3.3;                                      // 比 = 1 の高さ
    const X = t => xs + (xe - xs) * t / T;
    const Y = v => ay0 - (ay0 - y1) * v;
    const des = t => 1 - 0.17 * Math.exp(-t / 25) - 0.83 * Math.exp(-t / 0.8);
    const sevo = t => 1 - 0.4 * Math.exp(-t / 25) - 0.6 * Math.exp(-t / 1.3);
    const curve = f => {
      const pts = [];
      for (let k = 0; k <= 150; k++) { const t = T * k / 150; pts.push([r3(X(t)), r3(Y(f(t)))]); }
      return pts;
    };
    // 比 = 1 の点線
    s.hline({ x: ax0, y: r3(Y(1)), w: r3(xe - ax0), color: AXIS_C, width: 1.5, dash: 'dash' });
    axis(s, ax0, ay0, ax0, axTop);
    axis(s, ax0, ay0, axR, ay0);
    s.polyline(curve(sevo), { color: C.blue, width: 3.5 });
    s.polyline(curve(des), { color: C.teal, width: 3.5 });
    // 目盛り（縦軸の外）
    const tw = 0.4;
    s.label(fig(id, '1'), { x: r3(ax0 - 0.12 - tw), y: r3(Y(1) - L22 / 2), w: tw, h: L22, align: 'right' });
    s.label(fig(id, '0'), { x: r3(ax0 - 0.12 - tw), y: r3(ay0 - L22 / 2), w: tw, h: L22, align: 'right' });
    // 軸の名前（図の外）
    s.text(fig(id, '肺胞濃度 ÷ 吸入濃度'), { x: X0, y: 1.92, w: 3.5, h: L22, size: 22 });
    s.text(fig(id, '時間'), { x: r3(axR - 1.2), y: r3(ay0 + 0.17), w: 1.2, h: L22, size: 22, align: 'right' });
    // 曲線のラベル（図の外・上）
    const lx = 4.35, sw = 0.55;
    legend(s, { x: lx, y: 1.92, sw, w: 2.3, color: C.teal, text: fig(id, 'デスフルラン') });
    legend(s, { x: lx, y: r3(1.92 + L22 + 0.1), sw, w: 2.3, color: C.blue, text: fig(id, 'セボフルラン') });
    // --- 箇条書き（右） ---
    const b2 = L(id, '肺胞濃度（≈ 呼気濃度）が、脳の濃度の目安', '脳の');
    const b3 = L(id, '血液に溶けにくい薬ほど、立ち上がりと覚醒が速い', '立ち上がり');
    const b4 = L(id, '新鮮ガス流量・換気量・心拍出量で速さが変わる', '心拍出量');
    const items = [
      b2,
      E(b3, 0, '血液に', em('溶けにくい', C.teal), '薬ほど、'),
      b4,
    ];
    const bx = 7.7, igap = 0.42;
    const bh = bulletBlockHeight(items, 22, igap);
    const mid = (1.92 + ay0 + 0.17 + L22) / 2;
    bulletBlock(s, items, { x: bx, y: r3(mid - bh / 2), w: r3(XR - bx), size: 22, igap });
  }

  // ===== S3 MAC：最小肺胞濃度（上: 定義の呼び出し / 下: 表 3列×3行 と箇条書き2行） =====
  {
    const id = 'S3';
    const s = contentSlide(deck, id, 1);
    const def = L(id, 'MAC ＝ 50% の患者が皮膚切開で動かなくなる肺胞濃度')[0];
    const ch = 0.74;
    s.box({ x: X0, y: TOP, w: BW, h: ch, fill: C.white, line: { color: C.teal, width: 2.5 }, radius: 0.08,
      text: [R(def, em('MAC', C.teal), ' ＝ ', em('50%', C.vermilion), ' の患者が皮膚切開で動かなくなる肺胞濃度')],
      size: 28, bold: true, color: C.dark });
    const head = ['薬', 'MAC（成人の目安）', '備考'].map(t => L(id, t)[0]);
    const row = (a, m, c) => [{ text: L(id, a)[0], bold: true, fill: C.bg }, { text: L(id, m)[0], bold: true, align: 'center' }, L(id, c)[0]];
    const rows = [
      head,
      row('セボフルラン', '約 1.7〜2.0%', '年齢で変わる'),
      row('デスフルラン', '約 6%', '年齢で変わる'),
      row('亜酸化窒素', '約 105%', '単独では麻酔にならない'),
    ];
    const ty = r3(TOP + ch + 0.28);
    const t = s.table(rows, { x: X0, y: ty, colW: [3.3, 3.9, BW - 3.3 - 3.9], size: 24, rowH: 0.53, margin: [1, 8, 1, 8] });
    const l1 = L(id, '10 歳上がるごとに約 6〜7% 下がる（高齢者は減量）')[0];
    const l2 = L(id, 'オピオイド・亜酸化窒素の併用で下がる')[0];
    const by = r3(t.y + t.h + 0.28);
    s.bullets([
      R(l1, '10 歳上がるごとに', b('約 6〜7%'), ' 下がる（高齢者は減量）'),
      l2,
    ], { x: X0, y: by, w: 12.2, size: 24, gap: 10 });
  }

  // ===== S4 セボフルラン（3列カード） =====
  {
    const id = 'S4';
    const s = contentSlide(deck, id, 2);
    const cards = [
      {
        title: L(id, '特徴')[0], titleColor: C.blue,
        items: [
          L(id, '気道刺激が少ない'),
          L(id, 'マスク（緩徐）導入に使える', '導入'),
          L(id, '気管支拡張'),
        ],
      },
      {
        title: L(id, '用法（添付文書）')[0], titleColor: C.blue,
        items: [
          E(L(id, '導入：通常 0.5〜5.0%'), 0, '導入：通常 ', b('0.5〜5.0%')),
          E(L(id, '維持：通常 4.0% 以下（最小有効濃度）', '（'), 0, '維持：通常 ', b('4.0% 以下')),
        ],
      },
      {
        title: L(id, '注意')[0], titleColor: C.vermilion,
        items: [
          E(L(id, '悪性高熱症の誘因'), 0, danger('悪性高熱症'), 'の誘因'),
          L(id, '小児の覚醒時興奮'),
          L(id, 'コンパウンド A：低流量で増える', '低流量'),
          E(L(id, '乾いた吸収剤：異常発熱・発火', '異常'), 1, danger('異常発熱・発火')),
        ],
      },
    ];
    threeCards(s, id, cards, BOT2, CARD24);
  }

  // ===== S5 デスフルラン（3列カード） =====
  {
    const id = 'S5';
    const s = contentSlide(deck, id, 2);
    const cards = [
      {
        title: L(id, '特徴')[0], titleColor: C.blue,
        items: [
          E(L(id, '最も溶けにくい → 覚醒が速い', '→'), 1, '→ ', em('覚醒が速い', C.teal)),
          L(id, '長時間・肥満の手術で有利', '手術'),
          L(id, '代謝がほとんどない'),
        ],
      },
      {
        title: L(id, '用法（添付文書）')[0], titleColor: C.blue,
        items: [
          E(E(L(id, '維持：3.0% で開始、通常 7.6% 以下', '通常'), 0, '維持：', b('3.0%'), ' で開始、'), 1, '通常 ', b('7.6% 以下')),
          [R(L(id, '導入には使わない')[0], danger('導入には使わない'))],
        ],
      },
      {
        title: L(id, '注意')[0], titleColor: C.vermilion,
        items: [
          L(id, '気道刺激が強い：咳・喉頭痙攣', '咳'),
          L(id, '急に上げると頻脈・高血圧', '頻脈'),
          L(id, '専用の加温気化器'),
          L(id, '環境負荷が大きい'),
        ],
      },
    ];
    threeCards(s, id, cards, BOT2, CARD24);
  }

  // ===== S6 亜酸化窒素（3列カード） =====
  {
    const id = 'S6';
    const s = contentSlide(deck, id, 3);
    const cards = [
      {
        title: L(id, '特徴')[0], titleColor: C.blue,
        items: [
          L(id, '鎮痛作用がある'),
          E(L(id, 'MAC 約 105%：単独では麻酔にならない', '単独'), 0, b('MAC 約 105%'), '：'),
          L(id, '他の薬の必要量を減らす', '必要量'),
        ],
      },
      {
        title: L(id, '注意')[0], titleColor: C.vermilion,
        items: [
          E(L(id, '閉鎖腔に広がる：気胸・腸閉塞・中耳・眼内ガス', '腸閉塞'), 0, b('閉鎖腔に広がる'), '：気胸・'),
          L(id, 'カフ圧が上がる'),
          L(id, '終了時の拡散性低酸素'),
          L(id, '長時間（1 時間超）で PONV を増やす', 'PONV'),
        ],
      },
      {
        title: L(id, 'その他')[0], titleColor: C.blue,
        items: [
          L(id, 'ビタミン B12 を不活化', '不活化'),
          L(id, '温室効果ガス'),
          L(id, '使用は減っている'),
        ],
      },
    ];
    threeCards(s, id, cards, BOT2, CARD22);
  }

  // ===== S7 副作用と危険（表 2列 × 5行＋見出し行） =====
  {
    const id = 'S7';
    const s = contentSlide(deck, id, 4);
    const head = ['副作用', '対応'].map(t => L(id, t)[0]);
    const c1 = (str, ...parts) => ({ text: [parts.length ? R(L(id, str)[0], ...parts) : L(id, str)[0]], bold: true, fill: C.bg });
    const mh = L(id, '悪性高熱症（全揮発性麻酔薬）')[0];
    const mhCell = { text: [R(mh, { text: '悪性高熱症', color: C.red }, '（全揮発性麻酔薬）')], bold: true, fill: C.bg };
    const mhAct = L(id, '疑ったら即中止、ダントロレン → 6-9', 'ダントロレン');
    const rows = [
      head,
      [mhCell, { text: [R(mhAct[0], { text: '疑ったら即中止', color: C.red, bold: true }, '、'), mhAct[1]] }],
      [c1('血圧低下（血管拡張）'), L(id, '濃度を下げる、昇圧薬、輸液')[0]],
      [c1('呼吸抑制（CO2 への応答が鈍る）'), L(id, '自発呼吸では換気量を見る')[0]],
      [c1('覚醒時興奮（小児・セボフルラン）'), L(id, '鎮痛を十分に、静かな覚醒')[0]],
      [c1('PONV'), L(id, '予防薬、TIVA への変更')[0]],
    ];
    const colW = [6.22, BW - 6.22];
    s.table(rows, { x: X0, y: TOP, colW, size: 24, rowH: [0.66, 1.16, 0.66, 0.66, 0.66, 0.66] });
  }

  // ===== S8 使い方：導入から覚醒まで（5つのチップを矢印でつないだフロー 2段＋箇条書き2行） =====
  {
    const id = 'S8';
    const s = contentSlide(deck, id, 4);
    const size = 24, m = 0.1;
    const steps = [
      L(id, '① 挿管後、高流量で濃度を上げる', '濃度'),
      L(id, '② 呼気濃度が目標に達したら低流量へ', '目標'),
      L(id, '③ 年齢補正 MAC で調節', 'MAC'),
      L(id, '④ 終了前に濃度を下げる', '濃度'),
      L(id, '⑤ 最後は高流量で洗い出す', '洗い出す'),
    ];
    const needW = lines => Math.max(...lines.map(l => measure(textOf(l), size))) + 2 * m + 0.12;
    const ch = 1.16;                      // 2行の箱
    const r1 = TOP + 0.05, r2 = r3(r1 + ch + 0.62);
    // 1段目 ①②③（全幅に配る）。2段目 ④⑤ は ①② と同じ列にそろえる
    const g = 0.62;
    const w1 = steps.slice(0, 3).map(needW);
    const extra = (BW - 2 * g - w1.reduce((a, v) => a + v, 0)) / 3;
    if (extra < 0) throw new Error(`S8: 1段目のチップが収まりません（${extra.toFixed(2)}in）`);
    const ws = w1.map(v => r3(v + extra));
    const xs = [X0, r3(X0 + ws[0] + g), r3(X0 + ws[0] + g + ws[1] + g)];
    if (needW(steps[3]) > ws[0] || needW(steps[4]) > ws[1]) throw new Error('S8: 2段目のチップが1段目の列に収まりません');
    const chip = (lines, x, y, w, color) => s.box({ x, y, w, h: ch, fill: C.bg, line: { color, width: 2.25 }, radius: 0.12,
      text: lines, size, bold: true, color: C.dark, margin: m });
    const ac = 0.08;
    chip(steps[0], xs[0], r1, ws[0], C.blue);
    flowArrow(s, xs[0] + ws[0] + ac, r1 + ch / 2, xs[1] - ac, r1 + ch / 2);
    chip(steps[1], xs[1], r1, ws[1], C.blue);
    flowArrow(s, xs[1] + ws[1] + ac, r1 + ch / 2, xs[2] - ac, r1 + ch / 2);
    chip(steps[2], xs[2], r1, ws[2], C.teal);
    chip(steps[3], xs[0], r2, ws[0], C.blue);
    flowArrow(s, xs[0] + ws[0] + ac, r2 + ch / 2, xs[1] - ac, r2 + ch / 2);
    chip(steps[4], xs[1], r2, ws[1], C.blue);
    // ③ → ④ の折り返し（下 → 左 → 下）
    const ym = r3((r1 + ch + r2) / 2), c3 = r3(xs[2] + ws[2] / 2), c4 = r3(xs[0] + ws[0] / 2);
    flowArrow(s, c4, ym, c4, r3(r2 - ac), { from: [[c3, r3(r1 + ch + ac)], [c3, ym], [c4, ym]] });
    // 箇条書き
    const a1 = L(id, '筋弛緩薬の使用中は、呼気濃度の下限アラーム（目安 0.7 MAC）')[0];
    const a2 = L(id, '悪性高熱症を疑ったら、即座に中止して高流量')[0];
    const by = r3(r2 + ch + 0.34);
    s.bullets([
      R(a1, '筋弛緩薬の使用中は、呼気濃度の下限アラーム（目安 ', b('0.7 MAC'), '）'),
      R(a2, '悪性高熱症を疑ったら、', danger('即座に中止'), 'して高流量'),
    ], { x: X0, y: by, w: 12.2, size: 24, gap: 10 });
    if (by + heightFor(2, 24, 10) > BOT2 + 0.2) throw new Error('S8: 箇条書きが出典に近すぎます');
  }

  // ===== S9 まとめ（箇条書き5行） =====
  {
    const id = 'S9';
    const s = contentSlide(deck, id, 4);
    const items = [
      L(id, '呼気濃度で管理する。溶けにくい薬ほど速い')[0],
      R(L(id, 'MAC：セボフルラン約 1.7〜2%、デスフルラン約 6%、年齢で下がる')[0],
        'MAC：セボフルラン', b('約 1.7〜2%'), '、デスフルラン', b('約 6%'), '、年齢で下がる'),
      L(id, 'セボフルランはマスク導入に、デスフルランは覚醒の速さに')[0],
      L(id, '亜酸化窒素は閉鎖腔に広がる。終了時は 100% 酸素')[0],
      R(L(id, '悪性高熱症を疑ったら、即座に中止して高流量')[0], '悪性高熱症を疑ったら、', danger('即座に中止'), 'して高流量'),
    ];
    const gapPt = 28, h = heightFor(5, 24, gapPt);
    const y = r3(Math.min(BOT2 - h, TOP + (BOT2 + 0.2 - TOP - h) / 2));   // 推定高さは実際より大きいので少し下げる
    s.bullets(items, { x: X0, y, w: 12.2, size: 24, gap: gapPt });
  }

  checkOnSlide();
  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.render && res.render.images) console.log('[build] render: ' + (res.render.images[0] ? path.dirname(res.render.images[0]) : '(none)'));
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
  return res;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
