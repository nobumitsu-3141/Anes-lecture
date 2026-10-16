'use strict';
/**
 * 4-3 術前絶飲食と誤嚥リスク
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。q()/brk() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 図はすべてネイティブ図形（freeform）。図のラベルは図の外に置き、引き出し線でつなぐ。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, LAYOUT, fits, measure } = require('../_lib/decklib');

const NAME = '4-3_fasting_aspiration_risk';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（タイトル・章・ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
const SPEC_MD = fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8');
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const lines = part.split('\n');
    const head = lines[0];
    const id = (head.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const title = (head.match(/title=「([^」]*)」/) || [])[1];
    const chm = head.match(/\bch=(\d+)/);
    let notes = null, sources = null, figure = '';
    const fields = {};
    const quoted = new Set();
    let inOn = false;
    for (let i = 1; i < lines.length; i++) {
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
      const f = ln.match(/^(subtitle|meta):\s*(.*)$/);
      if (f) fields[f[1]] = f[2].trim();
      if (/^figure:/.test(ln)) figure = ln;
      if (/^on-slide:/.test(ln)) { inOn = true; continue; }
      if (inOn && !/^ {2}- /.test(ln)) inOn = false;
      // on-slide の「…」（入れ子の『』はそのまま）と、layout / figure 行の「…」（表の列名・図の注記）
      if (inOn || /^(layout|figure):/.test(ln)) for (const qm of ln.matchAll(/「([^」]*)」/g)) quoted.add(qm[1]);
    }
    out[id] = {
      title, ch: chm ? Number(chm[1]) : null, notes, figure, fields, quoted,
      sources: sources && sources !== '—' ? sources : null,
    };
  }
  return out;
}
const SPEC = parseSpec(SPEC_MD);
function notesOf(id) {
  const n = SPEC[id] && SPEC[id].notes;
  if (!n) throw new Error(`spec.md の ${id} に notes がありません`);
  return n;
}
/** spec の文言そのまま（照合のみ） */
function q(id, str) {
  if (!SPEC[id].quoted.has(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
  return str;
}
/** 図のラベル: spec の figure 行に含まれる語だけを使う */
function figWord(id, str) {
  if (!SPEC[id].figure.includes(str)) throw new Error(`${id}: 図のラベル「${str}」が spec の figure 行にありません`);
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
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める
//  改行は引用と引用の区切りでだけ入れる（1 つの引用を 2 行にまたがらせない）。
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 入らなければ区切りの前後の空白を詰める（全角の閉じ括弧の後ろだけ '/ '、次に全部 '/'）。
//  2) それでも入らなければ（）内の補足を後ろの文献から外す（' / '、'/' の順に試す）。
//  3) 以上で 2 行に入らないとき（引用 3 件がどれも長い）だけ、文言はそのままで引用の途中の語の切れ目で改行する
//     （著者「et al.」の後 > 誌名と年の間 > その他の空白）。文献そのものは落とさない。
//  元と違う形にしたものは DEVIATIONS に記録してビルド時に表示する。
// ---------------------------------------------------------------------------
const SRC = LAYOUT.source;
const DEVIATIONS = [];
/** スライド面の文言を spec から変えるとき（記録つき）。orig は spec の on-slide にあること */
function dev(id, orig, repl, why) {
  q(id, orig);
  DEVIATIONS.push(`${id} 本文: 「${orig}」→「${repl}」（${why}）`);
  return repl;
}
function libLines(items) {
  const lines = [];
  for (const it of items) {
    const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
    if (cur && fits(cur, SRC.size, SRC.w)) lines[lines.length - 1] = cur;
    else lines.push(it);
  }
  return lines;
}
function split2(str, sep) {
  const fitsLine = l => fits(l, SRC.size, SRC.w);
  const cands = [];
  for (let i = 1; i < str.length; i++) {
    let a, b, rank;
    if (str.startsWith(sep, i)) { a = str.slice(0, i); b = str.slice(i + sep.length); rank = 0; }
    else if (str[i] === ' ') {
      a = str.slice(0, i); b = str.slice(i + 1);
      if (a.endsWith('.')) rank = 1;                   // 著者名（et al. / 頭文字.）の後
      else if (/^\d/.test(b)) rank = 2;                // 誌名と年の間
      else rank = 3;
      if (a.endsWith('/') || b.startsWith('/')) continue;
    } else continue;
    a = a.trim(); b = b.trim();
    if (a && b && fitsLine(a) && fitsLine(b)) cands.push({ a, b, rank });
  }
  if (!cands.length) return null;
  cands.sort((x, y) => x.rank - y.rank || y.a.length - x.a.length);
  return [cands[0].a, cands[0].b];
}
function sourcesFor(id) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  const note = (form, lines) => {
    if (form !== spec) DEVIATIONS.push(`${id} 出典: ${spec} → ${lines.join(' ⏎ ')}`);
    return lines;
  };
  const fitsLine = l => fits(l, SRC.size, SRC.w);
  const items = spec.split(' / ');
  if (items.every(fitsLine) && libLines(items).length <= 2) return items;
  // 区切りの詰め方: そのまま ' / ' → 全角の閉じ括弧の後ろだけ詰める（字面の右に空きがあるので '」/ ' でも空いて見える）→ '/'
  const JOIN = [
    its => its.join(' / '),
    its => its.reduce((acc, it) => acc + (/[」）』】]$/.test(acc) ? '/ ' : ' / ') + it),
    its => its.join('/'),
  ];
  // 区切りの位置でだけ 2 行に割る（1 行目をできるだけ長く）
  const atSep = (its, join) => {
    for (let k = its.length - 1; k >= 1; k--) {
      const a = join(its.slice(0, k)), b = join(its.slice(k));
      if (fitsLine(a) && fitsLine(b)) return [a, b];
    }
    return null;
  };
  let r;
  for (const join of JOIN.slice(1)) {
    r = atSep(items, join);
    if (r) return note(join(items), r);
  }
  const cut = items.slice();
  for (let k = cut.length - 1; k >= 0; k--) {
    if (!/（[^）]*）$/.test(cut[k])) continue;
    cut[k] = cut[k].replace(/（[^）]*）$/, '');
    for (const join of JOIN) {
      r = atSep(cut, join);
      if (r) return note(join(cut), r);
    }
  }
  // 最後の手段: 文言はそのままで、引用の途中の語の切れ目で改行
  r = split2(spec, ' / ');
  if (r) {
    DEVIATIONS.push(`${id} 出典: 区切りの位置では 2 行に入らない（空白を詰め、（）の補足を外しても入らない）ため、引用の途中で改行: ${r.join(' ⏎ ')}`);
    return r;
  }
  throw new Error(`${id}: 出典が 2 行に収まりません`);
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const TOP = 1.85;                // 本文の上端
const TOP_MIN = 1.82;            // 金下線の下端（約 1.58in）から 0.24in
const BOT = 6.44;                // 出典 2 行（枠 y6.5・下揃え）のときの本文下端（全スライド出典 2 行）
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const r2 = v => Math.round(v * 100) / 100;
const r3 = v => Math.round(v * 1000) / 1000;

/**
 * 1 行 1 枠のテキスト（行送り em を指定）。各行は decklib の s.text（上揃え）で、
 * 枠の高さは decklib / lint の 1 行の下限（22pt 0.41in、24pt 0.45in、26pt 0.5in）。
 * 行送りが枠より狭いときは次の行の枠と少し重なる（文字どうしの枠の重なりは lint・decklib とも対象外）。
 */
const boxH = size => Math.ceil(Math.max(size / 72 * 1.2 * 1.18 + 6 / 72 - 0.3 / 2.54, size / 72 * 1.5 - 0.05) * 100 - 1e-9) / 100;
const pitch = (size, em) => size * em / 72;
/**
 * 1 行の枠の中での字面（インクの範囲）の上端・下端（em）。LibreOffice の描画で実測（22pt: +0.045 / +0.347in、
 * 24pt: +0.063 / +0.383in）。箱の上下の余白を「箱の縁から字面まで」で決めるのに使う。
 */
const GLYPH = { top: 0.15, bot: 1.14 };
/** 行頭の始め括弧（字面が右半分にある）。左揃えの行では半角ぶん左へぶら下げ、本文の開始位置をそろえる */
const OPEN_BRACKET = /^[『「（〔【《〈]/;
function lines(s, ls, o) {
  const size = o.size || 22, em = o.em || 1.2, p = pitch(size, em);
  const align = o.align || 'left';
  ls.forEach((line, i) => {
    const hang = align === 'left' && OPEN_BRACKET.test(runText(line)) ? 0.5 * size / 72 : 0;
    s.text([line], {
      x: r3(o.x - hang), y: r3(o.y + i * p), w: r3(o.w + hang), h: boxH(size),
      size, bold: !!o.bold, color: o.color || C.dark, align,
    });
  });
  return { top: o.y, bottom: r3(o.y + (ls.length - 1) * p + boxH(size)), h: ls.length * p };
}

/**
 * 箇条書きの列（行頭記号「•」を別の枠に置き、記号は常に本文の黒）。
 * items: 各項目 = 行の配列（2 行目以降は記号の右にぶら下げる）。
 * y 上端、size、em（行送り）、gap（項目間 in）、indent（記号ぶん in）
 */
function bulletList(s, items, o) {
  const size = o.size || 24, em = o.em || 1.5, ind = o.indent || 0.4, gap = o.gap != null ? o.gap : 0.3;
  const p = pitch(size, em);
  let y = o.y;
  items.forEach((ls, i) => {
    lines(s, ['•'], { x: o.x, y, w: Math.max(ind, measure('•', size) + 0.01), size, em });
    lines(s, ls, { x: r3(o.x + ind), y, w: r3(o.w - ind), size, em });
    y += ls.length * p + (i < items.length - 1 ? gap : 0);
  });
  const last = items[items.length - 1].length;
  return { top: o.y, bottom: r3(y - p + boxH(size)), textBottom: r3(y), lastBoxOver: boxH(size) - p, n: last };
}

/**
 * 強調ボックス（白地＋色枠、太字・中央）。1 行は decklib の box（上下中央）、
 * 2 行は行送りを指定して lines() で置く。戻り値 {y, h, bottom}
 */
const EMPH = { radius: 0.1, lineW: 3.5 };
function emphBox(s, o) {
  const x = o.x != null ? o.x : X0, w = o.w != null ? o.w : BW;
  const size = o.size || 26;
  if (o.lines.length === 1) {
    const h = o.h || r2(boxH(size) + 0.2);
    const y = r3(o.bottom != null ? o.bottom - h : o.y);
    s.box({
      x, y, w, h, fill: C.white, radius: EMPH.radius, line: { color: o.lineColor, width: EMPH.lineW },
      margin: 0.08, text: o.lines, size, bold: true, color: C.dark,
    });
    return { y, h, bottom: r3(y + h) };
  }
  const em = o.em || 1.2, p = pitch(size, em), pad = o.pad != null ? o.pad : 0.06;
  // h を渡したとき: 字面（GLYPH）の塊を箱の上下中央に置く（上下の余白が等しくなる）
  const ink = (o.lines.length - 1) * p + (GLYPH.bot - GLYPH.top) * size / 72;
  const h = o.h != null ? r3(o.h) : r3(2 * pad + o.lines.length * p);
  const y = r3(o.bottom != null ? o.bottom - h : o.y);
  const ty = o.h != null ? y + (h - ink) / 2 - GLYPH.top * size / 72 : y + pad;
  s.box({ x, y, w, h, fill: C.white, radius: EMPH.radius, line: { color: o.lineColor, width: EMPH.lineW } });
  const t = lines(s, o.lines, { x: x + 0.15, y: ty, w: w - 0.3, size, em, bold: true, align: 'center' });
  if (t.bottom > y + h + 0.06) throw new Error(`p${s.page} 強調ボックス: 最終行の枠が ${r3(t.bottom - y - h)}in はみ出します`);
  return { y, h, bottom: r3(y + h), clear: r3((h - ink) / 2) };
}

// ---------------------------------------------------------------------------
// 見出し＋項目のカード（角丸の塗り箱。文字はすべて箱の内側）
//  - 項目は「•」付き。1 項目を語句の切れ目で 2 行に割ったときは 2 行目を記号の右にぶら下げる。
//  - 左右の内側余白 padX は 0.2in（足りなければ最長行が入る範囲で 0.1in まで詰める）。
//  - 幅: 各カードの最長行（ライブラリ推定）＋記号ぶん＋左右余白を確保し、余りを等分する。equal: true なら全カード同幅。
//  - 高さ: y〜bottom。項目の間隔は全カード共通で、詰まった方のカードに合わせる（上限 maxGap）。
// ---------------------------------------------------------------------------
const CARD = { padX: 0.2, minPadX: 0.1, padTop: 0.08, padBot: 0.1, titleSize: 24, size: 22, titleGap: 0.06, ind: 0.26 };
const titleP = st => pitch(st.titleSize, st.titleEm || st.em);
function cardNeedW(c, st) {
  const t = measure(runText(c.title), st.titleSize);
  const b = Math.max(...c.items.flat().map(l => measure(runText(l), st.size))) + st.ind;
  return Math.max(t, b) + 2 * st.padX;
}
function cardFixedH(c, st) {
  const n = c.items.reduce((a, it) => a + it.length, 0);
  return st.padTop + titleP(st) + st.titleGap + n * pitch(st.size, st.em) + st.padBot;
}
function drawCard(s, c, x, y, w, h, gap, st) {
  s.box({ x: r3(x), y: r3(y), w: r3(w), h: r3(h), fill: C.bg, radius: 0.08 });
  const ix = x + st.padX, iw = w - 2 * st.padX;
  lines(s, [c.title], { x: ix, y: y + st.padTop, w: iw, size: st.titleSize, em: st.titleEm || st.em, bold: true, color: c.titleColor });
  const bl = bulletList(s, c.items, {
    x: ix, y: y + st.padTop + titleP(st) + st.titleGap, w: iw,
    size: st.size, em: st.em, indent: st.ind, gap,
  });
  if (bl.textBottom > y + h - st.padBot + 1e-6) throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さ不足（${r3(bl.textBottom)} > ${r3(y + h - st.padBot)}）`);
  if (bl.bottom > y + h + 0.06) throw new Error(`p${s.page} カード「${runText(c.title)}」: 最終行の枠が ${r3(bl.bottom - y - h)}in はみ出します`);
  return bl;
}
function cardRow(s, cards, o) {
  const st = { ...CARD, ...(o.style || {}) };
  const avail = BW - GAP * (cards.length - 1);
  // 内側余白: 既定 padX で入らなければ、入る範囲で minPadX まで詰める
  const needAt = px => cards.map(c => cardNeedW(c, { ...st, padX: px }));
  const sumAt = px => (o.equal ? Math.max(...needAt(px)) * cards.length : needAt(px).reduce((a, b) => a + b, 0));
  if (sumAt(st.padX) > avail + 1e-9) {
    const px = st.padX - (sumAt(st.padX) - avail) / (2 * cards.length);
    if (px < st.minPadX - 1e-9) throw new Error(`p${s.page}: カード幅が足りません（内側余白 ${r3(px)}in < ${st.minPadX}in）`);
    st.padX = Math.floor((px - 0.003) * 1000) / 1000;   // 丸めの余裕
  }
  const need = needAt(st.padX);
  const extra = avail - need.reduce((a, b) => a + b, 0);
  const ws = o.equal
    ? cards.map(() => Math.floor(avail / cards.length * 1000) / 1000)
    : need.map(v => Math.floor((v + extra / cards.length) * 1000) / 1000);
  ws.forEach((w, i) => { if (w + 1e-3 < need[i]) throw new Error(`p${s.page}: カード${i + 1} の幅 ${w}in < 必要 ${r3(need[i])}in`); });
  const h = r3(o.bottom - o.y);
  let gap = Math.min(o.maxGap, ...cards.map(c => (c.items.length > 1 ? (h - cardFixedH(c, st)) / (c.items.length - 1) : Infinity)));
  gap = Math.floor(gap * 1000) / 1000;
  if (gap < -1e-9) throw new Error(`p${s.page}: カードの高さが ${r3(-gap * (cards[0].items.length - 1))}in ほど足りません（h=${h}）`);
  gap = Math.max(0, gap);
  let x = X0;
  const bls = [];
  cards.forEach((c, i) => { bls.push(drawCard(s, c, x, o.y, ws[i], h, gap, st)); x += ws[i] + GAP; });
  // 最終行の字面の下端からカード下端まで（LibreOffice 実測の GLYPH で推定）
  const clearBot = r3(Math.min(...bls.map(b => o.y + h - (b.textBottom - pitch(st.size, st.em) + GLYPH.bot * st.size / 72))));
  console.log(`  p${s.page} カード: 幅 ${ws.join(' / ')}in（必要 ${need.map(r3).join(' / ')}）、内側余白 ${st.padX}in、` +
    `${o.y}〜${r3(o.y + h)}in、行送り ${st.em}em、項目間 ${gap}in、最終行の字面の下 ${clearBot}in`);
  return { gap, h, clearBot };
}

// ---------------------------------------------------------------------------
// 胃エコーの模式図（S6）。心窩部の矢状断（コンベックスの扇形）。
//  上から皮膚、肝左葉、前庭部（楕円、壁を描く）、膵、大動脈（円）。前庭部の外縁に「断面積をトレース」の破線。
//  扇形: 仮想の頂点 (cx, ya) から半径 r1（皮膚）〜 r2（深部）、半角 th。
// ---------------------------------------------------------------------------
const US = {
  cx: 4.3, yTop: 2.02, r1: 3.5, r2: 6.68, th: 12 * Math.PI / 180,
  bg: 'EDEDED', edge: 'A6A6A6', skin: '7F7F7F', skinW: 0.1,
  liverFill: 'F7D9C4', antrumFill: 'C9DEF0', lumen: 'FFFFFF', panFill: 'D0D0D0', panLine: '7F7F7F',
  aortaFill: '404040',
};
function polar(r, phi) { const ya = US.yTop - US.r1; return [r3(US.cx + r * Math.sin(phi)), r3(ya + r * Math.cos(phi))]; }
function arcPts(r, a0, a1, n) { const out = []; for (let k = 0; k <= n; k++) out.push(polar(r, a0 + (a1 - a0) * k / n)); return out; }
function ellPts(cx, cy, rx, ry, n = 72) {
  const out = [];
  for (let k = 0; k < n; k++) { const t = 2 * Math.PI * k / n; out.push([r3(cx + rx * Math.cos(t)), r3(cy + ry * Math.sin(t))]); }
  return out;
}
/** 扇形の左右の縁の x（高さ y） */
function edgeX(y, side) {
  const [xt, yt] = polar(US.r1, side * US.th), [xb, yb] = polar(US.r2, side * US.th);
  return xt + (xb - xt) * (y - yt) / (yb - yt);
}
function usFigure(s) {
  const th = US.th;
  // 扇形（超音波画像の範囲）
  s.polyline([...arcPts(US.r1, -th, th, 24), ...arcPts(US.r2, th, -th, 48)], { color: US.edge, width: 1.25, closed: true, fill: US.bg });
  // 皮膚（扇形の上縁の帯）
  s.polyline([...arcPts(US.r1, -th, th, 24), ...arcPts(US.r1 + US.skinW, th, -th, 24)], { color: US.skin, width: 0.75, closed: true, fill: US.skin });
  // 肝左葉（上の左＝頭側。尾側へ細くなる）
  const yL0 = 2.2, yL1 = 3.12;
  const liver = [
    [r3(edgeX(yL0, -1)), yL0], ...arcPts(US.r1 + 0.16, -th + 0.012, 0.06, 10).slice(1),
    [4.78, 2.25], [5.02, 2.4], [4.62, 2.68], [4.1, 2.92], [3.6, 3.07], [r3(edgeX(yL1, -1)), yL1],
  ];
  s.polyline(liver, { color: C.vermilion, width: 1.5, closed: true, fill: US.liverFill });
  // 前庭部（外側の楕円＝壁、内側＝内腔）と、外縁（漿膜を含む）をなぞる破線
  const A = { cx: 4.45, cy: 3.7, rx: 0.62, ry: 0.42, wall: 0.15 };
  s.polyline(ellPts(A.cx, A.cy, A.rx, A.ry), { color: C.blue, width: 2.25, closed: true, fill: US.antrumFill });
  s.polyline(ellPts(A.cx, A.cy, A.rx - A.wall, A.ry - A.wall), { color: C.blue, width: 1, closed: true, fill: US.lumen });
  s.polyline(ellPts(A.cx, A.cy, A.rx + 0.08, A.ry + 0.08), { color: C.teal, width: 2.5, closed: true, dash: 'dash' });
  // 膵（前庭部の後ろ）
  const P = { cx: 4.5, cy: 4.5, rx: 0.78, ry: 0.16 };
  s.polyline(ellPts(P.cx, P.cy, P.rx, P.ry), { color: US.panLine, width: 1.5, closed: true, fill: US.panFill });
  // 大動脈または上腸間膜動脈（さらに後ろの円）
  const V = { cx: 4.3, cy: 4.9, r: 0.17 };
  s.polyline(ellPts(V.cx, V.cy, V.r, V.r, 48), { color: C.dark, width: 1.5, closed: true, fill: US.aortaFill });
  return { A, P, V, bottom: polar(US.r2, 0)[1], left: edgeX(polar(US.r2, 0)[1], -1) };
}

async function main() {
  const CHAPTERS = ['絶飲食', 'リスク', 'GLP-1', '評価と\n対応', 'まとめ'];
  const chipSpec = '「絶飲食」「リスク」「GLP-1」「評価と\\n対応」「まとめ」';
  if (!SPEC_MD.includes(chipSpec)) throw new Error('章チップのラベルが spec と違います');
  const deck = new Deck({
    file: OUT,
    chapters: CHAPTERS,
    author: 'Anes-lecture',
    deckTitle: '4-3 術前絶飲食と誤嚥リスク',
  });
  const slide = id => deck.slide({ title: SPEC[id].title, chapter: SPEC[id].ch, notes: notesOf(id), sources: sourcesFor(id) });
  let S4_GLP_PAGE = null;   // S4 の「→ n ページ」が指すページ（S5 で照合）

  // ===== S1 表紙 =====
  {
    const S = SPEC.S1;
    if (S.fields.subtitle !== '4-3 術前評価' || S.fields.meta !== '初期研修医向け麻酔講義') throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: S.title, subtitle: S.fields.subtitle, meta: S.fields.meta, notes: notesOf('S1') });
  }

  // ===== S2 絶飲食の目安（表 3列 × 5行＋見出し行） =====
  {
    const id = 'S2';
    const s = slide(id);
    const ctr = (text, extra = {}) => ({ text: q(id, text), align: 'center', ...extra });
    const key = text => ctr(text, { color: C.blue, bold: true });
    // 列幅は spec のとおり 5.0 / 3.4 / 3.8in（合計 12.2in）。セルの左右余白を 6pt にして（decklib 既定 8pt）、
    // 「軽い食事（トースト、おかゆ）」（推定 4.791in）を 1 列目（使える幅 4.833in）に 1 行で入れる。
    // 1 列目の清澄水の例示は語句の切れ目で 3 行にする。
    const t = s.table([
      [q(id, '内容'), q(id, 'JSA 2012'), q(id, 'ASA 2017・2023')],
      [brk(id, '清澄水（水、お茶、果肉のないジュース、経口補水液）', '果肉', '経口').join('\n'), key('2 時間前まで'), key('2 時間前まで')],
      [q(id, '母乳'), ctr('4 時間'), ctr('4 時間')],
      [q(id, '人工乳・牛乳'), ctr('人工乳 6 時間'), ctr('6 時間')],
      [q(id, '軽い食事（トースト、おかゆ）'), ctr('推奨時間なし'), ctr('6 時間')],
      [q(id, '脂肪の多い食事、肉'), ctr('推奨時間なし'), ctr('8 時間以上')],
    ], { x: X0, y: 1.95, colW: [5.0, 3.4, 3.8], margin: [4, 6, 4, 6] });
    console.log(`  p${s.page} 表: ${t.y}〜${r3(t.y + t.h)}in（行高 ${t.rowH.join(', ')}）`);
    if (t.y + t.h > BOT + 1e-6) throw new Error(`${id} 表の下端 ${r3(t.y + t.h)} > ${BOT}`);
  }

  // ===== S3 指示の出し方（箇条書き5行＋強調ボックス） =====
  {
    const id = 'S3';
    const s = slide(id);
    const bl = bulletList(s, [
      [mark(q(id, '『0 時から絶飲食』は古い。清澄水は 2 時間前まで勧める'), ['古い', C.vermilion], ['2 時間前まで勧める', C.blue])],
      [mark(q(id, '固形物は『前日 21 時まで』など、時刻で書く'), ['時刻で書く', C.blue])],
      [q(id, '炭水化物飲料（ERAS）は 2 時間前までに飲み切る')],
      [q(id, '内服は少量の水で。飲む薬・止める薬を明記')],
      [mark(q(id, '手術が遅れたら、水を飲んでよい時刻を更新する'), ['時刻を更新する', C.blue])],
    ], { x: X0, y: TOP + 0.05, w: 12.2, size: 26, em: 1.5, indent: 0.45, gap: 0.2 });
    const e = emphBox(s, {
      bottom: BOT, size: 26, lineColor: C.teal,
      lines: [mark(q(id, '長い絶飲食は脱水・低血糖・不機嫌を招き、安全にはならない'), ['安全にはならない', C.vermilion])],
    });
    if (e.y < bl.bottom + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔 ${r3(e.y - bl.bottom)}in`);
    console.log(`  p${s.page} 箇条書き ${bl.top}〜${bl.bottom}in / 強調 ${e.y}〜${e.bottom}in`);
  }

  // ===== S4 胃が空でない患者（2列カード＋強調ボックス） =====
  {
    const id = 'S4';
    const s = slide(id);
    // 強調ボックスは他のスライドと同じ青緑の枠。強調語だけバーミリオン。
    const e = emphBox(s, {
      bottom: BOT, size: 26, lineColor: C.teal,
      lines: [mark(q(id, '『絶飲食の時間を守った』は『胃が空』を保証しない'), ['保証しない', C.vermilion])],
    });
    // spec の「→ S5」は spec 上のスライド番号。画面のページ番号（表紙を除く通し番号）で GLP-1 のスライドを指す。
    const glpPage = s.page + 1;
    const glp = dev(id, 'GLP-1 受容体作動薬（→ S5）', `GLP-1 受容体作動薬（→ ${glpPage} ページ）`,
      'S5 は spec のスライド番号で、画面のページ番号では GLP-1 のスライドは ' + glpPage);
    cardRow(s, [
      {
        title: q(id, '状況'), titleColor: C.vermilion,
        items: [
          [q(id, '緊急手術、絶飲食が守れていない')],
          brk(id, '外傷・強い痛み・オピオイド（胃排出が遅れ、予測できない）', '（胃排出'),
          [q(id, 'イレウス、腸閉塞、腹膜炎、腹水')],
          [q(id, '妊娠（陣痛中、緊急帝王切開）')],
          [q(id, '意識障害、嚥下障害')],
        ],
      },
      {
        title: q(id, '病気・薬'), titleColor: C.blue,
        items: [
          // 左右のカードの内側余白を 0.2in 近く取るため、左のカードと同じく（）の前で 2 行にする
          brk(id, '糖尿病（胃不全麻痺、自律神経障害）', '（胃不全'),
          [q(id, '症状のある逆流・食道裂孔ヘルニア')],
          [q(id, '食道アカラシア、胃の手術歴')],
          [glp],
          [q(id, '高度肥満（単独では議論あり）')],
        ],
      },
    ], { y: TOP, bottom: r3(e.y - GAP), maxGap: 0.2, style: { em: 1.3 } });
    S4_GLP_PAGE = glpPage;
  }

  // ===== S5 GLP-1 受容体作動薬（2列カード＋強調ボックス） =====
  {
    const id = 'S5';
    const s = slide(id);
    if (s.page !== S4_GLP_PAGE) throw new Error(`S4 の「→ ${S4_GLP_PAGE} ページ」が GLP-1 のスライド（p${s.page}）と合いません`);
    // 各カード 6 項目・9 行＋見出し、強調ボックス 2 行。どちらのカードも最長の語句で 9 行より減らせない
    // （左右の幅をどう配分しても片方の 1 行項目が折り返す）ので、文言は変えずに縦の配分で収める。
    //  - カード: 最終行の字面の下端からカード下端まで 0.15in 以上（GLYPH で推定）、見出しの字面の上 0.14in。
    //  - 強調ボックス: カードの GAP 下〜BOT5。2 行の字面の塊を上下中央に置く。
    //  - BOT5: 出典 2 行の枠（y6.5、文字は下揃えで上端 約6.82in）の 0.035in 上（decklib の警告の境 6.47in の内側）。
    const BOT5 = 6.465;
    const em = 1.06, emE = 1.1, p = pitch(22, em);
    const st = { em, titleEm: 1.08, padTop: 0.08, titleGap: 0.02, padBot: r3(0.15 + GLYPH.bot * 22 / 72 - p + 0.012) };
    const cards = [
      {
        title: q(id, '何が問題か'), titleColor: C.vermilion,
        items: [
          brk(id, '胃の排出を遅らせる。絶飲食を守っても残る', '絶飲食'),
          brk(id, '胃エコーで残胃 56% 対 19%（週 1 回製剤）', '（週').map((l, i) => (i === 0 ? mark(l, ['56% 対 19%', C.vermilion]) : l)),
          [q(id, '誤嚥の増加はメタ解析で示されず')],
          [q(id, '休薬の効果は報告で一致しない')],
          [q(id, '糖尿病だけでなく肥満症でも処方')],
          brk(id, 'セマグルチド添付文書（2026）：全身麻酔・深い鎮静下に注意', '全身麻酔'),
        ],
      },
      {
        title: q(id, 'どう扱うか（資料で違う）'), titleColor: C.blue,
        items: [
          brk(id, 'ASA 2023：週 1 回は 1 週前、1 日 1 回は当日休薬を検討', '1 日 1 回'),
          [q(id, '多学会 2024・英・豪：原則続ける')],
          brk(id, '24 時間の流動食（豪州・韓国：使用者全般、多学会：懸念例）', '使用者全般'),
          [q(id, '症状→延期を検討、緊急は満腹扱い')],
          [q(id, '無症状でも安全とは限らない')],
          brk(id, '迷ったら胃エコー、満腹なら RSI か延期', '満腹なら'),
        ],
      },
    ];
    const cardH = Math.max(...cards.map(c => cardFixedH(c, { ...CARD, ...st })));
    const cardBot = Math.ceil((TOP_MIN + cardH) * 1000) / 1000;
    const ey = r3(cardBot + GAP);
    const e = emphBox(s, {
      y: ey, h: r3(BOT5 - ey), size: 22, em: emE, lineColor: C.teal,
      lines: brk(id, '問診で薬の名前・用量・最後の投与日を聞く。『やせ薬』『週 1 回の注射』と具体的に', '『やせ薬』')
        .map((l, i) => (i === 0 ? mark(l, ['名前・用量・最後の投与日', C.blue]) : l)),
    });
    const r = cardRow(s, cards, { y: TOP_MIN, bottom: cardBot, maxGap: 0, style: st });
    console.log(`  p${s.page} 強調 ${e.y}〜${e.bottom}in（高さ ${e.h}in、字面の上下 ${e.clear}in）`);
    if (r.clearBot < 0.15 - 1e-3) throw new Error(`p${s.page}: カード最終行の字面の下が ${r.clearBot}in（0.15in 未満）`);
  }

  // ===== S6 胃エコーで胃内容を見る（左に模式図、右に箇条書き4項目、下に式） =====
  {
    const id = 'S6';
    const s = slide(id);
    // 式（下）: 22pt でも 1 行（推定 12.68in）が本文幅 12.23in に入らないので 2 行にする。
    // 左辺「胃内容量（mL）」を 1 行目、2 行目を「≈」で始めて右辺を続ける（2 行目は左辺から 0.5in 字下げ）。
    const fsz = 24, fem = 1.2, fpad = 0.07, fx = X0 + 0.25, find = 0.5;
    const fl = brk(id, '胃内容量（mL）≈ 27.0 ＋ 14.6 × 右側臥位の前庭部断面積（cm²）− 1.28 × 年齢', '≈');
    if (!fits(fl[1], fsz, XR - 0.1 - (fx + find))) throw new Error(`${id}: 式の 2 行目が 1 行に入りません`);
    const fh = r3(2 * fpad + pitch(fsz, fem) + boxH(fsz));
    const fy = r3(BOT - fh);
    s.box({ x: X0, y: fy, w: BW, h: fh, fill: C.bg, radius: 0.08 });
    lines(s, [mark(fl[0], ['胃内容量（mL）', C.blue])], { x: fx, y: fy + fpad, w: 4.0, size: fsz, em: fem });
    lines(s, [fl[1]], { x: fx + find, y: r3(fy + fpad + pitch(fsz, fem)), w: r3(XR - 0.1 - (fx + find)), size: fsz, em: fem });
    const limit = r3(fy - GAP);

    // 箇条書き（右）
    const BX = 7.35;
    const bl = bulletList(s, [
      brk(id, 'コンベックス、心窩部で矢状断。肝左葉・膵・大動脈が目印', '肝左葉'),
      [q(id, '仰臥位と右側臥位で前庭部を見る')],
      brk(id, 'Grade 0：両方で空、1：右側臥位だけ液体、2：両方で液体', '1：', '2：'),
      brk(id, '固形物が見える、または液体 1.5 mL/kg 以上 → 『満腹』', '液体').map((l, i) => (i === 1 ? mark(l, ['『満腹』', C.vermilion]) : l)),
    ], { x: BX, y: TOP, w: r3(XR - BX), size: 22, em: 1.2, indent: 0.28, gap: 0.12 });
    if (bl.bottom > limit + 1e-6) throw new Error(`${id}: 箇条書きの下端 ${bl.bottom} > ${limit}`);

    // 模式図（左）とラベル（図の外、引き出し線）
    const g = usFigure(s);
    const LR = 2.62, RL = 5.71;          // 左ラベルの右端・右ラベルの左端
    const lead = (x1, x2, y, color = '595959') => s.polyline([[r3(x1), r3(y)], [r3(x2), r3(y)]], { color, width: 1.25 });
    const leftLabel = (str, yc, tx, n = 1) => {
      const ls = Array.isArray(str) ? str : [str];
      const p = pitch(22, 1.2), h = (ls.length - 1) * p + boxH(22);
      lines(s, ls, { x: X0, y: r3(yc - h / 2), w: r3(LR - X0), size: 22, em: 1.2, align: 'right' });
      lead(LR + 0.07, tx, yc);
    };
    // 右ラベル: 左端は扇形の右縁から GAP 以上（RL より右に寄せる）
    const rightLabel = (str, yc, fx, color) => {
      const ls = Array.isArray(str) ? str : [str];
      const p = pitch(22, 1.2), h = (ls.length - 1) * p + boxH(22);
      const lx = r3(Math.max(RL, edgeX(yc + h / 2, 1) + GAP, edgeX(yc - h / 2, 1) + GAP));
      lines(s, ls, { x: lx, y: r3(yc - h / 2), w: r3(BX - GAP - lx), size: 22, em: 1.2, color: color || C.dark, bold: !!color });
      lead(fx, lx - 0.07, yc, color);
    };
    leftLabel(figWord(id, '皮膚'), 2.04, edgeX(2.04, -1) + 0.1);
    leftLabel(figWord(id, '肝左葉'), 2.62, edgeX(2.62, -1) + 0.3);
    leftLabel(figWord(id, '前庭部'), g.A.cy, g.A.cx - g.A.rx + 0.06);
    leftLabel(brk(id, '断面積をトレース', 'トレース').length && [figWord(id, '大動脈または'), figWord(id, '上腸間膜動脈')], 4.86, g.V.cx - g.V.r + 0.02);
    rightLabel(brk(id, '断面積をトレース', 'トレース'), g.A.cy, g.A.cx + g.A.rx + 0.08, C.teal);
    rightLabel(figWord(id, '膵'), g.P.cy, g.P.cx + g.P.rx - 0.06);
    if (!SPEC[id].figure.includes('大動脈または上腸間膜動脈')) throw new Error(`${id}: figure 行に「大動脈または上腸間膜動脈」がありません`);
    console.log(`  p${s.page} 図 ${US.yTop - 0.08}〜${r3(g.bottom)}in / 箇条書き ${bl.top}〜${bl.bottom}in / 式 ${fy}〜${r3(fy + fh)}in（上限 ${limit}）`);
  }

  // ===== S7 リスクが高いときの対応（3列カード＋強調ボックス） =====
  {
    const id = 'S7';
    const s = slide(id);
    const e = emphBox(s, {
      bottom: BOT, size: 26, lineColor: C.teal,
      lines: [mark(q(id, '迷ったら『満腹』として扱う。誤嚥は予防が治療'), ['『満腹』として扱う', C.vermilion])],
    });
    // 3 枚を同幅（各 約3.92in、間隔 0.24in）にする。内側の幅（約3.26in）に入らない項目は語句の切れ目で 2 行にする。
    cardRow(s, [
      {
        title: q(id, '待機手術'), titleColor: C.blue,
        items: [
          brk(id, '絶飲食と薬を整えて延期', '整えて'),
          brk(id, '胃エコーで確かめてから', '確かめて'),
          [q(id, '区域麻酔が選べるか')],
        ],
      },
      {
        title: q(id, '緊急手術'), titleColor: C.vermilion,
        items: [
          [q(id, '迅速導入（RSI）')],
          [q(id, '前酸素化を十分に')],
          brk(id, 'イレウスは胃管で減圧', '胃管'),
          [q(id, '覚醒してから抜管')],
        ],
      },
      {
        title: q(id, '準備'), titleColor: C.teal,
        items: [
          brk(id, '吸引を手元に、効くことを確認', '効く'),
          [q(id, '頭側を上げる体位')],
          [q(id, '上級医・応援に連絡')],
          brk(id, '筋弛緩薬は速く効く量', '速く'),
        ],
      },
    ], { y: TOP, bottom: r3(e.y - GAP), maxGap: 0.3, equal: true, style: { em: 1.3 } });
  }

  // ===== S8 まとめ（箇条書き5行） =====
  {
    const id = 'S8';
    const s = slide(id);
    const items = [
      [mark(q(id, '清澄水 2 時間、母乳 4 時間、人工乳・牛乳 6 時間'), ['2 時間', C.blue], ['4 時間', C.blue], ['6 時間', C.blue])],
      [mark(q(id, '『0 時から絶飲食』はやめる。時刻で書き、水を勧める'), ['はやめる', C.vermilion])],
      [mark(q(id, '時間を守っても胃が空でない患者がいる（緊急、外傷、DM、陣痛）'), ['胃が空でない', C.vermilion])],
      [mark(q(id, 'GLP-1：名前・用量・最後の投与日を聞く。症状があれば延期を検討'), ['最後の投与日', C.blue])],
      [mark(q(id, '迷ったら胃エコー。満腹なら延期か迅速導入'), ['胃エコー', C.blue])],
    ];
    const size = 24, em = 1.5, gap = 0.42;
    const total = items.length * pitch(size, em) + (items.length - 1) * gap;
    const bl = bulletList(s, items, { x: X0, y: r2(TOP + (BOT - TOP - total) / 2), w: 12.2, size, em, indent: 0.4, gap });
    console.log(`  p${s.page} 箇条書き ${bl.top}〜${bl.bottom}in`);
  }

  if (DEVIATIONS.length) console.log('[build] spec からの変更（deviations）:\n  ' + DEVIATIONS.join('\n  '));
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.render) console.log('[build] render: ' + res.render.images.length + ' images in ' + RENDER_DIR);
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.stack || e.message || e); process.exit(1); });
