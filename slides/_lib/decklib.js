'use strict';
/**
 * decklib.js — 川副式スライド (.pptx) 生成ライブラリ（pptxgenjs ラッパ）
 *
 * 使い方の要点は同じディレクトリの README.md を参照。
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 方針: lint（slide_lint.py）で落ちる組版はビルド時に Error を投げて止める。
 *   - 本文系ヘルパは 22pt 未満を受け付けない
 *   - 1段落が枠幅に収まらない（＝自動折返しが起きる）と Error
 *   - 推定描画高さが枠高さ + 0.3cm を超えると Error
 *   - タイトルが章チップに食い込むと Error
 *   - 本文ゾーンの重なり（lint と同じ判定）は save() 時に Error
 */

const fs = require('fs');
const path = require('path');
const { spawnSync } = require('child_process');

// ---------------------------------------------------------------------------
// 依存の読み込み（node_modules はリポジトリの外。NODE_PATH か PPTXGENJS_PATH で指定）
// ---------------------------------------------------------------------------
const PPTXGEN_ID = process.env.PPTXGENJS_PATH || 'pptxgenjs';
let pptxgen;
try {
  pptxgen = require(PPTXGEN_ID);
} catch (e) {
  throw new Error(
    '[decklib] pptxgenjs を読み込めません。`NODE_PATH=<pptxbuild>/node_modules node build.js` のように ' +
    'NODE_PATH を付けて実行するか、PPTXGENJS_PATH に pptxgenjs のパスを設定してください。\n' + e.message);
}
function loadJSZip() {
  try { return require('jszip'); } catch (e) { /* fallthrough */ }
  const base = path.dirname(require.resolve(PPTXGEN_ID));
  return require(require.resolve('jszip', { paths: [base] }));
}

// ---------------------------------------------------------------------------
// 定数
// ---------------------------------------------------------------------------
const FONT = 'メイリオ';
const LANG = 'ja-JP';

const C = {
  gold: 'BF9000', teal: '00A8AA', blue: '0072B2', vermilion: 'D55E00', red: 'C00000',
  gray: '7F7F7F', lightGray: 'D9D9D9', dark: '262626', bg: 'F2F2F2',
  white: 'FFFFFF', black: '000000', chipText: '404040', source: '595959', border: 'BFBFBF',
};

const SLIDE_W = 13.3333;
const SLIDE_H = 7.5;

/** レイアウト定数（インチ） */
const LAYOUT = Object.freeze({
  slideW: SLIDE_W,
  slideH: SLIDE_H,
  title: Object.freeze({ x: 0.55, y: 0.35, w: 9.3, h: 0.95, size: 44 }),
  underline: Object.freeze({ y: 1.52, width: 8 }),
  chip: Object.freeze({ w: 0.68, h: 0.64, y: 0.06, gap: 0.06, right: 13.13, radius: 0.1, size: 11 }),
  page: Object.freeze({ x: 12.3, y: 7.05, w: 0.8, h: 0.35, size: 12 }),
  source: Object.freeze({ x: 0.55, y: 6.85, w: 11.6, h: 0.5, size: 16, y2: 6.5, h2: 0.85 }),
  body: Object.freeze({ x: 0.55, y: 1.85, w: 12.23, right: 12.78, top: 1.6, bottom: 6.75 }),
  cover: Object.freeze({ titleBottom: 3.55, underlineY: 3.8, x: 0.55, w: 12.23 }),
});

// lint（slide_lint.py）と同じ係数・閾値
const CM = 2.54;
const JP_W = 1.12;   // Meiryo 横幅割増
const JP_H = 1.18;   // Meiryo 行高割増
const LINT = Object.freeze({
  tolIn: 0.20 / CM,                 // 枠外/内包の許容
  wrapTolIn: 0.15 / CM,             // 折返し判定の許容（lint 側）
  overflowTolIn: 0.30 / CM,         // あふれ判定の許容
  overlapMinIn2: (0.6 / CM) ** 2,   // 重なり最小面積
  overlapZone: [4.2 / CM, 17.8 / CM],
  textZone: [4.0 / CM, 18.2 / CM],
  chipGapIn: 0.2 / CM,              // タイトル推定右端とチップの最小間隔
  defaultSpaceAfterPt: 6.0,         // spcAft 未指定時に lint が仮定する値
});
const MEIRYO_LINE = 1.5;            // 実機 Meiryo の 1 行高（em）。警告用

const BODY_MIN_PT = 22;

// ---------------------------------------------------------------------------
// 推定器（slide_lint.py の _units / _text_w_emu / est_overflow と同一式）
// ---------------------------------------------------------------------------
/** 文字幅の単位数: 全角=1, ASCII/ラテン=0.56, 半角スペース=0.3 */
function units(text) {
  let u = 0;
  for (const ch of String(text)) {
    const o = ch.codePointAt(0);
    if (ch === ' ') u += 0.30;
    else if (o <= 0x24F) u += 0.56;
    else if (o >= 0x2080 && o <= 0x208E) u += 0.56;
    else u += 1.0;
  }
  return u;
}
/** 推定描画幅（インチ）= units × pt/72 × 1.12 */
function measure(text, size) {
  return units(text) * (size / 72) * JP_W;
}
/** 全角1字あたりの推定幅（インチ） */
function charW(size) { return (size / 72) * JP_W; }
/** その幅・サイズで1行に入る全角文字数 */
function maxChars(size, widthIn) {
  return Math.floor(widthIn / charW(size) + 1e-9);
}
/** 1行に収まるか（ライブラリ判定。lint より 0.15cm 厳しい = 許容なし） */
function fits(text, size, widthIn) {
  return measure(text, size) <= widthIn + 1e-9;
}
/** 詳細版: {ok, needIn, widthIn, units, overChars, maxChars} */
function fitInfo(text, size, widthIn) {
  const needIn = measure(text, size);
  const cw = charW(size);
  return {
    ok: needIn <= widthIn + 1e-9,
    needIn: +needIn.toFixed(3),
    widthIn: +(+widthIn).toFixed(3),
    units: +units(text).toFixed(2),
    overChars: +Math.max(0, (needIn - widthIn) / cw).toFixed(1),
    maxChars: maxChars(size, widthIn),
  };
}
/** lint 式の1行高（インチ）= pt/72 × 1.2 × 1.18 */
function lintLineIn(size) { return (size / 72) * 1.2 * JP_H; }
/**
 * lint 式の推定高さ（インチ）。paras: [{text, size, spaceAfter(pt)}]
 * 空段落は 0.2cm、spaceAfter 未指定/0 は 6pt として数える（lint と同じ）。
 */
function estHeightIn(paras, boxWIn) {
  let total = 0;
  for (const p of paras) {
    const t = String(p.text || '');
    if (!t.trim()) { total += 0.2 / CM; continue; }
    const lines = Math.max(1, Math.ceil(measure(t, p.size) / boxWIn - 1e-9));
    const sa = p.spaceAfter > 0 ? p.spaceAfter : LINT.defaultSpaceAfterPt;
    total += lines * lintLineIn(p.size) + sa / 72;
  }
  return total;
}
/** 実機 Meiryo 想定の高さ（インチ、警告用） */
function realHeightIn(paras) {
  let total = 0;
  paras.forEach((p, i) => {
    total += (p.size / 72) * MEIRYO_LINE;
    if (i < paras.length - 1 && p.spaceAfter > 0) total += p.spaceAfter / 72;
  });
  return total;
}
/** 必要な枠高さ（インチ）: lint 式（許容 0.3cm 込み）と実機推定の大きい方 */
function neededHeightIn(paras, boxWIn) {
  const lintNeed = estHeightIn(paras, boxWIn) - LINT.overflowTolIn;
  return Math.max(lintNeed, realHeightIn(paras));
}
/** 章チップ数 n のときのタイトル上限（全角文字数） */
function maxTitleChars(nChips) {
  const t = LAYOUT.title;
  let limitRight = t.x + t.w;
  if (nChips > 0) limitRight = Math.min(limitRight, chipLeft(nChips) - LINT.chipGapIn);
  return Math.floor((limitRight - t.x) / charW(t.size) + 1e-9);
}
/** n 行（各1段落）のテキストに必要な枠高さ（インチ）。text/bullets の h の目安 */
function heightFor(n, size = 24, gap = 0) {
  const paras = Array.from({ length: n }, () => ({ text: 'あ', size, spaceAfter: gap }));
  return Math.ceil(neededHeightIn(paras, 100) * 100) / 100;
}
/** 見出し付きカードに本文 n 行を入れるのに必要な最小 h（インチ）。card() の既定値前提 */
function minCardH(n, o = {}) {
  const pad = o.pad != null ? o.pad : 0.15;
  const ts = o.titleSize != null ? o.titleSize : 24;
  const bs = o.bodySize != null ? o.bodySize : 22;
  const gap = o.gap != null ? o.gap : 6;
  const th = o.title === false ? 0 : heightFor(1, ts, 0) + 0.08;
  return Math.ceil((2 * pad + th + (n > 0 ? heightFor(n, bs, gap) : 0)) * 100) / 100;
}
function chipLeft(n) {
  const ch = LAYOUT.chip;
  return ch.right - n * ch.w - (n - 1) * ch.gap;
}

// ---------------------------------------------------------------------------
// 小物
// ---------------------------------------------------------------------------
const r3 = v => Math.round(v * 1000) / 1000;
function clip(s, n = 16) { s = String(s).replace(/\n/g, '⏎'); return s.length > n ? s.slice(0, n) + '…' : s; }
function warn(msg) { console.warn('[decklib warn] ' + msg); }
function isNum(v) { return typeof v === 'number' && isFinite(v); }
function need(o, keys, where) {
  for (const k of keys) if (!isNum(o[k])) throw new Error(`[decklib] ${where}: ${k} (数値, インチ) が必要です`);
}
function color6(c, where) {
  if (c == null) return c;
  const s = String(c).replace(/^#/, '');
  if (!/^[0-9A-Fa-f]{6}$/.test(s)) throw new Error(`[decklib] ${where}: 色は 'RRGGBB' 形式で指定してください（${c}）`);
  return s.toUpperCase();
}
function lineSpec(line, where) {
  // line: undefined | null | 'RRGGBB' | {color,width,dash}
  if (!line) return undefined;
  if (typeof line === 'string') return { color: color6(line, where), width: 1.5 };
  const o = { color: color6(line.color || C.gray, where), width: line.width || 1.5 };
  if (line.dash) o.dashType = line.dash;
  return o;
}
function checkBodySize(size, where) {
  if (!isNum(size)) throw new Error(`[decklib] ${where}: size が数値ではありません`);
  if (size < BODY_MIN_PT) {
    throw new Error(`[decklib] ${where}: 本文系テキストは ${BODY_MIN_PT}pt 以上（指定 ${size}pt）。` +
      '小さくせず文言を減らすかスライドを分割してください（例外は出典16pt/ページ番号12pt/チップ11pt のみで、ライブラリが自動配置します）。');
  }
}

/**
 * 1行（=1段落）を正規化する。
 * line: 'テキスト' | ['通常', {text:'強調', color:C.red, bold:true}, ...] | {text, color, bold, size}
 */
function normLine(line, base, where) {
  let arr;
  if (line == null) arr = [''];
  else if (Array.isArray(line)) arr = line;
  else if (typeof line === 'object') arr = [line];
  else arr = [String(line)];
  const runs = arr.map(r => {
    const o = (typeof r === 'object' && r !== null) ? { ...r } : { text: String(r) };
    o.text = String(o.text == null ? '' : o.text);
    if (o.text.includes('\n')) throw new Error(`[decklib] ${where}: 1つの run の中に改行は入れられません（行ごとに分けてください）`);
    o.size = o.size != null ? o.size : base.size;
    checkBodySize(o.size, where);
    if (o.color) o.color = color6(o.color, where);
    return o;
  });
  const text = runs.map(r => r.text).join('');
  const size = Math.max(...runs.map(r => r.size));
  return { runs, text, size };
}
/** 文字列（\n 区切り）または行配列 → 行配列 */
function toLines(str) {
  if (Array.isArray(str)) return str;
  if (str == null) return [];
  return String(str).split('\n');
}

/** pptxgenjs の text 配列を組み立てる（段落ごとに run、段落末尾で breakLine） */
function buildRuns(paras, base) {
  const out = [];
  paras.forEach((p, pi) => {
    const last = pi === paras.length - 1;
    const runs = p.runs.filter(r => r.text !== '');
    if (runs.length === 0) {
      const o = { fontFace: FONT, lang: LANG, fontSize: p.size, color: base.color };
      if (!last) o.breakLine = true;
      out.push({ text: '', options: o });
      return;
    }
    runs.forEach((r, ri) => {
      const o = {
        fontFace: FONT, lang: LANG, fontSize: r.size,
        color: r.color || base.color,
      };
      const bold = r.bold != null ? r.bold : base.bold;
      if (bold) o.bold = true;
      if (r.italic) o.italic = true;
      if (r.underline) o.underline = { style: 'sng' };
      if (ri === 0 && p.bullet) o.bullet = p.bullet;
      if (ri === 0 && p.indentLevel) o.indentLevel = p.indentLevel;
      if (ri === runs.length - 1 && !last) o.breakLine = true;
      out.push({ text: r.text, options: o });
    });
  });
  return out;
}

// ---------------------------------------------------------------------------
// スライドラッパ
// ---------------------------------------------------------------------------
class SlideWrap {
  constructor(deck, slide, info) {
    this.deck = deck;
    this.raw = slide;               // pptxgenjs のスライド（逃げ道。ここに直接足すと検査対象外）
    this.page = info.page;          // lint の pN と一致（表紙を除く通し番号）
    this.title = info.title;
    this.isCover = !!info.isCover;
    this.hasSources = false;
    this.sourceTop = SLIDE_H;
    this._items = [];               // 重なり検査用 {kind, x,y,w,h, desc}
  }

  _where(what, txt) { return `${this.isCover ? '表紙' : 'p' + this.page}「${clip(this.title, 12)}」 ${what}${txt != null ? '「' + clip(txt) + '」' : ''}`; }

  _bounds(where, x, y, w, h) {
    if (x < -LINT.tolIn || y < -LINT.tolIn) throw new Error(`[decklib] ${where}: スライド左/上にはみ出しています (x=${r3(x)}, y=${r3(y)})`);
    if (x + w > SLIDE_W + LINT.tolIn * 0.5) throw new Error(`[decklib] ${where}: 右端 ${r3(x + w)}in > スライド幅 ${SLIDE_W}in`);
    if (y + h > SLIDE_H + LINT.tolIn * 0.5) throw new Error(`[decklib] ${where}: 下端 ${r3(y + h)}in > スライド高 ${SLIDE_H}in`);
    if (this.isCover) return;
    if (y < LAYOUT.body.top - 1e-6 && h > 0) warn(`${where}: 上端 y=${r3(y)}in が金下線(1.52in)直下の本文開始 ${LAYOUT.body.top}in より上です`);
    if (this.hasSources && y + h > this.sourceTop - 0.03 && h > 0) warn(`${where}: 下端 ${r3(y + h)}in が出典行（y=${this.sourceTop}in）に掛かっています`);
    else if (y + h > LAYOUT.body.bottom + 0.25) warn(`${where}: 下端 ${r3(y + h)}in がフッタ帯（>7.0in）に入っています`);
  }

  _reg(kind, x, y, w, h, desc) { this._items.push({ kind, x, y, w, h, desc }); }

  /** 段落群の幅・高さを検査（Error を投げる） */
  _checkParas(where, paras, availW, boxW, h) {
    paras.forEach((p, i) => {
      if (!p.text.trim()) return;
      const fi = fitInfo(p.text, p.size, availW);
      if (!fi.ok) {
        throw new Error(`[decklib] ${where} ${i + 1}行目「${clip(p.text, 24)}」が枠幅に収まりません: ` +
          `推定 ${fi.needIn}in > 使える幅 ${r3(availW)}in（全角換算 約${fi.overChars}字オーバー。` +
          `${p.size}pt・この幅では全角 ${fi.maxChars} 字まで。半角英数は0.56字）。` +
          '文言を短くするか、語の切れ目で "\\n" を入れて行を分けてください。');
      }
    });
    if (h != null) {
      const est = estHeightIn(paras, boxW);
      if (est > h + LINT.overflowTolIn + 1e-9) {
        throw new Error(`[decklib] ${where}: 縦にあふれます（lint 推定 ${r3(est)}in > 枠高 ${r3(h)}in + 0.12in）。` +
          `${paras.length} 段落。h を ${Math.ceil((est - LINT.overflowTolIn) * 100) / 100}in 以上にするか、行数を減らしてください。`);
      }
      const real = realHeightIn(paras);
      if (real > h + 0.05) warn(`${where}: 実機メイリオ推定 ${r3(real)}in > 枠高 ${r3(h)}in（PowerPoint で枠の下にはみ出す可能性）`);
    }
  }

  /**
   * テキストボックス。str は '\n' 区切り文字列、または行の配列（各行は文字列 or run 配列）。
   * opts: {x,y,w,h?, size=24, color, bold, align='left', valign='top', gap(pt)=0}
   */
  text(str, o = {}) {
    const where = this._where('text', Array.isArray(str) ? JSON.stringify(str).slice(0, 20) : str);
    need(o, ['x', 'y', 'w'], where);
    const size = o.size != null ? o.size : 24;
    checkBodySize(size, where);
    const base = { size, color: color6(o.color || C.dark, where), bold: !!o.bold };
    const gap = o.gap || 0;
    const paras = toLines(str).map(l => ({ ...normLine(l, base, where), spaceAfter: gap }));
    if (paras.length === 0) throw new Error(`[decklib] ${where}: テキストが空です`);
    const h = o.h != null ? o.h : Math.ceil(neededHeightIn(paras, o.w) * 100) / 100;
    this._bounds(where, o.x, o.y, o.w, h);
    this._checkParas(where, paras, o.w, o.w, h);
    this.raw.addText(buildRuns(paras, base), {
      x: o.x, y: o.y, w: o.w, h, margin: 0, isTextBox: true,
      fontFace: FONT, lang: LANG, fontSize: size,
      align: o.align || 'left', valign: o.valign || 'top',
      paraSpaceAfter: gap > 0 ? gap : undefined,
    });
    this._reg('txt', o.x, o.y, o.w, h, `text「${clip(paras[0].text)}」`);
    return this;
  }

  /** 小さめのラベル（既定 22pt・中央揃え・上下中央）。≥22pt。 */
  label(str, o = {}) {
    return this.text(str, {
      size: 22, align: 'center', valign: 'middle', ...o,
    });
  }

  /**
   * 箇条書き。items: 文字列 | run 配列 | {text: 文字列|run配列, level:0|1, color, bold}
   * opts: {x=0.55, y=1.85, w=12.2, h?, size=24, color, gap=10(pt), marker='•'}
   */
  bullets(items, o = {}) {
    const where = this._where('bullets', items && items[0] && (items[0].text || items[0]));
    const x = o.x != null ? o.x : LAYOUT.body.x;
    const y = o.y != null ? o.y : LAYOUT.body.y;
    const w = o.w != null ? o.w : 12.2;
    const size = o.size != null ? o.size : 24;
    checkBodySize(size, where);
    const gap = o.gap != null ? o.gap : 10;
    const base = { size, color: color6(o.color || C.dark, where), bold: !!o.bold };
    const indentPt = Math.round(size * 1.15);
    const bullet = { indent: indentPt };
    if (o.marker && o.marker !== '•') bullet.characterCode = o.marker.codePointAt(0).toString(16).toUpperCase().padStart(4, '0');
    if (!Array.isArray(items) || items.length === 0) throw new Error(`[decklib] ${where}: items が空です`);
    if (items.length > 6) warn(`${where}: 箇条書きが ${items.length} 個あります（1スライド最大5個程度が目安）`);
    const paras = items.map((it, i) => {
      let line = it, level = 0, b = base;
      if (it && typeof it === 'object' && !Array.isArray(it) && ('level' in it || Array.isArray(it.text))) {
        line = Array.isArray(it.text) ? it.text : [{ text: it.text, color: it.color, bold: it.bold, size: it.size }];
        level = it.level || 0;
      }
      const p = normLine(line, b, `${where} ${i + 1}項目`);
      if (!p.text.trim()) throw new Error(`[decklib] ${where}: ${i + 1}項目が空です`);
      return { ...p, level, spaceAfter: gap, bullet: { ...bullet }, indentLevel: level || undefined };
    });
    // 幅チェックは行頭記号のインデント分を差し引いた幅で行う（lint より厳しい）
    paras.forEach((p, i) => {
      const avail = w - (indentPt * (1 + p.level)) / 72;
      const fi = fitInfo(p.text, p.size, avail);
      if (!fi.ok) {
        throw new Error(`[decklib] ${where} ${i + 1}項目「${clip(p.text, 24)}」が1行に収まりません: ` +
          `推定 ${fi.needIn}in > 使える幅 ${r3(avail)}in（行頭記号ぶん差引後。全角換算 約${fi.overChars}字オーバー。` +
          `${p.size}pt では全角 ${fi.maxChars} 字まで）。1項目1行・体言止めで短くしてください。`);
      }
    });
    const h = o.h != null ? o.h : Math.ceil(neededHeightIn(paras, w) * 100) / 100;
    this._bounds(where, x, y, w, h);
    this._checkParas(where, paras, w, w, h);
    this.raw.addText(buildRuns(paras, base), {
      x, y, w, h, margin: 0, isTextBox: true,
      fontFace: FONT, lang: LANG, fontSize: size,
      align: 'left', valign: o.valign || 'top',
      paraSpaceAfter: gap > 0 ? gap : undefined,
    });
    this._reg('txt', x, y, w, h, `bullets「${clip(paras[0].text)}」`);
    return this;
  }

  /**
   * カード: 角丸の塗り箱 + 見出し + 本文（いずれも箱の内側。内側余白 0.15in）。
   * opts: {x,y,w,h, title, lines:[...], fill='F2F2F2', line, radius=0.08,
   *        titleColor, titleSize=24, titleAlign='left', bodySize=22, bodyColor, bodyAlign='left',
   *        bodyValign='top', bullets=false, gap=6}
   */
  card(o = {}) {
    const where = this._where('card', o.title || (o.lines && o.lines[0]));
    need(o, ['x', 'y', 'w', 'h'], where);
    const pad = o.pad != null ? o.pad : 0.15;
    const titleSize = o.titleSize != null ? o.titleSize : 24;
    const bodySize = o.bodySize != null ? o.bodySize : 22;
    checkBodySize(titleSize, where + ' 見出し');
    checkBodySize(bodySize, where + ' 本文');
    this._bounds(where, o.x, o.y, o.w, o.h);
    const fill = o.fill === null ? undefined : color6(o.fill || C.bg, where);
    this.raw.addShape('roundRect', {
      x: o.x, y: o.y, w: o.w, h: o.h,
      fill: fill ? { color: fill } : undefined,
      line: lineSpec(o.line, where),
      rectRadius: o.radius != null ? o.radius : 0.08,
    });
    this._reg('box', o.x, o.y, o.w, o.h, `card「${clip(o.title || '')}」`);
    const iw = o.w - 2 * pad;
    let cy = o.y + pad;
    const bottom = o.y + o.h - pad;
    if (o.title != null && o.title !== '') {
      const base = { size: titleSize, color: color6(o.titleColor || C.dark, where), bold: o.titleBold !== false };
      const tl = toLines(o.title).map(l => ({ ...normLine(l, base, where + ' 見出し'), spaceAfter: 0 }));
      const th = Math.ceil(neededHeightIn(tl, iw) * 100) / 100;
      if (cy + th > bottom + 1e-6) throw new Error(`[decklib] ${where}: カードが低すぎて見出しが入りません（h を ${r3(th + 2 * pad)}in 以上に）`);
      this._checkParas(where + ' 見出し', tl, iw, iw, th);
      this.raw.addText(buildRuns(tl, base), {
        x: o.x + pad, y: cy, w: iw, h: th, margin: 0, isTextBox: true,
        fontFace: FONT, lang: LANG, fontSize: titleSize,
        align: o.titleAlign || 'left', valign: 'top',
      });
      this._reg('txt', o.x + pad, cy, iw, th, `card見出し「${clip(tl[0].text)}」`);
      cy += th + 0.08;
    }
    const lines = o.lines || [];
    if (lines.length) {
      const bh = bottom - cy;
      if (bh <= 0.2) throw new Error(`[decklib] ${where}: 本文を入れる高さがありません（h を増やしてください）`);
      const gap = o.gap != null ? o.gap : 6;
      const base = { size: bodySize, color: color6(o.bodyColor || C.dark, where), bold: false };
      const indentPt = Math.round(bodySize * 1.15);
      const paras = lines.map(l => ({
        ...normLine(l, base, where + ' 本文'), spaceAfter: gap,
        bullet: o.bullets ? { indent: indentPt } : undefined,
      }));
      const avail = o.bullets ? iw - indentPt / 72 : iw;
      this._checkParas(where + ' 本文', paras, avail, iw, bh);
      this.raw.addText(buildRuns(paras, base), {
        x: o.x + pad, y: cy, w: iw, h: bh, margin: 0, isTextBox: true,
        fontFace: FONT, lang: LANG, fontSize: bodySize,
        align: o.bodyAlign || 'left', valign: o.bodyValign || 'top',
        paraSpaceAfter: gap > 0 ? gap : undefined,
      });
      this._reg('txt', o.x + pad, cy, iw, bh, `card本文「${clip(paras[0].text)}」`);
    }
    return this;
  }

  /**
   * 図形（塗り箱）。shape: 'rect' | 'roundRect' | 'ellipse'。text を渡すと図形内に文字（1図形）。
   * opts: {x,y,w,h, fill, line, radius, shape, text, size=24, color, bold, align='center', valign='middle', margin=0.08}
   */
  box(o = {}) {
    const where = this._where('box', o.text);
    need(o, ['x', 'y', 'w', 'h'], where);
    this._bounds(where, o.x, o.y, o.w, o.h);
    const shape = o.shape || (o.radius ? 'roundRect' : 'rect');
    const base = {
      x: o.x, y: o.y, w: o.w, h: o.h,
      fill: o.fill ? { color: color6(o.fill, where) } : undefined,
      line: lineSpec(o.line, where),
    };
    if (shape === 'roundRect') base.rectRadius = o.radius != null ? o.radius : 0.08;
    if (o.text != null && o.text !== '') {
      const size = o.size != null ? o.size : 24;
      checkBodySize(size, where);
      const m = o.margin != null ? o.margin : 0.08;
      const tb = { size, color: color6(o.color || C.dark, where), bold: !!o.bold };
      const paras = toLines(o.text).map(l => ({ ...normLine(l, tb, where), spaceAfter: 0 }));
      this._checkParas(where, paras, o.w - 2 * m, o.w, o.h);
      this.raw.addText(buildRuns(paras, tb), {
        ...base, shape, margin: m * 72,
        fontFace: FONT, lang: LANG, fontSize: size,
        align: o.align || 'center', valign: o.valign || 'middle',
      });
    } else {
      this.raw.addShape(shape, base);
    }
    this._reg('box', o.x, o.y, o.w, o.h, `box${o.text ? '「' + clip(o.text) + '」' : ''}`);
    return this;
  }

  /** 円（ellipse）。{x,y} は左上、{cx,cy} を渡すと中心指定。d=直径。 */
  circle(o = {}) {
    const d = o.d;
    if (!isNum(d)) throw new Error(`[decklib] ${this._where('circle')}: d (直径) が必要です`);
    const x = isNum(o.cx) ? o.cx - d / 2 : o.x;
    const y = isNum(o.cy) ? o.cy - d / 2 : o.y;
    return this.box({ ...o, x, y, w: d, h: d, shape: 'ellipse' });
  }

  /** 水平線 {x,y,w, color='7F7F7F', width=2(pt), dash} */
  hline(o = {}) {
    const where = this._where('hline');
    need(o, ['x', 'y', 'w'], where);
    this._bounds(where, o.x, o.y, o.w, 0);
    this.raw.addShape('line', { x: o.x, y: o.y, w: o.w, h: 0, line: lineSpec({ color: o.color || C.gray, width: o.width || 2, dash: o.dash }, where) });
    this._reg('box', o.x, o.y, o.w, 0, 'hline');
    return this;
  }

  /** 垂直線 {x,y,h, color='7F7F7F', width=2(pt), dash} */
  vline(o = {}) {
    const where = this._where('vline');
    need(o, ['x', 'y', 'h'], where);
    this._bounds(where, o.x, o.y, 0, o.h);
    this.raw.addShape('line', { x: o.x, y: o.y, w: 0, h: o.h, line: lineSpec({ color: o.color || C.gray, width: o.width || 2, dash: o.dash }, where) });
    this._reg('box', o.x, o.y, 0, o.h, 'vline');
    return this;
  }

  /** 矢印 (x1,y1)→(x2,y2)、終点に矢じり。{color='404040', width=3, dash, both=false} */
  arrow(o = {}) {
    const where = this._where('arrow');
    need(o, ['x1', 'y1', 'x2', 'y2'], where);
    const x = Math.min(o.x1, o.x2), y = Math.min(o.y1, o.y2);
    const w = Math.abs(o.x2 - o.x1), h = Math.abs(o.y2 - o.y1);
    this._bounds(where, x, y, w, h);
    const ln = lineSpec({ color: o.color || '404040', width: o.width || 3, dash: o.dash }, where);
    ln.endArrowType = o.head || 'triangle';
    if (o.both) ln.beginArrowType = o.head || 'triangle';
    this.raw.addShape('line', {
      x, y, w, h, line: ln,
      flipH: o.x2 < o.x1 || undefined,
      flipV: o.y2 < o.y1 || undefined,
    });
    if (w > 0.05 && h > 0.05) warn(`${where}: 斜めの矢印は外接矩形 ${r3(w)}×${r3(h)}in が lint の重なり判定対象になります`);
    this._reg('box', x, y, w, h, 'arrow');
    return this;
  }

  /**
   * 折れ線（波形など）を CUSTOM_GEOMETRY のネイティブ図形で描く。
   * points: [[x,y], ...]（スライド上の絶対インチ座標）
   * opts: {color='0072B2', width=3(pt), closed=false, fill}
   */
  polyline(points, o = {}) {
    const where = this._where('polyline');
    if (!Array.isArray(points) || points.length < 2) throw new Error(`[decklib] ${where}: points は2点以上 [[x,y],...]`);
    points.forEach((p, i) => { if (!isNum(p[0]) || !isNum(p[1])) throw new Error(`[decklib] ${where}: points[${i}] が不正`); });
    const xs = points.map(p => p[0]), ys = points.map(p => p[1]);
    const minX = Math.min(...xs), minY = Math.min(...ys);
    const w = Math.max(Math.max(...xs) - minX, 0.01), h = Math.max(Math.max(...ys) - minY, 0.01);
    this._bounds(where, minX, minY, w, h);
    const pts = points.map(p => ({ x: r3(p[0] - minX), y: r3(p[1] - minY) }));
    if (o.closed) pts.push({ close: true });
    this.raw.addShape('custGeom', {
      x: minX, y: minY, w, h, points: pts,
      line: lineSpec({ color: o.color || C.blue, width: o.width || 3, dash: o.dash }, where),
      fill: o.fill ? { color: color6(o.fill, where) } : undefined,
    });
    this._reg(null, minX, minY, w, h, 'polyline'); // freeform は lint の重なり判定対象外
    return this;
  }

  /**
   * 表（ネイティブ table）。rows: [[cell,...],...]、cell = 文字列（'\n' 可）| {text, color, bold, fill, align}
   * opts: {x,y,w?, colW:[...], size=22, header=true, headFill='D9D9D9', headColor, color, fill,
   *        border='BFBFBF', align='left', valign='middle', rowH(数値|配列), margin=[4,8,4,8](pt: 上右下左)}
   */
  table(rows, o = {}) {
    const where = this._where('table');
    need(o, ['x', 'y'], where);
    if (!Array.isArray(rows) || !rows.length) throw new Error(`[decklib] ${where}: rows が空です`);
    const size = o.size != null ? o.size : 22;
    checkBodySize(size, where);
    const nCol = Math.max(...rows.map(r => r.length));
    let colW = o.colW;
    if (!colW) {
      if (!isNum(o.w)) throw new Error(`[decklib] ${where}: colW か w を指定してください`);
      colW = Array(nCol).fill(r3(o.w / nCol));
    }
    if (colW.length !== nCol) throw new Error(`[decklib] ${where}: colW の数(${colW.length}) と列数(${nCol}) が違います`);
    const w = colW.reduce((a, b) => a + b, 0);
    const m = o.margin || [4, 8, 4, 8]; // pt: top,right,bottom,left
    const header = o.header !== false;
    const border = color6(o.border || C.border, where);
    const out = [];
    const rowH = [];
    rows.forEach((row, ri) => {
      const isHead = header && ri === 0;
      let maxLines = 1;
      const outRow = row.map((cell, ci) => {
        const c = (cell && typeof cell === 'object' && !Array.isArray(cell)) ? cell : { text: cell };
        const cs = c.size != null ? c.size : size;
        checkBodySize(cs, where);
        const lines = toLines(c.text == null ? '' : c.text);
        const avail = colW[ci] - (m[1] + m[3]) / 72;
        lines.forEach(l => {
          const p = normLine(l, { size: cs }, where);
          const fi = fitInfo(p.text, cs, avail);
          if (!fi.ok) {
            throw new Error(`[decklib] ${where} 行${ri + 1}列${ci + 1}「${clip(p.text, 20)}」がセル幅に収まりません: ` +
              `推定 ${fi.needIn}in > ${r3(avail)}in（全角換算 約${fi.overChars}字オーバー、全角 ${fi.maxChars} 字まで）。` +
              '短くするか "\\n" で分けるか colW を広げてください。');
          }
        });
        maxLines = Math.max(maxLines, lines.length);
        const colr = color6(c.color || (isHead ? (o.headColor || C.dark) : (o.color || C.dark)), where);
        const bold = c.bold != null ? c.bold : isHead;
        const fill = c.fill || (isHead ? (o.headFill || C.lightGray) : o.fill);
        const textArr = lines.map((l, li) => {
          const p = normLine(l, { size: cs }, where);
          return p.runs.map((r, k) => ({
            text: r.text,
            options: {
              fontFace: FONT, lang: LANG, fontSize: r.size, color: r.color || colr,
              bold: (r.bold != null ? r.bold : bold) || undefined,
              breakLine: (k === p.runs.length - 1 && li < lines.length - 1) || undefined,
            },
          }));
        }).flat();
        return {
          text: textArr.length ? textArr : '',
          options: {
            fill: fill ? { color: color6(fill, where) } : undefined,
            align: c.align || (isHead ? (o.headAlign || o.align || 'center') : (o.align || 'left')),
            valign: o.valign || 'middle',
            fontFace: FONT, fontSize: cs, color: colr, bold: bold || undefined,
          },
        };
      });
      out.push(outRow);
      const needH = maxLines * (size / 72) * MEIRYO_LINE + (m[0] + m[2]) / 72;
      let rh = Array.isArray(o.rowH) ? o.rowH[ri] : o.rowH;
      if (rh == null) rh = Math.ceil(needH * 100) / 100;
      else if (rh + 1e-6 < needH) throw new Error(`[decklib] ${where}: 行${ri + 1} の rowH ${rh}in では ${maxLines} 行が入りません（${r3(needH)}in 以上）`);
      rowH.push(rh);
    });
    const h = rowH.reduce((a, b) => a + b, 0);
    this._bounds(where, o.x, o.y, w, h);
    this.raw.addTable(out, {
      x: o.x, y: o.y, w, colW, rowH,
      fontFace: FONT, fontSize: size, margin: m,
      border: { type: 'solid', pt: o.borderPt || 1, color: border },
      valign: o.valign || 'middle',
    });
    this._reg('box', o.x, o.y, w, h, 'table');
    return { x: o.x, y: o.y, w, h, rowH };
  }

  /** ノートを追記 */
  notes(str) { if (str) this.raw.addNotes(String(str)); return this; }
}

// ---------------------------------------------------------------------------
// Deck
// ---------------------------------------------------------------------------
class Deck {
  /**
   * @param {object} o {file, chapters:[...], author, deckTitle, company, lintPath, validatePath, sofficePath}
   */
  constructor(o = {}) {
    if (!o.file) throw new Error('[decklib] Deck: file（出力 .pptx の絶対パス）が必要です');
    this.file = path.resolve(o.file);
    this.chapters = (o.chapters || []).map(String);
    this.opts = o;
    if (this.chapters.length > 6) throw new Error('[decklib] chapters は最大 6 個です');
    if (this.chapters.length === 1) warn('chapters が1個です（章チップは 2〜6 個を想定）');
    // 4 字（全角）の行があるデッキは、角丸の内接分で PowerPoint が折り返さないよう 10.5pt・角丸 0.05in にする
    // （roundRect の文字領域は角丸 r × (1−1/√2) ずつ内側に入る。11pt × 4 字 = 44pt は 0.68in・r0.1 の文字領域 44.7pt にほぼ等しい）
    const wide = this.chapters.some(lab => lab.split('\n').some(l => units(l) >= 3.5));
    this.chipSize = wide ? 10.5 : LAYOUT.chip.size;
    this.chipRadius = wide ? 0.05 : LAYOUT.chip.radius;
    this.chapters.forEach(lab => {
      const ls = lab.split('\n');
      if (ls.length > 2) throw new Error(`[decklib] 章ラベル「${lab}」: 2行まで`);
      ls.forEach(l => {
        // チップ 0.68in 幅・余白0 → 文字領域（角丸の内接分を引く）に全角 4 字まで
        const inner = (LAYOUT.chip.w - 2 * this.chipRadius * (1 - Math.SQRT1_2)) * 72;
        if (units(l) * this.chipSize > inner) {
          throw new Error(`[decklib] 章ラベル「${lab}」: チップ1行は全角4字まで（${units(l).toFixed(2)}字）。` +
            '5字以上は語の切れ目に "\\n" を入れて2行にしてください（例: "まとめ\\n確認"）。');
        }
      });
    });
    this.titleMax = maxTitleChars(this.chapters.length);

    const p = new pptxgen();
    p.layout = 'LAYOUT_WIDE';
    p.author = o.author || '';
    p.company = o.company || '';
    p.title = o.deckTitle || '';
    p.subject = o.subject || '';
    p.theme = { headFontFace: FONT, bodyFontFace: FONT };
    // NOTE: pptxgenjs はスライド側 addText({placeholder}) の位置・揃えをレイアウト側の値で上書きする
    //       （Object.assign(slideOpts, layoutOpts)）。そのため表紙用と本文用でレイアウトを分ける。
    const t = LAYOUT.title, cv = LAYOUT.cover;
    const titleStyle = {
      name: 'title', type: 'title', fontFace: FONT, fontSize: t.size, bold: true, color: C.gold,
      outline: { size: 2.25, color: C.black }, margin: [0, 0, 0, 0],
    };
    p.defineSlideMaster({
      title: 'KAWAZOE',
      background: { color: C.white },
      objects: [{ placeholder: { options: { ...titleStyle, x: t.x, y: t.y, w: t.w, h: t.h, align: 'left', valign: 'middle' }, text: '' } }],
    });
    p.defineSlideMaster({
      title: 'KAWAZOE_COVER',
      background: { color: C.white },
      objects: [{ placeholder: { options: { ...titleStyle, x: cv.x, y: cv.titleBottom - 1.9, w: cv.w, h: 1.9, align: 'center', valign: 'bottom' }, text: '' } }],
    });
    this.pptx = p;
    this.slides = [];
    this._page = 0;
    this._hasCover = false;
  }

  _addTitle(slide, text, box, align, valign = 'middle') {
    slide.addText(text, {
      placeholder: 'title',
      x: box.x, y: box.y, w: box.w, h: box.h,
      fontFace: FONT, lang: LANG, fontSize: LAYOUT.title.size, bold: true, color: C.gold,
      outline: { size: 2.25, color: C.black },
      align, valign, margin: 0,
    });
  }

  /** 表紙 {title, subtitle, meta, notes}。title/meta は '\n' で 2 行まで可。 */
  cover(o = {}) {
    if (this.slides.length) throw new Error('[decklib] cover() は最初に1回だけ呼んでください');
    if (!o.title) throw new Error('[decklib] cover: title が必要です');
    const slide = this.pptx.addSlide({ masterName: 'KAWAZOE_COVER' });
    const s = new SlideWrap(this, slide, { page: 0, title: o.title, isCover: true });
    const L = LAYOUT.cover;
    const tl = String(o.title).split('\n');
    if (tl.length > 2) throw new Error('[decklib] cover: title は2行まで');
    tl.forEach(l => {
      const fi = fitInfo(l, 44, L.w);
      if (!fi.ok) throw new Error(`[decklib] 表紙タイトル「${l}」が1行に収まりません（44pt・幅${L.w}in で全角 ${fi.maxChars} 字まで、約${fi.overChars}字オーバー）`);
    });
    // 位置はレイアウト KAWAZOE_COVER のプレースホルダ（下端 3.55in・下揃え・中央）で決まる
    this._addTitle(slide, tl.join('\n'), { x: L.x, y: L.titleBottom - 1.9, w: L.w, h: 1.9 }, 'center', 'bottom');
    slide.addShape('line', { x: 0, y: L.underlineY, w: SLIDE_W, h: 0, line: { color: C.gold, width: LAYOUT.underline.width } });
    let y = L.underlineY + 0.3;
    if (o.subtitle) {
      s.text(o.subtitle, { x: L.x, y, w: L.w, size: 28, align: 'center', color: C.dark });
      y += 0.3 + realHeightIn(toLines(o.subtitle).map(() => ({ size: 28 })));
    }
    if (o.meta) {
      s.text(o.meta, { x: L.x, y: y + 0.15, w: L.w, size: 22, align: 'center', color: C.source });
    }
    if (o.notes) slide.addNotes(String(o.notes));
    this.slides.push(s);
    this._hasCover = true;
    return s;
  }

  /**
   * 本文スライド {title, chapter, notes, sources:[...]}
   * chapter: chapters の 0 始まり番号（またはラベル文字列）
   */
  slide(o = {}) {
    if (!o.title) throw new Error('[decklib] slide: title が必要です');
    if (!this._hasCover && this.slides.length === 0) warn('cover() の前に slide() が呼ばれました（lint は1枚目を表紙として検査しません）');
    const page = ++this._page;
    const title = String(o.title);
    const tag = `p${page}「${clip(title, 14)}」`;
    if (title.includes('\n')) throw new Error(`[decklib] ${tag}: タイトルは1行（改行不可）`);
    // --- タイトル長さ（lint: TITLEパンくず侵入 / TITLE2行化）
    const t = LAYOUT.title;
    const n = this.chapters.length;
    const tright = t.x + measure(title, t.size);
    const limit = n > 0 ? Math.min(t.x + t.w, chipLeft(n) - LINT.chipGapIn) : t.x + t.w;
    if (tright > limit + 1e-9) {
      throw new Error(`[decklib] ${tag}: タイトルが長すぎます（推定右端 ${r3(tright)}in > 上限 ${r3(limit)}in、` +
        `全角換算 ${units(title).toFixed(2)} 字）。章チップ ${n} 個のときタイトルは全角 ${this.titleMax} 字まで` +
        '（半角英数 0.56 字・半角スペース 0.3 字換算）。助詞を削る・体言止めで短くしてください。');
    }
    // --- 章番号
    let ci = o.chapter;
    if (typeof ci === 'string') {
      const k = this.chapters.indexOf(ci);
      if (k < 0) throw new Error(`[decklib] ${tag}: chapter「${ci}」が chapters にありません`);
      ci = k;
    }
    if (n > 0 && !(Number.isInteger(ci) && ci >= 0 && ci < n)) {
      throw new Error(`[decklib] ${tag}: chapter は 0〜${n - 1} の整数（chapters の添字）で指定してください（指定: ${o.chapter}）`);
    }

    const slide = this.pptx.addSlide({ masterName: 'KAWAZOE' });
    const s = new SlideWrap(this, slide, { page, title });
    // タイトル（プレースホルダ idx=0 に後処理で正規化）
    this._addTitle(slide, title, t, 'left');
    // 金下線
    slide.addShape('line', { x: 0, y: LAYOUT.underline.y, w: SLIDE_W, h: 0, line: { color: C.gold, width: LAYOUT.underline.width } });
    // 章チップ
    if (n > 0) {
      const ch = LAYOUT.chip;
      let x = chipLeft(n);
      this.chapters.forEach((lab, k) => {
        const cur = k === ci;
        const ls = lab.split('\n');
        slide.addText(ls.map((l, li) => ({
          text: l,
          options: {
            fontFace: FONT, lang: LANG, fontSize: this.chipSize, color: cur ? C.white : C.chipText,
            bold: cur || undefined, breakLine: li < ls.length - 1 || undefined,
          },
        })), {
          shape: 'roundRect', x: r3(x), y: ch.y, w: ch.w, h: ch.h, rectRadius: this.chipRadius,
          fill: { color: cur ? C.teal : C.lightGray },
          fontFace: FONT, lang: LANG, fontSize: this.chipSize, align: 'center', valign: 'middle', margin: 0, wrap: false,
        });
        x += ch.w + ch.gap;
      });
    }
    // ページ番号（lint の pN と同じ、表紙を除く通し番号）
    const pg = LAYOUT.page;
    slide.addText(String(page), {
      x: pg.x, y: pg.y, w: pg.w, h: pg.h, margin: 0, isTextBox: true,
      fontFace: FONT, lang: LANG, fontSize: pg.size, color: C.gray, align: 'right', valign: 'bottom',
    });
    // 出典
    if (o.sources && o.sources.length) {
      const items = (Array.isArray(o.sources) ? o.sources : [o.sources]).map(String);
      const S = LAYOUT.source;
      const lines = [];
      for (const it of items) {
        if (!fits(it, S.size, S.w)) {
          const fi = fitInfo(it, S.size, S.w);
          throw new Error(`[decklib] ${tag}: 出典「${clip(it, 24)}」が1行に収まりません（16pt・全角 ${fi.maxChars} 字まで、約${fi.overChars}字オーバー）。略記にして全文はノートへ。`);
        }
        const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
        if (cur && fits(cur, S.size, S.w)) lines[lines.length - 1] = cur;
        else lines.push(it);
      }
      if (lines.length > 2) throw new Error(`[decklib] ${tag}: 出典が2行に収まりません（${items.length} 件）。件数を減らし全文はノートへ。`);
      const two = lines.length === 2;
      const y = two ? S.y2 : S.y, h = two ? S.h2 : S.h;
      slide.addText(lines.map((l, li) => ({
        text: l, options: { fontFace: FONT, lang: LANG, fontSize: S.size, color: C.source, breakLine: li < lines.length - 1 || undefined },
      })), {
        x: S.x, y, w: S.w, h, margin: 0, isTextBox: true,
        fontFace: FONT, lang: LANG, fontSize: S.size, align: 'left', valign: 'bottom',
      });
      s.hasSources = true;
      s.sourceTop = y;
      s._reg('txt', S.x, y, S.w, h, '出典');
    }
    if (o.notes) slide.addNotes(String(o.notes));
    this.slides.push(s);
    return s;
  }

  /** 重なり検査（slide_lint.py と同じ判定）。問題のリストを返す */
  checkOverlaps() {
    const issues = [];
    const tol = LINT.tolIn;
    const contains = (a, b) => a.x - tol <= b.x && a.y - tol <= b.y && a.x + a.w + tol >= b.x + b.w && a.y + a.h + tol >= b.y + b.h;
    const area = (a, b) => Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)) * Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
    for (const s of this.slides) {
      if (s.isCover) continue;
      const [z0, z1] = LINT.overlapZone;
      const items = s._items.filter(it => it.kind && (it.y + it.h / 2) >= z0 && (it.y + it.h / 2) <= z1);
      for (let i = 0; i < items.length; i++) {
        for (let j = i + 1; j < items.length; j++) {
          const a = items[i], b = items[j];
          if (a.kind === 'txt' && b.kind === 'txt') continue;
          if (contains(a, b) || contains(b, a)) continue;
          const ar = area(a, b);
          if (ar > LINT.overlapMinIn2) {
            issues.push(`p${s.page}「${clip(s.title, 12)}」 重なり ${a.desc} × ${b.desc}（${(ar * CM * CM).toFixed(1)}cm²）`);
          }
        }
      }
    }
    return issues;
  }

  /**
   * 書き出し → 後処理 → validate.py / slide_lint.py 実行。
   * opts: {strict=true（重なりで Error）, check=true（validate/lint を走らせる）, render=false（PDF+JPG を作る）, renderDir}
   * @returns {Promise<{file, validate, lint, violations, render}>}
   */
  async save(opts = {}) {
    const strict = opts.strict !== false;
    const ov = this.checkOverlaps();
    if (ov.length) {
      const msg = '[decklib] 本文ゾーンの重なり（lint と同判定）:\n  ' + ov.join('\n  ');
      if (strict) throw new Error(msg + '\n図形をずらす/小さくするか、文字は箱の内側に入れてください（save({strict:false}) で書き出しだけ可能）。');
      warn(msg);
    }
    if (!this._hasCover) warn('表紙がありません（lint は 1 枚目を表紙として検査対象外にします）');

    const buf = await this.pptx.write({ outputType: 'nodebuffer' });
    const out = await postProcess(buf);
    fs.mkdirSync(path.dirname(this.file), { recursive: true });
    fs.writeFileSync(this.file, out);
    console.log(`[decklib] wrote ${this.file} (${this.slides.length} slides)`);

    const res = { file: this.file, validate: null, lint: null, violations: null, render: null };
    if (opts.check !== false) {
      const tools = findTools(this.opts);
      if (tools.validate) {
        const r = spawnSync('python3', [tools.validate, this.file], { encoding: 'utf8' });
        res.validate = ((r.stdout || '') + (r.stderr || '')).trim();
        res.validateOk = r.status === 0;
        console.log('[decklib] validate.py: ' + res.validate);
        if (r.status !== 0) process.exitCode = 1;
      } else warn('validate.py が見つからないため OOXML 検証を省略（DECKLIB_VALIDATE で指定可）');
      if (tools.lint) {
        const r = spawnSync('python3', [tools.lint, this.file], { encoding: 'utf8' });
        res.lint = ((r.stdout || '') + (r.stderr || '')).trim();
        const m = res.lint.match(/違反 合計 (\d+) 件/);
        res.violations = m ? Number(m[1]) : null;
        console.log('[decklib] slide_lint.py:\n' + res.lint);
        if (res.violations !== 0) process.exitCode = 1;
      } else warn('slide_lint.py が見つからないため lint を省略（DECKLIB_LINT で指定可）');
      if (opts.render) res.render = render(this.file, tools.soffice, opts.renderDir);
    }
    return res;
  }
}

// ---------------------------------------------------------------------------
// 後処理（pptx の zip を直接編集）
// ---------------------------------------------------------------------------
async function postProcess(buf) {
  const JSZip = loadJSZip();
  const zip = await JSZip.loadAsync(buf);
  const names = Object.keys(zip.files);

  // 1) テーマフォント: major/minor の latin・ea・Jpan をメイリオに
  const themeName = 'ppt/theme/theme1.xml';
  let theme = await zip.file(themeName).async('string');
  theme = theme.replace(/<a:(majorFont|minorFont)>([\s\S]*?)<\/a:\1>/g, (m, tag, inner) => {
    inner = inner
      .replace(/<a:latin\b[^>]*\/>/, `<a:latin typeface="${FONT}"/>`)
      .replace(/<a:ea\b[^>]*\/>/, `<a:ea typeface="${FONT}"/>`)
      .replace(/<a:font script="Jpan" typeface="[^"]*"\/>/, `<a:font script="Jpan" typeface="${FONT}"/>`);
    return `<a:${tag}>${inner}</a:${tag}>`;
  });
  zip.file(themeName, theme);

  // 2) タイトルプレースホルダを idx なし（= idx 0）の type="title" に正規化
  const phFix = x => x.replace(/<p:ph\b[^>]*\btype="title"[^>]*\/>/g, '<p:ph type="title"/>');
  for (const n of names.filter(n => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n))) {
    let x = phFix(await zip.file(n).async('string'));
    // レイアウト側のタイトル枠も余白 0・アンカーを合わせる（PowerPoint で手動追加したスライド用）
    if (x.includes('<p:ph type="title"/>')) {
      const anchor = /<a:lvl1pPr algn="ctr"/.test(x) ? 'b' : 'ctr';
      x = x.replace('<a:bodyPr wrap="square" rtlCol="0"></a:bodyPr>',
        `<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="${anchor}"></a:bodyPr>`);
    }
    zip.file(n, x);
  }
  // 3) スライド: ph 正規化 / 段落内の重複 pPr 除去（pptxgenjs のバグ）/ shape id の振り直し
  for (const n of names.filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))) {
    let x = await zip.file(n).async('string');
    x = phFix(x);
    x = x.replace(/<p:cNvPr id="(\d+)" name="[^"]*"><\/p:cNvPr><p:cNvSpPr\/><p:nvPr><p:ph type="title"\/>/g,
      '<p:cNvPr id="$1" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/>');
    x = x.replace(/<a:p>([\s\S]*?)<\/a:p>/g, (m, inner) => {
      let seen = false;
      const fixed = inner.replace(/<a:pPr\b[^>]*?(?:\/>|>[\s\S]*?<\/a:pPr>)/g, (pp, off) => {
        if (!seen && off === 0) { seen = true; return pp; }
        seen = true;
        return '';
      });
      return '<a:p>' + fixed + '</a:p>';
    });
    let id = 0;
    x = x.replace(/<p:cNvPr id="\d+"/g, () => `<p:cNvPr id="${++id}"`);
    zip.file(n, x);
  }
  // 4) ノート: 改行ごとに段落を分ける（pptxgenjs は CRLF を1つの a:t に入れてしまう）
  for (const n of names.filter(n => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(n))) {
    let x = await zip.file(n).async('string');
    x = x.replace(/<a:p><a:r><a:rPr lang="en-US" dirty="0"\/><a:t>([\s\S]*?)<\/a:t><\/a:r><a:endParaRPr lang="en-US" dirty="0"\/><\/a:p>/,
      (m, t) => t.split(/\r?\n/).map(l => l === ''
        ? `<a:p><a:endParaRPr lang="${LANG}" altLang="en-US" dirty="0"/></a:p>`
        : `<a:p><a:r><a:rPr lang="${LANG}" altLang="en-US" dirty="0"/><a:t>${l}</a:t></a:r></a:p>`).join(''));
    zip.file(n, x);
  }
  // 検証: テーマに ea メイリオが入ったか
  const th2 = await zip.file(themeName).async('string');
  const nEa = (th2.match(new RegExp(`<a:ea typeface="${FONT}"/>`, 'g')) || []).length;
  if (nEa < 2) throw new Error('[decklib] テーマフォントの後処理に失敗しました（ea=メイリオ が major/minor に入っていません）');
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 6 } });
}

// ---------------------------------------------------------------------------
// 外部ツールの場所（スキルのディレクトリは UUID 付きなので探索する）
// ---------------------------------------------------------------------------
function findFile(roots, rel, depth = 4) {
  for (const root of roots) {
    if (!root || !fs.existsSync(root)) continue;
    const stack = [[root, 0]];
    while (stack.length) {
      const [d, k] = stack.pop();
      const cand = path.join(d, rel);
      if (fs.existsSync(cand)) return cand;
      if (k >= depth) continue;
      let ents = [];
      try { ents = fs.readdirSync(d, { withFileTypes: true }); } catch (e) { continue; }
      for (const e of ents) if (e.isDirectory() && !e.name.startsWith('.') && e.name !== 'node_modules') stack.push([path.join(d, e.name), k + 1]);
    }
  }
  return null;
}
function findTools(o = {}) {
  const home = process.env.HOME || '/root';
  const roots = [process.env.DECKLIB_SKILLS_DIR, path.join(home, '.claude', 'skills'), '/root/.claude/skills'];
  return {
    lint: o.lintPath || process.env.DECKLIB_LINT || findFile(roots, path.join('slide-format', 'scripts', 'slide_lint.py')),
    validate: o.validatePath || process.env.DECKLIB_VALIDATE || findFile(roots, path.join('pptx', 'scripts', 'office', 'validate.py')),
    soffice: o.sofficePath || process.env.DECKLIB_SOFFICE || findFile(roots, path.join('pptx', 'scripts', 'office', 'soffice.py')),
  };
}
/** PDF と slide-N.jpg を作る（目視確認用）。返り値: {pdf, images:[...]} */
function render(file, sofficePy, dir) {
  const outDir = dir || path.join(path.dirname(file), path.basename(file, '.pptx') + '_render');
  fs.mkdirSync(outDir, { recursive: true });
  const args = ['--headless', '--convert-to', 'pdf', '--outdir', outDir, file];
  const r = sofficePy ? spawnSync('python3', [sofficePy, ...args], { encoding: 'utf8' })
    : spawnSync('soffice', args, { encoding: 'utf8' });
  const pdf = path.join(outDir, path.basename(file, '.pptx') + '.pdf');
  if (!fs.existsSync(pdf)) { warn('PDF 変換に失敗: ' + (r.stderr || r.stdout)); return null; }
  for (const f of fs.readdirSync(outDir)) if (/^slide-\d+\.jpg$/.test(f)) fs.unlinkSync(path.join(outDir, f));
  spawnSync('pdftoppm', ['-jpeg', '-r', '80', pdf, path.join(outDir, 'slide')], { encoding: 'utf8' });
  const images = fs.readdirSync(outDir).filter(f => /^slide-\d+\.jpg$/.test(f)).sort().map(f => path.join(outDir, f));
  console.log(`[decklib] rendered ${images.length} images → ${outDir}`);
  return { pdf, images };
}

module.exports = {
  Deck, C, FONT, LAYOUT, LINT,
  units, measure, fits, fitInfo, maxChars, maxTitleChars, estHeightIn, realHeightIn, lintLineIn,
  heightFor, minCardH,
  postProcess, findTools,
};
