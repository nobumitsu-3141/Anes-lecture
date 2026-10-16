'use strict';
/**
 * 4-11 高齢者とフレイル・認知機能
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk()/q() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。spec から文言を変えた箇所があれば alt() で照合・記録し、ビルド時に一覧を出す。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら scratchpad の render/<NAME>）。
 *
 * 構成（spec の S1〜S10。内容スライド 9 枚）:
 *  - S2・S5・S8・S9: 箇条書き 5 行（24pt・段落後 10pt）＋全幅の強調ボックス（28pt 太字 2 行）。
 *  - S6・S7: 箇条書き 6 行（24pt・段落後 10pt）＋全幅の強調ボックス（24pt 太字 1 行）。
 *  - S3・S4: 2 列カード（幅 5.995in ずつ）＋強調ボックス（24pt 1 行）。カード本文は 22pt のまま、1 行 1 枠の s.text を
 *    行送り 1.1 em で置く（spcPts は slide_lint の高さ推定を壊し、spcPct は代替フォントで行送りが変わるため）。
 *    項目の頭は濃いグレーの小さな丸（ネイティブ図形）、見出しは青、オレンジは強調語だけ。
 *    カードの内側余白は字面から上下左右とも約 0.2in。項目の間隔は左右のカードで同じ値。
 *  - S10: まとめ（箇条書き 5 行）。
 *  - 箇条書きの行頭記号の色は 262626 に固定（fixBulletColor()）。
 *  - 強調色は 1 スライド 2 色まで。警告のレッド（太字＋下線）があるスライド（S5・S6・S10）は、ほかの強調を青にする
 *    （レッドとバーミリオンを並べない）。
 *  - 出典は全スライド 2 行（sourcesOf()）。区切りは全スライド ' / '。
 */
const fs = require('fs');
const path = require('path');
const { Deck, C, LAYOUT, fits, measure, heightFor } = require('../_lib/decklib');

const NAME = '4-11_elderly_frailty_cognition';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR ||
  path.join('/tmp/claude-0/-home-user-Anes-lecture/0b714c0d-c11e-535a-a97c-0e6ce9d53e16/scratchpad/render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（見出し・ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
const SPEC_MD = fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8');
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
    const head = part.split('\n')[0];
    const title = (head.match(/title=「([^」]*)」/) || [])[1];
    const ch = (head.match(/\bch=(\d+)/) || [])[1];
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
    out[id] = { title, ch: ch != null ? Number(ch) : null, notes, sources: sources && sources !== '—' ? sources : null, quoted };
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
/** spec の文言 str を、breaks の各語の直前で改行して行配列にする（文言は変えない。改行位置の半角スペースは落とす） */
function brk(id, str, ...breaks) {
  q(id, str);
  return split(str, breaks, id);
}
function split(str, breaks, id) {
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
/** spec の文言を収まりのため短くした箇所（spec 側の元の文言を照合し、ビルド時に一覧を出す） */
const DEVIATIONS = [];
function alt(id, orig, shown, why) {
  q(id, orig);
  DEVIATIONS.push(`${id}: 「${orig}」→「${shown}」（${why}）`);
  return shown;
}
/** alt() の結果を breaks の語の直前で改行する */
function altBrk(id, orig, shown, why, ...breaks) { return split(alt(id, orig, shown, why), breaks, id); }

// ---------------------------------------------------------------------------
// 出典: spec の sources 行を 16pt・幅 11.6in の 2 行に収める
//  0) ライブラリ既定（' / ' の区切りでだけ改行）で 2 行に入ればそのまま渡す。
//  1) 区切りだけでは入らないとき（この講義の出典は 3 件とも 1 件 5.4〜10.3in で、どの 2 件を並べても
//     11.6in を超える）は、文献の途中で改行するが、著者名・誌名の途中では割らない。改行してよい位置は
//     区切り（rank 0）＞ 著者と誌名の間（「et al.」「Theou O.」など著者の末尾の「.」の後。rank 1）＞
//     誌名と年の間（rank 2）の順。
//  2) （欠番）区切りの前後の空白を詰める（' / ' → '/'）段階は使わない（区切りを全スライドで ' / ' にそろえるため）。
//  3) 1) の位置で入らなければ（）内の補足を後ろの文献から外す（deviations に記録）。
//  4) どの段階でも 1) の位置で割れないときだけ、文言はそのままで残りの空白で改行する（ログに出す）。
//  5) それでも入らない（字数そのものが 2 行を超える）ときだけ、SRC_ABBR の略記を使う（deviations に記録）。
// 文献そのものは落とさない（落とさないと入らないときは Error）。
// ---------------------------------------------------------------------------
const SRC = LAYOUT.source;
function libLines(items) {
  const lines = [];
  for (const it of items) {
    const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
    if (cur && fits(cur, SRC.size, SRC.w)) lines[lines.length - 1] = cur;
    else lines.push(it);
  }
  return lines;
}
const SPLIT_KIND = ['区切り', '著者と誌名の間', '誌名と年の間', '語の切れ目'];
function split2(str, sep, maxRank) {
  const fitsLine = l => fits(l, SRC.size, SRC.w);
  const cands = [];
  for (let i = 1; i < str.length; i++) {
    let a, b, rank;
    if (str.startsWith(sep, i)) { a = str.slice(0, i); b = str.slice(i + sep.length); rank = 0; }
    else if (str[i] === ' ') {
      a = str.slice(0, i); b = str.slice(i + 1);
      if (a.endsWith('/') || b.startsWith('/')) continue;
      if (/ et$/.test(a)) continue;                   // 「et」と「al.」の間では切らない
      if (a.endsWith('.')) rank = 1;                   // 著者の末尾（et al. / Theou O. / Society. など）の後
      else if (/^\d/.test(b)) rank = 2;                // 誌名と年の間
      else rank = 3;
    } else continue;
    a = a.trim(); b = b.trim();
    if (rank <= maxRank && a && b && fitsLine(a) && fitsLine(b)) cands.push({ a, b, rank });
  }
  if (!cands.length) return null;
  cands.sort((x, y) => x.rank - y.rank || y.a.length - x.a.length);
  return cands[0];
}
/** spec の sources 行（文字列）を 2 行にする。入らなければ null。成功したときだけ deviations／ログを出す */
function trySources(id, spec, pre) {
  const items = spec.split(' / ');
  if (libLines(items).length <= 2) {
    if (pre) DEVIATIONS.push(`${id} 出典: ${pre}。区切り（' / '）で 2 行に入る`);
    return items;
  }
  // 区切りは全スライドで ' / ' にそろえる（'/' に詰めると S3 だけ区切りが違って見えるため、詰める段階は使わない）
  const stages = [
    { items, sep: ' / ', dev: null },
  ];
  const cut = items.slice(), removed = [];
  for (let k = cut.length - 1; k >= 0; k--) {
    if (!/（[^）]*）$/.test(cut[k])) continue;
    const was = cut[k];
    cut[k] = cut[k].replace(/（[^）]*）$/, '');
    removed.push(`「${was}」→「${cut[k]}」`);
    const its = cut.slice(), rm = removed.join('、');
    stages.push({ items: its, sep: ' / ', dev: `${rm}（（）内の補足を外した。区切りは ' / ' のまま）` });
  }
  for (const st of stages) {
    const r = split2(st.items.join(st.sep), st.sep, 2);
    if (!r) continue;
    const dev = [pre, st.dev].filter(Boolean).join('。さらに');
    if (dev) {
      DEVIATIONS.push(`${id} 出典: ${dev}。文言のままでは著者名・誌名を割らずに 16pt 2 行に入らないため。` +
        `改行は${SPLIT_KIND[r.rank]}: ${r.a} ⏎ ${r.b}`);
    } else console.log(`  ${id} 出典: ${SPLIT_KIND[r.rank]}で改行（文言はそのまま）: ${r.a} ⏎ ${r.b}`);
    return [r.a, r.b];
  }
  const r = split2(spec, ' / ', 3);
  if (r) {
    if (pre) DEVIATIONS.push(`${id} 出典: ${pre}。改行は語の切れ目: ${r.a} ⏎ ${r.b}`);
    else console.log(`  ${id} 出典: 区切りの空白を詰めても（）の補足を外しても、著者名・誌名を割らない位置では 2 行に入らない。` +
      `文言はそのままで語の切れ目で改行: ${r.a} ⏎ ${r.b}`);
    return [r.a, r.b];
  }
  return null;
}
/**
 * 5) 1)〜4) のどれでも 2 行に入らない出典（文字数そのものが 16pt 2 行＝23.2in を超える）だけ、文献は落とさずに
 *    著者名を略記にする。略記は spec のノートの「参考」行にある形を使う（deviations に記録）。
 */
const SRC_ABBR = {
  // S7: spec の文言のままで推定 25.23in（区切りを詰めても 24.93in）＞ 2 行ぶん 23.2in。（）の補足はない。
  //     「American Geriatrics Society.」を、S7 のノートの参考行と同じ「AGS.」にする（3.2in 短くなり、区切りで 2 行に入る）。
  S7: [['American Geriatrics Society. J Am Geriatr Soc 2015;63:142-150', 'AGS. J Am Geriatr Soc 2015;63:142-150']],
};
function sourcesOf(id) {
  const spec = SPEC[id] && SPEC[id].sources;
  if (!spec) return undefined;
  const r = trySources(id, spec, null);
  if (r) return r;
  if (SRC_ABBR[id]) {
    let s2 = spec;
    const what = [];
    for (const [from, to] of SRC_ABBR[id]) {
      if (!s2.includes(from)) throw new Error(`${id}: 出典に「${from}」がありません`);
      s2 = s2.replace(from, to);
      what.push(`「${from}」→「${to}」`);
    }
    const pre = `${what.join('、')}（著者名を略記。spec の文言のままでは区切りを詰めても 16pt・幅 11.6in の 2 行（23.2in）に入らず、` +
      '（）内の補足もないため。略記はノートの参考行と同じ形。文献は落としていない）';
    const r2 = trySources(id, s2, pre);
    if (r2) return r2;
  }
  throw new Error(`${id}: 出典が 2 行に収まりません`);
}
/**
 * 1 行の文字列に部分装飾を付けて run 配列にする。marks: [[語, 色, 追加の書式?], ...]（太字）。語は行の中に 1 回だけ現れること。
 * 警告（レッド）は色だけに頼らないよう下線も付ける（RED_U）。
 */
function mark(line, ...marks) {
  let runs = [line];
  for (const [w, color, extra] of marks) {
    const next = [];
    let found = 0;
    for (const r of runs) {
      if (typeof r !== 'string') { next.push(r); continue; }
      const k = r.indexOf(w);
      if (k < 0) { next.push(r); continue; }
      found++;
      if (k > 0) next.push(r.slice(0, k));
      next.push({ text: w, color, bold: true, ...(extra || {}) });
      if (k + w.length < r.length) next.push(r.slice(k + w.length));
    }
    if (found !== 1) throw new Error(`mark: 「${line}」に「${w}」が ${found} 回`);
    runs = next;
  }
  return runs;
}
const runText = l => (Array.isArray(l) ? l.map(r => (typeof r === 'string' ? r : r.text)).join('') : String(l));

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const Y0 = LAYOUT.body.y;        // 1.85
const Y0_T = 1.82;               // 詰めるスライドの上端（金下線の下端 1.576in から 0.244in）
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const r3 = v => Math.round(v * 1000) / 1000;
/**
 * 本文の下端。全スライドで出典が 2 行（枠 y6.5・下揃え。PowerPoint のメイリオ 16pt 2 行で文字の上端 約6.68in）
 * なので、その 0.24in 上。
 */
const BOTTOM = 6.44;

// ---------------------------------------------------------------------------
// 行送りを詰めたテキスト（カードと強調ボックス）。22pt 以上のまま、1 行ずつ decklib の s.text（上揃え）を
// 行送り em おきに置く。枠の高さは decklib / lint の 1 行の下限（22pt 0.41in、24pt 0.46in）で、
// 次の行の枠と少し重なる（文字どうしの枠の重なりは lint・decklib とも対象外）。
// ---------------------------------------------------------------------------
const lineIn = (size, em) => size * em / 72;
const BOX_H = { 22: 0.41, 24: 0.46 };
/**
 * o.hang: 行頭が全角「（」の行だけ、枠を左へ hang（in）ずらして広げる（行頭の全角括弧は字面が字送りの右半分に
 *         あるため、そのままだと半字ぶん字下げしたように見える。文言は変えず、位置で左端をそろえる）。
 * o.visRight: 字面の右端の上限（in。カードの右端－余白）。推定幅を実際のメイリオ相当（÷1.05）に直して検査する。
 */
function tightText(s, lines, o) {
  const size = o.size || 22;
  const em = o.em || 1.2;
  if (!BOX_H[size]) throw new Error(`tightText: size ${size} は 22 / 24 のみ`);
  lines.forEach((line, i) => {
    const hang = o.hang && /^（/.test(runText(line)) ? o.hang : 0;
    const x = o.x - hang;
    if (o.visRight != null && x + measure(runText(line), size) / 1.05 > o.visRight + 1e-6) {
      throw new Error(`p${s.page}「${runText(line)}」: 字面の右端がカードの余白に入ります`);
    }
    s.text([line], {
      x: r3(x), y: r3(o.y + i * lineIn(size, em)), w: r3(o.w + hang), h: BOX_H[size],
      size, bold: !!o.bold, color: o.color || C.dark, align: o.align || 'left',
    });
  });
  return lines.length * lineIn(size, em);
}

// ---------------------------------------------------------------------------
// カード（S3・S4）: 角丸の塗り箱＋見出し（青）＋項目。項目の頭は濃いグレーの小さな丸（ネイティブ図形。
// 箇条書きスライドの行頭記号と同じ色）。オレンジは強調語だけに使う。
// 項目は語句の切れ目で 2〜3 行に割り、2 行目以降も丸の右にそろえる。文字・丸はすべて箱の内側。
// 余白: 見出しと丸の左端はカードの左端から 0.2in。上は見出しの文字枠まで 0.14in（字面まで約 0.2in）、
// 下は最終行の文字枠から 0.16in（字面から約 0.2in）。右は文字枠の右端がカードの右端から 0.05in で、
// 字面（推定幅を実際のメイリオ相当の ÷1.05 に直したもの）の右端はカードの右端から 0.2in 以上内側（tightText で検査）。
// 行頭が全角「（」の行は半字（0.15in）左へずらして左端をそろえる（tightText の hang）。
// ---------------------------------------------------------------------------
const CARD = { padL: 0.2, padR: 0.05, visPad: 0.2, dot: 0.1, dotGap: 0.06, hang: 0.15, size: 22, titleSize: 24, dotColor: C.dark };
/** 項目間を除いたカードの高さ。padTop は箱の上端から見出しの文字枠まで、padBottom は最終行の文字枠の下端から箱の下端まで */
function cardFixedH(c, k) {
  const nLines = c.items.reduce((a, it) => a + it.length, 0);
  return k.padTop + lineIn(CARD.titleSize, k.em) + k.titleGap + (nLines - 1) * lineIn(CARD.size, k.em) + BOX_H[CARD.size] + k.padBottom;
}
function drawCard(s, c, x, y, w, h, itemGap, k) {
  s.box({ x: r3(x), y: r3(y), w: r3(w), h: r3(h), fill: C.bg, radius: 0.08 });
  const ix = x + CARD.padL, right = x + w - CARD.padR, visRight = x + w - CARD.visPad;
  const tx = ix + CARD.dot + CARD.dotGap;
  tightText(s, [c.title], { x: ix, y: y + k.padTop, w: right - ix, size: CARD.titleSize, bold: true, color: c.titleColor, em: k.em, visRight });
  let yy = y + k.padTop + lineIn(CARD.titleSize, k.em) + k.titleGap;
  const L = lineIn(CARD.size, k.em);
  c.items.forEach(lines => {
    // 丸は 1 行目の文字の高さの中ほど（メイリオ 22pt で枠上端から約 0.2in）に置く
    s.circle({ cx: r3(ix + CARD.dot / 2), cy: r3(yy + k.dotY), d: CARD.dot, fill: CARD.dotColor });
    tightText(s, lines, { x: tx, y: yy, w: right - tx, size: CARD.size, em: k.em, hang: CARD.hang, visRight });
    yy += lines.length * L + itemGap;
  });
  const end = yy - itemGap - L + BOX_H[CARD.size];   // 最終行の文字枠の下端
  if (end > y + h - k.padBottom + 1e-6) {
    throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さ不足（本文の下端 ${r3(end)} > ${r3(y + h - k.padBottom)}）`);
  }
  return end;
}
/**
 * 2 列のカード。widths: 各カードの幅（in）。高さは y〜bottom。
 * 項目の間隔は左右のカードで同じ値にする（どちらのカードにも入る最大値。上限 maxGap、下限 minGap）。
 * 行数の少ないカードは下に余りが出る（上揃え）。
 */
function cardRow(s, cards, o) {
  const k = { em: o.em, padTop: o.padTop, padBottom: o.padBottom, titleGap: o.titleGap, dotY: o.dotY };
  const sumW = o.widths.reduce((a, b) => a + b, 0) + GAP * (cards.length - 1);
  if (sumW > BW + 1e-6) throw new Error(`p${s.page}: カード幅の合計 ${r3(sumW)} > ${r3(BW)}`);
  const h = r3(o.bottom - o.y);
  const g = Math.floor(Math.min(o.maxGap, ...cards.map(c => {
    const free = h - cardFixedH(c, k);
    if (free < -1e-9) throw new Error(`p${s.page} カード「${runText(c.title)}」: 高さが足りません（h=${h}、必要 ${r3(cardFixedH(c, k))}）`);
    return c.items.length > 1 ? free / (c.items.length - 1) : Infinity;
  })) * 1000) / 1000;
  if (g < o.minGap) throw new Error(`p${s.page}: 項目間 ${g}in < ${o.minGap}in（行数を減らすか行送りを見直す）`);
  let x = X0;
  const ends = cards.map((c, i) => { const e = drawCard(s, c, x, o.y, o.widths[i], h, g, k); x += o.widths[i] + GAP; return e; });
  console.log(`  p${s.page} カード: 幅 ${o.widths.join(' / ')}in、${o.y}〜${o.bottom}in（高さ ${h}）、行送り ${o.em} em、` +
    `行数 ${cards.map(c => c.items.reduce((a, it) => a + it.length, 0)).join(' / ')}、項目間 ${g}in、` +
    `下余白（文字枠から） ${ends.map(e => r3(o.bottom - e)).join(' / ')}in`);
  return { h, gap: g };
}

/**
 * 行頭記号（•）の色を本文色 262626 に固定する（buClr）。pptxgenjs は buClr を出せず、記号の色が行頭の run の色
 * （S7「ベンゾジアゼピン」の橙、S5「高齢者評価：」の青など）になってしまうため、書き出した XML に足す。
 * deck.pptx.write を包み、decklib の後処理（postProcess）の前に適用する。buClr は buSzPct の直前（スキーマの順）。
 * （4-7 の build.js と同じ処理）
 */
function fixBulletColor(deck) {
  const JSZip = require('jszip');
  const write = deck.pptx.write.bind(deck.pptx);
  deck.pptx.write = async opts => {
    const zip = await JSZip.loadAsync(await write(opts));
    let n = 0;
    for (const name of Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))) {
      const x = await zip.file(name).async('string');
      const y = x.replace(/<a:buSzPct val="100000"\/><a:buChar /g, m => { n++; return `<a:buClr><a:srgbClr val="${C.dark}"/></a:buClr>${m}`; });
      zip.file(name, y);
    }
    if (!n) throw new Error('fixBulletColor: 行頭記号が見つかりません');
    console.log(`  行頭記号の色を ${C.dark} に固定（${n} 段落）`);
    return zip.generateAsync({ type: 'nodebuffer' });
  };
}

/** 強調ボックス（白地＋ティールの枠。文字は太字・中央）。decklib の box（通常の行送り）。 */
const EMPH = { radius: 0.1, lineW: 3 };
function emphBox(s, o) {
  s.box({
    x: r3(o.x != null ? o.x : X0), y: r3(o.y), w: r3(o.w != null ? o.w : BW), h: o.h, fill: C.white, radius: EMPH.radius,
    line: { color: o.lineColor || C.teal, width: EMPH.lineW },
    text: o.lines, size: o.size, bold: true, color: C.dark, margin: o.margin,
  });
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['フレイル', '認知\n機能', 'せん妄', '薬と\n麻酔', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-11 高齢者とフレイル・認知機能',
  });
  fixBulletColor(deck);
  /** 本文スライド。タイトルと章は spec の見出し行（## Sn  ch=…  title=「…」）から取る */
  const slide = id => {
    const sp = SPEC[id];
    if (!sp.title || sp.ch == null) throw new Error(`${id}: spec の見出し行に title / ch がありません`);
    return deck.slide({ title: sp.title, chapter: sp.ch, notes: notesOf(id), sources: sourcesOf(id) });
  };
  const VER = C.vermilion, BLUE = C.blue, RED = C.red;
  const RED_U = { underline: true };   // 警告のレッドは下線も付け、色だけに頼らない（同じスライドの他の強調は青）

  // ===== S1 表紙 =====
  {
    const subtitle = '4-11 術前評価', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: SPEC.S1.title, subtitle, meta, notes: notesOf('S1') });
  }

  // ---------------------------------------------------------------------------
  // 箇条書き＋強調ボックス（S2・S5〜S9）
  //  5 行: 箇条書き y1.9（24pt・段落後 10pt）、強調ボックス 28pt 太字 2 行（高さ 1.2in、下端 BOTTOM-0.04）。
  //  6 行: 箇条書き y1.85、強調ボックス 24pt 太字 1 行（高さ 0.58in、下端 BOTTOM）。
  //  どちらも箇条書きの枠の下端と強調ボックスの間を 0.24in 以上あける（足りなければ Error）。
  // ---------------------------------------------------------------------------
  const BUL5 = { y: 1.9, size: 24, gap: 10 };
  const BUL6 = { y: 1.85, size: 24, gap: 10 };
  const EMPH2_28 = 1.2;     // 28pt 2 行
  const EMPH1_24 = 0.58;    // 24pt 1 行
  function bulletsWithEmph(s, id, items, emphLines) {
    const six = items.length === 6;
    const B = six ? BUL6 : BUL5;
    if (!six && items.length !== 5) throw new Error(`${id}: 箇条書きは 5 か 6 行を想定（${items.length} 行）`);
    s.bullets(items, B);
    const bulBottom = B.y + heightFor(items.length, B.size, B.gap);
    const eh = six ? EMPH1_24 : EMPH2_28;
    const ey = r3(six ? BOTTOM - eh : BOTTOM - 0.04 - eh);
    if (ey < bulBottom + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(ey - bulBottom)}in`);
    if (emphLines.length !== (six ? 1 : 2)) throw new Error(`${id}: 強調ボックスの行数 ${emphLines.length}`);
    emphBox(s, { y: ey, h: eh, size: six ? 24 : 28, lines: emphLines });
    console.log(`  p${s.page} 箇条書き ${items.length} 行 ${B.y}〜${r3(bulBottom)}in、強調ボックス ${ey}〜${r3(ey + eh)}in（間 ${r3(ey - bulBottom)}in）`);
  }

  // ===== S2 フレイルが予後を決める（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S2';
    const s = slide(id);
    const lines = brk(id, 'ASA-PS や心リスク指数にフレイルを足すと、予測が良くなる', '予測');
    bulletsWithEmph(s, id, [
      mark(q(id, 'CFS でフレイル → 全死亡のオッズ比 4.0（メタ解析）'), ['オッズ比 4.0', VER]),
      mark(q(id, '1 年死亡 13.6% 対 4.8%（診断コード定義）、術直後ほど差が大きい'), ['13.6% 対 4.8%', VER]),
      mark(q(id, '合併症 OR 2.5、在院 1.7 倍、施設退院 OR 20（Makary 2010）'), ['施設退院 OR 20', VER]),
      mark(q(id, '術後せん妄 OR 2.1（メタ解析）'), ['OR 2.1', VER]),
      mark(q(id, 'CFS 4 以上で注意、5 以上をフレイル（閾値は研究で違う）'), ['4 以上で注意', BLUE], ['5 以上をフレイル', BLUE]),
    ], [mark(lines[0], ['フレイルを足す', VER]), lines[1]]);
  }

  // 2 列カード＋強調ボックス（1 行・24pt）の共通寸法（S3・S4）
  //  余白を上下左右とも字面から約 0.2in とるため、行送りは 1.1 em（22pt のまま）。8 行のカード（S3 右、S4 左右）で
  //  項目間は約 0.08in（項目間の字面の隙間は行間の約 2 倍）。左右のカードで同じ値にする。
  //  padBottom は最終行の文字枠の下端から（メイリオの字面の下端は文字枠の下端の約 0.05in 上）
  const KC = { em: 1.1, padTop: 0.14, padBottom: 0.16, titleGap: 0.06, dotY: 0.2, maxGap: 0.1, minGap: 0.07 };
  const halfW = r3((BW - GAP) / 2);   // 5.995

  // ===== S3 フレイルの測り方（2列カード + 強調ボックス） =====
  {
    const id = 'S3';
    const s = slide(id);
    const ey = r3(BOTTOM - EMPH1_24);
    emphBox(s, { y: ey, h: EMPH1_24, size: 24,
      lines: [mark(q(id, '転倒歴、眼鏡・補聴器、独居・介護者、本人の治療の目標も聞く'), ['本人の治療の目標', VER])] });
    cardRow(s, [
      {
        title: q(id, 'CFS（9 段階）'), titleColor: BLUE,
        items: [
          [mark(q(id, '約 2 週間前の病前の状態で付ける'), ['病前', VER])],
          brk(id, '移動、生活機能（IADL・ADL）、思考、健康の自己評価を聞く', '思考'),
          [mark(q(id, '迷ったら重い方'), ['重い方', VER])],
          brk(id, '若年者・固定した単一障害には使わない', '使わない'),
        ],
      },
      {
        title: q(id, 'J-CHS 基準（2020）と周辺'), titleColor: BLUE,
        items: [
          [q(id, '握力 男 28 kg 未満・女 18 kg 未満')],
          [q(id, '歩行速度 1.0 m/秒未満')],
          (() => { const l = brk(id, '体重減少、疲労感、身体活動の低下（3 項目以上でフレイル）', '（3'); return [l[0], mark(l[1], ['3 項目以上でフレイル', VER])]; })(),
          [q(id, '基本チェックリスト 8 点以上')],
          brk(id, '栄養：BMI 18.5 未満、Alb 3.0 g/dL 未満、6 か月で 10〜15% 超の体重減少', 'Alb', '6 か月'),
        ],
      },
    ], { ...KC, y: Y0_T, bottom: r3(ey - GAP), widths: [halfW, halfW] });
  }

  // ===== S4 Mini-Cog を全員に（2列カード + 強調ボックス） =====
  // 左右とも幅 5.995in（S3 と同じ）。右のカードは「せん妄 21% 対 7%、自宅以外への退院 67% 対 34%（Culley 2017）」を
  // 「退院」の前で割り、2 行に収める（3 行に割ると 9 行になり、強調ボックスの上に入らない）。
  // 左のカードの「（ASA Brain Health の専門家合意）」（全角換算 16.74 字）は行頭の全角括弧を半字左へずらして 1 行に入れる。
  {
    const id = 'S4';
    const s = slide(id);
    const ey = r3(BOTTOM - EMPH1_24);
    // 24pt で全幅 1 行（推定 12.10in）に入れるため、箱の内側余白を 0.05in にする
    emphBox(s, { y: ey, h: EMPH1_24, size: 24, margin: 0.05,
      lines: [mark(q(id, '65 歳を超える全員に『術後に混乱・記憶障害が起こりうる』と説明する'), ['『術後に混乱・記憶障害が起こりうる』', VER])] });
    cardRow(s, [
      {
        title: q(id, 'どう測るか'), titleColor: BLUE,
        items: [
          brk(id, '65 歳超の全員にベースラインの簡易スクリーニング（ASA Brain Health の専門家合意）', '簡易', '（ASA'),
          (() => { const l = brk(id, 'Mini-Cog：3 語の遅延再生＋時計描画、2〜4 分、2 点以下で陽性', '時計描画'); return [mark(l[0], ['Mini-Cog：', BLUE]), mark(l[1], ['2 点以下で陽性', VER])]; })(),
          brk(id, 'MMSE 7〜10 分（23 点以下）、HDS-R（20 点以下）', 'HDS-R'),
          [q(id, '眼鏡・補聴器をつけ、前投薬の前に')],
        ],
      },
      {
        title: q(id, '陽性なら'), titleColor: BLUE,
        items: [
          [mark(q(id, '診断ではない。詳しい評価へつなぐ'), ['診断ではない', VER])],
          brk(id, '認知症の診断がない人工関節患者の 24% が陽性', '人工関節'),
          brk(id, 'せん妄 21% 対 7%、自宅以外への退院 67% 対 34%（Culley 2017）', '退院'),
          brk(id, '家族に周術期神経認知障害のリスクを説明', 'リスク'),
          [q(id, '意思決定能力を確かめる（→ 4-5）')],
        ],
      },
    ], { ...KC, y: Y0_T, bottom: r3(ey - GAP), widths: [halfW, halfW] });
  }

  // ===== S5 術後せん妄：リスク因子（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S5';
    const s = slide(id);
    const em = brk(id, '低活動型は見逃される。回復室を出る前と病棟でスクリーニングする', '回復室');
    bulletsWithEmph(s, id, [
      mark(q(id, '高齢、ASA-PS 3 以上、Charlson 2 超、MMSE 25 未満（ESAIC）'), ['MMSE 25 未満', BLUE]),
      mark(q(id, '高齢者評価：フレイル、感覚障害、低栄養、多剤、貧血、社会的因子'), ['高齢者評価：', BLUE]),
      mark(q(id, '国内の加算：せん妄の既往、アルコール多飲、ベンゾジアゼピン'), ['国内の加算：', BLUE]),
      mark(q(id, '全身麻酔の手術は年齢を問わず対象（1 項目以上でハイリスク）'), ['年齢を問わず対象', BLUE]),
      mark(q(id, '骨折 35〜65%、大手術 20〜40% に起き、半数以上が見逃される'), ['35〜65%', BLUE], ['20〜40%', BLUE], ['半数以上が見逃される', RED, RED_U]),
    ], [mark(em[0], ['低活動型', BLUE]), em[1]]);
  }

  // ===== S6 術後せん妄：予防（箇条書き6行 + 強調ボックス） =====
  {
    const id = 'S6';
    const s = slide(id);
    bulletsWithEmph(s, id, [
      mark(q(id, '多因子の非薬物的介入（AGS 2015・ESAIC 2024 とも強い推奨）'), ['多因子の非薬物的介入', BLUE]),
      q(id, '見当識、睡眠、早期離床、眼鏡・補聴器、水分・栄養、鎮痛'),
      mark(q(id, '薬の見直し、酸素化、便秘。せん妄のオッズ 44% 減（メタ解析）'), ['44% 減', BLUE]),
      mark(q(id, '老年科医の関与でせん妄 32% 対 50%（大腿骨近位部骨折）'), ['32% 対 50%', BLUE]),
      mark(q(id, '予防薬は勧めない（ESAIC）。抗精神病薬の予防投与は根拠不十分'), ['予防薬は勧めない', RED, RED_U]),
      mark(q(id, 'コリンエステラーゼ阻害薬を予防のために新しく始めない（AGS 強）'), ['新しく始めない', RED, RED_U]),
    ], [mark(q(id, '薬で予防しない。入院初日から多職種で多因子介入を始める'), ['入院初日から', BLUE])]);
  }

  // ===== S7 高齢者で避ける薬（箇条書き6行 + 強調ボックス） =====
  {
    const id = 'S7';
    const s = slide(id);
    bulletsWithEmph(s, id, [
      mark(q(id, 'ベンゾジアゼピン（ミダゾラム、ジアゼパム）'), ['ベンゾジアゼピン', VER]),
      mark(q(id, '抗コリン薬：スコポラミン、アトロピン、プロメタジン'), ['抗コリン薬：', VER]),
      mark(q(id, 'ジフェンヒドラミン、ヒドロキシジン、H2 遮断薬（ファモチジン）'), ['H2 遮断薬', VER]),
      mark(q(id, 'メトクロプラミド、ペチジン（シバリングも加温が先）'), ['ペチジン', VER]),
      mark(q(id, 'ハロペリドールの予防投与・ルーチン投与'), ['予防投与・ルーチン投与', VER]),
      mark(q(id, '例外：徐脈へのアトロピン、離脱予防のベンゾジアゼピン'), ['例外：', BLUE]),
    ], [mark(q(id, '5〜6 剤以上は転倒・有害事象が増える。お薬手帳で市販薬まで確認'), ['5〜6 剤以上', VER], ['お薬手帳', BLUE])]);
  }

  // ===== S8 高齢者の麻酔の調節（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S8';
    const s = slide(id);
    const em = brk(id, '若年成人と同じ設定・用量にしない。少量から、効果を待って追加', '少量から');
    bulletsWithEmph(s, id, [
      mark(q(id, 'MAC は 10 歳ごとに約 6% 低下（80 歳のセボフルラン約 1.4%）'), ['約 6% 低下', VER]),
      mark(q(id, '年齢補正 MAC 比を見る。BIS は高く出ることがある'), ['年齢補正 MAC 比', BLUE]),
      mark(q(id, 'プロポフォール・ミダゾラムは減量し、ゆっくり（感受性が高い）'), ['減量し、ゆっくり', VER]),
      mark(q(id, 'ベースラインに対する相対的低血圧を避ける。正常体温を保つ'), ['相対的低血圧', VER]),
      mark(q(id, '鎮静の少ない多角的鎮痛。区域麻酔そのものではせん妄は減らない'), ['鎮静の少ない多角的鎮痛', BLUE]),
    ], [em[0], mark(em[1], ['少量から', VER])]);
  }

  // ===== S9 相談先と離島での工夫（箇条書き5行 + 強調ボックス） =====
  {
    const id = 'S9';
    const s = slide(id);
    const em = brk(id, '老年科医がいない島では、チェックリストと多職種で『老年科の目』を代わりに入れる', '『老年科の目』');
    bulletsWithEmph(s, id, [
      mark(q(id, '老年科・認知症ケアチーム・精神科リエゾンに術前から（遠隔でも）'), ['術前から', VER]),
      mark(q(id, '家族に：眼鏡・補聴器・時計・慣れた物を持参し、付き添いを'), ['付き添い', BLUE]),
      mark(q(id, '看護師・薬剤師・リハビリと多因子介入を入院初日から'), ['入院初日から', VER]),
      mark(q(id, '術前から栄養と運動（プレハビリテーション）、飲酒・睡眠薬の整理'), ['栄養と運動', BLUE]),
      mark(q(id, '退院先と治療の目標を本人・家族と前もって話す（→ 4-6）'), ['前もって話す', VER]),
    ], [em[0], mark(em[1], ['『老年科の目』', VER])]);
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const id = 'S10';
    const s = slide(id);
    s.bullets([
      mark(q(id, '年齢だけでなく CFS を病前で。4 以上は注意、5 以上はフレイル'), ['CFS を病前で', BLUE]),
      mark(q(id, '65 歳超は全員 Mini-Cog。陽性なら家族に説明し能力を確かめる'), ['全員 Mini-Cog', BLUE]),
      mark(q(id, 'せん妄は大手術で 20〜40%、半数以上を見逃す。多因子介入で予防'), ['多因子介入で予防', BLUE]),
      mark(q(id, '薬で予防しない。ベンゾジアゼピン・抗コリン薬・ペチジンを避ける'), ['予防しない', RED, RED_U]),
      mark(q(id, 'MAC は 10 歳で 6% 低下。相対的低血圧を避け、鎮静の少ない鎮痛'), ['10 歳で 6% 低下', BLUE]),
    ], { y: 2.0, size: 24, gap: 32 });
  }

  // spec の全スライドを組んだか（S1〜S10）
  const ids = Object.keys(SPEC);
  if (ids.length !== 10 || deck.slides.length !== ids.length) throw new Error(`spec ${ids.length} 枚 / deck ${deck.slides.length} 枚`);

  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
