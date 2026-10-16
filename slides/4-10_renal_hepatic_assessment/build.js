'use strict';
/**
 * 4-10 腎不全・透析と肝疾患の評価
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言はここに転記（改行位置のみ調整）、ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 2 列カード（S3・S6・S7）は 1 項目を語句の切れ目で 2〜3 行に割るため、項目内の行送りを 1.2 em に詰め
 * （1 行 1 テキスト枠。tightText）、項目の頭に「•」を置いて項目の切れ目を示す。左右で行数が違うカードは、
 * 行の少ない方の項目の間隔を広げて、両方の最終行を箱の下端にそろえる（fill）。
 * 行の多い表（S4・S5）は 3-11 の tightTable（行間を固定値に詰めたネイティブ表）を写して使う。
 * S4 は 6 行すべてが 2 行になり、1.2 em では出典 2 行の上に入らないので、行間 24pt。セルの余白は上だけに
 * 取る（固定行間では字面が行の上に寄り、余白を上下に分けると上の罫線に字が接するため）。
 * S8 の箇条書きは行頭の「•」を別の文字枠にして本文色にそろえる（箇条書きの記号は先頭の文字の色になるため）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, FONT, LAYOUT, fits, fitInfo } = require('../_lib/decklib');

const NAME = '4-10_renal_hepatic_assessment';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

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
 * 出典。spec の sources 行を ' / ' で項目に分けてライブラリに渡す（ライブラリが項目の区切りで 1〜2 行に詰め、
 * 区切りは ' / '）。1 文献を行の途中で折らない。
 * 文献の区切りでは 2 行に収まらないスライドだけ、次の扱いにする（deviations）:
 *   short: 各文献を「筆頭著者の姓 et al. 誌名 年」に略記する（巻・頁と著者のイニシャルを省く。
 *          （AGA）（VOCAL-Penn）の補足は残す）。S5〜S7 は、どの 2 文献も spec の表記のままでは
 *          16pt・11.6in の 1 行に並ばない（最短の組でも S5 14.50in、S6 12.61in、S7 14.29in）。
 *          S5 は巻・頁を省いてもイニシャルを残すと 1 行目が 12.41in になるため、イニシャルも省く
 *          （3 枚で表記をそろえる）。全文は spec.md の sources 行（S5・S6 はノートの「参考」にもある）。
 *   omit:  S8 は略記しても 3 文献が 2 行に入らない（Fielding-Singh と Northup の組で 12.69in）ため、
 *          S2・S3・S7 にも出る Fielding-Singh 2025 を落とす。
 */
const SRC_SHORT = {
  'Teh SH, et al. Gastroenterology 2007;132:1261-1269': 'Teh et al. Gastroenterology 2007',
  'Northup PG, et al. Clin Gastroenterol Hepatol 2019;17:595-606（AGA）': 'Northup et al. Clin Gastroenterol Hepatol 2019（AGA）',
  'Mahmud N, et al. Hepatology 2021;73:204-218（VOCAL-Penn）': 'Mahmud et al. Hepatology 2021（VOCAL-Penn）',
  "O'Leary JG, et al. Gastroenterology 2019;157:34-43（AGA）": "O'Leary et al. Gastroenterology 2019（AGA）",
  'Stöhr T, et al. Br J Anaesth 2021;127:415-423': 'Stöhr et al. Br J Anaesth 2021',
  'Khalil M, et al. Anesthesiology 1994;80:1241-1247': 'Khalil et al. Anesthesiology 1994',
  'Fielding-Singh V, et al. Anesthesiology 2025;143:1030-1048': 'Fielding-Singh et al. Anesthesiology 2025',
  'Neeff H, et al. J Gastrointest Surg 2011;15:1-11': 'Neeff et al. J Gastrointest Surg 2011',
};
const SRC_FIX = {
  S5: { short: true },
  S6: { short: true },
  S7: { short: true },
  S8: { omit: ['Fielding-Singh V, et al. Anesthesiology 2025;143:1030-1048'] },
};
const DEVIATIONS = [];
/** 項目の並びを、ライブラリと同じく項目の区切りで詰めた行にする（検査とログ用） */
function packSources(items) {
  const S = LAYOUT.source, lines = [];
  for (const it of items) {
    const cur = lines.length ? lines[lines.length - 1] + ' / ' + it : null;
    if (cur && fits(cur, S.size, S.w)) lines[lines.length - 1] = cur;
    else lines.push(it);
  }
  return lines;
}
function sourcesOf(id) {
  const raw = SPEC[id] && SPEC[id].sources;
  if (!raw) throw new Error(`spec.md の ${id} に sources がありません`);
  let items = raw.split(' / ');
  const f = SRC_FIX[id];
  if (f && f.omit) {
    f.omit.forEach(o => { if (!items.includes(o)) throw new Error(`${id} 出典に「${o}」がありません`); });
    items = items.filter(it => !f.omit.includes(it));
    DEVIATIONS.push(`${id} 出典: 「${f.omit.join('」「')}」を省略（略記しても 3 文献が 2 行に入らない）`);
  }
  if (f && f.short) {
    if (packSources(items).length <= 2) throw new Error(`${id} 出典: 略記しなくても 2 行に入ります`);
    items = items.map(it => {
      const a = SRC_SHORT[it];
      if (!a) throw new Error(`${id} 出典: 略記が未登録「${it}」`);
      // 略記は「姓」「誌名 年」「（補足）」を元の表記から引き継いでいること
      const sur = it.split(/[ ,]/)[0], ym = it.match(/ ((?:[A-Z][A-Za-z]* )*[A-Z][A-Za-z]*) (\d{4});/);
      const note = (it.match(/（[^）]*）$/) || [''])[0];
      if (!ym || !a.startsWith(sur + ' et al. ') || !a.includes(`${ym[1]} ${ym[2]}`) || !a.endsWith(note || ym[2])) {
        throw new Error(`${id} 出典: 略記「${a}」が元の「${it}」と対応しません`);
      }
      return a;
    });
    DEVIATIONS.push(`${id} 出典: 筆頭著者の姓 et al.・誌名・年に略記 → ${packSources(items).join(' ｜ ')}`);
  }
  const lines = packSources(items);
  if (lines.length > 2) throw new Error(`${id} 出典が文献の区切りで 2 行に収まりません: ${lines.join(' ｜ ')}`);
  return items;
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const TOP = 1.85;                // 本文の上端
const TOP_T = 1.82;              // 表・詰まったカードの上端（金下線 8pt の下端 1.576in から 0.24in 以上。4-12 と同じ）
const ZONE_BOT = 6.47;           // 表・詰まったカード・強調ボックスの下端の上限（出典 2 行の枠 y6.5 の 0.03in 上
                                 //  ＝decklib の警告の閾値。メイリオ 16pt 2 行の字面の上端 約6.73in から 0.26in）
const BOT = 6.44;                // 出典 2 行（枠 y6.5・下揃え。メイリオ 16pt 2 行で文字の上端 約6.68in）の 0.24in 上
const GAP = 0.26;                // 要素間（0.24in 以上）
const em = (text, color = C.vermilion, bold = true) => ({ text, color, bold });
const r3 = v => Math.round(v * 1000) / 1000;

// ---------------------------------------------------------------------------
// 行送りを詰めたテキスト（カードと 2 行の強調ボックス）。22pt 以上のまま行送りを 1.2 em にする
// （3-11 の tightTable と同じ考え方。4-13 と同じ組み方）。
// 段落の行間（spcPct / spcPts）は使わず、1 行ずつ decklib の s.text（上揃え）を 1.2 em おきに置く。
//  - spcPts の固定値は slide_lint の高さ推定を壊し、spcPct は代替フォント（LibreOffice の描画）で
//    行送りが変わる。1 行 1 枠なら行の位置が枠の y で決まり、PowerPoint でも描画確認でも同じ行送りになる。
//  - 枠の高さは decklib / lint の 1 行の下限（22pt 0.41in、24pt 0.46in）で、次の行の枠と少し重なる
//    （文字どうしの枠の重なりは lint・decklib とも対象外）。
// lines: 行の配列（行は文字列 or run 配列）。戻り値: 行送りぶんの高さ
// ---------------------------------------------------------------------------
const LS_EM = 1.2;
const lineIn = size => size * LS_EM / 72;            // 行送り（22pt 0.367in、24pt 0.4in）
const BOX_H = { 22: 0.41, 24: 0.46 };                // 1 行の枠の高さ
const tightH = (n, size) => n * lineIn(size);
function tightText(s, lines, o) {
  const size = o.size || 22;
  if (!BOX_H[size]) throw new Error(`tightText: size ${size} は 22 / 24 のみ`);
  lines.forEach((line, i) => {
    s.text([line], {
      x: r3(o.x), y: r3(o.y + i * lineIn(size)), w: r3(o.w), h: BOX_H[size],
      size, bold: !!o.bold, color: o.color || C.dark, align: o.align || 'left',
    });
  });
  return tightH(lines.length, size);
}

// ---------------------------------------------------------------------------
// 見出し＋項目のカード（項目の頭に「•」。1 項目を割った行は行間を詰め、項目の間だけ空ける）
// card: {title, titleColor, items: [[行, 行...], ...]}
// ---------------------------------------------------------------------------
const CARD = { padX: 0.1, padY: 0.1, titleSize: 24, tgap: 0.08, size: 22, indent: 0.27 };
function cardFixed(c) {
  const nLines = c.items.reduce((a, it) => a + it.length, 0);
  return 2 * CARD.padY + tightH(1, CARD.titleSize) + CARD.tgap + tightH(nLines, CARD.size);
}
function itemCard(s, c, o) {
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.bg, radius: 0.08 });
  const tx = o.x + CARD.padX, tw = o.w - 2 * CARD.padX;
  let y = o.y + CARD.padY;
  y += tightText(s, [c.title], { x: tx, y, w: tw, size: CARD.titleSize, bold: true, color: c.titleColor });
  y += CARD.tgap;
  c.items.forEach((it, i) => {
    // 「•」は推定器では全角 1 字（22pt で 0.342in）。枠は項目の文字枠に重ねて広めに取る（文字どうしの重なりは無害）
    tightText(s, ['•'], { x: tx, y, w: 0.36, size: CARD.size, color: c.titleColor });
    y += tightText(s, it, { x: tx + CARD.indent, y, w: tw - CARD.indent, size: CARD.size });
    if (i < c.items.length - 1) y += o.igap;
  });
  if (y + CARD.padY > o.y + o.h + 1e-6) throw new Error(`p${s.page} カード「${c.title}」: 高さ不足（必要 ${r3(y + CARD.padY - o.y)}in > ${o.h}in）`);
}
/**
 * 2 列カード。avail（上端 y からの高さ）に収まるよう項目の間隔 igap を決める（上限 igapMax）。
 * fill: true なら、行の少ないカードは項目の間隔を広げて最終行を箱の下端（余白 padY）にそろえる
 * （上限 fillMax。左右で中身の量が違っても、片方の下だけが空かないように）。
 * 戻り値: {h, igap, igaps}（カードの高さ、共通の間隔、カードごとの間隔）
 */
function cardRow(s, cards, o) {
  let igap = o.igapMax != null ? o.igapMax : 0.18;
  cards.forEach(c => {
    if (c.items.length > 1) igap = Math.min(igap, (o.avail - cardFixed(c)) / (c.items.length - 1));
  });
  if (igap < 0.06) throw new Error(`p${s.page}: カードの項目間隔が狭すぎます（${r3(igap)}in）`);
  const h = r3(Math.max(...cards.map(c => cardFixed(c) + igap * (c.items.length - 1))) + 1e-4);
  if (o.widths.reduce((a, b) => a + b, 0) + o.gap * (cards.length - 1) > BW + 1e-6) throw new Error(`p${s.page}: カードの幅の合計が本文幅を超えます`);
  const fillMax = o.fillMax != null ? o.fillMax : 0.4;
  const igaps = cards.map(c => (o.fill && c.items.length > 1)
    ? Math.min(fillMax, Math.max(igap, (h - 1e-4 - cardFixed(c)) / (c.items.length - 1)))
    : igap);
  let x = X0;
  cards.forEach((c, i) => {
    itemCard(s, c, { x, y: o.y, w: o.widths[i], h, igap: igaps[i] });
    x += o.widths[i] + o.gap;
  });
  return { h, igap, igaps: igaps.map(r3) };
}

/** 強調ボックス（白地＋ティールの枠、太字）。1 行は decklib の box、2 行は行間を詰めた文字を箱の内側に置く */
function callout(s, lines, o) {
  const size = o.size || 24;
  if (lines.length === 1) {
    s.box({ x: X0, y: o.y, w: BW, h: o.h, fill: C.white, line: { color: C.teal, width: 2.5 }, radius: 0.08, text: lines, size, bold: true, color: C.dark });
    return;
  }
  s.box({ x: X0, y: o.y, w: BW, h: o.h, fill: C.white, line: { color: C.teal, width: 2.5 }, radius: 0.08 });
  const th = tightH(lines.length, size);
  tightText(s, lines, { x: X0 + 0.2, y: o.y + (o.h - th) / 2, w: BW - 0.4, size, bold: true, align: 'center' });
}
const calloutH = (n, size) => (n === 1 ? r3(size / 72 * 1.5 + 0.24) : r3(tightH(n, size) + 0.16));

/**
 * 行の多い表（3-11_local_anesthetics/build.js の tightTable を写したもの）。decklib の s.table と同じ見た目
 * （見出し行 D9D9D9・太字、罫線 BFBFBF 1pt、上下中央）で、行間だけを固定値 LS pt（既定 1.2 em）に詰める。
 * セル幅の検査は decklib と同じ推定器（lint 式、全角 ×1.12）で行い、重なり検査用の登録も同じく行う。
 * cell: 文字列（'\n' で改行）| {text, color, bold, align}
 * o.round: 行の高さの切り上げ単位（既定 0.01in。S4 は 0.001in にして、余白を出典の上の限度まで使う）
 */
function tightTable(s, rows, o) {
  const size = o.size || 22, LS = o.lineSpacing || size * 1.2;
  const m = o.margin || [3, 6, 3, 6];   // pt: 上, 右, 下, 左
  // pptxgenjs は margin[0] が 1 未満だと 4 値すべてをインチとして扱うため、上の余白は 1pt 以上にする
  if (!(m[0] >= 1)) throw new Error(`p${s.page} 表: 上の余白 ${m[0]}pt は 1pt 以上にしてください`);
  const RND = o.round || 0.01;
  const colW = o.colW, nCol = colW.length;
  const w = colW.reduce((a, b) => a + b, 0);
  const rowH = [];
  const out = rows.map((row, ri) => {
    if (row.length !== nCol) throw new Error(`p${s.page} 表 行${ri + 1}: 列数 ${row.length} ≠ ${nCol}`);
    const head = ri === 0;
    let maxLines = 1;
    const cells = row.map((cell, ci) => {
      const c = typeof cell === 'string' ? { text: cell } : cell;
      const lines = c.text.split('\n');
      const avail = colW[ci] - (m[1] + m[3]) / 72;
      lines.forEach(l => {
        const fi = fitInfo(l, size, avail);
        if (!fi.ok) {
          throw new Error(`p${s.page} 表 行${ri + 1}列${ci + 1}「${l}」がセル幅に収まりません` +
            `（推定 ${fi.needIn}in > ${avail.toFixed(3)}in、約${fi.overChars}字オーバー）`);
        }
      });
      maxLines = Math.max(maxLines, lines.length);
      const color = c.color || C.dark, bold = c.bold != null ? c.bold : head;
      return {
        text: lines.map((l, li) => ({
          text: l,
          options: {
            fontFace: FONT, lang: 'ja-JP', fontSize: size, color, bold: bold || undefined,
            lineSpacing: LS, breakLine: li < lines.length - 1 || undefined,
          },
        })),
        options: {
          fill: head ? { color: C.lightGray } : undefined,
          align: c.align || 'left', valign: 'middle',
          fontFace: FONT, fontSize: size, color, bold: bold || undefined,
        },
      };
    });
    rowH.push(r3(Math.ceil((maxLines * LS / 72 + (m[0] + m[2]) / 72) / RND - 1e-6) * RND));
    return cells;
  });
  const h = rowH.reduce((a, b) => a + b, 0);
  s._bounds(`p${s.page} 表`, o.x, o.y, w, h);
  s.raw.addTable(out, {
    x: o.x, y: o.y, w, colW, rowH, fontFace: FONT, fontSize: size, margin: m,
    border: { type: 'solid', pt: 1, color: C.border }, valign: 'middle',
  });
  s._reg('box', o.x, o.y, w, h, 'table');
  return { x: o.x, y: o.y, w, h: r3(h), rowH };
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['透析', 'CKD\nと薬', '肝疾患', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '4-10 腎不全・透析と肝疾患の評価',
  });

  // ===== S1 表紙 =====
  deck.cover({
    title: '腎不全・透析と肝疾患の評価',
    subtitle: '4-10 術前評価',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 透析患者で確かめること =====
  {
    const s = deck.slide({ title: '透析患者で確かめること', chapter: 0, notes: notesOf('S2'), sources: sourcesOf('S2') });
    s.bullets([
      ['透析は', em('手術の前日'), 'に。間隔が長いほど術後 90 日死亡が増える'],
      '当日の透析直後は、除水による低血圧と回路のヘパリンに注意',
      'ドライウェイトと溢水の所見（浮腫、ラ音、起座呼吸、頸静脈怒張）',
      '透析の抗凝固（ヘパリンかナファモスタットか）と量を聞く',
      '残っている尿量、最近の透析中の低血圧、体重増加',
    ], { y: TOP, size: 24, gap: 18 });
    const bh = calloutH(1, 26);
    callout(s, [['腎臓内科・透析室と、透析の日と抗凝固を', em('前もって決める')]], { y: r3(BOT - bh), h: bh, size: 26 });
  }

  // ===== S3 カリウムとシャント =====
  {
    const s = deck.slide({ title: 'カリウムとシャント', chapter: 0, notes: notesOf('S3'), sources: sourcesOf('S3') });
    // 強調ボックス（2 行）は calloutH の 0.893in より上下に余裕を持たせて 0.95in。カードの上端を TOP_T、
    // 強調ボックスの下端を ZONE_BOT にして、右カード（7 行）の項目間隔 0.07in を保つ
    const bh = 0.95;
    if (bh < calloutH(2, 22)) throw new Error('S3 強調ボックスが低すぎます');
    const by = r3(ZONE_BOT - bh);
    const cw = r3((BW - GAP) / 2);
    // 左（カリウム 5 行）は右（シャント肢 7 行）より短いので、項目の間隔を広げて下端をそろえる（fill）
    const { h, igaps } = cardRow(s, [
      {
        title: 'カリウム', titleColor: C.blue, items: [
          [[em('5.5 mEq/L 超'), 'は延期して'], '透析か補正を考える（慣用の目安）'],
          ['透析直後の低値はすぐに補充しない'],
          ['週末明けは高 K が 2 倍以上'],
          ['心電図（T 波、QRS 幅）を必ず見る'],
        ],
      },
      {
        title: 'シャント肢', titleColor: C.vermilion, items: [
          ['血圧測定・採血・点滴・', '動脈ラインをしない'],
          ['体位で圧迫しない。', '術中・術後にスリルを確認'],
          ['鎖骨下静脈の', '中心静脈カテーテルは避ける'],
          ['透析用カテーテルは原則使わない'],
        ],
      },
    ], { y: TOP_T, avail: by - 0.24 - TOP_T, widths: [cw, cw], gap: GAP, fill: true });
    if (TOP_T + h > by - 0.24 + 1e-6) throw new Error('S3 カードと強調ボックスの間隔が足りません');
    if (igaps[1] < 0.07 - 1e-6) throw new Error(`S3 右カードの項目間隔 ${igaps[1]}in < 0.07in`);
    console.log(`  p${s.page} カード ${TOP_T}〜${r3(TOP_T + h)}in、項目間隔 左 ${igaps[0]} / 右 ${igaps[1]}in、強調ボックス ${by}〜${r3(by + bh)}in`);
    callout(s, [[em('K 5.5 超でも緊急なら進む：')], ['心電図・緊急度・補正の手段で個別に判断し、相談する']], { y: by, h: bh, size: 22 });
  }

  // ===== S4 腎機能で変わる薬と造影剤 =====
  {
    const s = deck.slide({ title: '腎機能で変わる薬と造影剤', chapter: 1, notes: notesOf('S4'), sources: sourcesOf('S4') });
    const drug = text => ({ text, bold: true });
    // 6 行すべてが 2 行になるため、行間 24pt に詰める（spec の列幅 2.6/5.0/4.6 は、
    // 「ガバペンチノイド、」と「避ける。フェンタニル主体に」が 1 行に入るよう 3.25/4.33/4.62 に変更）。
    // 固定行間 24pt（22pt の字）では字面が行の上に寄る（下は字の下に 約3.7pt の空きが残る）ため、
    // セルの余白は上だけに 3.1pt 取り、上の罫線から字を離す。上端は TOP_T、行の高さは 0.001in 単位で
    // 切り上げ、下端は出典 2 行の上の限度 ZONE_BOT（6.47in）に収める。
    const t = tightTable(s, [
      ['薬', '何が起きるか', 'どうする'],
      [drug('モルヒネ'), '活性代謝物 M6G が貯留。\n遅れて呼吸抑制', '避ける。フェンタニル主体に\n（→ 3-4）'],
      [drug('ロクロニウム'), 'クリアランス低下で\n遷延しうる', 'モニターで追加を控えめに\n（→ 3-5）'],
      [drug('スガマデクス'), 'CLcr 30 未満で排泄遅延。\n7 日後も複合体を検出', '国内は注意、米国非推奨。\n高流量透析で除去可'],
      [drug('エノキサパリン'), '蓄積。プロタミンで\n完全には中和できない', { text: 'CLcr 30 未満は禁忌（国内）', color: C.vermilion, bold: true }],
      [drug('ガバペンチノイド、\nトラマドール'), '腎排泄で蓄積\n（鎮静、ふらつき）', '減量か避ける'],
      [drug('ヨード造影剤'), '静注は eGFR 30 未満、\n動注は 60 未満でリスク', '説明と生食の補液。\n臨時透析は不要'],
    ], { x: X0, y: TOP_T, colW: [3.25, 4.33, 4.62], lineSpacing: 24, margin: [3.1, 6, 0, 6], round: 0.001 });
    if (t.y + t.h > ZONE_BOT + 1e-6) throw new Error(`S4 表の下端 ${r3(t.y + t.h)} > ${ZONE_BOT}`);
    console.log(`  p${s.page} 表 ${t.y}〜${r3(t.y + t.h)}in（行の高さ ${t.rowH.join(' / ')}in）`);
  }

  // ===== S5 Child-Pugh と MELD =====
  {
    const s = deck.slide({ title: 'Child-Pugh と MELD', chapter: 2, notes: notesOf('S5'), sources: sourcesOf('S5') });
    const ctr = text => ({ text, align: 'center' });
    const t = tightTable(s, [
      ['項目', ctr('1 点'), ctr('2 点'), ctr('3 点')],
      ['肝性脳症', ctr('なし'), ctr('軽度（I〜II 度）'), ctr('III 度以上')],
      ['腹水', ctr('なし'), ctr('軽度'), ctr('中等度以上')],
      ['総ビリルビン（mg/dL）', ctr('2.0 未満'), ctr('2.0〜3.0'), ctr('3.0 超')],
      ['アルブミン（g/dL）', ctr('3.5 超'), ctr('2.8〜3.5'), ctr('2.8 未満')],
      ['PT 活性（INR）', ctr('70% 超\n（1.7 未満）'), ctr('40〜70%\n（1.7〜2.3）'), ctr('40% 未満\n（2.3 超）')],
    ], { x: X0, y: TOP, colW: [3.87, 2.55, 2.95, 2.83] });
    const y = r3(t.y + t.h + 0.34);
    s.text([
      [em('A', C.blue), '：5〜6 点、', em('B', C.blue), '：7〜9 点、', em('C', C.blue), '：10〜15 点。'],
      [em('MELD', C.blue), ' はビリルビン・クレアチニン・INR から計算'],
    ], { x: X0, y, w: BW, size: 24, gap: 6 });
  }

  // ===== S6 肝疾患：凝固と薬 =====
  {
    const s = deck.slide({ title: '肝疾患：凝固と薬', chapter: 2, notes: notesOf('S6'), sources: sourcesOf('S6') });
    const cw = r3((BW - GAP) / 2);
    // 左（凝固）は 10 行。項目の間隔を 0.1in 以上にするため、カードを TOP_T〜ZONE_BOT に置く。
    // 「高リスク手技の目安：血小板 5 万超、」（5.87in）も「フィブリノゲン 120 超、Hct 25% 以上」（5.90in）も
    // 左カードの内側 5.52in に入らず、右カードは「動態はほぼ同じ。覚醒はやや遅れる」で幅を詰められないため、4 行のまま。
    // 右（薬）は 9 行・5 項目なので、間隔を広げて下端をそろえる（fill）
    const { igaps } = cardRow(s, [
      {
        title: '凝固', titleColor: C.blue, items: [
          ['PT-INR の延長は', '出血リスクを反映しない（再均衡）'],
          ['INR を下げるための', 'FFP 大量投与はしない'],
          ['高リスク手技の目安：', '血小板 5 万超、', 'フィブリノゲン 120 超、', 'Hct 25% 以上'],
          ['低リスク手技の前に', 'ルーチンで補正しない'],
        ],
      },
      {
        title: '薬', titleColor: C.vermilion, items: [
          ['プロポフォールは単回も持続も', '動態はほぼ同じ。覚醒はやや遅れる'],
          ['レミマゾラム・ミダゾラム：', 'Child C で遷延、脳症'],
          ['ロクロニウム：', '効き始めが遅く、回復が遷延'],
          ['レミフェンタニルは動態不変', '（換気抑制には敏感）'],
          ['オピオイド・ベンゾは脳症の誘因'],
        ],
      },
    ], { y: TOP_T, avail: ZONE_BOT - 0.001 - TOP_T, widths: [cw, cw], gap: GAP, fill: true });   // 0.001: 浮動小数で警告の閾値を越えないように
    if (Math.min(...igaps) < 0.1 - 1e-6) throw new Error(`S6 カードの項目間隔 ${igaps.join(' / ')}in < 0.1in`);
    console.log(`  p${s.page} 項目間隔 左 ${igaps[0]} / 右 ${igaps[1]}in`);
  }

  // ===== S7 延期と相談の目安 =====
  {
    const s = deck.slide({ title: '延期と相談の目安', chapter: 2, notes: notesOf('S7'), sources: sourcesOf('S7') });
    const bh = calloutH(1, 26);
    const by = r3(BOT - bh);
    // 「腹水・脳症・Na 130 未満は整えてから」を 1 行に入れるため、右のカードを広くする
    const w2 = 6.43, w1 = r3(BW - GAP - w2);
    const { h } = cardRow(s, [
      {
        title: '腎・透析', titleColor: C.blue, items: [
          [[em('K 5.5 超'), 'や心電図変化：'], '透析か補正してから'],
          ['溢水：透析してから。', '前日透析を手配'],
          ['透析の抗凝固と', '当日の時刻を決める'],
          ['相談：腎臓内科・透析室'],
        ],
      },
      {
        title: '肝', titleColor: C.vermilion, items: [
          [[em('Child C、MELD 20 超'), '：'], '緊急性の高い手術だけ'],
          [[em('MELD 15 以上'), '：待機手術の前に'], '肝臓専門医・移植施設へ'],
          ['腹水・脳症・Na 130 未満は整えてから'],
          ['緊急手術は死亡が待機の約 3〜5 倍。', '搬送は早めに'],
        ],
      },
    ], { y: TOP, avail: by - GAP - TOP, widths: [w1, w2], gap: GAP });
    if (TOP + h > by - 0.24 + 1e-6) throw new Error('S7 カードと強調ボックスの間隔が足りません');
    callout(s, [[em('一人で抱えない。'), '本土の専門医に電話し、搬送の判断を早く']], { y: by, h: bh, size: 26 });
  }

  // ===== S8 まとめ =====
  {
    const s = deck.slide({ title: 'まとめ', chapter: 3, notes: notesOf('S8'), sources: sourcesOf('S8') });
    // 箇条書きの記号は各行の先頭の文字の色になる（2・3・5 行目が橙になる）ため、記号を別の文字枠にして
    // 本文色にそろえる。字下げは s.bullets と同じ 28pt、行送りは 0.86in（PowerPoint での s.bullets の
    // 1 行 36pt＋段落後 26pt と同じ）、本文ゾーン（TOP〜BOT）の縦中央。
    const items = [
      ['透析は', em('前日'), 'に。ドライウェイト、K、抗凝固、シャントを確かめる'],
      [em('K 5.5 超'), 'は延期か補正。シャント肢には何もしない'],
      [em('モルヒネを避け'), '、ロクロニウムはモニター、スガマデクスは慎重に'],
      ['造影剤は ', em('eGFR 30（静注）・60（動注）'), '未満で説明と生食補液'],
      [em('Child C・MELD 20 超'), 'は緊急以外しない。INR は出血を反映しない'],
    ];
    const size = 24, ind = r3(Math.round(size * 1.15) / 72), pitch = 0.86, lh = 0.5;
    const y0 = r3(TOP + (BOT - TOP - (pitch * (items.length - 1) + lh)) / 2);
    items.forEach((it, i) => {
      const y = r3(y0 + i * pitch);
      s.text(['•'], { x: X0, y, w: ind, h: lh, size, color: C.dark });
      s.text([it], { x: r3(X0 + ind), y, w: r3(12.2 - ind), h: lh, size });
    });
  }

  if (DEVIATIONS.length) console.log('[deviations]\n  ' + DEVIATIONS.join('\n  '));
  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
