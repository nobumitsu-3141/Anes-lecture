'use strict';
/**
 * 2-3 JSA「安全な麻酔のためのモニター指針」と標準モニター — 組版スクリプト
 *
 *   cd slides/2-3_monitor_guideline
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）は同じディレクトリの spec.md が正本。
 *   - スライド面の文言は spec の on-slide をそのまま書き写している（変えているのは改行位置だけ）。
 *   - ノートは spec.md から実行時に読み込む（原文のまま全文）。
 *   - 出典は spec の sources 行を項目に分けて渡す（2行への折り返しはライブラリ）。
 *     例外: S11 は2行に収めるため URL の "https://" を省いた。
 * PDF / slide-*.jpg はリポジトリの外（RENDER_DIR または DECK_RENDER_DIR。既定は OS の一時ディレクトリ）に出す。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, heightFor } = require('../_lib/decklib');

const NAME = '2-3_monitor_guideline';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.RENDER_DIR || process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);
const RENDER = !process.argv.includes('--no-render');

// ---------------------------------------------------------------------------
// spec.md からノートと出典行を読む
// ---------------------------------------------------------------------------
function parseSpec(file) {
  const out = {};
  let cur = null;
  let inNotes = false;
  for (const ln of fs.readFileSync(file, 'utf8').split('\n')) {
    const h = ln.match(/^## (S\d+)\s/);
    if (h) { cur = out[h[1]] = { notes: [], sources: null }; inNotes = false; continue; }
    if (!cur) continue;
    if (/^notes:\s*\|\s*$/.test(ln)) { inNotes = true; continue; }
    if (inNotes) {
      if (ln.startsWith('  ')) { cur.notes.push(ln.slice(2)); continue; }
      if (ln.trim() === '') { cur.notes.push(''); continue; }
      inNotes = false;
    }
    const s = ln.match(/^sources:\s*(.*)$/);
    if (s) cur.sources = s[1].trim();
  }
  for (const k of Object.keys(out)) {
    const n = out[k].notes;
    while (n.length && n[n.length - 1] === '') n.pop();
    if (!n.length) throw new Error(`spec.md: ${k} の notes が見つかりません`);
    out[k].notes = n.join('\n');
  }
  return out;
}
const SPEC = parseSpec(path.join(__dirname, 'spec.md'));
const notes = id => {
  if (!SPEC[id]) throw new Error(`spec.md に ${id} がありません`);
  return SPEC[id].notes;
};

// 出典（spec の sources 行。1項目が16ptで1行に入らないものは語の切れ目で2項目に分けた）
const JSA = '日本麻酔科学会「安全な麻酔のためのモニター指針」（2025年8月改訂）';
const JSA_URL = 'https://anesth.or.jp/users/person/guide_line';
const SUZUKI = 'Suzuki T, et al. J Anesth 2026;40:173-181';
const SRC = {
  S2: [JSA, JSA_URL],
  S3: [JSA, SUZUKI],
  S4: [JSA],
  S5: [JSA],
  S6: [JSA, SUZUKI],
  S7: [JSA],
  S8: ['Wax DB, et al. Invasive and concomitant noninvasive intraoperative',
    'blood pressure monitoring. Anesthesiology 2011;115:973-978'],
  S9: [JSA],
  S10: [],
  // 3行になるため URL の "https://" を省いた（deviation に記載）
  S11: [JSA, 'anesth.or.jp/users/person/guide_line', SUZUKI],
};
// spec の sources 行と突き合わせ（区切りと空白を無視して比較。違えば警告）
function checkSources() {
  const norm = s => String(s).replace(/ \/ /g, '').replace(/\s+/g, '');
  for (const [id, items] of Object.entries(SRC)) {
    const spec = SPEC[id] && SPEC[id].sources;
    const want = (!spec || spec === '—') ? '' : norm(spec);
    const got = norm(items.join(' / '));
    if (want !== got && !(id === 'S11' && want === got.replace('anesth.or.jp', 'https://anesth.or.jp'))) {
      console.warn(`[build] ${id}: 出典が spec と違います\n  spec: ${spec}\n  got : ${items.join(' / ')}`);
    }
  }
}
checkSources();

// ---------------------------------------------------------------------------
// 見出し＋項目のカード。1項目を2行に割るときに、項目内は行間を詰め、項目間だけを空ける。
// （deck の card() は全段落に同じ段落後スペースを付けるため、割った行が別項目に見える）
function itemCard(s, o) {
  const pad = o.pad != null ? o.pad : 0.12;             // 左右の内側余白
  const padY = o.padY != null ? o.padY : pad;            // 上下の内側余白
  const ts = o.titleSize || 28, bs = o.bodySize || 22, ig = o.itemGap != null ? o.itemGap : 0.3;
  const tg = o.titleGap != null ? o.titleGap : 0.15;
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: o.fill || C.bg, radius: 0.08, line: o.line });
  const iw = o.w - 2 * pad;
  let y = o.y + padY;
  const th = heightFor(1, ts, 0);
  s.text(o.title, { x: o.x + pad, y, w: iw, h: th, size: ts, bold: true, color: o.titleColor || C.blue });
  y += th + tg;
  if (o.bodyY != null) {                                 // 項目の開始位置を外から指定（列どうしで1項目目の高さをそろえる）
    if (o.bodyY < y - 1e-6) throw new Error(`itemCard「${o.title}」: bodyY ${o.bodyY} が見出しに掛かります`);
    y = o.bodyY;
  }
  o.items.forEach((lines, k) => {
    const h = heightFor(lines.length, bs, 0);
    s.text(lines, { x: o.x + pad, y, w: iw, h, size: bs, color: C.dark });
    y += h + (k < o.items.length - 1 ? ig : 0);
  });
  if (y > o.y + o.h - padY + 1e-6) throw new Error(`itemCard「${o.title}」: 高さ不足（必要 ${(y + padY - o.y).toFixed(2)}in）`);
}

// itemCard の項目部分の高さ
const itemsH = (items, bs, ig) => items.reduce((a, l) => a + heightFor(l.length, bs, 0), 0) + ig * (items.length - 1);
// 見出しの下の残りの高さに、列のうち最も高い項目群を置いたときの開始位置（余白の 45% を上に）
function bodyStart(cards, o) {
  const top = o.y + o.padY + heightFor(1, o.titleSize, 0) + o.titleGap;
  const avail = o.y + o.h - o.padY - top;
  const need = Math.max(...cards.map(c => itemsH(c.items, o.bodySize, o.itemGap)));
  if (need > avail + 1e-6) throw new Error(`bodyStart: 高さ不足（必要 ${need.toFixed(2)} > ${avail.toFixed(2)}in）`);
  return top + (avail - need) * 0.45;
}

// ---------------------------------------------------------------------------
async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['指針', '7項目', '心電図', '血圧', '現場'],
    author: '初期研修医向け麻酔講義',
    deckTitle: '2-3 モニター指針と標準モニター',
  });

  // ===== S1 表紙 =====
  deck.cover({
    title: 'モニター指針と標準モニター',
    subtitle: '2-3 モニタリング',
    meta: '初期研修医向け麻酔講義｜日本麻酔科学会\n「安全な麻酔のためのモニター指針」（2025年8月改訂）に沿って',
    notes: notes('S1'),
  });

  // ===== S2 モニターは誰が見るか（引用ボックス＋対象の麻酔3枚＋一文） =====
  {
    const s = deck.slide({ title: 'モニターは誰が見るか', chapter: 0, notes: notes('S2'), sources: SRC.S2 });
    const top = 1.85, bottom = 5.3;
    const cy0 = top + 0.5 + 0.24, gap = 0.24;
    // 左: 指針の第1項目（原文どおり2行）。28pt が 7.45in 幅で1行に入る上限なので、
    // 文字を大きくする代わりに箱を右の3枚のカードの高さ（cy0〜bottom）にそろえて詰める
    s.box({
      x: 0.55, y: cy0, w: 7.45, h: bottom - cy0, fill: C.bg, radius: 0.12, margin: 0.2,
      text: ['現場に麻酔を担当する医師が居て、', '絶え間なく看視すること'],
      size: 28, bold: true, color: C.dark, align: 'center',
    });
    // 右: 対象となる麻酔
    const rx = 8.25, rw = 12.78 - 8.25;
    s.text('この指針が対象とする麻酔', { x: rx, y: top, w: rw, size: 24, bold: true, color: C.blue });
    const ch = (bottom - cy0 - 2 * gap) / 3;
    ['全身麻酔', '硬膜外麻酔', '脊髄くも膜下麻酔'].forEach((t, i) => {
      s.box({
        x: rx, y: cy0 + i * (ch + gap), w: rw, h: ch, fill: C.bg, radius: 0.08,
        line: { color: C.blue, width: 1.5 }, text: t, size: 26, color: C.dark,
      });
    });
    // 下: 一文
    s.text('機械は補助、判断するのは人', {
      x: 0.55, y: 5.6, w: 12.23, size: 32, bold: true, color: C.teal, align: 'center',
    });
  }

  // ===== S3 7項目の全体像（上段4枚・下段3枚のカード＋その下に全幅の補足ボックス） =====
  {
    const s = deck.slide({ title: '7項目の全体像', chapter: 1, notes: notes('S3'), sources: SRC.S3 });
    const gx = 0.25, ch = 1.34, gy = 0.24, y0 = 1.85;
    // 上段は4列、下段は3枚を全幅に並べる（下段4列目に空き枠を残さない）
    const rows = [
      { n: 4, items: [['① 担当医', '現場で看視'], ['② 酸素化', 'SpO2'], ['③ 換気', 'EtCO2'], ['④ 循環', 'ECG・血圧']] },
      { n: 3, items: [['⑤ 体温', '体温測定'], ['⑥ 筋弛緩', '定量（TOF比）'], ['⑦ 脳波', '必要に応じて']] },
    ];
    const colX = n => Array.from({ length: n }, (_, i) => 0.55 + i * ((12.23 - (n - 1) * gx) / n + gx));
    rows.forEach((r, ri) => {
      const cw = (12.23 - (r.n - 1) * gx) / r.n, xs = colX(r.n);
      r.items.forEach(([t, b], i) => {
        s.card({ x: xs[i], y: y0 + ri * (ch + gy), w: cw, h: ch, title: t, lines: [b], titleColor: C.blue });
      });
    });
    // 補足ボックス（3項目）: 22pt で1項目目が 5.85in・2項目目が 3.76in あり、カード1枠（4列の内側 2.57in）には入らない。
    // 3行に積むと 1.43in 要り、出典行（y6.5）との間を 0.24in 以上あけられないため、
    // 2段のカードの下に全幅の帯を置き、3項目を左揃えで横に並べる（1項目目だけ語の切れ目で2行）。
    // 各項目の左端は下段カード（⑤⑥⑦）の文字の左端にそろえる。
    const by = y0 + 2 * (ch + gy), bpy = 0.12;
    const supp = [
      ['②〜⑤は毎回', '（EtCO2 は全身麻酔で）'],
      ['⑥は筋弛緩薬を使うとき'],
      ['⑦は必要に応じて'],
    ];
    const th = heightFor(2, 22, 0), bh = th + 2 * bpy;
    s.box({ x: 0.55, y: by, w: 12.23, h: bh, fill: C.white, radius: 0.08, line: { color: C.teal, width: 2 } });
    const sx = colX(3).map(x => x + 0.15), tws = [3.85, 3.8, 12.78 - 0.15 - sx[2]];
    supp.forEach((t, i) => s.text(t, { x: sx[i], y: by + bpy, w: tws[i], h: th, size: 22, color: C.dark }));
  }

  // ===== S4 酸素化と換気（左右2枚の大きなカード＋案内1行） =====
  {
    const s = deck.slide({ title: '酸素化と換気', chapter: 1, notes: notes('S4'), sources: SRC.S4 });
    const cw = (12.23 - 0.3) / 2, cy = 1.95, ch = 3.15;
    const common = { y: cy, w: cw, h: ch, titleSize: 28, bodySize: 24, gap: 20 };
    s.card({ ...common, x: 0.55, title: '② 酸素化', titleColor: C.blue,
      lines: ['皮膚・粘膜・血液の色を看視', 'パルスオキシメータを装着'] });
    s.card({ ...common, x: 0.55 + cw + 0.3, title: '③ 換気', titleColor: C.blue,
      lines: ['胸郭・呼吸バッグの動き、呼吸音', '全身麻酔ではカプノメータ', '換気量モニターは適宜'] });
    s.text('原理と読み方 → 2-4 パルスオキシメータ／2-5 カプノグラフィ', {
      x: 0.55, y: cy + ch + 0.45, w: 12.23, size: 24, color: C.dark, align: 'center',
    });
  }

  // ===== S5 循環（3列カード） =====
  {
    const s = deck.slide({ title: '循環', chapter: 1, notes: notes('S5'), sources: SRC.S5 });
    // 3列とも同じ幅。26pt では「心電図モニターを用いる」(4.45in) が1行に入らないため語の切れ目で2行にし、
    // 割った項目は行間を詰めて項目どうしの間だけ空ける（itemCard）
    const gap = 0.24, cw = (12.23 - 2 * gap) / 3, cy = 1.95, ch = 4.2;
    const cards = [
      { title: '拍動を確かめる', items: [['心音・動脈の触診'], ['動脈波形・脈波'], ['いずれか一つを監視']] },
      { title: '心電図', items: [['心電図モニターを', '用いる'], ['リズムと心拍数']] },
      { title: '血圧', items: [['原則5分間隔'], ['必要なら頻回に'], ['観血式は必要時']] },
    ];
    const geo = { y: cy, h: ch, padY: 0.18, titleSize: 32, bodySize: 26, titleGap: 0.3, itemGap: 0.38 };
    const bodyY = bodyStart(cards, geo);
    cards.forEach((c, i) => itemCard(s, { ...geo, x: 0.55 + i * (cw + gap), w: cw, bodyY, ...c }));
  }

  // ===== S6 体温・筋弛緩・脳波（3列カード＋変遷の3段ボックス） =====
  {
    const s = deck.slide({ title: '体温・筋弛緩・脳波', chapter: 1, notes: notes('S6'), sources: SRC.S6 });
    // カードの列と変遷ボックスの列をそろえる（同じ幅・同じ間隔 az）。中央列は「抜管前に TOF比 0.9 以上」(22pt 3.85in) が入る幅
    const az = 0.6, w2 = 4.25, w1 = (12.23 - 2 * az - w2) / 2;
    const ws = [w1, w2, w1], cy = 1.85, ch = 2.43;
    const xs = [0.55, 0.55 + ws[0] + az, 0.55 + ws[0] + ws[1] + 2 * az];
    s.card({ x: xs[0], y: cy, w: ws[0], h: ch, title: '⑤ 体温', titleColor: C.blue, lines: ['体温測定を行う'] });
    s.card({
      x: xs[1], y: cy, w: ws[1], h: ch, title: '⑥ 筋弛緩', titleColor: C.teal, fill: C.white,
      line: { color: C.teal, width: 3 },
      lines: ['筋弛緩薬を使うとき', '深さを定量的に評価', ['抜管前に ', { text: 'TOF比 0.9 以上', bold: true }]],
    });
    s.card({ x: xs[2], y: cy, w: ws[2], h: ch, title: '⑦ 脳波', titleColor: C.blue, lines: ['必要に応じて装着'] });

    // 変遷: 2014 → 2019 → 2025（年を1行目、内容を2行目以降に）
    const fy = cy + ch + 0.32, fh = 1.5;
    const steps = [
      { text: [{ text: '2014', bold: true, color: C.gray }, '必要に応じて'], fill: C.bg },
      { text: [{ text: '2019', bold: true, color: C.blue }, '使う際はモニタリング'], fill: C.bg },
      { text: [{ text: '2025', bold: true, color: C.teal }, '定量評価・抜管前', 'TOF比 0.9 以上'], fill: C.white, line: { color: C.teal, width: 3 } },
    ];
    steps.forEach((st, i) => {
      const x = xs[i], fw = ws[i];
      s.box({ x, y: fy, w: fw, h: fh, radius: 0.08, size: 22, color: C.dark, ...st });
      if (i < 2) s.arrow({ x1: x + fw + 0.1, y1: fy + fh / 2, x2: x + fw + az - 0.1, y2: fy + fh / 2, color: C.chipText, width: 3 });
    });
  }

  // ===== S7 心電図：何を見るか（電極配置図＋注記＋縦3枚のカード） =====
  {
    const s = deck.slide({ title: '心電図：何を見るか', chapter: 2, notes: notes('S7'), sources: SRC.S7 });
    // --- 図: 胴体の輪郭（角丸四角）と電極（円）。ラベルは図の右に並べる
    const tx = 0.55, ty = 1.95, tw = 2.1, th = 3.55;
    s.box({ x: tx, y: ty, w: tw, h: th, fill: C.bg, radius: 0.4, line: { color: C.gray, width: 1.5 } });
    const EC = {
      R: { fill: C.red, ink: C.white },
      L: { fill: 'F0E442', line: { color: C.chipText, width: 1 }, ink: C.dark },
      N: { fill: C.dark, ink: C.white },
      F: { fill: '009E73', ink: C.white },
      C: { fill: C.white, line: { color: C.chipText, width: 1.5 }, ink: C.dark },
    };
    const ed = 0.46;
    const pos = {                       // 患者の右 = 図の左
      R: [tx + 0.5, ty + 0.55],         // 右鎖骨下
      L: [tx + tw - 0.5, ty + 0.55],    // 左鎖骨下
      N: [tx + 0.5, ty + th - 0.55],    // 右下腹部
      F: [tx + tw - 0.5, ty + th - 0.55], // 左下腹部
      C: [tx + tw - 0.36, ty + 1.6],    // 左前腋窩線・V4 と同じ高さ（第5肋間付近）= V5
    };
    // 電極は色だけでなく記号でも凡例と対応させる（円の中に R/L/N/F/C。文字は円そのものの文字＝1図形）
    for (const k of Object.keys(pos)) {
      const { ink, ...st } = EC[k];
      s.circle({ cx: pos[k][0], cy: pos[k][1], d: ed, ...st, text: k, size: 22, bold: true, color: ink, margin: 0.02 });
    }
    // ラベル（図の外）: 色の見本の小円＋文字
    const labels = [['R', '赤 R（右上）'], ['L', '黄 L（左上）'], ['N', '黒 N（右下）'], ['F', '緑 F（左下）'], ['C', '白 C（V5）']];
    const lx = tx + tw + 0.25, lh = 0.46, lstep = (th - lh) / (labels.length - 1);
    labels.forEach(([k, t], i) => {
      const y = ty + i * lstep;
      const { ink, ...st } = EC[k];
      s.circle({ cx: lx + 0.14, cy: y + lh / 2, d: 0.28, ...st });
      s.text(t, { x: lx + 0.42, y, w: 2.08, h: lh, size: 22, valign: 'middle' });
    });
    s.text('色は IEC 方式。記号で確認', { x: 0.55, y: ty + th + 0.3, w: 4.9, size: 22 });

    // --- 右: 縦3枚のカード（見出しは左、本文は右）
    const kx = 5.7, kw = 12.78 - kx, hx = kx + 0.15, hw = 1.52, bx = hx + hw + 0.18, bw = 12.78 - 0.15 - bx;
    const cards = [
      { title: '誘導', lines: ['II：リズム・P波', 'V5：虚血（ST）'] },
      { title: '見るもの', lines: ['心拍数', 'リズム', 'ST の変化'] },
      { title: '落とし穴', lines: ['電気メスのノイズ', '波形があっても脈がない（PEA）'] },
    ];
    let y = 1.85;
    cards.forEach(c => {
      const bh = heightFor(c.lines.length, 22, 0);
      const h = bh + 0.28;
      s.box({ x: kx, y, w: kw, h, fill: C.bg, radius: 0.08 });
      s.text(c.title, { x: hx, y: y + 0.12, w: hw, h: h - 0.24, size: 24, bold: true, color: C.blue, valign: 'middle' });
      // 本文も見出しと同じ枠・上下中央にそろえる（見出しの中心と本文の中心を合わせる）
      s.text(c.lines, { x: bx, y: y + 0.12, w: bw, h: h - 0.24, size: 22, color: C.dark, valign: 'middle' });
      y += h + 0.24;
    });
  }

  // ===== S8 非観血的血圧の原理（オシロメトリック法の図＋箇条書き） =====
  {
    const s = deck.slide({ title: '非観血的血圧の原理', chapter: 3, notes: notes('S8'), sources: SRC.S8 });
    // 軸
    const ox = 0.85, oy = 4.4, xEnd = 5.6, yTop = 2.35;
    s.text('カフ圧', { x: 0.55, y: 1.85, w: 1.3, size: 22, color: C.chipText });
    s.arrow({ x1: ox, y1: oy, x2: ox, y2: yTop, color: C.chipText, width: 2 });
    s.arrow({ x1: ox, y1: oy, x2: xEnd, y2: oy, color: C.chipText, width: 2 });
    s.text('時間', { x: xEnd - 0.75, y: oy + 0.1, w: 0.75, size: 22, color: C.chipText, align: 'right' });
    // カフ圧（右下がりの直線）と振動（振幅が増えてから減る）
    const x0 = 1.05, x1 = 5.35, p0 = 2.6, p1 = 4.05, xm = 3.05;
    const cuff = x => p0 + (p1 - p0) * (x - x0) / (x1 - x0);
    const amp = x => 0.28 * Math.exp(-(((x - xm) / 0.85) ** 2));
    const wave = [], envU = [], envL = [];
    const period = 0.24, n = Math.round((x1 - x0) / period * 16);
    for (let i = 0; i <= n; i++) {
      const x = x0 + (x1 - x0) * i / n;
      wave.push([x, cuff(x) - amp(x) * Math.sin(2 * Math.PI * (x - x0) / period)]);
    }
    for (let i = 0; i <= 60; i++) {
      const x = x0 + (x1 - x0) * i / 60;
      envU.push([x, cuff(x) - amp(x)]);
      envL.push([x, cuff(x) + amp(x)]);
    }
    s.polyline(envU, { color: C.gray, width: 1.25, dash: 'dash' });
    s.polyline(envL, { color: C.gray, width: 1.25, dash: 'dash' });
    s.polyline(wave, { color: C.vermilion, width: 2 });
    s.polyline([[x0, p0], [x1, p1]], { color: C.blue, width: 3.5 });
    // 振幅が最大の位置に縦の点線 → 下のラベルへ
    s.vline({ x: xm, y: yTop, h: oy + 0.62 - yTop, color: C.teal, width: 2.5, dash: 'sysDot' });
    s.text('振幅が最大 ＝ 平均血圧', { x: xm - 1.95, y: oy + 0.72, w: 3.9, size: 22, bold: true, color: C.teal, align: 'center' });
    s.text('収縮期・拡張期は計算値', { x: xm - 1.95, y: oy + 1.32, w: 3.9, size: 22, color: C.dark, align: 'center' });

    // 右: 箇条書き
    s.bullets([
      'カフを膨らませて徐々に下げる',
      '振動が最大 → 平均血圧',
      '収縮期・拡張期は装置の計算値',
      '不整脈・体動・シバリングで誤差',
    ], { x: 6.7, y: 2.1, w: 12.78 - 6.7, gap: 28 });
  }

  // ===== S9 カフの選び方と注意点（3列カード） =====
  {
    const s = deck.slide({ title: 'カフの選び方と注意点', chapter: 3, notes: notes('S9'), sources: SRC.S9 });
    // 3列とも同じ幅。24pt で1行に入らない項目は「・」の後（語の切れ目）で2行に割る。
    // 2行に割った項目は行間を詰め、項目どうしの間だけ広げる（itemCard）
    const gap = 0.24, cw = (12.23 - 2 * gap) / 3, cy = 1.95, ch = 4.2;
    const cards = [
      { title: 'カフの幅', items: [['上腕周の約 40%'], ['小さすぎると高く出る']] },
      { title: '装着', items: [['心臓の高さに'], ['点滴側・', 'シャント側は避ける']] },
      { title: '間隔', items: [['原則 5 分'], ['頻回なら皮下出血・', '神経障害に注意']] },
    ];
    // 左右の余白 0.09in: 24pt「小さすぎると高く出る」(3.733in) が内側幅に入る最小限
    const geo = { y: cy, h: ch, pad: 0.09, padY: 0.2, titleSize: 32, bodySize: 24, titleGap: 0.3, itemGap: 0.8 };
    const bodyY = bodyStart(cards, geo);
    cards.forEach((c, i) => itemCard(s, { ...geo, x: 0.55 + i * (cw + gap), w: cw, bodyY, ...c }));
  }

  // ===== S10 モニターの見方の型（5つのチップのフロー＋箇条書き3行） =====
  {
    const s = deck.slide({ title: 'モニターの見方の型', chapter: 4, notes: notes('S10') });
    const steps = ['患者を見る', 'SpO2', 'EtCO2', '血圧・心拍', '体温・TOF'];
    // チップを詰めて間隔 0.62in・矢印 0.42in を確保（チップ文字は 22pt。「患者を見る」1.71in が内側 1.79in に入る）
    const cw = 1.95, az = (12.23 - steps.length * cw) / (steps.length - 1), cy = 2.1, ch = 1.1;
    steps.forEach((t, i) => {
      const x = 0.55 + i * (cw + az);
      s.box({ x, y: cy, w: cw, h: ch, fill: C.blue, radius: 0.1, text: t, size: 22, bold: true, color: C.white, margin: 0.08 });
      if (i < steps.length - 1) s.arrow({ x1: x + cw + 0.1, y1: cy + ch / 2, x2: x + cw + az - 0.1, y2: cy + ch / 2, color: C.chipText, width: 3 });
    });
    s.bullets([
      'アラームの上下限を患者ごとに設定する',
      '数値より、波形と傾向を見る',
      '異常に気づいたら声に出して共有する',
    ], { y: 3.95, size: 28, gap: 30 });
  }

  // ===== S11 まとめ（箇条書き5行） =====
  {
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: notes('S11'), sources: SRC.S11 });
    s.bullets([
      '担当医が現場にいて、絶え間なく看視する',
      'SpO2・心電図・血圧・体温は毎回、全身麻酔では EtCO2 も',
      ['筋弛緩薬を使うなら定量モニター、抜管前に ', { text: 'TOF比 0.9 以上', bold: true, color: C.teal }],
      ['非観血的血圧は', { text: '低血圧で高めに出る', bold: true, color: C.red }, '。信頼できるのは平均血圧'],
      '決まった順番で見て、異常は声に出す',
    ], { y: 1.95, size: 26, gap: 24 });
  }

  const res = await deck.save({ render: RENDER, renderDir: RENDER_DIR });
  if (res.render) console.log('[build] render: ' + res.render.images.length + ' images in ' + RENDER_DIR);
  if (res.violations !== 0) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
