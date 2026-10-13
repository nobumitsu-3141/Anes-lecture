'use strict';
/**
 * 2-5b カプノグラフィ②：異常波形の読み方
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言はここに転記（改行位置のみ調整）、ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 *
 * 波形は s.polyline（ネイティブ図形）。正常部分はブルー、異常部分はバーミリオン。ラベルは図の外。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, fits, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '2-5b_capnography_abnormal';
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
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  return s ? s.split(' / ') : undefined;
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const em = (text, color, bold = true) => ({ text, color, bold });
const WAVE_W = 3.5;              // 波形の線幅(pt)
// 強調文字のティール。C.teal(00A8AA) は白地で 2.9:1・F2F2F2 地で 2.6:1 と薄いので、文字には濃いトーンを使う
// （白地 5.0:1 / F2F2F2 地 4.5:1）。枠線はチップと同じ C.teal のまま。
const TEAL_TXT = '007C80';

/**
 * 右列の箇条書き（手組み）。items: 各項目 = 行の配列（行は文字列 or run 配列）。
 * 1項目が長いときだけ語句の切れ目で2行にする（文言は変えない）。
 * y を渡すと上端、yc を渡すと縦中央に合わせる。戻り値 {top, bottom}
 */
function bulletCol(s, items, o) {
  const size = o.size || 24, ind = o.indent || 0.42, gap = o.gap != null ? o.gap : 0.3;
  const hs = items.map(it => heightFor(it.length, size, 0));
  const total = hs.reduce((a, b) => a + b, 0) + gap * (items.length - 1);
  const top = o.yc != null ? o.yc - total / 2 : o.y;
  let y = top;
  items.forEach((lines, i) => {
    s.text('•', { x: o.x, y, w: ind, size, color: C.dark });
    s.text(lines, { x: o.x + ind, y, w: o.w - ind, size, color: C.dark });
    y += hs[i] + gap;
  });
  return { top, bottom: top + total };
}

/** 図の軸（灰色の矢印）。zeroDash: 横軸の代わりに 0 の位置を点線で示す */
function axes(s, f, o = {}) {
  s.arrow({ x1: f.x, y1: f.yb, x2: f.x, y2: f.yt, color: C.gray, width: 2 });
  if (o.zeroDash) s.hline({ x: f.x, y: f.yb, w: f.xr - f.x, color: C.gray, width: 2, dash: 'dash' });
  else s.arrow({ x1: f.x, y1: f.yb, x2: f.xr, y2: f.yb, color: C.gray, width: 2 });
}

/**
 * 時間カプノグラムの 1 呼吸ぶんの点列（x から x+T まで）。
 *   yb: CO2 = 0 の y、H: 0 からプラトー終点（EtCO2）までの高さ(in)
 *   o.b0 / o.b1: 呼気開始時 / 呼気終了後の基線の高さ(in)（再呼吸で基線が浮く）
 *   o.kind: 'normal' | 'shark'（閉塞: 立ち上がりが鈍く斜め）| 'cleft'（くぼみ）| 'osc'（心拍の揺れ）
 */
function breath(x, T, yb, H, o = {}) {
  const b0 = o.b0 || 0, b1 = o.b1 != null ? o.b1 : b0;
  const y0 = yb - b0, y1 = yb - b1;
  const kind = o.kind || 'normal';
  const top1 = yb - H;                        // プラトー終点
  const top0 = yb - (b0 + (H - b0) * 0.93);   // プラトー始点（わずかに右上がり）
  const A = x + 0.12 * T;                     // 相 I の終わり
  const P1 = x + 0.70 * T;                    // 呼気の終わり
  const pts = [[x, y0], [A, y0]];
  if (kind === 'shark') {
    const n = 18;
    for (let k = 1; k <= n; k++) {
      const t = k / n, f = (1 - Math.exp(-2.4 * t)) / (1 - Math.exp(-2.4));
      pts.push([A + (P1 - A) * t, y0 + (top1 - y0) * f]);
    }
  } else {
    const rise = y0 - top0;
    const P0 = A + 0.075 * T;
    pts.push([A + 0.015 * T, y0 - 0.12 * rise], [A + 0.035 * T, y0 - 0.62 * rise],
      [A + 0.055 * T, y0 - 0.93 * rise], [P0, top0]);
    const yAt = xx => top0 + (top1 - top0) * (xx - P0) / (P1 - P0);
    if (kind === 'cleft') {
      const c = P0 + (P1 - P0) * 0.5, d = 0.38 * (y0 - top1);
      pts.push([c - 0.045 * T, yAt(c - 0.045 * T)], [c, yAt(c) + d], [c + 0.055 * T, yAt(c + 0.055 * T)]);
    }
    if (kind === 'osc') {
      // プラトーの後半〜下降脚の上半分に、心拍に同期した小さなさざ波。
      // 中心線（プラトー → ゆっくりした下降）に沿った道のりで位相を取り、中心線に垂直な方向へずらす
      // （急な下降脚でも波が大きくうねらず、細かいさざ波に見えるように）。
      // 折れ目（プラトー終点）と両端は振幅 0 の位置に置く。下半分は揺れなしで基線まで下がる。
      const wl = o.wl || 0.085, amp = o.amp || 0.022;
      const Pm = x + 0.45 * T;                       // さざ波の始まり（プラトー後半）
      const Dm = P1 + 0.22 * T;                      // 下降脚の上半分の終わり
      const yMid = top1 + 0.5 * (y1 - top1);
      const center = [[Pm, yAt(Pm)], [P1, top1], [Dm, yMid]];
      pts.push(center[0]);
      for (let i = 0; i < center.length - 1; i++) {
        const [ax, ay] = center[i], [bx, by] = center[i + 1];
        const L = Math.hypot(bx - ax, by - ay);
        const nHalf = Math.max(1, Math.round(L / (wl / 2)));   // 道のりを半波長の整数倍にそろえる
        const ux = (bx - ax) / L, uy = (by - ay) / L;          // 進行方向
        const nx = uy, ny = -ux;                               // 法線（上向き側）
        const m = nHalf * 8;
        for (let k = 1; k <= m; k++) {
          const t = k / m, off = amp * Math.sin(Math.PI * nHalf * t);
          pts.push([ax + (bx - ax) * t + nx * off, ay + (by - ay) * t + ny * off]);
        }
      }
      pts.push([Dm + 0.02 * T, y1 - 0.08 * (y1 - top1)], [Dm + 0.035 * T, y1], [x + T, y1]);
      return pts;
    }
    pts.push([P1, top1]);
  }
  // 相 0（吸気: 急な下降）
  const drop = y1 - top1;
  pts.push([P1 + 0.012 * T, top1 + 0.3 * drop], [P1 + 0.03 * T, y1 - 0.06 * drop], [P1 + 0.045 * T, y1], [x + T, y1]);
  return pts;
}

/** 呼吸の列を描く。list: [{T, H, ...opts}]。color ごとに連続部分を1本の折れ線にする。戻り値: 終端の x */
function waveRun(s, x, yb, list, color) {
  let pts = [];
  list.forEach(b => {
    const bp = breath(x, b.T, yb, b.H, b);
    pts = pts.length ? pts.concat(bp.slice(1)) : bp;
    x += b.T;
  });
  s.polyline(pts, { color, width: WAVE_W });
  return x;
}
const rep = (n, o) => Array.from({ length: n }, () => ({ ...o }));

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['順番', '消える・\n下がる', '上がる', '形', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '2-5b カプノグラフィ②：異常波形の読み方',
  });

  // ===== S1 表紙 =====
  deck.cover({
    title: 'カプノグラフィ② 異常波形の読み方',
    subtitle: '2-5 モニタリング',
    meta: '初期研修医向け麻酔講義',
    notes: notesOf('S1'),
  });

  // ===== S2 読む順番 =====
  {
    const s = deck.slide({ title: '読む順番', chapter: 0, notes: notesOf('S2'), sources: sourcesOf('S2') });
    // 上: 4つのチップを矢印でつないだフロー（番号と語句の間の空白で改行）
    const steps = [['①', '波形はあるか'], ['②', '高さ（値）は'], ['③', '形は'], ['④', '傾向は']];
    const cw = 2.55, ch = 1.25, cy = 1.95, g = (BW - 4 * cw) / 3;
    steps.forEach(([n, t], i) => {
      const x = X0 + i * (cw + g);
      s.box({ x, y: cy, w: cw, h: ch, fill: C.bg, line: { color: C.blue, width: 2 }, radius: 0.1,
        text: [[em(n, C.blue)].map(r => ({ ...r, size: 28 })), [em(t, C.dark)]], size: 24 });
      if (i < 3) s.arrow({ x1: x + cw + 0.12, y1: cy + ch / 2, x2: x + cw + g - 0.12, y2: cy + ch / 2, color: '404040', width: 3 });
    });
    // 中: 強調ボックス
    s.box({ x: X0, y: 3.7, w: BW, h: 0.95, fill: C.bg, line: { color: C.teal, width: 2.25 }, radius: 0.08,
      text: '異常は、まず患者（換気・循環）、次に回路', size: 28, bold: true, color: TEAL_TXT });
    // 下: 箇条書き2行
    s.bullets([
      '変化の速さで絞る：突然か、じわじわか',
      '気道内圧・SpO2・血圧と一緒に見る',
    ], { y: 5.1, size: 24, gap: 14 });
  }

  // ===== S3 突然消えた =====
  {
    const s = deck.slide({ title: '突然消えた', chapter: 1, notes: notesOf('S3'), sources: sourcesOf('S3') });
    const f = { x: 0.85, yt: 2.0, yb: 4.85, xr: 5.4 };
    axes(s, f);
    // 正常 3 呼吸（T=0.75）→ 平坦 約 2.0in（正常 1 呼吸の 2.7 倍）
    const x1 = waveRun(s, f.x + 0.08, f.yb, rep(3, { T: 0.75, H: 2.05 }), C.blue);
    s.polyline([[x1, f.yb], [f.xr - 0.18, f.yb]], { color: C.vermilion, width: WAVE_W });   // 突然平坦
    bulletCol(s, [
      ['回路外れ・サンプルラインの外れ'],
      ['チューブの閉塞・屈曲・事故抜管'],
      ['食道挿管（初めから出ない）'],
      [['蘇生中でも平坦 ＝ ', em('位置異常を疑う', C.red)]],
    ], { x: 5.85, w: XR - 5.85, yc: (f.yt + f.yb) / 2, gap: 0.32 });
    s.box({ x: X0, y: 5.45, w: BW, h: 0.85, fill: C.bg, line: { color: C.red, width: 2.25 }, radius: 0.08,
      text: 'まず患者：胸を見る・脈を触れる・用手換気', size: 28, bold: true, color: C.red });
  }

  // ===== S4 急に下がる =====
  {
    const s = deck.slide({ title: '急に下がる', chapter: 1, notes: notesOf('S4'), sources: sourcesOf('S4') });
    const f = { x: 0.85, yt: 2.05, yb: 6.05, xr: 5.4 };
    axes(s, f);
    const T = 0.86, H = 2.8;
    const x1 = waveRun(s, f.x + 0.08, f.yb, rep(3, { T, H }), C.blue);
    waveRun(s, x1, f.yb, rep(2, { T, H: H * 0.5 }), C.vermilion);
    bulletCol(s, [
      ['形は保たれ、高さだけ下がる', ['＝ ', em('肺血流の低下', TEAL_TXT)]],
      ['肺塞栓・空気塞栓'],
      ['急な低血圧・大量出血'],
      ['心拍出量の急な低下・心停止'],
      ['カフ漏れ・回路の漏れ（形も崩れる）'],
    ], { x: 5.85, w: XR - 5.85, yc: (f.yt + f.yb) / 2, gap: 0.3 });
  }

  // ===== S5 じわじわ下がる =====
  {
    const s = deck.slide({ title: 'じわじわ下がる', chapter: 1, notes: notesOf('S5'), sources: sourcesOf('S5') });
    // S4 と同じ図の枠。正常 2 呼吸 → 4 呼吸かけて階段状に低くなる（計 6 呼吸）
    const f = { x: 0.85, yt: 2.05, yb: 6.05, xr: 5.4 };
    axes(s, f);
    const T = 0.71;
    const x1 = waveRun(s, f.x + 0.08, f.yb, rep(2, { T, H: 2.8 }), C.blue);
    waveRun(s, x1, f.yb, [2.4, 2.0, 1.6, 1.2].map(H => ({ T, H })), C.vermilion);
    bulletCol(s, [
      ['過換気（換気量の設定が多い）'],
      ['低体温・代謝の低下'],
      ['徐々に減る心拍出量'],
      ['サンプルラインの漏れ・希釈'],
    ], { x: 5.85, w: XR - 5.85, yc: (f.yt + f.yb) / 2, gap: 0.42 });
  }

  // ===== S6 上がる =====
  {
    const s = deck.slide({ title: '上がる', chapter: 2, notes: notesOf('S6'), sources: sourcesOf('S6') });
    const f = { x: 0.85, yt: 2.0, yb: 4.95, xr: 5.4 };
    axes(s, f);
    const T = 0.73;
    const x1 = waveRun(s, f.x + 0.08, f.yb, rep(2, { T, H: 1.3 }), C.blue);
    waveRun(s, x1, f.yb, [1.55, 1.8, 2.05, 2.3].map(H => ({ T, H })), C.vermilion);
    bulletCol(s, [
      ['低換気', '（設定が少ない・自発呼吸が弱い）'],
      ['発熱・シバリング・悪性高熱症'],
      ['CO2 気腹・駆血帯の解除・重炭酸'],
      ['吸収剤の劣化（基線も上がる）'],
    ], { x: 5.85, w: XR - 5.85, y: 1.92, gap: 0.24 });
    s.box({ x: X0, y: 5.5, w: BW, h: 0.85, fill: C.bg, line: { color: C.red, width: 2.25 }, radius: 0.08,
      text: '説明のつかない上昇＋頻脈＋筋強直 → 悪性高熱症を疑う', size: 28, bold: true, color: C.red });
  }

  // ===== S7 基線が 0 に戻らない =====
  {
    const s = deck.slide({ title: '基線が 0 に戻らない', chapter: 2, notes: notesOf('S7'), sources: sourcesOf('S7') });
    const f = { x: 1.1, yt: 2.0, yb: 4.95, xr: 5.4 };
    axes(s, f, { zeroDash: true });
    s.text('0', { x: X0, y: f.yb - 0.25, w: 0.3, h: 0.5, size: 24, bold: true, color: C.gray, align: 'right' });
    const T = 0.85, e = 0.42;
    const x1 = waveRun(s, f.x + 0.08, f.yb, rep(2, { T, H: 2.1 }), C.blue);
    waveRun(s, x1, f.yb, [{ T, H: 2.3, b0: 0, b1: e }, { T, H: 2.35, b0: e }, { T, H: 2.4, b0: e }], C.vermilion);
    bulletCol(s, [
      ['二酸化炭素吸収剤の劣化（色を確認）'],
      ['一方弁の故障'],
      ['吸収剤なしの回路で流量不足'],
      ['測定の遅れ（呼吸数が多い小児）'],
    ], { x: 5.85, w: XR - 5.85, yc: (f.yt + f.yb) / 2, gap: 0.32 });
    s.box({ x: X0, y: 5.5, w: BW, h: 0.85, fill: C.bg, line: { color: C.teal, width: 2.25 }, radius: 0.08,
      text: '吸収剤を交換し、新鮮ガス流量を上げる', size: 28, bold: true, color: TEAL_TXT });
  }

  // ===== S8 プラトーが斜め（閉塞） =====
  {
    const s = deck.slide({ title: 'プラトーが斜め（閉塞）', chapter: 3, notes: notesOf('S8'), sources: sourcesOf('S8') });
    const f = { x: 0.85, yt: 2.05, yb: 5.85, xr: 5.4 };
    axes(s, f);
    const x1 = waveRun(s, f.x + 0.08, f.yb, rep(2, { T: 1.0, H: 2.7 }), C.blue);
    waveRun(s, x1, f.yb, rep(2, { T: 1.15, H: 2.55, kind: 'shark' }), C.vermilion);
    bulletCol(s, [
      ['気管支攣縮・喘息・COPD'],
      ['チューブの屈曲・', '分泌物による部分閉塞'],
      ['気道内圧の上昇・聴診とあわせて判断'],
      [['対処 ', em('→ 6-5 気道内圧上昇・換気困難', TEAL_TXT)]],
    ], { x: 5.85, w: XR - 5.85, yc: (f.yt + f.yb) / 2, gap: 0.36 });
  }

  // ===== S9 くぼみと揺れ =====
  {
    const s = deck.slide({ title: 'くぼみと揺れ', chapter: 3, notes: notesOf('S9'), sources: sourcesOf('S9') });
    const fw = 5.8, gx = BW - 2 * fw;
    const yt = 1.9, yb = 3.85;
    // 左図: プラトーのくぼみ
    const fL = { x: X0 + 0.15, yt, yb, xr: X0 + fw - 0.05 };
    axes(s, fL);
    const xl = waveRun(s, fL.x + 0.08, yb, rep(2, { T: 1.27, H: 1.5 }), C.blue);
    waveRun(s, xl, yb, rep(2, { T: 1.27, H: 1.5, kind: 'cleft' }), C.vermilion);
    // 右図: 心拍に同期した揺れ
    const rx = X0 + fw + gx;
    const fR = { x: rx + 0.15, yt, yb, xr: XR - 0.05 };
    axes(s, fR);
    const xr2 = waveRun(s, fR.x + 0.08, yb, rep(2, { T: 1.1, H: 1.5 }), C.blue);
    waveRun(s, xr2, yb, rep(2, { T: 1.43, H: 1.5, kind: 'osc', amp: 0.022, wl: 0.085 }), C.vermilion);
    // ラベル（各図の下、図の外）
    s.text('プラトーのくぼみ', { x: X0, y: yb + 0.3, w: fw, size: 24, bold: true, color: C.vermilion, align: 'center' });
    s.text('心拍に同期した揺れ', { x: rx, y: yb + 0.3, w: fw, size: 24, bold: true, color: C.vermilion, align: 'center' });
    // 下段: 箇条書き2行（全幅）
    s.bullets([
      'くぼみ ＝ 自発呼吸の努力 → 筋弛緩と麻酔深度を確認',
      '揺れ ＝ 心拍で肺が動く。徐呼吸で目立ち、異常ではない',
    ], { y: 5.15, size: 24, gap: 16 });
  }

  // ===== S10 一覧：形から原因へ =====
  {
    const s = deck.slide({ title: '一覧：形から原因へ', chapter: 4, notes: notesOf('S10'), sources: sourcesOf('S10') });
    const b = t => ({ text: t, bold: true });
    // 全セル 1 行。列幅 = 各列の最長セルの推定幅（22pt）＋左右余白 6pt（左5・右1）、合計 12.229in。
    // 「所見」7字・「最初にすること」13字は原文のまま必要幅いっぱい。「考える原因」は 15字
    // （1行目「回路外れ・閉塞・抜管・食道挿管」）が上限のため、2行目を 17字 → 15字に詰めた
    // （元「肺血流の低下（心停止を含む）、漏れ」→「肺血流低下（心停止含む）、漏れ」、助詞「の」「を」のみ削除）。
    s.table([
      ['所見', '考える原因', '最初にすること'],
      [b('突然消える'), '回路外れ・閉塞・抜管・食道挿管', '患者を見る、用手換気'],
      [b('急に下がる'), '肺血流低下（心停止含む）、漏れ', '血圧・脈波を見る'],
      [b('じわじわ下がる'), '過換気、低体温、心拍出量↓', '設定と体温を確認'],
      [b('上がる'), '低換気、発熱、気腹、駆血帯', '換気量を確認、悪性高熱症？'],
      [b('基線が上がる'), '吸収剤の劣化、弁の故障', '吸収剤交換、流量を上げる'],
      [b('プラトーが斜め'), '攣縮、部分閉塞', '気道内圧、聴診、吸引'],
      [b('くぼみ'), '自発呼吸の努力', 'TOF・麻酔深度を確認'],
    ], { x: X0, y: 1.88, colW: [2.479, 5.217, 4.533], size: 22, margin: [3, 1, 3, 5] });
  }

  // ===== S11 まとめ =====
  {
    // 出典: spec の3件は 16pt・2行に収まらない（約101字相当 > 2行93字相当）ため、
    // Kerslake の論文題名を省略し、全文はノートに残す（deviations に記載）。
    const full = SPEC.S11.sources;
    const l1 = 'Kodali BS. Anesthesiology 2013;118:192-201 / Kerslake I, Kelly F.';
    const l2 = 'BJA Education 2017;17:178-183 / Chrimes N, et al. Anaesthesia 2022;77:1395-1415';
    const S = LAYOUT.source;
    if (!fits(l1, S.size, S.w) || !fits(l2, S.size, S.w)) throw new Error('S11 出典: 2行に収まりません');
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: '出典（全文）: ' + full, sources: [l1, l2] });
    s.bullets([
      '波形があるか → 高さ → 形 → 傾向の順に見る',
      ['突然消えたら、', em('まず患者を見て用手換気', TEAL_TXT)],
      ['形を保って急に下がる ＝ ', em('肺血流の低下', TEAL_TXT)],
      ['上がり続けて頻脈・筋強直 ＝ ', em('悪性高熱症を疑う', C.red)],
      ['基線が浮く ＝ ', em('再呼吸', TEAL_TXT), '、吸収剤を交換'],
    ], { y: 1.95, size: 28, gap: 22 });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
