'use strict';
/**
 * 3-1 薬物動態の基本
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言はここに転記（改行位置のみ調整。spec にない文字列は Error）、ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, fits, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '3-1_pharmacokinetics';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
function parseSpec(md) {
  const out = {};
  for (const part of md.split(/^## /m).slice(1)) {
    const id = (part.match(/^(S\d+)\b/) || [])[1];
    if (!id) continue;
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
      if (inOn || /^(figure|layout):/.test(ln)) {
        for (const q of ln.matchAll(/「([^」]*)」/g)) quoted.add(q[1]);
      }
    }
    // 図の指示（layout / figure 行）の地の文。軸名・曲線名など「」なしで書かれた語の照合用
    const figText = lines.filter(l => /^(figure|layout):/.test(l)).join('\n');
    out[id] = { notes, sources: sources && sources !== '—' ? sources : null, quoted, figText, raw: part };
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
 * spec の on-slide（または figure/layout の「」）の文言 str を、breaks の各語の直前で改行して行配列にする。
 * str が spec に無い場合は Error（文言を変えていないことの確認）。改行位置の半角スペースは落とす。
 */
function brk(id, str, ...breaks) {
  if (!SPEC[id].quoted.has(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
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
/** spec の文言そのまま（照合のみ） */
function q(id, str) { return brk(id, str)[0]; }
/** 図の指示の地の文にある語（軸名・曲線名）。照合のみ */
function fig(id, str) {
  if (!SPEC[id].figText.includes(str)) throw new Error(`${id}: 「${str}」は spec の layout/figure にありません`);
  return str;
}
/** 1行を run に分けて一部を強調する。parts を連結すると spec の文言（行）に一致すること */
function runs(line, ...parts) {
  const joined = parts.map(p => (typeof p === 'string' ? p : p.text)).join('');
  if (joined !== line) throw new Error(`強調の分割が元の行と一致しません「${joined}」≠「${line}」`);
  return parts.map(p => (typeof p === 'string' ? { text: p } : p));
}

/**
 * 出典。既定は ' / ' 区切りの項目（ライブラリが詰めて最大2行に）。
 * breakAfter を渡すと、その文字列の直後の空白で2行に割る（文言は変えない）。
 * text を渡すと spec の代わりにその文字列を使う（表記の調整。deviations に記載）。
 */
function sourcesOf(id, breakAfter, text) {
  const s = text || (SPEC[id] && SPEC[id].sources);
  if (!s) return undefined;
  if (!breakAfter) return s.split(' / ');
  const k = s.indexOf(breakAfter);
  if (k < 0) throw new Error(`${id} 出典: 改行位置「${breakAfter}」が見つかりません`);
  const e = k + breakAfter.length;
  if (s[e] !== ' ') throw new Error(`${id} 出典: 改行位置の直後が空白ではありません`);
  let l1 = s.slice(0, e), l2 = s.slice(e + 1);
  if (l2.startsWith('/ ')) l2 = l2.slice(2);
  const S = LAYOUT.source;
  if (!fits(l1, S.size, S.w) || !fits(l2, S.size, S.w)) throw new Error(`${id} 出典: 2行に収まりません\n  ${l1}\n  ${l2}`);
  return [l1, l2];
}

// ---------------------------------------------------------------------------
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const YT = LAYOUT.body.y;        // 1.85
const BOT2 = 6.42;               // 出典2行（y 6.5）のスライドの本文下端
const BOT1 = 6.6;                // 出典1行（y 6.85）のスライドの本文下端
const r3 = v => Math.round(v * 1000) / 1000;
const em = (text, color, bold = true) => ({ text, color, bold });
const AXIS_C = C.gray;
const L22 = heightFor(1, 22, 0);   // 22pt 1行の枠高
const L24 = heightFor(1, 24, 0);

/**
 * 箇条書き（1 項目が 2 行以上になる場合用）。行頭記号と本文を別の枠に置き、折り返し行を本文の頭にそろえる。
 * items: [[行, 行...], ...]（行は文字列 or run 配列）。
 */
function bulletBlockHeight(items, size, igap) {
  return items.reduce((a, lines) => a + heightFor(lines.length, size, 0), 0) + igap * (items.length - 1);
}
function bulletBlock(s, items, o) {
  const size = o.size || 24, igap = o.igap != null ? o.igap : 0.3, ind = o.indent || (size >= 24 ? 0.42 : 0.38);
  let y = o.y;
  items.forEach(lines => {
    const h = heightFor(lines.length, size, 0);
    s.text('•', { x: o.x, y, w: ind, h: heightFor(1, size, 0), size, color: o.color });
    s.text(lines, { x: o.x + ind, y, w: o.w - ind, h, size, color: o.color });
    y += h + igap;
  });
  return y - igap - o.y;
}

/**
 * 矢印（ネイティブ freeform）。pts の折れ線を軸に、終点に塗りの三角の矢じり。
 * width(pt) に応じて矢じりを大きくする。
 */
function polyArrow(s, pts, o = {}) {
  const color = o.color || C.dark, width = o.width || 3;
  const wIn = width / 72;
  const len = o.head || Math.max(0.2, wIn * 2.6);
  const half = o.half || Math.max(0.11, wIn * 1.5);
  const tip = pts[pts.length - 1], prev = pts[pts.length - 2];
  const dx = tip[0] - prev[0], dy = tip[1] - prev[1], n = Math.hypot(dx, dy);
  const ux = dx / n, uy = dy / n, px = -uy, py = ux;
  const bx = tip[0] - ux * len, by = tip[1] - uy * len;
  // 軸は矢じりの付け根の少し先まで（太い線の角が矢じりの外に出ないように）
  const shaftEnd = [tip[0] - ux * len * 0.7, tip[1] - uy * len * 0.7];
  s.polyline([...pts.slice(0, -1), shaftEnd], { color, width });
  s.polyline([tip, [bx + px * half, by + py * half], [bx - px * half, by - py * half]],
    { color, width: 0.75, closed: true, fill: color });
}
/** 軸（線＋塗りの矢じり。始点→終点） */
function axis(s, x1, y1, x2, y2) {
  polyArrow(s, [[x1, y1], [x2, y2]], { color: AXIS_C, width: 2, head: 0.16, half: 0.08 });
}

/** 凡例 1 行: 短い線（曲線と同じ色・太さ）＋ラベル（同色・太字） */
function legend(s, o) {
  const yc = o.y + L22 / 2;
  s.polyline([[o.x, yc], [o.x + o.sw, yc]], { color: o.color, width: o.lw || 3.5 });
  s.text(o.text, { x: o.x + o.sw + 0.15, y: o.y, w: o.w, h: L22, size: 22, bold: true, color: o.color });
}

// ---------------------------------------------------------------------------
// 曲線のデータ（模式図。目盛りは置かない）
// ---------------------------------------------------------------------------
/** ボーラス後の血中濃度（2 指数）と効果部位濃度（ke0 の 1 次遅れ）。t: 分、値は Cp(0)=1 で正規化 */
function bolusCurves(T, dt) {
  const cp = t => 0.5 * Math.exp(-t / 1.0) + 0.5 * Math.exp(-t / 6);
  const ke0 = 0.7;   // 効果部位のピーク 約 1.9 分（模式）
  const out = [];
  let ce = 0;
  for (let t = 0; t <= T + 1e-9; t += dt) {
    out.push({ t, cp: cp(t), ce });
    // 1 次遅れ dCe/dt = ke0 (Cp - Ce)（細かい刻みで積分）
    const sub = 20, h = dt / sub;
    for (let k = 0; k < sub; k++) ce += ke0 * (cp(t + k * h) - ce) * h;
  }
  return out;
}
/**
 * 持続投与時間（h）→ 濃度が半分になるまでの時間（分）の概形（模式。Hughes 1992 の図の形をなぞる）
 * レミフェンタニル 約 3.5 分で一定、プロポフォール 緩やかに上がり 8 時間で 40 分未満、フェンタニル 急に上がる。
 */
const CSHT = {
  fentanyl: t => 3.5 + 258.5 * (1 - Math.exp(-Math.pow(t / 4.2, 1.6))),
  propofol: t => 3.5 + 32.5 * (1 - Math.exp(-t / 3)),
  remi: () => 3.5,
};

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['濃度', '分布', '消失', '持続', '患者差'],
    author: 'Anes-lecture',
    deckTitle: '3-1 薬物動態の基本',
  });

  // ===== S1 表紙 =====
  {
    const raw = SPEC.S1.raw;
    const title = '薬物動態の基本', subtitle = '3-1 麻酔薬理', meta = '初期研修医向け麻酔講義';
    for (const t of [title, `subtitle: ${subtitle}`, `meta: ${meta}`]) if (!raw.includes(t)) throw new Error(`S1: 「${t}」が spec にありません`);
    deck.cover({ title, subtitle, meta, notes: notesOf('S1') });
  }

  // ===== S2 濃度と効果のずれ（左: 血中濃度と効果部位濃度の模式図 / 右: 箇条書き3行） =====
  {
    const id = 'S2';
    const s = deck.slide({ title: '濃度と効果のずれ', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    // --- 図 ---
    const ax0 = 1.0, ay0 = 5.5, axTop = 2.95, axR = 7.0;
    const xs = ax0 + 0.04, xe = 6.75, T = 10;           // 0〜10 分
    const yTop = 3.2;                                    // Cp(0)=1 の高さ
    const X = t => xs + (xe - xs) * t / T;
    const Y = c => ay0 - (ay0 - yTop) * c;
    const data = bolusCurves(T, 0.05);
    const pk = data.reduce((a, d) => (d.ce > a.ce ? d : a), data[0]);
    // 効果部位濃度のピークの時刻（点線）と、静注からピークまでの遅れ（軸の下の両矢印）
    s.vline({ x: r3(X(pk.t)), y: r3(Y(pk.ce) + 0.12), h: r3(ay0 - Y(pk.ce) - 0.12), color: C.vermilion, width: 1.75, dash: 'dash' });
    const dy = ay0 + 0.24;
    s.arrow({ x1: X(0), y1: dy, x2: X(pk.t), y2: dy, color: C.vermilion, width: 2.5, both: true });
    // 軸
    axis(s, ax0, ay0, ax0, axTop);
    axis(s, ax0, ay0, axR, ay0);
    // 曲線
    s.polyline(data.map(d => [r3(X(d.t)), r3(Y(d.cp))]), { color: C.blue, width: 3.5 });
    s.polyline(data.map(d => [r3(X(d.t)), r3(Y(d.ce))]), { color: C.vermilion, width: 3.5 });
    // 軸の名前（図の外）
    s.text(fig(id, '濃度'), { x: X0, y: 1.95, w: 1.2, h: L22, size: 22 });
    s.text(fig(id, '時間'), { x: axR - 1.2, y: ay0 + 0.2, w: 1.2, h: L22, size: 22, align: 'right' });
    // 曲線のラベル（図の外・右上。色をそろえる）
    const lx = 3.85, sw = 0.55;
    legend(s, { x: lx, y: 1.95, sw, w: 2.6, color: C.blue, text: q(id, '血中濃度') });
    legend(s, { x: lx, y: 1.95 + L22 + 0.12, sw, w: 2.6, color: C.vermilion, text: q(id, '効果部位の濃度') });
    // --- 箇条書き（右） ---
    const b1 = brk(id, '薬は血中ではなく、効果部位（脳など）の濃度で効く', '（脳など）');
    const items = [
      [runs(b1[0], '薬は', em('血中', C.blue), 'ではなく、', em('効果部位', C.vermilion)), b1[1]],
      brk(id, '静注から効果のピークまで、数十秒〜数分の遅れ', '数十秒'),
      brk(id, '待たずに追加すると、重なって効きすぎる', '重なって'),
    ];
    const igap = 0.42, bx = 7.45;
    const bh = bulletBlockHeight(items, 24, igap);
    const mid = (1.95 + ay0 + 0.2 + L22) / 2;
    bulletBlock(s, items, { x: bx, y: r3(mid - bh / 2), w: XR - bx, size: 24, igap });
  }

  // ===== S3 分布と再分布（左: 区画の模式図 / 右: 箇条書き4行） =====
  {
    const id = 'S3';
    const s = deck.slide({ title: '分布と再分布', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    // --- 図: 上に脳・心臓、中央に血液、下に筋肉・脂肪 ---
    const cx = 2.725, bw = 4.3, bh = 0.72;
    const brain = { x: r3(cx - bw / 2), y: 1.95, w: bw, h: bh };
    const d = 1.4, cy = 4.08, rr = d / 2;
    const smW = 1.75, smY = 5.62;
    const muscle = { x: 0.6, y: smY, w: smW, h: bh };
    const fat = { x: r3(2 * cx - 0.6 - smW), y: smY, w: smW, h: bh };
    const boxStyle = { fill: C.bg, line: { color: C.gray, width: 1.25 }, radius: 0.1, size: 24, bold: true, color: C.dark };
    s.box({ ...brain, ...boxStyle, text: q(id, '脳・心臓（血流が多い）') });
    s.box({ ...muscle, ...boxStyle, text: q(id, '筋肉') });
    s.box({ ...fat, ...boxStyle, text: q(id, '脂肪') });
    s.circle({ cx, cy, d, fill: C.bg, line: { color: C.gray, width: 1.25 }, size: 24, bold: true, text: q(id, '血液') });
    // 矢印: 太さ＝血流の多さ（脳・心臓 太、筋肉 中、脂肪 細）。脳から血液へ戻る矢印はバーミリオン
    const gap = 0.08, off = 0.42;
    const circTopAt = dx => cy - Math.sqrt(rr * rr - dx * dx);
    polyArrow(s, [[cx - off, circTopAt(off) - gap], [cx - off, brain.y + bh + gap]], { color: C.blue, width: 11 });
    polyArrow(s, [[cx + off, brain.y + bh + gap], [cx + off, circTopAt(off) - gap]], { color: C.vermilion, width: 5 });
    const toBox = (b, w) => {
      const tx = b.x + b.w / 2, ty = b.y - gap;
      const vx = tx - cx, vy = ty - cy, n = Math.hypot(vx, vy);
      const sx = cx + vx / n * (rr + gap), sy = cy + vy / n * (rr + gap);
      polyArrow(s, [[sx, sy], [tx, ty]], { color: C.blue, width: w });
    };
    toBox(muscle, 6);
    toBox(fat, 2.5);
    // --- 箇条書き（右、1項目1行） ---
    const b2 = q(id, '次に筋肉へ移る（再分布）→ 効果が切れる');
    const b3 = q(id, '単回静注の導入薬は、主に再分布で切れる');
    const items = [
      q(id, '血流の多い脳・心臓にまず届く（効く）'),
      runs(b2, '次に筋肉へ移る（', em('再分布', C.vermilion), '）→ 効果が切れる'),
      runs(b3, '単回静注の導入薬は、主に', em('再分布', C.vermilion), 'で切れる'),
      q(id, '繰り返すと脂肪にたまり、切れにくくなる'),
    ];
    const bx = 5.18, gapPt = 26;
    const hB = heightFor(4, 24, gapPt);
    const mid = (brain.y + smY + bh) / 2;
    s.bullets(items, { x: bx, y: r3(mid - hB / 2), w: XR - bx, size: 24, gap: gapPt });
  }

  // ===== S4 消失：代謝と排泄（2×2 のカード） =====
  {
    const id = 'S4';
    const s = deck.slide({
      title: '消失：代謝と排泄', chapter: 2, notes: notesOf(id),
      sources: sourcesOf(id, 'Proost JH, et al.'),
    });
    const gx = 0.25, gy = 0.25;
    const cw = (XR - X0 - gx) / 2;
    const h1 = 2.32, h2 = BOT2 - (YT + h1 + gy);
    const cards = [
      { title: q(id, '肝臓で代謝'), titleColor: C.blue,
        lines: [q(id, 'プロポフォール、ミダゾラム、'), q(id, 'レミマゾラム'), q(id, '肝血流・肝機能で変わる')] },
      { title: q(id, '血液・組織の酵素'), titleColor: C.teal,
        lines: [q(id, 'レミフェンタニル'), q(id, '臓器の機能によらない')] },
      { title: q(id, '腎臓から排泄'), titleColor: C.vermilion,
        lines: [q(id, 'モルヒネ：活性代謝物がたまる'), q(id, 'ロクロニウム：一部は尿へ')] },
      { title: q(id, '肺から排泄'), titleColor: C.blue,
        lines: [q(id, '吸入麻酔薬'), q(id, '換気量と溶けやすさで決まる')] },
    ];
    cards.forEach((c, i) => {
      const col = i % 2, row = Math.floor(i / 2);
      s.card({
        x: r3(X0 + col * (cw + gx)), y: r3(row === 0 ? YT : YT + h1 + gy), w: r3(cw), h: r3(row === 0 ? h1 : h2),
        gap: 0, ...c,
      });
    });
  }

  // ===== S5 持続投与の切れ方（左: context-sensitive half-time の模式図 / 右: 箇条書き3行） =====
  {
    const id = 'S5';
    const s = deck.slide({
      title: '持続投与の切れ方', chapter: 3, notes: notesOf(id),
      sources: sourcesOf(id, 'Kapila A, et al. Anesthesiology'),
    });
    const ax0 = 1.0, ay0 = 5.6, axTop = 2.5;
    const xs = ax0 + 0.04, xe = 4.45, Tmax = 8;
    const axR = xe + 0.2;
    const X = t => xs + (xe - xs) * t / Tmax;
    // 縦: 模式図（目盛りなし）。3 本の順序と形を保ったまま、低い 2 本の終点のラベルが重ならないよう 0.8 乗で圧縮
    const base = 0.22, span = 2.7, cmin = 3.5, cmax = CSHT.fentanyl(Tmax);
    const Y = c => ay0 - (base + span * Math.pow(Math.max(0, c - cmin) / (cmax - cmin), 0.8));
    const curve = key => {
      const pts = [];
      for (let k = 0; k <= 160; k++) { const t = Tmax * k / 160; pts.push([r3(X(t)), r3(Y(CSHT[key](t)))]); }
      return pts;
    };
    axis(s, ax0, ay0, ax0, axTop);
    axis(s, ax0, ay0, axR, ay0);
    const series = [
      { key: 'fentanyl', color: C.vermilion, label: 'フェンタニル' },
      { key: 'propofol', color: C.blue, label: 'プロポフォール' },
      { key: 'remi', color: C.teal, label: 'レミフェンタニル' },
    ];
    series.forEach(sr => s.polyline(curve(sr.key), { color: sr.color, width: 3.5 }));
    // ラベル（図の右端の外。各曲線の終点の高さ）
    const lx = axR + 0.15;
    series.forEach(sr => {
      const yEnd = Y(CSHT[sr.key](Tmax));
      s.text(fig(id, sr.label), { x: r3(lx), y: r3(yEnd - L22 / 2), w: 2.8, h: L22, size: 22, bold: true, color: sr.color });
    });
    // 軸の名前（図の外）
    s.text(q(id, '半分になるまでの時間'), { x: X0, y: 1.92, w: 3.6, h: L22, size: 22 });
    s.text(q(id, '持続投与した時間'), { x: r3(axR - 3.0), y: ay0 + 0.2, w: 3.0, h: L22, size: 22, align: 'right' });
    // --- 箇条書き（右） ---
    const b2 = brk(id, 'レミフェンタニルは何時間入れても数分で切れる', '何時間');
    const b3 = brk(id, 'フェンタニルは長く入れるほど、切れにくくなる', '入れるほど');
    const items = [
      brk(id, '止めてから濃度が半分になる時間は、入れた時間で変わる', '時間は、'),
      [runs(b2[0], em('レミフェンタニル', C.teal), 'は'), b2[1]],
      [runs(b3[0], em('フェンタニル', C.vermilion), 'は長く'), b3[1]],
    ];
    const igap = 0.45, bx = 7.85;
    const bh = bulletBlockHeight(items, 22, igap);
    const mid = (1.92 + ay0 + 0.2 + L22) / 2;
    bulletBlock(s, items, { x: bx, y: r3(mid - bh / 2), w: XR - bx, size: 22, igap });
  }

  // ===== S6 用量反応と相乗作用（左: 2 本の S 字曲線 / 右: 箇条書き3行） =====
  {
    const id = 'S6';
    const s = deck.slide({ title: '用量反応と相乗作用', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    const ax0 = 1.6, ay0 = 5.75, axTop = 3.0, axR = 7.0;
    const xs = ax0 + 0.04, xe = 6.8;
    const y100 = 3.3;
    const X = u => xs + (xe - xs) * u;               // u: 0〜1（用量、模式）
    const Y = p => ay0 - (ay0 - y100) * p;           // p: 0〜1（効く人の割合）
    const sig = (u, c) => 1 / (1 + Math.exp(-(u - c) / 0.055));
    const cB = 0.64, cV = 0.36;                       // 鎮静薬だけ（右）/ オピオイド併用（左に移動）
    // 50% と 95% の横の点線
    for (const p of [0.5, 0.95]) s.hline({ x: ax0, y: r3(Y(p)), w: r3(xe - ax0), color: AXIS_C, width: 1.75, dash: 'dash' });
    axis(s, ax0, ay0, ax0, axTop);
    axis(s, ax0, ay0, axR, ay0);
    const curve = c => {
      const pts = [];
      for (let k = 0; k <= 120; k++) { const u = k / 120; pts.push([r3(X(u)), r3(Y(sig(u, c)))]); }
      return pts;
    };
    s.polyline(curve(cB), { color: C.blue, width: 3.5 });
    s.polyline(curve(cV), { color: C.vermilion, width: 3.5 });
    // 左への移動（25% の高さで、青の曲線から朱の曲線へ）
    const p25 = 0.25, du = 0.055 * Math.log(p25 / (1 - p25));
    polyArrow(s, [[X(cB + du) - 0.15, Y(p25)], [X(cV + du) + 0.15, Y(p25)]], { color: C.vermilion, width: 3 });
    // ED50 / ED95（縦軸の外、点線の高さ）
    const tw = 0.95;
    s.label(q(id, 'ED95'), { x: r3(ax0 - 0.12 - tw), y: r3(Y(0.95) - L22 / 2), w: tw, h: L22, align: 'right', bold: true });
    s.label(q(id, 'ED50'), { x: r3(ax0 - 0.12 - tw), y: r3(Y(0.5) - L22 / 2), w: tw, h: L22, align: 'right', bold: true });
    // 軸の名前（図の外）
    s.text(fig(id, '効く人の割合'), { x: X0, y: 1.95, w: 2.4, h: L22, size: 22 });
    s.text(fig(id, '用量'), { x: axR - 1.2, y: ay0 + 0.15, w: 1.2, h: L22, size: 22, align: 'right' });
    // 曲線のラベル（図の外・右上）
    const lx = 4.0, sw = 0.55;
    legend(s, { x: lx, y: 1.95, sw, w: 2.5, color: C.blue, text: fig(id, '鎮静薬だけ') });
    legend(s, { x: lx, y: 1.95 + L22 + 0.12, sw, w: 2.5, color: C.vermilion, text: fig(id, 'オピオイド併用') });
    // --- 箇条書き（右） ---
    const b2 = brk(id, '鎮静薬とオピオイドは相乗：併用で必要量が減る', '併用で');
    const items = [
      brk(id, 'ED50・ED95 は集団の値。個人差は大きい', '個人差'),
      [runs(b2[0], '鎮静薬とオピオイドは', em('相乗', C.vermilion), '：'), b2[1]],
      brk(id, '効き目は反応で見る：脳波、血圧、TOF', '脳波'),
    ];
    const igap = 0.45, bx = 7.45;
    const bh = bulletBlockHeight(items, 24, igap);
    const mid = (1.95 + ay0 + 0.15 + L22) / 2;
    bulletBlock(s, items, { x: bx, y: r3(mid - bh / 2), w: XR - bx, size: 24, igap });
  }

  // ===== S7 患者による違い（表 3列 × 6行＋見出し行） =====
  {
    const id = 'S7';
    const s = deck.slide({
      title: '患者による違い', chapter: 4, notes: notesOf(id),
      sources: sourcesOf(id, 'Panhuizen IF,'),
    });
    const c = str => q(id, str);
    const b = str => ({ text: q(id, str), bold: true });
    const rows = [
      [c('患者'), c('何が変わるか'), c('どうするか')],
      [b('高齢者'), c('感受性↑、分布と代謝↓'), c('減らして、ゆっくり')],
      [b('肥満'), c('脂溶性の薬の分布容積↑'), c('体重の使い分け（次）')],
      [b('腎機能低下'), c('活性代謝物が蓄積'), c('モルヒネを避ける')],
      [b('肝機能低下'), c('肝で処理される薬が遅れる'), brk(id, 'ミダゾラム・ロクロニウム遷延', 'ロクロニウム').join('\n')],
      [b('循環不全・出血'), c('効果部位に届くのが遅く、強く効く'), c('少量から、待って追加')],
      [b('小児（乳幼児）'), brk(id, '分布容積大、代謝が速い（新生児は未熟で遅い）', '（新生児').join('\n'),
        brk(id, 'mg/kg は多め、間隔は短め', '間隔').join('\n')],
    ];
    s.table(rows, { x: X0, y: YT, colW: [2.6, 5.85, 3.78], size: 22, margin: [2, 6, 2, 6] });
  }

  // ===== S8 体重の使い分け（表 2列 × 4行＋見出し行、下に一行の注意） =====
  {
    const id = 'S8';
    // 出典: spec の 4 件は 16pt・2 行（11.6in×2）にどの空白で割っても収まらない（最良の割り方で 1 行目が 0.09in 超過）。
    // 全角の丸括弧「（2025）」を半角「(2025)」にして 2 行に収める（文言は同じ）。
    const src = SPEC[id].sources.replace('（2025）', '(2025)');
    if (src === SPEC[id].sources) throw new Error('S8 出典の表記調整に失敗');
    const s = deck.slide({
      title: '体重の使い分け', chapter: 4, notes: notesOf(id),
      sources: sourcesOf(id, 'Horrow JC, et al. BMC', src),
    });
    const c = str => q(id, str);
    const b = str => ({ text: q(id, str), bold: true });
    const rows = [
      [c('体重の種類'), c('使う場面の例')],
      [b('総体重（実体重）'), c('スガマデクス、プロポフォール維持')],
      [b('理想体重（身長から）'), c('ロクロニウム、局所麻酔薬の最大用量')],
      [b('除脂肪体重'), c('プロポフォール導入、レミフェンタニル')],
      [{ text: brk(id, '補正体重（理想＋超過分の一部）', '（理想').join('\n'), bold: true }, c('施設の方式に従う')],
    ];
    const t = s.table(rows, { x: X0, y: YT, colW: [4.6, 7.63], size: 24 });
    const ny = r3(t.y + t.h + 0.32);
    s.box({
      x: X0, y: ny, w: XR - X0, h: r3(BOT2 - ny), fill: C.bg, radius: 0.08, margin: 0.2,
      text: q(id, 'どの体重を使うかは薬と施設の方式で決まる。迷ったら指導医に'), size: 24, align: 'left', color: C.dark,
    });
  }

  // ===== S9 現場での使い方（箇条書き5行） =====
  {
    const id = 'S9';
    const s = deck.slide({ title: '現場での使い方', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    const items = [
      q(id, '入れたら、効果のピークまで待つ'),
      q(id, '反応を見て、少量ずつ追加する'),
      q(id, '持続投与は、切れ方を考えて止める時刻を決める'),
      q(id, '高齢者・循環不全は、減らして（目安 半量）ゆっくり'),
      q(id, '併用するなら、それぞれの量を減らす'),
    ];
    const gapPt = 24, h = heightFor(5, 28, gapPt);
    s.bullets(items, { x: X0, y: r3(YT + (BOT2 - YT - h) / 2), w: 12.2, size: 28, gap: gapPt });
  }

  // ===== S10 まとめ（箇条書き5行） =====
  {
    const id = 'S10';
    const s = deck.slide({ title: 'まとめ', chapter: 4, notes: notesOf(id), sources: sourcesOf(id) });
    const items = [
      q(id, '効くのは効果部位の濃度。静注から効果まで数分の遅れ'),
      q(id, '単回静注は主に再分布で切れ、繰り返すと蓄積する'),
      q(id, '持続投与の切れ方は、入れた時間で変わる'),
      q(id, '鎮静薬とオピオイドは相乗。併用で量を減らす'),
      q(id, '高齢・肥満・腎肝・循環不全で、量と速さを変える'),
    ];
    const gapPt = 24, h = heightFor(5, 28, gapPt);
    s.bullets(items, { x: X0, y: r3(YT + (BOT2 - YT - h) / 2), w: 12.2, size: 28, gap: gapPt });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.render && res.render.images) console.log('[build] render: ' + (res.render.images[0] ? path.dirname(res.render.images[0]) : '(none)'));
  if (res.violations !== 0) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
