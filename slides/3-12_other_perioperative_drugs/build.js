'use strict';
/**
 * 3-12 周術期のその他の薬
 *   NODE_PATH=<pptxbuild>/node_modules node build.js
 *
 * 内容（構成・文言・ノート・出典）の正本は同じディレクトリの spec.md。
 * スライド面の文言は spec の on-slide をそのまま使い（改行位置のみ調整。brk()/q() が spec と照合する）、
 * ノートと出典は spec.md から読み込む。
 * render（PDF/JPG）はリポジトリの外に出す: DECK_RENDER_DIR（未指定なら OS の一時ディレクトリ）。
 */
const fs = require('fs');
const os = require('os');
const path = require('path');
const { Deck, C, measure, heightFor, LAYOUT } = require('../_lib/decklib');

const NAME = '3-12_other_perioperative_drugs';
const OUT = path.join(__dirname, NAME + '.pptx');
const RENDER_DIR = process.env.DECK_RENDER_DIR || path.join(os.tmpdir(), 'decklib-render', NAME);

// ---------------------------------------------------------------------------
// spec.md の読み込み（ノート・出典・スライド面の文言）
// ---------------------------------------------------------------------------
const SPEC_MD = fs.readFileSync(path.join(__dirname, 'spec.md'), 'utf8');
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
      if (inOn) {
        // 「…」の中身（入れ子の『』はそのまま）
        for (const q of ln.matchAll(/「([^」]*)」/g)) quoted.add(q[1]);
      }
    }
    out[id] = { notes, sources: sources && sources !== '—' ? sources : null, quoted };
  }
  return out;
}
const SPEC = parseSpec(SPEC_MD);
function notesOf(id) {
  const n = SPEC[id] && SPEC[id].notes;
  if (!n) throw new Error(`spec.md の ${id} に notes がありません`);
  return n;
}
/** 出典: spec の sources 行を ' / ' 区切りの項目にしてライブラリへ（ライブラリが詰めて最大2行） */
function sourcesOf(id) {
  const s = SPEC[id] && SPEC[id].sources;
  return s ? s.split(' / ') : undefined;
}
/** spec の文言そのまま（照合のみ） */
function q(id, str) {
  if (!SPEC[id].quoted.has(str)) throw new Error(`${id}: 「${str}」は spec の on-slide にありません`);
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
// 共通の小物
// ---------------------------------------------------------------------------
const X0 = LAYOUT.body.x;        // 0.55
const XR = LAYOUT.body.right;    // 12.78
const BW = XR - X0;              // 12.23
const Y0 = LAYOUT.body.y;        // 1.85
const GAP = 0.24;                // 要素間の最小間隔（0.6cm）
const r3 = v => Math.round(v * 1000) / 1000;
/** 出典の上端（2行なら 6.5、1行なら 6.85）から本文の下限を決める */
const BOTTOM_2SRC = 6.32;        // 出典2行（枠 y6.5・下揃え）のとき本文の下端
const BOTTOM_1SRC = 6.75;        // 出典1行（枠 y6.85・下揃え）のとき本文の下端

/**
 * 見出し＋本文の項目（各項目は1〜数行）を縦に並べたカード。文字はすべて箱の内側。
 * o: {x,y,w,h, title:[行], titleColor, items:[[行,...],...], pad, titleGap, itemGap, size, titleSize}
 */
function itemCard(s, o) {
  const pad = o.pad, size = o.size || 22, tsize = o.titleSize || 24;
  s.box({ x: o.x, y: o.y, w: o.w, h: o.h, fill: C.bg, radius: 0.08 });
  const iw = o.w - 2 * pad;
  const th = heightFor(o.title.length, tsize, 0);
  let cy = o.y + pad;
  s.text(o.title, { x: o.x + pad, y: cy, w: iw, h: th, size: tsize, bold: true, color: o.titleColor });
  cy += th + o.titleGap;
  o.items.forEach((lines, i) => {
    const h = heightFor(lines.length, size, 0);
    s.text(lines, { x: o.x + pad, y: cy, w: iw, h, size });
    cy += h + (i < o.items.length - 1 ? o.itemGap : 0);
  });
  if (cy > o.y + o.h - pad + 1e-6) throw new Error(`itemCard「${runText(o.title[0])}」: 高さ不足（必要 ${(cy + pad - o.y).toFixed(2)}in > ${o.h}in）`);
}
function itemCardMinW(o) {
  const size = o.size || 22, tsize = o.titleSize || 24;
  let w = Math.max(...o.title.map(t => measure(runText(t), tsize)));
  o.items.forEach(lines => lines.forEach(l => { w = Math.max(w, measure(runText(l), size)); }));
  return w + 2 * o.pad + 0.005;
}
function itemCardMinH(o) {
  const size = o.size || 22, tsize = o.titleSize || 24;
  return 2 * o.pad + heightFor(o.title.length, tsize, 0) + o.titleGap
    + o.items.reduce((a, lines) => a + heightFor(lines.length, size, 0), 0) + o.itemGap * (o.items.length - 1);
}
/**
 * 横に並べるカード: 各カードの最小幅（最長行）を満たしつつ、余りを均等に配る。
 * 高さは全カード共通（o.h、未指定なら最大の必要高さ）。gap はカード間。
 */
function cardRow(s, cards, o) {
  const common = { pad: o.pad, titleGap: o.titleGap, itemGap: o.itemGap, size: o.size, titleSize: o.titleSize };
  const cs = cards.map(c => ({ ...common, ...c }));
  const minW = cs.map(itemCardMinW);
  const free = BW - o.gap * (cs.length - 1) - minW.reduce((a, b) => a + b, 0);
  if (free < -1e-9) throw new Error(`p${s.page}: カード幅が足りません（${r3(-free)}in 不足。最小幅 ${minW.map(r3).join(', ')}）`);
  const needH = Math.max(...cs.map(itemCardMinH));
  const h = o.h != null ? o.h : Math.ceil(needH * 100) / 100;
  if (h + 1e-9 < needH) throw new Error(`p${s.page}: カードの高さ ${h}in < 必要 ${r3(needH)}in`);
  if (o.y + h > o.bottom + 1e-9) throw new Error(`p${s.page}: カード下端 ${r3(o.y + h)}in > 下限 ${o.bottom}in`);
  let x = X0;
  const widths = minW.map(w => w + free / cs.length);
  cs.forEach((c, i) => {
    itemCard(s, { ...c, x: r3(x), y: o.y, w: r3(widths[i]), h });
    x += widths[i] + o.gap;
  });
  return { h, widths, minW, needH };
}

async function main() {
  const deck = new Deck({
    file: OUT,
    chapters: ['予防的\n抗菌薬', '抗コリン\n薬', 'ステロ\nイド', 'まとめ'],
    author: 'Anes-lecture',
    deckTitle: '3-12 周術期のその他の薬',
  });

  // ===== S1 表紙 =====
  {
    const subtitle = '3-12 麻酔薬理', meta = '初期研修医向け麻酔講義';
    if (!SPEC_MD.includes('subtitle: ' + subtitle) || !SPEC_MD.includes('meta: ' + meta)) throw new Error('S1: 表紙の文言が spec と違います');
    deck.cover({ title: '周術期のその他の薬', subtitle, meta, notes: notesOf('S1') });
  }

  // ===== S2 予防的抗菌薬の目的（箇条書き4行 + 強調ボックス） =====
  {
    const id = 'S2';
    const s = deck.slide({ title: '予防的抗菌薬の目的', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    const size = 26, gap = 14;
    s.bullets([
      mark(q(id, '手術部位感染（SSI）を減らすため。治療ではない'), ['治療ではない', C.vermilion]),
      q(id, '切開の瞬間に、組織に十分な濃度があること'),
      mark(q(id, '狙うのは皮膚の常在菌（黄色ブドウ球菌など）→ セファゾリン'), ['セファゾリン', C.blue]),
      mark(q(id, '閉創後は 24 時間以内に終える（手術によっては 48 時間）'), ['24 時間以内', C.blue]),
    ], { y: 1.95, size, gap });
    const bh = heightFor(4, size, gap);
    // 強調ボックス（2行・28pt）。枠はティール（強調）
    const lines = brk(id, 'タイムアウト：執刀前 60 分以内に開始（バンコマイシンは 120 分）', '（バンコマイシン');
    const boxH = 1.25;
    const by = BOTTOM_2SRC - 0.06 - boxH;
    if (by < 1.95 + bh + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - 1.95 - bh)}in`);
    s.box({
      x: X0 + 0.6, y: r3(by), w: BW - 1.2, h: boxH, fill: C.white, radius: 0.1,
      line: { color: C.teal, width: 3.5 },
      text: [mark(lines[0], ['執刀前 60 分以内に開始', C.vermilion]), mark(lines[1], ['120 分', C.vermilion])],
      size: 28, bold: true, color: C.dark,
    });
  }

  // ===== S3 用量・再投与・アレルギー（3列カード） =====
  {
    const id = 'S3';
    const s = deck.slide({ title: '用量・再投与・アレルギー', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: brk(id, '用量（セファゾリン）', '（'), titleColor: C.blue,
        items: [
          brk(id, '成人 1 g（体重が重ければ 2 g）', '重ければ'),
          [q(id, '小児は体重あたり')],
          [q(id, '腎機能で間隔を調節')],
        ],
      },
      {
        title: [q(id, '術中の再投与')], titleColor: C.teal,
        items: [
          brk(id, '長い手術では半減期の約 2 倍ごと', '半減期'),
          brk(id, 'セファゾリンは 3〜4 時間ごと', '3〜4').map((l, i) => (i === 1 ? mark(l, ['3〜4 時間ごと', C.blue]) : l)),
          brk(id, '大量出血（1,500 mL 以上）でも', '（1,500'),
        ],
      },
      {
        title: [q(id, 'β-ラクタムアレルギー')], titleColor: C.vermilion,
        items: [
          brk(id, '重いアレルギー歴ならクリンダマイシンかバンコマイシン（腹部は陰性菌用を追加）', 'クリンダ', 'バンコ', '（腹部'),
          brk(id, '『ペニシリンアレルギー』の中身を聞く', 'の中身'),
        ],
      },
    ], { y: Y0, gap: GAP, pad: 0.15, titleGap: 0.12, itemGap: 0.26, bottom: BOTTOM_2SRC });
  }

  // ===== S4 抗菌薬のアナフィラキシー（箇条書き4行 + 強調ボックス） =====
  {
    const id = 'S4';
    const s = deck.slide({ title: '抗菌薬のアナフィラキシー', chapter: 0, notes: notesOf(id), sources: sourcesOf(id) });
    const size = 26, gap = 14;
    s.bullets([
      mark(q(id, '周術期アナフィラキシーの原因で最も多い（英国 NAP6：47%）'), ['47%', C.vermilion]),
      q(id, '麻酔導入と同じ時間帯に入るので、原因がまぎれる'),
      mark(q(id, '投与後 20 分ほどは血圧・気道内圧・皮膚を見る'), ['20 分ほど', C.blue]),
      q(id, '抗菌薬は、できれば導入の数分前に（反応を見分けやすい）'),
    ], { y: 1.95, size, gap });
    const bh = heightFor(4, size, gap);
    // 強調ボックス（1行・28pt）。危険なので枠と「アドレナリン」はレッド
    const boxH = 0.95;
    const by = BOTTOM_2SRC - 0.2 - boxH;
    if (by < 1.95 + bh + GAP - 1e-9) throw new Error(`${id}: 箇条書きと強調ボックスの間隔が ${r3(by - 1.95 - bh)}in`);
    s.box({
      x: X0 + 1.2, y: r3(by), w: BW - 2.4, h: boxH, fill: C.white, radius: 0.1,
      line: { color: C.red, width: 3.5 },
      text: [mark(q(id, '原因不明の低血圧＋気管支攣縮 → アドレナリン'), ['アドレナリン', C.red])],
      size: 28, bold: true, color: C.dark,
    });
  }

  // ===== S5 アトロピン（3列カード） =====
  {
    const id = 'S5';
    const s = deck.slide({ title: 'アトロピン', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: [q(id, '用量')], titleColor: C.blue,
        items: [
          brk(id, '添付文書：成人 0.5 mg 皮下・筋注（場合により静注）', '皮下', '（場合'),
          brk(id, '徐脈（蘇生 GL）：0.5 mg 静注を 3〜5 分ごと、総量 3 mg まで', '静注を', '総量'),
          brk(id, '小児 0.01〜0.02 mg/kg（添付文書外）', '（添付'),
        ],
      },
      {
        title: [q(id, '作用')], titleColor: C.teal,
        items: [
          brk(id, 'ムスカリン受容体の遮断', '遮断'),
          brk(id, '心拍数↑、分泌↓、気管支拡張', '気管支'),
          brk(id, '静注で 1〜2 分で効き始める', '効き'),
        ],
      },
      {
        title: [q(id, '注意')], titleColor: C.red,
        items: [
          brk(id, '成人で 0.5 mg 未満だと逆に徐脈になりうる', '逆に').map((l, i) => (i === 0 ? mark(l, ['0.5 mg 未満', C.red]) : l)),
          brk(id, '閉塞隅角緑内障、前立腺肥大による排尿障害、麻痺性イレウスは禁忌', '前立腺', '排尿', '麻痺性')
            .map((l, i) => (i === 3 ? mark(l, ['禁忌', C.red]) : l)),
          brk(id, '高齢者のせん妄、小児の発熱', '小児'),
        ],
      },
    ], { y: Y0, gap: GAP, pad: 0.12, titleGap: 0.08, itemGap: 0.12, h: BOTTOM_1SRC - Y0, bottom: BOTTOM_1SRC });
  }

  // ===== S6 抗コリン薬の使いどころ（左右2枚のカード） =====
  {
    const id = 'S6';
    const s = deck.slide({ title: '抗コリン薬の使いどころ', chapter: 1, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: [q(id, '使う場面')], titleColor: C.blue,
        items: [
          brk(id, '迷走神経反射の徐脈（眼球圧迫、腹膜牽引、気腹）', '（眼球'),
          [q(id, 'ネオスチグミンの併用（→ 3-6）')],
          [q(id, '気道操作・ケタミン前の分泌抑制')],
        ],
      },
      {
        title: [q(id, '考えること')], titleColor: C.vermilion,
        items: [
          brk(id, '徐脈の原因（低酸素、麻酔薬、刺激）を外すのが先', '外す').map((l, i) => (i === 1 ? mark(l, ['外すのが先', C.vermilion]) : l)),
          brk(id, '心拍数だけ上げても血圧が戻らないことがある', '血圧'),
          [q(id, 'エフェドリンとの使い分け（→ 3-8）')],
        ],
      },
    ], { y: Y0, gap: 0.32, pad: 0.15, titleGap: 0.16, itemGap: 0.3, bottom: BOTTOM_2SRC });
  }

  // ===== S7 ステロイドの 3 つの場面（3列カード） =====
  {
    const id = 'S7';
    const s = deck.slide({ title: 'ステロイドの 3 つの場面', chapter: 2, notes: notesOf(id), sources: sourcesOf(id) });
    cardRow(s, [
      {
        title: [q(id, 'PONV の予防')], titleColor: C.blue,
        items: [
          brk(id, 'デキサメタゾン 3.3〜6.6 mg（海外の 4〜8 mg）を導入時に', '（海外', '導入時'),
          brk(id, '効くまで時間がかかるので早めに', '早めに'),
          [q(id, '国内では適応外（→ 5-4）')],
        ],
      },
      {
        title: [q(id, 'ステロイドカバー')], titleColor: C.teal,
        items: [
          brk(id, '長期内服で副腎が抑制された患者', '抑制'),
          brk(id, 'ヒドロコルチゾン 100 mg を導入時', '100 mg').map((l, i) => (i === 1 ? mark(l, ['100 mg', C.blue]) : l)),
          brk(id, '以後 200 mg/24 時間の持続（倍量を内服できるまで）', '持続', '内服'),
        ],
      },
      {
        title: [q(id, '気道浮腫・補助')], titleColor: C.vermilion,
        items: [
          brk(id, '抜管後の気道浮腫の予防・治療', '予防'),
          brk(id, 'アナフィラキシーの二次治療（効果は遅い）', '二次', '（効果'),
          [q(id, '第一選択ではない')],
        ],
      },
    ], { y: Y0, gap: GAP, pad: 0.1, titleGap: 0.1, itemGap: 0.14, bottom: BOTTOM_2SRC });
  }

  // ===== S8 ステロイドの注意（箇条書き5行） =====
  {
    const id = 'S8';
    const s = deck.slide({ title: 'ステロイドの注意', chapter: 2, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, '高血糖：糖尿病患者では術後の血糖を見る'), ['高血糖：', C.blue]),
      mark(q(id, '創感染：単回投与で明らかな増加はないとされる'), ['創感染：', C.blue]),
      q(id, '急速静注で会陰部の灼熱感（デキサメタゾン）'),
      mark(q(id, 'カバー不足は副腎クリーゼ、過剰は高血糖・感染'), ['副腎クリーゼ', C.red]),
      [{ text: q(id, '迷ったら、カバーする方が安全'), color: C.teal, bold: true }],
    ], { y: 2.0, size: 28, gap: 22 });
  }

  // ===== S9 まとめ（箇条書き5行） =====
  {
    const id = 'S9';
    const s = deck.slide({ title: 'まとめ', chapter: 3, notes: notesOf(id), sources: sourcesOf(id) });
    s.bullets([
      mark(q(id, 'セファゾリン：執刀前 60 分以内開始、3〜4 時間ごと、24 時間以内'), ['セファゾリン：', C.blue]),
      mark(q(id, '抗菌薬は周術期アナフィラキシーの主な原因。投与後 20 分ほど見る'), ['抗菌薬', C.blue]),
      mark(q(id, 'アトロピン 0.5 mg：少量で逆に徐脈、禁忌を覚える'), ['アトロピン 0.5 mg：', C.blue]),
      mark(q(id, 'デキサメタゾン 3.3〜6.6 mg（海外の 4〜8 mg）を導入時に。適応外'), ['デキサメタゾン 3.3〜6.6 mg', C.blue]),
      mark(q(id, 'ステロイドカバー：ヒドロコルチゾン 100 mg、迷ったらカバー'), ['ステロイドカバー：', C.blue]),
    ], { y: 2.0, size: 24, gap: 28 });
  }

  const res = await deck.save({ render: true, renderDir: RENDER_DIR });
  if (res.violations !== 0 || res.validateOk === false) process.exitCode = 1;
}

main().catch(e => { console.error(e.message || e); process.exit(1); });
