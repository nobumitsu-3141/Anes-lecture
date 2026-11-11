'use strict';
// 章の要点（docs/src/ch<N>_keypoints.json）の機械検査。
//   node docs/src/check_keypoints.js          全章
//   node docs/src/check_keypoints.js 11 4     第 11 章と第 4 章だけ
// 誤り（ERROR）が 1 件でもあれば終了コード 1。注意（WARN）は目で確かめる項目。
const fs = require('fs');
const path = require('path');

const MAX_LEN = 40;          // 要点 1 個の字数の上限（全角・半角とも 1 字）
const MIN_N = 3, MAX_N = 4;  // 1 講義あたりの要点の数
// 略語として登録しなくてよい英字（単位と、綴りを添える必要のない記号）
const NO_ABBR = new Set(['mg', 'kg', 'g', 'mL', 'dL', 'L', 'mmHg', 'cmH2O', 'kPa', 'Torr', 'min', 'h', 'mEq', 'mmol', 'mol', 'mOsm', 'J', 'Fr', 'G', 'mm', 'cm', 'm', 'kcal', 'W', 'Hz', 'mA', 'mV', 'U', 'pH', 'ABO', 'Rh', 'A', 'B', 'AB', 'O', 'D', 'X', 'vs', 'I', 'II', 'III', 'IV', 'V', 'VI']);

const args = process.argv.slice(2);
const chapters = args.length ? args : Array.from({ length: 11 }, (_, i) => String(i + 1));
let errors = 0, warns = 0, total = 0;
const err = m => { errors++; console.log('ERROR ' + m); };
const warn = m => { warns++; console.log('WARN  ' + m); };
const norm = s => s.replace(/,/g, '').replace(/[０-９]/g, c => String.fromCharCode(c.charCodeAt(0) - 0xFEE0));

for (const n of chapters) {
  const srcPath = path.join(__dirname, `ch${n}_points.json`);
  const kpPath = path.join(__dirname, `ch${n}_keypoints.json`);
  if (!fs.existsSync(kpPath)) { err(`第${n}章: ${path.basename(kpPath)} がない`); continue; }
  let kp;
  try { kp = JSON.parse(fs.readFileSync(kpPath, 'utf8')); } catch (e) { err(`第${n}章: JSON として読めない（${e.message}）`); continue; }
  const src = JSON.parse(fs.readFileSync(srcPath, 'utf8'));
  const chapterText = norm(JSON.stringify(src)).toLowerCase();
  const abbrList = Array.isArray(kp.abbr) ? kp.abbr : [];
  if (!kp.decks || typeof kp.decks !== 'object') { err(`第${n}章: decks がない`); continue; }
  const ids = src.decks.map(d => d.id);
  for (const id of Object.keys(kp.decks)) if (!ids.includes(id)) err(`${id}: 元の要点集にない講義番号`);
  const seen = new Map();
  const used = new Set();
  let count = 0;
  for (const d of src.decks) {
    const list = kp.decks[d.id];
    if (!Array.isArray(list)) { err(`${d.id}: 要点がない`); continue; }
    if (list.length < MIN_N || list.length > MAX_N) err(`${d.id}: 要点が ${list.length} 個（${MIN_N}〜${MAX_N} 個にする）`);
    for (const [i, k] of list.entries()) {
      const tag = `${d.id} の ${i + 1} 個目`;
      count++;
      if (!k || typeof k.text !== 'string' || !k.text.trim()) { err(`${tag}: text がない`); continue; }
      const text = k.text;
      const len = Array.from(text).length;
      if (len > MAX_LEN) err(`${tag}: ${len} 字（上限 ${MAX_LEN} 字）「${text}」`);
      if (/要確認/.test(text)) err(`${tag}: 【要確認】の内容は要点にしない「${text}」`);
      if (/。$/.test(text)) err(`${tag}: 文末の「。」は付けない「${text}」`);
      if (/\n/.test(text)) err(`${tag}: 改行を入れない`);
      if (seen.has(text)) err(`${tag}: ${seen.get(text)} と同じ文`); else seen.set(text, tag);
      const ref = k.ref;
      if (!Array.isArray(ref) || !ref.length) { err(`${tag}: ref（解説の番号）がない`); continue; }
      const bad = ref.filter(r => !Number.isInteger(r) || r < 1 || r > d.points.length);
      if (bad.length) { err(`${tag}: ref ${JSON.stringify(bad)} は解説の番号の範囲外（1〜${d.points.length}）`); continue; }
      if (new Set(ref).size !== ref.length) err(`${tag}: ref に同じ番号が 2 回ある`);
      // 数値は、指した解説（見出し文と本文）か、その講義の数値表・落とし穴に同じ形で載っていること
      const refText = norm(ref.map(r => (d.summary ? d.summary[r - 1] : '') + '\n' + d.points[r - 1]).join('\n'));
      const deckText = norm(JSON.stringify([d.numbers || [], d.pitfalls || []]));
      for (const num of norm(text).match(/\d+(?:\.\d+)?/g) || []) {
        const re = new RegExp(`(^|[^\\d.])${num.replace('.', '\\.')}(?![\\d]|\\.\\d)`);
        if (re.test(refText)) continue;
        if (re.test(deckText)) warn(`${tag}: 数値 ${num} は指した解説になく、数値表・落とし穴にある「${text}」`);
        else err(`${tag}: 数値 ${num} が指した解説にも数値表にもない「${text}」`);
      }
      // 英字の略語は abbr に登録する
      for (const tok of text.match(/[A-Za-z][A-Za-z0-9]*/g) || []) {
        if (NO_ABBR.has(tok)) continue;
        const hit = abbrList.find(a => a && typeof a.abbr === 'string' && a.abbr.split(/[^A-Za-z0-9]+/).includes(tok));
        if (hit) used.add(hit.abbr); else err(`${tag}: 英字「${tok}」が abbr に登録されていない「${text}」`);
      }
    }
  }
  for (const a of abbrList) {
    if (!a || !a.abbr || !a.ja) { err(`第${n}章 abbr: abbr と ja は必須 ${JSON.stringify(a)}`); continue; }
    if (!used.has(a.abbr)) err(`第${n}章 abbr: 「${a.abbr}」はどの要点にも出てこない`);
    if (a.full && !chapterText.includes(norm(a.full).toLowerCase())) warn(`第${n}章 abbr: 「${a.abbr}」の綴り「${a.full}」が章の本文に見当たらない（綴りを確かめる）`);
    if (!a.full) warn(`第${n}章 abbr: 「${a.abbr}」に綴り（full）がない`);
  }
  total += count;
  console.log(`第${n}章: ${src.decks.length} 講義・要点 ${count} 個・略語 ${abbrList.length} 個`);
}
console.log(`合計: 要点 ${total} 個、ERROR ${errors} 件、WARN ${warns} 件`);
process.exit(errors ? 1 : 0);
