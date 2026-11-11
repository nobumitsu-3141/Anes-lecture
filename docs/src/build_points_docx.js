'use strict';
// 要点集（Word）の組版。docs/src/<ch>_points.json → docs/<ch>_points.docx
//   NODE_PATH=<pptxbuild>/node_modules node build_points_docx.js ch2_points.json ../第2章_モニタリング_要点集.docx
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, LevelFormat, PageBreak, Header, Footer, PageNumber,
} = require('docx');

//   node build_points_docx.js <入力 points.json> <出力 docx> [keypoints.json]
//   第 3 引数が無ければ、入力名の _points.json を _keypoints.json に置き換えたファイルを読む。
const [,, inName, outName, kpArg] = process.argv;
const data = JSON.parse(fs.readFileSync(path.join(__dirname, inName), 'utf8'));
const kpPath = kpArg ? path.resolve(process.cwd(), kpArg) : path.join(__dirname, inName.replace(/_points\.json$/, '_keypoints.json'));
let kp;
try { kp = JSON.parse(fs.readFileSync(kpPath, 'utf8')); }
catch (e) { console.error(`要点ファイルを読めません：${kpPath}（${e.code === 'ENOENT' ? 'ファイルがありません' : e.message}）`); process.exit(1); }
const kpDecks = (kp && kp.decks) || {};
const kpMissing = data.decks.filter(d => !Array.isArray(kpDecks[d.id]) || !kpDecks[d.id].length).map(d => d.id);
if (kpMissing.length) { console.error(`要点ファイル ${kpPath} に、次の講義の要点がありません：${kpMissing.join('、')}`); process.exit(1); }
const kpAbbr = Array.isArray(kp.abbr) ? kp.abbr : [];
const FONT = { ascii: 'Meiryo', eastAsia: 'Meiryo', hAnsi: 'Meiryo', cs: 'Meiryo' };
const GOLD = 'BF9000', BLUE = '0072B2', DARK = '262626', GRAY = '595959', LIGHT = 'F2F2F2', BORDER = 'BFBFBF';
const W = 9026; // A4 本文幅（DXA）: 11906 - 1440*2
const COLW = [2600, 3800, 2626];

const run = (text, o = {}) => new TextRun({ text, font: FONT, size: o.size || 21, bold: o.bold, color: o.color || DARK, italics: o.italics });
const para = (text, o = {}) => new Paragraph({ children: [run(text, o)], spacing: { after: o.after != null ? o.after : 120, line: 300 }, alignment: o.align });
const bullet = (text, ref = 'bullets') => new Paragraph({ children: [run(text)], numbering: { reference: ref, level: 0 }, spacing: { after: 80, line: 300 } });
const border = { style: BorderStyle.SINGLE, size: 4, color: BORDER };
const borders = { top: border, bottom: border, left: border, right: border };
const cell = (text, w, o = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA }, borders,
  shading: o.head ? { type: ShadingType.CLEAR, fill: LIGHT, color: 'auto' } : undefined,
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  children: [new Paragraph({ children: [run(text, { bold: o.head, size: 19 })], spacing: { after: 0, line: 280 } })],
});
function numbersTable(rows) {
  const head = new TableRow({ tableHeader: true, children: [cell('項目', COLW[0], { head: true }), cell('値・基準', COLW[1], { head: true }), cell('出典', COLW[2], { head: true })] });
  const body = rows.map(r => new TableRow({ children: [cell(r.item, COLW[0]), cell(r.value, COLW[1]), cell(r.source, COLW[2])] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: COLW, rows: [head, ...body] });
}
function abbrTable(rows) {
  const w = [1800, 4000, 3226];
  const head = new TableRow({ tableHeader: true, children: [cell('略語', w[0], { head: true }), cell('正式な綴り', w[1], { head: true }), cell('日本語', w[2], { head: true })] });
  const body = rows.map(r => new TableRow({ cantSplit: true, children: [cell(r.abbr || '', w[0]), cell(r.full || '', w[1]), cell(r.ja || '', w[2])] }));
  return new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: w, rows: [head, ...body] });
}
const sub = text => new Paragraph({ children: [run(text, { bold: true, color: BLUE, size: 22 })], spacing: { before: 160, after: 80 } });

const children = [];
// 表紙
children.push(new Paragraph({ children: [run(data.chapter, { size: 44, bold: true, color: GOLD })], spacing: { before: 2400, after: 200 }, alignment: AlignmentType.CENTER }));
children.push(new Paragraph({ children: [run(data.subtitle, { size: 28, color: GRAY })], spacing: { after: 200 }, alignment: AlignmentType.CENTER }));
children.push(new Paragraph({ children: [run('初期研修医向け麻酔講義（離島の一人麻酔科医による）', { size: 22, color: GRAY })], alignment: AlignmentType.CENTER, spacing: { after: 400 } }));
children.push(new Paragraph({ children: [run(data.cover_note || '講義の番号・タイトル・数値は各講義のスライド（spec.md）と同じです。原典に到達できなかった箇所は【要確認】と記しています。', { size: 20, color: GRAY })], alignment: AlignmentType.CENTER }));
children.push(new Paragraph({ children: [run('構成：この章の要点（全講義ぶん）→ 各講義の解説（ねらい → 解説 → 覚える数値・基準 → 落とし穴 → 主な出典）', { size: 20, color: GRAY })], alignment: AlignmentType.CENTER, spacing: { before: 200 } }));
children.push(new Paragraph({ children: [new PageBreak()] }));
// この章の要点
children.push(new Paragraph({ text: 'この章の要点', heading: HeadingLevel.HEADING_1 }));
children.push(para('各講義で持ち帰る要点を、先にまとめて示します。かっこ内は、その要点を詳しく述べている解説の番号です。', { size: 19, color: GRAY }));
for (const d of data.decks) {
  children.push(new Paragraph({ children: [run(`${d.id}　${d.title}`, { bold: true, color: BLUE, size: 22 })], keepNext: true, spacing: { before: 160, after: 80 } }));
  for (const k of kpDecks[d.id]) {
    const kids = [run(k.text, { bold: true, size: 22 })];
    if (Array.isArray(k.ref) && k.ref.length) kids.push(run(`\u3000（解説 ${k.ref.join('・')}）`, { size: 17, color: GRAY }));
    children.push(new Paragraph({ children: kids, numbering: { reference: 'bullets', level: 0 }, keepLines: true, spacing: { after: 80, line: 300 } }));
  }
}
if (kpAbbr.length) { children.push(sub('この一覧で使う略語')); children.push(abbrTable(kpAbbr)); }
children.push(new Paragraph({ children: [new PageBreak()] }));
// この章について
children.push(new Paragraph({ text: 'この章について', heading: HeadingLevel.HEADING_1 }));
children.push(para(data.intro, { after: 200 }));
children.push(new Paragraph({ children: [new PageBreak()] }));
// 各講義
data.decks.forEach((d, i) => {
  children.push(new Paragraph({ text: `${d.id}　${d.title}`, heading: HeadingLevel.HEADING_1 }));
  children.push(sub('ねらい'));
  children.push(para(d.purpose));
  if (Array.isArray(d.summary) && d.summary.length) {
    children.push(sub('解説'));
    d.points.forEach((p, k) => {
      children.push(new Paragraph({ children: [run(`${k + 1}. ${d.summary[k]}`, { bold: true, color: BLUE, size: 21 })], keepNext: true, spacing: { before: 160, after: 60 } }));
      children.push(new Paragraph({ children: [run(p)], indent: { left: 360 }, spacing: { after: 120, line: 300 } }));
    });
  } else {
  children.push(sub('ポイント'));
  d.points.forEach((p, k) => children.push(new Paragraph({ children: [run(p)], numbering: { reference: 'numbers', level: 0, instance: i }, spacing: { after: 80, line: 300 } })));
  }
  if (d.numbers && d.numbers.length) { children.push(sub('覚える数値・基準')); children.push(numbersTable(d.numbers)); children.push(para('', { after: 40 })); }
  if (d.pitfalls && d.pitfalls.length) { children.push(sub('落とし穴')); d.pitfalls.forEach(p => children.push(bullet(p))); }
  if (d.sources && d.sources.length) { children.push(new Paragraph({ children: [run('主な出典：', { size: 18, bold: true, color: GRAY }), run(d.sources.join('／'), { size: 18, color: GRAY })], spacing: { before: 160, after: 40, line: 276 } })); }
});

const numbering = { config: [
  { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 300 } } } }] },
  { reference: 'numbers', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 360 } } } }] },
] };

const doc = new Document({
  creator: '麻酔講義', title: data.chapter + ' 要点集',
  styles: {
    default: { document: { run: { font: FONT, size: 21, color: DARK } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 32, bold: true, color: GOLD }, paragraph: { keepNext: true, spacing: { before: 480, after: 160 }, border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: GOLD, space: 4 } }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 24, bold: true, color: BLUE }, paragraph: { keepNext: true, spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering,
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ children: [run(data.chapter + '　要点集', { size: 16, color: GRAY })], alignment: AlignmentType.RIGHT })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: GRAY })], alignment: AlignmentType.CENTER })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(buf => { const out = path.resolve(__dirname, outName); fs.writeFileSync(out, buf); console.log('wrote', out, buf.length, 'bytes'); });
