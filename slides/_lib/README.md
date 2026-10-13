# decklib — 川副式スライド生成ライブラリ

pptxgenjs のラッパ。川副式書式（slide-format スキル）の固定要素を自動で置き、
`slide_lint.py` で落ちる組版は **ビルド時に Error を投げて止める**。

## 実行方法

node_modules はリポジトリに置かない。pptxgenjs を入れたディレクトリを `NODE_PATH` で渡す。

```bash
B=/tmp/claude-0/-home-user-Anes-lecture/0b714c0d-c11e-535a-a97c-0e6ce9d53e16/scratchpad/pptxbuild   # npm i pptxgenjs 済み
NODE_PATH=$B/node_modules node slides/<講義>/build.js
# 見本: NODE_PATH=$B/node_modules node slides/_lib/sample_build.js $B/sample.pptx --render
```

- 環境を作り直すとき: `mkdir -p $B && cd $B && npm i pptxgenjs@4.0.1`
- `PPTXGENJS_PATH` で pptxgenjs の場所を直接指定してもよい。
- `save()` は書き出し後に `validate.py`（OOXML 検証）と `slide_lint.py` を自動実行して結果を表示する。
  スクリプトは `~/.claude/skills` 以下から自動探索（`DECKLIB_LINT` / `DECKLIB_VALIDATE` / `DECKLIB_SOFFICE` で上書き可）。
  lint 違反か検証エラーがあると `process.exitCode = 1`。

## API

```js
const { Deck, C } = require('../_lib/decklib');
const deck = new Deck({ file: '/abs/out.pptx', chapters: ['指針','7項目','心電図','血圧','現場'],
                        author: '…', deckTitle: '…' });
deck.cover({ title: '…', subtitle: '…', meta: '…', notes: '…' });   // 最初に1回。title は '\n' で2行まで
const s = deck.slide({ title: '…', chapter: 1, notes: '…', sources: ['出典A', '出典B'] });
s.bullets(['…', ['通常 ', { text: '強調', color: C.vermilion, bold: true }], { text: '下位', level: 1 }]);
s.card({ x, y, w, h, title: '見出し', lines: ['本文1', '本文2'], titleColor: C.blue });
const res = await deck.save({ render: true });   // {file, validate, lint, violations, render:{pdf, images}}
```

座標・寸法はすべてインチ。色は `'RRGGBB'`。`chapter` は `chapters` の 0 始まり添字（ラベル文字列も可）。

| メソッド | 主な引数（既定値） | 備考 |
|---|---|---|
| `s.text(str, o)` | `x,y,w,h?`, `size=24`, `color=262626`, `bold`, `align='left'`, `valign='top'`, `gap=0`(pt) | `str` は `'\n'` 区切り or 行配列。1行=1段落。余白 0 |
| `s.label(str, o)` | `size=22`, `align='center'`, `valign='middle'` | text の既定違い |
| `s.bullets(items, o)` | `x=0.55,y=1.85,w=12.2,h?`, `size=24`, `gap=10`, `marker='•'` | 1項目1段落。`{text, level:1}` で下位 |
| `s.card(o)` | `x,y,w,h`, `title`, `lines`, `fill=F2F2F2`, `radius=0.08`, `line`, `titleColor`, `titleSize=24`, `bodySize=22`, `bullets=false`, `gap=6` | 角丸箱＋見出し＋本文。文字は箱の内側（余白 0.15） |
| `s.box(o)` | `x,y,w,h`, `fill`, `line`, `radius`, `shape='rect'/'roundRect'/'ellipse'`, `text?`, `size=24` | text を渡すと図形内に文字（1図形） |
| `s.circle(o)` | `{x,y}`(左上) or `{cx,cy}`(中心), `d` | ellipse |
| `s.hline / s.vline` | `x,y,w` / `x,y,h`, `color=7F7F7F`, `width=2`(pt), `dash` | |
| `s.arrow(o)` | `x1,y1,x2,y2`, `color=404040`, `width=3`, `both` | 終点に矢じり |
| `s.polyline(pts, o)` | `[[x,y],…]` 絶対座標, `color=0072B2`, `width=3`, `closed`, `fill` | ネイティブ freeform（波形用） |
| `s.table(rows, o)` | `x,y`, `colW`, `size=22`, `header=true`, `headFill=D9D9D9`, `rowH?` | セルは文字列（`'\n'` 可）or `{text,color,bold,fill,align}`。戻り値 `{x,y,w,h,rowH}` |
| `s.notes(str)` / `s.raw` | | `raw` は素の pptxgenjs スライド（**検査対象外**） |

run（行内の部分装飾）は `{text, color, bold, italic, underline, size}`。サイズは 22 以上のみ。

補助関数: `fits(text, size, widthIn)`, `fitInfo(...)`, `maxChars(size, widthIn)`, `maxTitleChars(nChips)`,
`heightFor(nLines, size, gap)`, `minCardH(nLines, {titleSize, bodySize, gap})`, `measure`, `units`, 定数 `LAYOUT`, `C`。

## ライブラリが自動で置くもの

- 16:9（13.33×7.5in）、白背景、全テキスト メイリオ・`lang=ja-JP`。
- タイトル: 44pt 太字・金 `BF9000`・黒縁取り 2.25pt・左揃え・x0.55 y0.35 w9.3 h0.95・余白0・自動縮小なし。
  **タイトル用プレースホルダ（`<p:ph type="title"/>`, idx=0）** として出力（lint はこれでタイトルを認識）。
- 金下線: y1.52、全幅、8pt。章チップ: 0.68×0.64in、間隔0.06、右端13.13、現在章 `00A8AA`＋白太字、他 `D9D9D9`＋`404040`、11pt。
- ページ番号: 右下 12pt `7F7F7F`。**表紙を除く通し番号 = lint の `pN`** と一致。
- 出典: 16pt `595959`、y6.85（入り切らなければ項目の区切りで2行・y6.5 に上げる。3行以上は Error）。
- 表紙: タイトル中央（下端 3.55in）、金下線 y3.8、サブタイトル 28pt、meta 22pt。チップ・ページ番号なし。

## ビルド時に Error になるもの（＝lint 違反を出さないための関所）

- 本文系ヘルパで 22pt 未満。（16/12/11pt はライブラリ内部の出典・ページ番号・チップのみ）
- 1段落が枠幅に収まらない（推定幅 > 使える幅。箇条書きは行頭記号ぶんを差し引く）。
- 推定高さ（lint 式）> 枠高 + 0.3cm。`h` 省略時は必要高さを自動計算。
- タイトルが章チップに食い込む / 9.3in を超える / 改行を含む。
- 章ラベル1行が全角4字超（5字以上は `'まとめ\n確認'` のように2行に）。
- スライド枠外へのはみ出し、`chapter` 不正。
- `save()` 時: 本文ゾーンの重なり（lint と同一判定）。`save({strict:false})` で書き出しだけはできる。

警告（止めない）: 実機メイリオ推定（1行=1.5em）で枠を超える、下線より上/出典行やフッタ帯への侵入、箇条書き7個以上、斜め矢印。

## 文字数の目安（推定器: 全角=1、半角英数=0.56、半角スペース=0.3 字。幅 = 字数 × pt/72 × 1.12 in）

タイトル（44pt）の上限 = `maxTitleChars(章チップ数)`:

| 章チップ数 | 0〜4 | 5 | 6 |
|---|---|---|---|
| タイトル全角上限 | 13 | 12 | 11 |

本文（`maxChars(size, 幅)`、カードは内側余白 0.15×2 を引いた幅）:

| 枠 | 22pt | 24pt | 28pt |
|---|---|---|---|
| 全幅 12.23in | 35 | 32 | 28 |
| 箇条書き 12.2in（記号差引） | 34 | 31 | 27 |
| 半幅 6.0in | 17 | 16 | 13 |
| カード2列 w6.0（内側5.7） | 16 | 15 | 13 |
| カード3列 w3.9（内側3.6） | 10 | 9 | 8 |
| カード4列 w2.85（内側2.55） | 7 | 6 | 5 |
| 出典 16pt 11.6in | — | — | 46（16pt） |

必要な高さ: `heightFor(n, 24, 10)` = 1行 0.50 / 2行 1.14 / 3行 1.78 / 4行 2.42 / 5行 3.06 in。
見出し付きカード `minCardH(n)`（本文22pt）= 1行 1.34 / 2行 1.88 / 3行 2.43 / 4行 2.97 in。

## 落とし穴

- **重なり判定**: lint は「塗り箱・線・円・図形内テキスト（AUTO_SHAPE）」が絡む交差を違反にする。
  テキスト同士は対象外、箱の内側に完全に入った文字（card の中身）は対象外。
  斜めの矢印は外接矩形で判定される（水平/垂直なら面積0で無害）。polyline（freeform）と表は lint の対象外だが、
  表はライブラリ側で箱として重なり検査する。波形とラベルも見た目で重ねないこと（余白・引き出し線）。
  要素間は 0.6cm（0.24in）以上空けるのが書式ルール。
- 本文ゾーンは x 0.55〜12.78、y 1.6〜6.75（出典ありなら出典行の上まで）。
- `s.raw` に直接足した図形は検査されない（lint では検査される）。
- pptxgenjs はプレースホルダの位置・揃えをレイアウト側の値で上書きするため、表紙と本文でレイアウトを分けている。
- LibreOffice の描画（`--render`）はメイリオが無く代替フォント・**文字の縁取りは描かれない**。最終の実寸確認は PowerPoint で。

## 後処理（save() 内で zip を直接編集）

1. `ppt/theme/theme1.xml` の majorFont / minorFont の `latin`・`ea`・`Jpan` をメイリオに（既定の Yu Gothic / 空 ea への化けを防ぐ）。
   各 run にも `latin/ea/cs = メイリオ` が入っている。
2. タイトルの `<p:ph idx="100" type="title">`（pptxgenjs 既定）を `<p:ph type="title"/>`（idx 0）に正規化。レイアウト側も同様＋余白0。
3. pptxgenjs のバグ対策: 部分装飾した段落で `<a:pPr>` が run ごとに重複する → 先頭以外を削除。
4. 表を含むスライドで shape id が重複する → スライド内で振り直し。
5. ノートの改行を段落に分割（pptxgenjs は1つの `<a:t>` に CRLF を入れてしまう）。
