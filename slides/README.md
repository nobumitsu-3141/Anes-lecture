# slides/

講義スライド（.pptx）と、その元になる仕様・生成スクリプト。

```
slides/
  _lib/            組版ライブラリ（川副式書式を実装）と spec の書き方
  <番号>_<英名>/   講義ごとのディレクトリ
    spec.md        構成・本文・ノート・出典（内容の正本）
    build.js       spec を組版するスクリプト
    <番号>_<英名>.pptx  生成物
```

- 内容を直すときは spec.md を直してから build.js を実行する。
- 書式のルールと機械検査（lint）は `_lib/README.md` を参照。
- 番号は CURRICULUM.md の小項目番号に対応する。
