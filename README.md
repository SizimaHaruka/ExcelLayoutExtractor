# ExcelLayoutExtractor

Excel 帳票を解析し、帳票実装向けの中間レイアウト JSON として構造化するためのプロジェクト。

このプロジェクトは Excel 互換レンダラを作ることを目的とせず、既存帳票を半自動でトレースして、Web/PDF 帳票実装の初期工数を削減することを目的とする。

## Current MVP

- `.xlsx` の読み込み
- 印刷範囲または使用範囲の解析
- 行高、列幅の mm 変換
- 罫線からの `line` 抽出
- マージ領域ベースの `rect` 抽出
- 固定文字の `text` 抽出
- `{{key}}` / `{{key|format}}` の `variable` 抽出
- 埋め込み画像の `image` 抽出
- 中間レイアウト JSON 出力
- HTML プレビュー出力

## Commands

### Demo

サンプル帳票を生成し、そのまま抽出まで実行します。

```powershell
dotnet run --project .\src\ExcelLayoutExtractor.Cli\ExcelLayoutExtractor.Cli.csproj -- demo --output-dir .\samples\demo-output
```

### Extract

任意の Excel テンプレートから JSON と HTML を生成します。

```powershell
dotnet run --project .\src\ExcelLayoutExtractor.Cli\ExcelLayoutExtractor.Cli.csproj -- extract .\samples\invoice.xlsx --output-dir .\artifacts
```

オプション:

- `--sheet <name>`: 対象シート名
- `--output-dir <dir>`: 出力先

## Output

- `*.layout.json`: 中間レイアウト JSON
- `*.preview.html`: HTML プレビュー
- `images/`: 抽出画像
- `*.warnings.txt`: 解析警告一覧

## Notes

- 現時点では `{{...}}` がセル全体を占めるケースを主対象とする
- 高度な条件付き書式、複雑なページ分岐、Excel 計算依存帳票は未対応
- `table` 推定や編集 UI は後続フェーズ

## Documents

- `docs/requirements.md`: 要件書
- `docs/specification.md`: 仕様書
- `docs/technical-requirements.md`: 技術要件書
