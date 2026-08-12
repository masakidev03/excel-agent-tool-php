# PHPクライアント実行コマンドガイド

## 1. 実行準備

PowerShellでPHPプロジェクトへ移動し、依存関係を導入します。

```powershell
cd "C:\Users\masak\OneDrive\ドキュメント\llm-tool\excel-agent-tool-php"
composer install
php -v
php bin\excel-tool --help
```

PHPがPATHに登録されていない場合は、`php` の代わりに実行ファイルを指定します。

```powershell
& "C:\path\to\php.exe" bin\excel-tool --help
```

基本形式は次のとおりです。

```text
php bin/excel-tool <command> <arguments> [options]
```

パスに空白や日本語が含まれる場合は、PowerShellではパス全体をダブルクォートで囲みます。

## 2. Excelを読み取り、JSON化する

### ファイル情報の確認

```powershell
php bin\excel-tool inspect "C:\input\design.xlsx"
```

結果は標準出力へJSONで表示されます。ファイルの概要確認用で、完全なRawWorkbookは出力しません。

### RawWorkbook JSONをファイルへ出力

```powershell
php bin\excel-tool extract "C:\input\design.xlsx" `
  --output "C:\output\design.raw.json"
```

`--output` を省略すると、RawWorkbook JSON自体を標準出力へ表示します。

```powershell
php bin\excel-tool extract "C:\input\design.xlsx"
```

RawWorkbook JSONには、sheet名、cell address、行・列番号、値、値の型、数式、style、結合範囲、行高、列幅などの観測結果が含まれます。

### 今回のサンプルを再読取する

```powershell
$sampleDir = "outputs\019ff702-377a-7443-b39a-f00408cf7cba"

php bin\excel-tool extract `
  "$sampleDir\sample-order-management-design.xlsx" `
  --output "$sampleDir\sample-order-management-design.re-extracted.json"
```

既に存在する出力ファイルは上書きしないため、再実行時は別の出力名を指定してください。

## 3. JSONをもとに新しいExcelを作成する

`create` はRawWorkbookではなく、作成指示を表すCreatePlan JSONを受け取ります。

```powershell
php bin\excel-tool create "C:\plans\create-plan.json" "C:\output\created.xlsx"
```

最小のCreatePlan例です。

```json
{
  "schema_version": "1.0",
  "sheets": [
    {
      "name": "設計概要",
      "cells": [
        {
          "address": "A1",
          "value": "サンプル設計書",
          "value_type": "string",
          "style": {
            "font": {"bold": true},
            "fill": {"pattern": "solid", "foreground_color": "FFD9EAF7"}
          }
        },
        {
          "address": "A3",
          "value": "文書バージョン",
          "value_type": "string"
        },
        {
          "address": "B3",
          "value": "1.0",
          "value_type": "string"
        }
      ],
      "merged_ranges": [
        {"range": "A1:D1", "start": "A1", "end": "D1"}
      ],
      "row_properties": [],
      "column_properties": []
    }
  ]
}
```

## 4. 既存ExcelをJSONの指示で編集する

最初にEditPlanを検証します。

```powershell
php bin\excel-tool validate-edit "C:\input\design.xlsx" "C:\plans\edit-plan.json"
```

検証成功後、入力とは別のファイルへ出力します。

```powershell
php bin\excel-tool edit `
  "C:\input\design.xlsx" `
  "C:\plans\edit-plan.json" `
  "C:\output\design-edited.xlsx"
```

最小のEditPlan例です。

```json
{
  "schema_version": "1.0",
  "operations": [
    {
      "operation": "set_cell_value",
      "sheet": "概要",
      "address": "B6",
      "value": "1.1",
      "value_type": "string"
    },
    {
      "operation": "set_cell_style",
      "sheet": "概要",
      "address": "B6",
      "style": {"font": {"bold": true}}
    }
  ]
}
```

## 5. Retrieval JSONの生成と検索

```powershell
php bin\excel-tool retrieve "C:\input\design.xlsx" `
  --source-id "design-book-001" `
  --source-revision "rev-001" `
  --output "C:\output\design.retrieval.json"

php bin\excel-tool search "C:\output\design.retrieval.json" "受注 API" `
  --limit 10 `
  --sheet "API一覧"
```

## 6. JSON契約の使い分け

| 操作 | JSON | 用途 |
|---|---|---|
| `extract` | RawWorkbook | Excelから観測した事実 |
| `create` | CreatePlan | 新しいExcelの作成指示 |
| `validate-edit` / `edit` | EditPlan | 既存Excelへの編集指示 |
| `retrieve` | RetrievalBuild | 検索用の決定論的文書 |

`extract` が出力したRawWorkbook JSONを、そのまま `create` の入力にはできません。RawWorkbookを参照してCreatePlanまたはEditPlanを作成してください。

## 7. 標準出力と終了コード

成功結果は標準出力、失敗結果は標準エラーへJSONで出力します。

| 終了コード | 意味 |
|---:|---|
| `0` | 成功 |
| `1` | ファイルI/O・Excel処理などの実行時エラー |
| `2` | 引数またはJSON契約エラー |
| `3` | CreatePlan、EditPlan、retrieval profileなどの検証エラー |

PowerShellでは直前の終了コードを `$LASTEXITCODE` で確認できます。

```powershell
php bin\excel-tool inspect "C:\input\design.xlsx"
if ($LASTEXITCODE -ne 0) {
  throw "excel-tool failed: $LASTEXITCODE"
}
```

## 8. ファイル安全性

* 入力は通常のローカル `.xlsx` ファイルを指定します。
* `create`、`edit`、`extract --output`、`retrieve --output` は既存出力を上書きしません。
* `edit` の入力ファイルと出力ファイルは別パスにします。
* `.xlsm`、PHP stream wrapper、入力ファイルの直接上書きは対象外です。
