# Excel Agent Tool for PHP

Rust版 `excel-agent-tool` の中核を、PHP 8.0とPhpSpreadsheetで実装したExcelエンジンです。LLMや自動処理から、明示的なJSON操作だけで `.xlsx` の読取・新規作成・別名編集を行えます。RawWorkbookから決定論的なretrieval文書を生成・検索し、別レイヤーのsemantic assertionを検証する機能も含みます。

## 要件

* PHP 8.0（必要な拡張はPhpSpreadsheetのComposer要件に従います）
* Composer 2
* PhpSpreadsheet 1.30.6以降の1.30.x

直接依存する外部ライブラリは `phpoffice/phpspreadsheet` だけです。ComposerがPhpSpreadsheetの実行に必要な推移依存を自動的に導入します。PhpSpreadsheetの現行メジャー版はPHP 8.1以上を要求するため、PHP 8.0をサポートしセキュリティ修正がバックポートされた1.30系（最低1.30.6）に限定しています。`composer.lock` が実際の検証済みバージョンを固定します。

```powershell
composer install
composer test
php bin/excel-tool --help
```

## レイヤー

```text
CLI -> Engine ---------> JSON contract
          |
          +------------> PhpSpreadsheet
    -> Retrieval ------> RawWorkbook
    -> Semantic -------> RetrievalDocument
```

* `src/Contract`: JSON構文・必須フィールド・型・未知の書込スタイル属性を検証します。
* `src/Engine`: 座標、作成／編集計画、ファイル安全性を検証し、Excel処理を実行します。
* `src/Retrieval`: RawWorkbookの検証、決定論的chunk生成、provider-neutralな埋め込み／索引モデル、ローカル字句検索を扱います。
* `src/Semantic`: 意味解釈を生成せず、外部で作成されたassertionの来歴・confidence・evidenceを検証します。
* `src/Cli`: 引数とJSON入出力だけを扱う薄いアダプタです。

エンジンを直接利用する場合は `ExcelAgentTool\Engine\ExcelEngine` を使用します。PhpSpreadsheet固有型を公開JSON契約へ露出しません。

## CLI

```text
php bin/excel-tool inspect input.xlsx
php bin/excel-tool extract input.xlsx
php bin/excel-tool extract input.xlsx --output raw-workbook.json
php bin/excel-tool create create-plan.json output.xlsx
php bin/excel-tool validate-edit input.xlsx edit-plan.json
php bin/excel-tool edit input.xlsx edit-plan.json output.xlsx
php bin/excel-tool retrieve input.xlsx --source-id book-1 --source-revision rev-1 --output retrieval.json
php bin/excel-tool search retrieval.json "請求 合計" --limit 10 --sheet Data
```

PowerShellでのパス指定、全コマンド、終了コード、CreatePlan/EditPlan例は [PHPクライアント実行コマンドガイド](docs/client-command-guide.md) を参照してください。

成功結果はstdout、失敗はstderrへJSONで出力します。終了コードは、成功 `0`、実行時エラー `1`、引数／JSON契約エラー `2`、計画検証エラー `3` です。

### CreatePlan例

```json
{
  "schema_version": "1.0",
  "sheets": [
    {
      "name": "Data",
      "cells": [
        {
          "address": "A1",
          "value": "Amount",
          "value_type": "string",
          "style": {"font": {"bold": true}}
        },
        {
          "address": "B1",
          "value": 1250.5,
          "value_type": "number",
          "style": {"number_format": "#,##0.00"}
        }
      ],
      "merged_ranges": [],
      "row_properties": [],
      "column_properties": []
    }
  ]
}
```

### EditPlan例

```json
{
  "schema_version": "1.0",
  "operations": [
    {
      "operation": "set_cell_value",
      "sheet": "Data",
      "address": "B1",
      "value": 1500,
      "value_type": "number"
    },
    {
      "operation": "set_cell_style",
      "sheet": "Data",
      "address": "B1",
      "style": {"font": {"bold": true}}
    }
  ]
}
```

編集操作は `set_cell_value`、`set_cell_formula`、`set_cell_style`、`merge_cells`、`unmerge_cells`、`set_column_width`、`set_row_height`、`append_rows` をサポートします。シート作成・削除・改名はv1 Editorでは明示的な `unsupported_operation` になります。

## Retrievalとsemantic model

`retrieve` は既定の `worksheet_rows_v1` profileで、セル位置を保持したcanonical text、`source_refs`、置換用manifestを生成します。同じRawWorkbook、source、profileからはRust版と同じ `rdoc_<sha256>` IDが生成されます。`--profile profile.json` で分割上限などを変更できます。

`search` は生成済みretrieval JSONをプロセス内の字句索引へ読み込み、`source_id`、`source_revision`、`sheet` の完全一致filterを適用できます。埋め込み連携用には `RetrievalModel::embeddingRequest()`、`RetrievalModel::embeddingVector()` と検証APIを用意していますが、provider SDKやネットワーク呼出しは含みません。

semantic modelは、意味解釈をRawWorkbookへ混在させず、独立した `schema_version: "1.0"` 文書として扱います。

```php
use ExcelAgentTool\Semantic\SemanticModel;
use ExcelAgentTool\Semantic\SemanticValidator;

$knowledge = SemanticModel::knowledgeFromFile('semantic-knowledge.json');
$errors = SemanticValidator::validate($knowledge, $retrievalBuild['documents']);
```

検証エラーは `missing_required_value`、`duplicate_assertion_id`、`invalid_confidence`、`missing_evidence`、`unknown_retrieval_document`、`source_revision_mismatch`、`invalid_source_reference` の順序安定な配列です。契約と利用例の詳細は [retrieval・semantic仕様](docs/retrieval-semantic.md) を参照してください。

## 安全性

* 入力はローカルの通常 `.xlsx` ファイルだけを受け付けます。PHPストリームラッパーは拒否します。
* 編集入力は不変です。出力は必ず別パスにします。
* 既存出力は上書きしません。
* 全操作の検証が成功するまで出力を作りません。
* 同じ対象への重複・競合操作を暗黙の「後勝ち」にしません。
* 数式は評価しません。読取時はファイルに保存済みのキャッシュ値だけを観測します。

## 現在の移植範囲

Excelコア、CLI、retrieval、semantic modelを実装済みです。LLM呼出し、外部embedding provider、永続vector database、MCPアダプタは含みません。詳細は [初期移植計画](docs/exec-plans/php-port-v1.md)、[retrieval・semantic移植計画](docs/exec-plans/retrieval-semantic-port.md)、[制約](docs/limitations.md) を参照してください。
