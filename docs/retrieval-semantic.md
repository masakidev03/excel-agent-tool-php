# Retrieval・semantic model

## Retrieval

Retrieval文書はRawWorkbookの観測事実だけから生成されます。既定profileはRust版と同じ `worksheet_rows_v1` version `1` で、40行window、2行overlap、64列、4,000 addressed cells、24,000文字を上限とします。上限を超えたwindowは、行、列の順で再帰的に分割され、切り捨てません。分割不能な1セルが文字上限を超える場合は `text_limit_exceeded` です。

```powershell
php bin/excel-tool retrieve input.xlsx `
  --source-id book-1 `
  --source-revision sha256-or-object-version `
  --display-name "売上台帳" `
  --output retrieval.json

php bin/excel-tool search retrieval.json "顧客 売上" `
  --limit 10 `
  --source-id book-1 `
  --source-revision sha256-or-object-version `
  --sheet Data
```

source revisionは内容digestやobject versionなど、呼出し側が不変値を指定してください。出力済みJSONは上書きされません。

生成結果は次を含みます。

* `documents`: canonical text、worksheet座標、profile、source revision、厳密なsource rangeを持つ文書。
* `manifest`: 1つのsource revisionを完全置換するための文書ID列。
* `warnings`: profile設定による空sheet・数式・style・hidden cellの非致命的な除外通知。

埋め込みはRetrievalDocumentへ混在させず、派生artifactとして扱います。

```php
use ExcelAgentTool\Retrieval\RetrievalModel;

$request = RetrievalModel::embeddingRequest($documentId, 'embedding-profile-1', $canonicalText);
$vector = RetrievalModel::embeddingVector($documentId, 'embedding-profile-1', [0.1, -0.2]);
RetrievalModel::assertEmbeddingRequest($request);
RetrievalModel::assertEmbeddingVector($vector);
```

## Semantic knowledge

SemanticKnowledgeはRawWorkbookやRetrievalDocumentと独立した、fallibleかつreview可能な解釈です。`producer.kind` は `llm`、`human`、`process`、assertionの `status` は `unreviewed`、`accepted`、`rejected`、`superseded` を受け付けます。`llm` producerでは `provider`、`model`、`prompt_revision` が必須です。

```json
{
  "schema_version": "1.0",
  "knowledge_id": "knowledge-1",
  "source": {
    "source_id": "book-1",
    "source_revision": "rev-1",
    "raw_schema_version": "1.0"
  },
  "producer": {
    "kind": "llm",
    "provider": "example-provider",
    "model": "example-model",
    "prompt_revision": "prompt-1",
    "generated_at": "2026-08-13T00:00:00Z"
  },
  "assertions": [
    {
      "assertion_id": "assertion-1",
      "kind": "possible_header",
      "statement": "1行目は見出しである可能性があります。",
      "confidence": 0.8,
      "status": "unreviewed",
      "evidence": {
        "retrieval_document_ids": ["rdoc_..."],
        "source_refs": [
          {
            "sheet": "Data",
            "range": {"range": "A1:D1", "start": "A1", "end": "D1"}
          }
        ]
      }
    }
  ]
}
```

```php
use ExcelAgentTool\Semantic\SemanticModel;
use ExcelAgentTool\Semantic\SemanticValidator;

$knowledge = SemanticModel::knowledgeFromFile('semantic-knowledge.json');
$errors = SemanticValidator::validate($knowledge, $retrievalBuild['documents']);
if ($errors !== []) {
    // assertion_indexとassertion_idを含む構造化エラーとして処理する。
}
```

validatorはassertionの内容を正しいと認定しません。producer metadata、ID、confidence、evidence、source revision、source rangeの整合性だけを検証します。
