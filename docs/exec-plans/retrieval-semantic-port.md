# Retrieval and semantic model PHP port

## Goal

Add provider-neutral deterministic retrieval generation, lexical search, and independently versioned semantic assertion validation to the PHP 8.0 implementation, matching the existing Rust v1 JSON contracts and observable CLI behavior.

## Current State

The PHP repository already implements RawWorkbook extraction, CreatePlan/EditPlan processing, the Excel engine, and the `inspect`, `extract`, `create`, `validate-edit`, and `edit` CLI commands. The adjacent Rust repository defines retrieval contracts in `crates/excel-retrieval-model`, deterministic generation/search in `crates/excel-retrieval`, and semantic contracts/citation validation in `crates/excel-semantic-model`.

## Scope

Included:

* Retrieval schema v1 array contracts and strict JSON decoders.
* `worksheet_rows_v1` profile defaults and validation.
* RawWorkbook invariant validation, row/column windowing, recursive hard-limit splitting, canonical text, warnings, source references, and SHA-256 document IDs.
* Provider-neutral index records, an in-memory lexical index, and a manifest-validated process-local catalog.
* SemanticKnowledge schema v1 decoding and validation of producer metadata, assertion identity, confidence, evidence, source revisions, and range coverage.
* CLI `retrieve` and `search` commands compatible with the Rust adapter.
* PHP 8.0 tests and documentation.

Excluded:

* Calling an LLM or generating semantic assertions.
* Embedding-provider SDKs, persistent/vector databases, and network services.
* MCP transport, which remains a separate adapter task.

## Public Contract

Retrieval CLI:

```text
php bin/excel-tool retrieve <input.xlsx> --source-id <ID> --source-revision <REVISION>
  [--display-name <NAME>] [--profile <profile.json>] [--output <retrieval.json>]
php bin/excel-tool search <retrieval.json> <query> [--limit 10]
  [--source-id <ID>] [--source-revision <REVISION>] [--sheet <NAME>]
```

All retrieval and semantic top-level documents use `schema_version: "1.0"`. Retrieval IDs are `rdoc_` plus a lowercase SHA-256 digest of the Rust-compatible length-prefixed canonical identity.

Semantic validation is exposed through `ExcelAgentTool\Semantic\SemanticValidator::validate($knowledge, $documents)` and returns ordered `SemanticValidationError` arrays rather than invoking a producer.

## Implementation Steps

1. Implement retrieval/semantic JSON shape validators and structured exceptions.
2. Port deterministic retrieval builder behavior and canonical identity.
3. Port provider-neutral index record projection, lexical token scoring, and catalog replacement.
4. Add CLI parsing/output/error mapping for retrieve/search.
5. Add golden-style deterministic, limit, manifest, search, and semantic-citation tests.
6. Update public documentation and verify on PHP 8.0.30.

## Validation

```text
composer validate --strict --check-lock
composer check
composer audit --locked --no-interaction
composer check-platform-reqs --lock
php bin/excel-tool --help
```

## Progress

* [x] Extract Rust v1 contracts and algorithms.
* [x] Implement retrieval model and builder.
* [x] Implement lexical index and catalog.
* [x] Implement semantic model and validator.
* [x] Restore retrieve/search CLI commands.
* [x] Add tests and documentation.
* [x] Complete PHP 8.0 verification.

## Decisions

* Decision: preserve public contracts as associative arrays, consistent with the existing PHP RawWorkbook layer.
  Reason: exact snake_case JSON remains transparent and no serialization dependency is needed.
* Decision: reproduce Rust's canonical identity byte encoding, including unsigned 64-bit big-endian component lengths.
  Reason: identical source/profile/bounds should produce the same document IDs across Rust and PHP.
* Decision: represent embedding requests/vectors/index records as validated arrays and keep embeddings outside RetrievalDocument.
  Reason: this retains provider neutrality without introducing interfaces tied to a specific SDK.
* Decision: semantic code validates supplied assertions but never generates them.
  Reason: interpretations remain attributed, fallible output and do not contaminate factual RawWorkbook/RetrievalDocument data.

## Discoveries and Limitations

* PHP's Unicode tokenization uses PCRE Unicode letter/number classes and `mb_strtolower`; this is the closest standard-library equivalent to Rust `char::is_alphanumeric` and lowercase mapping.
* The lexical index is bounded, non-persistent, and intended for deterministic local/test use rather than production relevance ranking.
* Running both CLIs against `reader_mvp/input.xlsx` with identical source metadata produced fully equal normalized RetrievalBuild JSON. Both emitted two documents with identical IDs, canonical text, source references, warnings, and manifest.
* Verification used PHP 8.0.30: syntax checks passed for 26 files, 67 test assertions passed, Composer's strict lock validation and platform requirements passed, and the locked dependency audit found no known advisories.

## Completion Criteria

* Rust-compatible RetrievalBuild output is deterministic and source traceable.
* Profile/raw/manifest failures are explicit and structured.
* CLI retrieve/search support filters and no-overwrite JSON output.
* Semantic validation catches all Rust v1 error categories in stable order.
* No new Composer dependency is added.
* All checks pass on PHP 8.0.30 and known limitations are documented.
