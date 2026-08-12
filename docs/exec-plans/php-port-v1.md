# PHP 8.0 / PhpSpreadsheet port

## Goal

Provide a PHP 8.0 implementation of the Rust `excel-agent-tool` core so a caller can inspect, extract, create, validate, and edit `.xlsx` workbooks through the same versioned JSON plan concepts and a thin `excel-tool` CLI.

## Current State

The target repository initially contained only an empty Git repository. The source implementation is the adjacent Rust repository. Its public v1 contracts live in `crates/excel-model`, workbook behavior lives in `crates/excel-engine`, and CLI behavior lives in `crates/excel-cli`.

## Scope

Included:

* PHP 8.0-compatible source code.
* `schema_version: "1.0"` RawWorkbook, CreatePlan, EditPlan, and ValidationError shapes.
* deterministic contract and operation validation;
* `.xlsx` reading, workbook creation, and immutable-input editing;
* `inspect`, `extract`, `create`, `edit`, and `validate-edit` CLI commands;
* a dependency-free PHP test harness;
* Japanese usage documentation and explicit fidelity limitations.

Excluded from this first PHP port:

* MCP transport, remote services, FFI, and LLM-provider integration;
* the Rust retrieval and semantic-model layers;
* `.xls`, `.ods`, PDF, Word, and PowerPoint;
* overwrite-in-place editing;
* dependencies other than PhpSpreadsheet and its mandatory Composer transitive dependencies.

Retrieval and semantic support was added after this initial milestone under the separate [retrieval/semantic port plan](retrieval-semantic-port.md). The exclusions above describe only the first port.

## Public Contract

The PHP port accepts the same v1 plan envelopes:

```json
{
  "schema_version": "1.0",
  "sheets": [
    {
      "name": "Data",
      "cells": [
        {"address": "A1", "value": "Hello", "value_type": "string"}
      ],
      "merged_ranges": [],
      "row_properties": [],
      "column_properties": []
    }
  ]
}
```

```json
{
  "schema_version": "1.0",
  "operations": [
    {
      "operation": "set_cell_value",
      "sheet": "Data",
      "address": "A1",
      "value": "Updated",
      "value_type": "string"
    }
  ]
}
```

CLI syntax:

```text
php bin/excel-tool inspect <input.xlsx>
php bin/excel-tool extract <input.xlsx> [--output <raw-workbook.json>]
php bin/excel-tool create <plan.json> <output.xlsx>
php bin/excel-tool validate-edit <input.xlsx> <plan.json>
php bin/excel-tool edit <input.xlsx> <plan.json> <output.xlsx>
```

Success is JSON on stdout. Failure is JSON on stderr. Exit codes are 0 for success, 1 for operational failure, 2 for arguments/JSON contract failure, and 3 for plan validation failure.

## Implementation Steps

1. Add the Composer project with only `phpoffice/phpspreadsheet` as a direct library dependency.
2. Implement contract decoding, coordinate parsing, stable errors, and create/edit validators.
3. Implement the reader, writer, and editor behind a transport-independent engine facade.
4. Implement the CLI adapter without workbook logic.
5. Add a built-in test harness that creates and reopens targeted fixtures.
6. Document setup, compatibility, and fidelity boundaries.

## Validation

```text
composer validate --strict
composer install
composer test
php bin/excel-tool --help
```

The test harness must create a workbook, reopen and inspect it, edit a separate output, verify styles/formulas/merges/dimensions, reject invalid plans, and prove protected files are not overwritten.

## Progress

* [x] Capture the PHP compatibility and dependency decision.
* [x] Define the port scope and public CLI contract.
* [x] Implement contract and validation layers.
* [x] Implement reader, writer, and editor.
* [x] Implement CLI.
* [x] Add tests and documentation.
* [x] Run PHP 8.0-compatible verification.

## Decisions

* Decision: constrain PhpSpreadsheet to `~1.30.6` and commit the resolved lock file.
  Reason: current major versions require PHP 8.1 or newer; the maintained 1.30 branch supports PHP 8.0 and 1.30.6 contains the latest security backports available when this port was created.
* Decision: use associative arrays at the public JSON boundary and small typed engine classes internally.
  Reason: this preserves exact snake_case JSON while avoiding any serialization framework dependency.
* Decision: use a custom PHP test runner rather than PHPUnit.
  Reason: the user restricted the project to PhpSpreadsheet as its only direct library dependency.
* Decision: keep retrieval and MCP out of the initial port.
  Reason: they are adapters/layers beyond the Excel engine and would materially expand the migration; the core JSON plans remain transport-neutral for later adapters.

## Discoveries and Limitations

* The machine used to create the port did not initially expose `php` or `composer` on PATH, so verification used an official portable PHP 8.0.30 runtime in a temporary directory.
* PhpSpreadsheet loads and rewrites the workbook package. It cannot offer the Rust editor's byte-for-byte passthrough of untouched ZIP parts. Charts are loaded and written when supported, but unsupported OOXML extensions, signatures, macros, or advanced features may be changed or dropped.
* PHP 8.0 is end-of-life. The dependency is pinned to a security-backported compatibility branch; deployment should isolate untrusted files and plan a PHP upgrade.

## Completion Criteria

* The listed commands are implemented.
* Public JSON and CLI behavior are documented.
* The built-in test suite passes on PHP 8.0 with Composer's PHP platform fixed to 8.0.30. (Verified: 35 assertions.)
* Invalid operations produce structured errors and never create output.
* Input workbooks and existing output paths are never overwritten.
* Known PhpSpreadsheet fidelity differences are documented.
