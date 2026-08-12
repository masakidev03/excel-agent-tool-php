<?php

declare(strict_types=1);

namespace ExcelAgentTool\Retrieval;

use ExcelAgentTool\Contract\ContractException;
use JsonException;

final class RetrievalModel
{
    /** @return array<string, mixed> */
    public static function defaultProfile(): array
    {
        return [
            'schema_version' => '1.0',
            'profile' => 'worksheet_rows_v1',
            'profile_version' => '1',
            'target_rows' => 40,
            'overlap_rows' => 2,
            'max_columns' => 64,
            'max_addressed_cells' => 4000,
            'max_text_chars' => 24000,
            'include_formulas' => true,
            'include_basic_style_facts' => false,
            'include_hidden_cells' => true,
            'include_empty_sheets' => false,
        ];
    }

    /** @return array<string, mixed> */
    public static function source(string $sourceId, string $sourceRevision, ?string $displayName = null): array
    {
        $source = [
            'schema_version' => '1.0',
            'source_id' => $sourceId,
            'source_revision' => $sourceRevision,
        ];
        if ($displayName !== null) {
            $source['display_name'] = $displayName;
        }

        return $source;
    }

    /** @return array{document_id: string, embedding_profile_id: string, text: string} */
    public static function embeddingRequest(string $documentId, string $embeddingProfileId, string $text): array
    {
        return [
            'document_id' => $documentId,
            'embedding_profile_id' => $embeddingProfileId,
            'text' => $text,
        ];
    }

    /**
     * @param list<int|float> $values
     * @return array{document_id: string, embedding_profile_id: string, dimensions: int, values: list<float>}
     */
    public static function embeddingVector(string $documentId, string $embeddingProfileId, array $values): array
    {
        $normalized = [];
        foreach ($values as $index => $value) {
            if ((!is_int($value) && !is_float($value)) || !is_finite((float) $value)) {
                self::fail("embedding vector.values[$index] must be a finite number");
            }
            $normalized[] = (float) $value;
        }
        $vector = [
            'document_id' => $documentId,
            'embedding_profile_id' => $embeddingProfileId,
            'dimensions' => count($normalized),
            'values' => $normalized,
        ];
        self::assertEmbeddingVector($vector);

        return $vector;
    }

    /** @param array<string, mixed> $request */
    public static function assertEmbeddingRequest(array $request, string $context = 'embedding request'): void
    {
        self::exact($request, ['document_id', 'embedding_profile_id', 'text'], $context);
        foreach (['document_id', 'embedding_profile_id', 'text'] as $key) {
            self::string($request, $key, $context);
        }
    }

    /** @param array<string, mixed> $vector */
    public static function assertEmbeddingVector(array $vector, string $context = 'embedding vector'): void
    {
        self::exact($vector, ['document_id', 'embedding_profile_id', 'dimensions', 'values'], $context);
        self::string($vector, 'document_id', $context);
        self::string($vector, 'embedding_profile_id', $context);
        $dimensions = self::integer($vector, 'dimensions', $context);
        $values = self::list($vector, 'values', $context);
        foreach ($values as $index => $value) {
            if ((!is_int($value) && !is_float($value)) || !is_finite((float) $value)) {
                self::fail("$context.values[$index] must be a finite number");
            }
        }
        if ($dimensions !== count($values)) {
            self::fail("$context.dimensions must equal the number of values");
        }
    }

    /** @param array<string, mixed> $record */
    public static function assertIndexRecord(array $record, string $context = 'index record'): void
    {
        self::exactOptional(
            $record,
            ['document_id', 'text', 'source_id', 'source_revision', 'sheet', 'row_start', 'row_end'],
            ['sheet_index', 'embedding'],
            $context
        );
        foreach (['document_id', 'text', 'source_id', 'source_revision', 'sheet'] as $key) {
            self::string($record, $key, $context);
        }
        self::optionalInteger($record, 'sheet_index', $context);
        self::integer($record, 'row_start', $context);
        self::integer($record, 'row_end', $context);
        if (array_key_exists('embedding', $record) && $record['embedding'] !== null) {
            self::assertEmbeddingVector(self::object($record['embedding'], "$context.embedding"), "$context.embedding");
        }
    }

    /** @param array<string, mixed> $query */
    public static function assertSearchQuery(array $query, string $context = 'search query'): void
    {
        self::exactOptional($query, ['text', 'limit'], ['filter'], $context);
        self::string($query, 'text', $context);
        self::integer($query, 'limit', $context);
        if (array_key_exists('filter', $query)) {
            $filter = self::object($query['filter'], "$context.filter");
            self::exactOptional($filter, [], ['source_id', 'source_revision', 'sheet'], "$context.filter");
            foreach (['source_id', 'source_revision', 'sheet'] as $key) {
                self::optionalString($filter, $key, "$context.filter");
            }
        }
    }

    /** @return array<string, mixed> */
    public static function profileFromFile(string $path): array
    {
        $profile = self::readObject($path, 'retrieval profile');
        self::assertProfile($profile, 'retrieval profile');

        return $profile;
    }

    /** @return array<string, mixed> */
    public static function buildFromFile(string $path): array
    {
        $build = self::readObject($path, 'retrieval build');
        self::assertBuild($build);

        return $build;
    }

    /** @param array<string, mixed> $profile */
    public static function assertProfile(array $profile, string $context = 'profile'): void
    {
        self::exact($profile, [
            'schema_version', 'profile', 'profile_version', 'target_rows', 'overlap_rows',
            'max_columns', 'max_addressed_cells', 'max_text_chars', 'include_formulas',
            'include_basic_style_facts', 'include_hidden_cells', 'include_empty_sheets',
        ], $context);
        self::version($profile, $context);
        self::string($profile, 'profile', $context);
        self::string($profile, 'profile_version', $context);
        foreach (['target_rows', 'overlap_rows', 'max_columns', 'max_addressed_cells', 'max_text_chars'] as $key) {
            self::integer($profile, $key, $context);
        }
        foreach (['include_formulas', 'include_basic_style_facts', 'include_hidden_cells', 'include_empty_sheets'] as $key) {
            self::boolean($profile, $key, $context);
        }
    }

    /** @param array<string, mixed> $build */
    public static function assertBuild(array $build): void
    {
        self::exact($build, ['schema_version', 'documents', 'manifest', 'warnings'], 'retrieval build');
        self::version($build, 'retrieval build');
        $documents = self::list($build, 'documents', 'retrieval build');
        foreach ($documents as $index => $document) {
            self::assertDocument(self::object($document, "documents[$index]"), "documents[$index]");
        }
        self::assertManifest(self::requiredObject($build, 'manifest', 'retrieval build'));
        $warnings = self::list($build, 'warnings', 'retrieval build');
        foreach ($warnings as $index => $warning) {
            $context = "warnings[$index]";
            $warning = self::object($warning, $context);
            self::exactOptional($warning, ['code', 'message'], ['sheet'], $context);
            $code = self::string($warning, 'code', $context);
            if (!in_array($code, ['empty_sheet_skipped', 'formulas_excluded', 'style_facts_excluded', 'hidden_cells_excluded'], true)) {
                self::fail("$context.code is unsupported");
            }
            self::string($warning, 'message', $context);
            self::optionalString($warning, 'sheet', $context);
        }
    }

    /** @param array<string, mixed> $document */
    public static function assertDocument(array $document, string $context = 'retrieval document'): void
    {
        self::exact($document, ['schema_version', 'document_id', 'source', 'location', 'chunk', 'content', 'source_refs'], $context);
        self::version($document, $context);
        self::string($document, 'document_id', $context);
        self::assertRetrievalSource(self::requiredObject($document, 'source', $context), "$context.source");

        $location = self::requiredObject($document, 'location', $context);
        self::exactOptional(
            $location,
            ['sheet', 'range', 'row_start', 'row_end', 'column_start', 'column_end'],
            ['sheet_index'],
            "$context.location"
        );
        self::string($location, 'sheet', "$context.location");
        self::optionalInteger($location, 'sheet_index', "$context.location");
        self::assertRange(self::requiredObject($location, 'range', "$context.location"), "$context.location.range");
        foreach (['row_start', 'row_end', 'column_start', 'column_end'] as $key) {
            self::integer($location, $key, "$context.location");
        }

        $chunk = self::requiredObject($document, 'chunk', $context);
        self::exact($chunk, ['profile', 'profile_version', 'ordinal', 'overlap_from_previous_rows'], "$context.chunk");
        self::string($chunk, 'profile', "$context.chunk");
        self::string($chunk, 'profile_version', "$context.chunk");
        self::integer($chunk, 'ordinal', "$context.chunk");
        self::integer($chunk, 'overlap_from_previous_rows', "$context.chunk");

        $content = self::requiredObject($document, 'content', $context);
        self::exact($content, ['media_type', 'text', 'cell_count', 'formula_count', 'merged_range_count'], "$context.content");
        self::string($content, 'media_type', "$context.content");
        self::string($content, 'text', "$context.content");
        foreach (['cell_count', 'formula_count', 'merged_range_count'] as $key) {
            self::integer($content, $key, "$context.content");
        }

        $refs = self::list($document, 'source_refs', $context);
        foreach ($refs as $index => $sourceRef) {
            self::assertSourceRef(self::object($sourceRef, "$context.source_refs[$index]"), "$context.source_refs[$index]");
        }
    }

    /** @param array<string, mixed> $source */
    public static function assertRetrievalSource(array $source, string $context = 'source'): void
    {
        self::exact($source, ['source_id', 'source_revision', 'raw_schema_version'], $context);
        self::string($source, 'source_id', $context);
        self::string($source, 'source_revision', $context);
        self::string($source, 'raw_schema_version', $context);
    }

    /** @param array<string, mixed> $sourceRef */
    public static function assertSourceRef(array $sourceRef, string $context = 'source_ref'): void
    {
        self::exact($sourceRef, ['sheet', 'range'], $context);
        self::string($sourceRef, 'sheet', $context);
        self::assertRange(self::requiredObject($sourceRef, 'range', $context), "$context.range");
    }

    /** @param array<string, mixed> $range */
    public static function assertRange(array $range, string $context = 'range'): void
    {
        self::exact($range, ['range', 'start', 'end'], $context);
        self::string($range, 'range', $context);
        self::string($range, 'start', $context);
        self::string($range, 'end', $context);
    }

    /** @param array<string, mixed> $manifest */
    private static function assertManifest(array $manifest): void
    {
        $context = 'manifest';
        self::exact($manifest, ['schema_version', 'source_id', 'source_revision', 'profile', 'profile_version', 'document_ids'], $context);
        self::version($manifest, $context);
        foreach (['source_id', 'source_revision', 'profile', 'profile_version'] as $key) {
            self::string($manifest, $key, $context);
        }
        $ids = self::list($manifest, 'document_ids', $context);
        foreach ($ids as $index => $id) {
            if (!is_string($id)) {
                self::fail("$context.document_ids[$index] must be a string");
            }
        }
    }

    /** @return array<string, mixed> */
    private static function readObject(string $path, string $description): array
    {
        $contents = @file_get_contents($path);
        if ($contents === false) {
            throw new ContractException("failed to read $description $path");
        }
        try {
            $value = json_decode($contents, true, 512, JSON_THROW_ON_ERROR);
        } catch (JsonException $exception) {
            throw new ContractException("invalid $description JSON in $path: " . $exception->getMessage(), 0, $exception);
        }

        return self::object($value, $description);
    }

    /** @param array<string, mixed> $value */
    private static function version(array $value, string $context): void
    {
        if (($value['schema_version'] ?? null) !== '1.0') {
            self::fail("$context.schema_version must be \"1.0\"");
        }
    }

    /** @param mixed $value @return array<string, mixed> */
    private static function object($value, string $context): array
    {
        if (!is_array($value) || ($value !== [] && self::isList($value))) {
            self::fail("$context must be an object");
        }

        return $value;
    }

    /** @param array<string, mixed> $value @return array<string, mixed> */
    private static function requiredObject(array $value, string $key, string $context): array
    {
        if (!array_key_exists($key, $value)) {
            self::fail("$context.$key is required");
        }

        return self::object($value[$key], "$context.$key");
    }

    /** @param array<string, mixed> $value @return list<mixed> */
    private static function list(array $value, string $key, string $context): array
    {
        if (!array_key_exists($key, $value) || !is_array($value[$key]) || !self::isList($value[$key])) {
            self::fail("$context.$key must be an array");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function string(array $value, string $key, string $context): string
    {
        if (!array_key_exists($key, $value) || !is_string($value[$key])) {
            self::fail("$context.$key must be a string");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function optionalString(array $value, string $key, string $context): void
    {
        if (array_key_exists($key, $value) && $value[$key] !== null && !is_string($value[$key])) {
            self::fail("$context.$key must be a string or null");
        }
    }

    /** @param array<string, mixed> $value */
    private static function integer(array $value, string $key, string $context): int
    {
        if (!array_key_exists($key, $value) || !is_int($value[$key]) || $value[$key] < 0) {
            self::fail("$context.$key must be a non-negative integer");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function optionalInteger(array $value, string $key, string $context): void
    {
        if (array_key_exists($key, $value) && $value[$key] !== null && (!is_int($value[$key]) || $value[$key] < 0)) {
            self::fail("$context.$key must be a non-negative integer or null");
        }
    }

    /** @param array<string, mixed> $value */
    private static function boolean(array $value, string $key, string $context): bool
    {
        if (!array_key_exists($key, $value) || !is_bool($value[$key])) {
            self::fail("$context.$key must be a boolean");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value @param list<string> $keys */
    private static function exact(array $value, array $keys, string $context): void
    {
        self::exactOptional($value, $keys, [], $context);
    }

    /** @param array<string, mixed> $value @param list<string> $required @param list<string> $optional */
    private static function exactOptional(array $value, array $required, array $optional, string $context): void
    {
        foreach ($required as $key) {
            if (!array_key_exists($key, $value)) {
                self::fail("$context.$key is required");
            }
        }
        $allowed = array_merge($required, $optional);
        foreach (array_keys($value) as $key) {
            if (!in_array($key, $allowed, true)) {
                self::fail("$context contains unknown field $key");
            }
        }
    }

    /** @param array<mixed> $value */
    private static function isList(array $value): bool
    {
        $expected = 0;
        foreach ($value as $key => $_) {
            if ($key !== $expected++) {
                return false;
            }
        }

        return true;
    }

    private static function fail(string $message): void
    {
        throw new ContractException($message);
    }
}
