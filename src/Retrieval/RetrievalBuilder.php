<?php

declare(strict_types=1);

namespace ExcelAgentTool\Retrieval;

use ExcelAgentTool\Engine\Coordinate;
use JsonException;

final class RetrievalBuilder
{
    /**
     * @param array<string, mixed> $workbook
     * @param array<string, mixed> $source
     * @param array<string, mixed>|null $profile
     * @return array<string, mixed>
     */
    public function build(array $workbook, array $source, ?array $profile = null): array
    {
        $profile = $profile ?? RetrievalModel::defaultProfile();
        $this->validateSource($source);
        RetrievalModel::assertProfile($profile);
        $this->validateProfile($profile);
        $this->validateWorkbook($workbook);

        $warnings = [];
        $candidates = [];
        foreach ($this->orderedSheets($workbook['sheets']) as $worksheet) {
            $bounds = $this->worksheetBounds($worksheet);
            if ($bounds === null) {
                if ($profile['include_empty_sheets']) {
                    $candidates[] = ['worksheet' => $worksheet, 'bounds' => $this->single(1, 1), 'overlap_rows' => 0];
                } else {
                    $warnings[] = $this->warning('empty_sheet_skipped', 'empty worksheet skipped by retrieval profile', $worksheet['name']);
                }
                continue;
            }
            $this->appendConfigurationWarnings($worksheet, $profile, $warnings);
            $this->appendWindows($worksheet, $bounds, $profile, $candidates);
        }

        $splitCandidates = [];
        foreach ($candidates as $candidate) {
            $this->splitForLimits($candidate, $profile, $splitCandidates);
        }

        $sourceValue = [
            'source_id' => $source['source_id'],
            'source_revision' => $source['source_revision'],
            'raw_schema_version' => $workbook['schema_version'],
        ];
        $documents = [];
        $identities = [];
        foreach ($splitCandidates as $ordinal => $candidate) {
            $rendered = $this->render($candidate['worksheet'], $candidate['bounds'], $profile);
            $identity = $this->canonicalIdentity($source, $profile, $candidate, $ordinal);
            $documentId = 'rdoc_' . hash('sha256', $identity);
            if (isset($identities[$documentId]) && $identities[$documentId] !== $identity) {
                throw $this->failure(
                    'document_id_collision',
                    'different canonical chunks produced the same document ID',
                    $candidate['worksheet']['name'],
                    $this->a1Range($candidate['bounds'])['range']
                );
            }
            $identities[$documentId] = $identity;
            $range = $this->a1Range($candidate['bounds']);
            $documents[] = [
                'schema_version' => '1.0',
                'document_id' => $documentId,
                'source' => $sourceValue,
                'location' => array_filter([
                    'sheet' => $candidate['worksheet']['name'],
                    'sheet_index' => $candidate['worksheet']['sheet_index'] ?? null,
                    'range' => $range,
                    'row_start' => $candidate['bounds']['row_start'],
                    'row_end' => $candidate['bounds']['row_end'],
                    'column_start' => $candidate['bounds']['column_start'],
                    'column_end' => $candidate['bounds']['column_end'],
                ], static fn ($value): bool => $value !== null),
                'chunk' => [
                    'profile' => $profile['profile'],
                    'profile_version' => $profile['profile_version'],
                    'ordinal' => $ordinal,
                    'overlap_from_previous_rows' => $candidate['overlap_rows'],
                ],
                'content' => $rendered,
                'source_refs' => $this->sourceRefs($candidate['worksheet'], $candidate['bounds']),
            ];
        }

        $documentIds = array_map(static fn (array $document): string => $document['document_id'], $documents);

        return [
            'schema_version' => '1.0',
            'documents' => $documents,
            'manifest' => [
                'schema_version' => '1.0',
                'source_id' => $source['source_id'],
                'source_revision' => $source['source_revision'],
                'profile' => $profile['profile'],
                'profile_version' => $profile['profile_version'],
                'document_ids' => $documentIds,
            ],
            'warnings' => $warnings,
        ];
    }

    /**
     * @param array<string, mixed> $workbook
     * @param array<string, mixed> $source
     * @param array<string, mixed>|null $profile
     * @return list<array<string, mixed>>
     */
    public function documents(array $workbook, array $source, ?array $profile = null): array
    {
        return $this->build($workbook, $source, $profile)['documents'];
    }

    /** @param array<string, mixed> $source */
    private function validateSource(array $source): void
    {
        $allowed = ['schema_version', 'source_id', 'source_revision', 'display_name'];
        foreach (array_keys($source) as $key) {
            if (!in_array($key, $allowed, true)) {
                throw $this->failure('invalid_source', "source contains unknown field $key");
            }
        }
        if (($source['schema_version'] ?? null) !== '1.0') {
            throw $this->failure('invalid_source', 'source schema_version must be "1.0"');
        }
        foreach (['source_id', 'source_revision'] as $key) {
            $value = $source[$key] ?? null;
            if (!is_string($value) || $this->blank($value) || $this->length($value) > 512 || preg_match('/[\x00-\x1F\x7F]/', $value)) {
                throw $this->failure(
                    'invalid_source',
                    "$key must be non-empty, at most 512 characters, and contain no control characters"
                );
            }
        }
        if (array_key_exists('display_name', $source) && $source['display_name'] !== null && !is_string($source['display_name'])) {
            throw $this->failure('invalid_source', 'display_name must be a string or null');
        }
    }

    /** @param array<string, mixed> $profile */
    private function validateProfile(array $profile): void
    {
        if ($profile['profile'] !== 'worksheet_rows_v1' || $profile['profile_version'] !== '1') {
            throw $this->failure('invalid_profile', 'only profile worksheet_rows_v1 version 1 is supported');
        }
        if (
            $profile['target_rows'] === 0
            || $profile['overlap_rows'] >= $profile['target_rows']
            || $profile['max_columns'] === 0
            || $profile['max_columns'] > Coordinate::MAX_COLUMNS
            || $profile['max_addressed_cells'] === 0
            || $profile['max_text_chars'] === 0
        ) {
            throw $this->failure(
                'invalid_profile',
                'profile limits must be positive, overlap_rows must be less than target_rows, and max_columns must fit XLSX limits'
            );
        }
    }

    /** @param array<string, mixed> $workbook */
    private function validateWorkbook(array $workbook): void
    {
        if (($workbook['schema_version'] ?? null) !== '1.0' || !isset($workbook['sheets']) || !is_array($workbook['sheets'])) {
            throw $this->failure('invalid_raw_workbook', 'RawWorkbook must use schema version 1.0 and contain sheets');
        }
        $names = [];
        $indexes = [];
        foreach ($workbook['sheets'] as $worksheet) {
            $sheet = is_array($worksheet) && is_string($worksheet['name'] ?? null) ? $worksheet['name'] : null;
            if ($sheet === null || $sheet === '' || isset($names[$sheet])) {
                throw $this->rawFailure($worksheet, 'worksheet names must be non-empty and unique');
            }
            $names[$sheet] = true;
            if (isset($worksheet['sheet_index'])) {
                if (!is_int($worksheet['sheet_index']) || $worksheet['sheet_index'] < 0 || isset($indexes[$worksheet['sheet_index']])) {
                    throw $this->rawFailure($worksheet, 'worksheet indexes must be unique when present');
                }
                $indexes[$worksheet['sheet_index']] = true;
            }

            $coordinates = [];
            foreach ($worksheet['cells'] ?? [] as $cell) {
                $parsed = is_array($cell) && is_string($cell['address'] ?? null) ? $this->parseCell($cell['address']) : null;
                $row = $cell['row'] ?? null;
                $column = $cell['column'] ?? null;
                $key = is_int($row) && is_int($column) ? "$row:$column" : '';
                $value = $cell['value'] ?? null;
                $valueType = $cell['value_type'] ?? null;
                if (
                    $parsed === null
                    || $parsed !== ['row' => $row, 'column' => $column]
                    || isset($coordinates[$key])
                    || !$this->compatible($valueType, $value)
                    || ((is_float($value) || is_int($value)) && !is_finite((float) $value))
                ) {
                    throw $this->rawFailure(
                        $worksheet,
                        'cell address, coordinates, value type, and finite value must be consistent and unique',
                        is_array($cell) && is_string($cell['address'] ?? null) ? $cell['address'] : null
                    );
                }
                $coordinates[$key] = true;
            }

            $merges = [];
            foreach ($worksheet['merged_ranges'] ?? [] as $merged) {
                $rangeText = is_array($merged) ? ($merged['range'] ?? null) : null;
                $bounds = is_string($rangeText) ? $this->parseRange($rangeText) : null;
                if (
                    $bounds === null
                    || $this->parseCell($merged['start'] ?? '') !== ['row' => $bounds['row_start'], 'column' => $bounds['column_start']]
                    || $this->parseCell($merged['end'] ?? '') !== ['row' => $bounds['row_end'], 'column' => $bounds['column_end']]
                    || isset($merges[$rangeText])
                ) {
                    throw $this->rawFailure($worksheet, 'merged range, endpoints, and uniqueness must be consistent', is_string($rangeText) ? $rangeText : null);
                }
                $merges[$rangeText] = true;
            }

            $rows = [];
            foreach ($worksheet['row_properties'] ?? [] as $properties) {
                $row = $properties['row'] ?? null;
                $height = $properties['height'] ?? null;
                if (!is_int($row) || $row < 1 || $row > Coordinate::MAX_ROWS || isset($rows[$row]) || ($height !== null && (!is_numeric($height) || !is_finite((float) $height) || $height <= 0))) {
                    throw $this->rawFailure($worksheet, 'row properties must use unique valid rows and finite positive heights', is_scalar($row) ? (string) $row : null);
                }
                $rows[$row] = true;
            }
            $columns = [];
            foreach ($worksheet['column_properties'] ?? [] as $properties) {
                $column = $properties['column'] ?? null;
                $width = $properties['width'] ?? null;
                if (!is_int($column) || $column < 1 || $column > Coordinate::MAX_COLUMNS || isset($columns[$column]) || ($width !== null && (!is_numeric($width) || !is_finite((float) $width) || $width <= 0))) {
                    throw $this->rawFailure($worksheet, 'column properties must use unique valid columns and finite positive widths', is_scalar($column) ? (string) $column : null);
                }
                $columns[$column] = true;
            }
        }
    }

    /** @param list<array<string, mixed>> $sheets @return list<array<string, mixed>> */
    private function orderedSheets(array $sheets): array
    {
        usort($sheets, static function (array $left, array $right): int {
            $leftIndex = $left['sheet_index'] ?? PHP_INT_MAX;
            $rightIndex = $right['sheet_index'] ?? PHP_INT_MAX;
            $byIndex = $leftIndex <=> $rightIndex;

            return $byIndex !== 0 ? $byIndex : strcmp($left['name'], $right['name']);
        });

        return $sheets;
    }

    /** @param array<string, mixed> $worksheet @return array<string, int>|null */
    private function worksheetBounds(array $worksheet): ?array
    {
        $bounds = null;
        foreach ($worksheet['cells'] as $cell) {
            $this->extendBounds($bounds, $this->single($cell['row'], $cell['column']));
        }
        foreach ($worksheet['merged_ranges'] as $merged) {
            $range = $this->parseRange($merged['range']);
            if ($range === null) {
                throw $this->rawFailure($worksheet, 'merged range could not be parsed while computing bounds', $merged['range']);
            }
            $this->extendBounds($bounds, $range);
        }
        foreach ($worksheet['row_properties'] as $row) {
            $this->extendBounds($bounds, $this->single($row['row'], 1));
        }
        foreach ($worksheet['column_properties'] as $column) {
            $this->extendBounds($bounds, $this->single(1, $column['column']));
        }

        return $bounds;
    }

    /** @param array<string, int>|null $current @param array<string, int> $additional */
    private function extendBounds(?array &$current, array $additional): void
    {
        if ($current === null) {
            $current = $additional;
            return;
        }
        $current = [
            'row_start' => min($current['row_start'], $additional['row_start']),
            'row_end' => max($current['row_end'], $additional['row_end']),
            'column_start' => min($current['column_start'], $additional['column_start']),
            'column_end' => max($current['column_end'], $additional['column_end']),
        ];
    }

    /** @param array<string, mixed> $worksheet @param array<string, int> $bounds @param array<string, mixed> $profile @param list<array<string, mixed>> $candidates */
    private function appendWindows(array $worksheet, array $bounds, array $profile, array &$candidates): void
    {
        $rowStart = $bounds['row_start'];
        $first = true;
        while (true) {
            $rowEnd = min($bounds['row_end'], $rowStart + $profile['target_rows'] - 1);
            $columnStart = $bounds['column_start'];
            while ($columnStart <= $bounds['column_end']) {
                $columnEnd = min($bounds['column_end'], $columnStart + $profile['max_columns'] - 1);
                $candidates[] = [
                    'worksheet' => $worksheet,
                    'bounds' => [
                        'row_start' => $rowStart,
                        'row_end' => $rowEnd,
                        'column_start' => $columnStart,
                        'column_end' => $columnEnd,
                    ],
                    'overlap_rows' => $first ? 0 : $profile['overlap_rows'],
                ];
                if ($columnEnd === $bounds['column_end']) {
                    break;
                }
                $columnStart = $columnEnd + 1;
            }
            if ($rowEnd === $bounds['row_end']) {
                break;
            }
            $rowStart = $rowEnd + 1 - $profile['overlap_rows'];
            $first = false;
        }
    }

    /** @param array<string, mixed> $candidate @param array<string, mixed> $profile @param list<array<string, mixed>> $output */
    private function splitForLimits(array $candidate, array $profile, array &$output): void
    {
        if ($this->area($candidate['bounds']) > $profile['max_addressed_cells']) {
            [$left, $right] = $this->splitCandidate($candidate);
            $this->splitForLimits($left, $profile, $output);
            $this->splitForLimits($right, $profile, $output);
            return;
        }
        $content = $this->render($candidate['worksheet'], $candidate['bounds'], $profile);
        if ($this->length($content['text']) > $profile['max_text_chars']) {
            if ($this->height($candidate['bounds']) === 1 && $this->width($candidate['bounds']) === 1) {
                throw $this->failure(
                    'text_limit_exceeded',
                    'one worksheet cell cannot fit max_text_chars without truncation',
                    $candidate['worksheet']['name'],
                    $this->a1Range($candidate['bounds'])['range']
                );
            }
            [$left, $right] = $this->splitCandidate($candidate);
            $this->splitForLimits($left, $profile, $output);
            $this->splitForLimits($right, $profile, $output);
            return;
        }
        $output[] = $candidate;
    }

    /** @param array<string, mixed> $candidate @return array{0: array<string, mixed>, 1: array<string, mixed>} */
    private function splitCandidate(array $candidate): array
    {
        $left = $candidate;
        $right = $candidate;
        $right['overlap_rows'] = 0;
        if ($this->height($candidate['bounds']) > 1) {
            $middle = $candidate['bounds']['row_start'] + intdiv($this->height($candidate['bounds']), 2) - 1;
            $left['bounds']['row_end'] = $middle;
            $right['bounds']['row_start'] = $middle + 1;
        } else {
            $middle = $candidate['bounds']['column_start'] + intdiv($this->width($candidate['bounds']), 2) - 1;
            $left['bounds']['column_end'] = $middle;
            $right['bounds']['column_start'] = $middle + 1;
        }

        return [$left, $right];
    }

    /** @param array<string, mixed> $worksheet @param array<string, int> $bounds @param array<string, mixed> $profile @return array<string, mixed> */
    private function render(array $worksheet, array $bounds, array $profile): array
    {
        $range = $this->a1Range($bounds);
        $text = 'sheet' . "\t" . $this->json($worksheet['name']) . "\nrange\t" . $range['range'] . "\n";
        $hiddenRows = [];
        foreach ($worksheet['row_properties'] as $properties) {
            if ($properties['hidden']) {
                $hiddenRows[$properties['row']] = true;
            }
        }
        $hiddenColumns = [];
        foreach ($worksheet['column_properties'] as $properties) {
            if ($properties['hidden']) {
                $hiddenColumns[$properties['column']] = true;
            }
        }

        $cells = array_values(array_filter($worksheet['cells'], function (array $cell) use ($bounds, $profile, $hiddenRows, $hiddenColumns): bool {
            if (!$this->containsCell($bounds, $cell)) {
                return false;
            }
            return $profile['include_hidden_cells'] || (!isset($hiddenRows[$cell['row']]) && !isset($hiddenColumns[$cell['column']]));
        }));
        usort($cells, static fn (array $left, array $right): int => [$left['row'], $left['column']] <=> [$right['row'], $right['column']]);
        $formulaCount = 0;
        foreach ($cells as $cell) {
            $text .= 'cell' . "\t" . $cell['address'] . "\ttype=" . $cell['value_type'] . "\tvalue=" . $this->json($cell['value']);
            if ($profile['include_formulas'] && isset($cell['formula'])) {
                ++$formulaCount;
                $text .= "\tformula=" . $this->json($cell['formula']);
            }
            if ($profile['include_hidden_cells'] && (isset($hiddenRows[$cell['row']]) || isset($hiddenColumns[$cell['column']]))) {
                $text .= "\thidden=true";
            }
            if ($profile['include_basic_style_facts'] && isset($cell['style'])) {
                $text .= $this->renderStyle($cell['style']);
            }
            $text .= "\n";
        }

        $merged = [];
        foreach ($worksheet['merged_ranges'] as $item) {
            $mergedBounds = $this->parseRange($item['range']);
            if ($mergedBounds !== null && $this->intersects($bounds, $mergedBounds)) {
                $merged[] = ['item' => $item, 'bounds' => $mergedBounds];
            }
        }
        usort($merged, static fn (array $left, array $right): int => [
            $left['bounds']['row_start'], $left['bounds']['column_start'], $left['bounds']['row_end'], $left['bounds']['column_end'],
        ] <=> [
            $right['bounds']['row_start'], $right['bounds']['column_start'], $right['bounds']['row_end'], $right['bounds']['column_end'],
        ]);
        foreach ($merged as $entry) {
            $text .= "merged_range\t" . $entry['item']['range'] . "\n";
        }
        foreach ($worksheet['row_properties'] as $properties) {
            if ($properties['row'] >= $bounds['row_start'] && $properties['row'] <= $bounds['row_end']) {
                $text .= "row\t" . $properties['row'] . "\theight=" . $this->json($properties['height'] ?? null)
                    . "\thidden=" . ($properties['hidden'] ? 'true' : 'false') . "\n";
            }
        }
        foreach ($worksheet['column_properties'] as $properties) {
            if ($properties['column'] >= $bounds['column_start'] && $properties['column'] <= $bounds['column_end']) {
                $text .= "column\t" . $properties['column'] . "\twidth=" . $this->json($properties['width'] ?? null)
                    . "\thidden=" . ($properties['hidden'] ? 'true' : 'false') . "\n";
            }
        }

        return [
            'media_type' => 'text/plain',
            'text' => $text,
            'cell_count' => count($cells),
            'formula_count' => $formulaCount,
            'merged_range_count' => count($merged),
        ];
    }

    /** @param array<string, mixed> $style */
    private function renderStyle(array $style): string
    {
        $text = '';
        if (isset($style['number_format'])) {
            $text .= "\tstyle.number_format=" . $this->json($style['number_format']);
        }
        if (isset($style['font'])) {
            foreach (['bold', 'italic'] as $key) {
                if (array_key_exists($key, $style['font'])) {
                    $text .= "\tstyle.font.$key=" . ($style['font'][$key] ? 'true' : 'false');
                }
            }
            if (isset($style['font']['color'])) {
                $text .= "\tstyle.font.color=" . $this->json($style['font']['color']);
            }
        }
        if (isset($style['fill'])) {
            if (isset($style['fill']['pattern'])) {
                $text .= "\tstyle.fill.pattern=" . $this->json($style['fill']['pattern']);
            }
            if (isset($style['fill']['foreground_color'])) {
                $text .= "\tstyle.fill.foreground_color=" . $this->json($style['fill']['foreground_color']);
            }
        }

        return $text;
    }

    /** @param array<string, mixed> $worksheet @param array<string, int> $bounds @return list<array<string, mixed>> */
    private function sourceRefs(array $worksheet, array $bounds): array
    {
        $refs = [['sheet' => $worksheet['name'], 'range' => $this->a1Range($bounds)]];
        foreach ($worksheet['merged_ranges'] as $merged) {
            $mergedBounds = $this->parseRange($merged['range']);
            if ($mergedBounds === null) {
                throw $this->rawFailure($worksheet, 'merged range could not be parsed for source references', $merged['range']);
            }
            if ($this->intersects($bounds, $mergedBounds) && $mergedBounds !== $bounds) {
                $refs[] = [
                    'sheet' => $worksheet['name'],
                    'range' => ['range' => $merged['range'], 'start' => $merged['start'], 'end' => $merged['end']],
                ];
            }
        }
        foreach ($refs as $ref) {
            if ($ref['sheet'] !== $worksheet['name'] || $this->parseRange($ref['range']['range']) === null) {
                throw $this->failure('invalid_source_reference', 'generated source reference could not be validated', $worksheet['name'], $ref['range']['range']);
            }
        }

        return $refs;
    }

    /** @param array<string, mixed> $worksheet @param array<string, mixed> $profile @param list<array<string, mixed>> $warnings */
    private function appendConfigurationWarnings(array $worksheet, array $profile, array &$warnings): void
    {
        if (!$profile['include_formulas'] && array_filter($worksheet['cells'], static fn (array $cell): bool => isset($cell['formula']))) {
            $warnings[] = $this->warning('formulas_excluded', 'formula text excluded by retrieval profile', $worksheet['name']);
        }
        if (!$profile['include_basic_style_facts'] && array_filter($worksheet['cells'], static fn (array $cell): bool => isset($cell['style']))) {
            $warnings[] = $this->warning('style_facts_excluded', 'basic style facts excluded by retrieval profile', $worksheet['name']);
        }
        if (!$profile['include_hidden_cells']) {
            $hiddenRows = [];
            foreach ($worksheet['row_properties'] as $properties) {
                if ($properties['hidden']) {
                    $hiddenRows[$properties['row']] = true;
                }
            }
            $hiddenColumns = [];
            foreach ($worksheet['column_properties'] as $properties) {
                if ($properties['hidden']) {
                    $hiddenColumns[$properties['column']] = true;
                }
            }
            foreach ($worksheet['cells'] as $cell) {
                if (isset($hiddenRows[$cell['row']]) || isset($hiddenColumns[$cell['column']])) {
                    $warnings[] = $this->warning('hidden_cells_excluded', 'cells in hidden rows or columns excluded by retrieval profile', $worksheet['name']);
                    break;
                }
            }
        }
    }

    /** @param array<string, mixed> $source @param array<string, mixed> $profile @param array<string, mixed> $candidate */
    private function canonicalIdentity(array $source, array $profile, array $candidate, int $ordinal): string
    {
        $components = [
            '1.0', $source['source_id'], $source['source_revision'], $profile['profile'], $profile['profile_version'],
            (string) $profile['target_rows'], (string) $profile['overlap_rows'], (string) $profile['max_columns'],
            (string) $profile['max_addressed_cells'], (string) $profile['max_text_chars'],
            $profile['include_formulas'] ? 'true' : 'false',
            $profile['include_basic_style_facts'] ? 'true' : 'false',
            $profile['include_hidden_cells'] ? 'true' : 'false',
            $profile['include_empty_sheets'] ? 'true' : 'false',
            isset($candidate['worksheet']['sheet_index']) ? (string) $candidate['worksheet']['sheet_index'] : '',
            $candidate['worksheet']['name'],
            (string) $candidate['bounds']['row_start'], (string) $candidate['bounds']['row_end'],
            (string) $candidate['bounds']['column_start'], (string) $candidate['bounds']['column_end'],
            (string) $ordinal,
        ];
        $output = '';
        foreach ($components as $component) {
            $output .= pack('J', strlen($component)) . $component;
        }

        return $output;
    }

    /** @param mixed $value */
    private function json($value): string
    {
        try {
            return json_encode(
                $value,
                JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_LINE_TERMINATORS
                | JSON_PRESERVE_ZERO_FRACTION | JSON_THROW_ON_ERROR
            );
        } catch (JsonException $exception) {
            throw $this->failure('invalid_raw_workbook', 'value could not be rendered as JSON: ' . $exception->getMessage());
        }
    }

    /** @return array<string, int> */
    private function single(int $row, int $column): array
    {
        return ['row_start' => $row, 'row_end' => $row, 'column_start' => $column, 'column_end' => $column];
    }

    /** @param array<string, int> $bounds */
    private function height(array $bounds): int
    {
        return $bounds['row_end'] - $bounds['row_start'] + 1;
    }

    /** @param array<string, int> $bounds */
    private function width(array $bounds): int
    {
        return $bounds['column_end'] - $bounds['column_start'] + 1;
    }

    /** @param array<string, int> $bounds */
    private function area(array $bounds): int
    {
        return $this->height($bounds) * $this->width($bounds);
    }

    /** @param array<string, int> $bounds @param array<string, mixed> $cell */
    private function containsCell(array $bounds, array $cell): bool
    {
        return $cell['row'] >= $bounds['row_start'] && $cell['row'] <= $bounds['row_end']
            && $cell['column'] >= $bounds['column_start'] && $cell['column'] <= $bounds['column_end'];
    }

    /** @param array<string, int> $left @param array<string, int> $right */
    private function intersects(array $left, array $right): bool
    {
        return $left['row_start'] <= $right['row_end'] && $left['row_end'] >= $right['row_start']
            && $left['column_start'] <= $right['column_end'] && $left['column_end'] >= $right['column_start'];
    }

    /** @param array<string, int> $bounds @return array<string, string> */
    private function a1Range(array $bounds): array
    {
        $start = Coordinate::address($bounds['row_start'], $bounds['column_start']);
        $end = Coordinate::address($bounds['row_end'], $bounds['column_end']);

        return ['range' => "$start:$end", 'start' => $start, 'end' => $end];
    }

    /** @return array{row: int, column: int}|null */
    private function parseCell(string $address): ?array
    {
        if (!preg_match('/^([A-Z]+)([0-9]+)$/D', $address, $matches)) {
            return null;
        }
        $column = 0;
        foreach (str_split($matches[1]) as $letter) {
            $column = ($column * 26) + ord($letter) - ord('A') + 1;
        }
        $row = (int) $matches[2];
        if ($row < 1 || $row > Coordinate::MAX_ROWS || $column < 1 || $column > Coordinate::MAX_COLUMNS) {
            return null;
        }

        return ['row' => $row, 'column' => $column];
    }

    /** @return array<string, int>|null */
    private function parseRange(string $value): ?array
    {
        if (substr_count($value, ':') !== 1) {
            return null;
        }
        [$startText, $endText] = explode(':', $value, 2);
        $start = $this->parseCell($startText);
        $end = $this->parseCell($endText);
        if ($start === null || $end === null || $start['row'] > $end['row'] || $start['column'] > $end['column']) {
            return null;
        }

        return [
            'row_start' => $start['row'], 'row_end' => $end['row'],
            'column_start' => $start['column'], 'column_end' => $end['column'],
        ];
    }

    /** @param mixed $value */
    private function compatible($valueType, $value): bool
    {
        switch ($valueType) {
            case 'blank':
                return $value === null;
            case 'string':
            case 'error':
                return is_string($value);
            case 'number':
                return is_int($value) || is_float($value);
            case 'boolean':
                return is_bool($value);
            default:
                return false;
        }
    }

    /** @return array<string, mixed> */
    private function warning(string $code, string $message, string $sheet): array
    {
        return ['code' => $code, 'message' => $message, 'sheet' => $sheet];
    }

    /** @param array<string, mixed> $worksheet */
    private function rawFailure(array $worksheet, string $message, ?string $address = null): RetrievalException
    {
        return $this->failure('invalid_raw_workbook', $message, is_string($worksheet['name'] ?? null) ? $worksheet['name'] : null, $address);
    }

    private function failure(string $code, string $message, ?string $sheet = null, ?string $address = null): RetrievalException
    {
        $issue = ['schema_version' => '1.0', 'code' => $code, 'message' => $message];
        if ($sheet !== null) {
            $issue['sheet'] = $sheet;
        }
        if ($address !== null) {
            $issue['address_or_range'] = $address;
        }

        return new RetrievalException($issue);
    }

    private function length(string $value): int
    {
        return function_exists('mb_strlen') ? mb_strlen($value, 'UTF-8') : strlen($value);
    }

    private function blank(string $value): bool
    {
        return preg_match('/^[\s\p{Z}]*$/u', $value) === 1;
    }
}
