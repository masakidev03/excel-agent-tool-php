<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Throwable;

final class Editor
{
    private Reader $reader;
    private Writer $writer;

    public function __construct(?Reader $reader = null, ?Writer $writer = null)
    {
        $this->reader = $reader ?? new Reader();
        $this->writer = $writer ?? new Writer();
    }

    /** @param array<string, mixed> $plan */
    public function validateOperations(string $inputPath, array $plan): void
    {
        $input = PathSafety::input($inputPath);
        $spreadsheet = $this->reader->load($input);
        try {
            $this->applyPlan($spreadsheet, $plan);
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    /** @param array<string, mixed> $plan */
    public function validate(string $inputPath, string $outputPath, array $plan): void
    {
        $input = PathSafety::input($inputPath);
        PathSafety::output($outputPath, $input);
        $spreadsheet = $this->reader->load($input);
        try {
            $this->applyPlan($spreadsheet, $plan);
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    /** @param array<string, mixed> $plan */
    public function edit(string $inputPath, string $outputPath, array $plan): string
    {
        $input = PathSafety::input($inputPath);
        $output = PathSafety::output($outputPath, $input);
        $spreadsheet = $this->reader->load($input);
        try {
            $this->applyPlan($spreadsheet, $plan);
            PathSafety::reserve($output);
            try {
                $xlsx = new Xlsx($spreadsheet);
                $xlsx->setPreCalculateFormulas(false);
                if (method_exists($xlsx, 'setIncludeCharts')) {
                    $xlsx->setIncludeCharts(true);
                }
                $xlsx->save($output);
            } catch (Throwable $exception) {
                @unlink($output);
                throw new EngineException(
                    'output_write_failed',
                    sprintf('failed to write edited XLSX output %s: %s', $output, $exception->getMessage()),
                    $output,
                    $exception
                );
            }

            return $output;
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    /** @param array<string, mixed> $plan */
    private function applyPlan(Spreadsheet $spreadsheet, array $plan): void
    {
        $errors = [];
        $sheets = [];
        foreach ($spreadsheet->getWorksheetIterator() as $worksheet) {
            $sheets[$worksheet->getTitle()] = $worksheet;
        }
        $claims = [
            'exact' => [], 'content' => [], 'style' => [], 'column' => [], 'row' => [], 'merges' => [],
        ];

        foreach ($plan['operations'] as $index => $operation) {
            $before = count($errors);
            $exactKey = $this->canonicalJson($operation);
            if (isset($claims['exact'][$exactKey])) {
                $errors[] = $this->collision('duplicate_operation', $index, $claims['exact'][$exactKey], $operation);
                continue;
            }

            $kind = $operation['operation'];
            switch ($kind) {
                case 'set_cell_value':
                    $target = $this->targetCell($operation, $index, $sheets, $errors);
                    if ($target === null) {
                        break;
                    }
                    [$worksheet, $address, $coordinate] = $target;
                    Validator::literal($operation['value'], $operation['value_type'], $operation['sheet'], $address, $index, false, $errors);
                    $this->validateEditable($worksheet, $coordinate, $operation['sheet'], $address, $index, $errors);
                    if (count($errors) === $before && $this->claim(
                        $claims['content'],
                        $operation['sheet'] . '!' . $address,
                        ['value' => $operation['value'], 'value_type' => $operation['value_type']],
                        $index,
                        $operation,
                        $errors
                    )) {
                        $this->writer->writeValue($worksheet->getCell($address), $operation['value'], $operation['value_type']);
                    }
                    break;

                case 'set_cell_formula':
                    $target = $this->targetCell($operation, $index, $sheets, $errors);
                    if ($target === null) {
                        break;
                    }
                    [$worksheet, $address, $coordinate] = $target;
                    Validator::formula($operation['formula'], $operation['sheet'], $address, $index, $errors);
                    $this->validateEditable($worksheet, $coordinate, $operation['sheet'], $address, $index, $errors);
                    if (count($errors) === $before && $this->claim(
                        $claims['content'],
                        $operation['sheet'] . '!' . $address,
                        ['formula' => $operation['formula']],
                        $index,
                        $operation,
                        $errors
                    )) {
                        $worksheet->getCell($address)->setValueExplicit($operation['formula'], DataType::TYPE_FORMULA);
                    }
                    break;

                case 'set_cell_style':
                    $target = $this->targetCell($operation, $index, $sheets, $errors);
                    if ($target === null) {
                        break;
                    }
                    [$worksheet, $address] = $target;
                    Validator::style($operation['style'], $operation['sheet'], $address, $index, $errors);
                    if (count($errors) === $before && $this->claim(
                        $claims['style'],
                        $operation['sheet'] . '!' . $address,
                        $operation['style'],
                        $index,
                        $operation,
                        $errors
                    )) {
                        StyleMapper::apply($worksheet->getCell($address), $operation['style']);
                    }
                    break;

                case 'merge_cells':
                    $worksheet = $this->targetSheet($operation['sheet'], $index, $sheets, $errors);
                    if ($worksheet === null) {
                        break;
                    }
                    $range = Validator::mergedRange($operation['range'], $operation['sheet'], $index, $errors);
                    if ($range === null) {
                        break;
                    }
                    $rangeText = $this->rangeText($range);
                    $tentativeMerges = $claims['merges'];
                    if (!$this->claimMerge($tentativeMerges, $operation['sheet'], $range, 'merge', $index, $operation, $errors)) {
                        break;
                    }
                    foreach ($this->worksheetRanges($worksheet) as $existing) {
                        if (Coordinate::rangesOverlap($existing, $range)) {
                            $errors[] = Validator::error('invalid_range', 'merged range overlaps an existing merged range', $index, $operation['sheet'], null, $rangeText);
                            break 2;
                        }
                    }
                    if ($this->rangeHasContentOutsideAnchor($worksheet, $range)) {
                        $errors[] = Validator::error('invalid_range', 'merged range contains content outside its top-left anchor', $index, $operation['sheet'], null, $rangeText);
                        break;
                    }
                    $claims['merges'] = $tentativeMerges;
                    $worksheet->mergeCells($rangeText);
                    break;

                case 'unmerge_cells':
                    $worksheet = $this->targetSheet($operation['sheet'], $index, $sheets, $errors);
                    if ($worksheet === null) {
                        break;
                    }
                    $range = Validator::mergedRange($operation['range'], $operation['sheet'], $index, $errors);
                    if ($range === null) {
                        break;
                    }
                    $rangeText = $this->rangeText($range);
                    $exists = false;
                    foreach ($this->worksheetRanges($worksheet) as $existing) {
                        if ($existing === $range) {
                            $exists = true;
                            break;
                        }
                    }
                    if (!$exists) {
                        $errors[] = Validator::error('invalid_range', 'unmerge range does not exactly match an existing merged range', $index, $operation['sheet'], null, $rangeText);
                        break;
                    }
                    if ($this->claimMerge($claims['merges'], $operation['sheet'], $range, 'unmerge', $index, $operation, $errors)) {
                        $worksheet->unmergeCells($rangeText);
                    }
                    break;

                case 'set_column_width':
                    $worksheet = $this->targetSheet($operation['sheet'], $index, $sheets, $errors);
                    if ($worksheet === null) {
                        break;
                    }
                    if ($operation['column'] < 1 || $operation['column'] > Coordinate::MAX_COLUMNS) {
                        $errors[] = Validator::error('invalid_range', 'column must be in the range 1..=16384', $index, $operation['sheet']);
                    }
                    if (!$this->finiteRange($operation['width'], 255.0)) {
                        $errors[] = Validator::error('invalid_value', 'column width must be finite and in the range 0 < width <= 255', $index, $operation['sheet']);
                    }
                    if (count($errors) === $before && $this->claim(
                        $claims['column'],
                        $operation['sheet'] . '!' . $operation['column'],
                        $operation['width'],
                        $index,
                        $operation,
                        $errors
                    )) {
                        $columnName = preg_replace('/[0-9]+$/', '', Coordinate::address(1, $operation['column']));
                        $worksheet->getColumnDimension($columnName)->setWidth((float) $operation['width']);
                    }
                    break;

                case 'set_row_height':
                    $worksheet = $this->targetSheet($operation['sheet'], $index, $sheets, $errors);
                    if ($worksheet === null) {
                        break;
                    }
                    if ($operation['row'] < 1 || $operation['row'] > Coordinate::MAX_ROWS) {
                        $errors[] = Validator::error('invalid_range', 'row must be in the range 1..=1048576', $index, $operation['sheet']);
                    }
                    if (!$this->finiteRange($operation['height'], 409.5)) {
                        $errors[] = Validator::error('invalid_value', 'row height must be finite and in the range 0 < height <= 409.5', $index, $operation['sheet']);
                    }
                    if (count($errors) === $before && $this->claim(
                        $claims['row'],
                        $operation['sheet'] . '!' . $operation['row'],
                        $operation['height'],
                        $index,
                        $operation,
                        $errors
                    )) {
                        $worksheet->getRowDimension($operation['row'])->setRowHeight((float) $operation['height']);
                    }
                    break;

                case 'append_rows':
                    $worksheet = $this->targetSheet($operation['sheet'], $index, $sheets, $errors);
                    if ($worksheet !== null) {
                        $this->appendRows($worksheet, $operation, $index, $claims, $errors);
                    }
                    break;

                case 'create_sheet':
                case 'delete_sheet':
                case 'rename_sheet':
                    $errors[] = Validator::error(
                        'unsupported_operation',
                        'Editor MVP does not support worksheet creation, deletion, or rename',
                        $index
                    );
                    break;
            }

            if (count($errors) === $before) {
                $claims['exact'][$exactKey] = $index;
            }
        }

        if ($errors !== []) {
            throw new ValidationException('edit plan failed validation', $errors);
        }
    }

    /** @param array<string, mixed> $operation @param array<string, Worksheet> $sheets
     * @param list<array<string, mixed>> $errors
     * @return array{0: Worksheet, 1: string, 2: array{row: int, column: int}}|null
     */
    private function targetCell(array $operation, int $index, array $sheets, array &$errors): ?array
    {
        $worksheet = $this->targetSheet($operation['sheet'], $index, $sheets, $errors);
        if ($worksheet === null) {
            return null;
        }
        $coordinate = Coordinate::parseCell($operation['address']);
        if ($coordinate === null) {
            $errors[] = Validator::error(
                'invalid_cell_reference',
                sprintf('cell address %s is invalid or outside XLSX limits', json_encode($operation['address'])),
                $index,
                $operation['sheet'],
                $operation['address']
            );
            return null;
        }
        $address = Coordinate::address($coordinate['row'], $coordinate['column']);

        return [$worksheet, $address, $coordinate];
    }

    /** @param array<string, Worksheet> $sheets @param list<array<string, mixed>> $errors */
    private function targetSheet(string $name, int $index, array $sheets, array &$errors): ?Worksheet
    {
        if (!isset($sheets[$name])) {
            $errors[] = Validator::error('worksheet_not_found', sprintf('worksheet %s does not exist', json_encode($name)), $index, $name);
            return null;
        }

        return $sheets[$name];
    }

    /** @param array{row: int, column: int} $coordinate @param list<array<string, mixed>> $errors */
    private function validateEditable(Worksheet $worksheet, array $coordinate, string $sheet, string $address, int $index, array &$errors): void
    {
        foreach ($this->worksheetRanges($worksheet) as $range) {
            if (Coordinate::rangeContains($range, $coordinate) && $coordinate !== $range['start']) {
                $errors[] = Validator::error('invalid_cell_reference', 'cell is a subordinate member of a merged range', $index, $sheet, $address);
                return;
            }
        }
    }

    /** @param array<string, mixed> $operation @param array<string, mixed> $claims @param list<array<string, mixed>> $errors */
    private function appendRows(Worksheet $worksheet, array $operation, int $index, array &$claims, array &$errors): void
    {
        $before = count($errors);
        $maxRow = 0;
        foreach ($worksheet->getCellCollection()->getCoordinates() as $address) {
            $coordinate = Coordinate::parseCell($address);
            if ($coordinate !== null) {
                $maxRow = max($maxRow, $coordinate['row']);
            }
        }
        $startRow = $maxRow + 1;
        if ($startRow + count($operation['rows']) - 1 > Coordinate::MAX_ROWS) {
            $errors[] = Validator::error('invalid_range', 'appended rows exceed the XLSX row limit', $index, $operation['sheet']);
            return;
        }

        $prepared = [];
        foreach ($operation['rows'] as $rowOffset => $row) {
            if (count($row) > Coordinate::MAX_COLUMNS) {
                $errors[] = Validator::error('invalid_range', 'an appended row exceeds the XLSX column limit', $index, $operation['sheet']);
                continue;
            }
            foreach ($row as $columnOffset => $input) {
                $rowNumber = $startRow + $rowOffset;
                $columnNumber = $columnOffset + 1;
                $address = Coordinate::address($rowNumber, $columnNumber);
                Validator::cellInput($input, $operation['sheet'], $address, $index, false, $errors);
                $prepared[] = [$address, ['row' => $rowNumber, 'column' => $columnNumber], $input];
            }
        }
        if (count($errors) !== $before) {
            return;
        }

        $tentativeContent = $claims['content'];
        $tentativeStyle = $claims['style'];
        foreach ($prepared as [$address, $coordinate, $input]) {
            if ($input['value'] !== null || isset($input['formula'])) {
                $this->validateEditable($worksheet, $coordinate, $operation['sheet'], $address, $index, $errors);
                $claim = isset($input['formula'])
                    ? ['formula' => $input['formula']]
                    : ['value' => $input['value'], 'value_type' => $input['value_type']];
                $this->claim($tentativeContent, $operation['sheet'] . '!' . $address, $claim, $index, $operation, $errors);
            }
            if (isset($input['style'])) {
                $this->claim($tentativeStyle, $operation['sheet'] . '!' . $address, $input['style'], $index, $operation, $errors);
            }
        }
        if (count($errors) !== $before) {
            return;
        }
        $claims['content'] = $tentativeContent;
        $claims['style'] = $tentativeStyle;

        foreach ($prepared as [$address, $_coordinate, $input]) {
            $cell = $worksheet->getCell($address);
            if (isset($input['formula'])) {
                $cell->setValueExplicit($input['formula'], DataType::TYPE_FORMULA);
            } elseif ($input['value'] !== null) {
                $this->writer->writeValue($cell, $input['value'], $input['value_type']);
            }
            if (isset($input['style'])) {
                StyleMapper::apply($cell, $input['style']);
            }
        }
    }

    /** @param array<string, array{index: int, claim: mixed}> $domain @param mixed $claim
     * @param array<string, mixed> $operation @param list<array<string, mixed>> $errors
     */
    private function claim(array &$domain, string $target, $claim, int $index, array $operation, array &$errors): bool
    {
        if (isset($domain[$target])) {
            $code = $this->canonicalJson($domain[$target]['claim']) === $this->canonicalJson($claim)
                ? 'duplicate_operation'
                : 'conflicting_operation';
            $errors[] = $this->collision($code, $index, $domain[$target]['index'], $operation);
            return false;
        }
        $domain[$target] = ['index' => $index, 'claim' => $claim];

        return true;
    }

    /** @param list<array<string, mixed>> $mergeClaims
     * @param array{start: array{row: int, column: int}, end: array{row: int, column: int}} $range
     * @param array<string, mixed> $operation @param list<array<string, mixed>> $errors
     */
    private function claimMerge(array &$mergeClaims, string $sheet, array $range, string $claim, int $index, array $operation, array &$errors): bool
    {
        foreach ($mergeClaims as $previous) {
            if ($previous['sheet'] !== $sheet) {
                continue;
            }
            if ($previous['range'] === $range) {
                $code = $previous['claim'] === $claim ? 'duplicate_operation' : 'conflicting_operation';
                $errors[] = $this->collision($code, $index, $previous['index'], $operation);
                return false;
            }
            if ($claim === 'merge' && $previous['claim'] === 'merge' && Coordinate::rangesOverlap($previous['range'], $range)) {
                $errors[] = $this->collision('conflicting_operation', $index, $previous['index'], $operation);
                return false;
            }
        }
        $mergeClaims[] = ['sheet' => $sheet, 'range' => $range, 'claim' => $claim, 'index' => $index];

        return true;
    }

    /** @param array<string, mixed> $operation @return array<string, mixed> */
    private function collision(string $code, int $index, int $previous, array $operation): array
    {
        return Validator::error(
            $code,
            sprintf('%s conflicts with operation %d', $code === 'duplicate_operation' ? 'duplicate operation' : 'operation', $previous),
            $index,
            $operation['sheet'] ?? null,
            $operation['address'] ?? null,
            isset($operation['range']) ? $operation['range']['range'] : null
        );
    }

    /** @return list<array{start: array{row: int, column: int}, end: array{row: int, column: int}}> */
    private function worksheetRanges(Worksheet $worksheet): array
    {
        $ranges = [];
        foreach ($worksheet->getMergeCells() as $rangeText) {
            $range = Coordinate::parseRange(str_replace('$', '', $rangeText));
            if ($range !== null) {
                $ranges[] = $range;
            }
        }

        return $ranges;
    }

    /** @param array{start: array{row: int, column: int}, end: array{row: int, column: int}} $range */
    private function rangeText(array $range): string
    {
        return Coordinate::address($range['start']['row'], $range['start']['column'])
            . ':' . Coordinate::address($range['end']['row'], $range['end']['column']);
    }

    /** @param array{start: array{row: int, column: int}, end: array{row: int, column: int}} $range */
    private function rangeHasContentOutsideAnchor(Worksheet $worksheet, array $range): bool
    {
        foreach ($worksheet->getCellCollection()->getCoordinates() as $address) {
            $coordinate = Coordinate::parseCell($address);
            if ($coordinate !== null && $coordinate !== $range['start'] && Coordinate::rangeContains($range, $coordinate)) {
                if ($worksheet->getCell($address)->getValue() !== null) {
                    return true;
                }
            }
        }

        return false;
    }

    /** @param mixed $value */
    private function canonicalJson($value): string
    {
        return (string) json_encode($this->canonical($value), JSON_PRESERVE_ZERO_FRACTION | JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    }

    /** @param mixed $value @return mixed */
    private function canonical($value)
    {
        if (!is_array($value)) {
            return $value;
        }
        $isList = true;
        $expected = 0;
        foreach ($value as $key => $_) {
            if ($key !== $expected++) {
                $isList = false;
                break;
            }
        }
        if (!$isList) {
            ksort($value);
        }
        foreach ($value as $key => $item) {
            $value[$key] = $this->canonical($item);
        }

        return $value;
    }

    /** @param mixed $value */
    private function finiteRange($value, float $maximum): bool
    {
        return (is_int($value) || is_float($value)) && is_finite((float) $value) && $value > 0 && $value <= $maximum;
    }
}
