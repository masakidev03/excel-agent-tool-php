<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Throwable;

final class Reader
{
    /** @return array<string, mixed> */
    public function read(string $path): array
    {
        $resolved = PathSafety::input($path);
        $spreadsheet = $this->load($resolved);
        try {
            $sheets = [];
            foreach ($spreadsheet->getWorksheetIterator() as $index => $worksheet) {
                $sheets[] = $this->readWorksheet($worksheet, $index);
            }

            return [
                'schema_version' => '1.0',
                'metadata' => ['source_format' => 'xlsx', 'worksheet_count' => count($sheets)],
                'sheets' => $sheets,
            ];
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    public function load(string $resolvedPath): Spreadsheet
    {
        try {
            $reader = new Xlsx();
            $reader->setReadDataOnly(false);
            if (method_exists($reader, 'setIncludeCharts')) {
                $reader->setIncludeCharts(true);
            }

            return $reader->load($resolvedPath);
        } catch (Throwable $exception) {
            throw new EngineException(
                'invalid_workbook',
                sprintf('invalid XLSX workbook %s: %s', $resolvedPath, $exception->getMessage()),
                $resolvedPath,
                $exception
            );
        }
    }

    /** @return array<string, mixed> */
    private function readWorksheet(Worksheet $worksheet, int $index): array
    {
        $comments = method_exists($worksheet, 'getComments') ? $worksheet->getComments() : [];
        $hyperlinks = method_exists($worksheet, 'getHyperlinkCollection') ? $worksheet->getHyperlinkCollection() : [];
        $coordinates = [];
        foreach ($worksheet->getCellCollection()->getCoordinates() as $address) {
            $coordinates[strtoupper($address)] = true;
        }
        foreach (array_keys($comments) as $address) {
            $coordinates[strtoupper((string) $address)] = true;
        }
        foreach (array_keys($hyperlinks) as $address) {
            $coordinates[strtoupper((string) $address)] = true;
        }

        $ordered = [];
        foreach (array_keys($coordinates) as $address) {
            $coordinate = Coordinate::parseCell($address);
            if ($coordinate !== null) {
                $ordered[] = ['address' => Coordinate::address($coordinate['row'], $coordinate['column'])] + $coordinate;
            }
        }
        usort($ordered, static function (array $left, array $right): int {
            return [$left['row'], $left['column']] <=> [$right['row'], $right['column']];
        });

        $cells = [];
        $maxRow = 0;
        $maxColumn = 0;
        foreach ($ordered as $coordinate) {
            $cell = $worksheet->getCell($coordinate['address']);
            $style = StyleMapper::observe($cell);
            $comment = $comments[$coordinate['address']] ?? null;
            $hyperlink = $hyperlinks[$coordinate['address']] ?? null;
            if ($cell->getValue() === null && $style === null && $comment === null && $hyperlink === null) {
                continue;
            }
            $raw = $this->readCell($cell, $coordinate['row'], $coordinate['column']);
            if ($style !== null) {
                $raw['style'] = $style;
            }
            if ($comment !== null) {
                $raw['comment'] = [
                    'author' => $comment->getAuthor(),
                    'text' => $comment->getText()->getPlainText(),
                ];
                if ($raw['comment']['author'] === '') {
                    unset($raw['comment']['author']);
                }
            }
            if ($hyperlink !== null && $hyperlink->getUrl() !== '') {
                $raw['hyperlink'] = ['target' => $hyperlink->getUrl()];
                if ($hyperlink->getTooltip() !== '') {
                    $raw['hyperlink']['tooltip'] = $hyperlink->getTooltip();
                }
            }
            $cells[] = $raw;
            $maxRow = max($maxRow, $coordinate['row']);
            $maxColumn = max($maxColumn, $coordinate['column']);
        }

        $mergedRanges = [];
        foreach ($worksheet->getMergeCells() as $rangeText) {
            $range = Coordinate::parseRange(str_replace('$', '', $rangeText));
            if ($range === null) {
                continue;
            }
            $start = Coordinate::address($range['start']['row'], $range['start']['column']);
            $end = Coordinate::address($range['end']['row'], $range['end']['column']);
            $mergedRanges[] = ['range' => "$start:$end", 'start' => $start, 'end' => $end];
            $maxRow = max($maxRow, $range['end']['row']);
            $maxColumn = max($maxColumn, $range['end']['column']);
        }

        $rowProperties = [];
        foreach ($worksheet->getRowDimensions() as $row => $dimension) {
            $height = $dimension->getRowHeight();
            if ($height < 0 && !$dimension->getVisible()) {
                $rowProperties[] = ['row' => (int) $row, 'hidden' => true];
            } elseif ($height >= 0 || !$dimension->getVisible()) {
                $property = ['row' => (int) $row, 'hidden' => !$dimension->getVisible()];
                if ($height >= 0) {
                    $property['height'] = (float) $height;
                }
                $rowProperties[] = $property;
            }
        }
        usort($rowProperties, static fn (array $left, array $right): int => $left['row'] <=> $right['row']);

        $columnProperties = [];
        foreach ($worksheet->getColumnDimensions() as $columnName => $dimension) {
            $coordinate = Coordinate::parseCell($columnName . '1');
            if ($coordinate === null) {
                continue;
            }
            $width = $dimension->getWidth();
            if ($width >= 0 || !$dimension->getVisible()) {
                $property = ['column' => $coordinate['column'], 'hidden' => !$dimension->getVisible()];
                if ($width >= 0) {
                    $property['width'] = (float) $width;
                }
                $columnProperties[] = $property;
            }
        }
        usort($columnProperties, static fn (array $left, array $right): int => $left['column'] <=> $right['column']);

        return [
            'name' => $worksheet->getTitle(),
            'sheet_index' => $index,
            'visibility' => $this->visibility($worksheet),
            'used_range' => $maxRow === 0 ? null : 'A1:' . Coordinate::address($maxRow, $maxColumn),
            'max_row' => $maxRow,
            'max_column' => $maxColumn,
            'cells' => $cells,
            'merged_ranges' => $mergedRanges,
            'row_properties' => $rowProperties,
            'column_properties' => $columnProperties,
        ];
    }

    /** @return array<string, mixed> */
    private function readCell(Cell $cell, int $row, int $column): array
    {
        $formula = null;
        if ($cell->getDataType() === DataType::TYPE_FORMULA) {
            $formula = (string) $cell->getValue();
            $value = method_exists($cell, 'getOldCalculatedValue') ? $cell->getOldCalculatedValue() : null;
            [$value, $valueType] = $this->naturalValue($value, null);
        } else {
            [$value, $valueType] = $this->naturalValue($cell->getValue(), $cell->getDataType());
        }
        $result = [
            'address' => Coordinate::address($row, $column),
            'row' => $row,
            'column' => $column,
            'value' => $value,
            'value_type' => $valueType,
        ];
        if ($formula !== null) {
            $result['formula'] = $formula;
        }

        return $result;
    }

    /** @param mixed $value @return array{0: mixed, 1: string} */
    private function naturalValue($value, ?string $dataType): array
    {
        if ($value instanceof RichText) {
            return [$value->getPlainText(), 'string'];
        }
        if ($value === null) {
            return [null, 'blank'];
        }
        if ($dataType === DataType::TYPE_ERROR) {
            return [(string) $value, 'error'];
        }
        if (is_string($value) && in_array($value, ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA'], true)) {
            return [$value, 'error'];
        }
        if (is_bool($value)) {
            return [$value, 'boolean'];
        }
        if (is_int($value) || is_float($value)) {
            return [(float) $value, 'number'];
        }

        return [(string) $value, 'string'];
    }

    private function visibility(Worksheet $worksheet): string
    {
        switch ($worksheet->getSheetState()) {
            case Worksheet::SHEETSTATE_HIDDEN:
                return 'hidden';
            case Worksheet::SHEETSTATE_VERYHIDDEN:
                return 'very_hidden';
            default:
                return 'visible';
        }
    }
}
