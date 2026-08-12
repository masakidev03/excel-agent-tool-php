<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

final class Coordinate
{
    public const MAX_ROWS = 1048576;
    public const MAX_COLUMNS = 16384;

    /** @return array{row: int, column: int}|null */
    public static function parseCell(string $reference): ?array
    {
        if (!preg_match('/^([A-Za-z]{1,3})([1-9][0-9]*)$/D', $reference, $matches)) {
            return null;
        }

        $column = 0;
        foreach (str_split(strtoupper($matches[1])) as $character) {
            $column = ($column * 26) + (ord($character) - ord('A') + 1);
        }
        $row = (int) $matches[2];
        if ($row < 1 || $row > self::MAX_ROWS || $column < 1 || $column > self::MAX_COLUMNS) {
            return null;
        }

        return ['row' => $row, 'column' => $column];
    }

    /** @return array{start: array{row: int, column: int}, end: array{row: int, column: int}}|null */
    public static function parseRange(string $reference): ?array
    {
        if (substr_count($reference, ':') !== 1) {
            return null;
        }
        [$startText, $endText] = explode(':', $reference, 2);
        $start = self::parseCell($startText);
        $end = self::parseCell($endText);
        if ($start === null || $end === null || $start['row'] > $end['row'] || $start['column'] > $end['column']) {
            return null;
        }

        return ['start' => $start, 'end' => $end];
    }

    public static function address(int $row, int $column): string
    {
        if ($row < 1 || $row > self::MAX_ROWS || $column < 1 || $column > self::MAX_COLUMNS) {
            throw new EngineException('coordinate_out_of_range', 'cell coordinate is outside XLSX limits');
        }

        $letters = '';
        while ($column > 0) {
            --$column;
            $letters = chr(ord('A') + ($column % 26)) . $letters;
            $column = intdiv($column, 26);
        }

        return $letters . $row;
    }

    /** @param array{start: array{row: int, column: int}, end: array{row: int, column: int}} $left
     *  @param array{start: array{row: int, column: int}, end: array{row: int, column: int}} $right
     */
    public static function rangesOverlap(array $left, array $right): bool
    {
        return $left['start']['row'] <= $right['end']['row']
            && $right['start']['row'] <= $left['end']['row']
            && $left['start']['column'] <= $right['end']['column']
            && $right['start']['column'] <= $left['end']['column'];
    }

    /** @param array{start: array{row: int, column: int}, end: array{row: int, column: int}} $range
     *  @param array{row: int, column: int} $cell
     */
    public static function rangeContains(array $range, array $cell): bool
    {
        return $cell['row'] >= $range['start']['row'] && $cell['row'] <= $range['end']['row']
            && $cell['column'] >= $range['start']['column'] && $cell['column'] <= $range['end']['column'];
    }
}

