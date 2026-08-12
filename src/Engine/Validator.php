<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

final class Validator
{
    private const VALUE_TYPES = ['blank', 'string', 'number', 'boolean', 'error'];
    private const FILL_PATTERNS = [
        'none', 'solid', 'medium_gray', 'dark_gray', 'light_gray', 'dark_horizontal',
        'dark_vertical', 'dark_down', 'dark_up', 'dark_grid', 'dark_trellis',
        'light_horizontal', 'light_vertical', 'light_down', 'light_up', 'light_grid',
        'light_trellis', 'gray_125', 'gray_0625',
    ];

    /** @param array<string, mixed> $plan @return list<array<string, mixed>> */
    public static function createPlan(array $plan): array
    {
        $errors = [];
        $names = [];
        if ($plan['sheets'] === []) {
            $errors[] = self::error('unsupported_operation', 'create plan must contain at least one worksheet');
        }

        foreach ($plan['sheets'] as $sheet) {
            $name = $sheet['name'];
            if (!self::validWorksheetName($name)) {
                $errors[] = self::error(
                    'invalid_worksheet_name',
                    sprintf('worksheet name %s violates Excel naming rules', json_encode($name)),
                    null,
                    $name
                );
            }
            $normalizedName = self::lower($name);
            if (isset($names[$normalizedName])) {
                $errors[] = self::error(
                    'duplicate_worksheet_name',
                    sprintf('worksheet name %s is duplicated', json_encode($name)),
                    null,
                    $name
                );
            }
            $names[$normalizedName] = true;

            $ranges = [];
            foreach ($sheet['merged_ranges'] as $descriptor) {
                $range = self::mergedRange($descriptor, $name, null, $errors);
                if ($range === null) {
                    continue;
                }
                foreach ($ranges as $existing) {
                    if (Coordinate::rangesOverlap($existing, $range)) {
                        $errors[] = self::error(
                            'invalid_range',
                            sprintf('merged range %s overlaps an earlier merged range', $descriptor['range']),
                            null,
                            $name,
                            null,
                            $descriptor['range']
                        );
                        continue 2;
                    }
                }
                $ranges[] = $range;
            }

            $coordinates = [];
            foreach ($sheet['cells'] as $cell) {
                $coordinate = Coordinate::parseCell($cell['address']);
                if ($coordinate === null) {
                    $errors[] = self::error(
                        'invalid_cell_reference',
                        sprintf('cell address %s is invalid or outside XLSX limits', json_encode($cell['address'])),
                        null,
                        $name,
                        $cell['address']
                    );
                    continue;
                }
                $key = $coordinate['row'] . ':' . $coordinate['column'];
                if (isset($coordinates[$key])) {
                    $errors[] = self::error(
                        'invalid_cell_reference',
                        sprintf('cell address %s is specified more than once', json_encode($cell['address'])),
                        null,
                        $name,
                        $cell['address']
                    );
                }
                $coordinates[$key] = true;
                foreach ($ranges as $range) {
                    if (Coordinate::rangeContains($range, $coordinate) && $coordinate !== $range['start']) {
                        $errors[] = self::error(
                            'invalid_cell_reference',
                            sprintf('cell %s targets a non-anchor cell inside a merged range', json_encode($cell['address'])),
                            null,
                            $name,
                            $cell['address']
                        );
                        break;
                    }
                }
                self::cellInput($cell, $name, $cell['address'], null, true, $errors);
            }

            $rows = [];
            foreach ($sheet['row_properties'] as $row) {
                if ($row['row'] < 1 || $row['row'] > Coordinate::MAX_ROWS || isset($rows[$row['row']])) {
                    $errors[] = self::error('invalid_range', 'row property index is invalid or duplicated', null, $name);
                }
                $rows[$row['row']] = true;
                if (!isset($row['height']) || !self::finiteInRange($row['height'] ?? null, 0.0, 409.5, false)) {
                    $errors[] = self::error('invalid_value', 'row height must be finite and in the range 0 < height <= 409.5', null, $name);
                }
                if ($row['hidden']) {
                    $errors[] = self::error('unsupported_operation', 'Writer MVP does not support hidden rows', null, $name);
                }
                if (isset($row['style'])) {
                    $errors[] = self::error('unsupported_operation', 'Writer MVP does not support row styles', null, $name);
                }
            }

            $columns = [];
            foreach ($sheet['column_properties'] as $column) {
                if ($column['column'] < 1 || $column['column'] > Coordinate::MAX_COLUMNS || isset($columns[$column['column']])) {
                    $errors[] = self::error('invalid_range', 'column property index is invalid or duplicated', null, $name);
                }
                $columns[$column['column']] = true;
                if (!isset($column['width']) || !self::finiteInRange($column['width'] ?? null, 0.0, 255.0, false)) {
                    $errors[] = self::error('invalid_value', 'column width must be finite and in the range 0 < width <= 255', null, $name);
                }
                if ($column['hidden']) {
                    $errors[] = self::error('unsupported_operation', 'Writer MVP does not support hidden columns', null, $name);
                }
                if (isset($column['style'])) {
                    $errors[] = self::error('unsupported_operation', 'Writer MVP does not support column styles', null, $name);
                }
            }
        }

        return $errors;
    }

    /** @param array<string, mixed> $input @param list<array<string, mixed>> $errors */
    public static function cellInput(
        array $input,
        string $sheet,
        string $address,
        ?int $operationIndex,
        bool $creating,
        array &$errors
    ): void {
        $formula = $input['formula'] ?? null;
        $valid = $formula !== null
            ? $input['value'] === null && $input['value_type'] === 'blank'
            : self::compatible($input['value_type'], $input['value']);
        if (!$valid) {
            $errors[] = self::error(
                'invalid_value',
                $creating ? 'cell value does not match the declared value type' : 'appended cell value, value_type, and formula are inconsistent',
                $operationIndex,
                $sheet,
                $address
            );
        } elseif ($formula !== null) {
            if ($creating) {
                $errors[] = self::error('unsupported_operation', 'Writer MVP does not support formulas', $operationIndex, $sheet, $address);
            } else {
                self::formula($formula, $sheet, $address, $operationIndex, $errors);
            }
        } else {
            self::literal($input['value'], $input['value_type'], $sheet, $address, $operationIndex, $creating, $errors);
        }

        if (isset($input['comment'])) {
            $errors[] = self::error(
                'unsupported_operation',
                $creating ? 'Writer MVP does not support comments' : 'Editor MVP does not support comments in appended rows',
                $operationIndex,
                $sheet,
                $address
            );
        }
        if (isset($input['hyperlink'])) {
            $errors[] = self::error(
                'unsupported_operation',
                $creating ? 'Writer MVP does not support hyperlinks' : 'Editor MVP does not support hyperlinks in appended rows',
                $operationIndex,
                $sheet,
                $address
            );
        }
        if (isset($input['style'])) {
            self::style($input['style'], $sheet, $address, $operationIndex, $errors);
        }
    }

    /** @param mixed $value @param list<array<string, mixed>> $errors */
    public static function literal(
        $value,
        string $valueType,
        string $sheet,
        string $address,
        ?int $operationIndex,
        bool $creating,
        array &$errors
    ): void {
        if (!self::compatible($valueType, $value)) {
            $errors[] = self::error('invalid_value', 'cell value is incompatible with value_type', $operationIndex, $sheet, $address);
            return;
        }
        if ($valueType === 'error') {
            $errors[] = self::error(
                'unsupported_operation',
                ($creating ? 'Writer' : 'Editor') . ' MVP does not support writing Excel error cells',
                $operationIndex,
                $sheet,
                $address
            );
        } elseif (is_float($value) && !is_finite($value)) {
            $errors[] = self::error('invalid_value', 'cell number must be finite', $operationIndex, $sheet, $address);
        } elseif (is_string($value)) {
            $validString = $value !== '' && self::length($value) <= 32767 && self::validXmlText($value);
            if (!$validString) {
                $errors[] = self::error(
                    $creating && $value === '' ? 'unsupported_operation' : 'invalid_value',
                    $creating && $value === ''
                        ? 'Writer MVP does not support empty strings distinct from blank cells'
                        : 'cell string must contain 1 to 32,767 XML-compatible characters',
                    $operationIndex,
                    $sheet,
                    $address
                );
            }
        }
    }

    /** @param list<array<string, mixed>> $errors */
    public static function formula(string $formula, string $sheet, string $address, ?int $operationIndex, array &$errors): void
    {
        if (substr($formula, 0, 1) !== '=' || self::length($formula) < 2 || self::length($formula) > 8193 || !self::validXmlText($formula)) {
            $errors[] = self::error(
                'invalid_value',
                "formula must start with '=' and contain 1 to 8,192 XML-compatible characters",
                $operationIndex,
                $sheet,
                $address
            );
        }
    }

    /** @param array<string, mixed> $style @param list<array<string, mixed>> $errors */
    public static function style(array $style, string $sheet, string $address, ?int $operationIndex, array &$errors): void
    {
        $hasProperty = isset($style['number_format'])
            || (isset($style['font']) && $style['font'] !== [])
            || (isset($style['fill']) && $style['fill'] !== [])
            || (isset($style['alignment']) && $style['alignment'] !== []);
        if (!$hasProperty) {
            $errors[] = self::error('invalid_value', 'cell style must contain at least one property', $operationIndex, $sheet, $address);
            return;
        }
        if (isset($style['number_format'])) {
            $format = $style['number_format'];
            if ($format === '' || self::length($format) > 255 || strpos($format, "\0") !== false) {
                $errors[] = self::error('invalid_value', 'number format must contain 1 to 255 valid characters', $operationIndex, $sheet, $address);
            }
        }
        if (isset($style['font'])) {
            $font = $style['font'];
            if (isset($font['name']) && ($font['name'] === '' || self::length($font['name']) > 255 || preg_match('/[\x00-\x1F\x7F]/', $font['name']))) {
                $errors[] = self::error('invalid_value', 'font name must contain 1 to 255 non-control characters', $operationIndex, $sheet, $address);
            }
            if (isset($font['size']) && !self::finiteInRange($font['size'], 1.0, 409.0, true)) {
                $errors[] = self::error('invalid_value', 'font size must be finite and in the range 1 <= size <= 409', $operationIndex, $sheet, $address);
            }
            if (isset($font['color'])) {
                self::color($font['color'], $sheet, $address, $operationIndex, $errors);
            }
        }
        if (isset($style['fill'])) {
            $fill = $style['fill'];
            $pattern = $fill['pattern'] ?? null;
            if ($pattern !== null && !in_array($pattern, self::FILL_PATTERNS, true)) {
                $errors[] = self::error('invalid_value', sprintf('unsupported fill pattern %s', json_encode($pattern)), $operationIndex, $sheet, $address);
            }
            if ($pattern === null && (isset($fill['foreground_color']) || isset($fill['background_color']))) {
                $errors[] = self::error('invalid_value', 'fill colors require an explicit fill pattern', $operationIndex, $sheet, $address);
            }
            if ($pattern === 'none' && (isset($fill['foreground_color']) || isset($fill['background_color']))) {
                $errors[] = self::error('invalid_value', "fill pattern 'none' cannot carry visible colors", $operationIndex, $sheet, $address);
            }
            foreach (['foreground_color', 'background_color'] as $key) {
                if (isset($fill[$key])) {
                    self::color($fill[$key], $sheet, $address, $operationIndex, $errors);
                }
            }
        }
        if (isset($style['alignment'])) {
            $alignment = $style['alignment'];
            if (isset($alignment['horizontal']) && !in_array($alignment['horizontal'], ['general', 'left', 'center', 'right', 'fill', 'justify', 'distributed'], true)) {
                $errors[] = self::error('invalid_value', 'unsupported horizontal alignment', $operationIndex, $sheet, $address);
            }
            if (isset($alignment['vertical']) && !in_array($alignment['vertical'], ['top', 'center', 'bottom', 'justify', 'distributed'], true)) {
                $errors[] = self::error('invalid_value', 'unsupported vertical alignment', $operationIndex, $sheet, $address);
            }
            if (isset($alignment['text_rotation'])) {
                $rotation = $alignment['text_rotation'];
                if (($rotation < 0 || $rotation > 90) && $rotation !== 255) {
                    $errors[] = self::error('invalid_value', 'text rotation must be in 0..=90 or 255 for stacked text', $operationIndex, $sheet, $address);
                }
            }
        }
    }

    /** @param array<string, mixed> $descriptor @param list<array<string, mixed>> $errors
     * @return array{start: array{row: int, column: int}, end: array{row: int, column: int}}|null
     */
    public static function mergedRange(array $descriptor, string $sheet, ?int $operationIndex, array &$errors): ?array
    {
        $range = Coordinate::parseRange($descriptor['range']);
        $start = Coordinate::parseCell($descriptor['start']);
        $end = Coordinate::parseCell($descriptor['end']);
        if ($range === null || $start === null || $end === null || $range['start'] !== $start || $range['end'] !== $end || $start === $end) {
            $errors[] = self::error(
                'invalid_range',
                'range text and ordered endpoints must identify at least two cells',
                $operationIndex,
                $sheet,
                null,
                $descriptor['range']
            );
            return null;
        }

        return $range;
    }

    public static function validWorksheetName(string $name): bool
    {
        return $name !== ''
            && self::length($name) <= 31
            && strpbrk($name, '*?:[]\\/') === false
            && substr($name, 0, 1) !== "'"
            && substr($name, -1) !== "'"
            && !preg_match('/[\x00-\x1F\x7F]/', $name);
    }

    /** @param mixed $value */
    public static function compatible(string $valueType, $value): bool
    {
        if (!in_array($valueType, self::VALUE_TYPES, true)) {
            return false;
        }
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
        }

        return false;
    }

    /** @return array<string, mixed> */
    public static function error(
        string $code,
        string $message,
        ?int $operationIndex = null,
        ?string $sheet = null,
        ?string $address = null,
        ?string $range = null
    ): array {
        $error = ['schema_version' => '1.0', 'code' => $code, 'message' => $message];
        foreach (['operation_index' => $operationIndex, 'sheet' => $sheet, 'address' => $address, 'range' => $range] as $key => $value) {
            if ($value !== null) {
                $error[$key] = $value;
            }
        }

        return $error;
    }

    private static function color(string $color, string $sheet, string $address, ?int $operationIndex, array &$errors): void
    {
        if (!preg_match('/^ff[0-9a-f]{6}$/iD', $color)) {
            $errors[] = self::error('invalid_value', 'color must be an opaque eight-digit ARGB value beginning with FF', $operationIndex, $sheet, $address);
        }
    }

    /** @param mixed $value */
    private static function finiteInRange($value, float $minimum, float $maximum, bool $includeMinimum): bool
    {
        if ((!is_int($value) && !is_float($value)) || !is_finite((float) $value)) {
            return false;
        }
        $number = (float) $value;

        return ($includeMinimum ? $number >= $minimum : $number > $minimum) && $number <= $maximum;
    }

    private static function validXmlText(string $value): bool
    {
        return preg_match('//u', $value) === 1 && preg_match('/[\x00-\x08\x0B\x0C\x0E-\x1F]/', $value) !== 1;
    }

    private static function length(string $value): int
    {
        return function_exists('mb_strlen') ? mb_strlen($value, 'UTF-8') : strlen($value);
    }

    private static function lower(string $value): string
    {
        return function_exists('mb_strtolower') ? mb_strtolower($value, 'UTF-8') : strtolower($value);
    }
}
