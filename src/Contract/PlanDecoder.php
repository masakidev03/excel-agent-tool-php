<?php

declare(strict_types=1);

namespace ExcelAgentTool\Contract;

use JsonException;

final class PlanDecoder
{
    /** @return array<string, mixed> */
    public static function createPlanFromFile(string $path): array
    {
        $plan = self::decodeFile($path);
        self::validateCreatePlanShape($plan);

        return $plan;
    }

    /** @return array<string, mixed> */
    public static function editPlanFromFile(string $path): array
    {
        $plan = self::decodeFile($path);
        self::validateEditPlanShape($plan);

        return $plan;
    }

    /** @return array<string, mixed> */
    private static function decodeFile(string $path): array
    {
        $contents = @file_get_contents($path);
        if ($contents === false) {
            throw new ContractException(sprintf('failed to read plan %s', $path));
        }

        try {
            $value = json_decode($contents, true, 512, JSON_THROW_ON_ERROR);
        } catch (JsonException $exception) {
            throw new ContractException(
                sprintf('invalid plan JSON in %s: %s', $path, $exception->getMessage()),
                0,
                $exception
            );
        }

        if (!is_array($value) || self::isList($value)) {
            throw new ContractException('plan JSON must be an object');
        }

        return $value;
    }

    /** @param array<string, mixed> $plan */
    private static function validateCreatePlanShape(array $plan): void
    {
        self::schema($plan);
        $sheets = self::requiredList($plan, 'sheets', 'create plan');
        foreach ($sheets as $sheetIndex => $sheet) {
            $sheet = self::object($sheet, "sheets[$sheetIndex]");
            self::requiredString($sheet, 'name', "sheets[$sheetIndex]");
            $cells = self::requiredList($sheet, 'cells', "sheets[$sheetIndex]");
            foreach ($cells as $cellIndex => $cell) {
                $cell = self::object($cell, "sheets[$sheetIndex].cells[$cellIndex]");
                self::requiredString($cell, 'address', "sheets[$sheetIndex].cells[$cellIndex]");
                self::cellInput($cell, "sheets[$sheetIndex].cells[$cellIndex]");
            }
            $ranges = self::requiredList($sheet, 'merged_ranges', "sheets[$sheetIndex]");
            foreach ($ranges as $rangeIndex => $range) {
                self::mergedRange($range, "sheets[$sheetIndex].merged_ranges[$rangeIndex]");
            }
            $rows = self::requiredList($sheet, 'row_properties', "sheets[$sheetIndex]");
            foreach ($rows as $rowIndex => $row) {
                $context = "sheets[$sheetIndex].row_properties[$rowIndex]";
                $row = self::object($row, $context);
                self::rejectUnknown($row, ['row', 'height', 'hidden', 'style'], $context);
                self::requiredInt($row, 'row', $context);
                self::optionalNumber($row, 'height', $context);
                self::requiredBool($row, 'hidden', $context);
                self::optionalStyle($row, 'style', $context);
            }
            $columns = self::requiredList($sheet, 'column_properties', "sheets[$sheetIndex]");
            foreach ($columns as $columnIndex => $column) {
                $context = "sheets[$sheetIndex].column_properties[$columnIndex]";
                $column = self::object($column, $context);
                self::rejectUnknown($column, ['column', 'width', 'hidden', 'style'], $context);
                self::requiredInt($column, 'column', $context);
                self::optionalNumber($column, 'width', $context);
                self::requiredBool($column, 'hidden', $context);
                self::optionalStyle($column, 'style', $context);
            }
        }
    }

    /** @param array<string, mixed> $plan */
    private static function validateEditPlanShape(array $plan): void
    {
        self::schema($plan);
        $operations = self::requiredList($plan, 'operations', 'edit plan');
        foreach ($operations as $index => $operation) {
            $context = "operations[$index]";
            $operation = self::object($operation, $context);
            $kind = self::requiredString($operation, 'operation', $context);
            switch ($kind) {
                case 'set_cell_value':
                    self::sheetAddress($operation, $context);
                    self::requiredScalarOrNull($operation, 'value', $context);
                    self::requiredString($operation, 'value_type', $context);
                    break;
                case 'set_cell_formula':
                    self::sheetAddress($operation, $context);
                    self::requiredString($operation, 'formula', $context);
                    break;
                case 'set_cell_style':
                    self::sheetAddress($operation, $context);
                    if (!array_key_exists('style', $operation)) {
                        self::fail("$context.style is required");
                    }
                    self::style($operation['style'], "$context.style");
                    break;
                case 'merge_cells':
                case 'unmerge_cells':
                    self::requiredString($operation, 'sheet', $context);
                    if (!array_key_exists('range', $operation)) {
                        self::fail("$context.range is required");
                    }
                    self::mergedRange($operation['range'], "$context.range");
                    break;
                case 'set_column_width':
                    self::requiredString($operation, 'sheet', $context);
                    self::requiredInt($operation, 'column', $context);
                    self::requiredNumber($operation, 'width', $context);
                    break;
                case 'set_row_height':
                    self::requiredString($operation, 'sheet', $context);
                    self::requiredInt($operation, 'row', $context);
                    self::requiredNumber($operation, 'height', $context);
                    break;
                case 'append_rows':
                    self::requiredString($operation, 'sheet', $context);
                    $rows = self::requiredList($operation, 'rows', $context);
                    foreach ($rows as $rowIndex => $row) {
                        if (!is_array($row) || !self::isList($row)) {
                            self::fail("$context.rows[$rowIndex] must be an array");
                        }
                        foreach ($row as $columnIndex => $cell) {
                            $cell = self::object($cell, "$context.rows[$rowIndex][$columnIndex]");
                            self::cellInput($cell, "$context.rows[$rowIndex][$columnIndex]");
                        }
                    }
                    break;
                case 'create_sheet':
                case 'delete_sheet':
                    self::requiredString($operation, 'name', $context);
                    break;
                case 'rename_sheet':
                    self::requiredString($operation, 'from', $context);
                    self::requiredString($operation, 'to', $context);
                    break;
                default:
                    self::fail(sprintf('%s.operation has unsupported value %s', $context, $kind));
            }
        }
    }

    /** @param array<string, mixed> $value */
    private static function schema(array $value): void
    {
        if (($value['schema_version'] ?? null) !== '1.0') {
            self::fail('schema_version must be "1.0"');
        }
    }

    /** @param array<string, mixed> $value */
    private static function sheetAddress(array $value, string $context): void
    {
        self::requiredString($value, 'sheet', $context);
        self::requiredString($value, 'address', $context);
    }

    /** @param array<string, mixed> $value */
    private static function cellInput(array $value, string $context): void
    {
        self::requiredScalarOrNull($value, 'value', $context);
        $valueType = self::requiredString($value, 'value_type', $context);
        if (array_key_exists('formula', $value) && $value['formula'] !== null && !is_string($value['formula'])) {
            self::fail("$context.formula must be a string");
        }
        $formulaPresent = array_key_exists('formula', $value) && $value['formula'] !== null;
        $compatible = $formulaPresent
            ? $value['value'] === null && $valueType === 'blank'
            : self::compatibleValue($valueType, $value['value']);
        if (!$compatible) {
            self::fail("$context has an incompatible value, value_type, and formula combination");
        }
        self::optionalStyle($value, 'style', $context);
        if (array_key_exists('comment', $value) && $value['comment'] !== null && !is_array($value['comment'])) {
            self::fail("$context.comment must be an object");
        }
        if (array_key_exists('hyperlink', $value) && $value['hyperlink'] !== null && !is_array($value['hyperlink'])) {
            self::fail("$context.hyperlink must be an object");
        }
    }

    /** @param mixed $value */
    private static function mergedRange($value, string $context): void
    {
        $value = self::object($value, $context);
        self::requiredString($value, 'range', $context);
        self::requiredString($value, 'start', $context);
        self::requiredString($value, 'end', $context);
    }

    /** @param array<string, mixed> $value */
    private static function optionalStyle(array $value, string $key, string $context): void
    {
        if (array_key_exists($key, $value) && $value[$key] !== null) {
            self::style($value[$key], "$context.$key");
        }
    }

    /** @param mixed $value */
    private static function style($value, string $context): void
    {
        $style = self::object($value, $context);
        self::rejectUnknown($style, ['number_format', 'font', 'fill', 'alignment'], $context);
        if (array_key_exists('number_format', $style) && $style['number_format'] !== null && !is_string($style['number_format'])) {
            self::fail("$context.number_format must be a string");
        }
        if (array_key_exists('font', $style) && $style['font'] !== null) {
            $font = self::object($style['font'], "$context.font");
            self::rejectUnknown($font, ['name', 'size', 'bold', 'italic', 'underline', 'color'], "$context.font");
            foreach (['name', 'color'] as $key) {
                if (array_key_exists($key, $font) && $font[$key] !== null && !is_string($font[$key])) {
                    self::fail("$context.font.$key must be a string");
                }
            }
            self::optionalNumber($font, 'size', "$context.font");
            foreach (['bold', 'italic', 'underline'] as $key) {
                if (array_key_exists($key, $font) && $font[$key] !== null && !is_bool($font[$key])) {
                    self::fail("$context.font.$key must be a boolean");
                }
            }
        }
        if (array_key_exists('fill', $style) && $style['fill'] !== null) {
            $fill = self::object($style['fill'], "$context.fill");
            self::rejectUnknown($fill, ['pattern', 'foreground_color', 'background_color'], "$context.fill");
            foreach (['pattern', 'foreground_color', 'background_color'] as $key) {
                if (array_key_exists($key, $fill) && $fill[$key] !== null && !is_string($fill[$key])) {
                    self::fail("$context.fill.$key must be a string");
                }
            }
        }
        if (array_key_exists('alignment', $style) && $style['alignment'] !== null) {
            $alignment = self::object($style['alignment'], "$context.alignment");
            self::rejectUnknown($alignment, ['horizontal', 'vertical', 'wrap_text', 'text_rotation'], "$context.alignment");
            foreach (['horizontal', 'vertical'] as $key) {
                if (array_key_exists($key, $alignment) && $alignment[$key] !== null && !is_string($alignment[$key])) {
                    self::fail("$context.alignment.$key must be a string");
                }
            }
            if (array_key_exists('wrap_text', $alignment) && $alignment['wrap_text'] !== null && !is_bool($alignment['wrap_text'])) {
                self::fail("$context.alignment.wrap_text must be a boolean");
            }
            if (array_key_exists('text_rotation', $alignment) && $alignment['text_rotation'] !== null && !is_int($alignment['text_rotation'])) {
                self::fail("$context.alignment.text_rotation must be an integer");
            }
        }
    }

    /** @param array<string, mixed> $value @return list<mixed> */
    private static function requiredList(array $value, string $key, string $context): array
    {
        if (!array_key_exists($key, $value) || !is_array($value[$key]) || !self::isList($value[$key])) {
            self::fail("$context.$key must be an array");
        }

        return $value[$key];
    }

    /** @param mixed $value @return array<string, mixed> */
    private static function object($value, string $context): array
    {
        if (!is_array($value) || ($value !== [] && self::isList($value))) {
            self::fail("$context must be an object");
        }

        return $value;
    }

    /** @param array<string, mixed> $value */
    private static function requiredString(array $value, string $key, string $context): string
    {
        if (!array_key_exists($key, $value) || !is_string($value[$key])) {
            self::fail("$context.$key must be a string");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function requiredInt(array $value, string $key, string $context): int
    {
        if (!array_key_exists($key, $value) || !is_int($value[$key])) {
            self::fail("$context.$key must be an integer");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function requiredNumber(array $value, string $key, string $context): float
    {
        if (!array_key_exists($key, $value) || (!is_int($value[$key]) && !is_float($value[$key]))) {
            self::fail("$context.$key must be a number");
        }

        return (float) $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function optionalNumber(array $value, string $key, string $context): void
    {
        if (array_key_exists($key, $value) && $value[$key] !== null && !is_int($value[$key]) && !is_float($value[$key])) {
            self::fail("$context.$key must be a number");
        }
    }

    /** @param array<string, mixed> $value */
    private static function requiredBool(array $value, string $key, string $context): bool
    {
        if (!array_key_exists($key, $value) || !is_bool($value[$key])) {
            self::fail("$context.$key must be a boolean");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function requiredScalarOrNull(array $value, string $key, string $context): void
    {
        if (!array_key_exists($key, $value) || (!is_null($value[$key]) && !is_scalar($value[$key]))) {
            self::fail("$context.$key must be a JSON scalar or null");
        }
    }

    /** @param array<string, mixed> $value @param list<string> $allowed */
    private static function rejectUnknown(array $value, array $allowed, string $context): void
    {
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
            if ($key !== $expected) {
                return false;
            }
            ++$expected;
        }

        return true;
    }

    /** @param mixed $value */
    private static function compatibleValue(string $valueType, $value): bool
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

    private static function fail(string $message): void
    {
        throw new ContractException($message);
    }
}
