<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

final class PathSafety
{
    public static function input(string $path): string
    {
        self::localXlsx($path, 'input');
        $real = realpath($path);
        if ($real === false || !is_file($real)) {
            throw new EngineException('input_read_failed', sprintf('input workbook does not exist or is not a regular file: %s', $path), $path);
        }

        return $real;
    }

    public static function output(string $path, ?string $protectedInput = null): string
    {
        self::localXlsx($path, 'output');
        $parent = dirname($path);
        $realParent = realpath($parent);
        if ($realParent === false || !is_dir($realParent)) {
            throw new EngineException('output_write_failed', sprintf('output parent is not an existing directory: %s', $parent), $path);
        }
        $resolved = $realParent . DIRECTORY_SEPARATOR . basename($path);
        $existingResolved = file_exists($path) ? realpath($path) : false;
        if ($protectedInput !== null && (
            self::samePath($protectedInput, $resolved)
            || ($existingResolved !== false && self::samePath($protectedInput, $existingResolved))
        )) {
            throw new ValidationException('edit plan failed validation', [
                Validator::error('input_output_path_conflict', 'input and output resolve to the same protected path'),
            ]);
        }
        if (file_exists($path)) {
            throw new EngineException('output_exists', sprintf('refusing to overwrite existing output: %s', $path), $path);
        }

        return $resolved;
    }

    public static function reserve(string $path): void
    {
        $handle = @fopen($path, 'x+b');
        if ($handle === false) {
            $code = file_exists($path) ? 'output_exists' : 'output_write_failed';
            throw new EngineException($code, sprintf('failed to create output: %s', $path), $path);
        }
        fclose($handle);
    }

    private static function localXlsx(string $path, string $kind): void
    {
        if (preg_match('#^[A-Za-z][A-Za-z0-9+.-]*://#', $path) || strtolower(pathinfo($path, PATHINFO_EXTENSION)) !== 'xlsx') {
            throw new EngineException('unsupported_file_type', sprintf('expected a local .xlsx %s: %s', $kind, $path), $path);
        }
    }

    private static function samePath(string $left, string $right): bool
    {
        if (DIRECTORY_SEPARATOR === '\\') {
            return strtolower($left) === strtolower($right);
        }

        return $left === $right;
    }
}
