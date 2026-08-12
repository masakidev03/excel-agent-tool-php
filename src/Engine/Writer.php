<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Throwable;

final class Writer
{
    /** @param array<string, mixed> $plan */
    public function validate(array $plan): void
    {
        $errors = Validator::createPlan($plan);
        if ($errors !== []) {
            throw new ValidationException('create plan failed validation', $errors);
        }
    }

    /** @param array<string, mixed> $plan */
    public function write(array $plan, string $outputPath): string
    {
        $this->validate($plan);
        $output = PathSafety::output($outputPath);
        $spreadsheet = new Spreadsheet();
        try {
            foreach ($plan['sheets'] as $index => $sheetPlan) {
                if ($index === 0) {
                    $worksheet = $spreadsheet->getActiveSheet();
                    $worksheet->setTitle($sheetPlan['name']);
                } else {
                    $worksheet = $spreadsheet->createSheet();
                    $worksheet->setTitle($sheetPlan['name']);
                }
                foreach ($sheetPlan['cells'] as $input) {
                    $cell = $worksheet->getCell(strtoupper($input['address']));
                    $this->writeValue($cell, $input['value'], $input['value_type']);
                    if (isset($input['style'])) {
                        StyleMapper::apply($cell, $input['style']);
                    }
                }
                foreach ($sheetPlan['merged_ranges'] as $range) {
                    $worksheet->mergeCells(strtoupper($range['range']));
                }
                foreach ($sheetPlan['row_properties'] as $properties) {
                    $worksheet->getRowDimension($properties['row'])->setRowHeight((float) $properties['height']);
                }
                foreach ($sheetPlan['column_properties'] as $properties) {
                    $column = preg_replace('/[0-9]+$/', '', Coordinate::address(1, $properties['column']));
                    $worksheet->getColumnDimension($column)->setWidth((float) $properties['width']);
                }
            }
            $spreadsheet->setActiveSheetIndex(0);
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
                    sprintf('failed to write XLSX output %s: %s', $output, $exception->getMessage()),
                    $output,
                    $exception
                );
            }

            return $output;
        } finally {
            $spreadsheet->disconnectWorksheets();
        }
    }

    /** @param mixed $value */
    public function writeValue(\PhpOffice\PhpSpreadsheet\Cell\Cell $cell, $value, string $valueType): void
    {
        switch ($valueType) {
            case 'blank':
                $cell->setValueExplicit(null, DataType::TYPE_NULL);
                break;
            case 'string':
                $cell->setValueExplicit($value, DataType::TYPE_STRING);
                break;
            case 'number':
                $cell->setValueExplicit($value, DataType::TYPE_NUMERIC);
                break;
            case 'boolean':
                $cell->setValueExplicit($value, DataType::TYPE_BOOL);
                break;
        }
    }
}

