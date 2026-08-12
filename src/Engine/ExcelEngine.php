<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

final class ExcelEngine
{
    private Reader $reader;
    private Writer $writer;
    private Editor $editor;

    public function __construct()
    {
        $this->reader = new Reader();
        $this->writer = new Writer();
        $this->editor = new Editor($this->reader, $this->writer);
    }

    /** @return array<string, mixed> */
    public function readXlsx(string $inputPath): array
    {
        return $this->reader->read($inputPath);
    }

    /** @param array<string, mixed> $plan */
    public function validateCreatePlan(array $plan): void
    {
        $this->writer->validate($plan);
    }

    /** @param array<string, mixed> $plan */
    public function writeXlsx(array $plan, string $outputPath): string
    {
        return $this->writer->write($plan, $outputPath);
    }

    /** @param array<string, mixed> $plan */
    public function validateEditPlanOperations(string $inputPath, array $plan): void
    {
        $this->editor->validateOperations($inputPath, $plan);
    }

    /** @param array<string, mixed> $plan */
    public function validateEditPlan(string $inputPath, string $outputPath, array $plan): void
    {
        $this->editor->validate($inputPath, $outputPath, $plan);
    }

    /** @param array<string, mixed> $plan */
    public function editXlsx(string $inputPath, string $outputPath, array $plan): string
    {
        return $this->editor->edit($inputPath, $outputPath, $plan);
    }
}

