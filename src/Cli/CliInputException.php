<?php

declare(strict_types=1);

namespace ExcelAgentTool\Cli;

use RuntimeException;

final class CliInputException extends RuntimeException
{
    private string $stableCode;

    public function __construct(string $stableCode, string $message)
    {
        parent::__construct($message);
        $this->stableCode = $stableCode;
    }

    public function stableCode(): string
    {
        return $this->stableCode;
    }
}

