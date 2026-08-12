<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

use RuntimeException;
use Throwable;

class EngineException extends RuntimeException
{
    private string $stableCode;
    private ?string $relatedPath;

    public function __construct(string $stableCode, string $message, ?string $path = null, ?Throwable $previous = null)
    {
        parent::__construct($message, 0, $previous);
        $this->stableCode = $stableCode;
        $this->relatedPath = $path;
    }

    public function stableCode(): string
    {
        return $this->stableCode;
    }

    public function relatedPath(): ?string
    {
        return $this->relatedPath;
    }
}

