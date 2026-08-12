<?php

declare(strict_types=1);

namespace ExcelAgentTool\Engine;

final class ValidationException extends EngineException
{
    /** @var list<array<string, mixed>> */
    private array $errors;

    /** @param list<array<string, mixed>> $errors */
    public function __construct(string $message, array $errors)
    {
        parent::__construct('validation_failed', $message);
        $this->errors = $errors;
    }

    /** @return list<array<string, mixed>> */
    public function errors(): array
    {
        return $this->errors;
    }
}

