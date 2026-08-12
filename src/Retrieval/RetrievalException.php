<?php

declare(strict_types=1);

namespace ExcelAgentTool\Retrieval;

use RuntimeException;

final class RetrievalException extends RuntimeException
{
    /** @var array<string, mixed> */
    private array $issue;

    /** @param array<string, mixed> $issue */
    public function __construct(array $issue)
    {
        parent::__construct((string) $issue['message']);
        $this->issue = $issue;
    }

    /** @return array<string, mixed> */
    public function issue(): array
    {
        return $this->issue;
    }
}

