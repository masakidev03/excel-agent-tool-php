<?php

declare(strict_types=1);

namespace ExcelAgentTool\Retrieval;

final class InMemoryIndex
{
    public const MAX_SEARCH_RESULTS = 20;

    /** @var array<string, array<string, mixed>> */
    private array $records = [];

    /** @param list<array<string, mixed>> $documents @return list<array<string, mixed>> */
    public static function recordsFromDocuments(array $documents): array
    {
        return array_map(static function (array $document): array {
            $record = [
                'document_id' => $document['document_id'],
                'text' => $document['content']['text'],
                'source_id' => $document['source']['source_id'],
                'source_revision' => $document['source']['source_revision'],
                'sheet' => $document['location']['sheet'],
                'row_start' => $document['location']['row_start'],
                'row_end' => $document['location']['row_end'],
            ];
            if (isset($document['location']['sheet_index'])) {
                $record['sheet_index'] = $document['location']['sheet_index'];
            }

            return $record;
        }, $documents);
    }

    /** @param list<array<string, mixed>> $records */
    public function replaceSource(string $sourceId, array $records): void
    {
        $ids = [];
        foreach ($records as $record) {
            $documentId = is_string($record['document_id'] ?? null) ? $record['document_id'] : '';
            try {
                RetrievalModel::assertIndexRecord($record);
            } catch (\ExcelAgentTool\Contract\ContractException $exception) {
                throw new IndexException('invalid index record: ' . $exception->getMessage(), 0, $exception);
            }
            if (($record['source_id'] ?? null) !== $sourceId) {
                throw new IndexException(sprintf('record %s belongs to another source', json_encode($documentId)));
            }
            if (isset($ids[$documentId])) {
                throw new IndexException(sprintf('duplicate document ID %s', json_encode($documentId)));
            }
            $ids[$documentId] = true;
            if (isset($record['embedding'])) {
                $embedding = $record['embedding'];
                if (
                    !is_array($embedding)
                    || ($embedding['document_id'] ?? null) !== $documentId
                    || !is_int($embedding['dimensions'] ?? null)
                    || !is_array($embedding['values'] ?? null)
                    || $embedding['dimensions'] !== count($embedding['values'])
                ) {
                    throw new IndexException(sprintf('record %s has inconsistent embedding dimensions', json_encode($documentId)));
                }
            }
        }
        foreach ($this->records as $id => $record) {
            if ($record['source_id'] === $sourceId) {
                unset($this->records[$id]);
            }
        }
        foreach ($records as $record) {
            $this->records[$record['document_id']] = $record;
        }
    }

    public function deleteSource(string $sourceId): void
    {
        foreach ($this->records as $id => $record) {
            if ($record['source_id'] === $sourceId) {
                unset($this->records[$id]);
            }
        }
    }

    /**
     * @param array{text: string, filter?: array<string, string|null>, limit: int} $query
     * @return list<array{document_id: string, score: float}>
     */
    public function search(array $query): array
    {
        try {
            RetrievalModel::assertSearchQuery($query);
        } catch (\ExcelAgentTool\Contract\ContractException $exception) {
            throw new IndexException('invalid search query: ' . $exception->getMessage(), 0, $exception);
        }
        $limit = $query['limit'];
        if ($limit < 1 || $limit > self::MAX_SEARCH_RESULTS) {
            throw new IndexException('search limit must be between 1 and 20');
        }
        $queryTokens = $this->tokens($query['text']);
        if ($queryTokens === []) {
            throw new IndexException('search query must contain searchable text');
        }
        $filter = $query['filter'] ?? [];
        $hits = [];
        foreach ($this->records as $record) {
            if (!$this->matchesFilter($record, $filter)) {
                continue;
            }
            $recordTokens = $this->tokens($record['text']);
            $overlap = count(array_intersect_key($queryTokens, $recordTokens));
            if ($overlap > 0) {
                $hits[] = [
                    'document_id' => $record['document_id'],
                    'score' => $overlap / max(count($queryTokens), 1),
                ];
            }
        }
        usort($hits, static function (array $left, array $right): int {
            if ($left['score'] === $right['score']) {
                return strcmp($left['document_id'], $right['document_id']);
            }

            return $left['score'] < $right['score'] ? 1 : -1;
        });

        return array_slice($hits, 0, $limit);
    }

    /** @return array<string, true> */
    private function tokens(string $text): array
    {
        $parts = preg_split('/[^\p{L}\p{N}]+/u', $text, -1, PREG_SPLIT_NO_EMPTY);
        if ($parts === false) {
            return [];
        }
        $tokens = [];
        foreach ($parts as $token) {
            $normalized = function_exists('mb_strtolower') ? mb_strtolower($token, 'UTF-8') : strtolower($token);
            $tokens[$normalized] = true;
        }

        return $tokens;
    }

    /** @param array<string, mixed> $record @param array<string, string|null> $filter */
    private function matchesFilter(array $record, array $filter): bool
    {
        foreach (['source_id', 'source_revision', 'sheet'] as $key) {
            if (isset($filter[$key]) && $filter[$key] !== $record[$key]) {
                return false;
            }
        }

        return true;
    }
}
