<?php

declare(strict_types=1);

namespace ExcelAgentTool\Retrieval;

use ExcelAgentTool\Contract\ContractException;

final class RetrievalCatalog
{
    private InMemoryIndex $index;
    /** @var array<string, array{document: array<string, mixed>, source_locator: string|null}> */
    private array $documents = [];
    /** @var array<string, array{source_revision: string, document_ids: list<string>}> */
    private array $activeSources = [];

    public function __construct(?InMemoryIndex $index = null)
    {
        $this->index = $index ?? new InMemoryIndex();
    }

    /** @param array<string, mixed> $build @return array<string, mixed> */
    public function replaceBuild(array $build, ?string $sourceLocator = null): array
    {
        $this->validateBuild($build);
        $sourceId = $build['manifest']['source_id'];
        $sourceRevision = $build['manifest']['source_revision'];
        $indexedDocumentIds = $build['manifest']['document_ids'];
        foreach ($indexedDocumentIds as $documentId) {
            if (isset($this->documents[$documentId]) && $this->documents[$documentId]['document']['source']['source_id'] !== $sourceId) {
                throw new CatalogException(sprintf('retrieval document ID %s is already owned by another source', json_encode($documentId)));
            }
        }

        $nextIndex = clone $this->index;
        try {
            $nextIndex->replaceSource($sourceId, InMemoryIndex::recordsFromDocuments($build['documents']));
        } catch (IndexException $exception) {
            throw new CatalogException('retrieval index rejected operation: ' . $exception->getMessage(), 0, $exception);
        }

        $previous = $this->activeSources[$sourceId] ?? null;
        $removed = $previous['document_ids'] ?? [];
        $removed = array_values(array_filter($removed, static fn (string $id): bool => !in_array($id, $indexedDocumentIds, true)));
        sort($removed, SORT_STRING);
        if ($previous !== null) {
            foreach ($previous['document_ids'] as $documentId) {
                unset($this->documents[$documentId]);
            }
        }
        foreach ($build['documents'] as $document) {
            $this->documents[$document['document_id']] = ['document' => $document, 'source_locator' => $sourceLocator];
        }
        $this->activeSources[$sourceId] = ['source_revision' => $sourceRevision, 'document_ids' => $indexedDocumentIds];
        $this->index = $nextIndex;

        $summary = [
            'source_id' => $sourceId,
            'source_revision' => $sourceRevision,
            'removed_document_ids' => $removed,
            'indexed_document_ids' => $indexedDocumentIds,
        ];
        if ($previous !== null) {
            $summary['previous_source_revision'] = $previous['source_revision'];
        }

        return $summary;
    }

    /** @param array{text: string, filter?: array<string, string|null>, limit: int} $query @return list<array<string, mixed>> */
    public function search(array $query): array
    {
        try {
            $indexHits = $this->index->search($query);
        } catch (IndexException $exception) {
            throw new CatalogException('retrieval index rejected operation: ' . $exception->getMessage(), 0, $exception);
        }
        $hits = [];
        foreach ($indexHits as $hit) {
            if (!isset($this->documents[$hit['document_id']])) {
                throw new CatalogException(sprintf('retrieval index references missing catalog document %s', json_encode($hit['document_id'])));
            }
            $entry = $this->documents[$hit['document_id']];
            $result = ['score' => $hit['score'], 'document' => $entry['document']];
            if ($entry['source_locator'] !== null) {
                $result['source_locator'] = $entry['source_locator'];
            }
            $hits[] = $result;
        }

        return $hits;
    }

    public function documentCount(): int
    {
        return count($this->documents);
    }

    /** @param array<string, mixed> $build */
    private function validateBuild(array $build): void
    {
        try {
            RetrievalModel::assertBuild($build);
        } catch (ContractException $exception) {
            throw new CatalogException('invalid retrieval manifest: ' . $exception->getMessage(), 0, $exception);
        }
        $manifest = $build['manifest'];
        if ($this->blank($manifest['source_id']) || $this->blank($manifest['source_revision'])) {
            throw new CatalogException('invalid retrieval manifest: source_id and source_revision must be non-empty');
        }
        $documentIds = array_map(static fn (array $document): string => $document['document_id'], $build['documents']);
        if ($documentIds !== $manifest['document_ids']) {
            throw new CatalogException('invalid retrieval manifest: document_ids must exactly match documents in deterministic order');
        }
        $unique = [];
        foreach ($build['documents'] as $document) {
            $documentId = $document['document_id'];
            if (isset($unique[$documentId])) {
                throw new CatalogException(sprintf('invalid retrieval manifest: document ID %s occurs more than once', json_encode($documentId)));
            }
            $unique[$documentId] = true;
            if ($document['source']['source_id'] !== $manifest['source_id'] || $document['source']['source_revision'] !== $manifest['source_revision']) {
                throw new CatalogException(sprintf('invalid retrieval manifest: document %s belongs to another source revision', json_encode($documentId)));
            }
            if ($document['chunk']['profile'] !== $manifest['profile'] || $document['chunk']['profile_version'] !== $manifest['profile_version']) {
                throw new CatalogException(sprintf('invalid retrieval manifest: document %s belongs to another retrieval profile', json_encode($documentId)));
            }
        }
    }

    private function blank(string $value): bool
    {
        return preg_match('/^[\s\p{Z}]*$/u', $value) === 1;
    }
}
