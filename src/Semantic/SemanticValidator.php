<?php

declare(strict_types=1);

namespace ExcelAgentTool\Semantic;

final class SemanticValidator
{
    /**
     * @param array<string, mixed> $knowledge
     * @param list<array<string, mixed>> $documents
     * @return list<array<string, mixed>>
     */
    public static function validate(array $knowledge, array $documents): array
    {
        SemanticModel::assertKnowledge($knowledge);
        $errors = [];
        if (
            self::blank($knowledge['knowledge_id'])
            || self::blank($knowledge['source']['source_id'])
            || self::blank($knowledge['source']['source_revision'])
            || self::blank($knowledge['producer']['generated_at'])
        ) {
            $errors[] = self::error(
                'missing_required_value',
                'knowledge_id, source identity, source revision, and generated_at must be non-empty'
            );
        }
        if ($knowledge['producer']['kind'] === 'llm') {
            foreach (['provider', 'model', 'prompt_revision'] as $key) {
                if (!isset($knowledge['producer'][$key]) || self::blank($knowledge['producer'][$key])) {
                    $errors[] = self::error('missing_required_value', 'llm producers require provider, model, and prompt_revision');
                    break;
                }
            }
        }

        $documentsById = [];
        foreach ($documents as $document) {
            $documentsById[$document['document_id']] = $document;
        }
        $assertionIds = [];
        foreach ($knowledge['assertions'] as $index => $assertion) {
            $assertionId = $assertion['assertion_id'];
            if (self::blank($assertionId) || self::blank($assertion['kind']) || self::blank($assertion['statement'])) {
                $errors[] = self::error(
                    'missing_required_value',
                    'assertion_id, kind, and statement must be non-empty',
                    $index,
                    $assertionId
                );
            }
            if (isset($assertionIds[$assertionId])) {
                $errors[] = self::error(
                    'duplicate_assertion_id',
                    'assertion_id must be unique within SemanticKnowledge',
                    $index,
                    $assertionId
                );
            }
            $assertionIds[$assertionId] = true;
            $confidence = (float) $assertion['confidence'];
            if (!is_finite($confidence) || $confidence < 0.0 || $confidence > 1.0) {
                $errors[] = self::error(
                    'invalid_confidence',
                    'confidence must be finite and between 0.0 and 1.0',
                    $index,
                    $assertionId
                );
            }
            $evidence = $assertion['evidence'];
            $documentIds = $evidence['retrieval_document_ids'] ?? [];
            $sourceRefs = $evidence['source_refs'] ?? [];
            if ($documentIds === [] && $sourceRefs === []) {
                $errors[] = self::error(
                    'missing_evidence',
                    'each assertion must cite a retrieval document or direct source range',
                    $index,
                    $assertionId
                );
            }

            $cited = [];
            foreach ($documentIds as $documentId) {
                if (isset($cited[$documentId])) {
                    continue;
                }
                $cited[$documentId] = true;
                if (!isset($documentsById[$documentId])) {
                    $errors[] = self::error(
                        'unknown_retrieval_document',
                        sprintf('unknown retrieval document %s', json_encode($documentId)),
                        $index,
                        $assertionId
                    );
                    continue;
                }
                if (!self::sameSource($documentsById[$documentId]['source'], $knowledge['source'])) {
                    $errors[] = self::error(
                        'source_revision_mismatch',
                        sprintf('retrieval document %s belongs to another source revision', json_encode($documentId)),
                        $index,
                        $assertionId
                    );
                }
            }

            foreach ($sourceRefs as $sourceRef) {
                $parsed = self::parseRange($sourceRef['range']['range']);
                $endpoints = [self::parseCell($sourceRef['range']['start']) ?? [0, 0], self::parseCell($sourceRef['range']['end']) ?? [0, 0]];
                $validRange = $parsed !== null && $parsed === $endpoints;
                $covered = false;
                foreach ($documents as $document) {
                    if (!self::sameSource($document['source'], $knowledge['source'])) {
                        continue;
                    }
                    foreach ($document['source_refs'] as $candidate) {
                        if ($candidate['sheet'] === $sourceRef['sheet'] && self::contains($candidate, $sourceRef)) {
                            $covered = true;
                            break 2;
                        }
                    }
                }
                if (!$validRange || !$covered) {
                    $errors[] = self::error(
                        'invalid_source_reference',
                        sprintf(
                            'source reference %s!%s is malformed or not covered by this source revision',
                            $sourceRef['sheet'],
                            $sourceRef['range']['range']
                        ),
                        $index,
                        $assertionId
                    );
                }
            }
        }

        return $errors;
    }

    private static function blank(string $value): bool
    {
        return preg_match('/^[\s\p{Z}]*$/u', $value) === 1;
    }

    /** @param array<string, mixed> $left @param array<string, mixed> $right */
    private static function sameSource(array $left, array $right): bool
    {
        foreach (['source_id', 'source_revision', 'raw_schema_version'] as $key) {
            if (($left[$key] ?? null) !== ($right[$key] ?? null)) {
                return false;
            }
        }

        return true;
    }

    /** @return array<string, mixed> */
    private static function error(string $code, string $message, ?int $index = null, ?string $id = null): array
    {
        $error = ['schema_version' => '1.0', 'code' => $code, 'message' => $message];
        if ($index !== null) {
            $error['assertion_index'] = $index;
        }
        if ($id !== null) {
            $error['assertion_id'] = $id;
        }

        return $error;
    }

    /** @return array{0: int, 1: int}|null */
    private static function parseCell(string $address): ?array
    {
        if (!preg_match('/^([A-Z]+)([0-9]+)$/D', $address, $matches)) {
            return null;
        }
        $column = 0;
        foreach (str_split($matches[1]) as $letter) {
            $column = ($column * 26) + ord($letter) - ord('A') + 1;
        }
        $row = (int) $matches[2];

        return $row > 0 && $column > 0 ? [$row, $column] : null;
    }

    /** @return array{0: array{0: int, 1: int}, 1: array{0: int, 1: int}}|null */
    private static function parseRange(string $range): ?array
    {
        if (substr_count($range, ':') !== 1) {
            return null;
        }
        [$startText, $endText] = explode(':', $range, 2);
        $start = self::parseCell($startText);
        $end = self::parseCell($endText);
        if ($start === null || $end === null || $start[0] > $end[0] || $start[1] > $end[1]) {
            return null;
        }

        return [$start, $end];
    }

    /** @param array<string, mixed> $outer @param array<string, mixed> $inner */
    private static function contains(array $outer, array $inner): bool
    {
        $outerRange = self::parseRange($outer['range']['range']);
        $innerRange = self::parseRange($inner['range']['range']);
        if ($outerRange === null || $innerRange === null) {
            return false;
        }

        return $outerRange[0][0] <= $innerRange[0][0]
            && $outerRange[0][1] <= $innerRange[0][1]
            && $outerRange[1][0] >= $innerRange[1][0]
            && $outerRange[1][1] >= $innerRange[1][1];
    }
}
