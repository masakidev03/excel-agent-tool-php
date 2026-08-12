<?php

declare(strict_types=1);

namespace ExcelAgentTool\Semantic;

use ExcelAgentTool\Contract\ContractException;
use ExcelAgentTool\Retrieval\RetrievalModel;
use JsonException;

final class SemanticModel
{
    /** @return array<string, mixed> */
    public static function knowledgeFromFile(string $path): array
    {
        $contents = @file_get_contents($path);
        if ($contents === false) {
            throw new ContractException("failed to read semantic knowledge $path");
        }
        try {
            $knowledge = json_decode($contents, true, 512, JSON_THROW_ON_ERROR);
        } catch (JsonException $exception) {
            throw new ContractException("invalid semantic knowledge JSON in $path: " . $exception->getMessage(), 0, $exception);
        }
        $knowledge = self::object($knowledge, 'semantic knowledge');
        self::assertKnowledge($knowledge);
        foreach ($knowledge['assertions'] as &$assertion) {
            $assertion['evidence']['retrieval_document_ids'] = $assertion['evidence']['retrieval_document_ids'] ?? [];
            $assertion['evidence']['source_refs'] = $assertion['evidence']['source_refs'] ?? [];
        }
        unset($assertion);

        return $knowledge;
    }

    /** @param array<string, mixed> $knowledge */
    public static function assertKnowledge(array $knowledge): void
    {
        self::exact($knowledge, ['schema_version', 'knowledge_id', 'source', 'producer', 'assertions'], 'semantic knowledge');
        self::version($knowledge, 'semantic knowledge');
        self::string($knowledge, 'knowledge_id', 'semantic knowledge');
        RetrievalModel::assertRetrievalSource(self::requiredObject($knowledge, 'source', 'semantic knowledge'), 'semantic knowledge.source');

        $producer = self::requiredObject($knowledge, 'producer', 'semantic knowledge');
        self::exactOptional($producer, ['kind', 'generated_at'], ['provider', 'model', 'prompt_revision'], 'semantic knowledge.producer');
        $kind = self::string($producer, 'kind', 'semantic knowledge.producer');
        if (!in_array($kind, ['llm', 'human', 'process'], true)) {
            self::fail('semantic knowledge.producer.kind is unsupported');
        }
        foreach (['provider', 'model', 'prompt_revision'] as $key) {
            self::optionalString($producer, $key, 'semantic knowledge.producer');
        }
        self::string($producer, 'generated_at', 'semantic knowledge.producer');

        $assertions = self::list($knowledge, 'assertions', 'semantic knowledge');
        foreach ($assertions as $index => $assertion) {
            $context = "semantic knowledge.assertions[$index]";
            $assertion = self::object($assertion, $context);
            self::exact($assertion, ['assertion_id', 'kind', 'statement', 'confidence', 'status', 'evidence'], $context);
            foreach (['assertion_id', 'kind', 'statement'] as $key) {
                self::string($assertion, $key, $context);
            }
            if (!is_int($assertion['confidence']) && !is_float($assertion['confidence'])) {
                self::fail("$context.confidence must be a number");
            }
            $status = self::string($assertion, 'status', $context);
            if (!in_array($status, ['unreviewed', 'accepted', 'rejected', 'superseded'], true)) {
                self::fail("$context.status is unsupported");
            }
            $evidence = self::requiredObject($assertion, 'evidence', $context);
            self::exactOptional($evidence, [], ['retrieval_document_ids', 'source_refs'], "$context.evidence");
            if (array_key_exists('retrieval_document_ids', $evidence)) {
                $ids = self::list($evidence, 'retrieval_document_ids', "$context.evidence");
                foreach ($ids as $idIndex => $id) {
                    if (!is_string($id)) {
                        self::fail("$context.evidence.retrieval_document_ids[$idIndex] must be a string");
                    }
                }
            }
            if (array_key_exists('source_refs', $evidence)) {
                $refs = self::list($evidence, 'source_refs', "$context.evidence");
                foreach ($refs as $refIndex => $ref) {
                    RetrievalModel::assertSourceRef(self::object($ref, "$context.evidence.source_refs[$refIndex]"), "$context.evidence.source_refs[$refIndex]");
                }
            }
        }
    }

    /** @param array<string, mixed> $value */
    private static function version(array $value, string $context): void
    {
        if (($value['schema_version'] ?? null) !== '1.0') {
            self::fail("$context.schema_version must be \"1.0\"");
        }
    }

    /** @param mixed $value @return array<string, mixed> */
    private static function object($value, string $context): array
    {
        if (!is_array($value) || ($value !== [] && self::isList($value))) {
            self::fail("$context must be an object");
        }

        return $value;
    }

    /** @param array<string, mixed> $value @return array<string, mixed> */
    private static function requiredObject(array $value, string $key, string $context): array
    {
        if (!array_key_exists($key, $value)) {
            self::fail("$context.$key is required");
        }

        return self::object($value[$key], "$context.$key");
    }

    /** @param array<string, mixed> $value @return list<mixed> */
    private static function list(array $value, string $key, string $context): array
    {
        if (!array_key_exists($key, $value) || !is_array($value[$key]) || !self::isList($value[$key])) {
            self::fail("$context.$key must be an array");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function string(array $value, string $key, string $context): string
    {
        if (!array_key_exists($key, $value) || !is_string($value[$key])) {
            self::fail("$context.$key must be a string");
        }

        return $value[$key];
    }

    /** @param array<string, mixed> $value */
    private static function optionalString(array $value, string $key, string $context): void
    {
        if (array_key_exists($key, $value) && $value[$key] !== null && !is_string($value[$key])) {
            self::fail("$context.$key must be a string or null");
        }
    }

    /** @param array<string, mixed> $value @param list<string> $required */
    private static function exact(array $value, array $required, string $context): void
    {
        self::exactOptional($value, $required, [], $context);
    }

    /** @param array<string, mixed> $value @param list<string> $required @param list<string> $optional */
    private static function exactOptional(array $value, array $required, array $optional, string $context): void
    {
        foreach ($required as $key) {
            if (!array_key_exists($key, $value)) {
                self::fail("$context.$key is required");
            }
        }
        $allowed = array_merge($required, $optional);
        foreach (array_keys($value) as $key) {
            if (!in_array($key, $allowed, true)) {
                self::fail("$context contains unknown field $key");
            }
        }
    }

    /** @param array<mixed> $value */
    private static function isList(array $value): bool
    {
        $expected = 0;
        foreach ($value as $key => $_) {
            if ($key !== $expected++) {
                return false;
            }
        }

        return true;
    }

    private static function fail(string $message): void
    {
        throw new ContractException($message);
    }
}
