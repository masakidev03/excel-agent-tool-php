<?php

declare(strict_types=1);

use ExcelAgentTool\Contract\PlanDecoder;
use ExcelAgentTool\Cli\Application;
use ExcelAgentTool\Engine\ExcelEngine;
use ExcelAgentTool\Engine\ValidationException;
use ExcelAgentTool\Retrieval\CatalogException;
use ExcelAgentTool\Retrieval\InMemoryIndex;
use ExcelAgentTool\Retrieval\IndexException;
use ExcelAgentTool\Retrieval\RetrievalBuilder;
use ExcelAgentTool\Retrieval\RetrievalCatalog;
use ExcelAgentTool\Retrieval\RetrievalException;
use ExcelAgentTool\Retrieval\RetrievalModel;
use ExcelAgentTool\Semantic\SemanticModel;
use ExcelAgentTool\Semantic\SemanticValidator;

require dirname(__DIR__) . '/vendor/autoload.php';

set_error_handler(static function (int $severity, string $message, string $file, int $line): bool {
    if ((error_reporting() & $severity) === 0) {
        return false;
    }
    throw new ErrorException($message, 0, $severity, $file, $line);
});

final class TestRunner
{
    private int $assertions = 0;

    public function run(): void
    {
        $directory = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'excel-agent-tool-' . bin2hex(random_bytes(8));
        if (!mkdir($directory, 0700) && !is_dir($directory)) {
            throw new RuntimeException('failed to create test directory');
        }

        try {
            $this->testCoordinateAndContract($directory);
            $this->testCreateReadEditAndSafety($directory);
            $this->testRetrievalAndSemantic($directory);
            fwrite(STDOUT, sprintf("OK (%d assertions)\n", $this->assertions));
        } finally {
            $this->removeTree($directory);
        }
    }

    private function testCoordinateAndContract(string $directory): void
    {
        $valid = $directory . DIRECTORY_SEPARATOR . 'contract.json';
        file_put_contents($valid, json_encode($this->createPlan(), JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES));
        $decoded = PlanDecoder::createPlanFromFile($valid);
        $this->same('1.0', $decoded['schema_version'], 'schema version should decode');

        $invalid = $directory . DIRECTORY_SEPARATOR . 'invalid-contract.json';
        file_put_contents($invalid, json_encode([
            'schema_version' => '2.0',
            'sheets' => [],
        ]));
        $this->throws(
            static fn () => PlanDecoder::createPlanFromFile($invalid),
            \ExcelAgentTool\Contract\ContractException::class,
            'unsupported schema version must fail during decoding'
        );

        $mismatch = $directory . DIRECTORY_SEPARATOR . 'mismatched-contract.json';
        $plan = $this->createPlan();
        $plan['sheets'][0]['cells'][0]['value_type'] = 'number';
        file_put_contents($mismatch, json_encode($plan));
        $this->throws(
            static fn () => PlanDecoder::createPlanFromFile($mismatch),
            \ExcelAgentTool\Contract\ContractException::class,
            'incompatible cell value and value_type must fail during decoding'
        );
    }

    private function testCreateReadEditAndSafety(string $directory): void
    {
        $engine = new ExcelEngine();
        $input = $directory . DIRECTORY_SEPARATOR . 'created.xlsx';
        $engine->validateCreatePlan($this->createPlan());
        $engine->writeXlsx($this->createPlan(), $input);
        $this->true(is_file($input), 'writer should create an XLSX file');

        $raw = $engine->readXlsx($input);
        $this->same(2, $raw['metadata']['worksheet_count'], 'reader should observe both sheets');
        $this->same('Data', $raw['sheets'][0]['name'], 'sheet order should be retained');
        $this->same('A1:B2', $raw['sheets'][0]['used_range'], 'used range should include merge and data');
        $this->same('A1:B1', $raw['sheets'][0]['merged_ranges'][0]['range'], 'merge should round-trip');
        $this->same(28.0, $raw['sheets'][0]['row_properties'][0]['height'], 'row height should round-trip');
        $amount = $this->cell($raw, 'Data', 'B2');
        $this->same(12.5, $amount['value'], 'numeric cell should retain its value');
        $this->same('0.00', $amount['style']['number_format'], 'number format should round-trip');

        $invalidCreateOutput = $directory . DIRECTORY_SEPARATOR . 'invalid-create.xlsx';
        $invalidCreate = $this->createPlan();
        $invalidCreate['sheets'][0]['cells'][0]['address'] = 'A0';
        try {
            $engine->writeXlsx($invalidCreate, $invalidCreateOutput);
            throw new RuntimeException('invalid create unexpectedly succeeded');
        } catch (ValidationException $exception) {
            $this->same('invalid_cell_reference', $exception->errors()[0]['code'], 'invalid create address should be structured');
        }
        $this->true(!file_exists($invalidCreateOutput), 'invalid create must not create output');

        $before = file_get_contents($input);
        $edited = $directory . DIRECTORY_SEPARATOR . 'edited.xlsx';
        $plan = [
            'schema_version' => '1.0',
            'operations' => [
                [
                    'operation' => 'set_cell_value',
                    'sheet' => 'Data',
                    'address' => 'A2',
                    'value' => 'Updated',
                    'value_type' => 'string',
                ],
                [
                    'operation' => 'set_cell_formula',
                    'sheet' => 'Data',
                    'address' => 'B2',
                    'formula' => '=2+3',
                ],
                [
                    'operation' => 'set_cell_style',
                    'sheet' => 'Data',
                    'address' => 'A2',
                    'style' => ['font' => ['bold' => true]],
                ],
                [
                    'operation' => 'merge_cells',
                    'sheet' => 'Data',
                    'range' => ['range' => 'C1:D1', 'start' => 'C1', 'end' => 'D1'],
                ],
                [
                    'operation' => 'set_column_width',
                    'sheet' => 'Data',
                    'column' => 3,
                    'width' => 17.0,
                ],
                [
                    'operation' => 'set_row_height',
                    'sheet' => 'Data',
                    'row' => 2,
                    'height' => 24.0,
                ],
                [
                    'operation' => 'append_rows',
                    'sheet' => 'Data',
                    'rows' => [[
                        ['value' => 'Appended', 'value_type' => 'string'],
                        ['value' => true, 'value_type' => 'boolean'],
                    ]],
                ],
            ],
        ];
        $engine->validateEditPlan($input, $edited, $plan);
        $this->true(!file_exists($edited), 'validation-only API must not create output');
        $engine->editXlsx($input, $edited, $plan);
        $this->same($before, file_get_contents($input), 'editing must not mutate the input workbook');

        $rawEdited = $engine->readXlsx($edited);
        $this->same('Updated', $this->cell($rawEdited, 'Data', 'A2')['value'], 'value edit should be observable');
        $this->true($this->cell($rawEdited, 'Data', 'A2')['style']['font']['bold'], 'style patch should be observable');
        $this->same('=2+3', $this->cell($rawEdited, 'Data', 'B2')['formula'], 'formula edit should be observable');
        $this->same('Appended', $this->cell($rawEdited, 'Data', 'A3')['value'], 'append should start below used cells');
        $this->same('C1:D1', $rawEdited['sheets'][0]['merged_ranges'][1]['range'], 'new merge should be observable');

        $invalidOutput = $directory . DIRECTORY_SEPARATOR . 'invalid-output.xlsx';
        $invalidPlan = [
            'schema_version' => '1.0',
            'operations' => [
                [
                    'operation' => 'set_cell_value',
                    'sheet' => 'Missing',
                    'address' => 'A1',
                    'value' => 'x',
                    'value_type' => 'string',
                ],
                [
                    'operation' => 'set_cell_formula',
                    'sheet' => 'Data',
                    'address' => 'A0',
                    'formula' => '=1',
                ],
            ],
        ];
        try {
            $engine->editXlsx($input, $invalidOutput, $invalidPlan);
            throw new RuntimeException('invalid edit unexpectedly succeeded');
        } catch (ValidationException $exception) {
            $this->same('worksheet_not_found', $exception->errors()[0]['code'], 'missing sheet should be structured');
            $this->same(0, $exception->errors()[0]['operation_index'], 'missing sheet index should be retained');
            $this->same('invalid_cell_reference', $exception->errors()[1]['code'], 'invalid address should be structured');
        }
        $this->true(!file_exists($invalidOutput), 'invalid edit must not create output');

        $duplicateOutput = $directory . DIRECTORY_SEPARATOR . 'duplicate-output.xlsx';
        $duplicate = [
            'schema_version' => '1.0',
            'operations' => [
                ['operation' => 'set_cell_value', 'sheet' => 'Data', 'address' => 'A2', 'value' => 'one', 'value_type' => 'string'],
                ['operation' => 'set_cell_value', 'sheet' => 'Data', 'address' => 'a2', 'value' => 'one', 'value_type' => 'string'],
            ],
        ];
        try {
            $engine->editXlsx($input, $duplicateOutput, $duplicate);
            throw new RuntimeException('duplicate edit unexpectedly succeeded');
        } catch (ValidationException $exception) {
            $this->same('duplicate_operation', $exception->errors()[0]['code'], 'normalized duplicate should be rejected');
            $this->same(1, $exception->errors()[0]['operation_index'], 'later duplicate index should be retained');
        }

        $duplicateMergeOutput = $directory . DIRECTORY_SEPARATOR . 'duplicate-merge-output.xlsx';
        $duplicateMerge = [
            'schema_version' => '1.0',
            'operations' => [
                ['operation' => 'merge_cells', 'sheet' => 'Data', 'range' => ['range' => 'C1:D1', 'start' => 'C1', 'end' => 'D1']],
                ['operation' => 'merge_cells', 'sheet' => 'Data', 'range' => ['range' => 'c1:d1', 'start' => 'c1', 'end' => 'd1']],
            ],
        ];
        try {
            $engine->editXlsx($input, $duplicateMergeOutput, $duplicateMerge);
            throw new RuntimeException('duplicate merge unexpectedly succeeded');
        } catch (ValidationException $exception) {
            $this->same('duplicate_operation', $exception->errors()[0]['code'], 'normalized duplicate merge should be rejected');
        }
        $this->true(!file_exists($duplicateMergeOutput), 'duplicate merge must not create output');

        $createPlan = $this->createPlan();
        $this->throws(
            static fn () => $engine->writeXlsx($createPlan, $input),
            \ExcelAgentTool\Engine\EngineException::class,
            'writer must refuse existing output'
        );
        $this->throws(
            static fn () => $engine->editXlsx($input, $input, ['schema_version' => '1.0', 'operations' => []]),
            ValidationException::class,
            'editor must refuse input/output identity'
        );

        $stdout = fopen('php://temp', 'w+b');
        $stderr = fopen('php://temp', 'w+b');
        if ($stdout === false || $stderr === false) {
            throw new RuntimeException('failed to open CLI test streams');
        }
        $application = new Application(null, null, null, $stdout, $stderr);
        $status = $application->run(['excel-tool', 'inspect', $input]);
        rewind($stdout);
        rewind($stderr);
        $cliResult = json_decode((string) stream_get_contents($stdout), true, 512, JSON_THROW_ON_ERROR);
        $this->same(0, $status, 'inspect CLI should succeed');
        $this->same('inspect', $cliResult['command'], 'inspect CLI should return the command name');
        $this->same('', (string) stream_get_contents($stderr), 'successful CLI should not write stderr');
        fclose($stdout);
        fclose($stderr);

        $stdout = fopen('php://temp', 'w+b');
        $stderr = fopen('php://temp', 'w+b');
        $application = new Application(null, null, null, $stdout, $stderr);
        $status = $application->run(['excel-tool', 'unknown']);
        rewind($stderr);
        $cliError = json_decode((string) stream_get_contents($stderr), true, 512, JSON_THROW_ON_ERROR);
        $this->same(2, $status, 'unknown CLI command should be an input error');
        $this->same('invalid_arguments', $cliError['error']['code'], 'CLI input error should be structured');
        fclose($stdout);
        fclose($stderr);
    }

    private function testRetrievalAndSemantic(string $directory): void
    {
        $engine = new ExcelEngine();
        $input = $directory . DIRECTORY_SEPARATOR . 'retrieval-input.xlsx';
        $engine->writeXlsx($this->createPlan(), $input);
        $raw = $engine->readXlsx($input);
        $source = RetrievalModel::source('book-1', 'revision-1', 'Test workbook');
        $builder = new RetrievalBuilder();
        $first = $builder->build($raw, $source);
        $second = $builder->build($raw, $source);
        $this->same($first, $second, 'retrieval builds must be deterministic');
        $this->same(2, count($first['documents']), 'one default chunk should be emitted per non-empty sheet');
        $this->true(strpos($first['documents'][0]['document_id'], 'rdoc_') === 0, 'retrieval ID should use rdoc prefix');
        $this->same(69, strlen($first['documents'][0]['document_id']), 'retrieval ID should contain a SHA-256 digest');
        $this->true(strpos($first['documents'][0]['content']['text'], "cell\tA1") !== false, 'canonical text should include source addresses');
        $this->same($first['manifest']['document_ids'][0], $first['documents'][0]['document_id'], 'manifest order should match documents');

        $embeddingRequest = RetrievalModel::embeddingRequest(
            $first['documents'][0]['document_id'],
            'embedding-profile-1',
            $first['documents'][0]['content']['text']
        );
        RetrievalModel::assertEmbeddingRequest($embeddingRequest);
        $embedding = RetrievalModel::embeddingVector($embeddingRequest['document_id'], 'embedding-profile-1', [0.25, -0.5]);
        $this->same(2, $embedding['dimensions'], 'embedding vectors should declare their exact dimensions');
        $this->throws(
            static fn () => RetrievalModel::embeddingVector('rdoc_invalid', 'embedding-profile-1', ['not-a-number']),
            \ExcelAgentTool\Contract\ContractException::class,
            'embedding values must not be silently coerced'
        );
        $records = InMemoryIndex::recordsFromDocuments($first['documents']);
        $records[0]['embedding'] = $embedding;
        $embeddingIndex = new InMemoryIndex();
        $embeddingIndex->replaceSource('book-1', $records);
        $this->throws(
            static function () use ($embeddingIndex, $records): void {
                $invalidRecords = $records;
                $invalidRecords[0]['embedding']['dimensions'] = 3;
                $embeddingIndex->replaceSource('book-1', $invalidRecords);
            },
            IndexException::class,
            'embedding dimension mismatches must be rejected'
        );

        $profile = RetrievalModel::defaultProfile();
        $profile['target_rows'] = 2;
        $profile['overlap_rows'] = 1;
        $profile['max_addressed_cells'] = 1;
        $split = $builder->build($raw, $source, $profile);
        $this->true(count($split['documents']) > count($first['documents']), 'hard area limit should split chunks');
        foreach ($split['documents'] as $document) {
            $area = ($document['location']['row_end'] - $document['location']['row_start'] + 1)
                * ($document['location']['column_end'] - $document['location']['column_start'] + 1);
            $this->true($area <= 1, 'split chunks must obey max_addressed_cells');
        }

        $tooSmall = RetrievalModel::defaultProfile();
        $tooSmall['max_text_chars'] = 1;
        try {
            $builder->build($raw, $source, $tooSmall);
            throw new RuntimeException('text overflow unexpectedly succeeded');
        } catch (RetrievalException $exception) {
            $this->same('text_limit_exceeded', $exception->issue()['code'], 'indivisible text overflow should be structured');
        }

        $catalog = new RetrievalCatalog();
        $summary = $catalog->replaceBuild($first, 'retrieval-input.xlsx');
        $this->same('revision-1', $summary['source_revision'], 'catalog should activate the build revision');
        $hits = $catalog->search(['text' => 'Report Amount', 'filter' => ['source_id' => 'book-1'], 'limit' => 10]);
        $this->true(count($hits) >= 1, 'lexical catalog should return overlapping tokens');
        $this->same('retrieval-input.xlsx', $hits[0]['source_locator'], 'catalog should retain adapter locator outside documents');

        $revisedSource = RetrievalModel::source('book-1', 'revision-2');
        $revised = $builder->build($raw, $revisedSource);
        $replacement = $catalog->replaceBuild($revised, 'retrieval-input.xlsx');
        $this->same('revision-1', $replacement['previous_source_revision'], 'replacement should report previous revision');
        $this->same([], $catalog->search(['text' => 'Report', 'filter' => ['source_revision' => 'revision-1'], 'limit' => 10]), 'stale revision must be removed');

        $invalidBuild = $first;
        $invalidBuild['manifest']['document_ids'] = [];
        $this->throws(
            static fn () => (new RetrievalCatalog())->replaceBuild($invalidBuild),
            CatalogException::class,
            'manifest mismatch must be rejected'
        );

        $document = $first['documents'][0];
        $knowledge = [
            'schema_version' => '1.0',
            'knowledge_id' => 'knowledge-1',
            'source' => $document['source'],
            'producer' => [
                'kind' => 'llm',
                'provider' => 'provider-a',
                'model' => 'model-a',
                'prompt_revision' => 'prompt-1',
                'generated_at' => '2026-08-13T00:00:00Z',
            ],
            'assertions' => [[
                'assertion_id' => 'assertion-1',
                'kind' => 'possible_header',
                'statement' => 'The first row may be a header.',
                'confidence' => 0.8,
                'status' => 'unreviewed',
                'evidence' => [
                    'retrieval_document_ids' => [$document['document_id']],
                    'source_refs' => [[
                        'sheet' => $document['location']['sheet'],
                        'range' => ['range' => 'A1:A1', 'start' => 'A1', 'end' => 'A1'],
                    ]],
                ],
            ]],
        ];
        SemanticModel::assertKnowledge($knowledge);
        $this->same([], SemanticValidator::validate($knowledge, $first['documents']), 'valid semantic citations should pass');

        $invalidKnowledge = $knowledge;
        $invalidKnowledge['assertions'][0]['confidence'] = 1.5;
        $invalidKnowledge['assertions'][0]['evidence']['retrieval_document_ids'] = ['missing'];
        $semanticErrors = SemanticValidator::validate($invalidKnowledge, $first['documents']);
        $codes = array_column($semanticErrors, 'code');
        $this->true(in_array('invalid_confidence', $codes, true), 'invalid confidence should be structured');
        $this->true(in_array('unknown_retrieval_document', $codes, true), 'unknown retrieval document should be structured');

        $duplicateKnowledge = $knowledge;
        $duplicateKnowledge['assertions'][] = $duplicateKnowledge['assertions'][0];
        $codes = array_column(SemanticValidator::validate($duplicateKnowledge, $first['documents']), 'code');
        $this->true(in_array('duplicate_assertion_id', $codes, true), 'duplicate semantic assertion IDs should be rejected');

        $nullEvidenceKnowledge = $knowledge;
        $nullEvidenceKnowledge['assertions'][0]['evidence']['source_refs'] = null;
        $this->throws(
            static fn () => SemanticModel::assertKnowledge($nullEvidenceKnowledge),
            \ExcelAgentTool\Contract\ContractException::class,
            'semantic evidence arrays must reject null instead of treating it as absent'
        );

        $retrievalPath = $directory . DIRECTORY_SEPARATOR . 'retrieval-build.json';
        $stdout = fopen('php://temp', 'w+b');
        $stderr = fopen('php://temp', 'w+b');
        $application = new Application(null, null, null, $stdout, $stderr);
        $status = $application->run([
            'excel-tool', 'retrieve', $input,
            '--source-id', 'book-cli',
            '--source-revision', 'revision-cli',
            '--output', $retrievalPath,
        ]);
        $this->same(0, $status, 'retrieve CLI should succeed');
        $this->true(is_file($retrievalPath), 'retrieve CLI should create requested JSON output');
        fclose($stdout);
        fclose($stderr);

        $stdout = fopen('php://temp', 'w+b');
        $stderr = fopen('php://temp', 'w+b');
        $application = new Application(null, null, null, $stdout, $stderr);
        $status = $application->run(['excel-tool', 'search', $retrievalPath, 'Report', '--limit', '5']);
        rewind($stdout);
        $search = json_decode((string) stream_get_contents($stdout), true, 512, JSON_THROW_ON_ERROR);
        $this->same(0, $status, 'search CLI should succeed');
        $this->same('search', $search['command'], 'search CLI should identify its command');
        $this->true(count($search['hits']) >= 1, 'search CLI should return complete document hits');
        fclose($stdout);
        fclose($stderr);
    }

    /** @return array<string, mixed> */
    private function createPlan(): array
    {
        return [
            'schema_version' => '1.0',
            'sheets' => [
                [
                    'name' => 'Data',
                    'cells' => [
                        [
                            'address' => 'A1',
                            'value' => 'Report',
                            'value_type' => 'string',
                            'style' => [
                                'font' => ['bold' => true, 'color' => 'FFFFFFFF'],
                                'fill' => ['pattern' => 'solid', 'foreground_color' => 'FF1F4E78'],
                                'alignment' => ['horizontal' => 'center'],
                            ],
                        ],
                        ['address' => 'A2', 'value' => 'Amount', 'value_type' => 'string'],
                        [
                            'address' => 'B2',
                            'value' => 12.5,
                            'value_type' => 'number',
                            'style' => ['number_format' => '0.00'],
                        ],
                    ],
                    'merged_ranges' => [
                        ['range' => 'A1:B1', 'start' => 'A1', 'end' => 'B1'],
                    ],
                    'row_properties' => [
                        ['row' => 1, 'height' => 28.0, 'hidden' => false],
                    ],
                    'column_properties' => [
                        ['column' => 1, 'width' => 18.0, 'hidden' => false],
                        ['column' => 2, 'width' => 14.0, 'hidden' => false],
                    ],
                ],
                [
                    'name' => 'Notes',
                    'cells' => [
                        ['address' => 'A1', 'value' => 'Ready', 'value_type' => 'string'],
                    ],
                    'merged_ranges' => [],
                    'row_properties' => [],
                    'column_properties' => [],
                ],
            ],
        ];
    }

    /** @param array<string, mixed> $raw @return array<string, mixed> */
    private function cell(array $raw, string $sheetName, string $address): array
    {
        foreach ($raw['sheets'] as $sheet) {
            if ($sheet['name'] !== $sheetName) {
                continue;
            }
            foreach ($sheet['cells'] as $cell) {
                if ($cell['address'] === $address) {
                    return $cell;
                }
            }
        }
        throw new RuntimeException("cell not found: $sheetName!$address");
    }

    /** @param mixed $expected @param mixed $actual */
    private function same($expected, $actual, string $message): void
    {
        ++$this->assertions;
        if ($expected !== $actual) {
            throw new RuntimeException(sprintf('%s; expected %s, got %s', $message, var_export($expected, true), var_export($actual, true)));
        }
    }

    private function true(bool $condition, string $message): void
    {
        ++$this->assertions;
        if (!$condition) {
            throw new RuntimeException($message);
        }
    }

    private function throws(callable $callable, string $class, string $message): void
    {
        ++$this->assertions;
        try {
            $callable();
        } catch (Throwable $exception) {
            if ($exception instanceof $class) {
                return;
            }
            throw new RuntimeException(sprintf('%s; expected %s, got %s', $message, $class, get_class($exception)), 0, $exception);
        }
        throw new RuntimeException($message . '; no exception was thrown');
    }

    private function removeTree(string $directory): void
    {
        if (!is_dir($directory)) {
            return;
        }
        foreach (scandir($directory) ?: [] as $name) {
            if ($name === '.' || $name === '..') {
                continue;
            }
            $path = $directory . DIRECTORY_SEPARATOR . $name;
            if (is_dir($path)) {
                $this->removeTree($path);
            } else {
                @unlink($path);
            }
        }
        @rmdir($directory);
    }
}

(new TestRunner())->run();
