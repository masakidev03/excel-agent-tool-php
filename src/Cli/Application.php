<?php

declare(strict_types=1);

namespace ExcelAgentTool\Cli;

use ExcelAgentTool\Contract\ContractException;
use ExcelAgentTool\Contract\PlanDecoder;
use ExcelAgentTool\Engine\Editor;
use ExcelAgentTool\Engine\EngineException;
use ExcelAgentTool\Engine\Reader;
use ExcelAgentTool\Engine\ValidationException;
use ExcelAgentTool\Engine\Writer;
use ExcelAgentTool\Retrieval\CatalogException;
use ExcelAgentTool\Retrieval\RetrievalBuilder;
use ExcelAgentTool\Retrieval\RetrievalCatalog;
use ExcelAgentTool\Retrieval\RetrievalException;
use ExcelAgentTool\Retrieval\RetrievalModel;
use JsonException;
use Throwable;

final class Application
{
    private Reader $reader;
    private Writer $writer;
    private Editor $editor;
    /** @var resource */
    private $stdout;
    /** @var resource */
    private $stderr;

    /** @param resource|null $stdout @param resource|null $stderr */
    public function __construct(
        ?Reader $reader = null,
        ?Writer $writer = null,
        ?Editor $editor = null,
        $stdout = null,
        $stderr = null
    )
    {
        $this->reader = $reader ?? new Reader();
        $this->writer = $writer ?? new Writer();
        $this->editor = $editor ?? new Editor($this->reader, $this->writer);
        $this->stdout = $stdout ?? STDOUT;
        $this->stderr = $stderr ?? STDERR;
    }

    /** @param list<string> $arguments */
    public function run(array $arguments): int
    {
        array_shift($arguments);
        if ($arguments === ['--help'] || $arguments === ['-h'] || $arguments === []) {
            fwrite($this->stdout, $this->help());
            return 0;
        }
        if ($arguments === ['--version'] || $arguments === ['-V']) {
            fwrite($this->stdout, "excel-tool 0.1.0\n");
            return 0;
        }

        $command = array_shift($arguments);
        try {
            switch ($command) {
                case 'inspect':
                    $this->requireCount($command, $arguments, 1);
                    $raw = $this->reader->read($arguments[0]);
                    $sheets = [];
                    foreach ($raw['sheets'] as $sheet) {
                        $sheets[] = [
                            'name' => $sheet['name'],
                            'sheet_index' => $sheet['sheet_index'],
                            'visibility' => $sheet['visibility'],
                            'used_range' => $sheet['used_range'],
                            'max_row' => $sheet['max_row'],
                            'max_column' => $sheet['max_column'],
                            'cell_count' => count($sheet['cells']),
                            'merged_range_count' => count($sheet['merged_ranges']),
                        ];
                    }
                    $this->writeJson($this->stdout, [
                        'schema_version' => '1.0',
                        'ok' => true,
                        'command' => 'inspect',
                        'input' => $arguments[0],
                        'summary' => [
                            'source_format' => 'xlsx',
                            'worksheet_count' => count($sheets),
                            'sheets' => $sheets,
                        ],
                    ]);
                    break;

                case 'extract':
                    [$input, $output] = $this->extractArguments($arguments);
                    $raw = $this->reader->read($input);
                    if ($output === null) {
                        $this->writeJson($this->stdout, $raw);
                    } else {
                        $this->writeNewJson($output, $raw);
                        $this->success('extract', $output);
                    }
                    break;

                case 'retrieve':
                    $options = $this->retrieveArguments($arguments);
                    $profile = RetrievalModel::defaultProfile();
                    if ($options['profile'] !== null) {
                        try {
                            $profile = RetrievalModel::profileFromFile($options['profile']);
                        } catch (ContractException $exception) {
                            $code = strpos($exception->getMessage(), 'failed to read') === 0
                                ? 'profile_read_failed'
                                : 'invalid_retrieval_profile_json';
                            throw new CliInputException($code, $exception->getMessage());
                        }
                    }
                    $raw = $this->reader->read($options['input']);
                    $source = RetrievalModel::source(
                        $options['source_id'],
                        $options['source_revision'],
                        $options['display_name']
                    );
                    $build = (new RetrievalBuilder())->build($raw, $source, $profile);
                    if ($options['output'] === null) {
                        $this->writeJson($this->stdout, $build);
                    } else {
                        $this->writeNewJson($options['output'], $build);
                        $this->success('retrieve', $options['output']);
                    }
                    break;

                case 'search':
                    $options = $this->searchArguments($arguments);
                    try {
                        $build = RetrievalModel::buildFromFile($options['retrieval']);
                    } catch (ContractException $exception) {
                        $code = strpos($exception->getMessage(), 'failed to read') === 0
                            ? 'retrieval_read_failed'
                            : 'invalid_retrieval_json';
                        throw new CliInputException($code, $exception->getMessage());
                    }
                    try {
                        $catalog = new RetrievalCatalog();
                        $catalog->replaceBuild($build);
                    } catch (CatalogException $exception) {
                        throw new CliInputException('invalid_retrieval_build', $exception->getMessage());
                    }
                    try {
                        $hits = $catalog->search([
                            'text' => $options['query'],
                            'filter' => array_filter([
                                'source_id' => $options['source_id'],
                                'source_revision' => $options['source_revision'],
                                'sheet' => $options['sheet'],
                            ], static fn ($value): bool => $value !== null),
                            'limit' => $options['limit'],
                        ]);
                    } catch (CatalogException $exception) {
                        throw new EngineException('retrieval_search_failed', $exception->getMessage(), null, $exception);
                    }
                    $this->writeJson($this->stdout, [
                        'schema_version' => '1.0',
                        'ok' => true,
                        'command' => 'search',
                        'hits' => $hits,
                    ]);
                    break;

                case 'create':
                    $this->requireCount($command, $arguments, 2);
                    $plan = PlanDecoder::createPlanFromFile($arguments[0]);
                    $this->writer->write($plan, $arguments[1]);
                    $this->success('create', $arguments[1]);
                    break;

                case 'validate-edit':
                    $this->requireCount($command, $arguments, 2);
                    $plan = PlanDecoder::editPlanFromFile($arguments[1]);
                    $this->editor->validateOperations($arguments[0], $plan);
                    $this->success('validate_edit');
                    break;

                case 'edit':
                    $this->requireCount($command, $arguments, 3);
                    $plan = PlanDecoder::editPlanFromFile($arguments[1]);
                    $this->editor->edit($arguments[0], $arguments[2], $plan);
                    $this->success('edit', $arguments[2]);
                    break;

                default:
                    throw new ContractException(sprintf('unknown command: %s', $command));
            }

            return 0;
        } catch (ValidationException $exception) {
            $this->failure($command, $exception->stableCode(), $exception->getMessage(), $exception->relatedPath(), $exception->errors());
            return 3;
        } catch (RetrievalException $exception) {
            $this->failure($command, 'retrieval_validation_failed', $exception->getMessage(), null, null, $exception->issue());
            return 3;
        } catch (CliInputException $exception) {
            $this->failure($command, $exception->stableCode(), $exception->getMessage());
            return 2;
        } catch (ContractException $exception) {
            $code = strpos($exception->getMessage(), 'failed to read plan') === 0 ? 'plan_read_failed' : 'invalid_plan_json';
            if (strpos($exception->getMessage(), 'unknown command') === 0 || strpos($exception->getMessage(), 'expects ') !== false) {
                $code = 'invalid_arguments';
            }
            $this->failure($command ?: 'cli', $code, $exception->getMessage());
            return 2;
        } catch (EngineException $exception) {
            $this->failure($command, $exception->stableCode(), $exception->getMessage(), $exception->relatedPath());
            return in_array($exception->stableCode(), ['unsupported_file_type'], true) ? 2 : 1;
        } catch (Throwable $exception) {
            $this->failure($command, 'internal_error', $exception->getMessage());
            return 1;
        }
    }

    /** @param list<string> $arguments */
    private function requireCount(string $command, array $arguments, int $count): void
    {
        if (count($arguments) !== $count) {
            throw new ContractException(sprintf('%s expects %d argument(s)', $command, $count));
        }
    }

    /** @param list<string> $arguments @return array{0: string, 1: string|null} */
    private function extractArguments(array $arguments): array
    {
        if (count($arguments) === 1) {
            return [$arguments[0], null];
        }
        if (count($arguments) === 3 && in_array($arguments[1], ['--output', '-o'], true)) {
            return [$arguments[0], $arguments[2]];
        }
        throw new ContractException('extract expects <input.xlsx> [--output <raw-workbook.json>]');
    }

    /** @param list<string> $arguments @return array<string, mixed> */
    private function retrieveArguments(array $arguments): array
    {
        if ($arguments === [] || substr($arguments[0], 0, 1) === '-') {
            throw new ContractException('retrieve expects <input.xlsx> and required source options');
        }
        $result = [
            'input' => array_shift($arguments),
            'source_id' => null,
            'source_revision' => null,
            'display_name' => null,
            'profile' => null,
            'output' => null,
        ];
        $map = [
            '--source-id' => 'source_id',
            '--source-revision' => 'source_revision',
            '--display-name' => 'display_name',
            '--profile' => 'profile',
            '--output' => 'output',
            '-o' => 'output',
        ];
        $seen = [];
        while ($arguments !== []) {
            $option = array_shift($arguments);
            if (!isset($map[$option]) || $arguments === []) {
                throw new ContractException("invalid retrieve option: $option");
            }
            $key = $map[$option];
            if (isset($seen[$key])) {
                throw new ContractException("retrieve option $option was supplied more than once");
            }
            $seen[$key] = true;
            $result[$key] = array_shift($arguments);
        }
        if ($result['source_id'] === null || $result['source_revision'] === null) {
            throw new ContractException('retrieve requires --source-id and --source-revision');
        }

        return $result;
    }

    /** @param list<string> $arguments @return array<string, mixed> */
    private function searchArguments(array $arguments): array
    {
        if (count($arguments) < 2) {
            throw new ContractException('search expects <retrieval.json> <query>');
        }
        $result = [
            'retrieval' => array_shift($arguments),
            'query' => array_shift($arguments),
            'limit' => 10,
            'source_id' => null,
            'source_revision' => null,
            'sheet' => null,
        ];
        $map = [
            '--limit' => 'limit',
            '--source-id' => 'source_id',
            '--source-revision' => 'source_revision',
            '--sheet' => 'sheet',
        ];
        $seen = [];
        while ($arguments !== []) {
            $option = array_shift($arguments);
            if (!isset($map[$option]) || $arguments === []) {
                throw new ContractException("invalid search option: $option");
            }
            $key = $map[$option];
            if (isset($seen[$key])) {
                throw new ContractException("search option $option was supplied more than once");
            }
            $seen[$key] = true;
            $value = array_shift($arguments);
            if ($key === 'limit') {
                if (!preg_match('/^[0-9]+$/D', $value)) {
                    throw new ContractException('search --limit must be an integer from 1 through 20');
                }
                $value = (int) $value;
                if ($value < 1 || $value > 20) {
                    throw new ContractException('search --limit must be an integer from 1 through 20');
                }
            }
            $result[$key] = $value;
        }

        return $result;
    }

    private function success(string $command, ?string $output = null): void
    {
        $document = ['schema_version' => '1.0', 'ok' => true, 'command' => $command];
        if ($output !== null) {
            $document['output'] = $output;
        }
        $this->writeJson($this->stdout, $document);
    }

    /** @param resource $stream @param mixed $document */
    private function writeJson($stream, $document): void
    {
        try {
            $json = json_encode(
                $document,
                JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_PRESERVE_ZERO_FRACTION | JSON_THROW_ON_ERROR
            );
        } catch (JsonException $exception) {
            throw new EngineException('json_serialization_failed', $exception->getMessage(), null, $exception);
        }
        if (fwrite($stream, $json . PHP_EOL) === false) {
            throw new EngineException('stdout_write_failed', 'failed to write JSON output');
        }
    }

    /** @param mixed $document */
    private function writeNewJson(string $path, $document): void
    {
        $handle = @fopen($path, 'xb');
        if ($handle === false) {
            $code = file_exists($path) ? 'output_exists' : 'output_write_failed';
            throw new EngineException($code, sprintf('failed to create JSON output: %s', $path), $path);
        }
        try {
            $json = json_encode(
                $document,
                JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES | JSON_UNESCAPED_UNICODE | JSON_PRESERVE_ZERO_FRACTION | JSON_THROW_ON_ERROR
            );
            if (fwrite($handle, $json . PHP_EOL) === false) {
                throw new EngineException('output_write_failed', sprintf('failed to write JSON output: %s', $path), $path);
            }
        } catch (Throwable $exception) {
            fclose($handle);
            @unlink($path);
            if ($exception instanceof EngineException) {
                throw $exception;
            }
            throw new EngineException('json_serialization_failed', $exception->getMessage(), $path, $exception);
        }
        fclose($handle);
    }

    /** @param list<array<string, mixed>>|null $validationErrors @param array<string, mixed>|null $retrievalIssue */
    private function failure(
        string $command,
        string $code,
        string $message,
        ?string $path = null,
        ?array $validationErrors = null,
        ?array $retrievalIssue = null
    ): void {
        $error = ['code' => $code, 'message' => $message];
        if ($path !== null) {
            $error['path'] = $path;
        }
        if ($validationErrors !== null && $validationErrors !== []) {
            $error['validation_errors'] = $validationErrors;
        }
        if ($retrievalIssue !== null) {
            $error['retrieval_issue'] = $retrievalIssue;
        }
        try {
            $this->writeJson($this->stderr, [
                'schema_version' => '1.0',
                'ok' => false,
                'command' => $command,
                'error' => $error,
            ]);
        } catch (Throwable $_) {
            // There is no safe fallback if stderr itself is unavailable.
        }
    }

    private function help(): string
    {
        return <<<'TEXT'
excel-tool 0.1.0

Usage:
  excel-tool inspect <input.xlsx>
  excel-tool extract <input.xlsx> [--output <raw-workbook.json>]
  excel-tool retrieve <input.xlsx> --source-id <ID> --source-revision <REVISION>
    [--display-name <NAME>] [--profile <profile.json>] [--output <retrieval.json>]
  excel-tool search <retrieval.json> <query> [--limit 10]
    [--source-id <ID>] [--source-revision <REVISION>] [--sheet <NAME>]
  excel-tool create <plan.json> <output.xlsx>
  excel-tool validate-edit <input.xlsx> <plan.json>
  excel-tool edit <input.xlsx> <plan.json> <output.xlsx>

Options:
  -h, --help       Show this help
  -V, --version    Show version

TEXT;
    }
}
