<?php

declare(strict_types=1);

$root = dirname(__DIR__);
$paths = [$root . '/src', $root . '/tests', $root . '/bin'];
$files = [];
foreach ($paths as $path) {
    $iterator = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($path));
    foreach ($iterator as $file) {
        if (!$file->isFile()) {
            continue;
        }
        if ($file->getExtension() === 'php' || $file->getFilename() === 'excel-tool') {
            $files[] = $file->getPathname();
        }
    }
}
sort($files);

$failed = false;
foreach ($files as $file) {
    $command = escapeshellarg(PHP_BINARY) . ' -n -l ' . escapeshellarg($file);
    $output = [];
    $status = 0;
    exec($command, $output, $status);
    if ($status !== 0) {
        $failed = true;
        fwrite(STDERR, implode(PHP_EOL, $output) . PHP_EOL);
    }
}
if ($failed) {
    exit(1);
}

fwrite(STDOUT, sprintf("Syntax OK (%d files)\n", count($files)));

