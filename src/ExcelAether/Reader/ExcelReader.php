<?php

namespace ExcelAether\Reader;

use Generator;

class ExcelReader
{
    private $inputFile = '';

    private $tmpPath = '';

    public function load(string $file)
    {
        $this->inputFile = $file;
    }

    public function read(): Generator
    {
        $this->tmpPath();

        $tmpName = $this->tmpPath . '/' . md5(uniqid()) . '.csv';

        $script = $this->scriptPath();

        $com = "$script $this->inputFile $tmpName";

        exec($com);

        return $this->readCsv($tmpName);
    }

    private function scriptPath(): string
    {
        $reflect = new \ReflectionClass($this);
        $path = substr($reflect->getFileName(), 0, -15);

        $os = PHP_OS;

        if ($os == 'WINNT') {
            return $path.'reader.exe';
        } else {
            return './'.$path . 'main';
        }
    }

    private function tmpPath()
    {
        if (!$this->inputFile) {
            die('excel is null');
        }

        $this->tmpPath = dirname($this->inputFile);
    }

    private function readCsv($file): Generator
    {
        if (!is_file($file) && !file_exists($file)) {
            die('文件错误');
        }

        $cvsFile = fopen($file, 'r');

        while ($filData = fgetcsv($cvsFile)) {
            yield $filData;
        }
        fclose($cvsFile);
        unlink($file);
    }
}