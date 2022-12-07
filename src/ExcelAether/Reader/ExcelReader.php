<?php

namespace ExcelAether\Reader;

use Exception;
use Generator;

class ExcelReader
{
    private $inputFile = '';

    private $tmpPath = '';

    private $tmpFile = '';

    private $count = 0;

    public function setTmpPath(string $tmpPath)
    {
        $this->tmpPath = $tmpPath;
    }

    public function count(): int
    {
        return $this->count;
    }

    public function load(string $file)
    {
        $this->inputFile = $file;

        $this->tmpPath();

        $tmpName = $this->tmpPath . DIRECTORY_SEPARATOR . md5(uniqid()) . '.csv';

        if (!file_exists($this->inputFile)) {
            throw new Exception('文件不存在');
        }
        $this->tmpFile = $tmpName;
        $script = $this->scriptPath();

        exec($script);

        $this->getCount();

    }

    private function tmpPath()
    {
        if (!$this->inputFile) {
            throw new Exception('excel is null');
        }
        if (!$this->tmpPath) {
            $this->tmpPath = dirname($this->inputFile);
        } else {
            if (!is_dir($this->tmpPath) and !mkdir($this->tmpPath, 0755, true)) {
                throw new Exception('文件夹创建失败');
            }
        }
    }

    private function scriptPath(): string
    {
        $tmpName = $this->tmpFile;
        $reflect = new \ReflectionClass($this);
        $path = substr($reflect->getFileName(), 0, -15);

        $os = PHP_OS;

        if ($os == 'WINNT') {
            return "$path\ExcelReader.exe $this->inputFile $tmpName";
        } else {
            return "ExcelReader $this->inputFile $tmpName";
        }
    }

    private function getCount()
    {
        $fp = file($this->tmpFile, FILE_SKIP_EMPTY_LINES);
        $this->count = count($fp);
    }

    public function read(): Generator
    {
        return $this->readCsv($this->tmpFile);
    }

    private function readCsv($file): Generator
    {
        if (!is_file($file) && !file_exists($file)) {
            throw new Exception('文件错误');
        }

        $cvsFile = fopen($file, 'r');

        while ($filData = fgetcsv($cvsFile)) {
            yield $filData;
        }
        fclose($cvsFile);
        $this->tmpDelete();
    }

    public function tmpDelete()
    {
        unlink($this->tmpFile);
    }

    private function scriptPath2(): string
    {
        $reflect = new \ReflectionClass($this);
        $path = substr($reflect->getFileName(), 0, -15);

        $os = PHP_OS;

        if ($os == 'WINNT') {
            return $path . 'ExcelReader.py';
        } else {
            return 'ExcelReader.py';
        }
    }
}