<?php


namespace ExcelAether\CsvToExcel;

class CsvToExcel
{
    private $inputFile;

    private $outFile;

    public function csv2Excel($inputFile, $outFile)
    {
        $this->inputFile = $inputFile;

        $this->outFile = $outFile;

        $script = $this->scriptPath();

        exec($script);
    }

    private function scriptPath(): string
    {
        $reflect = new \ReflectionClass($this);
        $path = substr($reflect->getFileName(), 0, -15);

        $os = PHP_OS;

        if ($os == 'WINNT') {
            return "$path\CsvToExcel.exe $this->inputFile $this->outFile";
        } else {
            return "CsvToExcel '$this->inputFile' '$this->outFile'";
        }
    }
}