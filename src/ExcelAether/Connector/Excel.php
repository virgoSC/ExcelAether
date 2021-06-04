<?php

namespace ExcelAether\Connector;

/**
 * Class Excel
 * @package ExcelAether\Connector
 */
class Excel
{

    /**
     * @var string $dir
     */
    private $dir;

    /**
     * @var string $fileName
     */
    private $fileName;

    /**
     * Excel constructor.
     */
    public function __construct(string $dir, string $fileName)
    {
        $this->dir = $dir;
        $this->fileName = $fileName;
    }

    public static function construct(string $dir, string $fileName): Excel
    {
        return (new self($dir, $fileName));
    }

    /**
     * @return string
     */
    public function getDir(): string
    {
        return $this->dir;
    }

    /**
     * @param string $dir
     * @return Excel
     */
    public function setDir(string $dir): Excel
    {
        $this->dir = $dir;
        return $this;
    }

    /**
     * @return string
     */
    public function getFileName(): string
    {
        return $this->fileName;
    }

    /**
     * @param string $fileName
     * @return Excel
     */
    public function setFileName(string $fileName): Excel
    {
        $this->fileName = $fileName;
        return $this;
    }


}
