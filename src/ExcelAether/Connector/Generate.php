<?php

namespace ExcelAether\Connector;

class Generate
{

    /**
     * @var array $header
     */
    private $header;

    /**
     * @var array $headerKey
     */
    private $headerKey;

    /**
     * @var array $list
     */
    private $list;

    /**
     * @var string $fileName
     */
    private $fileName;

    /**
     * @var string $title
     */
    private $title;

    /**
     * @var string $dir
     */
    private $dir;

    public function __construct(array $header, array $list, string $fileName, string $dir, string $title = '')
    {
        $this->setHeader($header);
        $this->setList($list);
        $this->setFileName($fileName);
        $this->setDir($dir);
        $this->setTitle($title);
        $this->parse();
    }

    public static function construct($header, $list, $fileName, $dir, $title = ''): Generate
    {
        return (new self($header, $list, $fileName, $dir, $title));
    }

    /**
     * @return array
     */
    public function getHeader(): array
    {
        return $this->header;
    }

    /**
     * @param array $header
     * @return Generate
     */
    public function setHeader(array $header): Generate
    {
        $this->header = $header;
        return $this;
    }

    /**
     * @return array
     */
    public function getList(): array
    {
        return $this->list;
    }

    /**
     * @param array $list
     * @return Generate
     */
    public function setList(array $list): Generate
    {
        $this->list = $list;
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
     * @return Generate
     */
    public function setFileName(string $fileName): Generate
    {
        $this->fileName = rtrim($fileName, '.xls') . '.xls';


        return $this;
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
     * @return Generate
     */
    public function setDir(string $dir): Generate
    {
        $this->dir = rtrim($dir, '/') . '/';
        return $this;
    }

    /**
     * @return string
     */
    public function getTitle(): string
    {
        return $this->title;
    }

    /**
     * @param string $title
     * @return Generate
     */
    public function setTitle(string $title): Generate
    {
        $this->title = $title;
        return $this;
    }

    private function parse()
    {
        $columns = array_keys($this->header);

        //索引数组和关联数据区分
        if ($columns !== range(0, count($this->header) - 1)) {
            //关联数组
            $tmp = array_combine($columns, array_fill(0, count($columns), ''));
            //字段
            foreach ($this->list as &$value) {

                $value = array_merge($tmp,array_intersect_key($value, $tmp));
            }
            $this->headerKey = $columns;

            $this->header = array_values($this->header);

        } else {
            //索引数组
            $this->headerKey = $this->header;
        }
    }
}
