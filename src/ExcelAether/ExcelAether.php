<?php

namespace ExcelAether;

use ExcelAether\Connector\Excel;
use ExcelAether\Connector\PhpSpreadsheet;
use ExcelAether\Connector\Generate;
use Exception;

/**
 * Class ExcelAether
 * @package ExcelAether
 */
class ExcelAether
{
    /**
     * @throws Exception
     */
    public static function ExcelCreateBySpreadsheet(array $header, array $list, string $fileName, string $dir, string $title = ''): Excel
    {
        try {
            $generate = Generate::construct($header, $list, $fileName, $dir, $title);
            PhpSpreadsheet::createExcel($generate);
            return Excel::construct($generate->getDir(), $generate->getFileName());

        } catch (\PhpOffice\PhpSpreadsheet\Exception $exception) {
            throw new Exception($exception->getMessage());
        }
    }

    public static function ExcelCreateByHtml(array $header, array $list, string $fileName, string $dir, string $title = '')
    {

    }
}