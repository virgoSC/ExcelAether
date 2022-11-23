<?php


namespace ExcelAether\Connector;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\IOFactory;

/**
 * Class PhpSpreadsheet
 * @package ExcelAether\Connector
 */
class PhpSpreadsheet
{

    /**
     * @throws Exception
     */
    public static function createExcel(Generate $generate)
    {
        $spreadsheet = new Spreadsheet();

        $config = $generate->getConfig();

        $headerStyle = $config['headerStyle'] ?? [];
        $cellStyle = $config['cellStyle'] ?? [];


        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->getRowDimension('1')->setRowHeight($config['height'] ?? 20);

        $worksheet->setTitle('sheet1');
        //全局样式
        $styleArray = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, //水平居中
                'vertical' => Alignment::VERTICAL_CENTER, //垂直居中
            ],
        ];
        $styleBorders = [
            'borders' => [  //黑框
                'outline' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => '000000'],
                ],
                'inside' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => '000000'],
                ],
            ],
        ];

        //横坐标
        $horizontal = 0;
        //纵坐标
        $vertical = 1;

        //宽度设置
        if ($configWidth = $config['width'] ?? []) {
            foreach ($configWidth as $l => $w) {
                $worksheet->getColumnDimension(self::IntToChr($l + 1))->setWidth($w);
            }
        }

        //设置标题
        if ($generate->getTitle()) {
            $cellCode = self::IntToChr($horizontal) . $vertical;
            $worksheet->setCellValue($cellCode, $generate->getTitle());
            $worksheet->getStyle($cellCode)->applyFromArray($styleArray);
            $worksheet->mergeCells(self::IntToChr($horizontal) . $vertical . ":" . self::IntToChr(count($generate->getHeader()) - 1) . $vertical);
            $vertical += 1;
        }

        //设置表头
        foreach ($generate->getHeader() as $value) {
            $cellCode = self::IntToChr($horizontal) . $vertical;
            $worksheet->setCellValue($cellCode, $value);
            if ($headerStyle) {
                $worksheet->getStyle($cellCode)->applyFromArray($headerStyle);
            } else {
                $worksheet->getStyle($cellCode)->applyFromArray($styleArray)->applyFromArray($styleBorders);
            }

            if (!$configWidth) {
                $worksheet->getDefaultColumnDimension()->setWidth(15);
            }
            $horizontal += 1;
        }

        $vertical += 1;
        //设置数据
        foreach ($generate->getList() as $datum) {
            $horizontal = 0;

            foreach ($datum as $value) {

                $cellCode = self::IntToChr($horizontal) . $vertical;

                if (is_int($value) or is_numeric($value) or is_float($value)) {
                    $value = $value . "\t";
                }

                $worksheet->setCellValue($cellCode, $value);

                if ($cellStyle) {
                    $worksheet->getStyle($cellCode)->applyFromArray($cellStyle);
                } else {
                    $worksheet->getStyle($cellCode)->applyFromArray($styleArray);
                }

                $horizontal += 1;
            }
            $vertical += 1;
        }
        //全局样式

        //文件生成
        if (!is_dir($generate->getDir()) and !mkdir($generate->getDir(), 0755, true)) {
            return false;
        }

        IOFactory::createWriter($spreadsheet, 'Xls')->save($generate->getDir() . $generate->getFileName());

        return $generate->getFileName();
    }


    /**
     * @param $fileName
     * @param string $sheet
     * @param int $maxColumn 行数
     * @param int $maxRow 列数
     * @param int $beginColumn
     * @param int $beginRow
     * @return array|void
     * @throws Exception
     * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
     */
    public static function readerExcel($fileName, string $sheet = '', int $maxColumn = 0, int $maxRow = 0, int $beginColumn = 0, int $beginRow = 1)
    {

        $fileType = IOFactory::identify($fileName);

        $reader = IOFactory::createReader($fileType);

        $canRead = $reader->canRead($fileName);

        if (!$canRead) {
            return [];
        }

        @$spreadsheet = $reader->load($fileName);

        if ($sheet) {
            $activeSheet = $spreadsheet->setActiveSheetIndexByName($sheet);
        } else {
            $activeSheet = $spreadsheet->getActiveSheet();
        }

        if ($maxColumn) {
            $column = $maxColumn;
        } else {
            $column = self::chrToInt($activeSheet->getHighestColumn());
        }

        if ($maxRow) {
            $row = $maxRow;
        } else {
            $row = $activeSheet->getHighestRow();
        }

        $data = [];
        for ($i = $beginRow; $i <= $row; $i++) {
            $columnData = [];
            for ($j = $beginColumn; $j < $column; $j++) {
                $columnData[] = $activeSheet->getCellByColumnAndRow($j, $i)->getValue();
            }
            $data[] = $columnData;
        }

        return $data;
    }

    public static function IntToChr($index, $start = 65): string
    {
        $str = '';
        if (floor($index / 26) > 0) {
            $str .= self::IntToChr(floor($index / 26) - 1);
        }
        return $str . chr($index % 26 + $start);
    }

    private static function chrToInt($chr): int
    {
        $int = 0;
        $length = strlen($chr);
        for ($i = 0; $i < $length; $i++) {
            $t = ord($chr[$i]) - 65 + 1;
            $int += $t;
        }

        return $int;
    }
}