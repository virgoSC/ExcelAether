<?php


namespace ExcelAether\Connector;

use PhpOffice\PhpSpreadsheet\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\IOFactory;

class PhpSpreadsheet
{

    /**
     * @throws Exception
     */
    public static function createExcel(Generate $generate)
    {

        $spreadsheet = new Spreadsheet();

        $worksheet = $spreadsheet->getActiveSheet();
        $worksheet->getRowDimension('1')->setRowHeight(20);

        $cellCode = 'A1';

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
            $worksheet->getStyle($cellCode)->applyFromArray($styleArray);
            $worksheet->getDefaultColumnDimension()->setWidth(15);
            $horizontal += 1;
        }
        $horizontal = 0;

        $worksheet->getStyle(self::IntToChr($horizontal) . $vertical . ":$cellCode")->applyFromArray($styleBorders);

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

                $worksheet->getStyle($cellCode)->applyFromArray($styleArray);

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

    public static function IntToChr(int $index, int $start = 65): string
    {
        $str = '';
        if (floor($index / 26) > 0) {
            $str .= self::IntToChr(floor($index / 26) - 1);
        }
        return $str . chr($index % 26 + $start);
    }
}