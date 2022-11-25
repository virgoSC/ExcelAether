<?php

use ExcelAether\Reader\ExcelReader;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

require "vendor/autoload.php";

//生成excel
if (0) {
    $name = '1111.xls';

    $dir = './excel';

    $title = '';

    $header = [
        'name' => 'tmpName1',
        'phone' => 'tmpPhone1',
        'address' => 'tmpAddress1'
    ];

    $list = [
        [
            'name' => 'xixi',
            'phone' => '181',
            'address' => 'address1-2',
            'ttt' => '22'
        ],
        [
            'name' => 'xixi',
            'address' => 'address2-2',
            'phone' => '182'
        ],
        [
            'name' => 'xixi',
            'address' => 'address3-2',
            'phone' => '183'
        ]
    ];


//or

//    $header = [
//        'name1', 'phone1', 'address1'
//    ];

//    $list = [
//        ['xixi', '22233.02', '1856263'],
//        ['xixi', '22233', '1856263'],
//        ['xixi', '22233', '1856263'],
//        ['xixi', '22233', '1856263'],
//        ['xixi', '22233', '1856263'],
//    ];

    $config = [
        'width' => [5, 10, 10],
        'height' => '20',
        'headerStyle' => [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, //水平居中
                'vertical' => Alignment::VERTICAL_CENTER, //垂直居中
            ],
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
            'font' => [
                'bold' => true,
            ]]
        ,
        'cellStyle' => [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, //水平居中
                'vertical' => Alignment::VERTICAL_CENTER, //垂直居中
            ],
        ],
    ];

    $re = \ExcelAether\ExcelAether::ExcelCreateBySpreadsheet($header, $list, $name, $dir, $title, $config);

    var_dump($re->getDir() . $re->getFileName());

}

//读取excel
if (1) {
    $inputFile = './excel/t1.xlsx';

    $reader = new ExcelReader();

    $reader->load($inputFile);

    foreach ($reader->read() as $v) {
        var_dump($v[0]);
    }
}


