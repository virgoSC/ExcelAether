<?php

require "vendor/autoload.php";

//生成excel
if (1) {
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
            'address' => '3323',
            'phone' => '122222',
            'ttt' => '22'
        ],
        [
            'name' => 'xixi',
            'address' => '3323',
            'phone' => '122222'
        ],
        [
            'name' => 'xixi',
            'address' => '3323',
            'phone' => '122222'
        ]
    ];


//or

    $header = [
        'name1', 'phone1', 'address1'
    ];

    $list = [
        ['xixi', '22233.02', '1856263'],
        ['xixi', '22233', '1856263'],
        ['xixi', '22233', '1856263'],
        ['xixi', '22233', '1856263'],
        ['xixi', '22233', '1856263'],
    ];

    $re = \ExcelAether\ExcelAether::ExcelCreateBySpreadsheet($header, $list, $name, $dir, $title);

    var_dump($re->getDir().$re->getFileName());

}


//读取excel
if (0) {
    $file = './1.xlsx';

    $re = \ExcelAether\ExcelAether::excelReadBySpreadsheet($file,'',0,0,0,2);

    file_put_contents('./1.txt',json_encode($re,JSON_UNESCAPED_UNICODE));

//    var_dump($re);

}