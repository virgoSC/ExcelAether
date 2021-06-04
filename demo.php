<?php

require "vendor/autoload.php";


$name = '1111.xls';

$dir = './excel';

$title = '';

$header = [
    'name' => 'name1',
    'phone' => 'phone1',
    'address' => 'address1'
];

$list = [
    [
        'name' => 'xixi',
        'address' => '3323',
        'phone' => '122222'
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
