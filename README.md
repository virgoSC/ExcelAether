# ExcelAether
简单的excel 生成器 (偷懒用)

安装
```shell
composer require virgo/excel-aether
```

使用说明

```php
//文件名称
$name = 'demo.xls';
//or
$name = 'demo';

//地址
$dir = './excel';
//or
$dir = './excel/';

//标题
$title = 'demo';

//表头
$header = [
    'name' => '姓名',
    'phone' => '电话',
    'address' => '地址'
];

//数据格式
$list = [
    ['name' => 'zhang',
     'address' => 'beijing',
     'phone' => '1821'],
    ['name' => 'wen',
     'address' => 'shanghia',
     'phone' => '1822'],
    ['name' => 'liu',
     'address' => 'chengdu',
     'phone' => '1823']

];

//或者

//表头
$header = [
    '姓名', '电话', '地址'
];

//数据格式
$list = [
    ['zhang', '1825', 'beijing'],
    ['wen', '1823', 'shengzhen'],
    ['liu', '1822', 'shanghai'],
    ['xu', '1827', 'nancong'],
    ['qian', '1821', 'dongjing'],
];

//使用
$re = \ExcelAether\ExcelAether::ExcelCreateBySpreadsheet($header, $list, $name, $dir, $title);

var_dump($re->getDir().$re->getFileName());

```









