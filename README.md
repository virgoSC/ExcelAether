# ExcelAether

简单的excel 生成器 (偷懒用)

安装

```shell
composer require virgo/excel-aether
```

## 写入
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
$list = [       //key=>value 格式可以打乱
    [
        'name' => 'zhang',
        'address' => 'beijing',
        'phone' => '1821',
        'tt' => '333' //与表头字段不同，会自动过滤
    ],
    [
        'name' => 'wen',
        'address' => 'shanghia',
        'phone' => '1822'
    ],
    [
        'name' => 'liu',
        'phone' => '1823'
        'address' => 'chengdu',
    ]
];

//或者

//表头
//-------------------
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

//------------------

//样式配置
$config = [
        //竖列 宽度
        'width' => [5, 10, 10],
        //高度
        'height' => '20',
        //表头样式
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
        //非表头高度
        'cellStyle' => [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER, //水平居中
                'vertical' => Alignment::VERTICAL_CENTER, //垂直居中
            ],
        ],
    ];

    $re = \ExcelAether\ExcelAether::ExcelCreateBySpreadsheet($header, $list, $name, $dir, $title, $config);

    var_dump($re->getDir() . $re->getFileName());

```

## 读取
PS：弃用 phpspreadsheet 会造成内存占用过大，等到内存回收，导致程序错误

使用python 将xls xlxs 转成csv 再用PHP 读取


````shell
#添加到 /etc/profile
export PATH=$PATH:/var/www/swoft/vendor/virgo/excel-aether/src/ExcelAether/Reader

#执行
source /etc/profile
````

### 读取demo

```php
$inputFile = './excel/t1.xlsx';
$reader = new ExcelReader();
$reader->load($inputFile);
foreach ($reader->read() as $v) {
    var_dump($v[0]);
}

```

###### 更新
###### 1 自动读取csv 为utf8编码 






