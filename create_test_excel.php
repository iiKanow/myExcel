<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// 创建新的Spreadsheet对象
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// 设置表头
$sheet->setCellValue('A1', '姓名');
$sheet->setCellValue('B1', '年龄');
$sheet->setCellValue('C1', '邮箱');

// 准备测试数据
$testData = [
    ['张三', 25, 'zhangsan@example.com'],
    ['李四', 30, 'lisi@example.com'],
    ['王五', 28, 'wangwu@example.com'],
    ['赵六', 35, 'zhaoliu@example.com'],
    ['钱七', 27, 'qianqi@example.com'],
];

// 填充数据
$row = 2;
foreach ($testData as $item) {
    $sheet->setCellValue('A' . $row, $item[0]);
    $sheet->setCellValue('B' . $row, $item[1]);
    $sheet->setCellValue('C' . $row, $item[2]);
    $row++;
}

// 设置列宽
$sheet->getColumnDimension('A')->setWidth(15);
$sheet->getColumnDimension('B')->setWidth(10);
$sheet->getColumnDimension('C')->setWidth(25);

// 设置表头样式
$sheet->getStyle('A1:C1')->getFont()->setBold(true);
$sheet->getStyle('A1:C1')->getFill()
    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
    ->getStartColor()->setARGB('FFCCCCCC');

// 创建写入对象
$writer = new Xlsx($spreadsheet);

// 保存文件
$writer->save('test_import.xlsx');

echo "测试文件 test_import.xlsx 已创建成功！\n"; 