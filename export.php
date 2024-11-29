<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// 示例数据
$data = [
    ['姓名', '年龄', '部门'],
    ['张三', 25, '技术部'],
    ['李四', 30, '市场部'],
    ['王五', 28, '财务部']
];

// 创建新的Excel文档
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// 写入数据
foreach ($data as $rowIndex => $row) {
    foreach ($row as $columnIndex => $value) {
        $sheet->setCellValueByColumnAndRow($columnIndex + 1, $rowIndex + 1, $value);
    }
}

// 设置响应头
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="示例数据.xlsx"');
header('Cache-Control: max-age=0');

// 输出Excel文件
$writer = new Xlsx($spreadsheet);
$writer->save('php://output'); 