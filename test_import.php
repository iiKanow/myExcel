<?php

require 'vendor/autoload.php';

use App\Controllers\ExcelController;

$controller = new ExcelController();
$result = $controller->import('test_import.xlsx');

if ($result['success']) {
    echo "导入成功！\n";
    echo "导入的数据：\n";
    print_r($result['data']);
} else {
    echo "导入失败：" . $result['message'] . "\n";
} 