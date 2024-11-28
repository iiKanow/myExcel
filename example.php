<?php

require 'vendor/autoload.php';
use App\Controllers\ExcelController;  // 添加 use 语句

// 导出Excel
$excelController = new ExcelController();  // 移除完整命名空间前缀
$excelController->export();

// 导入Excel
$file = $_FILES['excel_file']['tmp_name'];
$result = $excelController->import($file);

if ($result['success']) {
    echo "导入成功！";
    print_r($result['data']);
} else {
    echo "导入失败：" . $result['message'];
} 