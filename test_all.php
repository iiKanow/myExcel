<?php

require 'vendor/autoload.php';

use App\Controllers\ExcelController;

echo "=== 开始测试 ===\n\n";

// 1. 测试生成Excel文件
echo "1. 生成测试Excel文件...\n";
require 'create_test_excel.php';
echo "\n";

// 2. 测试导入功能
echo "2. 测试导入功能...\n";
$controller = new ExcelController();
$result = $controller->import('test_import.xlsx');

if ($result['success']) {
    echo "导入成功！\n";
    echo "导入的数据：\n";
    print_r($result['data']);
} else {
    echo "导入失败：" . $result['message'] . "\n";
}
echo "\n";

// 3. 测试导出功能
echo "3. 测试导出功能...\n";
echo "注意：导出功能需要在Web环境中测试，因为它会设置HTTP头信息。\n\n";

echo "=== 测试完成 ===\n";