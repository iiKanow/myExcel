<?php

use App\Controllers\ExcelController;  // 添加 use 语句

// 启动内置Web服务器的路由脚本
$uri = parse_url($_SERVER['REQUEST_URI'], PHP_URL_PATH);

if ($uri === '/excel/export') {
    require 'vendor/autoload.php';
    $controller = new ExcelController();  // 移除完整命名空间前缀
    $controller->export();
    exit;
}

if ($uri === '/excel/import') {
    require 'vendor/autoload.php';
    $controller = new ExcelController();  // 移除完整命名空间前缀
    $result = $controller->import($_FILES['excel_file']['tmp_name']);
    header('Content-Type: application/json');
    echo json_encode($result);
    exit;
}

// 默认返回HTML页面
if ($uri === '/' || $uri === '/index.html') {
    include 'upload.html';
    exit;
}

http_response_code(404);
echo "404 Not Found"; 