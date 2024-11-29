<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;

header('Access-Control-Allow-Origin: *');
header('Content-Type: application/json');

try {
    if (!isset($_FILES['excel_file'])) {
        throw new Exception('请选择文件');
    }

    $file = $_FILES['excel_file'];
    if ($file['error'] !== UPLOAD_ERR_OK) {
        throw new Exception('文件上传失败');
    }

    $spreadsheet = IOFactory::load($file['tmp_name']);

    // 获取模板sheet
    $templateSheet = $spreadsheet->getSheetByName('模板');
    if (!$templateSheet) {
        throw new Exception('未找到"模板"工作表');
    }

    // 创建或获取sheet2
    if ($spreadsheet->sheetNameExists('sheet2')) {
        $sheet2 = $spreadsheet->getSheet(1);
    } else {
        $sheet2 = $spreadsheet->createSheet();
        $sheet2->setTitle('sheet2');
    }

    // 1. 设置列宽
    $columnWidths = [
        'A' => 4,     // 序号
        'B' => 8,     // 姓名
        'C' => 12,    // 职务(岗位)工资
        'D' => 12,    // 级别(技术等级、薪级)工资
        'E' => 8,     // 地区补贴
        'F' => 8,     // 保留补贴
        'G' => 8,     // 其他
        'H' => 10,    // 基本工资合计
        'I' => 8,     // 扣养老保险
        'J' => 8,     // 扣职业年金
        'K' => 8,     // 扣住房公积金
        'L' => 8,     // 扣医疗保险
        'M' => 8,     // 扣大额医疗
        'N' => 8,     // 扣个人所得税
        '0' => 10,    // 基本工资实发合计
        'P' => 8,     // 基础性绩效
        'Q' => 8,     // 取暖费
        'R' => 8,     // 实际执行工资
        'S' => 10     // 实发合计
    ];

    foreach ($columnWidths as $col => $width) {
        $sheet2->getColumnDimension($col)->setWidth($width);
    }

    // 2. 设置标题
    $month = date('n');
    $sheet2->setCellValue('A1', $month . '月份工资明细及汇总表');
    $sheet2->mergeCells('A1:U1');

    // 3. 设置表头
    $headers = [
        'A2' => '序号',
        'B2' => '姓名',
        'C2' => '职务(岗位)工资',
        'D2' => '级别(技术等级、薪级)工资',
        'E2' => '地区补贴',
        'F2' => '保留补贴',
        'G2' => '其他',
        'I2' => '基本工资合计',
        'J2' => '扣养老保险',
        'K2' => '扣职业年金',
        'L2' => '扣住房公积金',
        'M2' => '扣医疗保险',
        'N2' => '扣大额医疗',
        'O2' => '扣个人所得税',
        'Q2' => '基本工资实发合计',
        'R2' => '基础性绩效',
        'S2' => '取暖费',
        'T2' => '实际执行工资',
        'U2' => '实发合计'
    ];

    foreach ($headers as $cell => $value) {
        $sheet2->setCellValue($cell, $value);
    }

    // 4. 复制数据（从第3行开始）
    $templateData = $templateSheet->toArray();
    $dataStartRow = 3;
    $rowNumber = 1;

    foreach (array_slice($templateData, 1) as $rowIndex => $row) {
        // 跳过第一行（工资性质等信息）
        if ($rowIndex === 0) {
            continue;
        }

        // 检查是否为空行
        if (empty($row[6]) && empty($row[7]) && empty($row[8])) { // 检查姓名和相关字段
            continue;
        }

        $currentRow = $dataStartRow + $rowNumber - 1;

        // 设置序号
        $sheet2->setCellValue('A' . $currentRow, $rowNumber);

        // 映射数据到对应的列（根据实际数据位置更新）
        $columnMapping = [
            'B' => 6,  // 姓名 (第7列)
            'C' => 15, // 职务(岗位)工资 (第14列)
            'D' => 16, // 级别工资 (第15列)
            'E' => 17, // 地区补贴 (第17列)
            'F' => 18, // 保留补贴 (第18列)
            'G' => 19, // 其他 (第19列)
            'I' => 21, // 基本工资合计
            'J' => 22, // 扣养老保险
            'K' => 23, // 扣职业年金
            'L' => 23, // 扣住房公积金
            'M' => 24, // 扣医疗保险
            'N' => 25, // 扣大额医疗
            'O' => 26, // 扣个人所得税
            'Q' => 28, // 基本工资实发合计
            'R' => 29, // 基础性绩效
            'S' => 31, // 取暖费
            'T' => 33, // 实际执行工资
            'U' => 34  // 实发合计
        ];

        foreach ($columnMapping as $targetCol => $sourceIndex) {
            if (isset($row[$sourceIndex])) {
                $value = $row[$sourceIndex];
                // 如果是数字，确保格式正确
                if (is_numeric($value)) {
                    $value = floatval($value);
                }
                $sheet2->setCellValue($targetCol . $currentRow, $value);
            }
        }

        $rowNumber++;
    }

    // 5. 设置整体样式
    $lastRow = $sheet2->getHighestRow();
    $styleArray = [
        'font' => [
            'name' => '宋体',
            'size' => 9
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical' => Alignment::VERTICAL_CENTER,
            'wrapText' => true
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['rgb' => '000000']
            ]
        ]
    ];
    $sheet2->getStyle('A1:U' . $lastRow)->applyFromArray($styleArray);

    // 6. 设置标题样式
    $sheet2->getStyle('A1')->applyFromArray([
        'font' => [
            'name' => '宋体',
            'size' => 14,
            'bold' => true
        ]
    ]);

    // 7. 设置表头样式
    $sheet2->getStyle('A2:U2')->applyFromArray([
        'font' => [
            'name' => '宋体',
            'size' => 9,
            'bold' => true
        ],
        'fill' => [
            'fillType' => Fill::FILL_SOLID,
            'color' => ['rgb' => 'D9D9D9']
        ]
    ]);

    // 8. 设置数字格式
    $numberFormat = '#,##0.00';
    $sheet2->getStyle('C3:U' . $lastRow)->getNumberFormat()->setFormatCode($numberFormat);

    // 9. 设置行高
    $sheet2->getRowDimension(1)->setRowHeight(30);
    $sheet2->getRowDimension(2)->setRowHeight(40);
    for ($i = 3; $i <= $lastRow; $i++) {
        $sheet2->getRowDimension($i)->setRowHeight(25);
    }

    // 10. 设置打印格式
    $sheet2->getPageSetup()
        ->setOrientation(PageSetup::ORIENTATION_LANDSCAPE)
        ->setPaperSize(PageSetup::PAPERSIZE_A4)
        ->setFitToPage(true)
        ->setFitToWidth(1)
        ->setFitToHeight(0);

    // 11. 设置页边距（厘米）
    $sheet2->getPageMargins()
        ->setTop(1)
        ->setRight(0.5)
        ->setBottom(1)
        ->setLeft(0.5)
        ->setHeader(0.5)
        ->setFooter(0.5);

    // 12. 设置打印标题重复
    $sheet2->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 2);

    // 保存文件
    $outputDir = 'output';
    if (!file_exists($outputDir)) {
        mkdir($outputDir, 0777, true);
    }

    $outputFile = $outputDir . '/' . $month . '月份工资明细表_' . date('YmdHis') . '.xlsx';
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($outputFile);

    echo json_encode([
        'success' => true,
        'message' => '转换成功',
        'file_path' => $outputFile
    ]);

} catch (Exception $e) {
    echo json_encode([
        'success' => false,
        'message' => '处理失败：' . $e->getMessage()
    ]);
}