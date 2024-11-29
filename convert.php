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

    // 4. 复制数据（从第3行开始）
    $templateData = $templateSheet->toArray();
    // 根据模板数据，动态设置表头和列宽
    $allHeaders = $templateData[1];
    // 数组取交集，得到需要输出的字段
    $headers = array_intersect($allHeaders, ['序号', '姓名', '职务（岗位）工资', '级别（技术等级、薪级）工资', '地区补贴', '保留补贴', '其他', '基本工资合计', '扣养老保险', '扣职业年金', '扣住房公积金', '扣医疗保险', '扣大额医疗', '扣个人所得税', '基本工资实发合计', '基础性绩效', '取暖费', '实际执行工资', '实发合计', '补发工资', '降温费']);
    // 总列数
    $count = count($headers);
    // 确保列数不超过 26
    if ($count > 26) {
        throw new Exception('输出字段数量超过最大列数限制');
    }
    // 把序号转化为字母
    $stringMap = range('A', 'Z'); // 使用 range 函数生成字母数组
    $maxCol = $stringMap[$count - 1];

    // 设置标题
    // 月份根据模板工资月份列获取
    $monthStr = $templateSheet->getCell('C3')->getValueString();
    $month = explode('-', $monthStr)[1];
    $sheet2->setCellValue('A1', $month . '月份工资明细及汇总表');
    $sheet2->mergeCells('A1:' . $maxCol . '1');

    // 设置表头
    $widthMap = [
        '序号' => 6,
        '姓名' => 8,
        '职务（岗位）工资' => 10,
        '级别（技术等级、薪级）工资' => 10,
        '地区补贴' => 6,
        '保留补贴' => 6,
        '其他' => 8,
        '基本工资合计' => 10,
        '扣养老保险' => 10,
        '扣职业年金' => 8,
        '扣住房公积金' => 8,
        '扣医疗保险' => 8,
        '扣大额医疗' => 6,
        '扣个人所得税' => 6,
        '基本工资实发合计' => 12,
        '基础性绩效' => 8,
        '取暖费' => 8,
        '实际执行工资' => 10,
        '实发合计' => 10,
        '补发工资' => 4,
        '降温费' => 8,
    ];
    $column = 'A';
    foreach ($headers as $col => $value) {
        // 设置列宽，确保内容完整显示
        $width = $widthMap[$value];
        $width = isset($widthMap[$value]) ? $widthMap[$value] : 10;  // 默认宽度为 10
        $sheet2->getColumnDimension($column)->setWidth($width);
        // 设置表头
        $sheet2->setCellValue($column . '2', $value);
        $column++;
    }
    $dataStartRow = 3;
    $other = $basicSalary = $warmOrCoolingSalary = 0;
    $total = count($templateData);
    // 翻转index和字段名对应
    $columnNameIndexMap = array_flip($headers);
    foreach (array_slice($templateData, 2) as $rowIndex => $row) {
        $column = 'A';
        // 根据headers循环输出
        foreach ($headers as $key => $value) {
            // 检查行和列是否有效
            $val = isset($row[$key]) && !empty($row[$key]) ? $row[$key] : '';
            $sheet2->setCellValue($column . ($rowIndex + 3), $val);
            $column++;
        }
        if ($rowIndex + 3 == $total) {
            $other = $row[$columnNameIndexMap['其他']];
            $basicSalary = $row[$columnNameIndexMap['基本工资合计']];
            $warmOrCoolingSalary = (isset($columnNameIndexMap['取暖费']) && isset($row[$columnNameIndexMap['取暖费']]) && $row[$columnNameIndexMap['取暖费']] ? $row[$columnNameIndexMap['取暖费']] : 0) + (isset($columnNameIndexMap['降温费']) && isset($row[$columnNameIndexMap['降温费']]) && $row[$columnNameIndexMap['降温费']] ? $row[$columnNameIndexMap['降温费']] : 0);
        }
    }

    // 5. 设置整体样式
    $lastRow = $sheet2->getHighestRow();
    if ($lastRow < 1) {
        throw new Exception('没有有效的数据行');
    }
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
    $sheet2->getStyle('A1:' . $maxCol . $lastRow)->applyFromArray($styleArray);

    // 6. 设置标题样式
    $sheet2->getStyle('A1')->applyFromArray([
        'font' => [
            'name' => '宋体',
            'size' => 14,
            'bold' => true
        ]
    ]);

    // 7. 设置表头样式
    $sheet2->getStyle('A2:' . $maxCol . '2')->applyFromArray([
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
    $sheet2->getStyle('C3:' . $maxCol . $lastRow)->getNumberFormat()->setFormatCode($numberFormat);

    // 9. 设置行高
    $sheet2->getRowDimension(1)->setRowHeight(25);
    $sheet2->getRowDimension(2)->setRowHeight(30);
    for ($i = 3; $i <= $lastRow; $i++) {
        $sheet2->getRowDimension($i)->setRowHeight(23);
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

    $outputFile = $outputDir . '/' . $month . '月份工资明细表_' . date('Ymd') . '.xlsx';
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($outputFile);

    // 处理工资指标分配明细表
    // 读取 Excel 文件
    $inputFileName = '/Users/surge/Desktop/防汛办12月份工资指标分配明细表.xls'; // 输入文件名
    $spreadsheet = IOFactory::load($inputFileName);
    $sheet = $spreadsheet->getActiveSheet();

    // 遍历行并修改特定字段
    foreach ($sheet->getRowIterator() as $row) {
        $cellIterator = $row->getCellIterator();
        $cellIterator->setIterateOnlyExistingCells(false); // 允许遍历所有单元格

        foreach ($cellIterator as $cell) {
            if ($cell->getColumn() == 'H') {
                if ($row->getRowIndex() == 1) {
                    $cell->setValue($month . '月份工资明细');
                } elseif ($row->getRowIndex() == 2) {
                    $cell->setValue($other + $basicSalary + $warmOrCoolingSalary + 83131.64);
                } elseif ($row->getRowIndex() == 4) {
                    $cell->setValue($basicSalary);
                } elseif ($row->getRowIndex() == 6) {
                    $cell->setValue($warmOrCoolingSalary);
                } elseif ($row->getRowIndex() == 10) {
                    $cell->setValue($other);
                }
            }
        }
    }

    // 保存修改后的文件
    $outputFile2 = $outputDir . '/' . $month . '月份工资指标分配明细表_' . date('Ymd') . '.xlsx';
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($outputFile2);

    echo json_encode([
        'success' => true,
        'message' => '转换成功',
        'file_path' => $outputFile,
        'file_path2' => $outputFile2
    ]);

} catch (Exception $e) {
    echo json_encode([
        'success' => false,
        'message' => '处理失败：' . $e->getMessage()
    ]);
}