<?php

namespace App\Controllers;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

class ExcelController
{
    /**
     * 导出Excel文件
     */
    public function export()
    {
        // 创建新的Spreadsheet对象
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // 设置表头
        $sheet->setCellValue('A1', '姓名');
        $sheet->setCellValue('B1', '年龄');
        $sheet->setCellValue('C1', '邮箱');

        // 示例数据
        $data = [
            ['张三', 25, 'zhangsan@example.com'],
            ['李四', 30, 'lisi@example.com'],
            ['王五', 28, 'wangwu@example.com'],
        ];

        // 填充数据
        $row = 2;
        foreach ($data as $item) {
            $sheet->setCellValue('A' . $row, $item[0]);
            $sheet->setCellValue('B' . $row, $item[1]);
            $sheet->setCellValue('C' . $row, $item[2]);
            $row++;
        }

        // 获取最大行数
        $maxRow = count($data) + 1;

        // 设置列宽
        $sheet->getColumnDimension('A')->setWidth(15);
        $sheet->getColumnDimension('B')->setWidth(10);
        $sheet->getColumnDimension('C')->setWidth(30);

        // 设置表头样式
        $sheet->getStyle('A1:C1')->applyFromArray([
            'font' => [
                'bold' => true,
                'size' => 12
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => [
                    'rgb' => 'CCCCCC',
                ],
            ],
        ]);

        // 设置所有单元格的样式
        $sheet->getStyle('A1:C' . $maxRow)->applyFromArray([
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ]);

        // 自动换行
        $sheet->getStyle('A1:C' . $maxRow)->getAlignment()->setWrapText(true);

        // 设置行高
        for ($i = 1; $i <= $maxRow; $i++) {
            $sheet->getRowDimension($i)->setRowHeight(25);
        }

        // 创建写入对象并输出文件
        $writer = new Xlsx($spreadsheet);

        // 设置header
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="用户数据.xlsx"');
        header('Cache-Control: max-age=0');

        // 输出到浏览器
        $writer->save('php://output');
    }

    /**
     * 导入Excel文件
     */
    public function import($filePath)
    {
        try {
            // 加载Excel文件
            $spreadsheet = IOFactory::load($filePath);
            $sheet = $spreadsheet->getActiveSheet();

            // 获取最大行数
            $maxRow = $sheet->getHighestRow();

            // 存储数据的数组
            $data = [];

            // 从第二行开始读取数据（第一行是表头）
            for ($row = 2; $row <= $maxRow; $row++) {
                $rowData = [
                    'name' => $sheet->getCell('A' . $row)->getValue(),
                    'age' => $sheet->getCell('B' . $row)->getValue(),
                    'email' => $sheet->getCell('C' . $row)->getValue(),
                ];
                $data[] = $rowData;
            }

            // 返回读取到的数据
            return [
                'success' => true,
                'data' => $data
            ];

        } catch (\Exception $e) {
            return [
                'success' => false,
                'message' => $e->getMessage()
            ];
        }
    }
}