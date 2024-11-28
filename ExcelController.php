<?php

namespace App\Controllers;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

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
    public function import($file)
    {
        try {
            // 加载Excel文件
            $spreadsheet = IOFactory::load($file);
            $sheet = $spreadsheet->getActiveSheet();
            
            // 获取最大行数
            $maxRow = $sheet->getHighestRow();
            
            // 存储数据的数组
            $data = [];
            
            // 从第二行开始读取数据（第一行是表头）
            for ($row = 2; $row <= $maxRow; $row++) {
                $rowData = [
                    'name' => $sheet->getCellValue('A' . $row),
                    'age' => $sheet->getCellValue('B' . $row),
                    'email' => $sheet->getCellValue('C' . $row),
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