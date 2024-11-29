<?php

namespace App\Http\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use Illuminate\Http\Request;

class ExcelController extends Controller
{
    public function convertSalaryTemplate(Request $request)
    {
        try {
            if (!$request->hasFile('excel_file')) {
                return response()->json(['success' => false, 'message' => '请选择文件']);
            }

            $file = $request->file('excel_file');
            $spreadsheet = IOFactory::load($file->getPathname());
            
            // 获取模板sheet
            $templateSheet = $spreadsheet->getSheetByName('模板');
            if (!$templateSheet) {
                return response()->json(['success' => false, 'message' => '未找到"模板"工作表']);
            }
            
            // 获取sheet1的格式作为参考
            $sheet1 = $spreadsheet->getSheet(0);
            
            // 创建新的sheet2
            if ($spreadsheet->sheetNameExists('sheet2')) {
                $sheet2 = $spreadsheet->getSheet(1);
            } else {
                $sheet2 = $spreadsheet->createSheet();
                $sheet2->setTitle('sheet2');
            }
            
            // 复制模板数据到sheet2
            $templateData = $templateSheet->toArray();
            foreach ($templateData as $rowIndex => $row) {
                foreach ($row as $colIndex => $value) {
                    $sheet2->setCellValueByColumnAndRow($colIndex + 1, $rowIndex + 1, $value);
                }
            }
            
            // 设置列宽（参考sheet1的设置）
            foreach (range('A', $sheet1->getHighestColumn()) as $col) {
                $sheet2->getColumnDimension($col)->setWidth(
                    $sheet1->getColumnDimension($col)->getWidth()
                );
            }
            
            // 设置行高
            $highestRow = $sheet2->getHighestRow();
            for ($row = 1; $row <= $highestRow; $row++) {
                $sheet2->getRowDimension($row)->setRowHeight(-1); // 自动行高
            }
            
            // 设置打印格式
            $sheet2->getPageSetup()
                ->setOrientation(PageSetup::ORIENTATION_LANDSCAPE) // 横向打印
                ->setPaperSize(PageSetup::PAPERSIZE_A4)
                ->setFitToWidth(1)
                ->setFitToHeight(0);
            
            // 保存文件
            $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
            $newFilePath = storage_path('app/public/converted_salary.xlsx');
            $writer->save($newFilePath);
            
            return response()->json([
                'success' => true,
                'message' => '转换成功',
                'file_path' => '/storage/converted_salary.xlsx'
            ]);
            
        } catch (\Exception $e) {
            return response()->json([
                'success' => false,
                'message' => '处理失败：' . $e->getMessage()
            ]);
        }
    }
} 