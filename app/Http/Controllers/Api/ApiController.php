<?php

namespace App\Http\Controllers\Api;

use App\Models\Users;
use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use PhpOffice\PhpSpreadsheet\IOFactory as SheetIOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;

class ApiController extends Controller
{
    /**
     * @param Request $Request
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     *
     * PHP Excel Export
     */
    public function phpExcel(Request $Request){
        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        # 设置工作表标题名称
        $worksheet->setTitle('用户信息');
        # 表头
        # 设置单元格内容
        $worksheet->setCellValueByColumnAndRow(1, 1, '用户信息');
        $worksheet->setCellValueByColumnAndRow(1, 2, '姓名');
        $worksheet->setCellValueByColumnAndRow(2, 2, '手机号');
        $worksheet->setCellValueByColumnAndRow(3, 2, 'openid');
        $worksheet->setCellValueByColumnAndRow(4, 2, '真实姓名');
        $worksheet->setCellValueByColumnAndRow(5, 2, '余额');
        # 设置字体色
        $color = Color::COLOR_GREEN;
        $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->getColor()->setARGB($color);
//         $spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->getColor()->setARGB($color);
//         $spreadsheet->getActiveSheet()->getStyle('C1')->getFont()->getColor()->setARGB($color);
//         $spreadsheet->getActiveSheet()->getStyle('D1')->getFont()->getColor()->setARGB($color);
//         $spreadsheet->getActiveSheet()->getStyle('E1')->getFont()->getColor()->setARGB($color);
        # 设置背景色
        $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_DARKGREEN);
        # 列宽
        $spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(12);
        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(35);
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(40);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(15);
        # 合并单元格
        $worksheet->mergeCells('A1:E1');
        $styleArray = [
            'font' => [
                'bold' => true
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
            ],
        ];
        # 设置单元格样式
        $worksheet->getStyle('A1')->applyFromArray($styleArray)->getFont()->setSize(28);
        $worksheet->getStyle('A2:E2')->applyFromArray($styleArray)->getFont()->setSize(14);
        $data = Users::query()->get();
        $len = count($data);
//        $j = 0;
        foreach ($data as $key=>$val) {
            $j = $key + 3; # 从表格第3行开始
            $worksheet->setCellValueByColumnAndRow(1, $j, $val['username']);
            $worksheet->setCellValueByColumnAndRow(2, $j, $val['phone']);
            $worksheet->setCellValueByColumnAndRow(3, $j, $val['openid']);
            $worksheet->setCellValueByColumnAndRow(4, $j, $val['real_name']);
            $worksheet->setCellValueByColumnAndRow(5, $j, $val['balance'] + $val['uncash_balance']);
        }
        $styleArrayBody = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => '666666'],
                ],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
            ],
        ];
        $total_rows = $len + 2;
        # 添加所有边框/居中
        $worksheet->getStyle('A1:E'.$total_rows)->applyFromArray($styleArrayBody);
        $filename = '用户信息.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="'.$filename.'"');
        header('Cache-Control: max-age=0');
        $writer = SheetIOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
    }
}
