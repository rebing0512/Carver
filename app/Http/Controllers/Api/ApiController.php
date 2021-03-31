<?php

namespace App\Http\Controllers\Api;

use App\Models\ReviewCriterionInitialType;
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

    # 初步评审test
    public function test1(){
        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        # 冻结首列/首行、第二行、第三行、第四行
        $spreadsheet->getActiveSheet()->freezePane("B5");
        # 设置工作表标题名称
        $title = '附件二：评选标准表（初步评审）';
        $item_name = '项目名称：中国电信股份有限公司滁州分公司2021-2023年法律顾问服务采购项目';
        $item_num = '项目编号：AHCZ20210003';
        $candidate = [
            '参选人1','参选人2','参选人3','参选人4','参选人5'
        ];
        # 表头
        $worksheet->setCellValueByColumnAndRow(1, 1, $title);
        # 项目名称
        $worksheet->setCellValueByColumnAndRow(1, 2, $item_name);
        # 项目编号
        $worksheet->setCellValueByColumnAndRow(1, 3, $item_num);
        # 设置单元格内容
        $worksheet->setCellValueByColumnAndRow(2, 4, '评审因素');
        $worksheet->setCellValueByColumnAndRow(3, 4, '评审标准');
        $i = 4;
        foreach($candidate as $key=>$val) {
            $worksheet->setCellValueByColumnAndRow($i, 4, $val);
            $i++;
        }
        # 默认列宽
        $spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(14);
        # C列
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(120);
        # 默认行高
        $worksheet->getDefaultRowDimension()->setRowHeight(40);
        # 合并单元格
        $worksheet->mergeCells('A1:C1');
        $worksheet->mergeCells('A2:C2');
        $worksheet->mergeCells('A3:C3');
        # 设置单元格样式
        $styleArray = [
            'font' => [
                'bold' => true
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical'  => Alignment::VERTICAL_CENTER,
            ],
        ];
        $worksheet->getStyle('A1:C1')->applyFromArray($styleArray)->getFont()->setSize(14);
        $styleArray4 = [
            'font' => [
                'bold' => true
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical'  => Alignment::VERTICAL_CENTER,
            ],
        ];
        $worksheet->getStyle('A4:Z4')->applyFromArray($styleArray4)->getFont()->setSize(14);
        # 获取数据
        $data = ReviewCriterionInitialType::query()->with('review')->get()->toArray();
        $rows_counts = 0;
        if($data){
            $i = 5;
            $j = 5;
            foreach($data as $index=>$item){
                $type_name = $item['type_name'];
                $review_data = $item['review'];
                $len = count($review_data);
                $rows_counts = $rows_counts + $len;
                foreach($review_data as $key=>$val) {
                    $styleArrayText = [
                        'alignment' => [
                            'horizontal' => Alignment::HORIZONTAL_LEFT,
                            'vertical' => Alignment::VERTICAL_CENTER,
                        ],
                    ];
//                    $j = $key + 5; # 从表格第5行开始
                    # 设置样式
                    $worksheet->getStyle('A' . $j . ':Z' . $j)->applyFromArray($styleArrayText);
                    # 行高
                    $spreadsheet->getActiveSheet()->getRowDimension($j)->setRowHeight(40);
                    # 设置自动换行
                    $worksheet->getStyle('B' . $j)->getAlignment()->setWrapText(true);
                    $worksheet->getStyle('C' . $j)->getAlignment()->setWrapText(true);
                    $worksheet->setCellValueByColumnAndRow(1, $j, $type_name);
                    $worksheet->setCellValueByColumnAndRow(2, $j, $val['element']);
                    $worksheet->setCellValueByColumnAndRow(3, $j, $val['criterion']);
                    $j++;
                }
                # 合并单元格 todo:合并异常
//                $worksheet->mergeCells('A'.$i.':A'.$j);
                $i = $len + $i;
            }
        }
        $styleArrayBody = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => Border::BORDER_THIN,
                    'color' => ['argb' => '666666'],
                ],
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
            ],
        ];
        $total_rows = $rows_counts + 5;
        # 添加所有边框/居中
        $worksheet->getStyle('A5:I' . $total_rows)->applyFromArray($styleArrayBody);
        $filename = $title.'.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer = SheetIOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
    }
}
