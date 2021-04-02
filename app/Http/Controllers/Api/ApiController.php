<?php

namespace App\Http\Controllers\Api;

use App\Models\CustomerInformation;
use App\Models\ReviewCriterionInitialType;
use App\Models\Users;
use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use PhpOffice\PhpSpreadsheet\IOFactory as SheetIOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Borders;
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
        # 设置对角线
        $spreadsheet->getActiveSheet()->getStyle('F1')->getBorders()->setDiagonalDirection(Borders::DIAGONAL_DOWN );
        $spreadsheet->getActiveSheet()->getStyle('F1')->getBorders()->getDiagonal()-> setBorderStyle(Border::BORDER_THIN);
//        $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()->setFillType(Fill::FILL_SOLID)
//            ->getStartColor()->setARGB(Color::COLOR_DARKGREEN);
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
//        $spreadsheet->getActiveSheet()->getProtection()->setPassword('PhpSpreadsheet');
//        $spreadsheet->getActiveSheet()->getProtection()->setSheet(true);
//        $spreadsheet->getActiveSheet()->getProtection()->setSort(true);
//        $spreadsheet->getActiveSheet()->getProtection()->setInsertRows(true);
//        $spreadsheet->getActiveSheet()->getProtection()->setFormatCells(true);
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
        # 默认列宽
        $spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(14);
        # 默认行高
        $spreadsheet->getActiveSheet()->getDefaultRowDimension()->setRowHeight(40);
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
        # C列
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(120);
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
                \Log::notice('初步评审test--$i:'.$i);
                \Log::notice('初步评审test--$j:'.($j-1));
                \Log::notice('初步评审test--$len:'.$len);
                # 合并单元格
                $worksheet->mergeCells('A'.$i.':A'.($j-1));
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
        $total_rows = $rows_counts + 4;
        \Log::notice('初步评审test--$total_rows:'.$total_rows);
        # 添加所有边框/居中
        $worksheet->getStyle('A5:Z' . $total_rows)->applyFromArray($styleArrayBody);
        $filename = $title.'.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer = SheetIOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
    }

    # 唱价记录表test2
    public function test2(){
        $spreadsheet = new Spreadsheet();
        $worksheet = $spreadsheet->getActiveSheet();
        # 默认列宽
        $spreadsheet->getActiveSheet()->getDefaultColumnDimension()->setWidth(18);
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(25);
        # 行高
        $spreadsheet->getActiveSheet()->getRowDimension(1)->setRowHeight(30);
        $spreadsheet->getActiveSheet()->getRowDimension(2)->setRowHeight(30);
        $spreadsheet->getActiveSheet()->getRowDimension(3)->setRowHeight(30);
        $spreadsheet->getActiveSheet()->getRowDimension(4)->setRowHeight(30);
        $spreadsheet->getActiveSheet()->getRowDimension(5)->setRowHeight(60);
        # 冻结首列/首行、第二行、第三行、第四行、第五行
        $spreadsheet->getActiveSheet()->freezePane("B6");
        # 设置工作表标题名称
        $title = '附件一：唱价记录表';
        $worksheet->setCellValueByColumnAndRow(1, 1, $title);
        # 项目名称
        $item_name = '项目名称：中国电信股份有限公司滁州分公司2021-2023年法律顾问服务采购项目';
        $worksheet->setCellValueByColumnAndRow(1, 2, $item_name);
        # 时间
        $time = '唱价时间：2021年3月25日8时30分00秒';
        $worksheet->setCellValueByColumnAndRow(7, 2, $time);
        # 项目编号
        $item_num = '项目编号：AHCZ20210003';
        $worksheet->setCellValueByColumnAndRow(1, 3, $item_num);
        # 地点
        $address = '唱价地点：安徽省滁州市琅琊区南谯北路1108号电信公司五楼采购中心旁党员活动室';
        $worksheet->setCellValueByColumnAndRow(7, 3, $address);
        # 比选代理
        $agency = '比选代理机构：上海信产管理咨询有限公司';
        $worksheet->setCellValueByColumnAndRow(1, 4, $agency);
        # 委托公证方式
        $consign_type = '唱价过程是否采用委托公证的方式:□是，公证机构名称:  /  ；■否';
        $worksheet->setCellValueByColumnAndRow(7, 4, $consign_type);
        # 设置单元格内容
        $worksheet->setCellValueByColumnAndRow(1, 5, '序号');
        $title1 = '报价内容\n\n参选人';
        $worksheet->setCellValueByColumnAndRow(2, 5, $title1);
        $worksheet->setCellValueByColumnAndRow(3, 5, '法律顾问费用包年价（不含税，元/年）');
        $worksheet->setCellValueByColumnAndRow(4, 5, '税率（%）');
        $worksheet->setCellValueByColumnAndRow(5, 5, '法律顾问费用包年价（含税，元/年）');
        $worksheet->setCellValueByColumnAndRow(6, 5, '每起案件法律顾问费用（含税，元/件）');
        $worksheet->setCellValueByColumnAndRow(7, 5, '税率（%）');
        $worksheet->setCellValueByColumnAndRow(8, 5, '每起案件法律顾问费用（含税，元/件）');
        $worksheet->setCellValueByColumnAndRow(9, 5, '备注');
        $worksheet->setCellValueByColumnAndRow(10, 5, '密封情况');
        $worksheet->setCellValueByColumnAndRow(11, 5, '对现场唱价环节是否有异议');
        $worksheet->setCellValueByColumnAndRow(12, 5, '参选人签名');
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
        $worksheet->getStyle('A1:L1')->applyFromArray($styleArray)->getFont()->setSize(14);
        $styleArray2 = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
                'vertical'  => Alignment::VERTICAL_CENTER,
            ],
        ];
        $worksheet->getStyle('A2:L4')->applyFromArray($styleArray2)->getFont()->setSize(11);
        # 合并单元格
        $worksheet->mergeCells('A1:L1');
        $worksheet->mergeCells('A2:F2');
        $worksheet->mergeCells('G2:L2');
        $worksheet->mergeCells('A3:F3');
        $worksheet->mergeCells('G3:L3');
        $worksheet->mergeCells('A4:F4');
        $worksheet->mergeCells('G4:L4');
        $styleArray4 = [
            'font' => [
                'bold' => true
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical'  => Alignment::VERTICAL_CENTER,
            ],
        ];
        $worksheet->getStyle('A5:Z5')->applyFromArray($styleArray4)->getFont()->setSize(11);
        # 获取数据
        $data = CustomerInformation::query()->oldest('sort')->pluck('name')->toArray();
        $len = count($data);
        $i = 5;
        if($data){
            foreach ($data as $key => $val) {
                $styleArrayText = [
                    'alignment' => [
                        'horizontal' => Alignment::HORIZONTAL_CENTER,
                        'vertical' => Alignment::VERTICAL_CENTER,
                    ],
                ];
                $j = $key + 6; # 从表格第7行开始
                # 设置样式
                $worksheet->getStyle('A1')->applyFromArray($styleArrayText);
                $worksheet->getStyle('A' . $j . ':Z' . $j)->applyFromArray($styleArrayText);
                # 行高
                $spreadsheet->getActiveSheet()->getRowDimension($j)->setRowHeight(40);
                # 设置自动换行
                $worksheet->getStyle('B' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('C' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('D' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('E' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('F' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('G' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('H' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('I' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('J' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('K' . $i)->getAlignment()->setWrapText(true);
                $worksheet->getStyle('L' . $i)->getAlignment()->setWrapText(true);
                $worksheet->setCellValueByColumnAndRow(1, $j, $key+ 1);
                $worksheet->setCellValueByColumnAndRow(2, $j, $val);
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
                'horizontal' => Alignment::HORIZONTAL_CENTER,
                'vertical' => Alignment::VERTICAL_CENTER,
            ],
        ];
        $total_rows = $len + $i;
        \Log::notice('唱价记录--$total_rows:'.$total_rows);
        # 添加所有边框/居中
        $worksheet->getStyle('A5:Z' . $total_rows)->applyFromArray($styleArrayBody);
        $filename = $title.'.xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer = SheetIOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
    }
}
