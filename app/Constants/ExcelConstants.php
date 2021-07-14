<?php
namespace App\Constants\Common;

use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

class ExcelConstants
{
    const DEFAULT_HEADER = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
    ];

    const HEADER_BORDER = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const HEADER_BORDER_BOLD_INCREASE = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
            'color' => ['argb' => 'FFC40000'],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const DEFAULT_DATA =[
        'font' => [
            'bold' => false,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
    ];

    const DATA_BORDER =[
        'font' => [
            'bold' => false,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const DATA_BORDER_STRING = [
        'font' => [
            'bold' => false,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
        'dataType' => DataType::TYPE_STRING,
    ];

    const DATA_BORDER_NUMBER_FORMAT =[
        'font' => [
            'bold' => false,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
        'numberFormat' => [
            'formatCode' => '#,##0',
        ],
    ];

    const DATA_BORDER_LEFT =[
        'font' => [
            'bold' => false,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_LEFT,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const DATA_BORDER_BOLD =[
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allBorders' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const DATA_BORDER_INCREASE = [
        'font' => [
            'name' => '맑은 고딕',
            'size' => 9,
            'color' => ['argb' => 'FFC40000'],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const DATA_BORDER_BOLD_INCREASE = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
            'color' => ['argb' => 'FFC40000'],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const DATA_BORDER_BOLD_DECREASE = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
            'color' => ['argb' => 'FF4472C4'],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const DEFAULT_SUMMARY = [
        'font' => [
            'bold' => false,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
    ];

    const SUMMARY_BOLD_BORDER = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const SUMMARY_BOLD_BORDER_NUMBER_FORMAT = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
        'numberFormat' => [
            'formatCode' => '#,##0',
        ],
    ];

    const DEFAULT_DATA_NUMBER_FORMAT = [
        'font' => [
            'bold' => false,
            'name' => '맑은 고딕',
            'size' => 9,
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'numberFormat' => [
            'formatCode' => '#,##0',
        ],
    ];

    const SUMMARY_BOLD_BORDER_INCREASE = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
            'color' => ['argb' => 'FFC40000'],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];

    const SUMMARY_BOLD_BORDER_DECREASE = [
        'font' => [
            'bold' => true,
            'name' => '맑은 고딕',
            'size' => 9,
            'color' => ['argb' => 'FF4472C4'],
        ],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'outline' => [
                'borderStyle' => Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
    ];
}
