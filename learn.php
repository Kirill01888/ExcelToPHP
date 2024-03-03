<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet;

//str_replace

function getInner($arr){
    $str = '';
    foreach($arr as $k => $v){
    }
    if (is_array($arr)){
        return getInner($arr);
    }
    return $str.$arr;
}

function makeStyleTemplate()
{
    $baseTemplate = [
            '_variable_->getActiveSheet()->getStyle($cell)->applyFromArray(',
            'alingment' => "'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_,]
            ],",
            ');'
        ];

    $str = '';

    foreach($baseTemplate as $value){
        if(is_array($value)){
            foreach($value as $v){
                echo $v."<br>";
            }
        }
        echo $value."<br>";
    }
}

$align = [
    "'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::<HOR>,
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::<VER>,
        'textWrap' => <WRAP>,
        'size' => <SIZE>,
        'name' => <NAME>
    ],"

];

$bord = [
    'borders' => [
        'top' => "'top' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'bottom' => "'bottom' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'left' => "'left' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'right' => "'right' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'diagonal' => "'diagonal' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'allBorders' => "'allBorders' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'outline' => "'outline' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'inside' => "'inside' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'vertical' => "'vertical' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
        'horizontal' => "'horizontal' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],",
    ],
];

$arrBorders = [
    'none' => 'BORDER_NONE',
    'dashDot' => 'BORDER_DASHDOT',
    'dashDotDot' => 'BORDER_DASHDOTDOT',
    'dashed' => 'BORDER_DASHED',
    'dotted' => 'BORDER_DOTTED',
    'double' => 'BORDER_DOUBLE',
    'hair' => 'BORDER_HAIR',
    'medium' => 'BORDER_MEDIUM',
    'mediumDashDot' => 'BORDER_MEDIUMDASHDOT',
    'mediumDashDotDot' => 'BORDER_MEDIUMDASHDOTDOT',
    'mediumDashed' => 'BORDER_MEDIUMDASHED',
    'slantDashDot' => 'BORDER_SLANTDASHDOT',
    'thick' => 'BORDER_THICK',
    'thin' => 'BORDER_THIN',
    'omit' => 'BORDER_OMIT'
];

$arrAlignmentHorizontal = [
    'general' => 'HORIZONTAL_GENERAL',
    'left' => 'HORIZONTAL_LEFT',
    'right' => 'HORIZONTAL_RIGHT',
    'center' => 'HORIZONTAL_CENTER',
    'centerContinuous' => 'HORIZONTAL_CENTER_CONTINUOUS',
    'justify' => 'HORIZONTAL_JUSTIFY',
    'fill' => 'HORIZONTAL_FILL',
    'distributed' => 'HORIZONTAL_DISTRIBUTED'
];

$arrAlignmentVertical = [
    'bottom' => 'VERTICAL_BOTTOM',
    'top' => 'VERTICAL_TOP',
    'center' => 'VERTICAL_CENTER',
    'justify' => 'VERTICAL_JUSTIFY',
    'distributed' => 'VERTICAL_DISTRIBUTED'
];

$arrFonts = [
    'none' => 'UNDERLINE_NONE',
    'double' => 'UNDERLINE_DOUBLE',
    'doubleAccounting' => 'UNDERLINE_DOUBLEACCOUNTING',
    'single' => 'UNDERLINE_SINGLE',
    'singleAccounting' => 'UNDERLINE_SINGLEACCOUNTING',
];

$arrFills = [
    'none' => 'FILL_NONE',
    'solid' => 'FILL_SOLID',
    'linear' => 'FILL_GRADIENT_LINEAR',
    'path' => 'FILL_GRADIENT_PATH',
    'darkDown' => 'FILL_PATTERN_DARKDOWN',
    'darkGray' => 'FILL_PATTERN_DARKGRAY',
    'darkGrid' => 'FILL_PATTERN_DARKGRID',
    'darkHorizontal' => 'FILL_PATTERN_DARKHORIZONTAL',
    'darkTrellis' => 'FILL_PATTERN_DARKTRELLIS',
    'darkUp' => 'FILL_PATTERN_DARKUP',
    'darkVertical' => 'FILL_PATTERN_DARKVERTICAL',
    'gray0625' => 'FILL_PATTERN_GRAY0625',
    'gray125' => 'FILL_PATTERN_GRAY125',
    'lightDown' => 'FILL_PATTERN_LIGHTDOWN',
    'lightGray' => 'FILL_PATTERN_LIGHTGRAY',
    'lightGrid' => 'FILL_PATTERN_LIGHTGRID',
    'lightHorizontal' => 'FILL_PATTERN_LIGHTHORIZONTAL',
    'lightTrellis' => 'FILL_PATTERN_LIGHTTRELLIS',
    'lightUp' => 'FILL_PATTERN_LIGHTUP',
    'lightVertical' => 'FILL_PATTERN_LIGHTVERTICAL',
    'mediumGray' => 'FILL_PATTERN_MEDIUMGRAY',
];

// ($exportedArray, 'baseLine', $this->getBaseLine());
// ($exportedArray, 'bold', $this->getBold());
// ($exportedArray, 'cap', $this->getCap());
// ($exportedArray, 'chartColor', $this->getChartColor());
// ($exportedArray, 'color', $this->getColor());
// ($exportedArray, 'complexScript', $this->getComplexScript());
// ($exportedArray, 'eastAsian', $this->getEastAsian());
// ($exportedArray, 'italic', $this->getItalic());
// ($exportedArray, 'latin', $this->getLatin());
// ($exportedArray, 'name', $this->getName());
// ($exportedArray, 'scheme', $this->getScheme());
// ($exportedArray, 'size', $this->getSize());
// ($exportedArray, 'strikethrough', $this->getStrikethrough());
// ($exportedArray, 'strikeType', $this->getStrikeType());
// ($exportedArray, 'subscript', $this->getSubscript());
// ($exportedArray, 'superscript', $this->getSuperscript());
// ($exportedArray, 'underline', $this->getUnderline());
// ($exportedArray, 'underlineColor', $this->getUnderlineColor());

// foreach($arrFills as $a){
//     $b = explode(' ', $a);
//     echo join(' ', array_reverse($b)).",";
//     echo "<br>";
// }


// $writer = new Xlsx($spreadsheet);

// $writer->save('New One.xlsx');



$reader = IOFactory::createReader('Xlsx');
$spr = $reader->load('New One.xlsx');
// // Только чтение данных
$reader->setReadDataOnly(true);
$cur = $spr->getActiveSheet();

// 'font' => [
//     'bold' => true,
// ],
// 'alignment' => [
//     'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_RIGHT,
// ],
// 'borders' => [
//     'top' => [
//         'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
//     ],
// ],
// 'fill' => [
//     'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
//     'rotation' => 90,
//     'startColor' => [
//         'argb' => 'FFA0A0A0',
//     ],
//     'endColor' => [
//         'argb' => 'FFFFFFFF',
//     ],
// ],

for ($i = 1; $i <= 1; $i++) {
    echo "<h1>------------------------</h1>";
    $cell = $cur->getCell('A1');
    $cellStyle = $cell->getStyle()->getBorders()->getBottom()->getBorderStyle();
    $cellAlingHor = $cell->getStyle()->getAlignment()->getHorizontal();
    $cellAlingVer = $cell->getStyle()->getAlignment()->getVertical();
    $cellFont = $cell->getStyle()->getFont()->getCap();
    // $cellAlingVer = $cell->getStyle()->;
    // $cellAlingVer = $cell->getStyle()->;
    // $cellAlingVer = $cell->getStyle()->;
    echo "</pre>";
    var_dump($cellAlingHor, $cellAlingVer);
    echo "<h1>------------------------</h1>";
    echo "</pre>";
}
