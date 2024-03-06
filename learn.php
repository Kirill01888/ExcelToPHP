<?php

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

// $cellParams = [
//     'alingment' => [
//         'Horizontal' => $cell->getStyle('B1')->getAlignment()->getHorizontal(),
//         'Vertical' => $cell->getStyle('B1')->getAlignment()->getVertical(),
//         'Rotation' => $cell->getStyle('B1')->getAlignment()->getTextRotation(),
//         'TextWrap' => $cell->getStyle('B1')->getAlignment()->getWrapText()
//     ],
//     'cell' => [
//         'rowHeight' => $cell->getRowDimension('1')->getRowHeight(),
//         'columnWidth' => $cell->getColumnDimension('B')->getWidth(),
//         'cellValue' => $cell->getCell('B1')->getValue()
//     ],
//     'font' => [
//         'size' => $cell->getStyle('B1')->getFont()->getSize(),
//         'name' => $cell->getStyle('B1')->getFont()->getName(),
//         'color' => $cell->getStyle('B1')->getFont()->getColor()->getARGB(),
//         'subscript' => $cell->getStyle('B1')->getFont()->getSubscript(),
//         'superscript' => $cell->getStyle('B1')->getFont()->getSuperscript(),
//         'underline' => $cell->getStyle('B1')->getFont()->getUnderline(),
//         'bold' => $cell->getStyle('B1')->getFont()->getBold(),
//         'italic' => $cell->getStyle('B1')->getFont()->getItalic(),
//         'strikethrough' => $cell->getStyle('B1')->getFont()->getStrikethrough()
//     ],
//     'fill' => [
//         'fillType' => $cell->getStyle('B1')->getFill()->getFillType(),
//         'rotation' => $cell->getStyle('B1')->getFill()->getRotation(),
//         'startColor' => $cell->getStyle('B1')->getFill()->getStartColor()->getARGB(),
//         'endColor' => $cell->getStyle('B1')->getFill()->getEndColor()->getARGB(),
//     ],
//     'borders' => [
//         'left' => $cell->getStyle('B1')->getBorders()->getLeft()->getBorderStyle(),
//         'right' => $cell->getStyle('B1')->getBorders()->getRight()->getBorderStyle(),
//         'top' => $cell->getStyle('B1')->getBorders()->getTop()->getBorderStyle(),
//         'bottom' => $cell->getStyle('B1')->getBorders()->getBottom()->getBorderStyle(),
//         'diagonal' => $cell->getStyle('B1')->getBorders()->getDiagonal()->getBorderStyle(),
// 'outline' => $cell->getStyle('B1')->getBorders()->getOutline()->getBorderStyle(),
// 'inside' => $cell->getStyle('B1')->getBorders()->getInside()->getBorderStyle(),
// 'vertical' => $cell->getStyle('B1')->getBorders()->getVertical()->getBorderStyle(),
// 'horizontal' => $cell->getStyle('B1')->getBorders()->getHorizontal()->getBorderStyle(),
//     ]

// ];
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet;

// class MakeTemplate{

$bord = [
        'top' => ['top' => "'top' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],"],
        'bottom' => ['bottom' => "'bottom' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],"],
        'left' => ['left' => "'left' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],"],
        'right' => ['right' => "'right' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],"],
        'diagonal' => ['diagonal' => "'diagonal' => [
                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                ],"],
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
                ],"
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
    'path' => 'FILL_GRADIENT_PATH'
];

$arrUnderLine = [
    'none' => 'UNDERLINE_NONE',
    'double' => 'UNDERLINE_DOUBLE',
    'doubleAccounting' => 'UNDERLINE_DOUBLEACCOUNTING',
    'single' => 'UNDERLINE_SINGLE',
    'singleAccounting' => 'UNDERLINE_SINGLEACCOUNTING'
];
// function __uct($pathToXlsxFile){
$reader = IOFactory::createReader('Xlsx');
$spr = $reader->load('New One.xlsx');
// // Только чтение данных
$reader->setReadDataOnly(true);
$cell = $spr->getActiveSheet();
// }

// function makeStyleTemplate($WordSheet, $startLetter=1, $endLetter=1, $startRow=9, $endRow=7)
// {

$myHash = '';

$end = ');';

$afterRes = [];

$baseTemplate = [
    '_variable_->getActiveSheet()->getStyle($cell)->applyFromArray(',
    'alignment' => "'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::<HOR>,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::<VER>,
                'textWrap' => <WRAP>,
                'size' => <SIZE>,
                'name' => <NAME>,
                'rotation' => <R>
            ],",
    'font' => "'font' => [
                'bold' => <BOLD>,
                'italic' => <ITAL>,
                'subscript' => <SUB>,
                'superscript' => <SUP>,
                'color' => <ARGB>,
                'underline' => <UND>,
                'strikethrough' => <STK>
            ],"
];

//alignment
$c = $cell->getStyle('B1')->getAlignment()->getHorizontal();
$myHash .= substr($c, 0, 3);
$baseTemplate["alignment"] = str_replace('<HOR>', $arrAlignmentHorizontal[$c], $baseTemplate['alignment']);
$c = $cell->getStyle('B1')->getAlignment()->getVertical();
$myHash .= substr($c, 0, 3);
$baseTemplate["alignment"] = str_replace('<VER>', $arrAlignmentVertical[$c], $baseTemplate["alignment"]);
$c = $cell->getStyle('B1')->getAlignment()->getTextRotation();
$baseTemplate["alignment"] = str_replace('<R>', $c, $baseTemplate["alignment"]);
$c = $cell->getStyle('B1')->getAlignment()->getWrapText() ? 'true' : 'false';
$myHash .= substr($c, 0, 3);
$baseTemplate["alignment"] = str_replace('<WRAP>', $c, $baseTemplate["alignment"]);
$c = $cell->getRowDimension('1')->getRowHeight();
array_push($afterRes, '_variable_->getRowDimension()->setRowHeight();');
$c = $cell->getColumnDimension('B')->getWidth();
array_push($afterRes, '_variable_->getColumnDimension()->setWidth();');
$c = $cell->getCell('B1')->getValue();
array_push($afterRes, '_variable_->getCell()->setValue();');

$font =
    "'font' => 
            'bold' => <BOLD>,
            'italic' => <ITAL>,
            'subscript' => <SUB>,
            'superscript' => <SUP>,
            'color' => <ARGB>,
            'underline' => \PhpOffice\PhpSpreadsheet\Style\Font::<UND>,
            'strikethrough' => <STK>,
            ";
$fill =
    "'fill' => 'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::<FILL>,
        <IFGRAD>";
$fillGrad = "'rotation' => <R>,
            'startColor' => [
            'argb' => <COLS>,
            ],
            'endColor' => [
            'argb' => '<COLE>',
            ],";
//font
$c = $cell->getStyle('B1')->getFont()->getSize();
$baseTemplate["alignment"] = str_replace('<SIZE>', $c, $baseTemplate["alignment"]);
$c = $cell->getStyle('B1')->getFont()->getName();
$myHash .= substr($c, 0, 3);
$baseTemplate["alignment"] = str_replace('<NAME>', $c, $baseTemplate["alignment"]);
$c = $cell->getStyle('B1')->getFont()->getColor()->getARGB();
$baseTemplate["font"] = str_replace('<ARGB>', $c, $baseTemplate["font"]);
$c = $cell->getStyle('B1')->getFont()->getSubscript() ? 'true' : 'false';
$myHash .= substr($c, 0, 3);
$baseTemplate["font"] = str_replace('<SUB>', $c, $baseTemplate["font"]);
$c = $cell->getStyle('B1')->getFont()->getSuperscript() ? 'true' : 'false';
$myHash .= substr($c, 0, 3);
$baseTemplate["font"] = str_replace('<SUP>', $c, $baseTemplate["font"]);
$c = $cell->getStyle('B1')->getFont()->getUnderline();
$myHash .= substr($c, 0, 3);
$baseTemplate["font"] = str_replace('<UND>', $arrUnderLine[$c], $baseTemplate["font"]);
$c = $cell->getStyle('B1')->getFont()->getBold() ? 'true' : 'false';
$myHash .= substr($c, 0, 3);
$baseTemplate["font"] = str_replace('<BOLD>', $c, $baseTemplate["font"]);
$c = $cell->getStyle('B1')->getFont()->getItalic() ? 'true' : 'false';
$myHash .= substr($c, 0, 3);
$baseTemplate["font"] = str_replace('<ITAL>', $c, $baseTemplate["font"]);
$c = $cell->getStyle('B1')->getFont()->getStrikethrough() ? 'true' : 'false';
$myHash .= substr($c, 0, 3);
$baseTemplate["font"] = str_replace('<STK>', $c, $baseTemplate["font"]);

//fill
$c = $cell->getStyle('B1')->getFill()->getFillType();
echo $c=='path';
if ($c == 'linear' or $c == 'path') {
    $baseTemplate['fill'] = $fill;
    $myHash .= substr($c, 0, 3);
    $baseTemplate["fill"] = str_replace('<IFGRAD>', $fillGrad, $baseTemplate["fill"]);
    $baseTemplate["fill"] = str_replace('<FILL>', $arrFills[$c], $baseTemplate["fill"]);
    $myHash .= substr($c, 0, 3);
    $c = $cell->getStyle('B1')->getFill()->getRotation();
    $baseTemplate["fill"] = str_replace('<R>', $c, $baseTemplate["fill"]);
    $c = $cell->getStyle('B1')->getFill()->getStartColor()->getARGB();
    $baseTemplate["fill"] = str_replace('<COLS>', $c, $baseTemplate["fill"]);
    $c = $cell->getStyle('B1')->getFill()->getEndColor()->getARGB();
    $baseTemplate["fill"] = str_replace('<COLE>', $c, $baseTemplate["fill"]);
} else {
    $baseTemplate["fill"] = str_replace('<FILL>', $arrFills[$c], $baseTemplate["fill"]);
}


//borders
$c = $cell->getStyle('B1')->getBorders()->getLeft()->getBorderStyle();
if($c != 'none'){
    $myHash .= substr($c, 0, 3);
    $baseTemplate['borders'] = $bord['left'];
    $baseTemplate['borders']['left'] = str_replace('<BRD>', $c, $baseTemplate['borders']['left']);
}

$c = $cell->getStyle('B1')->getBorders()->getRight()->getBorderStyle();
if($c != 'none'){
    $myHash .= substr($c, 0, 3);
    $baseTemplate['borders'] = $bord['right'];
    $baseTemplate['borders']['right'] = str_replace('<BRD>', $c, $baseTemplate['borders']['right']);
}
$c = $cell->getStyle('B1')->getBorders()->getTop()->getBorderStyle();
if($c != 'none'){
    $myHash .= substr($c, 0, 3);
    $baseTemplate['borders'] = $bord['top'];
    $baseTemplate['borders']['top'] = str_replace('<BRD>', $c, $baseTemplate['borders']['top']);
}
$c = $cell->getStyle('B1')->getBorders()->getBottom()->getBorderStyle();
if($c != 'none'){
    $myHash .= substr($c, 0, 3);
    $baseTemplate['borders'] = $bord['bottom'];
    $baseTemplate['borders']['bottom'] = str_replace('<BRD>', $c, $baseTemplate['borders']['bottom']);
}
$c = $cell->getStyle('B1')->getBorders()->getDiagonal()->getBorderStyle();
if($c != 'none'){
    $myHash .= substr($c, 0, 3);
    $baseTemplate['borders'] = $bord['diagonal'];
    $baseTemplate['borders']['diagonal'] = str_replace('<BRD>', $c, $baseTemplate['borders']['diagonal']);
}
// $c = $cell->getStyle('B1')->getBorders()->getOutline()->getBorderStyle();
// $c = $cell->getStyle('B1')->getBorders()->getInside()->getBorderStyle();
// $c = $cell->getStyle('B1')->getBorders()->getVertical()->getBorderStyle();
// $c = $cell->getStyle('B1')->getBorders()->getHorizontal()->getBorderStyle();

echo "<pre>";
echo $myHash . '<br>';
echo "<br>";
var_dump($baseTemplate);
echo "</pre>";
    // }
// }
