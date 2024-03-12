<?php

namespace ExcelToPhp;

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

class Converter
{
    private $col = "'color' => [
        'rgb' => <C>
        ]";

    private $alignment = "\n'alignment' => [\n
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::<HOR>,\n
        'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::<VER>,\n
        'textWrap' => <WRAP>,\n
        'size' => <SIZE>,\n
        'name' => <NAME>,\n
        'rotation' => <R>\n
        ],";

    private $bord = [
        'top' => "'top' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::<BRD>,
                        <COL>
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
                        ],"
    ];

    private $arrBorders = [
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

    private $arrAlignmentHorizontal = [
        'general' => 'HORIZONTAL_GENERAL',
        'left' => 'HORIZONTAL_LEFT',
        'right' => 'HORIZONTAL_RIGHT',
        'center' => 'HORIZONTAL_CENTER',
        'centerContinuous' => 'HORIZONTAL_CENTER_CONTINUOUS',
        'justify' => 'HORIZONTAL_JUSTIFY',
        'fill' => 'HORIZONTAL_FILL',
        'distributed' => 'HORIZONTAL_DISTRIBUTED'
    ];

    private $arrAlignmentVertical = [
        'bottom' => 'VERTICAL_BOTTOM',
        'top' => 'VERTICAL_TOP',
        'center' => 'VERTICAL_CENTER',
        'justify' => 'VERTICAL_JUSTIFY',
        'distributed' => 'VERTICAL_DISTRIBUTED'
    ];

    private $arrFonts = [
        'none' => 'UNDERLINE_NONE',
        'double' => 'UNDERLINE_DOUBLE',
        'doubleAccounting' => 'UNDERLINE_DOUBLEACCOUNTING',
        'single' => 'UNDERLINE_SINGLE',
        'singleAccounting' => 'UNDERLINE_SINGLEACCOUNTING',
    ];

    private $arrFills = [
        'none' => 'FILL_NONE',
        'solid' => 'FILL_SOLID',
        'linear' => 'FILL_GRADIENT_LINEAR',
        'path' => 'FILL_GRADIENT_PATH'
    ];

    private $arrUnderLine = [
        'none' => 'UNDERLINE_NONE',
        'double' => 'UNDERLINE_DOUBLE',
        'doubleAccounting' => 'UNDERLINE_DOUBLEACCOUNTING',
        'single' => 'UNDERLINE_SINGLE',
        'singleAccounting' => 'UNDERLINE_SINGLEACCOUNTING'
    ];

    private $font =
    "'font' => [\n
            'bold' => <BOLD>,\n
            'italic' => <ITAL>,\n
            'subscript' => <SUB>,\n
            'superscript' => <SUP>,\n
            'color' => <ARGB>,\n
            'underline' => <UND>,\n
            'strikethrough' => <STK>\n
            ],";
    private $fill =
    "\n'fill' => [
        \n'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::<FILL>,\n
        <IFGRAD>\n
        ]";
    private $fillGrad = 
        "'rotation' => <R>,\n
        'startColor' => [\n
        'argb' => <COLS>,\n
        ],\n
        'endColor' => [\n
        'argb' => '<COLE>',\n
        ],\n";
    
    private static $instance;
    private static $fileName;
    private static $range;
    private static $fileReader;

    private function __construct($file, $r)
    {
        $fileR = IOFactory::createReader('Xlsx');
        self::$fileReader = $fileR->load($file)->getActiveSheet();
        self::$range = $r;
    }

    public static function getInstance($fileN, $r): Converter
    {
        if (self::$instance === null || self::$fileName != $fileN || self::$range != $r) {
            self::$instance = new Converter($fileN, $r);
            return self::$instance;
        } else {
            return self::$instance;
        }
    }

    public function makeTemplate()
    {
        $res = '';
        $end = ');';
        $afterRes = [];
        $range = explode(':', self::$range);
        $start = $range[0];
        $end = $range[1];
        $s = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        
        for($startWord = strpos($s, $start[0]); $startWord <= strpos($s, $end[0]); $start++){
            $word = $s[$startWord];
            for($startRow = (int)$start[1]; $startRow < (int)$end[1]; $startRow++){
                $baseTemplate = [
                    '_variable_->getActiveSheet()->getStyle($cells)->applyFromArray(',
                    'alignment' => "",
                    'font' => "",
                    'fill' => "",
                    'borders' => ""
                ];
                $baseTemplate['alignment'] = $this->alignment;
                $baseTemplate['font'] = $this->font;
                $myHash = '';
                $arrOfHashAndCell = [];
                //alignment
                $c = self::$fileReader->getStyle($word.$startRow)->getAlignment()->getHorizontal();
                $myHash .= substr($c, 0, 3);
                $baseTemplate["alignment"] = str_replace('<HOR>', $this->arrAlignmentHorizontal[$c], $baseTemplate['alignment']);
                $c = self::$fileReader->getStyle($word.$startRow)->getAlignment()->getVertical();
                $myHash .= substr($c, 0, 3);
                $baseTemplate["alignment"] = str_replace('<VER>', $this->arrAlignmentVertical[$c], $baseTemplate["alignment"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getAlignment()->getTextRotation();
                $baseTemplate["alignment"] = str_replace('<R>', $c, $baseTemplate["alignment"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getAlignment()->getWrapText() ? 'true' : 'false';
                $myHash .= substr($c, 0, 3);
                $baseTemplate["alignment"] = str_replace('<WRAP>', $c, $baseTemplate["alignment"]);
                $c = self::$fileReader->getRowDimension('1')->getRowHeight();
                array_push($afterRes, '_variable_->getRowDimension("'.$startRow.'")->setRowHeight('.$c.');');
                $c = self::$fileReader->getColumnDimension('B')->getWidth();
                array_push($afterRes, '_variable_->getColumnDimension()->setWidth('.$c.');');
                $c = self::$fileReader->getCell($word.$startRow)->getValue();
                array_push($afterRes, '_variable_->getCell()->setValue('.$c.');');
        
                //font
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getSize();
                $baseTemplate["alignment"] = str_replace('<SIZE>', $c, $baseTemplate["alignment"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getName();
                $myHash .= substr($c, 0, 3);
                $baseTemplate["alignment"] = str_replace('<NAME>', $c, $baseTemplate["alignment"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getColor()->getARGB();
                $baseTemplate["font"] = str_replace('<ARGB>', $c, $baseTemplate["font"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getSubscript() ? 'true' : 'false';
                $myHash .= substr($c, 0, 3);
                $baseTemplate["font"] = str_replace('<SUB>', $c, $baseTemplate["font"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getSuperscript() ? 'true' : 'false';
                $myHash .= substr($c, 0, 3);
                $baseTemplate["font"] = str_replace('<SUP>', $c, $baseTemplate["font"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getUnderline();
                $myHash .= substr($c, 0, 3);
                $baseTemplate["font"] = str_replace('<UND>', $this->arrUnderLine[$c], $baseTemplate["font"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getBold() ? 'true' : 'false';
                $myHash .= substr($c, 0, 3);
                $baseTemplate["font"] = str_replace('<BOLD>', $c, $baseTemplate["font"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getItalic() ? 'true' : 'false';
                $myHash .= substr($c, 0, 3);
                $baseTemplate["font"] = str_replace('<ITAL>', $c, $baseTemplate["font"]);
                $c = self::$fileReader->getStyle($word.$startRow)->getFont()->getStrikethrough() ? 'true' : 'false';
                $myHash .= substr($c, 0, 3);
                $baseTemplate["font"] = str_replace('<STK>', $c, $baseTemplate["font"]);
        
                //fill
                $c = self::$fileReader->getStyle($word.$startRow)->getFill()->getFillType();
                echo $c == 'path';
                if ($c == 'linear' or $c == 'path') {
                    $baseTemplate['fill'] = $this->fill;
                    $myHash .= substr($c, 0, 3);
                    $baseTemplate["fill"] = str_replace('<IFGRAD>', $this->fillGrad, $baseTemplate["fill"]);
                    $baseTemplate["fill"] = str_replace('<FILL>', $this->arrFills[$c], $baseTemplate["fill"]);
                    $myHash .= substr($c, 0, 3);
                    $c = self::$fileReader->getStyle($word.$startRow)->getFill()->getRotation();
                    $baseTemplate["fill"] = str_replace('<R>', $c, $baseTemplate["fill"]);
                    $c = self::$fileReader->getStyle($word.$startRow)->getFill()->getStartColor()->getARGB();
                    $baseTemplate["fill"] = str_replace('<COLS>', $c, $baseTemplate["fill"]);
                    $c = self::$fileReader->getStyle($word.$startRow)->getFill()->getEndColor()->getARGB();
                    $baseTemplate["fill"] = str_replace('<COLE>', $c, $baseTemplate["fill"]);
                } else {
                    $baseTemplate["fill"] = str_replace('<FILL>', $this->arrFills[$c], $baseTemplate["fill"]);
                }
        
                //borders
                $c = self::$fileReader->getStyle($word.$startRow)->getBorders()->getLeft()->getBorderStyle();
                if ($c != 'none') {
                    $myHash .= substr($c, 0, 3);
                    $baseTemplate['borders'] = $this->bord['left'];
                    $baseTemplate['borders']['left'] = str_replace('<BRD>', $this->arrBorders[$c], $baseTemplate['borders']['left']);
                }
        
                $c = self::$fileReader->getStyle($word.$startRow)->getBorders()->getRight()->getBorderStyle();
                if ($c != 'none') {
                    $myHash .= substr($c, 0, 3);
                    $baseTemplate['borders'] = $this->bord['right'];
                    $baseTemplate['borders']['right'] = str_replace('<BRD>', $this->arrBorders[$c], $baseTemplate['borders']['right']);
                }
                $c = self::$fileReader->getStyle($word.$startRow)->getBorders()->getTop()->getBorderStyle();
                if ($c != 'none') {
                    $myHash .= substr($c, 0, 3);
                    $baseTemplate['borders'] = $this->bord['top'];
                    $baseTemplate['borders']['top'] = str_replace('<BRD>', $this->arrBorders[$c], $baseTemplate['borders']['top']);
                }
                $c = self::$fileReader->getStyle($word.$startRow)->getBorders()->getBottom()->getBorderStyle();
                if ($c != 'none') {
                    $myHash .= substr($c, 0, 3);
                    $baseTemplate['borders'] = $this->bord['bottom'];
                    $baseTemplate['borders']['bottom'] = str_replace('<BRD>', $this->arrBorders[$c], $baseTemplate['borders']['bottom']);
                }
                $c = self::$fileReader->getStyle($word.$startRow)->getBorders()->getDiagonal()->getBorderStyle();
                if ($c != 'none') {
                    $myHash .= substr($c, 0, 3);
                    $baseTemplate['borders'] = $this->bord['diagonal'];
                    $baseTemplate['borders']['diagonal'] = str_replace('<BRD>', $this->arrBorders[$c], $baseTemplate['borders']['diagonal']);
                }
            }
        }
    }
}

// $sh = IOFactory::createReader('Xlsx')->load('New One.xlsx')->getActiveSheet();

// echo $sh->getCell("A1")->getCoordinate();

$s = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
$p = strpos($s, 'Z');
