<?php

require './vendor/autoload.php';
$file = './tmp/test.xlsx';

PHPExcel_Calculation::getInstance()->setCalculationCacheEnabled(false);

try {
 $objPHPExcel = PHPExcel_IOFactory::load($file);
} catch(Exception $e) {
 die('Error loading file "'.pathinfo($file,PATHINFO_BASENAME).'": '.$e->getMessage());
}

$getSheet = $objPHPExcel->getActiveSheet();
$value = $getSheet->getCell('B2')->getValue();
echo $getSheet->getCell('A2')->getCalculatedValue()."\n";
echo $getSheet->getCell('A3')->getCalculatedValue()."\n";
echo $getSheet->getCell('B2')->getValue()."\n";
echo $getSheet->getCell('B2')->getCalculatedValue()."\n";
echo $getSheet->getCell('B2')->getFormattedValue()."\n";

echo "\n";

$getSheet->setCellValue('A2',12);
PHPExcel_Calculation::getInstance( $getSheet->getParent() )->flushInstance();
echo $getSheet->getCell('A2')->getCalculatedValue()."\n";
echo $getSheet->getCell('A3')->getCalculatedValue()."\n";
echo $getSheet->getCell('B2')->getValue()."\n";
echo $getSheet->getCell('B2')->getCalculatedValue()."\n";
echo $getSheet->getCell('B2')->getFormattedValue()."\n";

?>
