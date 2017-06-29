<?php
  ini_set('memory_limit', '-1');
  date_default_timezone_set('Asia/Taipei');
  require './vendor/autoload.php';

  ### USL7
  $file = './tmp/USL7.xlsx';
  $min_ins_amount = 5000;
  $max_ins_amount = 7000;
  $distance = 1000;
  $ins_amount_pos = 'R8';
  $ins_fee_pos = 'J19';

  ### FID
  // $file = './tmp/FID.xlsx';
  // $ins_amount = 1;
  // $ins_amount_pos = 'J36';
  // $ins_fee_pos = 'R34';

  try {
   $objPHPExcel = PHPExcel_IOFactory::load($file);
  } catch(Exception $e) {
   die('Error loading file "'.pathinfo($file,PATHINFO_BASENAME).'": '.$e->getMessage());
  }
  PHPExcel_Calculation::getInstance($objPHPExcel)->disableCalculationCache();

  $getSheet = $objPHPExcel->getActiveSheet();
  $ins_amount = $min_ins_amount;
  while ($ins_amount <= $max_ins_amount) {
    echo "{$ins_amount_pos}: ".$getSheet->getCell($ins_amount_pos)->getCalculatedValue()."\n";
    echo "{$ins_fee_pos}: ".$getSheet->getCell($ins_fee_pos)->getCalculatedValue()."\n";
    $getSheet->setCellValue($ins_amount_pos,$ins_amount);

    echo "{$ins_amount_pos}: ".$getSheet->getCell($ins_amount_pos)->getCalculatedValue()."\n";
    echo "{$ins_fee_pos}: ".$getSheet->getCell($ins_fee_pos)->getCalculatedValue()."\n";
    echo "\n";

    $ins_amount = $ins_amount + $distance;
    echo "============================\n";
  }
?>
