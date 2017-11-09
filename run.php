<?php
  ini_set('memory_limit', '-1');
  date_default_timezone_set('Asia/Taipei');
  require './vendor/autoload.php';

  ### USL7
  // $file = './tmp/USL7.xlsx';
  // $min_ins_amount = 5000;
  // $max_ins_amount = 7000;
  // $distance = 1000;
  // $ins_amount_pos = 'R8';
  // $ins_fee_pos = 'J19';

  ### DY1
  // $file = './tmp/DY1.xls';
  // $min_ins_amount = 100;
  // $max_ins_amount = 300;
  // $distance = 100;
  // $ins_amount_pos = 'I15';
  // $ins_fee_pos = 'I47';

  ### ISJ
  // $file = './tmp/ISJ.xlsx';
  // $min_ins_amount = 100;
  // $max_ins_amount = 300;
  // $distance = 100;
  // $ins_amount_pos = 'P10';
  // $ins_fee_pos = 'T10';

  ### EW
  $file = './tmp/EW.xls';
  $min_ins_amount = 100;
  $max_ins_amount = 300;
  $distance = 100;
  $ins_amount_pos = 'J8';
  $ins_fee_pos = 'H16';

  echo "Filepath: {$file}\n";
  try {
   $objPHPExcel = PHPExcel_IOFactory::load($file);
  } catch(Exception $e) {
   die('Error loading file "'.pathinfo($file,PATHINFO_BASENAME).'": '.$e->getMessage());
  }
  PHPExcel_Calculation::getInstance($objPHPExcel)->disableCalculationCache();

  $getSheet = $objPHPExcel->getActiveSheet();
  $ins_amount = $min_ins_amount;
  while ($ins_amount <= $max_ins_amount) {
    // echo "{$ins_amount_pos}: ".$getSheet->getCell($ins_amount_pos)->getCalculatedValue()."\n";
    // echo "{$ins_fee_pos}: ".$getSheet->getCell($ins_fee_pos)->getCalculatedValue()."\n";
    $getSheet->setCellValue($ins_amount_pos,$ins_amount);

    echo "{$ins_amount_pos}: ".$getSheet->getCell($ins_amount_pos)->getCalculatedValue()."\n";
    echo "{$ins_fee_pos}: ".$getSheet->getCell($ins_fee_pos)->getCalculatedValue()."\n";
    echo "\n";

    $ins_amount = $ins_amount + $distance;
    echo "============================\n";
  }
?>
