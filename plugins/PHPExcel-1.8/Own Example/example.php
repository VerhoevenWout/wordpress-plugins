<?php


/**
 *
 * WRITE TO EXCEL SIMPLE
 *
 */
require_once('lib/PHPExcel.php');
$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'Hello')
            ->setCellValue('B2', 'world!')
            ->setCellValue('C1', 'Hello')
            ->setCellValue('D2', 'world!');
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save('/Users/woutverhoeven/Desktop/some_excel_file.xlsx');





/**
 *
 * WRITE TO EXCEL DIFFERENT METHOD
 *
 */
require_once('lib/PHPExcel.php');
$user_display_name = str_replace(' ', '_', $user_display_name);
$export_name = 'xls_export_'.$user_display_name.'.xls';
// $upload_dir = wp_upload_dir()['baseurl'].'/xls_exports/'.$export_name;
// $upload_dir = wp_upload_dir()['basedir'];
$upload_dir = '/Users/woutverhoeven/Desktop/test.xls';

$excel 		= [];
$excel[] 	= ['test'];
$excel[] 	= [];


$objPHPExcel = new PHPExcel();
$objPHPExcel->setActiveSheetIndex(0);
$objPHPExcel->getActiveSheet()->fromArray($excel, null, 'A1');
// $objPHPExcel->setActiveSheetIndex(0)
//             ->setCellValue('A1', 'Hello')
//             ->setCellValue('B2', 'world!')
//             ->setCellValue('C1', 'Hello')
//             ->setCellValue('D2', 'world!');
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$objWriter->save($upload_dir);





/**
 *
 * VENUES ONLINE EXAMPLE
 *
 */
require_once('lib/PHPExcel.php');
$doc = new \PHPExcel();

$export_id = '1';
$export_name = 'xls_export_'.$offerteId.'.xls';
$upload_dir = $this->theme->getUploadDir().'/xls_exports/'.$export_name;
$upload_name = 'Web_offerteaanvraag_';

$excel 		= [];
$excel[] 	= [$export_name];
$excel[] 	= [];

foreach ($currentuser as $key => $value) {
	$excel[] = [$key,$value];
}

$excel[] = [];
$excel[] = ['opmerking:'];
$excel[] = [$opmerking];
$excel[] = [];
$excel[] = [
	'ArtikelNr' => 'ArtikelNr',
	'Titel' => 'Titel', 
	'Aantal' => 'Aantal'
];
$excel[] = [];

$doc->setActiveSheetIndex(0);
$doc->getActiveSheet()->fromArray($excel, null, 'A1');
$doc->getActiveSheet()->getStyle('A19:C19')->applyFromArray(
        array(
            'font' => array(
                'bold' => true
            )
        )
);

foreach(range('A','G') as $columnID) {
	$doc->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
}
$doc->getActiveSheet()->getStyle('A1:A100')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$doc->getActiveSheet()->getStyle('B1:B100')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$doc->getActiveSheet()->getStyle('C1:C100')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="your_name.xls"');
header('Cache-Control: max-age=0');

$writer = \PHPExcel_IOFactory::createWriter($doc, 'Excel5');
$writer->save($uploadDir);





/**
 *
 * MULLER EXAMPLE
 *
 */
$fileOfferteNaam = $offerteNaam ? '-'.$offerteNaam : '';
$uploadName = 'Web_offerteaanvraag_'.$date.'-'.$display_name.'-'.$offerteId.$fileOfferteNaam.'.xls';
$uploadDir = $this->theme->getUploadDir().'/offerte/muller_order_'.$offerteId.'.xls';
$excel = [];
$excel[] = [$fullRef];
$excel[] = [$offerteNaam];
$excel[] = [];

foreach ($currentuser as $key => $value) {
	$excel[] = [$key,$value];
}


$excel[] = [];

$excel[] = ['opmerking:'];
$excel[] = [$opmerking];

$excel[] = [];

$excel[] = [
	'ArtikelNr' => 'ArtikelNr',
	'Titel' => 'Titel', 
	'Aantal' => 'Aantal'
];



$cart = json_decode(stripcslashes($_POST['products']));
foreach ($cart as $key => $product) {
	$excel[] = [
		'ArtikelNr' => $this->theme->product->formatnr($product->intern_artikelnr),
		'Titel' => $product->post_title, 
		'Aantal' => $product->count
	];
}

$excel[] = [];

require_once('lib/PHPExcel.php');
$doc = new \PHPExcel();
$doc->setActiveSheetIndex(0);
$doc->getActiveSheet()->fromArray($excel, null, 'A1');
$doc->getActiveSheet()->getStyle('A19:C19')->applyFromArray(
        array(
            'font' => array(
                'bold' => true
            )
        )
);

foreach(range('A','G') as $columnID) {
	$doc->getActiveSheet()->getColumnDimension($columnID)->setAutoSize(true);
}
$doc->getActiveSheet()->getStyle('A1:A100')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$doc->getActiveSheet()->getStyle('B1:B100')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);
$doc->getActiveSheet()->getStyle('C1:C100')->getAlignment()->setHorizontal(\PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="your_name.xls"');
header('Cache-Control: max-age=0');

$writer = \PHPExcel_IOFactory::createWriter($doc, 'Excel5');
$writer->save($uploadDir);