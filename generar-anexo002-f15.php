<?php
error_reporting(0);

require 'phpspreadsheet/vendor/autoload.php';

use \ConvertApi\ConvertApi;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

$txtNombres = $_POST['nombres'];
$txtCargo = $_POST['cargo'];
$txtDni = $_POST['dni'];

$txtDia = $_POST['dia'];
$txtMes = $_POST['mes'];
$txtAnio = $_POST['anio'];
$txtSolicitanteNombres = $_POST['solicita_nombres'];
$txtSolicitanteDni = $_POST['solicita_dni'];

$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("SERVICIOS_DE_TRANSPORTE_DE_MERCANCIAS_EN_GENERAL.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

$rowNumber = 50;

for ($i = 0; $i < count($txtNombres); $i++) { 
    $worksheet->setCellValue('H' . $rowNumber, $txtNombres[$i]);
    $worksheet->setCellValue('BD' . $rowNumber, $txtCargo[$i]);
    $worksheet->setCellValue('BS' . $rowNumber, $txtDni[$i]);
    $rowNumber = $rowNumber + 6;
}

$worksheet->setCellValue('M96', $txtDia);
$worksheet->setCellValue('W96', $txtMes);
$worksheet->setCellValue('AS96', $txtAnio);


//$worksheet->setCellValue('BM147', $txtSolicita_nombres);
//$worksheet->setCellValue('BD150', $txtSolicita_dni);

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = "-SERVICIOS_DE_TRANSPORTE_DE_MERCANCIAS_EN_GENERAL.xlsx";
$writer->save($archivo);
echo '<a href="'.$archivo.'">Descargar SERVICIOS DE TRANSPORTE DE MERCANCIAS EN GENERAL</a>';
?>
