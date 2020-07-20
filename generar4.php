<?php
/**
 * Creado: Jordan Diaz Diaz
 * Uso: Inscripcion y Reactivacion SUNAT -Persona Natural
 * Fecha: 16/07/2020
 */

error_reporting(0);

require 'phpspreadsheet/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

$txtNumeroDocumento = $_POST['txtNumeroDocumento'];
$txtPrimerApellido= $_POST['txtPrimerApellido'];
$txtSegundoApellido = $_POST['txtSegundoApellido'];
$txtNombres = $_POST['txtNombres'];
$txtCalidad = $_POST['txtCalidad'];
$txtEmpresa = $_POST['txtEmpresa'];
$txtDomicilio = $_POST['txtDomicilio'];
$txtPartidaRegistral = $_POST['txtPartidaRegistral'];


$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("SERVICIOS_DE_TRANSPORTE_TERRESTR_ DE_MERCANCIAS.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

$nombre_completo = $txtPrimerApellido." ".$txtSegundoApellido." ".$txtNombres;
$worksheet->setCellValue('K38', $nombre_completo);
$worksheet->setCellValue('Q42', $txtNumeroDocumento);
$worksheet->setCellValue('AT50', $txtDomicilio);
$worksheet->setCellValue('AZ42', $txtCalidad);
$worksheet->setCellValue('H46', $txtEmpresa);
$worksheet->setCellValue('H50', $txtPartidaRegistral);
$time = time();
$meses_ES = array("Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre");
$meses_EN = array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December");
$nombreMes = str_replace($meses_EN, $meses_ES, date("F", $time));
$worksheet->setCellValue('M107', date("d", $time));
$worksheet->setCellValue('W107', $nombreMes );
$worksheet->setCellValue('AS107', date("Y", $time));

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento.date("H-i-s", $time)."-SERVICIOS_DE_TRANSPORTE_TERRESTR_ DE_MERCANCIAS.xlsx";
$writer->save($archivo);
echo '<a href="'.$archivo.'">Descargar SERVICIOS_DE_TRANSPORTE_TERRESTR_ DE_MERCANCÍAS.xlsx</a>';
?>
