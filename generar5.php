<?php
/**
 * Creado: Jordan Diaz Diaz
 * Uso: Inscripcion y Reactivacion SUNAT -Persona Natural
 * Fecha: 16/07/2020
 */

//error_reporting(0);

require 'phpspreadsheet/vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\Spreadsheet;
$txtRazonSocial  = $_POST['txtRazonSocial'];
$txtDomicilioLegal = $_POST['txtDomicilioLegal'];
$txtDistrito = $_POST['txtDistrito'];
$txtProvincia= $_POST['txtProvincia'];
$txtDepartameno = $_POST['txtDepartameno'];
$txtNumeroDocumento = $_POST['txtNumeroDocumento'];
$txtTipoDocumento = $_POST['txtTipoDocumento'];
$txtRuc = $_POST['txtRuc'];
$txtTelefono = $_POST['txtTelefono'];
$txtCelular =  $_POST['txtCelular'];
$txtCorreo = $_POST['txtCorreo'];
$txtNombres =  $_POST['txtNombres'];
$txtDomicilioRepresentante =  $_POST['txtDomicilioRepresentante'];
$txtTipoDocumentoRepresentante = $_POST['txtTipoDocumentoRepresentante'];
$txtDniRepresentante =  $_POST['txtDniRepresentante'];
$txtPoderPartida  = $_POST['txtPoderPartida'];
$txtFechaPartida = $_POST['txtFechaPartida'];
$txtOficinaRegistral = $_POST['txtOficinaRegistral'];
$txtServicio = $_POST['txtServicio'];
$txtRecibo = $_POST['txtRecibo'];
$txtOperacion = $_POST['txtOperacion'];
$txtFechaPago = $_POST['txtFechaPago'];
//echo $txttipodomicilio;



$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("SERVICIO_DE_TRANSPORTE_TERRESTRE_DE_AMBITO_NACIONAL.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

$worksheet->setCellValue('F27', $txtRazonSocial);
$worksheet->setCellValue('F33', $txtDomicilioLegal);
$worksheet->setCellValue('F39', $txtDistrito);
$worksheet->setCellValue('AL39', $txtProvincia);
$worksheet->setCellValue('BR39', $txtDepartameno);
$dni = $txtNumeroDocumento;
$worksheet->setCellValue('F45', $dni[0]);
$worksheet->setCellValue('J45', $dni[1]);
$worksheet->setCellValue('N45', $dni[2]);
$worksheet->setCellValue('R45', $dni[3]);
$worksheet->setCellValue('V45', $dni[4]);
$worksheet->setCellValue('Z45', $dni[5]);
$worksheet->setCellValue('AD45', $dni[6]);
$worksheet->setCellValue('AH45', $dni[7]);
if ($txtTipoDocumento == "CE") {
    $worksheet->setCellValue('AV43', 'X');
}
if ($txtTipoDocumento == "CI") {
    $worksheet->setCellValue('BH43', 'X');
}
$ruc = $txtRuc;
$worksheet->setCellValue('BR45', $ruc[0]);
$worksheet->setCellValue('BU45', $ruc[1]);
$worksheet->setCellValue('BX45', $ruc[2]);
$worksheet->setCellValue('CA45', $ruc[3]);
$worksheet->setCellValue('CD45', $ruc[4]);
$worksheet->setCellValue('CG45', $ruc[5]);
$worksheet->setCellValue('CJ45', $ruc[6]);
$worksheet->setCellValue('CM45', $ruc[7]);
$worksheet->setCellValue('CP45', $ruc[8]);
$worksheet->setCellValue('CS45', $ruc[9]);

$worksheet->setCellValue('F51', $txtTelefono);
$worksheet->setCellValue('AL51', $txtCelular);
$worksheet->setCellValue('BR51', $txtCorreo);
$worksheet->setCellValue('CJ59', "X");
$worksheet->setCellValue('F65', $txtNombres);
$worksheet->setCellValue('F71', $txtDomicilioRepresentante);
$worksheet->setCellValue('BX71', $txtDniRepresentante);
if ($txtTipoDocumentoRepresentante == "CE") {
    $worksheet->setCellValue('CK69', 'X');
}
if ($txtTipoDocumentoRepresentante == "CI") {
    $worksheet->setCellValue('CR69', 'X');
}
if ($txtTipoDocumentoRepresentante == "DNI") {
    $worksheet->setCellValue('CC69', 'X');
}
$worksheet->setCellValue('T75', $txtPoderPartida);
$worksheet->setCellValue('AT75', $txtFechaPartida);
$worksheet->setCellValue('BZ75', $txtOficinaRegistral);

if($txtServicio=="DGTT-021"){
    $worksheet->setCellValue('AV84', 'X');
}
if($txtServicio=="DGTT-022"){
    $worksheet->setCellValue('AV88', 'X');
}
if($txtServicio=="DGTT-023"){
    $worksheet->setCellValue('AV92', 'X');
}
if($txtServicio=="DGTT-024"){
    $worksheet->setCellValue('AV96', 'X');
}
if($txtServicio=="DGTT-025"){
    $worksheet->setCellValue('CP84', 'X');
}
if($txtServicio=="DGTT-027"){
    $worksheet->setCellValue('CP88', 'X');
}
if($txtServicio=="DGTT-028"){
    $worksheet->setCellValue('CP92', 'X');
}
if($txtServicio=="DGTT-029"){
    $worksheet->setCellValue('CP96', 'X');
}

$worksheet->setCellValue('P108', $txtRecibo);
$worksheet->setCellValue('AX108', $txtOperacion);
$worksheet->setCellValue('CK108', $txtFechaPago);

$time = time();
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento.date("H-i-s", $time)."-SERVICIO_DE_TRANSPORTE_TERRESTRE_DE_AMBITO_NACIONAL.xlsx";
$writer->save($archivo);
echo '<a href="'.$archivo.'">Descargar SERVICIO_DE_TRANSPORTE_TERRESTRE_DE_AMBITO_NACIONAL_(PERSONAS Y MERCANCÍAS)</a>';
?>
