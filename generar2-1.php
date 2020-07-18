<?php
/**
 * Creado: Jordan Diaz Diaz
 * Uso: Inscripcion y Reactivacion SUNAT -Persona Natural
 * Fecha: 16/07/2020
 */

error_reporting(0);

require 'phpspreadsheet/vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\Spreadsheet;
$txtTipoDocumento  = $_POST['txtTipoDocumento'];
$txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
$txtPrimerApellido  = $_POST['txtPrimerApellido'];
$txtSegundoApellido  = $_POST['txtSegundoApellido'];
$txtNombres = $_POST['txtNombres'];
$txtFechaNacimiento = $_POST['txtFechaNacimiento'];
$txtRazonSocial = $_POST['txtRazonSocial'];
$txtPaisRecidencia = $_POST['txtPaisRecidencia'];
$txtTipoVehiculo  = $_POST['txtTipoVehiculo'];
$txtPorcentajeParticipacion  = $_POST['txtPorcentajeParticipacion'];
$txtTipoEstablecimiento  = $_POST['txtTipoEstablecimiento'];
$txtDistrito  = $_POST['txtDistrito'];
$txtProvincia  = $_POST['txtProvincia'];
$txtDepartameno  = $_POST['txtDepartameno'];
$txttipodomicilio  = $_POST['txttipodomicilio'];
$txtDireccionEstablecimiento = $_POST['txtDireccionEstablecimiento'];

$txtNumeroDocumento_Anterior = $_COOKIE["NumeroDocumento"];
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($txtNumeroDocumento_Anterior."-GUIA_PERSONA_JURIDICA_14052020_01.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

$worksheet->setCellValue('EQ12', "X");
$worksheet->setCellValue('EU11', $txtNumeroDocumento[0]);
$worksheet->setCellValue('EZ11', $txtNumeroDocumento[1]);
$worksheet->setCellValue('FC11', $txtNumeroDocumento[2]);
$worksheet->setCellValue('FF11', $txtNumeroDocumento[3]);
$worksheet->setCellValue('FI11', $txtNumeroDocumento[4]);
$worksheet->setCellValue('FL11', $txtNumeroDocumento[5]);
$worksheet->setCellValue('FO11', $txtNumeroDocumento[6]);
$worksheet->setCellValue('FR11', $txtNumeroDocumento[7]);
$worksheet->setCellValue('EU16', $txtPrimerApellido);
$worksheet->setCellValue('FW16', $txtSegundoApellido);
$worksheet->setCellValue('EU21', $txtNombres);
$fechanacimiento = $txtFechaNacimiento;
$fecha_nacimiento = explode('/', $fechanacimiento);
$worksheet->setCellValue('FJ24',$fecha_nacimiento[0]);
$worksheet->setCellValue('FN24',  $fecha_nacimiento[1]);
$worksheet->setCellValue('FR24',$fecha_nacimiento[2]);
$worksheet->setCellValue('GE24', $txtPaisRecidencia);
$worksheet->setCellValue('EM28', $txtRazonSocial);
$worksheet->setCellValue('EC31', $txtTipoVehiculo);
$time = time();
$worksheet->setCellValue('GL31', date("d", $time));
$worksheet->setCellValue('GP31', date("m", $time));
$worksheet->setCellValue('GT31', date("Y", $time));
$worksheet->setCellValue('EB36', $txtPorcentajeParticipacion);

if($txtTipoEstablecimiento=="CASA MATRIZ"){
    $worksheet->setCellValue('EB46', "X");
}
if($txtTipoEstablecimiento=="AGENCIA"){
    $worksheet->setCellValue('EB50', "X");
}
if($txtTipoEstablecimiento=="DEPOSITO O ALMACEN"){
    $worksheet->setCellValue('EY46', "X");
}
if($txtTipoEstablecimiento=="SUCURSAL"){
    $worksheet->setCellValue('EY50', "X");
}
if($txtTipoEstablecimiento=="SEDE PRODUCTIVA"){
    $worksheet->setCellValue('FV46', "X");
}
if($txtTipoEstablecimiento=="OFICIA ADMINISTRATIVA"){
    $worksheet->setCellValue('FV50', "X");
}
if($txtTipoEstablecimiento=="LOCAL COMERCIAL O DE SERVICIOS"){
    $worksheet->setCellValue('GS48', "X");
}
$worksheet->setCellValue('EG54', $txtDireccionEstablecimiento);
$worksheet->setCellValue('DQ57', $txtDistrito);
$worksheet->setCellValue('EW57', $txtProvincia);
$worksheet->setCellValue('GE57', $txtDepartameno);
if($txttipodomicilio=="PROPIO"){
    $worksheet->setCellValue('EJ61', "X");
}
if($txttipodomicilio=="ALQUILADO"){
    $worksheet->setCellValue('EY61', "X");
}
if($txttipodomicilio=="CEDIDO"){
    $worksheet->setCellValue('FP61', "X");
}
if($txttipodomicilio=="OTROS"){
    $worksheet->setCellValue('GI61', "X");
}
setcookie("NumeroDocumento2",$txtNumeroDocumento_Anterior);
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento_Anterior."-2-GUIA_PERSONA_JURIDICA_14052020_01.xlsx";
$writer->save($archivo);
echo '<a href="formulario2-2.html">Continuar con el Siguiente Formulario</a>';
?>
