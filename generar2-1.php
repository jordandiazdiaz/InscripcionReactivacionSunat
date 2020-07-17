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



$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_01.xlsx");
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




 
 $worksheet->setCellValue('F47', $txtPartidaRegistral);
 $worksheet->setCellValue('AL47', $txtZonaRegistral);
 $worksheet->setCellValue('BR47', $txtOficinaRegistral);
if($txtTipoRepresentacion=="INDISTINTA"){
    $worksheet->setCellValue('AU51', "X");
}
if($txtTipoRepresentacion=="CONJUNTA"){
    $worksheet->setCellValue('BN51', "X");
}
if($txtTipoRepresentacion=="SUCESIVA"){
    $worksheet->setCellValue('CF51', "X");
}
if($txtOrigenCapital == "NACIONAL"){
    $worksheet->setCellValue('AU56', "X");
}
if($txtOrigenCapital == "EXTRANJERO"){
    $worksheet->setCellValue('BN56', "X");
}
$worksheet->setCellValue('CI56', $txtPais);
$worksheet->setCellValue('G62', $txtDomicilioFiscal);
$worksheet->setCellValue('G67', $txtDistrito);
$worksheet->setCellValue('AL67', $txtProvincia);
$worksheet->setCellValue('BR67', $txtDepartameno);
if($txttipodomicilio=="PROPIO"){
    $worksheet->setCellValue('H75', "X");
}
if($txttipodomicilio=="ALQUILADO"){
    $worksheet->setCellValue('R75', "X");
}
if($txttipodomicilio=="CEDIDO"){
    $worksheet->setCellValue('AD75', "X");
}
if($txttipodomicilio=="OTROS"){
    $worksheet->setCellValue('AS75', "X");
}

$worksheet->setCellValue('BD74', $txtCorreo);
$worksheet->setCellValue('H82', $txtTelefonoMovil[0]);
$worksheet->setCellValue('M82', $txtTelefonoMovil[1]);
$worksheet->setCellValue('R82', $txtTelefonoMovil[2]);
$worksheet->setCellValue('W82', $txtTelefonoMovil[3]);
$worksheet->setCellValue('AB82', $txtTelefonoMovil[4]);
$worksheet->setCellValue('AG82', $txtTelefonoMovil[5]);
$worksheet->setCellValue('AL82', $txtTelefonoMovil[6]);
$worksheet->setCellValue('AQ82', $txtTelefonoMovil[7]);
$worksheet->setCellValue('AV82', $txtTelefonoMovil[8]);

$worksheet->setCellValue('BD81', $txtActividadEconomica);
$worksheet->setCellValue('AG89',"X");
$worksheet->setCellValue('BQ89',"X");
$time = time();
$worksheet->setCellValue('AH93', date("d", $time));
$worksheet->setCellValue('AO93', date("m", $time));
$worksheet->setCellValue('AV93', date("Y", $time));

if($txtRegimenGeneral =="REGIMEN GENERAL"){
    $worksheet->setCellValue('BD99', "X");
}
if($txtRegimenMype =="REGIMEN MYPE TRIBUTARIO -RMT"){
    $worksheet->setCellValue('CS99', "X");
}
if($txtRegimenEspecial =="REGIMEN ESPECIAL DE RENTA -RER"){
    $worksheet->setCellValue('BD103', "X");
}
$worksheet->setCellValue('BY102', $txtOtrosTributos);
$worksheet->setCellValue('AP113', "X");

$dni = $txtNumeroDocumento;

$worksheet->setCellValue('AL129',$txtRazonSocial);
$worksheet->setCellValue('AF132',$txtTipoCargo);
$fecha_cargo = explode("/",$txtFechaInicioCarg);
$worksheet->setCellValue('CK132',$fecha_cargo[0]);
$worksheet->setCellValue('CO132',$fecha_cargo[1]);
$worksheet->setCellValue('CS132',$fecha_cargo[2]);
$worksheet->setCellValue('P137',$txtDomicilioRespresentante);
$worksheet->setCellValue('O140',$txtDistritoRepresentante);
$worksheet->setCellValue('AU140',$txtProvinciaRepresentante);
$worksheet->setCellValue('CC140',$txtDepartamentoRepresentante);

if($txtCondicionDomicilio=="PROPIO"){
    $worksheet->setCellValue('AM144',"X");
}
if($txtCondicionDomicilio=="ALQUILADO"){
    $worksheet->setCellValue('BB144',"X");
}
if($txtCondicionDomicilio=="CEDIDO"){
    $worksheet->setCellValue('BS144',"X");
}
if($txtCondicionDomicilio=="OTROS"){
    $worksheet->setCellValue('CL144',"X");
}

$worksheet->setCellValue('W148',$txtCorreoRepresentante);
$worksheet->setCellValue('CE148', $txtTelefonoRepresentante[0]);
$worksheet->setCellValue('CG148', $txtTelefonoRepresentante[1]);
$worksheet->setCellValue('CI148', $txtTelefonoRepresentante[2]);
$worksheet->setCellValue('CK148', $txtTelefonoRepresentante[3]);
$worksheet->setCellValue('CM148', $txtTelefonoRepresentante[4]);
$worksheet->setCellValue('CO148', $txtTelefonoRepresentante[5]);
$worksheet->setCellValue('CQ148', $txtTelefonoRepresentante[6]);
$worksheet->setCellValue('CS148', $txtTelefonoRepresentante[7]);
$worksheet->setCellValue('CU148', $txtTelefonoRepresentante[8]);


$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento."-GUIA_PERSONA_JURIDICA_14052020_01.xlsx";
$writer->save($archivo);
//cho '<a href="'.$archivo.'">Descargar GUIA_PERSONA_SIN_NEGOCIO_14052020</a>';
?>
