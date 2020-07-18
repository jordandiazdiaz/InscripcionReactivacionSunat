<?php
/**
 * Creado: Jordan Diaz Diaz
 * Uso: Inscripcion y Reactivacion SUNAT -Persona Natural
 * Fecha: 16/07/2020
 */

//error_reporting(0);

require 'phpspreadsheet/vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\Spreadsheet;
$txtTipoDocumento  = $_POST['txtTipoDocumento'];
$txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
$txtPrimerApellido  = $_POST['txtPrimerApellido'];
$txtSegundoApellido  = $_POST['txtSegundoApellido'];
$txtNombres = $_POST['txtNombres'];
$txtFechaNacimiento = $_POST['txtFechaNacimiento'];
$txtRazonSocial = $_POST['txtRazonSocial'];
$txtTipoCargo  = $_POST['txtTipoCargo'];
$txtDomicilio = $_POST['txtDomicilio'];
$txtDistrito  = $_POST['txtDistrito'];
$txtProvincia  = $_POST['txtProvincia'];
$txtDepartameno  = $_POST['txtDepartameno'];
$txttipodomicilio  = $_POST['txttipodomicilio'];
$txtCorreo  = $_POST['txtCorreo'];
$txtTelefonoMovil  = $_POST['txtTelefonoMovil'];

$txtNumeroDocumento_Anterior = $_COOKIE["NumeroDocumento2"];
print "Archivo: ".$txtNumeroDocumento_Anterior;
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($txtNumeroDocumento_Anterior."-2-GUIA_PERSONA_JURIDICA_14052020_01.xlsx");
$worksheet = $spreadsheet->getActiveSheet();
if($txtTipoDocumento=="Carnet de Extranjeria"){
    $worksheet->setCellValue('IR17', "X");
}
if($txtTipoDocumento=="Documento Nacional de Identidad"){
    $worksheet->setCellValue('IR15', "X");
}
if($txtTipoDocumento=="Pasaporte"){
    $worksheet->setCellValue('IR19', "X");
}

$worksheet->setCellValue('JA14', $txtNumeroDocumento[0]);
$worksheet->setCellValue('JD14', $txtNumeroDocumento[1]);
$worksheet->setCellValue('JG14', $txtNumeroDocumento[2]);
$worksheet->setCellValue('JJ14', $txtNumeroDocumento[3]);
$worksheet->setCellValue('JM14', $txtNumeroDocumento[4]);
$worksheet->setCellValue('JP14', $txtNumeroDocumento[5]);
$worksheet->setCellValue('JS14', $txtNumeroDocumento[6]);
$worksheet->setCellValue('JV14', $txtNumeroDocumento[7]);
$worksheet->setCellValue('IV19', $txtPrimerApellido);
$worksheet->setCellValue('JX19', $txtSegundoApellido);
$worksheet->setCellValue('IV24', $txtNombres);
$fechanacimiento = $txtFechaNacimiento;
$fecha_nacimiento = explode('/', $fechanacimiento);
$worksheet->setCellValue('KD27',$fecha_nacimiento[0]);
$worksheet->setCellValue('KK27',  $fecha_nacimiento[1]);
$worksheet->setCellValue('KR27',$fecha_nacimiento[2]);
$worksheet->setCellValue('IN31', $txtRazonSocial);
$worksheet->setCellValue('IH34', $txtTipoCargo);
$time = time();
$worksheet->setCellValue('KD27', date("d", $time));
$worksheet->setCellValue('KK27', date("m", $time));
$worksheet->setCellValue('KR27', date("Y", $time));
if($txttipodomicilio=="PROPIO"){
    $worksheet->setCellValue('IO46', "X");
}
if($txttipodomicilio=="ALQUILADO"){
    $worksheet->setCellValue('JD46', "X");
}
if($txttipodomicilio=="CEDIDO"){
    $worksheet->setCellValue('JU46', "X");
}
if($txttipodomicilio=="OTROS"){
    $worksheet->setCellValue('KN46', "X");
}
$worksheet->setCellValue('HR39', $txtDomicilio);
$worksheet->setCellValue('HQ42', $txtDistrito);
$worksheet->setCellValue('IW42', $txtProvincia);
$worksheet->setCellValue('KE42', $txtDepartameno);
$worksheet->setCellValue('HY50', $txtCorreo);
$worksheet->setCellValue('KG50', $txtTelefonoMovil[0]);
$worksheet->setCellValue('KI50', $txtTelefonoMovil[1]);
$worksheet->setCellValue('KK50', $txtTelefonoMovil[2]);
$worksheet->setCellValue('KM50', $txtTelefonoMovil[3]);
$worksheet->setCellValue('KO50', $txtTelefonoMovil[4]);
$worksheet->setCellValue('KQ50', $txtTelefonoMovil[5]);
$worksheet->setCellValue('KS50', $txtTelefonoMovil[6]);
$worksheet->setCellValue('KU50', $txtTelefonoMovil[7]);
$worksheet->setCellValue('KW50', $txtTelefonoMovil[8]);

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
$txtTipoDocumento2  = $_POST['txtTipoDocumento2'];
$txtNumeroDocumento2  = $_POST['txtNumeroDocumento2'];
$txtPrimerApellido2  = $_POST['txtPrimerApellido2'];
$txtSegundoApellido2  = $_POST['txtSegundoApellido2'];
$txtNombres2 = $_POST['txtNombres2'];
$txtFechaNacimiento2 = $_POST['txtFechaNacimiento2'];
$txtRazonSocial2 = $_POST['txtRazonSocial2'];
$txtTipoCargo2  = $_POST['txtTipoCargo2'];
$txtDomicilio2 = $_POST['txtDomicilio2'];
$txtDistrito2 = $_POST['txtDistrito2'];
$txtProvincia2 = $_POST['txtProvincia2'];
$txtDepartameno2  = $_POST['txtDepartameno2'];
$txttipodomicilio2  = $_POST['txttipodomicilio2'];
$txtCorreo2  = $_POST['txtCorreo2'];
$txtTelefonoMovil2  = $_POST['txtTelefonoMovil2'];

if($txtTipoDocumento2=="Carnet de Extranjeria"){
    $worksheet->setCellValue('IR64', "X");
}
if($txtTipoDocumento2=="Documento Nacional de Identidad"){
    $worksheet->setCellValue('IR62', "X");
}
if($txtTipoDocumento2=="Pasaporte"){
    $worksheet->setCellValue('IR66', "X");
}

$worksheet->setCellValue('JA61', $txtNumeroDocumento2[0]);
$worksheet->setCellValue('JD61', $txtNumeroDocumento2[1]);
$worksheet->setCellValue('JG61', $txtNumeroDocumento2[2]);
$worksheet->setCellValue('JJ61', $txtNumeroDocumento2[3]);
$worksheet->setCellValue('JM61', $txtNumeroDocumento2[4]);
$worksheet->setCellValue('JP61', $txtNumeroDocumento2[5]);
$worksheet->setCellValue('JS61', $txtNumeroDocumento2[6]);
$worksheet->setCellValue('JV61', $txtNumeroDocumento2[7]);
$worksheet->setCellValue('IV66', $txtPrimerApellido2);
$worksheet->setCellValue('JX66', $txtSegundoApellido2);
$worksheet->setCellValue('IV71', $txtNombres2);
$fechanacimiento2 = $txtFechaNacimiento2;
$fecha_nacimiento2 = explode('/', $fechanacimiento2);
$worksheet->setCellValue('KD74',$fecha_nacimiento2[0]);
$worksheet->setCellValue('KK74',  $fecha_nacimiento2[1]);
$worksheet->setCellValue('KR74',$fecha_nacimiento2[2]);
$worksheet->setCellValue('IN78', $txtRazonSocial2);
$worksheet->setCellValue('IH81', $txtTipoCargo2);
$time2 = time();
$worksheet->setCellValue('KD27', date("d", $time2));
$worksheet->setCellValue('KK27', date("m", $time2));
$worksheet->setCellValue('KR27', date("Y", $time2));
if($txttipodomicilio2=="PROPIO"){
    $worksheet->setCellValue('IO93', "X");
}
if($txttipodomicilio2=="ALQUILADO"){
    $worksheet->setCellValue('JD93', "X");
}
if($txttipodomicilio2=="CEDIDO"){
    $worksheet->setCellValue('JU93', "X");
}
if($txttipodomicilio2=="OTROS"){
    $worksheet->setCellValue('KN93', "X");
}
$worksheet->setCellValue('HR86', $txtDomicilio2);
$worksheet->setCellValue('HQ89', $txtDistrito2);
$worksheet->setCellValue('IW89', $txtProvincia2);
$worksheet->setCellValue('KE89', $txtDepartameno2);
$worksheet->setCellValue('HY97', $txtCorreo2);
$worksheet->setCellValue('KG97', $txtTelefonoMovil2[0]);
$worksheet->setCellValue('KI97', $txtTelefonoMovil2[1]);
$worksheet->setCellValue('KK97', $txtTelefonoMovil2[2]);
$worksheet->setCellValue('KM97', $txtTelefonoMovil2[3]);
$worksheet->setCellValue('KO97', $txtTelefonoMovil2[4]);
$worksheet->setCellValue('KQ97', $txtTelefonoMovil2[5]);
$worksheet->setCellValue('KS97', $txtTelefonoMovil2[6]);
$worksheet->setCellValue('KU97', $txtTelefonoMovil2[7]);
$worksheet->setCellValue('KW97', $txtTelefonoMovil2[8]);


$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento_Anterior."-3-GUIA_PERSONA_JURIDICA_14052020_01.xlsx";
$writer->save($archivo);
echo '<a href="'.$archivo.'">Descargar GUIA_PERSONA_SIN_NEGOCIO_14052020</a>';
?>
