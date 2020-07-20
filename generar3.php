<?php
/**
 * Creado: Jordan Diaz Diaz
 * Uso: Inscripcion y Reactivacion SUNAT -Persona Natural
 * Fecha: 16/07/2020
 */

error_reporting(0);

require 'phpspreadsheet/vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\Spreadsheet;
$txtinsreact  = $_POST['txtinsreact'];
$txtTipoDocumento = $_POST['txtTipoDocumento'];
$txtNumeroDocumento = $_POST['txtNumeroDocumento'];
$txtPrimerApellido= $_POST['txtPrimerApellido'];
$txtSegundoApellido = $_POST['txtSegundoApellido'];
$txtNombres = $_POST['txtNombres'];
$txtFechaNacimiento = $_POST['txtFechaNacimiento'];
$txtCorreo = $_POST['txtCorreo'];
$txtCorreoRepetido = $_POST['txtCorreoRepetido'];
$txttipodomicilio =  $_POST['txttipodomicilio'];
$txtDomicilioFiscal = $_POST['txtDomicilioFiscal'];
$txtDistrito =  $_POST['txtDistrito'];
$txtProvincia =  $_POST['txtProvincia'];
$txtDepartameno = $_POST['txtDepartameno'];
$txtTelefonoMovil =  $_POST['txtTelefonoMovil'];
$txtActividadEconomica  = $_POST['txtActividadEconomica'];
$txtPaisNacionalidad = $_POST['txtPaisNacionalidad'];
$txtSexo = $_POST['txtSexo'];
$txtTributosAfectos = $_POST['txtTributosAfectos'];
$txtOtros = $_POST['txtOtros'];
//echo $txttipodomicilio;



$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_CON_NEGOCIO_1405202 01.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

if ($txtinsreact=="INSCRIPCION") {
    $worksheet->setCellValue('F12', 'X');

}
else
{
    $worksheet->setCellValue('T12', 'X');

}


$worksheet->setCellValue('H38', 'X');

if ($txtTipoDocumento == "3") {
    $worksheet->setCellValue('BH38', 'X');
}
if ($txtTipoDocumento == "1") {
    $worksheet->setCellValue('BH35', 'X');
}
if ($txtTipoDocumento == "187") {
    $worksheet->setCellValue('BH41', 'X');
}
$dni = $txtNumeroDocumento;
$fechanacimiento = $txtFechaNacimiento;
$fecha_nacimiento = explode('/', $fechanacimiento);


$worksheet->setCellValue('BM40', $dni[0]);
$worksheet->setCellValue('BP40', $dni[1]);
$worksheet->setCellValue('BS40', $dni[2]);
$worksheet->setCellValue('BV40', $dni[3]);
$worksheet->setCellValue('BY40', $dni[4]);
$worksheet->setCellValue('CB40', $dni[5]);
$worksheet->setCellValue('CE40', $dni[6]);
$worksheet->setCellValue('CH40', $dni[7]);
$worksheet->setCellValue('F53', $txtPrimerApellido);
$worksheet->setCellValue('BC53', $txtSegundoApellido);
$worksheet->setCellValue('F58', $txtNombres);
$worksheet->setCellValue('J65',$fecha_nacimiento[0]);
$worksheet->setCellValue('R65',  $fecha_nacimiento[1]);
$worksheet->setCellValue('Z65',$fecha_nacimiento[2]);
$sexo = $txtSexo;

if ($sexo == "MASCULINO") {
    $worksheet->setCellValue('BB64', 'X');
}
else
{
    $worksheet->setCellValue('AN64', 'X');
}
$worksheet->setCellValue('BR63', $txtPaisNacionalidad);
$worksheet->setCellValue('G70', $txtDomicilioFiscal);
$worksheet->setCellValue('G75', $txtDistrito);
$worksheet->setCellValue('AL75', $txtProvincia);
$worksheet->setCellValue('BR75', $txtDepartameno);

if ($txttipodomicilio == "PROPIO") {
    $worksheet->setCellValue('J82', 'X');
}
if ($txttipodomicilio == "ALQUILADO") {
    $worksheet->setCellValue('S82', 'X');
}
if ($txttipodomicilio == "CEDIDO") {
    $worksheet->setCellValue('AE82', 'X');
}
if ($txttipodomicilio == "OTROS") {
    $worksheet->setCellValue('AS82', 'X');
}


$worksheet->setCellValue('BD81',$txtCorreo);
$worksheet->setCellValue('H89', $txtTelefonoMovil[0]);
$worksheet->setCellValue('M89', $txtTelefonoMovil[1]);
$worksheet->setCellValue('R89', $txtTelefonoMovil[2]);
$worksheet->setCellValue('W89', $txtTelefonoMovil[3]);
$worksheet->setCellValue('AB89', $txtTelefonoMovil[4]);
$worksheet->setCellValue('AG89', $txtTelefonoMovil[5]);
$worksheet->setCellValue('AL89', $txtTelefonoMovil[6]);
$worksheet->setCellValue('AQ89', $txtTelefonoMovil[7]);
$worksheet->setCellValue('AV89', $txtTelefonoMovil[8]);
$worksheet->setCellValue('BD88', $txtActividadEconomica);
$worksheet->setCellValue('AG96', 'X');
$worksheet->setCellValue('BQ96', 'X');
$time = time();
$worksheet->setCellValue('AH100', date("d", $time));
$worksheet->setCellValue('AO100', date("m", $time));
$worksheet->setCellValue('AV100', date("Y", $time));
if($txtTributosAfectos=="NUEVO RUS - NRUS"){
    $worksheet->setCellValue('BD106', 'X');
}
if($txtTributosAfectos=="REGIMEN ESPECIAL DE RENTA -RER"){
    $worksheet->setCellValue('BD110', 'X');
}
if($txtTributosAfectos=="REGIMEN MYPE TRIBUTARIO"){
    $worksheet->setCellValue('CS106', 'X');
}
if($txtTributosAfectos=="REGIMEN GENERA"){
    $worksheet->setCellValue('CS110', "X");
}
if($txtOtros==""){
    $worksheet->setCellValue('AR113',"");
}
else
{
    $worksheet->setCellValue('AR113',$txtOtros);
}


$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento.date("H-i-s", $time)."-GUIA_PERSONA_CON_NEGOCIO_14052020.xlsx";
$writer->save($archivo);
echo '<a href="'.$archivo.'">Descargar GUIA_PERSONA_CON_NEGOCIO_14052020</a>';
?>
