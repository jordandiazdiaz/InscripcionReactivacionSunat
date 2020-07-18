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
$txtNumeroSocios  = $_POST['txtNumeroSocios'];
$txtTipoContribuyente = $_POST['txtTipoContribuyente'];
$txtDenominacion = $_POST['txtDenominacion'];
$txtPartidaRegistral = $_POST['txtPartidaRegistral'];
$txtZonaRegistral = $_POST['txtZonaRegistral'];
$txtOficinaRegistral = $_POST['txtOficinaRegistral'];
$txtTipoRepresentacion = $_POST['txtTipoRepresentacion'];
$txtOrigenCapital = $_POST['txtOrigenCapital'];
$txtPais = $_POST['txtPais'];
$txtDomicilioFiscal = $_POST['txtDomicilioFiscal'];
$txtDistrito = $_POST['txtDistrito'];
$txtProvincia = $_POST['txtProvincia'];
$txtDepartameno = $_POST['txtDepartameno'];
$txttipodomicilio = $_POST['txttipodomicilio'];
$txtCorreo = $_POST['txtCorreo'];
$txtTelefonoMovil = $_POST['txtTelefonoMovil'];
$txtActividadEconomica = $_POST['txtActividadEconomica'];
$txtRegimenGeneral = $_POST['txtRegimenGeneral'];
$txtRegimenEspecial = $_POST['txtRegimenEspecial'];
$txtRegimenMype = $_POST['txtRegimenMype'];
$txtOtrosTributos = $_POST['txtOtrosTributos'];
$txtTipoDocumento = $_POST['txtTipoDocumento'];
$txtNumeroDocumento = $_POST['txtNumeroDocumento'];
$txtPrimerApellido = $_POST['txtPrimerApellido'];
$txtSegundoApellido = $_POST['txtSegundoApellido'];
$txtNombres = $_POST['txtNombres'];
$txtFechaNacimiento = $_POST['txtFechaNacimiento'];
$txtRazonSocial = $_POST['txtRazonSocial'];
$txtTipoCargo = $_POST['txtTipoCargo'];
$txtDomicilioRespresentante = $_POST['txtDomicilioRespresentante'];
$txtDistritoRepresentante = $_POST['txtDistritoRepresentante'];
$txtProvinciaRepresentante = $_POST['txtProvinciaRepresentante'];
$txtDepartamentoRepresentante = $_POST['txtDepartamentoRepresentante'];
$txttipodomicilio = $_POST['txttipodomicilio'];
$txtCorreoRepresentante = $_POST['txtCorreoRepresentante'];
$txtTelefonoRepresentante = $_POST['txtTelefonoRepresentante'];
$txtTelefonoRepresentante = $_POST['txtTelefonoRepresentante'];
$txtFechaInicioCargo = $_POST['txtFechaInicioCargo'];
$txtCondicionDomicilio = $_POST['txtCondicionDomicilio'];

if($txtNumeroSocios=="1"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_01.xlsx");
}
if($txtNumeroSocios=="2"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_02.xlsx");
}
if($txtNumeroSocios=="3"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_03.xlsx");
}
if($txtNumeroSocios=="4"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_04.xlsx");
}
if($txtNumeroSocios=="5"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_05.xlsx");
}
if($txtNumeroSocios=="6"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_06.xlsx");
}
if($txtNumeroSocios=="7"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_07.xlsx");
}
if($txtNumeroSocios=="8"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_08.xlsx");
}
if($txtNumeroSocios=="9"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_09.xlsx");
}
if($txtNumeroSocios=="10"){
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA_PERSONA_JURIDICA_14052020_10.xlsx");
}
setcookie("NumeroSocios",$txtNumeroSocios);


$worksheet = $spreadsheet->getActiveSheet();

if ($txtinsreact=="INSCRIPCION") {
    $worksheet->setCellValue('F17', 'X');

}
else
{
    $worksheet->setCellValue('T17', 'X');

}


 $worksheet->setCellValue('AF33', $txtTipoContribuyente);
 $worksheet->setCellValue('F42', $txtDenominacion);
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
$fechanacimiento = $txtFechaNacimiento;
$fecha_nacimiento = explode('/', $fechanacimiento);


$worksheet->setCellValue('AY112', $dni[0]);
$worksheet->setCellValue('BB112', $dni[1]);
$worksheet->setCellValue('BE112', $dni[2]);
$worksheet->setCellValue('BH112', $dni[3]);
$worksheet->setCellValue('BK112', $dni[4]);
$worksheet->setCellValue('BN112', $dni[5]);
$worksheet->setCellValue('BQ112', $dni[6]);
$worksheet->setCellValue('BT112', $dni[7]);
$worksheet->setCellValue('AT117', $txtPrimerApellido);
$worksheet->setCellValue('BV117', $txtSegundoApellido);
$worksheet->setCellValue('AT122', $txtNombres);
$worksheet->setCellValue('CB125',$fecha_nacimiento[0]);
$worksheet->setCellValue('CI125',  $fecha_nacimiento[1]);
$worksheet->setCellValue('CP125',$fecha_nacimiento[2]);
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

//Variable dni ira a una variable de session para poder ser usado en el segundo formulario
setcookie("NumeroDocumento",$txtNumeroDocumento);

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento."-GUIA_PERSONA_JURIDICA_14052020_01.xlsx";
$writer->save($archivo);

echo '<a href="formulario2-1.html">Continuar con el Siguiente Formulario</a>';
?>
