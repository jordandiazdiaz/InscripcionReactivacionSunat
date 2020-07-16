<?php
/**
 * Creado: Jordan Diaz Diaz
 * Uso: Inscripcion y Reactivacion SUNAT -Persona Natural
 * Fecha: 16/07/1988
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
//echo $txttipodomicilio;



$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("GUIA.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

if ($txtinsreact=="INSCRIPCION") {
    $worksheet->setCellValue('G17', 'X');

}
else
{
    $worksheet->setCellValue('U17', 'X');

}


$worksheet->setCellValue('I43', 'X');

if ($txtTipoDocumento == "3") {
    $worksheet->setCellValue('BI43', 'X');
}
if ($txtTipoDocumento == "1") {
    $worksheet->setCellValue('BI40', 'X');
}
if ($txtTipoDocumento == "187") {
    $worksheet->setCellValue('BI43', 'X');
}
$dni = $txtNumeroDocumento;
$fechanacimiento = $txtFechaNacimiento;
$date =$fechanacimiento;
list($month, $day, $year) = split('[/.-]', $date);


$worksheet->setCellValue('BN45', $dni[0]);
$worksheet->setCellValue('BQ45', $dni[1]);
$worksheet->setCellValue('BT45', $dni[2]);
$worksheet->setCellValue('BW45', $dni[3]);
$worksheet->setCellValue('BZ45', $dni[4]);
$worksheet->setCellValue('CC45', $dni[5]);
$worksheet->setCellValue('CF45', $dni[6]);
$worksheet->setCellValue('CI45', $dni[7]);
$worksheet->setCellValue('G58', $txtPrimerApellido);
$worksheet->setCellValue('BD58', $txtSegundoApellido);
$worksheet->setCellValue('G63', $txtNombres);
$worksheet->setCellValue('G69',$month);
$worksheet->setCellValue('Q69',  $day);
$worksheet->setCellValue('AA69',$year);
$sexo = $txtSexo;

if ($sexo == "MASCULINO") {
    $worksheet->setCellValue('AO69', 'X');
}
else
{
    $worksheet->setCellValue('BC69', 'X');
}
$worksheet->setCellValue('BS68', $txtPaisNacionalidad);
$worksheet->setCellValue('H75', $txtDomicilioFiscal);
$worksheet->setCellValue('H80', $txtDistrito);
$worksheet->setCellValue('AM80', $txtProvincia);
$worksheet->setCellValue('BS80', $txtDepartameno);

if ($txttipodomicilio == "PROPIO") {
    $worksheet->setCellValue('K87', 'X');
}
if ($txttipodomicilio == "ALQUILADO") {
    $worksheet->setCellValue('T87', 'X');
}
if ($txttipodomicilio == "CEDIDO") {
    $worksheet->setCellValue('AF87', 'X');
}
if ($txttipodomicilio == "OTROS") {
    $worksheet->setCellValue('AT87', 'X');
}


$worksheet->setCellValue('BE86',$txtCorreo);
$worksheet->setCellValue('I94', $txtTelefonoMovil[0]);
$worksheet->setCellValue('N94', $txtTelefonoMovil[1]);
$worksheet->setCellValue('S94', $txtTelefonoMovil[2]);
$worksheet->setCellValue('X94', $txtTelefonoMovil[3]);
$worksheet->setCellValue('AC94', $txtTelefonoMovil[4]);
$worksheet->setCellValue('AH94', $txtTelefonoMovil[5]);
$worksheet->setCellValue('AM94', $txtTelefonoMovil[6]);
$worksheet->setCellValue('AR94', $txtTelefonoMovil[7]);
$worksheet->setCellValue('AW94', $txtTelefonoMovil[8]);
$worksheet->setCellValue('BE93', $txtActividadEconomica);
$worksheet->setCellValue('AH101', 'X');
$worksheet->setCellValue('CF101', 'X');
$worksheet->setCellValue('CR122', 'X');
$time = time();
$worksheet->setCellValue('AI105', date("d", $time));
$worksheet->setCellValue('AP105', date("m", $time));
$worksheet->setCellValue('AW105', date("Y", $time));

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento.date("H-i-s", $time)."-GUIA_PERSONA_SIN_NEGOCIO_14052020.xlsx";
$writer->save($archivo);
echo '<a href="'.$archivo.'">Descargar GUIA_PERSONA_SIN_NEGOCIO_14052020</a>';
?>
