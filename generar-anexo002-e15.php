<?php
error_reporting(0);

require 'phpspreadsheet/vendor/autoload.php';

use \ConvertApi\ConvertApi;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

$txtPlaca = $_POST['placa'];
$txtSoat = $_POST['soat'];
$txtCITV = $_POST['citv'];
$txtCaf = $_POST['caf'];
$txtCao = $_POST['cao'];

$txtInstalacion = $_POST['instalacion'];

$txtDomicilio = $_POST['domicilio'];
$txtDistrito = $_POST['distrito'];
$txtProvincia = $_POST['provincia'];
$txtDepartamento = $_POST['departamento'];

$txtDeclaracionJurada = $_POST['declaracion_jurada'];

$txtConductor_nombres1 = $_POST['conductor_nombres1'];
$txtConductor_dni1 = $_POST['conductor_dni1'];
$txtConductor_edad1 = $_POST['conductor_edad1'];
$txtConductor_licencia1 = $_POST['conductor_licencia1'];
$txtConductor_categoria1 = $_POST['conductor_categoria1'];

$txtConductor_nombres2 = $_POST['conductor_nombres2'];
$txtConductor_dni2 = $_POST['conductor_dni2'];
$txtConductor_edad2 = $_POST['conductor_edad2'];
$txtConductor_licencia2 = $_POST['conductor_licencia2'];
$txtConductor_categoria2 = $_POST['conductor_categoria2'];

$txtConductor_nombres3 = $_POST['conductor_nombres3'];
$txtConductor_dni3 = $_POST['conductor_dni3'];
$txtConductor_edad3 = $_POST['conductor_edad3'];
$txtConductor_licencia3 = $_POST['conductor_licencia3'];
$txtConductor_categoria3 = $_POST['conductor_categoria3'];

$txtConductor_nombres4 = $_POST['conductor_nombres4'];
$txtConductor_dni4 = $_POST['conductor_dni4'];
$txtConductor_edad4 = $_POST['conductor_edad4'];
$txtConductor_licencia4 = $_POST['conductor_licencia4'];
$txtConductor_categoria4 = $_POST['conductor_categoria4'];

$txtConductor_nombres5 = $_POST['conductor_nombres5'];
$txtConductor_dni5 = $_POST['conductor_dni5'];
$txtConductor_edad5 = $_POST['conductor_edad5'];
$txtConductor_licencia5 = $_POST['conductor_licencia5'];
$txtConductor_categoria5 = $_POST['conductor_categoria5'];

$txtSolicita_nombres  = $_POST['solicita_nombres'];
$txtSolicita_dni  = $_POST['solicita_dni'];



$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("Autorizacion-Renovacion-o-Sustitucion-Transporte-de-Mercancias.xlsx");
$worksheet = $spreadsheet->getActiveSheet();

$rowOneNumber = 37;
$rowTwoNumber = 37;
for ($i = 0; $i < count($txtPlaca); $i++) {
    if ($i < 10) {
        $worksheet->setCellValue('J' . $rowOneNumber, $txtPlaca[$i]);
        $worksheet->setCellValue('V' . $rowOneNumber, $txtSoat[$i]);
        $worksheet->setCellValue('AH' . $rowOneNumber, $txtCITV[$i]);
        $worksheet->setCellValue('AT' . $rowOneNumber, $txtCaf[$i]);
        $worksheet->setCellValue('AX' . $rowOneNumber, $txtCao[$i]);
        $rowOneNumber = $rowOneNumber + 3;
    } else {
        $worksheet->setCellValue('BD' . $rowTwoNumber, $txtPlaca[$i]);
        $worksheet->setCellValue('BP' . $rowTwoNumber, $txtSoat[$i]);
        $worksheet->setCellValue('CB' . $rowTwoNumber, $txtCITV[$i]);
        $worksheet->setCellValue('CN' . $rowTwoNumber, $txtCaf[$i]);
        $worksheet->setCellValue('CR' . $rowTwoNumber, $txtCao[$i]);
        $rowTwoNumber = $rowTwoNumber + 3;
        echo $rowTwoNumber;
    }
}




if ($txtInstalacion == 'propia') $worksheet->setCellValue('AL69', 'X');
else $worksheet->setCellValue('AU69', 'X');

$worksheet->setCellValue('L75', $txtDomicilio);
$worksheet->setCellValue('L82', $txtDistrito);
$worksheet->setCellValue('AN82', $txtProvincia);
$worksheet->setCellValue('BO82', $txtDepartamento);

for ($i = 0; $i < count($txtDeclaracionJurada); $i++){

    switch ($txtDeclaracionJurada[$i]) {
        case 1:
            $worksheet->setCellValue('CJ96', 'X');
            break;
        case 2:
            $worksheet->setCellValue('CJ99', 'X');
            break;
        case 3:
            $worksheet->setCellValue('CJ102', 'X');
            break;
        case 4:
            $worksheet->setCellValue('CJ105', 'X');
            break;
        
        default:
            # code...
            break;
    }
}


$worksheet->setCellValue('I121', $txtConductor_nombres1);
$worksheet->setCellValue('I123', $txtConductor_nombres2);
$worksheet->setCellValue('I125', $txtConductor_nombres3);
$worksheet->setCellValue('I127', $txtConductor_nombres4);
$worksheet->setCellValue('I129', $txtConductor_nombres5);

$worksheet->setCellValue('AX121', $txtConductor_dni1);
$worksheet->setCellValue('AX123', $txtConductor_dni2);
$worksheet->setCellValue('AX125', $txtConductor_dni3);
$worksheet->setCellValue('AX127', $txtConductor_dni4);
$worksheet->setCellValue('AX129', $txtConductor_dni5);

$worksheet->setCellValue('BM121', $txtConductor_edad1);
$worksheet->setCellValue('BM123', $txtConductor_edad2);
$worksheet->setCellValue('BM125', $txtConductor_edad3);
$worksheet->setCellValue('BM127', $txtConductor_edad4);
$worksheet->setCellValue('BM129', $txtConductor_edad5);

$worksheet->setCellValue('BQ121', $txtConductor_licencia1);
$worksheet->setCellValue('BQ123', $txtConductor_licencia2);
$worksheet->setCellValue('BQ125', $txtConductor_licencia3);
$worksheet->setCellValue('BQ127', $txtConductor_licencia4);
$worksheet->setCellValue('BQ129', $txtConductor_licencia5);

$worksheet->setCellValue('CF121', $txtConductor_categoria1);
$worksheet->setCellValue('CF123', $txtConductor_categoria2);
$worksheet->setCellValue('CF125', $txtConductor_categoria3);
$worksheet->setCellValue('CF127', $txtConductor_categoria4);
$worksheet->setCellValue('CF129', $txtConductor_categoria5);

$worksheet->setCellValue('BM147', $txtSolicita_nombres);
$worksheet->setCellValue('BD150', $txtSolicita_dni);

$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = "-Autorizacion-Renovacion-o-Sustitucion-Transporte-de-Mercancias.xlsx";
$writer->save($archivo);
echo '<a href="'.$archivo.'">Descargar Autorizacion-Renovacion-o-Sustitucion-Transporte-de-Mercancias</a>';
?>
