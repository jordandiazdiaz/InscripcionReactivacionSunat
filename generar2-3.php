<?php
/**
 * Creado: Jordan Diaz Diaz
 * Uso: Inscripcion y Reactivacion SUNAT -Persona Natural
 * Fecha: 16/07/2020
 */

//error_reporting(0);

require 'phpspreadsheet/vendor/autoload.php';


use PhpOffice\PhpSpreadsheet\Spreadsheet;

$txtNumeroDocumento_Anterior = $_COOKIE["NumeroDocumento2"];
$NumeroSocios = $_COOKIE["NumeroSocios"];
$spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($txtNumeroDocumento_Anterior."-3-GUIA_PERSONA_JURIDICA_14052020_01.xlsx");
$worksheet = $spreadsheet->getActiveSheet();



if($txtNumeroDocumento=="1"){
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
    $txtPrimerApellido  = $_POST['txtPrimerApellido'];
    $txtSegundoApellido  = $_POST['txtSegundoApellido'];
    $txtNombres = $_POST['txtNombres'];
    $txtFechaNacimiento = $_POST['txtFechaNacimiento'];
    $txtPaisResidencia = $_POST['txtPaisResidencia'];
    $txtRazonSocial = $_POST['txtRazonSocial'];
    $txtTipoVehiculo = $_POST['txtTipoVehiculo'];
    $txtPorcentajeParticipacion  = $_POST['txtPorcentajeParticipacion'];
    
    if($txtTipoDocumento=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS17', "X");
    }
    if($txtTipoDocumento=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS15', "X");
    }
    if($txtTipoDocumento=="Pasaporte"){
        $worksheet->setCellValue('MS19', "X");
    }
    
    $worksheet->setCellValue('NB14', $txtNumeroDocumento[0]);
    $worksheet->setCellValue('NE14', $txtNumeroDocumento[1]);
    $worksheet->setCellValue('NH14', $txtNumeroDocumento[2]);
    $worksheet->setCellValue('NK14', $txtNumeroDocumento[3]);
    $worksheet->setCellValue('NN14', $txtNumeroDocumento[4]);
    $worksheet->setCellValue('NQ14', $txtNumeroDocumento[5]);
    $worksheet->setCellValue('NT14', $txtNumeroDocumento[6]);
    $worksheet->setCellValue('NW14', $txtNumeroDocumento[7]);
    $worksheet->setCellValue('MW19', $txtPrimerApellido);
    $worksheet->setCellValue('NY19', $txtSegundoApellido);
    $worksheet->setCellValue('MW24', $txtNombres);
    $fechanacimiento = $txtFechaNacimiento;
    $fecha_nacimiento = explode('/', $fechanacimiento);
    $worksheet->setCellValue('NL27',$fecha_nacimiento[0]);
    $worksheet->setCellValue('NP27',$fecha_nacimiento[1]);
    $worksheet->setCellValue('NT27',$fecha_nacimiento[2]);
    $worksheet->setCellValue('OG27', $txtPaisResidencia);
    $worksheet->setCellValue('MO31', $txtRazonSocial);
    $worksheet->setCellValue('ME34', $txtTipoVehiculo);
    $time = time();
    $worksheet->setCellValue('ON34', date("d", $time));
    $worksheet->setCellValue('OR34', date("m", $time));
    $worksheet->setCellValue('OV34', date("Y", $time));
    $worksheet->setCellValue('MD39', $txtPorcentajeParticipacion);
}
if($txtNumeroDocumento=="2"){
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
    $txtPrimerApellido  = $_POST['txtPrimerApellido'];
    $txtSegundoApellido  = $_POST['txtSegundoApellido'];
    $txtNombres = $_POST['txtNombres'];
    $txtFechaNacimiento = $_POST['txtFechaNacimiento'];
    $txtPaisResidencia = $_POST['txtPaisResidencia'];
    $txtRazonSocial = $_POST['txtRazonSocial'];
    $txtTipoVehiculo = $_POST['txtTipoVehiculo'];
    $txtPorcentajeParticipacion  = $_POST['txtPorcentajeParticipacion'];
    
    if($txtTipoDocumento=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS17', "X");
    }
    if($txtTipoDocumento=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS15', "X");
    }
    if($txtTipoDocumento=="Pasaporte"){
        $worksheet->setCellValue('MS19', "X");
    }
    
    $worksheet->setCellValue('NB14', $txtNumeroDocumento[0]);
    $worksheet->setCellValue('NE14', $txtNumeroDocumento[1]);
    $worksheet->setCellValue('NH14', $txtNumeroDocumento[2]);
    $worksheet->setCellValue('NK14', $txtNumeroDocumento[3]);
    $worksheet->setCellValue('NN14', $txtNumeroDocumento[4]);
    $worksheet->setCellValue('NQ14', $txtNumeroDocumento[5]);
    $worksheet->setCellValue('NT14', $txtNumeroDocumento[6]);
    $worksheet->setCellValue('NW14', $txtNumeroDocumento[7]);
    $worksheet->setCellValue('MW19', $txtPrimerApellido);
    $worksheet->setCellValue('NY19', $txtSegundoApellido);
    $worksheet->setCellValue('MW24', $txtNombres);
    $fechanacimiento = $txtFechaNacimiento;
    $fecha_nacimiento = explode('/', $fechanacimiento);
    $worksheet->setCellValue('NL27',$fecha_nacimiento[0]);
    $worksheet->setCellValue('NP27',$fecha_nacimiento[1]);
    $worksheet->setCellValue('NT27',$fecha_nacimiento[2]);
    $worksheet->setCellValue('OG27', $txtPaisResidencia);
    $worksheet->setCellValue('MO31', $txtRazonSocial);
    $worksheet->setCellValue('ME34', $txtTipoVehiculo);
    $time = time();
    $worksheet->setCellValue('ON34', date("d", $time));
    $worksheet->setCellValue('OR34', date("m", $time));
    $worksheet->setCellValue('OV34', date("Y", $time));
    $worksheet->setCellValue('MD39', $txtPorcentajeParticipacion);
////////////////////////////////////2//////////////////////////////////

    $txtTipoDocumento2  = $_POST['txtTipoDocumento2'];
    $txtNumeroDocumento2  = $_POST['txtNumeroDocumento2'];
    $txtPrimerApellido2  = $_POST['txtPrimerApellido2'];
    $txtSegundoApellido2  = $_POST['txtSegundoApellido2'];
    $txtNombres2 = $_POST['txtNombres2'];
    $txtFechaNacimiento2 = $_POST['txtFechaNacimiento2'];
    $txtPaisResidencia2 = $_POST['txtPaisResidencia2'];
    $txtRazonSocial2 = $_POST['txtRazonSocial2'];
    $txtTipoVehiculo2 = $_POST['txtTipoVehiculo2'];
    $txtPorcentajeParticipacion2  = $_POST['txtPorcentajeParticipacion2'];

    if($txtTipoDocumento2=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS56', "X");
    }
    if($txtTipoDocumento2=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS54', "X");
    }
    if($txtTipoDocumento2=="Pasaporte"){
        $worksheet->setCellValue('MS58', "X");
    }

    $worksheet->setCellValue('NB53', $txtNumeroDocumento2[0]);
    $worksheet->setCellValue('NE53', $txtNumeroDocumento2[1]);
    $worksheet->setCellValue('NH53', $txtNumeroDocumento2[2]);
    $worksheet->setCellValue('NK53', $txtNumeroDocumento2[3]);
    $worksheet->setCellValue('NN53', $txtNumeroDocumento2[4]);
    $worksheet->setCellValue('NQ53', $txtNumeroDocumento2[5]);
    $worksheet->setCellValue('NT53', $txtNumeroDocumento2[6]);
    $worksheet->setCellValue('NW53', $txtNumeroDocumento2[7]);
    $worksheet->setCellValue('MW58', $txtPrimerApellido2);
    $worksheet->setCellValue('NY58', $txtSegundoApellido2);
    $worksheet->setCellValue('MW63', $txtNombres2);
    $fechanacimiento2 = $txtFechaNacimiento2;
    $fecha_nacimiento2= explode('/', $fechanacimiento2);
    $worksheet->setCellValue('NL66',$fecha_nacimiento2[0]);
    $worksheet->setCellValue('NP66',$fecha_nacimiento2[1]);
    $worksheet->setCellValue('NT66',$fecha_nacimiento2[2]);
    $worksheet->setCellValue('OG66', $txtPaisResidencia2);
    $worksheet->setCellValue('MO70', $txtRazonSocial2);
    $worksheet->setCellValue('ME73', $txtTipoVehiculo2);
    $time = time();
    $worksheet->setCellValue('ON73', date("d", $time));
    $worksheet->setCellValue('OR73', date("m", $time));
    $worksheet->setCellValue('OV73', date("Y", $time));
    $worksheet->setCellValue('MD78', $txtPorcentajeParticipacion2);

    
}
if($txtNumeroDocumento=="3"){
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
    $txtPrimerApellido  = $_POST['txtPrimerApellido'];
    $txtSegundoApellido  = $_POST['txtSegundoApellido'];
    $txtNombres = $_POST['txtNombres'];
    $txtFechaNacimiento = $_POST['txtFechaNacimiento'];
    $txtPaisResidencia = $_POST['txtPaisResidencia'];
    $txtRazonSocial = $_POST['txtRazonSocial'];
    $txtTipoVehiculo = $_POST['txtTipoVehiculo'];
    $txtPorcentajeParticipacion  = $_POST['txtPorcentajeParticipacion'];
    
    if($txtTipoDocumento=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS17', "X");
    }
    if($txtTipoDocumento=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS15', "X");
    }
    if($txtTipoDocumento=="Pasaporte"){
        $worksheet->setCellValue('MS19', "X");
    }
    
    $worksheet->setCellValue('NB14', $txtNumeroDocumento[0]);
    $worksheet->setCellValue('NE14', $txtNumeroDocumento[1]);
    $worksheet->setCellValue('NH14', $txtNumeroDocumento[2]);
    $worksheet->setCellValue('NK14', $txtNumeroDocumento[3]);
    $worksheet->setCellValue('NN14', $txtNumeroDocumento[4]);
    $worksheet->setCellValue('NQ14', $txtNumeroDocumento[5]);
    $worksheet->setCellValue('NT14', $txtNumeroDocumento[6]);
    $worksheet->setCellValue('NW14', $txtNumeroDocumento[7]);
    $worksheet->setCellValue('MW19', $txtPrimerApellido);
    $worksheet->setCellValue('NY19', $txtSegundoApellido);
    $worksheet->setCellValue('MW24', $txtNombres);
    $fechanacimiento = $txtFechaNacimiento;
    $fecha_nacimiento = explode('/', $fechanacimiento);
    $worksheet->setCellValue('NL27',$fecha_nacimiento[0]);
    $worksheet->setCellValue('NP27',$fecha_nacimiento[1]);
    $worksheet->setCellValue('NT27',$fecha_nacimiento[2]);
    $worksheet->setCellValue('OG27', $txtPaisResidencia);
    $worksheet->setCellValue('MO31', $txtRazonSocial);
    $worksheet->setCellValue('ME34', $txtTipoVehiculo);
    $time = time();
    $worksheet->setCellValue('ON34', date("d", $time));
    $worksheet->setCellValue('OR34', date("m", $time));
    $worksheet->setCellValue('OV34', date("Y", $time));
    $worksheet->setCellValue('MD39', $txtPorcentajeParticipacion);
////////////////////////////////////2//////////////////////////////////

    $txtTipoDocumento2  = $_POST['txtTipoDocumento2'];
    $txtNumeroDocumento2  = $_POST['txtNumeroDocumento2'];
    $txtPrimerApellido2  = $_POST['txtPrimerApellido2'];
    $txtSegundoApellido2  = $_POST['txtSegundoApellido2'];
    $txtNombres2 = $_POST['txtNombres2'];
    $txtFechaNacimiento2 = $_POST['txtFechaNacimiento2'];
    $txtPaisResidencia2 = $_POST['txtPaisResidencia2'];
    $txtRazonSocial2 = $_POST['txtRazonSocial2'];
    $txtTipoVehiculo2 = $_POST['txtTipoVehiculo2'];
    $txtPorcentajeParticipacion2  = $_POST['txtPorcentajeParticipacion2'];

    if($txtTipoDocumento2=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS56', "X");
    }
    if($txtTipoDocumento2=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS54', "X");
    }
    if($txtTipoDocumento2=="Pasaporte"){
        $worksheet->setCellValue('MS58', "X");
    }

    $worksheet->setCellValue('NB53', $txtNumeroDocumento2[0]);
    $worksheet->setCellValue('NE53', $txtNumeroDocumento2[1]);
    $worksheet->setCellValue('NH53', $txtNumeroDocumento2[2]);
    $worksheet->setCellValue('NK53', $txtNumeroDocumento2[3]);
    $worksheet->setCellValue('NN53', $txtNumeroDocumento2[4]);
    $worksheet->setCellValue('NQ53', $txtNumeroDocumento2[5]);
    $worksheet->setCellValue('NT53', $txtNumeroDocumento2[6]);
    $worksheet->setCellValue('NW53', $txtNumeroDocumento2[7]);
    $worksheet->setCellValue('MW58', $txtPrimerApellido2);
    $worksheet->setCellValue('NY58', $txtSegundoApellido2);
    $worksheet->setCellValue('MW63', $txtNombres2);
    $fechanacimiento2 = $txtFechaNacimiento2;
    $fecha_nacimiento2= explode('/', $fechanacimiento2);
    $worksheet->setCellValue('NL66',$fecha_nacimiento2[0]);
    $worksheet->setCellValue('NP66',$fecha_nacimiento2[1]);
    $worksheet->setCellValue('NT66',$fecha_nacimiento2[2]);
    $worksheet->setCellValue('OG66', $txtPaisResidencia2);
    $worksheet->setCellValue('MO70', $txtRazonSocial2);
    $worksheet->setCellValue('ME73', $txtTipoVehiculo2);
    $time = time();
    $worksheet->setCellValue('ON73', date("d", $time));
    $worksheet->setCellValue('OR73', date("m", $time));
    $worksheet->setCellValue('OV73', date("Y", $time));
    $worksheet->setCellValue('MD78', $txtPorcentajeParticipacion2);

    ////////////////////3/////////////////////////////////////////
    $txtTipoDocumento3  = $_POST['txtTipoDocumento3'];
    $txtNumeroDocumento3  = $_POST['txtNumeroDocumento3'];
    $txtPrimerApellido3  = $_POST['txtPrimerApellido3'];
    $txtSegundoApellido3  = $_POST['txtSegundoApellido3'];
    $txtNombres3 = $_POST['txtNombres3'];
    $txtFechaNacimiento3 = $_POST['txtFechaNacimiento3'];
    $txtPaisResidencia3 = $_POST['txtPaisResidencia3'];
    $txtRazonSocial3 = $_POST['txtRazonSocial3'];
    $txtTipoVehiculo3 = $_POST['txtTipoVehiculo3'];
    $txtPorcentajeParticipacion3  = $_POST['txtPorcentajeParticipacion3'];

    if($txtTipoDocumento3=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS95', "X");
    }
    if($txtTipoDocumento3=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS93', "X");
    }
    if($txtTipoDocumento3=="Pasaporte"){
        $worksheet->setCellValue('MS97', "X");
    }

    $worksheet->setCellValue('NB92', $txtNumeroDocumento3[0]);
    $worksheet->setCellValue('NE92', $txtNumeroDocumento3[1]);
    $worksheet->setCellValue('NH92', $txtNumeroDocumento3[2]);
    $worksheet->setCellValue('NK92', $txtNumeroDocumento3[3]);
    $worksheet->setCellValue('NN92', $txtNumeroDocumento3[4]);
    $worksheet->setCellValue('NQ92', $txtNumeroDocumento3[5]);
    $worksheet->setCellValue('NT92', $txtNumeroDocumento3[6]);
    $worksheet->setCellValue('NW92', $txtNumeroDocumento3[7]);
    $worksheet->setCellValue('MW97', $txtPrimerApellido3);
    $worksheet->setCellValue('NY97', $txtSegundoApellido3);
    $worksheet->setCellValue('MW102', $txtNombres3);
    $fechanacimiento3 = $txtFechaNacimiento3;
    $fecha_nacimiento3= explode('/', $fechanacimiento3);
    $worksheet->setCellValue('NL105',$fecha_nacimiento3[0]);
    $worksheet->setCellValue('NP105',$fecha_nacimiento3[1]);
    $worksheet->setCellValue('NT105',$fecha_nacimiento3[2]);
    $worksheet->setCellValue('OG105', $txtPaisResidencia3);
    $worksheet->setCellValue('MO109', $txtRazonSocial3);
    $worksheet->setCellValue('ME112', $txtTipoVehiculo3);
    $time = time();
    $worksheet->setCellValue('ON112', date("d", $time));
    $worksheet->setCellValue('OR112', date("m", $time));
    $worksheet->setCellValue('OV112', date("Y", $time));
    $worksheet->setCellValue('MD117', $txtPorcentajeParticipacion3);
}
if($NumeroSocios=="4"){
    
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
    $txtPrimerApellido  = $_POST['txtPrimerApellido'];
    $txtSegundoApellido  = $_POST['txtSegundoApellido'];
    $txtNombres = $_POST['txtNombres'];
    $txtFechaNacimiento = $_POST['txtFechaNacimiento'];
    $txtPaisResidencia = $_POST['txtPaisResidencia'];
    $txtRazonSocial = $_POST['txtRazonSocial'];
    $txtTipoVehiculo = $_POST['txtTipoVehiculo'];
    $txtPorcentajeParticipacion  = $_POST['txtPorcentajeParticipacion'];
    
    if($txtTipoDocumento=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS17', "X");
    }
    if($txtTipoDocumento=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS15', "X");
    }
    if($txtTipoDocumento=="Pasaporte"){
        $worksheet->setCellValue('MS19', "X");
    }
    
    $worksheet->setCellValue('NB14', $txtNumeroDocumento[0]);
    $worksheet->setCellValue('NE14', $txtNumeroDocumento[1]);
    $worksheet->setCellValue('NH14', $txtNumeroDocumento[2]);
    $worksheet->setCellValue('NK14', $txtNumeroDocumento[3]);
    $worksheet->setCellValue('NN14', $txtNumeroDocumento[4]);
    $worksheet->setCellValue('NQ14', $txtNumeroDocumento[5]);
    $worksheet->setCellValue('NT14', $txtNumeroDocumento[6]);
    $worksheet->setCellValue('NW14', $txtNumeroDocumento[7]);
    $worksheet->setCellValue('MW19', $txtPrimerApellido);
    $worksheet->setCellValue('NY19', $txtSegundoApellido);
    $worksheet->setCellValue('MW24', $txtNombres);
    $fechanacimiento = $txtFechaNacimiento;
    $fecha_nacimiento = explode('/', $fechanacimiento);
    $worksheet->setCellValue('NL27',$fecha_nacimiento[0]);
    $worksheet->setCellValue('NP27',$fecha_nacimiento[1]);
    $worksheet->setCellValue('NT27',$fecha_nacimiento[2]);
    $worksheet->setCellValue('OG27', $txtPaisResidencia);
    $worksheet->setCellValue('MO31', $txtRazonSocial);
    $worksheet->setCellValue('ME34', $txtTipoVehiculo);
    $time = time();
    $worksheet->setCellValue('ON34', date("d", $time));
    $worksheet->setCellValue('OR34', date("m", $time));
    $worksheet->setCellValue('OV34', date("Y", $time));
    $worksheet->setCellValue('MD39', $txtPorcentajeParticipacion);
////////////////////////////////////2//////////////////////////////////

    $txtTipoDocumento2  = $_POST['txtTipoDocumento2'];
    $txtNumeroDocumento2  = $_POST['txtNumeroDocumento2'];
    $txtPrimerApellido2  = $_POST['txtPrimerApellido2'];
    $txtSegundoApellido2  = $_POST['txtSegundoApellido2'];
    $txtNombres2 = $_POST['txtNombres2'];
    $txtFechaNacimiento2 = $_POST['txtFechaNacimiento2'];
    $txtPaisResidencia2 = $_POST['txtPaisResidencia2'];
    $txtRazonSocial2 = $_POST['txtRazonSocial2'];
    $txtTipoVehiculo2 = $_POST['txtTipoVehiculo2'];
    $txtPorcentajeParticipacion2  = $_POST['txtPorcentajeParticipacion2'];

    if($txtTipoDocumento2=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS56', "X");
    }
    if($txtTipoDocumento2=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS54', "X");
    }
    if($txtTipoDocumento2=="Pasaporte"){
        $worksheet->setCellValue('MS58', "X");
    }

    $worksheet->setCellValue('NB53', $txtNumeroDocumento2[0]);
    $worksheet->setCellValue('NE53', $txtNumeroDocumento2[1]);
    $worksheet->setCellValue('NH53', $txtNumeroDocumento2[2]);
    $worksheet->setCellValue('NK53', $txtNumeroDocumento2[3]);
    $worksheet->setCellValue('NN53', $txtNumeroDocumento2[4]);
    $worksheet->setCellValue('NQ53', $txtNumeroDocumento2[5]);
    $worksheet->setCellValue('NT53', $txtNumeroDocumento2[6]);
    $worksheet->setCellValue('NW53', $txtNumeroDocumento2[7]);
    $worksheet->setCellValue('MW58', $txtPrimerApellido2);
    $worksheet->setCellValue('NY58', $txtSegundoApellido2);
    $worksheet->setCellValue('MW63', $txtNombres2);
    $fechanacimiento2 = $txtFechaNacimiento2;
    $fecha_nacimiento2= explode('/', $fechanacimiento2);
    $worksheet->setCellValue('NL66',$fecha_nacimiento2[0]);
    $worksheet->setCellValue('NP66',$fecha_nacimiento2[1]);
    $worksheet->setCellValue('NT66',$fecha_nacimiento2[2]);
    $worksheet->setCellValue('OG66', $txtPaisResidencia2);
    $worksheet->setCellValue('MO70', $txtRazonSocial2);
    $worksheet->setCellValue('ME73', $txtTipoVehiculo2);
    $time = time();
    $worksheet->setCellValue('ON73', date("d", $time));
    $worksheet->setCellValue('OR73', date("m", $time));
    $worksheet->setCellValue('OV73', date("Y", $time));
    $worksheet->setCellValue('MD78', $txtPorcentajeParticipacion2);

    ////////////////////3/////////////////////////////////////////
    $txtTipoDocumento3  = $_POST['txtTipoDocumento3'];
    $txtNumeroDocumento3  = $_POST['txtNumeroDocumento3'];
    $txtPrimerApellido3  = $_POST['txtPrimerApellido3'];
    $txtSegundoApellido3  = $_POST['txtSegundoApellido3'];
    $txtNombres3 = $_POST['txtNombres3'];
    $txtFechaNacimiento3 = $_POST['txtFechaNacimiento3'];
    $txtPaisResidencia3 = $_POST['txtPaisResidencia3'];
    $txtRazonSocial3 = $_POST['txtRazonSocial3'];
    $txtTipoVehiculo3 = $_POST['txtTipoVehiculo3'];
    $txtPorcentajeParticipacion3  = $_POST['txtPorcentajeParticipacion3'];

    if($txtTipoDocumento3=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS95', "X");
    }
    if($txtTipoDocumento3=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS93', "X");
    }
    if($txtTipoDocumento3=="Pasaporte"){
        $worksheet->setCellValue('MS97', "X");
    }

    $worksheet->setCellValue('NB92', $txtNumeroDocumento3[0]);
    $worksheet->setCellValue('NE92', $txtNumeroDocumento3[1]);
    $worksheet->setCellValue('NH92', $txtNumeroDocumento3[2]);
    $worksheet->setCellValue('NK92', $txtNumeroDocumento3[3]);
    $worksheet->setCellValue('NN92', $txtNumeroDocumento3[4]);
    $worksheet->setCellValue('NQ92', $txtNumeroDocumento3[5]);
    $worksheet->setCellValue('NT92', $txtNumeroDocumento3[6]);
    $worksheet->setCellValue('NW92', $txtNumeroDocumento3[7]);
    $worksheet->setCellValue('MW97', $txtPrimerApellido3);
    $worksheet->setCellValue('NY97', $txtSegundoApellido3);
    $worksheet->setCellValue('MW102', $txtNombres3);
    $fechanacimiento3 = $txtFechaNacimiento3;
    $fecha_nacimiento3= explode('/', $fechanacimiento3);
    $worksheet->setCellValue('NL105',$fecha_nacimiento3[0]);
    $worksheet->setCellValue('NP105',$fecha_nacimiento3[1]);
    $worksheet->setCellValue('NT105',$fecha_nacimiento3[2]);
    $worksheet->setCellValue('OG105', $txtPaisResidencia3);
    $worksheet->setCellValue('MO109', $txtRazonSocial3);
    $worksheet->setCellValue('ME112', $txtTipoVehiculo3);
    $time = time();
    $worksheet->setCellValue('ON112', date("d", $time));
    $worksheet->setCellValue('OR112', date("m", $time));
    $worksheet->setCellValue('OV112', date("Y", $time));
    $worksheet->setCellValue('MD117', $txtPorcentajeParticipacion3);
    ////////////////////4/////////////////////////////////////////
    $txtTipoDocumento4  = $_POST['txtTipoDocumento4'];
    $txtNumeroDocumento4  = $_POST['txtNumeroDocumento4'];
    $txtPrimerApellido4  = $_POST['txtPrimerApellido4'];
    $txtSegundoApellido4  = $_POST['txtSegundoApellido4'];
    $txtNombres4 = $_POST['txtNombres4'];
    $txtFechaNacimiento4 = $_POST['txtFechaNacimiento4'];
    $txtPaisResidencia4 = $_POST['txtPaisResidencia4'];
    $txtRazonSocial4 = $_POST['txtRazonSocial4'];
    $txtTipoVehiculo4 = $_POST['txtTipoVehiculo4'];
    $txtPorcentajeParticipacion4  = $_POST['txtPorcentajeParticipacion4'];

    if($txtTipoDocumento4=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS134', "X");
    }
    if($txtTipoDocumento4=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS132', "X");
    }
    if($txtTipoDocumento4=="Pasaporte"){
        $worksheet->setCellValue('MS136', "X");
    }

    $worksheet->setCellValue('NB131', $txtNumeroDocumento4[0]);
    $worksheet->setCellValue('NE131', $txtNumeroDocumento4[1]);
    $worksheet->setCellValue('NH131', $txtNumeroDocumento4[2]);
    $worksheet->setCellValue('NK131', $txtNumeroDocumento4[3]);
    $worksheet->setCellValue('NN131', $txtNumeroDocumento4[4]);
    $worksheet->setCellValue('NQ131', $txtNumeroDocumento4[5]);
    $worksheet->setCellValue('NT131', $txtNumeroDocumento4[6]);
    $worksheet->setCellValue('NW131', $txtNumeroDocumento4[7]);
    $worksheet->setCellValue('MW136', $txtPrimerApellido4);
    $worksheet->setCellValue('NY136', $txtSegundoApellido4);
    $worksheet->setCellValue('MW141', $txtNombres4);
    $fechanacimiento4 = $txtFechaNacimiento4;
    $fecha_nacimiento4= explode('/', $fechanacimiento4);
    $worksheet->setCellValue('NL144',$fecha_nacimiento4[0]);
    $worksheet->setCellValue('NP144',$fecha_nacimiento4[1]);
    $worksheet->setCellValue('NT144',$fecha_nacimiento4[2]);
    $worksheet->setCellValue('OG144', $txtPaisResidencia4);
    $worksheet->setCellValue('MO148', $txtRazonSocial4);
    $worksheet->setCellValue('ME151', $txtTipoVehiculo4);
    $time = time();
    $worksheet->setCellValue('ON151', date("d", $time));
    $worksheet->setCellValue('OR151', date("m", $time));
    $worksheet->setCellValue('OV151', date("Y", $time));
    $worksheet->setCellValue('MD156', $txtPorcentajeParticipacion4);
}
if($NumeroSocios=="5"){
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
    $txtPrimerApellido  = $_POST['txtPrimerApellido'];
    $txtSegundoApellido  = $_POST['txtSegundoApellido'];
    $txtNombres = $_POST['txtNombres'];
    $txtFechaNacimiento = $_POST['txtFechaNacimiento'];
    $txtPaisResidencia = $_POST['txtPaisResidencia'];
    $txtRazonSocial = $_POST['txtRazonSocial'];
    $txtTipoVehiculo = $_POST['txtTipoVehiculo'];
    $txtPorcentajeParticipacion  = $_POST['txtPorcentajeParticipacion'];
    
    if($txtTipoDocumento=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS17', "X");
    }
    if($txtTipoDocumento=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS15', "X");
    }
    if($txtTipoDocumento=="Pasaporte"){
        $worksheet->setCellValue('MS19', "X");
    }
    
    $worksheet->setCellValue('NB14', $txtNumeroDocumento[0]);
    $worksheet->setCellValue('NE14', $txtNumeroDocumento[1]);
    $worksheet->setCellValue('NH14', $txtNumeroDocumento[2]);
    $worksheet->setCellValue('NK14', $txtNumeroDocumento[3]);
    $worksheet->setCellValue('NN14', $txtNumeroDocumento[4]);
    $worksheet->setCellValue('NQ14', $txtNumeroDocumento[5]);
    $worksheet->setCellValue('NT14', $txtNumeroDocumento[6]);
    $worksheet->setCellValue('NW14', $txtNumeroDocumento[7]);
    $worksheet->setCellValue('MW19', $txtPrimerApellido);
    $worksheet->setCellValue('NY19', $txtSegundoApellido);
    $worksheet->setCellValue('MW24', $txtNombres);
    $fechanacimiento = $txtFechaNacimiento;
    $fecha_nacimiento = explode('/', $fechanacimiento);
    $worksheet->setCellValue('NL27',$fecha_nacimiento[0]);
    $worksheet->setCellValue('NP27',$fecha_nacimiento[1]);
    $worksheet->setCellValue('NT27',$fecha_nacimiento[2]);
    $worksheet->setCellValue('OG27', $txtPaisResidencia);
    $worksheet->setCellValue('MO31', $txtRazonSocial);
    $worksheet->setCellValue('ME34', $txtTipoVehiculo);
    $time = time();
    $worksheet->setCellValue('ON34', date("d", $time));
    $worksheet->setCellValue('OR34', date("m", $time));
    $worksheet->setCellValue('OV34', date("Y", $time));
    $worksheet->setCellValue('MD39', $txtPorcentajeParticipacion);
////////////////////////////////////2//////////////////////////////////

    $txtTipoDocumento2  = $_POST['txtTipoDocumento2'];
    $txtNumeroDocumento2  = $_POST['txtNumeroDocumento2'];
    $txtPrimerApellido2  = $_POST['txtPrimerApellido2'];
    $txtSegundoApellido2  = $_POST['txtSegundoApellido2'];
    $txtNombres2 = $_POST['txtNombres2'];
    $txtFechaNacimiento2 = $_POST['txtFechaNacimiento2'];
    $txtPaisResidencia2 = $_POST['txtPaisResidencia2'];
    $txtRazonSocial2 = $_POST['txtRazonSocial2'];
    $txtTipoVehiculo2 = $_POST['txtTipoVehiculo2'];
    $txtPorcentajeParticipacion2  = $_POST['txtPorcentajeParticipacion2'];

    if($txtTipoDocumento2=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS56', "X");
    }
    if($txtTipoDocumento2=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS54', "X");
    }
    if($txtTipoDocumento2=="Pasaporte"){
        $worksheet->setCellValue('MS58', "X");
    }

    $worksheet->setCellValue('NB53', $txtNumeroDocumento2[0]);
    $worksheet->setCellValue('NE53', $txtNumeroDocumento2[1]);
    $worksheet->setCellValue('NH53', $txtNumeroDocumento2[2]);
    $worksheet->setCellValue('NK53', $txtNumeroDocumento2[3]);
    $worksheet->setCellValue('NN53', $txtNumeroDocumento2[4]);
    $worksheet->setCellValue('NQ53', $txtNumeroDocumento2[5]);
    $worksheet->setCellValue('NT53', $txtNumeroDocumento2[6]);
    $worksheet->setCellValue('NW53', $txtNumeroDocumento2[7]);
    $worksheet->setCellValue('MW58', $txtPrimerApellido2);
    $worksheet->setCellValue('NY58', $txtSegundoApellido2);
    $worksheet->setCellValue('MW63', $txtNombres2);
    $fechanacimiento2 = $txtFechaNacimiento2;
    $fecha_nacimiento2= explode('/', $fechanacimiento2);
    $worksheet->setCellValue('NL66',$fecha_nacimiento2[0]);
    $worksheet->setCellValue('NP66',$fecha_nacimiento2[1]);
    $worksheet->setCellValue('NT66',$fecha_nacimiento2[2]);
    $worksheet->setCellValue('OG66', $txtPaisResidencia2);
    $worksheet->setCellValue('MO70', $txtRazonSocial2);
    $worksheet->setCellValue('ME73', $txtTipoVehiculo2);
    $time = time();
    $worksheet->setCellValue('ON73', date("d", $time));
    $worksheet->setCellValue('OR73', date("m", $time));
    $worksheet->setCellValue('OV73', date("Y", $time));
    $worksheet->setCellValue('MD78', $txtPorcentajeParticipacion2);

    ////////////////////3/////////////////////////////////////////
    $txtTipoDocumento3  = $_POST['txtTipoDocumento3'];
    $txtNumeroDocumento3  = $_POST['txtNumeroDocumento3'];
    $txtPrimerApellido3  = $_POST['txtPrimerApellido3'];
    $txtSegundoApellido3  = $_POST['txtSegundoApellido3'];
    $txtNombres3 = $_POST['txtNombres3'];
    $txtFechaNacimiento3 = $_POST['txtFechaNacimiento3'];
    $txtPaisResidencia3 = $_POST['txtPaisResidencia3'];
    $txtRazonSocial3 = $_POST['txtRazonSocial3'];
    $txtTipoVehiculo3 = $_POST['txtTipoVehiculo3'];
    $txtPorcentajeParticipacion3  = $_POST['txtPorcentajeParticipacion3'];

    if($txtTipoDocumento3=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS95', "X");
    }
    if($txtTipoDocumento3=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS93', "X");
    }
    if($txtTipoDocumento3=="Pasaporte"){
        $worksheet->setCellValue('MS97', "X");
    }

    $worksheet->setCellValue('NB92', $txtNumeroDocumento3[0]);
    $worksheet->setCellValue('NE92', $txtNumeroDocumento3[1]);
    $worksheet->setCellValue('NH92', $txtNumeroDocumento3[2]);
    $worksheet->setCellValue('NK92', $txtNumeroDocumento3[3]);
    $worksheet->setCellValue('NN92', $txtNumeroDocumento3[4]);
    $worksheet->setCellValue('NQ92', $txtNumeroDocumento3[5]);
    $worksheet->setCellValue('NT92', $txtNumeroDocumento3[6]);
    $worksheet->setCellValue('NW92', $txtNumeroDocumento3[7]);
    $worksheet->setCellValue('MW97', $txtPrimerApellido3);
    $worksheet->setCellValue('NY97', $txtSegundoApellido3);
    $worksheet->setCellValue('MW102', $txtNombres3);
    $fechanacimiento3 = $txtFechaNacimiento3;
    $fecha_nacimiento3= explode('/', $fechanacimiento3);
    $worksheet->setCellValue('NL105',$fecha_nacimiento3[0]);
    $worksheet->setCellValue('NP105',$fecha_nacimiento3[1]);
    $worksheet->setCellValue('NT105',$fecha_nacimiento3[2]);
    $worksheet->setCellValue('OG105', $txtPaisResidencia3);
    $worksheet->setCellValue('MO109', $txtRazonSocial3);
    $worksheet->setCellValue('ME112', $txtTipoVehiculo3);
    $time = time();
    $worksheet->setCellValue('ON112', date("d", $time));
    $worksheet->setCellValue('OR112', date("m", $time));
    $worksheet->setCellValue('OV112', date("Y", $time));
    $worksheet->setCellValue('MD117', $txtPorcentajeParticipacion3);    
      ////////////////////4/////////////////////////////////////////
      $txtTipoDocumento4  = $_POST['txtTipoDocumento4'];
      $txtNumeroDocumento4  = $_POST['txtNumeroDocumento4'];
      $txtPrimerApellido4  = $_POST['txtPrimerApellido4'];
      $txtSegundoApellido4  = $_POST['txtSegundoApellido4'];
      $txtNombres4 = $_POST['txtNombres4'];
      $txtFechaNacimiento4 = $_POST['txtFechaNacimiento4'];
      $txtPaisResidencia4 = $_POST['txtPaisResidencia4'];
      $txtRazonSocial4 = $_POST['txtRazonSocial4'];
      $txtTipoVehiculo4 = $_POST['txtTipoVehiculo4'];
      $txtPorcentajeParticipacion4  = $_POST['txtPorcentajeParticipacion4'];
  
      if($txtTipoDocumento4=="Carnet de Extranjeria"){
          $worksheet->setCellValue('MS134', "X");
      }
      if($txtTipoDocumento4=="Documento Nacional de Identidad"){
          $worksheet->setCellValue('MS132', "X");
      }
      if($txtTipoDocumento4=="Pasaporte"){
          $worksheet->setCellValue('MS136', "X");
      }
  
      $worksheet->setCellValue('NB131', $txtNumeroDocumento4[0]);
      $worksheet->setCellValue('NE131', $txtNumeroDocumento4[1]);
      $worksheet->setCellValue('NH131', $txtNumeroDocumento4[2]);
      $worksheet->setCellValue('NK131', $txtNumeroDocumento4[3]);
      $worksheet->setCellValue('NN131', $txtNumeroDocumento4[4]);
      $worksheet->setCellValue('NQ131', $txtNumeroDocumento4[5]);
      $worksheet->setCellValue('NT131', $txtNumeroDocumento4[6]);
      $worksheet->setCellValue('NW131', $txtNumeroDocumento4[7]);
      $worksheet->setCellValue('MW136', $txtPrimerApellido4);
      $worksheet->setCellValue('NY136', $txtSegundoApellido4);
      $worksheet->setCellValue('MW141', $txtNombres4);
      $fechanacimiento4 = $txtFechaNacimiento4;
      $fecha_nacimiento4= explode('/', $fechanacimiento4);
      $worksheet->setCellValue('NL144',$fecha_nacimiento4[0]);
      $worksheet->setCellValue('NP144',$fecha_nacimiento4[1]);
      $worksheet->setCellValue('NT144',$fecha_nacimiento4[2]);
      $worksheet->setCellValue('OG144', $txtPaisResidencia4);
      $worksheet->setCellValue('MO148', $txtRazonSocial4);
      $worksheet->setCellValue('ME151', $txtTipoVehiculo4);
      $time = time();
      $worksheet->setCellValue('ON151', date("d", $time));
      $worksheet->setCellValue('OR151', date("m", $time));
      $worksheet->setCellValue('OV151', date("Y", $time));
      $worksheet->setCellValue('MD156', $txtPorcentajeParticipacion4);
    ////////////////////5/////////////////////////////////////////
    $txtTipoDocumento5  = $_POST['txtTipoDocumento5'];
    $txtNumeroDocumento5  = $_POST['txtNumeroDocumento5'];
    $txtPrimerApellido5  = $_POST['txtPrimerApellido5'];
    $txtSegundoApellido5  = $_POST['txtSegundoApellido5'];
    $txtNombres5 = $_POST['txtNombres5'];
    $txtFechaNacimiento5 = $_POST['txtFechaNacimiento5'];
    $txtPaisResidencia5 = $_POST['txtPaisResidencia5'];
    $txtRazonSocial5 = $_POST['txtRazonSocial5'];
    $txtTipoVehiculo5 = $_POST['txtTipoVehiculo5'];
    $txtPorcentajeParticipacion5  = $_POST['txtPorcentajeParticipacion5'];

    if($txtTipoDocumento5=="Carnet de Extranjeria"){
        $worksheet->setCellValue('QT17', "X");
    }
    if($txtTipoDocumento5=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('QT15', "X");
    }
    if($txtTipoDocumento5=="Pasaporte"){
        $worksheet->setCellValue('QT19', "X");
    }

    $worksheet->setCellValue('RC14', $txtNumeroDocumento5[0]);
    $worksheet->setCellValue('RF14', $txtNumeroDocumento5[1]);
    $worksheet->setCellValue('RI14', $txtNumeroDocumento5[2]);
    $worksheet->setCellValue('RL14', $txtNumeroDocumento5[3]);
    $worksheet->setCellValue('RO14', $txtNumeroDocumento5[4]);
    $worksheet->setCellValue('RR14', $txtNumeroDocumento5[5]);
    $worksheet->setCellValue('RU14', $txtNumeroDocumento5[6]);
    $worksheet->setCellValue('RX14', $txtNumeroDocumento5[7]);
    $worksheet->setCellValue('QX19', $txtPrimerApellido5);
    $worksheet->setCellValue('RZ19', $txtSegundoApellido5);
    $worksheet->setCellValue('QX24', $txtNombres5);
    $fechanacimiento5 = $txtFechaNacimiento5;
    $fecha_nacimiento5= explode('/', $fechanacimiento5);
    $worksheet->setCellValue('RM27',$fecha_nacimiento5[0]);
    $worksheet->setCellValue('RQ27',$fecha_nacimiento5[1]);
    $worksheet->setCellValue('RU27',$fecha_nacimiento5[2]);
    $worksheet->setCellValue('SH27', $txtPaisResidencia5);
    $worksheet->setCellValue('QP31', $txtRazonSocial5);
    $worksheet->setCellValue('QF34', $txtTipoVehiculo5);
    $time = time();
    $worksheet->setCellValue('SO34', date("d", $time));
    $worksheet->setCellValue('SS34', date("m", $time));
    $worksheet->setCellValue('SW34', date("Y", $time));
    $worksheet->setCellValue('QE39', $txtPorcentajeParticipacion5);
}
if($NumeroSocios=="6"){
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
    $txtPrimerApellido  = $_POST['txtPrimerApellido'];
    $txtSegundoApellido  = $_POST['txtSegundoApellido'];
    $txtNombres = $_POST['txtNombres'];
    $txtFechaNacimiento = $_POST['txtFechaNacimiento'];
    $txtPaisResidencia = $_POST['txtPaisResidencia'];
    $txtRazonSocial = $_POST['txtRazonSocial'];
    $txtTipoVehiculo = $_POST['txtTipoVehiculo'];
    $txtPorcentajeParticipacion  = $_POST['txtPorcentajeParticipacion'];
    
    if($txtTipoDocumento=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS17', "X");
    }
    if($txtTipoDocumento=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS15', "X");
    }
    if($txtTipoDocumento=="Pasaporte"){
        $worksheet->setCellValue('MS19', "X");
    }
    
    $worksheet->setCellValue('NB14', $txtNumeroDocumento[0]);
    $worksheet->setCellValue('NE14', $txtNumeroDocumento[1]);
    $worksheet->setCellValue('NH14', $txtNumeroDocumento[2]);
    $worksheet->setCellValue('NK14', $txtNumeroDocumento[3]);
    $worksheet->setCellValue('NN14', $txtNumeroDocumento[4]);
    $worksheet->setCellValue('NQ14', $txtNumeroDocumento[5]);
    $worksheet->setCellValue('NT14', $txtNumeroDocumento[6]);
    $worksheet->setCellValue('NW14', $txtNumeroDocumento[7]);
    $worksheet->setCellValue('MW19', $txtPrimerApellido);
    $worksheet->setCellValue('NY19', $txtSegundoApellido);
    $worksheet->setCellValue('MW24', $txtNombres);
    $fechanacimiento = $txtFechaNacimiento;
    $fecha_nacimiento = explode('/', $fechanacimiento);
    $worksheet->setCellValue('NL27',$fecha_nacimiento[0]);
    $worksheet->setCellValue('NP27',$fecha_nacimiento[1]);
    $worksheet->setCellValue('NT27',$fecha_nacimiento[2]);
    $worksheet->setCellValue('OG27', $txtPaisResidencia);
    $worksheet->setCellValue('MO31', $txtRazonSocial);
    $worksheet->setCellValue('ME34', $txtTipoVehiculo);
    $time = time();
    $worksheet->setCellValue('ON34', date("d", $time));
    $worksheet->setCellValue('OR34', date("m", $time));
    $worksheet->setCellValue('OV34', date("Y", $time));
    $worksheet->setCellValue('MD39', $txtPorcentajeParticipacion);
////////////////////////////////////2//////////////////////////////////

    $txtTipoDocumento2  = $_POST['txtTipoDocumento2'];
    $txtNumeroDocumento2  = $_POST['txtNumeroDocumento2'];
    $txtPrimerApellido2  = $_POST['txtPrimerApellido2'];
    $txtSegundoApellido2  = $_POST['txtSegundoApellido2'];
    $txtNombres2 = $_POST['txtNombres2'];
    $txtFechaNacimiento2 = $_POST['txtFechaNacimiento2'];
    $txtPaisResidencia2 = $_POST['txtPaisResidencia2'];
    $txtRazonSocial2 = $_POST['txtRazonSocial2'];
    $txtTipoVehiculo2 = $_POST['txtTipoVehiculo2'];
    $txtPorcentajeParticipacion2  = $_POST['txtPorcentajeParticipacion2'];

    if($txtTipoDocumento2=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS56', "X");
    }
    if($txtTipoDocumento2=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS54', "X");
    }
    if($txtTipoDocumento2=="Pasaporte"){
        $worksheet->setCellValue('MS58', "X");
    }

    $worksheet->setCellValue('NB53', $txtNumeroDocumento2[0]);
    $worksheet->setCellValue('NE53', $txtNumeroDocumento2[1]);
    $worksheet->setCellValue('NH53', $txtNumeroDocumento2[2]);
    $worksheet->setCellValue('NK53', $txtNumeroDocumento2[3]);
    $worksheet->setCellValue('NN53', $txtNumeroDocumento2[4]);
    $worksheet->setCellValue('NQ53', $txtNumeroDocumento2[5]);
    $worksheet->setCellValue('NT53', $txtNumeroDocumento2[6]);
    $worksheet->setCellValue('NW53', $txtNumeroDocumento2[7]);
    $worksheet->setCellValue('MW58', $txtPrimerApellido2);
    $worksheet->setCellValue('NY58', $txtSegundoApellido2);
    $worksheet->setCellValue('MW63', $txtNombres2);
    $fechanacimiento2 = $txtFechaNacimiento2;
    $fecha_nacimiento2= explode('/', $fechanacimiento2);
    $worksheet->setCellValue('NL66',$fecha_nacimiento2[0]);
    $worksheet->setCellValue('NP66',$fecha_nacimiento2[1]);
    $worksheet->setCellValue('NT66',$fecha_nacimiento2[2]);
    $worksheet->setCellValue('OG66', $txtPaisResidencia2);
    $worksheet->setCellValue('MO70', $txtRazonSocial2);
    $worksheet->setCellValue('ME73', $txtTipoVehiculo2);
    $time = time();
    $worksheet->setCellValue('ON73', date("d", $time));
    $worksheet->setCellValue('OR73', date("m", $time));
    $worksheet->setCellValue('OV73', date("Y", $time));
    $worksheet->setCellValue('MD78', $txtPorcentajeParticipacion2);

    ////////////////////3/////////////////////////////////////////
    $txtTipoDocumento3  = $_POST['txtTipoDocumento3'];
    $txtNumeroDocumento3  = $_POST['txtNumeroDocumento3'];
    $txtPrimerApellido3  = $_POST['txtPrimerApellido3'];
    $txtSegundoApellido3  = $_POST['txtSegundoApellido3'];
    $txtNombres3 = $_POST['txtNombres3'];
    $txtFechaNacimiento3 = $_POST['txtFechaNacimiento3'];
    $txtPaisResidencia3 = $_POST['txtPaisResidencia3'];
    $txtRazonSocial3 = $_POST['txtRazonSocial3'];
    $txtTipoVehiculo3 = $_POST['txtTipoVehiculo3'];
    $txtPorcentajeParticipacion3  = $_POST['txtPorcentajeParticipacion3'];

    if($txtTipoDocumento3=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS95', "X");
    }
    if($txtTipoDocumento3=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS93', "X");
    }
    if($txtTipoDocumento3=="Pasaporte"){
        $worksheet->setCellValue('MS97', "X");
    }

    $worksheet->setCellValue('NB92', $txtNumeroDocumento3[0]);
    $worksheet->setCellValue('NE92', $txtNumeroDocumento3[1]);
    $worksheet->setCellValue('NH92', $txtNumeroDocumento3[2]);
    $worksheet->setCellValue('NK92', $txtNumeroDocumento3[3]);
    $worksheet->setCellValue('NN92', $txtNumeroDocumento3[4]);
    $worksheet->setCellValue('NQ92', $txtNumeroDocumento3[5]);
    $worksheet->setCellValue('NT92', $txtNumeroDocumento3[6]);
    $worksheet->setCellValue('NW92', $txtNumeroDocumento3[7]);
    $worksheet->setCellValue('MW97', $txtPrimerApellido3);
    $worksheet->setCellValue('NY97', $txtSegundoApellido3);
    $worksheet->setCellValue('MW102', $txtNombres3);
    $fechanacimiento3 = $txtFechaNacimiento3;
    $fecha_nacimiento3= explode('/', $fechanacimiento3);
    $worksheet->setCellValue('NL105',$fecha_nacimiento3[0]);
    $worksheet->setCellValue('NP105',$fecha_nacimiento3[1]);
    $worksheet->setCellValue('NT105',$fecha_nacimiento3[2]);
    $worksheet->setCellValue('OG105', $txtPaisResidencia3);
    $worksheet->setCellValue('MO109', $txtRazonSocial3);
    $worksheet->setCellValue('ME112', $txtTipoVehiculo3);
    $time = time();
    $worksheet->setCellValue('ON112', date("d", $time));
    $worksheet->setCellValue('OR112', date("m", $time));
    $worksheet->setCellValue('OV112', date("Y", $time));
    $worksheet->setCellValue('MD117', $txtPorcentajeParticipacion3);    
      ////////////////////4/////////////////////////////////////////
      $txtTipoDocumento4  = $_POST['txtTipoDocumento4'];
      $txtNumeroDocumento4  = $_POST['txtNumeroDocumento4'];
      $txtPrimerApellido4  = $_POST['txtPrimerApellido4'];
      $txtSegundoApellido4  = $_POST['txtSegundoApellido4'];
      $txtNombres4 = $_POST['txtNombres4'];
      $txtFechaNacimiento4 = $_POST['txtFechaNacimiento4'];
      $txtPaisResidencia4 = $_POST['txtPaisResidencia4'];
      $txtRazonSocial4 = $_POST['txtRazonSocial4'];
      $txtTipoVehiculo4 = $_POST['txtTipoVehiculo4'];
      $txtPorcentajeParticipacion4  = $_POST['txtPorcentajeParticipacion4'];
  
      if($txtTipoDocumento4=="Carnet de Extranjeria"){
          $worksheet->setCellValue('MS134', "X");
      }
      if($txtTipoDocumento4=="Documento Nacional de Identidad"){
          $worksheet->setCellValue('MS132', "X");
      }
      if($txtTipoDocumento4=="Pasaporte"){
          $worksheet->setCellValue('MS136', "X");
      }
  
      $worksheet->setCellValue('NB131', $txtNumeroDocumento4[0]);
      $worksheet->setCellValue('NE131', $txtNumeroDocumento4[1]);
      $worksheet->setCellValue('NH131', $txtNumeroDocumento4[2]);
      $worksheet->setCellValue('NK131', $txtNumeroDocumento4[3]);
      $worksheet->setCellValue('NN131', $txtNumeroDocumento4[4]);
      $worksheet->setCellValue('NQ131', $txtNumeroDocumento4[5]);
      $worksheet->setCellValue('NT131', $txtNumeroDocumento4[6]);
      $worksheet->setCellValue('NW131', $txtNumeroDocumento4[7]);
      $worksheet->setCellValue('MW136', $txtPrimerApellido4);
      $worksheet->setCellValue('NY136', $txtSegundoApellido4);
      $worksheet->setCellValue('MW141', $txtNombres4);
      $fechanacimiento4 = $txtFechaNacimiento4;
      $fecha_nacimiento4= explode('/', $fechanacimiento4);
      $worksheet->setCellValue('NL144',$fecha_nacimiento4[0]);
      $worksheet->setCellValue('NP144',$fecha_nacimiento4[1]);
      $worksheet->setCellValue('NT144',$fecha_nacimiento4[2]);
      $worksheet->setCellValue('OG144', $txtPaisResidencia4);
      $worksheet->setCellValue('MO148', $txtRazonSocial4);
      $worksheet->setCellValue('ME151', $txtTipoVehiculo4);
      $time = time();
      $worksheet->setCellValue('ON151', date("d", $time));
      $worksheet->setCellValue('OR151', date("m", $time));
      $worksheet->setCellValue('OV151', date("Y", $time));
      $worksheet->setCellValue('MD156', $txtPorcentajeParticipacion4);
    ////////////////////5/////////////////////////////////////////
    $txtTipoDocumento5  = $_POST['txtTipoDocumento5'];
    $txtNumeroDocumento5  = $_POST['txtNumeroDocumento5'];
    $txtPrimerApellido5  = $_POST['txtPrimerApellido5'];
    $txtSegundoApellido5  = $_POST['txtSegundoApellido5'];
    $txtNombres5 = $_POST['txtNombres5'];
    $txtFechaNacimiento5 = $_POST['txtFechaNacimiento5'];
    $txtPaisResidencia5 = $_POST['txtPaisResidencia5'];
    $txtRazonSocial5 = $_POST['txtRazonSocial5'];
    $txtTipoVehiculo5 = $_POST['txtTipoVehiculo5'];
    $txtPorcentajeParticipacion5  = $_POST['txtPorcentajeParticipacion5'];

    if($txtTipoDocumento4=="Carnet de Extranjeria"){
        $worksheet->setCellValue('QT17', "X");
    }
    if($txtTipoDocumento4=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('QT15', "X");
    }
    if($txtTipoDocumento4=="Pasaporte"){
        $worksheet->setCellValue('QT19', "X");
    }

    $worksheet->setCellValue('RC14', $txtNumeroDocumento5[0]);
    $worksheet->setCellValue('RF14', $txtNumeroDocumento5[1]);
    $worksheet->setCellValue('RI14', $txtNumeroDocumento5[2]);
    $worksheet->setCellValue('RL14', $txtNumeroDocumento5[3]);
    $worksheet->setCellValue('RO14', $txtNumeroDocumento5[4]);
    $worksheet->setCellValue('RR14', $txtNumeroDocumento5[5]);
    $worksheet->setCellValue('RU14', $txtNumeroDocumento5[6]);
    $worksheet->setCellValue('RX14', $txtNumeroDocumento5[7]);
    $worksheet->setCellValue('QX19', $txtPrimerApellido5);
    $worksheet->setCellValue('RZ19', $txtSegundoApellido5);
    $worksheet->setCellValue('QX24', $txtNombres5);
    $fechanacimiento5 = $txtFechaNacimiento5;
    $fecha_nacimiento5= explode('/', $fechanacimiento5);
    $worksheet->setCellValue('RM27',$fecha_nacimiento5[0]);
    $worksheet->setCellValue('RQ27',$fecha_nacimiento5[1]);
    $worksheet->setCellValue('RU27',$fecha_nacimiento5[2]);
    $worksheet->setCellValue('SH27', $txtPaisResidencia5);
    $worksheet->setCellValue('QP31', $txtRazonSocial5);
    $worksheet->setCellValue('QF34', $txtTipoVehiculo5);
    $time = time();
    $worksheet->setCellValue('SO34', date("d", $time));
    $worksheet->setCellValue('SS34', date("m", $time));
    $worksheet->setCellValue('SW34', date("Y", $time));
    $worksheet->setCellValue('QE39', $txtPorcentajeParticipacion5);
        ////////////////////5/////////////////////////////////////////
    $txtTipoDocumento5  = $_POST['txtTipoDocumento5'];
    $txtNumeroDocumento5  = $_POST['txtNumeroDocumento5'];
    $txtPrimerApellido5  = $_POST['txtPrimerApellido5'];
    $txtSegundoApellido5  = $_POST['txtSegundoApellido5'];
    $txtNombres5 = $_POST['txtNombres5'];
    $txtFechaNacimiento5 = $_POST['txtFechaNacimiento5'];
    $txtPaisResidencia5 = $_POST['txtPaisResidencia5'];
    $txtRazonSocial5 = $_POST['txtRazonSocial5'];
    $txtTipoVehiculo5 = $_POST['txtTipoVehiculo5'];
    $txtPorcentajeParticipacion5  = $_POST['txtPorcentajeParticipacion5'];

    if($txtTipoDocumento5=="Carnet de Extranjeria"){
        $worksheet->setCellValue('QT17', "X");
    }
    if($txtTipoDocumento5=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('QT15', "X");
    }
    if($txtTipoDocumento5=="Pasaporte"){
        $worksheet->setCellValue('QT19', "X");
    }

    $worksheet->setCellValue('RC14', $txtNumeroDocumento5[0]);
    $worksheet->setCellValue('RF14', $txtNumeroDocumento5[1]);
    $worksheet->setCellValue('RI14', $txtNumeroDocumento5[2]);
    $worksheet->setCellValue('RL14', $txtNumeroDocumento5[3]);
    $worksheet->setCellValue('RO14', $txtNumeroDocumento5[4]);
    $worksheet->setCellValue('RR14', $txtNumeroDocumento5[5]);
    $worksheet->setCellValue('RU14', $txtNumeroDocumento5[6]);
    $worksheet->setCellValue('RX14', $txtNumeroDocumento5[7]);
    $worksheet->setCellValue('QX19', $txtPrimerApellido5);
    $worksheet->setCellValue('RZ19', $txtSegundoApellido5);
    $worksheet->setCellValue('QX24', $txtNombres5);
    $fechanacimiento5 = $txtFechaNacimiento5;
    $fecha_nacimiento5= explode('/', $fechanacimiento5);
    $worksheet->setCellValue('RM27',$fecha_nacimiento5[0]);
    $worksheet->setCellValue('RQ27',$fecha_nacimiento5[1]);
    $worksheet->setCellValue('RU27',$fecha_nacimiento5[2]);
    $worksheet->setCellValue('SH27', $txtPaisResidencia5);
    $worksheet->setCellValue('QP31', $txtRazonSocial5);
    $worksheet->setCellValue('QF34', $txtTipoVehiculo5);
    $time = time();
    $worksheet->setCellValue('SO34', date("d", $time));
    $worksheet->setCellValue('SS34', date("m", $time));
    $worksheet->setCellValue('SW34', date("Y", $time));
    $worksheet->setCellValue('QE39', $txtPorcentajeParticipacion5);
    ////////////////////6/////////////////////////////////////////
    $txtTipoDocumento6  = $_POST['txtTipoDocumento6'];
    $txtNumeroDocumento6  = $_POST['txtNumeroDocumento6'];
    $txtPrimerApellido6  = $_POST['txtPrimerApellido6'];
    $txtSegundoApellido6  = $_POST['txtSegundoApellido6'];
    $txtNombres6 = $_POST['txtNombres6'];
    $txtFechaNacimiento6 = $_POST['txtFechaNacimiento6'];
    $txtPaisResidencia6 = $_POST['txtPaisResidencia6'];
    $txtRazonSocial6 = $_POST['txtRazonSocial6'];
    $txtTipoVehiculo6 = $_POST['txtTipoVehiculo6'];
    $txtPorcentajeParticipacion6  = $_POST['txtPorcentajeParticipacion6'];

    if($txtTipoDocumento6=="Carnet de Extranjeria"){
        $worksheet->setCellValue('QT56', "X");
    }
    if($txtTipoDocumento6=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('QT54', "X");
    }
    if($txtTipoDocumento6=="Pasaporte"){
        $worksheet->setCellValue('QT58', "X");
    }

    $worksheet->setCellValue('RC53', $txtNumeroDocumento6[0]);
    $worksheet->setCellValue('RF53', $txtNumeroDocumento6[1]);
    $worksheet->setCellValue('RI53', $txtNumeroDocumento6[2]);
    $worksheet->setCellValue('RL53', $txtNumeroDocumento6[3]);
    $worksheet->setCellValue('RO53', $txtNumeroDocumento6[4]);
    $worksheet->setCellValue('RR53', $txtNumeroDocumento6[5]);
    $worksheet->setCellValue('RU53', $txtNumeroDocumento6[6]);
    $worksheet->setCellValue('RX53', $txtNumeroDocumento6[7]);
    $worksheet->setCellValue('QX58', $txtPrimerApellido6);
    $worksheet->setCellValue('RZ58', $txtSegundoApellido6);
    $worksheet->setCellValue('QX63', $txtNombres6);
    $fechanacimiento6 = $txtFechaNacimiento6;
    $fecha_nacimiento6= explode('/', $fechanacimiento6);
    $worksheet->setCellValue('RM66',$fecha_nacimiento6[0]);
    $worksheet->setCellValue('RQ66',$fecha_nacimiento6[1]);
    $worksheet->setCellValue('RU66',$fecha_nacimiento6[2]);
    $worksheet->setCellValue('SH66', $txtPaisResidencia6);
    $worksheet->setCellValue('QP70', $txtRazonSocial6);
    $worksheet->setCellValue('QF73', $txtTipoVehiculo6);
    $time = time();
    $worksheet->setCellValue('SO73', date("d", $time));
    $worksheet->setCellValue('SS73', date("m", $time));
    $worksheet->setCellValue('SW73', date("Y", $time));
    $worksheet->setCellValue('QE78', $txtPorcentajeParticipacion6);
}
if($NumeroSocios=="7"){
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtTipoDocumento  = $_POST['txtTipoDocumento'];
    $txtNumeroDocumento  = $_POST['txtNumeroDocumento'];
    $txtPrimerApellido  = $_POST['txtPrimerApellido'];
    $txtSegundoApellido  = $_POST['txtSegundoApellido'];
    $txtNombres = $_POST['txtNombres'];
    $txtFechaNacimiento = $_POST['txtFechaNacimiento'];
    $txtPaisResidencia = $_POST['txtPaisResidencia'];
    $txtRazonSocial = $_POST['txtRazonSocial'];
    $txtTipoVehiculo = $_POST['txtTipoVehiculo'];
    $txtPorcentajeParticipacion  = $_POST['txtPorcentajeParticipacion'];
    
    if($txtTipoDocumento=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS17', "X");
    }
    if($txtTipoDocumento=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS15', "X");
    }
    if($txtTipoDocumento=="Pasaporte"){
        $worksheet->setCellValue('MS19', "X");
    }
    
    $worksheet->setCellValue('NB14', $txtNumeroDocumento[0]);
    $worksheet->setCellValue('NE14', $txtNumeroDocumento[1]);
    $worksheet->setCellValue('NH14', $txtNumeroDocumento[2]);
    $worksheet->setCellValue('NK14', $txtNumeroDocumento[3]);
    $worksheet->setCellValue('NN14', $txtNumeroDocumento[4]);
    $worksheet->setCellValue('NQ14', $txtNumeroDocumento[5]);
    $worksheet->setCellValue('NT14', $txtNumeroDocumento[6]);
    $worksheet->setCellValue('NW14', $txtNumeroDocumento[7]);
    $worksheet->setCellValue('MW19', $txtPrimerApellido);
    $worksheet->setCellValue('NY19', $txtSegundoApellido);
    $worksheet->setCellValue('MW24', $txtNombres);
    $fechanacimiento = $txtFechaNacimiento;
    $fecha_nacimiento = explode('/', $fechanacimiento);
    $worksheet->setCellValue('NL27',$fecha_nacimiento[0]);
    $worksheet->setCellValue('NP27',$fecha_nacimiento[1]);
    $worksheet->setCellValue('NT27',$fecha_nacimiento[2]);
    $worksheet->setCellValue('OG27', $txtPaisResidencia);
    $worksheet->setCellValue('MO31', $txtRazonSocial);
    $worksheet->setCellValue('ME34', $txtTipoVehiculo);
    $time = time();
    $worksheet->setCellValue('ON34', date("d", $time));
    $worksheet->setCellValue('OR34', date("m", $time));
    $worksheet->setCellValue('OV34', date("Y", $time));
    $worksheet->setCellValue('MD39', $txtPorcentajeParticipacion);
////////////////////////////////////2//////////////////////////////////

    $txtTipoDocumento2  = $_POST['txtTipoDocumento2'];
    $txtNumeroDocumento2  = $_POST['txtNumeroDocumento2'];
    $txtPrimerApellido2  = $_POST['txtPrimerApellido2'];
    $txtSegundoApellido2  = $_POST['txtSegundoApellido2'];
    $txtNombres2 = $_POST['txtNombres2'];
    $txtFechaNacimiento2 = $_POST['txtFechaNacimiento2'];
    $txtPaisResidencia2 = $_POST['txtPaisResidencia2'];
    $txtRazonSocial2 = $_POST['txtRazonSocial2'];
    $txtTipoVehiculo2 = $_POST['txtTipoVehiculo2'];
    $txtPorcentajeParticipacion2  = $_POST['txtPorcentajeParticipacion2'];

    if($txtTipoDocumento2=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS56', "X");
    }
    if($txtTipoDocumento2=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS54', "X");
    }
    if($txtTipoDocumento2=="Pasaporte"){
        $worksheet->setCellValue('MS58', "X");
    }

    $worksheet->setCellValue('NB53', $txtNumeroDocumento2[0]);
    $worksheet->setCellValue('NE53', $txtNumeroDocumento2[1]);
    $worksheet->setCellValue('NH53', $txtNumeroDocumento2[2]);
    $worksheet->setCellValue('NK53', $txtNumeroDocumento2[3]);
    $worksheet->setCellValue('NN53', $txtNumeroDocumento2[4]);
    $worksheet->setCellValue('NQ53', $txtNumeroDocumento2[5]);
    $worksheet->setCellValue('NT53', $txtNumeroDocumento2[6]);
    $worksheet->setCellValue('NW53', $txtNumeroDocumento2[7]);
    $worksheet->setCellValue('MW58', $txtPrimerApellido2);
    $worksheet->setCellValue('NY58', $txtSegundoApellido2);
    $worksheet->setCellValue('MW63', $txtNombres2);
    $fechanacimiento2 = $txtFechaNacimiento2;
    $fecha_nacimiento2= explode('/', $fechanacimiento2);
    $worksheet->setCellValue('NL66',$fecha_nacimiento2[0]);
    $worksheet->setCellValue('NP66',$fecha_nacimiento2[1]);
    $worksheet->setCellValue('NT66',$fecha_nacimiento2[2]);
    $worksheet->setCellValue('OG66', $txtPaisResidencia2);
    $worksheet->setCellValue('MO70', $txtRazonSocial2);
    $worksheet->setCellValue('ME73', $txtTipoVehiculo2);
    $time = time();
    $worksheet->setCellValue('ON73', date("d", $time));
    $worksheet->setCellValue('OR73', date("m", $time));
    $worksheet->setCellValue('OV73', date("Y", $time));
    $worksheet->setCellValue('MD78', $txtPorcentajeParticipacion2);

    ////////////////////3/////////////////////////////////////////
    $txtTipoDocumento3  = $_POST['txtTipoDocumento3'];
    $txtNumeroDocumento3  = $_POST['txtNumeroDocumento3'];
    $txtPrimerApellido3  = $_POST['txtPrimerApellido3'];
    $txtSegundoApellido3  = $_POST['txtSegundoApellido3'];
    $txtNombres3 = $_POST['txtNombres3'];
    $txtFechaNacimiento3 = $_POST['txtFechaNacimiento3'];
    $txtPaisResidencia3 = $_POST['txtPaisResidencia3'];
    $txtRazonSocial3 = $_POST['txtRazonSocial3'];
    $txtTipoVehiculo3 = $_POST['txtTipoVehiculo3'];
    $txtPorcentajeParticipacion3  = $_POST['txtPorcentajeParticipacion3'];

    if($txtTipoDocumento3=="Carnet de Extranjeria"){
        $worksheet->setCellValue('MS95', "X");
    }
    if($txtTipoDocumento3=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('MS93', "X");
    }
    if($txtTipoDocumento3=="Pasaporte"){
        $worksheet->setCellValue('MS97', "X");
    }

    $worksheet->setCellValue('NB92', $txtNumeroDocumento3[0]);
    $worksheet->setCellValue('NE92', $txtNumeroDocumento3[1]);
    $worksheet->setCellValue('NH92', $txtNumeroDocumento3[2]);
    $worksheet->setCellValue('NK92', $txtNumeroDocumento3[3]);
    $worksheet->setCellValue('NN92', $txtNumeroDocumento3[4]);
    $worksheet->setCellValue('NQ92', $txtNumeroDocumento3[5]);
    $worksheet->setCellValue('NT92', $txtNumeroDocumento3[6]);
    $worksheet->setCellValue('NW92', $txtNumeroDocumento3[7]);
    $worksheet->setCellValue('MW97', $txtPrimerApellido3);
    $worksheet->setCellValue('NY97', $txtSegundoApellido3);
    $worksheet->setCellValue('MW102', $txtNombres3);
    $fechanacimiento3 = $txtFechaNacimiento3;
    $fecha_nacimiento3= explode('/', $fechanacimiento3);
    $worksheet->setCellValue('NL105',$fecha_nacimiento3[0]);
    $worksheet->setCellValue('NP105',$fecha_nacimiento3[1]);
    $worksheet->setCellValue('NT105',$fecha_nacimiento3[2]);
    $worksheet->setCellValue('OG105', $txtPaisResidencia3);
    $worksheet->setCellValue('MO109', $txtRazonSocial3);
    $worksheet->setCellValue('ME112', $txtTipoVehiculo3);
    $time = time();
    $worksheet->setCellValue('ON112', date("d", $time));
    $worksheet->setCellValue('OR112', date("m", $time));
    $worksheet->setCellValue('OV112', date("Y", $time));
    $worksheet->setCellValue('MD117', $txtPorcentajeParticipacion3);    
      ////////////////////4/////////////////////////////////////////
      $txtTipoDocumento4  = $_POST['txtTipoDocumento4'];
      $txtNumeroDocumento4  = $_POST['txtNumeroDocumento4'];
      $txtPrimerApellido4  = $_POST['txtPrimerApellido4'];
      $txtSegundoApellido4  = $_POST['txtSegundoApellido4'];
      $txtNombres4 = $_POST['txtNombres4'];
      $txtFechaNacimiento4 = $_POST['txtFechaNacimiento4'];
      $txtPaisResidencia4 = $_POST['txtPaisResidencia4'];
      $txtRazonSocial4 = $_POST['txtRazonSocial4'];
      $txtTipoVehiculo4 = $_POST['txtTipoVehiculo4'];
      $txtPorcentajeParticipacion4  = $_POST['txtPorcentajeParticipacion4'];
  
      if($txtTipoDocumento4=="Carnet de Extranjeria"){
          $worksheet->setCellValue('MS134', "X");
      }
      if($txtTipoDocumento4=="Documento Nacional de Identidad"){
          $worksheet->setCellValue('MS132', "X");
      }
      if($txtTipoDocumento4=="Pasaporte"){
          $worksheet->setCellValue('MS136', "X");
      }
  
      $worksheet->setCellValue('NB131', $txtNumeroDocumento4[0]);
      $worksheet->setCellValue('NE131', $txtNumeroDocumento4[1]);
      $worksheet->setCellValue('NH131', $txtNumeroDocumento4[2]);
      $worksheet->setCellValue('NK131', $txtNumeroDocumento4[3]);
      $worksheet->setCellValue('NN131', $txtNumeroDocumento4[4]);
      $worksheet->setCellValue('NQ131', $txtNumeroDocumento4[5]);
      $worksheet->setCellValue('NT131', $txtNumeroDocumento4[6]);
      $worksheet->setCellValue('NW131', $txtNumeroDocumento4[7]);
      $worksheet->setCellValue('MW136', $txtPrimerApellido4);
      $worksheet->setCellValue('NY136', $txtSegundoApellido4);
      $worksheet->setCellValue('MW141', $txtNombres4);
      $fechanacimiento4 = $txtFechaNacimiento4;
      $fecha_nacimiento4= explode('/', $fechanacimiento4);
      $worksheet->setCellValue('NL144',$fecha_nacimiento4[0]);
      $worksheet->setCellValue('NP144',$fecha_nacimiento4[1]);
      $worksheet->setCellValue('NT144',$fecha_nacimiento4[2]);
      $worksheet->setCellValue('OG144', $txtPaisResidencia4);
      $worksheet->setCellValue('MO148', $txtRazonSocial4);
      $worksheet->setCellValue('ME151', $txtTipoVehiculo4);
      $time = time();
      $worksheet->setCellValue('ON151', date("d", $time));
      $worksheet->setCellValue('OR151', date("m", $time));
      $worksheet->setCellValue('OV151', date("Y", $time));
      $worksheet->setCellValue('MD156', $txtPorcentajeParticipacion4);
    ////////////////////5/////////////////////////////////////////
    $txtTipoDocumento5  = $_POST['txtTipoDocumento5'];
    $txtNumeroDocumento5  = $_POST['txtNumeroDocumento5'];
    $txtPrimerApellido5  = $_POST['txtPrimerApellido5'];
    $txtSegundoApellido5  = $_POST['txtSegundoApellido5'];
    $txtNombres5 = $_POST['txtNombres5'];
    $txtFechaNacimiento5 = $_POST['txtFechaNacimiento5'];
    $txtPaisResidencia5 = $_POST['txtPaisResidencia5'];
    $txtRazonSocial5 = $_POST['txtRazonSocial5'];
    $txtTipoVehiculo5 = $_POST['txtTipoVehiculo5'];
    $txtPorcentajeParticipacion5  = $_POST['txtPorcentajeParticipacion5'];

    if($txtTipoDocumento4=="Carnet de Extranjeria"){
        $worksheet->setCellValue('QT17', "X");
    }
    if($txtTipoDocumento4=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('QT15', "X");
    }
    if($txtTipoDocumento4=="Pasaporte"){
        $worksheet->setCellValue('QT19', "X");
    }

    $worksheet->setCellValue('RC14', $txtNumeroDocumento5[0]);
    $worksheet->setCellValue('RF14', $txtNumeroDocumento5[1]);
    $worksheet->setCellValue('RI14', $txtNumeroDocumento5[2]);
    $worksheet->setCellValue('RL14', $txtNumeroDocumento5[3]);
    $worksheet->setCellValue('RO14', $txtNumeroDocumento5[4]);
    $worksheet->setCellValue('RR14', $txtNumeroDocumento5[5]);
    $worksheet->setCellValue('RU14', $txtNumeroDocumento5[6]);
    $worksheet->setCellValue('RX14', $txtNumeroDocumento5[7]);
    $worksheet->setCellValue('QX19', $txtPrimerApellido5);
    $worksheet->setCellValue('RZ19', $txtSegundoApellido5);
    $worksheet->setCellValue('QX24', $txtNombres5);
    $fechanacimiento5 = $txtFechaNacimiento5;
    $fecha_nacimiento5= explode('/', $fechanacimiento5);
    $worksheet->setCellValue('RM27',$fecha_nacimiento5[0]);
    $worksheet->setCellValue('RQ27',$fecha_nacimiento5[1]);
    $worksheet->setCellValue('RU27',$fecha_nacimiento5[2]);
    $worksheet->setCellValue('SH27', $txtPaisResidencia5);
    $worksheet->setCellValue('QP31', $txtRazonSocial5);
    $worksheet->setCellValue('QF34', $txtTipoVehiculo5);
    $time = time();
    $worksheet->setCellValue('SO34', date("d", $time));
    $worksheet->setCellValue('SS34', date("m", $time));
    $worksheet->setCellValue('SW34', date("Y", $time));
    $worksheet->setCellValue('QE39', $txtPorcentajeParticipacion5);
        ////////////////////5/////////////////////////////////////////
    $txtTipoDocumento5  = $_POST['txtTipoDocumento5'];
    $txtNumeroDocumento5  = $_POST['txtNumeroDocumento5'];
    $txtPrimerApellido5  = $_POST['txtPrimerApellido5'];
    $txtSegundoApellido5  = $_POST['txtSegundoApellido5'];
    $txtNombres5 = $_POST['txtNombres5'];
    $txtFechaNacimiento5 = $_POST['txtFechaNacimiento5'];
    $txtPaisResidencia5 = $_POST['txtPaisResidencia5'];
    $txtRazonSocial5 = $_POST['txtRazonSocial5'];
    $txtTipoVehiculo5 = $_POST['txtTipoVehiculo5'];
    $txtPorcentajeParticipacion5  = $_POST['txtPorcentajeParticipacion5'];

    if($txtTipoDocumento5=="Carnet de Extranjeria"){
        $worksheet->setCellValue('QT17', "X");
    }
    if($txtTipoDocumento5=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('QT15', "X");
    }
    if($txtTipoDocumento5=="Pasaporte"){
        $worksheet->setCellValue('QT19', "X");
    }

    $worksheet->setCellValue('RC14', $txtNumeroDocumento5[0]);
    $worksheet->setCellValue('RF14', $txtNumeroDocumento5[1]);
    $worksheet->setCellValue('RI14', $txtNumeroDocumento5[2]);
    $worksheet->setCellValue('RL14', $txtNumeroDocumento5[3]);
    $worksheet->setCellValue('RO14', $txtNumeroDocumento5[4]);
    $worksheet->setCellValue('RR14', $txtNumeroDocumento5[5]);
    $worksheet->setCellValue('RU14', $txtNumeroDocumento5[6]);
    $worksheet->setCellValue('RX14', $txtNumeroDocumento5[7]);
    $worksheet->setCellValue('QX19', $txtPrimerApellido5);
    $worksheet->setCellValue('RZ19', $txtSegundoApellido5);
    $worksheet->setCellValue('QX24', $txtNombres5);
    $fechanacimiento5 = $txtFechaNacimiento5;
    $fecha_nacimiento5= explode('/', $fechanacimiento5);
    $worksheet->setCellValue('RM27',$fecha_nacimiento5[0]);
    $worksheet->setCellValue('RQ27',$fecha_nacimiento5[1]);
    $worksheet->setCellValue('RU27',$fecha_nacimiento5[2]);
    $worksheet->setCellValue('SH27', $txtPaisResidencia5);
    $worksheet->setCellValue('QP31', $txtRazonSocial5);
    $worksheet->setCellValue('QF34', $txtTipoVehiculo5);
    $time = time();
    $worksheet->setCellValue('SO34', date("d", $time));
    $worksheet->setCellValue('SS34', date("m", $time));
    $worksheet->setCellValue('SW34', date("Y", $time));
    $worksheet->setCellValue('QE39', $txtPorcentajeParticipacion5);
    ////////////////////6/////////////////////////////////////////
    $txtTipoDocumento6  = $_POST['txtTipoDocumento6'];
    $txtNumeroDocumento6  = $_POST['txtNumeroDocumento6'];
    $txtPrimerApellido6  = $_POST['txtPrimerApellido6'];
    $txtSegundoApellido6  = $_POST['txtSegundoApellido6'];
    $txtNombres6 = $_POST['txtNombres6'];
    $txtFechaNacimiento6 = $_POST['txtFechaNacimiento6'];
    $txtPaisResidencia6 = $_POST['txtPaisResidencia6'];
    $txtRazonSocial6 = $_POST['txtRazonSocial6'];
    $txtTipoVehiculo6 = $_POST['txtTipoVehiculo6'];
    $txtPorcentajeParticipacion6  = $_POST['txtPorcentajeParticipacion6'];

    if($txtTipoDocumento6=="Carnet de Extranjeria"){
        $worksheet->setCellValue('QT56', "X");
    }
    if($txtTipoDocumento6=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('QT54', "X");
    }
    if($txtTipoDocumento6=="Pasaporte"){
        $worksheet->setCellValue('QT58', "X");
    }

    $worksheet->setCellValue('RC53', $txtNumeroDocumento6[0]);
    $worksheet->setCellValue('RF53', $txtNumeroDocumento6[1]);
    $worksheet->setCellValue('RI53', $txtNumeroDocumento6[2]);
    $worksheet->setCellValue('RL53', $txtNumeroDocumento6[3]);
    $worksheet->setCellValue('RO53', $txtNumeroDocumento6[4]);
    $worksheet->setCellValue('RR53', $txtNumeroDocumento6[5]);
    $worksheet->setCellValue('RU53', $txtNumeroDocumento6[6]);
    $worksheet->setCellValue('RX53', $txtNumeroDocumento6[7]);
    $worksheet->setCellValue('QX58', $txtPrimerApellido6);
    $worksheet->setCellValue('RZ58', $txtSegundoApellido6);
    $worksheet->setCellValue('QX63', $txtNombres6);
    $fechanacimiento6 = $txtFechaNacimiento6;
    $fecha_nacimiento6= explode('/', $fechanacimiento6);
    $worksheet->setCellValue('RM66',$fecha_nacimiento6[0]);
    $worksheet->setCellValue('RQ66',$fecha_nacimiento6[1]);
    $worksheet->setCellValue('RU66',$fecha_nacimiento6[2]);
    $worksheet->setCellValue('SH66', $txtPaisResidencia6);
    $worksheet->setCellValue('QP70', $txtRazonSocial6);
    $worksheet->setCellValue('QF73', $txtTipoVehiculo6);
    $time = time();
    $worksheet->setCellValue('SO73', date("d", $time));
    $worksheet->setCellValue('SS73', date("m", $time));
    $worksheet->setCellValue('SW73', date("Y", $time));
    $worksheet->setCellValue('QE78', $txtPorcentajeParticipacion6);
    ////////////////////7/////////////////////////////////////////
    $txtTipoDocumento7  = $_POST['txtTipoDocumento7'];
    $txtNumeroDocumento7  = $_POST['txtNumeroDocumento7'];
    $txtPrimerApellido7  = $_POST['txtPrimerApellido7'];
    $txtSegundoApellido7  = $_POST['txtSegundoApellido7'];
    $txtNombres7 = $_POST['txtNombres7'];
    $txtFechaNacimiento7 = $_POST['txtFechaNacimiento7'];
    $txtPaisResidencia7 = $_POST['txtPaisResidencia7'];
    $txtRazonSocial7 = $_POST['txtRazonSocial7'];
    $txtTipoVehiculo7 = $_POST['txtTipoVehiculo7'];
    $txtPorcentajeParticipacion7  = $_POST['txtPorcentajeParticipacion7'];

    if($txtTipoDocumento7=="Carnet de Extranjeria"){
        $worksheet->setCellValue('QT95', "X");
    }
    if($txtTipoDocumento7=="Documento Nacional de Identidad"){
        $worksheet->setCellValue('QT93', "X");
    }
    if($txtTipoDocumento7=="Pasaporte"){
        $worksheet->setCellValue('QT97', "X");
    }

    $worksheet->setCellValue('RC92', $txtNumeroDocumento7[0]);
    $worksheet->setCellValue('RF92', $txtNumeroDocumento7[1]);
    $worksheet->setCellValue('RI92', $txtNumeroDocumento7[2]);
    $worksheet->setCellValue('RL92', $txtNumeroDocumento7[3]);
    $worksheet->setCellValue('RO92', $txtNumeroDocumento7[4]);
    $worksheet->setCellValue('RR92', $txtNumeroDocumento7[5]);
    $worksheet->setCellValue('RU92', $txtNumeroDocumento7[6]);
    $worksheet->setCellValue('RX92', $txtNumeroDocumento7[7]);
    $worksheet->setCellValue('QX97', $txtPrimerApellido7);
    $worksheet->setCellValue('RZ97', $txtSegundoApellido7);
    $worksheet->setCellValue('QX102', $txtNombres7);
    $fechanacimiento7 = $txtFechaNacimiento7;
    $fecha_nacimiento7= explode('/', $fechanacimiento7);
    $worksheet->setCellValue('RM105',$fecha_nacimiento7[0]);
    $worksheet->setCellValue('RQ105',$fecha_nacimiento7[1]);
    $worksheet->setCellValue('RU105',$fecha_nacimiento7[2]);
    $worksheet->setCellValue('SH105', $txtPaisResidencia7);
    $worksheet->setCellValue('QP109', $txtRazonSocial7);
    $worksheet->setCellValue('QF112', $txtTipoVehiculo7);
    $time = time();
    $worksheet->setCellValue('SO112', date("d", $time));
    $worksheet->setCellValue('SS112', date("m", $time));
    $worksheet->setCellValue('SW112', date("Y", $time));
    $worksheet->setCellValue('QE117', $txtPorcentajeParticipacion7);
}
$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, "Xlsx");
$archivo = $txtNumeroDocumento_Anterior."-4-GUIA_PERSONA_JURIDICA_14052020_01.xlsx";
$writer->save($archivo);
echo '<a href="'.$archivo.'">Descargar GUIA_PERSONA_SIN_NEGOCIO_14052020</a>';
?>
