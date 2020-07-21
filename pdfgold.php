<?php

require('fpdf/fpdf.php');
class PDF extends FPDF {
    // Cabecera de página
    function Header()
    {
        //Title
        $this->SetTitle('ProCont Businness Precio del Oro');
        // Logo
        $this->Image('logo.jpeg',10,8,33);
        // Arial bold 15
        $this->SetFont('Arial','B',15);
        // Movernos a la derecha
        $this->Cell(80);
        // Título
        $this->Cell(30,10,'Calculadora de Precio de Oro',0,0,'C');
        // Salto de línea
        $this->Ln(20);
    }

    // Pie de página
    function Footer()
    {
        // Posición: a 1,5 cm del final
        $this->SetY(-15);
        // Arial italic 8
        $this->SetFont('Arial','I',8);
        // Número de página
        //$this->Cell(0,10,'Page '.$this->PageNo().'/{nb}',0,0,'C');
    }

    
}

// Creación del objeto de la clase heredada
$pdf = new PDF();
$pdf->AliasNbPages();
$pdf->AddPage();
$pdf->SetFont('Times','',12);

$sunat_precio_compra= $_POST['sunat_precio_compra'];
$sunat_precio_venta= $_POST['sunat_precio_venta'];
$kitco_am= $_POST['kitco_am'];
$kitco_pm= $_POST['kitco_pm'];

$ruc = $_POST['ruc'];
$razon_social = $_POST['razon_social'];
$numero_liq = $_POST['numero_liq'];
$precio_inter = $_POST['precio_inter'];
$tipo_cambio = $_POST['tipo_cambio'];
$peso_oro = $_POST['peso_oro'];
$pureza_oro = $_POST['pureza_oro'];
$descuento = $_POST['descuento'];
$detraccion = $_POST['detraccion'];
$precio_oro_peru = $_POST['precio_oro_peru'];

//Cell($w, $h=0, $txt='', $border=0, $ln=0, $align='', $fill=false, $link='')

$pdf->SetX(45);
$pdf->Cell(60, 10, 'GOLD' , 1, 0, 'C');
$pdf->Cell(60, 10, 'PRECIO DOLAR' , 1, 0, 'C');
$pdf->Ln();

$pdf->SetX(45);
$pdf->Cell(30, 10, 'AM' , 1, 0, 'C');
$pdf->Cell(30, 10, 'PM' , 1, 0, 'C');
$pdf->Cell(30, 10, 'COMPRA' , 1, 0, 'C');
$pdf->Cell(30, 10, 'VENTA' , 1, 0, 'C');
$pdf->Ln();

$pdf->SetX(45);
$pdf->Cell(30, 10, $kitco_am , 1, 0, 'C');
$pdf->Cell(30, 10, $kitco_pm , 1, 0, 'C');
$pdf->Cell(30, 10, $sunat_precio_compra , 1, 0, 'C');
$pdf->Cell(30, 10, $sunat_precio_venta , 1, 0, 'C');

$pdf->Ln(20);
//Ruc
$pdf->SetX(45);
$pdf->Cell(30, 10, 'RUC: ' . $ruc , 0, 0);
$pdf->Ln();

$pdf->SetX(45);
$pdf->MultiCell(120, 5, $razon_social , 0, 'L');


$pdf->SetX(45);
$pdf->Cell(50, 10, utf8_decode('Nª LIQUIDACIÓN: ' . $numero_liq) , 0, 0);

$pdf->Ln(15);

$pdf->Rect(45, $pdf->GetY(), 120, 70);
$pdf->SetX(45);
$pdf->Cell(60, 10, 'Precio internacional' , 0, 0);
$pdf->Cell(30, 10, $precio_inter , 1, 0, 'C');
$pdf->Cell(30, 10, utf8_decode('Dólares/Onza') , 0, 0, 'L');
$pdf->Ln();

$pdf->SetX(45);
$pdf->Cell(60, 10, 'Tipo de Cambio' , 0, 0);
$pdf->Cell(30, 10, $tipo_cambio , 1, 0, 'C');
$pdf->Cell(30, 10, utf8_decode('Soles/Dólar') , 0, 0, 'L');
$pdf->Ln();

$pdf->SetX(45);
$pdf->Cell(60, 10, 'Peso del Oro' , 0, 0);
$pdf->Cell(30, 10, $peso_oro , 1, 0, 'C');
$pdf->Cell(30, 10, 'Gramos' , 0, 0, 'L');
$pdf->Ln();

$pdf->SetX(45);
$pdf->Cell(60, 10, 'Pureza del Oro' , 0, 0);
$pdf->Cell(30, 10, $pureza_oro , 1, 0, 'C');
$pdf->Cell(30, 10, utf8_decode('Milésimos') , 0, 0, 'L');
$pdf->Ln();

$pdf->SetX(45);
$pdf->Cell(60, 10, 'Descuento' , 0, 0);
$pdf->Cell(30, 10, $descuento , 1, 0, 'C');
$pdf->Cell(30, 10, '%' , 0, 0, 'L');
$pdf->Ln();

$pdf->SetX(45);
$pdf->Cell(60, 10, utf8_decode('SPOT Detracción') , 0, 0);
$pdf->Cell(30, 10, $detraccion , 1, 0, 'C');
$pdf->Cell(30, 10, '%' , 0, 0, 'L');
$pdf->Ln();

$pdf->SetX(45);
$pdf->Cell(60, 10, utf8_decode('Precio del Oro Perú') , 0, 0);
$pdf->Cell(30, 10, $precio_oro_peru , 1, 0, 'C');
$pdf->Cell(30, 10, 'Soles' , 0, 0, 'L');
$pdf->Ln();


$pdf->Output('I','ProCont Businness Precio del Oro.pdf');

$archivo = "contador.txt";
$contador = 0;

$fp = fopen($archivo,"r");
$contador = fgets($fp, 26);
fclose($fp);

++$contador;

$fp = fopen($archivo,"w+");
fwrite($fp, $contador, 26);
fclose($fp);