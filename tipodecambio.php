<?php
/**
 * Autor: Jordan Diaz Diaz
 * Descripcion: Tipo de Cambio Sunat
 * Fecha: 21/07/2020
 */
libxml_use_internal_errors(true);
$dom = new DomDocument;
$dom->loadHtmlFile('http://www.sunat.gob.pe/cl-at-ittipcam/tcS01Alias');
$xpath = new DomXPath($dom);
$nodes = $xpath->query('//div[2]/center/table/tr[last()]/td[last()]');
if ($nodes->length) {
    //Imprime la venta
<<<<<<< HEAD
    echo $nodes[0]->textContent;
=======
    $precio_compra =  $nodes[0]->textContent;
>>>>>>> 0f98030231d97706329f86ba75da5339d68ece54
}    
$nodes = $xpath->query('//div[2]/center/table/tr[last()]/td[last()-1]');
if ($nodes->length) {
    //Imprime la compra
<<<<<<< HEAD
    echo $nodes[0]->textContent;
=======
    $precio_venta = $nodes[0]->textContent;
>>>>>>> 0f98030231d97706329f86ba75da5339d68ece54
}    
?>