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
    echo $nodes[0]->textContent;
}    
$nodes = $xpath->query('//div[2]/center/table/tr[last()]/td[last()-1]');
if ($nodes->length) {
    //Imprime la compra
    echo $nodes[0]->textContent;
}    
?>