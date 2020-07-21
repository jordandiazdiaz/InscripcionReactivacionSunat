<?php
/**
 * Autor: Jordan Diaz Diaz
 * Descripcion: Precio Horo kitco
 * Fecha: 21/07/2020
 */
libxml_use_internal_errors(true);
$dom = new DomDocument;
$dom->loadHtmlFile('https://www.kitco.com/gold.londonfix.html');
$xpath = new DomXPath($dom);
$nodes = $xpath->query('//div[3]/div[3]/div[1]/div[1]/div[1]/div[2]/table/tr[last()]/td[1]');
if ($nodes->length) {
    //Imprime la AM
    echo $nodes[0]->textContent;
}    
echo "---";
$nodes = $xpath->query('//div[3]/div[3]/div[1]/div[1]/div[1]/div[2]/table/tr[last()]/td[2]');
if ($nodes->length) {
    //Imprime la PM
    echo $nodes[0]->textContent;
}    
?>