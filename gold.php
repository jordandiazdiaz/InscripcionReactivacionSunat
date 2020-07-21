<?php
    include_once('tipodecambio.php');
    include_once('preciooro.php');


    $archivo = "contador.txt";
    $contador = 0;

    $fp = fopen($archivo,"r");
    $contador = fgets($fp, 26);
    fclose($fp);

    //++$contador;

    /*$fp = fopen($archivo,"w+");
    fwrite($fp, $contador, 26);
    fclose($fp);
*/

    if ($contador < 10) $liq = '0000';
    if ($contador >= 10 && $contador < 100) $liq = '000';
    if ($contador >= 100 && $contador < 1000) $liq = '00';
    if ($contador >= 1000 && $contador < 10000) $liq = '0';

    $liq = $liq . $contador;

?>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <meta http-equiv="X-UA-Compatible" content="IE=10; IE=9; IE=8; IE=7; IE=EDGE" />
    <title>
        PROCONT :: CALCULADORA DE PRECIO ORO
    </title>
    <link type="text/css" href="https://plaft.sbs.gob.pe/sisdel/css/bootstrap/bootstrap.min.css" rel="stylesheet" /><link type="text/css" href="https://plaft.sbs.gob.pe/sisdel/css/bootstrap/bootstrap-theme.min.css" rel="stylesheet" /><link type="text/css" href="https://plaft.sbs.gob.pe/sisdel/css/site.css" rel="stylesheet" /><link type="text/css" href="https://plaft.sbs.gob.pe/sisdel/css/bootstrap-typeahead.css" rel="stylesheet" />
    
    <!--[if lte IE 8]>        
       <script src="/sisdel/js/html5shiv.js"></script>
       <script src="/sisdel/js/respond.min.js"></script>  
    <![endif]-->
    
    <script type="text/javascript" src="https://plaft.sbs.gob.pe/sisdel/js/jquery-3.4.1.min.js"></script>
    <script type="text/javascript" src="https://plaft.sbs.gob.pe/sisdel/js/bootstrap/bootstrap.min.js"></script>
    <script type="text/javascript" src="https://plaft.sbs.gob.pe/sisdel/js/bootstrap-typeahead.min.js"></script>    
    <script type="text/javascript" src="https://plaft.sbs.gob.pe/sisdel/js/jquery.alphanum.js"></script>
    <script type="text/javascript" src="https://plaft.sbs.gob.pe/sisdel/js/site.js"></script>    
    <script src="https://www.google.com/recaptcha/api.js" type="text/javascript" ></script>

    <style>
        .navbar {
            padding: 0;
        }
        .brand {
            display: block;
            width: 100%;
            height: auto;
        }
        .t-width {
            width: 300px;
        }
        .search-ruc, .d-flex {
            display: flex;
        }
        .d-flex {
            margin-bottom: 1rem;
        }
        .first {
            margin-right: 1rem;
        }
        .form-group {
            margin-bottom: 0;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="row">
            <div class="col-md-12">
                <nav class="navbar">
                    <div class="navbar-logo" >                
                        <img alt="Logo" src="logo.jpg" class="brand">
                    </div>                    
                </nav>
                <form action="pdfgold.php" method="post" class="form-horizontal">
                    <div class="panel-group" id="gold">

                        <div class="d-flex">
                            <!--Panel: Precio del Oro-->
                            <div class="panel panel-default t-width first">
                                <div class="panel-heading  panel-heading-custom">
                                    <h3 class="panel-title">
                                        Precio Internacinal (Dólares/Onza)
                                    </h3>
                                </div>
                                <div class="panel-body">
                                    <table class="table table-bordered">
                                        <thead>
                                            <tr>
                                                <th>AM</th>
                                                <th>PM</th>
                                            </tr>
                                        </thead>
                                        <tfoot>
                                            <tr>
                                                <td colspan="2">Fuente: kitco</td>
                                            </tr>
                                        </tfoot>
                                        <tbody>
                                            <tr>
                                                <td><?= $am ?></td>
                                                <td><?= $pm ?></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                            </div>

                            <!--Panel: Precio del Oro-->
                            <div class="panel panel-default t-width">
                                <div class="panel-heading  panel-heading-custom">
                                    <h3 class="panel-title">
                                        Tipo de Cambio
                                    </h3>
                                </div>
                                <div class="panel-body">
                                    <table class="table table-bordered">
                                        <thead>
                                            <tr>
                                                <th>COMPRA</th>
                                                <th>VENTA</th>
                                            </tr>
                                        </thead>
                                        <tfoot>
                                            <tr>
                                                <td colspan="2">Fuente: SUNAT</td>
                                            </tr>
                                        </tfoot>
                                        <tbody>
                                            <tr>
                                                <td><?= $precio_compra ?></td>
                                                <td><?= $precio_venta ?></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
    
                        <!--Panel: Precio del Oro-->
                        <div class="panel panel-default">
                            <div class="panel-heading  panel-heading-custom">
                                <h3 class="panel-title">
                                    Calculadora
                                </h3>
                            </div>
                            <div class="panel-body">

                                <div class="form-group">
                                    <label for="ruc" class="col-sm-4 control-label">RUC</label>
                                    <div class="col-sm-8">
                                        <div class="input-group search-ruc">
                                            <input type="number" class="form-control input-sm" id="ruc" name="ruc" placeholder="RUC (11 dígitos)">
                                            <button type="button" onclick="getRuc()" class="btn btn-success">BUSCAR</button>
                                        </div>
                                    </div>
                                </div>

                                <div class="form-group">
                                    <label for="razon_social" class="col-sm-4 control-label">Razón Social</label>
                                    <div class="col-sm-8">
                                        <input type="text" class="form-control input-sm" id="razon_social" name="razon_social" placeholder="Razón Social">
                                    </div>
                                </div>

                                <div class="form-group">
                                    <label for="numero_liq" class="col-sm-4 control-label">N° Liquidación</label>
                                    <div class="col-sm-8">
                                        <input type="text" class="form-control input-sm" id="numero_liq" name="numero_liq" placeholder="Número de Liq" readonly value="001-<?= $liq ?>">
                                    </div>
                                </div>

                                <hr class="bd-primary">

                                <div class="form-group">
                                    <label for="precio_inter" class="col-sm-4 control-label">Precio Internacional</label>
                                    <div class="col-sm-8">
                                        <input type="text" class="form-control input-sm" id="precio_inter" name="precio_inter" aria-label="help_precio_inter"  onkeyup="getPriceGold()">
                                        <span id="help_precio_inter" class="help-block">Dólares/Onza</span>
                                    </div>
                                </div>

                                <div class="form-group">
                                    <label for="tipo_cambio" class="col-sm-4 control-label">Tipo de Cambio</label>
                                    <div class="col-sm-8">
                                        <input type="text" class="form-control input-sm" id="tipo_cambio" name="tipo_cambio" aria-label="help_tipo_cambio" onkeyup="getPriceGold()">
                                        <span id="help_tipo_cambio" class="help-block">Soles/Dólares</span>
                                    </div>
                                </div>

                                <div class="form-group">
                                    <label for="peso_oro" class="col-sm-4 control-label">Peso de Oro</label>
                                    <div class="col-sm-8">
                                        <input type="text" class="form-control input-sm" id="peso_oro" name="peso_oro" aria-label="help_peso_oro" onkeyup="getPriceGold()">
                                        <span id="help_peso_oro" class="help-block">Gramos</span>
                                    </div>
                                </div>

                                <div class="form-group">
                                    <label for="pureza_oro" class="col-sm-4 control-label">Pureza de Oro</label>
                                    <div class="col-sm-8">
                                        <input type="text" class="form-control input-sm" id="pureza_oro" name="pureza_oro" aria-label="help_pureza_oro" onkeyup="getPriceGold()">
                                        <span id="help_pureza_oro" class="help-block">Milésimos</span>
                                    </div>
                                </div>
                                
                                <div class="form-group">
                                    <label for="descuento" class="col-sm-4 control-label">Descuentos</label>
                                    <div class="col-sm-8">
                                        <div class="input-group">
                                            <input type="text" class="form-control input-sm" id="descuento" name="descuento" onkeyup="getPriceGold()">
                                            <div class="input-group-addon">%</div>
                                        </div>
                                    </div>
                                </div>

                                <div class="form-group">
                                    <label for="detraccion" class="col-sm-4 control-label">SPOT Detracción</label>
                                    <div class="col-sm-8">
                                        <div class="input-group">
                                            <input type="text" class="form-control input-sm" id="detraccion" name="detraccion" onkeyup="getPriceGold()">
                                            <div class="input-group-addon">%</div>
                                        </div>
                                    </div>
                                </div>

                                <div class="form-group">
                                    <label for="precio_oro_peru" class="col-sm-4 control-label">Precio del Oro en el Perú</label>
                                    <div class="col-sm-8">
                                        <input type="text" class="form-control input-sm" id="precio_oro_peru" name="precio_oro_peru" aria-label="help_precio_oro_peru" onclick="getPriceGold()" readonly>
                                        <span id="help_precio_oro_peru" class="help-block">Soles</span>
                                    </div>
                                </div>

                                <input type="hidden" name="sunat_precio_compra" value="<?= $precio_compra ?>">
                                <input type="hidden" name="sunat_precio_venta" value="<?= $precio_venta ?>">
                                <input type="hidden" name="kitco_am" value="<?= $am ?>">
                                <input type="hidden" name="kitco_pm" value="<?= $pm ?>">

                                <div class="form-group">
                                    <div class="col-sm-4"></div>
                                    <div class="col-sm-8 text-right">
                                        <button class="btn btn-default" type="reset">LIMPIAR</button>
                                        <button class="btn btn-primary" type="submit">IMPRIMIR PDF</button>
                                    </div>
                                </div>

                            </div>
                        </div>

                    </div>
                </form>
            </div>
        </div> 
    </div>

    <script>

        let rucInput = document.getElementById('ruc')
        let razonInput = document.getElementById('razon_social')
        function getRuc() {
            // Create an FormData object
            var data = new FormData();
            // If you want to add an extra field for the FormData
            data.append("action", "getnumero");
            data.append("numero", rucInput.value);

            // disabled the submit button
            //$("#btnSubmit").prop("disabled", true);

            $.ajax({
                type: "POST",
                enctype: 'multipart/form-data',
                url: "https://incared.com/api/apirest",
                data: data,
                processData: false,
                contentType: false,
                cache: false,
                timeout: 600000,
                success: function (data) {
                    console.log("SUCCESS : ", data);
                    data = JSON.parse(data)
                    razonInput.value = data.rs
                },
                error: function (e) {
                    console.log("ERROR : ", e);
                }
            });
        }

        let precio_inter = document.getElementById('precio_inter')
        let tipo_cambio = document.getElementById('tipo_cambio')
        let peso_oro = document.getElementById('peso_oro')
        let pureza_oro = document.getElementById('pureza_oro')
        let descuento = document.getElementById('descuento')
        let detraccion = document.getElementById('detraccion')

        let precio_oro_peru = document.getElementById('precio_oro_peru')

        function getPriceGold() {
            let precio =
                (precio_inter.value / 31.1035) *
                tipo_cambio.value *
                peso_oro.value *
                (pureza_oro.value / 100) *
                ((100 - descuento.value - detraccion.value) / 100)

            precio_oro_peru.value = Math.round((precio + Number.EPSILON) * 100) / 100
            
        }


    </script>
</body>
</html>