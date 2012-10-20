<?php

/*
 * plugin name: Producto XLS file Importer
 * author: Mahibul Hasan
 * 
 * */

define("PRODUCTOXLS_DIR", dirname(__FILE__));
define("PRODUCTOXLS_FILE", __FILE__);
define("PRODUCTOXLS_URL", plugins_url('/', __FILE__));

include PRODUCTOXLS_DIR . '/classes/class.producto-xls.php';
Producto_xls_parser::init();
