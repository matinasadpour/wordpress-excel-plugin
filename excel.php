<?php
/**
* Plugin Name: wordpress-excel-plugin
* Plugin URI: 
* Description: import/export/update woocommerce products
* Version: 0.1
* Author: Matin Asadpour
* Author URI: 
**/

include 'SimpleXLSXGen.php';
include 'SimpleXLSX.php';
include 'functions.php';

add_action( 'admin_menu', 'plugin_menu' );
function plugin_menu(){    
  $page_title = 'Excel';
  $menu_title = 'Excel';
  $capability = 'manage_options';
  $menu_slug  = 'wordpress-excel-plugin';
  $function   = 'plugin_page';
  $icon_url   = 'dashicons-media-spreadsheet';
  $position   = 3;
  add_menu_page( $page_title, $menu_title, $capability, $menu_slug, $function, $icon_url, $position );
}

function plugin_page(){
  ?>
  <h1>WordPress Excel Plugin</h1>
  <br/>
  <h2>Export</h2>
  <form method="post">
    <label for="instock">InStock?</label>
    <input type="checkbox" name="instock" value="true" checked />
    <br/>
    <input type="submit" name="export" class="button" value="export" />
  </form>
  <br><hr>
  <h2>Import</h2>
  <form meyhod="post">
    <input type="file" name="file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
    <br/>
    <input type="submit" name="update" class="button" value="import & update" />
  </form>
  <?php 
}

if( array_key_exists( 'export', $_POST ) and array_key_exists( 'instock', $_POST ) ){
  $products = get_products(true);
  $xlsx = Shuchkin\SimpleXLSXGen::fromArray( $products );
  $xlsx->downloadAs('export.xlsx');
} elseif (array_key_exists( 'export', $_POST )){
  $products = get_products(false);
  $xlsx = Shuchkin\SimpleXLSXGen::fromArray( $products );
  $xlsx->downloadAs('export.xlsx');
}