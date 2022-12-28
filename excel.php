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
  <input type="submit" name="export" class="button" value="export" />
  </form>
  <br>
  <h2>Import</h2>
  <?php 
}

if( array_key_exists( 'export', $_POST ) ){
  $products = get_products();
  $xlsx = Shuchkin\SimpleXLSXGen::fromArray( $products );
  $xlsx->downloadAs('export.xlsx');
}