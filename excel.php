<?php
/**
* Plugin Name: wordpress-excel-plugin
* Plugin URI: 
* Description: import/export/update woocommerce products
* Version: 1.1.4
* Author: Matin Asadpour
* Author URI: 
**/

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
  <form method="POST">
    <label for="instock">InStock?</label>
    <input type="checkbox" name="wxp_instock" value="true" checked />
    <br/>
    <input type="submit" name="wxp_export" class="button" value="export" />
  </form>
  <br><hr>
  <h2>Import</h2>
  <form method="POST" enctype="multipart/form-data">
    <input type="file" name="wxp_file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet">
    <br/>
    <?php echo wp_nonce_field( 'upload_wxp_file', 'wxp_nonce', true, false ); ?>
    <input type="submit" name="wxp_update" class="button" value="import & update" />
  </form>
  <br><hr>
  <h2>Update with Google Doc Sheets</h2>
  <form method="POST" enctype="multipart/form-data">
    <input type="text" name="wxp_url" value="<?php echo wxp_get_url(); ?>" size="135">
    <br/>
    <input type="submit" name="wxp_url_update" class="button" value="update" />
  </form>
  <?php 
}

