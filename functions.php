<?php

include 'SimpleXLSXGen.php';
include 'SimpleXLSX.php';

use Shuchkin\SimpleXLSXGen;
use Shuchkin\SimpleXLSX;

function get_products($instock){
  global $wpdb;
  $prefix = $wpdb->prefix;

  $products = $wpdb->get_results( "SELECT ID, post_type, post_title, post_parent FROM " . $prefix . "posts WHERE post_type = 'product' OR post_type = 'product_variation' ORDER BY `" . $prefix . "posts`.`ID`  DESC");
  $repeat = array();
  foreach($products as $product){
    if($product->post_parent){
      array_push($repeat, $product->post_parent);
    }
  }
  $repeat = array_unique($repeat);
  
  foreach($products as $key=>$value){
      if(in_array($value->ID, $repeat)){
          unset($products[$key]);
      }
  }

  $price = $wpdb->get_results( "SELECT post_id, meta_key, meta_value FROM " . $prefix . "postmeta WHERE meta_key = '_regular_price'  ORDER BY `" . $prefix . "postmeta`.`post_id`  DESC");
  $sale = $wpdb->get_results( "SELECT post_id, meta_key, meta_value FROM " . $prefix . "postmeta WHERE meta_key = '_sale_price'  ORDER BY `" . $prefix . "postmeta`.`post_id`  DESC");
  $stock = $wpdb->get_results( "SELECT post_id, meta_key, meta_value FROM " . $prefix . "postmeta WHERE meta_key = '_stock_status'  ORDER BY `" . $prefix . "postmeta`.`post_id`  DESC");

  foreach($price as $key=>$value){
    if(in_array($value->post_id, $repeat)){
        unset($price[$key]);
    }
  }
  foreach($sale as $key=>$value){
    if(in_array($value->post_id, $repeat)){
        unset($sale[$key]);
    }
  }
  foreach($stock as $key=>$value){
    if(in_array($value->post_id, $repeat)){
        unset($stock[$key]);
    }
  }

  foreach($products as $product){
    foreach($price as $value){
      if($product->ID == $value->post_id){
        $product->price = $value->meta_value;
      }
    }
    foreach($sale as $value){
      if($product->ID == $value->post_id){
        $product->sale = $value->meta_value;
      }
    }
    foreach($stock as $value){
      if($product->ID == $value->post_id){
        $product->stock = $value->meta_value;
      }
    }
  }

  $excel = array(['<center><b>ID</b></center>', '<center><b>Post Title</b></center>', '<center><b>Price</b></center>', '<center><b>Sale Price</b></center>']);
  foreach($products as $product){
    if($instock and $product->stock=='outofstock') continue;
    array_push($excel, [$product->ID, $product->post_title, $product->price, $product->sale]);
  }
  
  return $excel;
}

function wxp_submit(){
  if( isset( $_POST['wxp_export']) and isset( $_POST['wxp_instock']) ){
    $products = get_products(true);
    $xlsx = SimpleXLSXGen::fromArray( $products );
    $xlsx->downloadAs('export.xlsx');
  } elseif (isset( $_POST['wxp_export'])){
    $products = get_products(false);
    $xlsx = SimpleXLSXGen::fromArray( $products );
    $xlsx->downloadAs('export.xlsx');
  }

  if(isset( $_POST['wxp_update'])){
    if ( ! wp_verify_nonce( $_POST['wxp_nonce'], 'upload_wxp_file' ) ) {
			wp_die( esc_html__( 'Nonce mismatched', 'theme-text-domain' ) );
		}
    if ( ! $_FILES['wxp_file']['name'] ) {
			wp_die( esc_html__( 'Please choose a file', 'theme-text-domain' ) );
		}
		$allowed_extensions = array( 'xlsx' );
		$file_type = wp_check_filetype( $_FILES['wxp_file']['name'] );
		$file_extension = $file_type['ext'];

		// Check for valid file extension
		if ( ! in_array( $file_extension, $allowed_extensions ) ) {
			wp_die( sprintf(  esc_html__( 'Invalid file extension, only allowed: %s', 'theme-text-domain' ), implode( ', ', $allowed_extensions ) ) );
		}

    $uploadedfile = $_FILES['wxp_file'];
    $attachment_id = wp_handle_upload( $uploadedfile, array( 'test_form' => false ) );

		if ( is_wp_error( $attachment_id ) ) {
			// There was an error uploading the image.
			wp_die( $attachment_id->get_error_message() );
		} else {
      wxp_update_products($attachment_id['url'], 0, 2, 3);

      // We will redirect the user to the attachment page after uploading the file successfully.
			wp_redirect( menu_page_url('wordpress-excel-plugin', true) );
		}
  }

  if(isset( $_POST['wxp_url_update'])){
    global $wpdb;
    $prefix = $wpdb->prefix;

    if(!$wpdb->update( $prefix .'options', array('option_value'=>$_POST['wxp_url']), array('option_name'=>'wxp_url'))){
      $wpdb->insert( $prefix .'options', array('option_name'=>'wxp_url', 'option_value'=>$_POST['wxp_url']));
    }

    wxp_update_products($_POST['wxp_url'], 4, 2, 3);
  }
}

add_action( 'init', 'wxp_submit' );

function wxp_get_url(){
  global $wpdb;
  $prefix = $wpdb->prefix;
  $res = $wpdb->get_results( "SELECT option_value FROM " . $prefix . "options WHERE option_name = 'wxp_url'");
  return $res[0]->option_value;
}

function wxp_update_products($file_path, $id_col, $price_col, $sale_col){
  global $wpdb;
  $prefix = $wpdb->prefix;

  $curl = curl_init($file_path);
  curl_setopt($curl, CURLOPT_RETURNTRANSFER, true);
  curl_setopt($curl, CURLOPT_FOLLOWLOCATION, true);
  $res = curl_exec($curl);
  curl_close();

  if ( $xlsx = SimpleXLSX::parseData($res) ) {
    foreach($xlsx->rows() as $key=>$row){
      if(empty($row[$id_col])) continue;
      if(!empty($row[$price_col])) {
        $wpdb->update( $prefix .'postmeta', array('meta_value'=>'instock'), array('meta_key'=>'_stock_status', 'post_id'=>$row[$id_col]));
        $reg_price = $wpdb->get_results( "SELECT * FROM " . $prefix . "postmeta WHERE meta_key = '_regular_price' AND post_id = ".$row[$id_col]);
        if($reg_price){
          $wpdb->update( $prefix .'postmeta', array('meta_value'=>$row[$price_col]), array('meta_key'=>'_regular_price', 'post_id'=>$row[$id_col]));
        }else{
          $wpdb->insert( $prefix .'postmeta', array('post_id'=>$row[$id_col], 'meta_key'=>'_regular_price', 'meta_value'=>$row[$price_col]));
        }

        if(!empty($row[$sale_col])){
          $sale_price = $wpdb->get_results( "SELECT * FROM " . $prefix . "postmeta WHERE meta_key = '_sale_price' AND post_id = ".$row[$id_col]);
          if($sale_price){
            $wpdb->update( $prefix .'postmeta', array('meta_value'=>$row[$sale_col]), array('meta_key'=>'_sale_price', 'post_id'=>$row[$id_col]));
          }else{
            $wpdb->insert( $prefix .'postmeta', array('post_id'=>$row[$id_col], 'meta_key'=>'_sale_price', 'meta_value'=>$row[$sale_col]));
          }
        }else{
          $wpdb->delete( $prefix .'postmeta', array('meta_key'=>'_sale_price', 'post_id'=>$row[$id_col]));
        }

        if(!empty($row[$sale_col])){
          $wpdb->update( $prefix .'postmeta', array('meta_value'=>$row[$sale_col]), array('meta_key'=>'_price', 'post_id'=>$row[$id_col]));
        }elseif(!empty($row[$price_col])){
          $wpdb->update( $prefix .'postmeta', array('meta_value'=>$row[$price_col]), array('meta_key'=>'_price', 'post_id'=>$row[$id_col]));
        }
      }else{
        $wpdb->update( $prefix .'postmeta', array('meta_value'=>'outofstock'), array('meta_key'=>'_stock_status', 'post_id'=>$row[$id_col]));
      }
    }

    $products = $wpdb->get_results( "SELECT ID FROM " . $prefix . "posts WHERE post_type = 'product' AND post_parent=0 ORDER BY `" . $prefix . "posts`.`ID`  DESC");
    foreach($products as $key=>$product){
      $variations = $wpdb->get_results( "SELECT ID FROM " . $prefix . "posts WHERE post_type = 'product_variation' AND post_parent=" . $product->ID . " ORDER BY `" . $prefix . "posts`.`ID`  DESC");
      if(!$variations) continue;
      $instock = false;
      foreach($variations as $key=>$variation){
        $status = $wpdb->get_results( "SELECT meta_value FROM " . $prefix . "postmeta WHERE meta_key = '_stock_status' AND post_id=" . $product->ID);
        if($status->meta_value=='instock') $instock = true;
      }
      if($instock){
        $wpdb->update( $prefix .'postmeta', array('meta_value'=>'instock'), array('meta_key'=>'_stock_status', 'post_id'=>$product->ID));
      }else{
        $wpdb->update( $prefix .'postmeta', array('meta_value'=>'outofstock'), array('meta_key'=>'_stock_status', 'post_id'=>$product->ID));
      }
    }
  } else {
    wp_die( esc_html__( SimpleXLSX::parseError(), 'theme-text-domain' ) );
  } 
}
