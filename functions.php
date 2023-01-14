<?php

include 'SimpleXLSXGen.php';
include 'SimpleXLSX.php';

use Shuchkin\SimpleXLSXGen;
use Shuchkin\SimpleXLSX;

function get_products($instock){
  global $wpdb;

  $products = $wpdb->get_results( "SELECT ID, post_type, post_title, post_parent FROM wp_posts WHERE post_type = 'product' OR post_type = 'product_variation' ORDER BY `wp_posts`.`ID`  DESC");
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

  $price = $wpdb->get_results( "SELECT post_id, meta_key, meta_value FROM wp_postmeta WHERE meta_key = '_regular_price'  ORDER BY `wp_postmeta`.`post_id`  DESC");
  $sale = $wpdb->get_results( "SELECT post_id, meta_key, meta_value FROM wp_postmeta WHERE meta_key = '_sale_price'  ORDER BY `wp_postmeta`.`post_id`  DESC");
  $stock = $wpdb->get_results( "SELECT post_id, meta_key, meta_value FROM wp_postmeta WHERE meta_key = '_stock_status'  ORDER BY `wp_postmeta`.`post_id`  DESC");

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
  global $wpdb;

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
			// We will redirect the user to the attachment page after uploading the file successfully.
			wp_redirect( menu_page_url('wordpress-excel-plugin', true) );

      if ( $xlsx = SimpleXLSX::parseData(file_get_contents($attachment_id['url'])) ) {
          foreach($xlsx->rows() as $key=>$row){
            if($key==0) continue;
            if($row[2]) {
              if(!$wpdb->update('wp_postmeta', array('meta_value'=>$row[2]), array('meta_key'=>'_regular_price', 'post_id'=>$row[0]))){
                $wpdb->insert('wp_postmeta', array('post_id'=>$row[0], 'meta_key'=>'_regular_price', 'meta_value'=>$row[2]));
              }
            }else{
              $wpdb->update('wp_postmeta', array('meta_value'=>'outofstock'), array('meta_key'=>'_stock_status', 'post_id'=>$row[0]));
            }
            if($row[3]){
              if(!$wpdb->update('wp_postmeta', array('meta_value'=>$row[3]), array('meta_key'=>'_sale_price', 'post_id'=>$row[0]))){
                $wpdb->insert('wp_postmeta', array('post_id'=>$row[0], 'meta_key'=>'_sale_price', 'meta_value'=>$row[3]));
              }
            }else{
              $wpdb->delete('wp_postmeta', array('meta_key'=>'_sale_price', 'post_id'=>$row[0]));
            }

          }
      } else {
        wp_die( esc_html__( SimpleXLSX::parseError(), 'theme-text-domain' ) );
      }
		}
  }

  if(isset( $_POST['wxp_url_update'])){
    if(!$wpdb->update('wp_options', array('option_value'=>$_POST['wxp_url']), array('option_name'=>'wxp_url'))){
      $wpdb->insert('wp_options', array('option_name'=>'wxp_url', 'option_value'=>$_POST['wxp_url']));
    }

    if ( $xlsx = SimpleXLSX::parseData(file_get_contents($_POST['wxp_url'])) ) {
      foreach($xlsx->rows() as $key=>$row){
        if(!$row[4]) continue;
        if($row[2]) {
          if(!$wpdb->update('wp_postmeta', array('meta_value'=>$row[2]), array('meta_key'=>'_regular_price', 'post_id'=>$row[4]))){
            $wpdb->insert('wp_postmeta', array('post_id'=>$row[4], 'meta_key'=>'_regular_price', 'meta_value'=>$row[2]));
          }
        }else{
          $wpdb->update('wp_postmeta', array('meta_value'=>'outofstock'), array('meta_key'=>'_stock_status', 'post_id'=>$row[4]));
        }
        if($row[3]){
          if(!$wpdb->update('wp_postmeta', array('meta_value'=>$row[3]), array('meta_key'=>'_sale_price', 'post_id'=>$row[4]))){
            $wpdb->insert('wp_postmeta', array('post_id'=>$row[4], 'meta_key'=>'_sale_price', 'meta_value'=>$row[3]));
          }
        }else{
          $wpdb->delete('wp_postmeta', array('meta_key'=>'_sale_price', 'post_id'=>$row[4]));
        }
      }
    } else {
      wp_die( esc_html__( SimpleXLSX::parseError(), 'theme-text-domain' ) );
    }
  }
}

add_action( 'init', 'wxp_submit' );

function wxp_get_url(){
  global $wpdb;
  $res = $wpdb->get_results( "SELECT option_value FROM wp_options WHERE option_name = 'wxp_url'");
  return $res[0]->option_value;
}
