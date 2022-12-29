<?php

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
  $sale = $wpdb->get_results( "SELECT post_id, meta_key, meta_value FROM wp_postmeta WHERE meta_key = '_price'  ORDER BY `wp_postmeta`.`post_id`  DESC");
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