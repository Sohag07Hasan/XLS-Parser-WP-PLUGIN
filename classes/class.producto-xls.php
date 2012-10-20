<?php
set_time_limit(0);

/*
 * create admin interfaces to import the xls files
 * */

class Producto_xls_parser{
	
	
	/*
	 * some constants for custom post types
	 * */
	const post_type = "producto";
	const post_category = "producto_category";
	const post_tag = "producto_tag";
	const meta_key = "price";
	
	//initialize the hooks
	static function init(){
		add_action('admin_menu', array(get_class(), 'add_submenu_page'));
		add_action('init', array(get_class(), 'xls_parser'));
	}
	
	
	/*
	 * some variables
	 **/
	static $headers = array();
	static $descriptions = array();
	static $current_row = array();
	static $current_post = array();
	static $posted = 0;
	static $skipped = 0;
	static $exc_time = 0;
	static $is_xml = false;
	
	/*
	 * creates admin interface
	 * */
	static function add_submenu_page(){
		//add_options_page('parsing xls sheet', 'Import XLS', 'manage_options', 'import-xls', array(get_class(), 'options_page_content'));
		add_submenu_page('edit.php?post_type='.self::post_type, 'parsing xls sheet', 'Import XLS', 'manage_options', 'import-xls', array(get_class(), 'options_page_content'));
	}
	
	
	/*
	 * options page
	 * */
	static function options_page_content(){
		include PRODUCTOXLS_DIR . '/includes/import-form.php';
	}
	
	
	/*
	 * xls_parser
	 * */
	static function xls_parser(){
		if($_POST['producto-xls-importer-form'] == "Y") :			
			if(!empty($_FILES['producto-xls-file']['tmp_name'])){

				$start = microtime(true);
				
				$xls = self::get_xls_object($_FILES['producto-xls-file']['tmp_name']);
				
				if($xls){
					self::$is_xml = true;
				}
				
				$sheet = 0;
				for($row=1;$row<=$xls->rowcount($sheet);$row++){
					self::$current_row = array();
					
					for($col=1;$col<=$xls->colcount($sheet);$col++) {
						$rowspan = $xls->rowspan($row,$col,$sheet);
						$colspan = $xls->colspan($row,$col,$sheet);
						for($i=0;$i<$rowspan;$i++) {
							for($j=0;$j<$colspan;$j++) {
								if ($i>0 || $j>0) {
									$xls->sheets[$sheet]['cellsInfo'][$row+$i][$col+$j]['dontprint']=1;
								}
							}
						}
						
						if(!$xls->sheets[$sheet]['cellsInfo'][$row][$col]['dontprint']) {
							$val = $xls->val($row,$col,$sheet);
							if($row == 1){
								self::$headers[] = explode("|", $val);
							//	self::$headers[] = $val;
							}
							if($row == 2){
								self::$descriptions[] = $val;
							}
							
							if($row >= 3){
								self::$current_row[] = $val;
							}
						}
					}

					self::create_post();
					
				}			
				
				$end = microtime(true);
				self::$exc_time = $end - $start;
				
			}
		endif;
	}
	
	
	/*
	 * return the xls object
	 * */
	static function get_xls_object($file){
		if(!class_exists('Spreadsheet_Excel_Reader')){
			include PRODUCTOXLS_DIR . '/classes/class.parser.php';
		}
		
		return new Spreadsheet_Excel_Reader($file, false);
	}
	
	
	/*
	 * create posts
	 * */
	static function create_post(){
		
		
		if(!empty(self::$current_row)){
			foreach(self::$headers as $key => $header){						
				$post = array();
				if($key >= 2 && !empty($header[0])){					
					$post['post_title'] = self::$current_row[0] . " " . ucwords(strtolower(trim($header[0])));
					$post['post_content'] = self::$descriptions[$key];
					$post['tag'] = self::$current_row[1];
					$post['category'] = self::create_categories($header, self::$current_row[0]);
					$post['price'] = self::sanitize_price(self::$current_row[$key]);
					self::insert_post($post);					
				}
				
			}			
						
		}		
		
	}
	
	
	/**
	 * creates the categoires
	 * */
	static function create_categories($header, $title){
		$categories = array();
		foreach($header as $h){
			$categories[] = ucwords(strtolower(trim($h)));
		}
		
		$categories[] = ucwords(strtolower(trim($title)));
		
		return $categories;
	}
	
	/*
	 * sanitize Price
	 * */
	static function sanitize_price($price){
		$price = preg_replace("#[^0-9.]#", '', $price);
		if($price){
			return number_format((float)$price, 2, '.', '');
		}
		else{
			return false;
		}
	}
	
	
	/*
	 * inserting the post into database
	 * */
	static function insert_post($post){
		if($post['price']){
			$sanitized_post = array(
				'post_title' => $post['post_title'],
				'post_content' => $post['post_content'],
				'post_type' => self::post_type,
				'post_status' => 'publish'
			);
			$post_id = wp_insert_post($sanitized_post);
			if($post_id){
				update_post_meta($post_id, self::meta_key, $post['price']);
				wp_set_object_terms($post_id, $post['category'], self::post_category);
				if(!empty($post['tag'])){
					wp_set_object_terms($post_id, $post['tag'], self::post_tag);
				}
				self::$posted ++;
			}
			else{
				self::$skipped ++;
			}
		}
		else{
			self::$skipped ++;
		}

		return;
	}
	
	
	/*
	 * get demo xls url
	 * */
	static function get_demo_xls(){
		return PRODUCTOXLS_URL . "sample/demo.xls";
	}
	
	
	/*
	 * print the message
	 * */
	static function print_message(){
		$message = "";
		if($is_xml) :
			$message .= "<div class='updated'>";
			$message .= "<p>XLS is successfully parsed</p>";
			$message .= "<p> Created Posts " . self::$posted . "</p>";
			$message .= "<p> Skipped " . self::$skipped . " </p>";
			$message .= "<p> Time Required " . self::$exc_time . "</p>";
			$message .= "</div>";
		endif;
		return $message;	
	}
	
}