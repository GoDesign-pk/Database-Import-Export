<?php 

/*
  Plugin Name: Import/Export Mysql Table
  Plugin URI: http://www.stillbreathing.co.uk/wordpress/database-browser/
  Description: Easily browse the data in your database, and download in CSV, XML, SQL and JSON format
  Author: Agile solution
  Version: 1.4.4
 */

function myplugin_register_options_page() {
  add_options_page('Page Title', 'Import Export MySql Table', 'manage_options', 'myplugin', 'myplugin_options_page');
}
add_action('admin_menu', 'myplugin_register_options_page');

function myplugin_options_page()
{
?>
  <div>
  <?php 
        screen_icon(); 
        global $wpdb;
		$sql = "SHOW TABLES;";
		$results = $wpdb->get_results( $sql, ARRAY_N );
		$tables = array();
		$i=0;
		foreach ( $results as $result ) {
		    $i++;
		    if($i>12){
			    $tables[] = $result[0];
		    }
		}
  ?>
  <h2>Import/Export Database Multiple Tables From Single Excel File With Multiplesheets</h2>
  <form method="post" name="import_export_form">
  <?php settings_fields( 'myplugin_options_group' ); ?>
  <table>
  <tr valign="top">
  <td><input type="submit" id="export_btn" name="exportbtn" style="cursor:pointer;background: #03a9f4;padding: 8px 10px;border-radius: 5px;border: 1px solid #03a9f4;color: white;" value="Export Tables" /></td>
  <th scope="row"><label for="myplugin_option_name"><?= implode(", ",$tables); ?></label></th>
  </tr>
  </table>
  </form>
  <br><br>
    <form  method="post" enctype="multipart/form-data">
        Upload excel file : 
        <input type="file" name="uploadFile" value="" />
        <input type="submit" id="import_btn" class="btn btn-info" style="cursor:pointer;background: #ffc107;padding: 8px 10px;border-radius: 5px;border: 1px solid #ffc107;color: white;" name="importbtn" value="Import Now"/> 
    </form>
  </div>
<?php
} 

if(isset($_POST['exportbtn'])){
    export();
}

if(isset($_POST['importbtn'])){
    import();
}

function checkExist($id,$tblname)
{
    global $wpdb;
    $rowcount = $wpdb->get_var("SELECT COUNT(*) FROM $tblname WHERE id = '$id' ");
    return $rowcount;
}
    
function import(){
    global $wpdb;
    if(isset($_FILES['uploadFile']['name']) && $_FILES['uploadFile']['name'] != "") {
        $allowedExtensions = array("xls","xlsx");
        $ext = pathinfo($_FILES['uploadFile']['name'], PATHINFO_EXTENSION);
        if(in_array($ext, $allowedExtensions)) {
            
               $file = plugin_dir_path( __FILE__ ) ."uploads/".$_FILES['uploadFile']['name'];
               $isUploaded = move_uploaded_file($_FILES['uploadFile']['tmp_name'], $file);
               if($isUploaded) {
                    include("Classes/PHPExcel/IOFactory.php");
                    try {
                        //Load the excel(.xls/.xlsx) file
                        $objPHPExcel = PHPExcel_IOFactory::load($file);
                    } catch (Exception $e) {
                         die('Error loading file "' . pathinfo($file, PATHINFO_BASENAME). '": ' . $e->getMessage());
                    }
                    
                    $sheets = $objPHPExcel->getSheetNames();
                    $sheet_index=0;
                    foreach($sheets as $sheetname){
                        $pos = strpos($sheetname, "-");
                            $arr = explode("-", $sheetname, 2);
                            $tblname = $arr[0];
                           
                            $total_num_cols = substr($sheetname, $pos+1);
                            //An excel file may contains many sheets, so you have to specify which one you need to read or work with.
                            $sheet = $objPHPExcel->getSheet($sheet_index);
                            //It returns the highest number of rows
                            $total_rows = $sheet->getHighestRow();
                            //It returns the highest number of columns
                            $total_columns = $sheet->getHighestColumn();
                            $update_query = "Update `$tblname` set ";
                            $insert_query = "insert into `$tblname` VALUES (";
                            for($row =2; $row <= $total_rows; $row++) {
                                //Read a single row of data and store it as a array.
                                //This line of code selects range of the cells like A1:D1
                                $query="";
                                $headings = $sheet->rangeToArray('A1:' . $total_columns . 1,
                                                    NULL,
                                                    TRUE,
                                                    FALSE);
                                $single_row = $sheet->rangeToArray('A' . $row . ':' . $total_columns . $row, NULL, TRUE, FALSE);
                                $single_row[0] = array_combine($headings[0], $single_row[0]);
                                $i=0;
                                $row_id=0;
                                //Creating a dynamic query based on the rows from the excel file
                                //Print each cell of the current row
                                
                                foreach($single_row[0] as $key=>$value) {
                                    if($key=='id'){
                                        $row_id = $value;
                                    }
                                    
                                    if(checkExist($row_id,$tblname)>0){
                                        
                                        if($i<$total_num_cols){
                                            $query .= $key."='".$value."',";
                                        }
                                        
                                        $complete_query = $update_query . $query;
                                    }else{
                                        if($i<$total_num_cols){
                                          $query .= "'".$value."',";
                                        }
                                        $complete_query = $insert_query . $query;
                                    }
                                    $i++;
                                }
                                $complete_query = substr($complete_query, 0, -1);
                                if(checkExist($row_id,$tblname)>0){
                                    $complete_query .= "";
                                }else{
                                   $complete_query .= ")";
                                }
                                //echo $complete_query;
                                $mysqli_affected_rows = $wpdb->query($complete_query);         
                                        
                                if($mysqli_affected_rows > 0) {
                                    $result = '<span class="msg">Database table updated!</span>';
                                }else{
                                    $result = '<span class="msg">Something Went Wrong!</span>';
                                }
                            }
                            
                            
                            
                        
                        $sheet_index++;
                    }
                    
                    echo '<span class="msg">Database table updated!</span>';
                    // Finally we will remove the file from the uploads folder (optional) 
                    unlink($file);
                } else {
                    echo '<span class="msg">File not uploaded!</span>';
                }
            
        } else {
            echo '<span class="msg">This type of file not allowed!</span>';
        }
    } else {
        echo '<span class="msg">Select an excel file first!</span>';
    }
}


function export(){
    require_once 'database.php';
    require_once 'Classes/PHPExcel.php';
    require_once 'Classes/PHPExcel/IOFactory.php';
    global $wpdb;
    $sql = "SHOW TABLES;";
	$results = $wpdb->get_results( $sql, ARRAY_N );
	$tables = array();
	$i=0;
	foreach ( $results as $result ) {
	    $i++;
	    if($i>12){
		    $tables[] = $result[0];
	    }
	}
	$active_index = 0;
	
	foreach ( $tables as $table ) {
        $tbl_result = $wpdb->get_results ( "SELECT * FROM {$table}" );
        if($tbl_result){
    
            if($active_index==0){
                /* Create new PHPExcel object*/
                $objPHPExcel = new PHPExcel();
                $objPHPExcel->setActiveSheetIndex($active_index);
                
                $columns = $wpdb->get_results("SHOW COLUMNS FROM {$table}");
                $column_count = count($columns);
                $colindex =0;
                foreach($columns as $v) {
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($colindex, 1, $v->Field);
                    $colindex++;
                }
                $objPHPExcel->getActiveSheet()->getStyle('A1:IV1')->getFont()->setBold(true)->setName('Verdana')->setSize(10);
	            
                $row = 2;
                foreach ($tbl_result as $res) {
                    $col = 0;
                    foreach($res as $v){
                    	$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $v);
                    	$col++;
                    }
                    $row++;
                }
                
                $objPHPExcel->getActiveSheet()->setTitle($table."-".$column_count);
            }else{
                $objPHPExcel->createSheet();
                $objPHPExcel->setActiveSheetIndex($active_index);
                
                $columns = $wpdb->get_results("SHOW COLUMNS FROM {$table}");
                $column_count = count($columns);
                $colindex = 0;
                foreach($columns as $v) {
                    $objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($colindex, 1, $v->Field);
                    $colindex++;
                }
                $objPHPExcel->getActiveSheet()->getStyle('A1:IV1')->getFont()->setBold(true)->setName('Verdana')->setSize(10);
                
                $row = 2;
                $row = 2;
                foreach ($tbl_result as $res) {
                    $col = 0;
                    foreach($res as $v){
                    	$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $v);
                    	$col++;
                    }
                    $row++;
                }
                $objPHPExcel->getActiveSheet()->setTitle($table."-".$column_count);
            }
            $active_index++;
        }
	}
    
    
    
    
    
    // /* Create a new worksheet, after the default sheet*/
    // $objPHPExcel->createSheet();
    
    // /* Add some data to the second sheet, resembling some different data types*/
    // $objPHPExcel->setActiveSheetIndex(1);
    // $objPHPExcel->getActiveSheet()->setCellValue('A1', 'Salary');
    // $i=2;
    // while($row1= mysqli_fetch_array($result1)) {
    // 	$salary=$row1['title'];
    // 	$objPHPExcel->getActiveSheet()->setCellValue("A$i",$salary);
    // $i++;
    // }
    
    // /* Rename 2nd sheet*/
    // $objPHPExcel->getActiveSheet()->setTitle('WP Data Tables');
    
    /* Redirect output to a client’s web browser (Excel5)*/
    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment;filename="wpdatatable.xls"');
    header('Cache-Control: max-age=0');
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
    $objWriter->save('php://output');

}

?>