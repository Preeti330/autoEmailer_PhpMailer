<?php
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

//load the composers autoload.php
require 'vendor/autoload.php';
include "db.php";
$db_connect = $conn;
$date=date('Y-m-d');

$sql="
			SELECT 
					d.id,d.skus,total_voucher_code,used_vouchers AS total_sent_vouchers,(total_voucher_code-used_vouchers) AS voucher_codes_left
				--	,case when total_vouchers_sent_today IS NOT NULL THEN 
				-- total_vouchers_sent_today
			     --   ELSE 0 END AS credited_date  
			FROM denominations AS d 
			Left JOIN 
			(
			SELECT  count(used_vouchers) AS used_vouchers,COUNT(v.id) AS total_voucher_code,v.denomination_id
			FROM voucher_codes AS v 
			left JOIN ( SELECT COUNT(id) AS used_vouchers,voucher_code_id from stashfin_orders WHERE voucher_code_id IS NOT NULL AND denomination_id IS NOT NULL GROUP BY voucher_code_id ) AS s ON (s.voucher_code_id=v.id) 
			GROUP BY v.denomination_id ) AS asss
			ON (asss.denomination_id=d.id)
			LEFT JOIN (
			SELECT s1.denomination_id,COUNT(s1.id) AS total_vouchers_sent_today,credited_date FROM stashfin_orders AS s1 WHERE s1.credited_date='".$date."' GROUP BY s1.denomination_id,credited_date ) AS ass2 ON (ass2.denomination_id=d.id)
			ORDER BY id ASC 
";

$getRedords = pg_query($GLOBALS['db_connect'], $sql);
$getAllRedords    = pg_fetch_all($getRedords);


$file='reports/'.rand().'_'.date("Y-m-d").'.xlsx';
generateExcelReport($getAllRedords,$file);

 function generateExcelReport($queryResultData, $file)
{
	$spreadsheet = new Spreadsheet();   
	$activesheet = $spreadsheet->getActiveSheet();
	//create file directory
	// $dir = 'uploads/';
	// if (!file_exists($dir)) 
	// {
	// 	mkdir($dir, 777, true);
	// }
	// $date=date('Y-m-d');
	// $dirpath = '/reports/'.$SpreadSheetTitle.'/';
	// if (!file_exists($dirpath)) {
	// 	mkdir($dirpath, 0777, true);
	// }
	// $file_name = uniqid() . '.Xls';
	// $file_name = $dirpath .$date.'-'.$file_name;
	//set title of spreadsheet
	//$activesheet->setTitle($SpreadSheetTitle);
	//set row as i=2;
	$i = 2;
	// print_r($queryResultData);exit;
	foreach ($queryResultData as $key => $val) {
		//$count = count($queryResultData[$key]);
		$j = 0;
		foreach ($val as $keys => $value) {
			//assci val of 65 is A and Every Time increase it by one to get next column (like B,C.......)               
			$charval = 65 + $j;
			
			 $char=chr($charval);                       
			//set headers using $key for $char.1(A1,B1....);
			if($i==2)
			{
				if($char>='A' && $char<='Z'){
					$activesheet->setCellValue($char.'1',$keys);
					//set style to headers 
					$spreadsheet->getActiveSheet()->getStyle($char.'1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()
					->setRGB('003300');
					$spreadsheet->getActiveSheet()->getStyle($char.'1')->getFont()->setBold(true)->setSize(13)->getColor()->setRGB('FFFFFF');
				}
			}
			if ($charval >= 65 && $charval <= 90) 
			{
				$char = chr($charval);
				
			   // echo "<br>";
				//  set values remaining cloumns 
				$activesheet->setCellValue($char . $i, $val[$keys]);
			   
			}
			if ($charval >= 91 && $charval<=116){
				//set values for Column AA,AB,AC,.....
				$charval2 = $charval - 26;//91-26=65(i.e A)
				$charA = 'A'.chr($charval2);
				if($i==2){
					if($charA>='AA' && $charA<='AZ'){
					   // print_r($charA);
					   // echo "<br>";
						$activesheet->setCellValue($charA.'1',$keys);
						//set style to headers 
						$spreadsheet->getActiveSheet()->getStyle($charA.'1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('green');
						$spreadsheet->getActiveSheet()->getStyle($charA.'1')->getFont()->setBold(true);
					}
				}
			  //  echo "charval $charval2 $charA $keys : $val[$keys]<br>";
				 $activesheet->setCellValue($charA.$i,$val[$keys]);
			}
			
			if($charval>=117 && $charval<=142)
			{
				 //set values for Column BA,BB,BC,BD,.....
				$charval2 = $charval - 52;
				$charB = 'B'.chr($charval2);
				if($i==2){
				if($charB>='BA' && $charB<='BZ'){
					$activesheet->setCellValue($charB.'1',$keys);
					$spreadsheet->getActiveSheet()->getStyle($charB.'1')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('green');
					$spreadsheet->getActiveSheet()->getStyle($charB.'1')->getFont()->setBold(true);
				}
			}
				 $activesheet->setCellValue($charB.$i,$val[$keys]);
			  
			}             
			//incriment j by one inside second loop to change char value 
			$j = $j + 1;
		}
		//increment i for new row
		$i++;
	}
	$writer = new Xlsx($spreadsheet);
	//$writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xls');
	$writer->save($file);
	
 //  return $file;
}


$mail =new PHPMailer(true);

try {
	$mail->SMTPDebug = 0;									 
	$mail->isSMTP();										 
	$mail->Host	 = 'smtp.gmail.com';			 
	$mail->SMTPAuth = true;							 
	$mail->Username = 'preeti@gmail.com';				 
	$mail->Password = 'p123';					 
	$mail->SMTPSecure = 'tls';							 
	$mail->Port	 = 587; 
	
	$mail->setFrom('preeti@gmail.com', 'Preeti');
	$mail->addCC('ravi@gamil.com');
	// $mail->addAddress('support@xzy.com');
	$filepath='../InventoryMail/'.$file;
	$mail->addReplyTo('noreplay@bigcity.in');
	$mail->addAttachment($filepath);
	
   
    $mail->isHTML(true);                                  
    $mail->Subject = 'Inventory Notification On '.date('Y-m-d H:i:s');
    $mail->Body    = '<!DOCTYPE html>
	<html>
	<body>
	<p>Inventory Report</p>
	</body>
	</html>
	 ';
    $mail->AltBody = 'Body in plain text for non-HTML mail clients';
    $mail->send();
    echo "Mail has been sent successfully!";
} catch (Exception $e) {
    echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
}

?>
