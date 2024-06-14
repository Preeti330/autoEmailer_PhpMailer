<?php

$host="localhost";
$user_name="postgres";
$password="p1234";
$db="amazon_db";
$connstr="host=localhost port=5432 dbname=amazon_db user=postgres password=p1234";

 $conn= pg_connect($connstr);

if($conn){
     echo "yes";
}else{
    echo "connection error";    
} 
exit;
?>