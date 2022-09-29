<?php

// ##### НАСТРОЙКИ #####
$from_email		 = 'san@kompensator.by'; // от кого будет письмо
$recipient_email = 'san@kompensator.by'; // кому отправить письмо (можно несколько через запятую)
$reply_to_email = 'san@kompensator.by'; // спец настройка почтового сервера "кому отправлять ответ" (обычно = $from_email)

// тут словарь переводов

$dictionary = [
    
    'Симметричный'=>'Concentric',
    'С двойным эксцентриситетом'=>'Double Offset',
    'С тройным эксцентриситетом'=>'Triple Offset',
    'Дымовой'=>'Air Damper',

    'Под приварку'=>'????',
    'Межфланцевое'=>'Wafer',
    'Фланцевое'=>'???????',
    'Межфланцевое с резьбовыми проушинами'=>'Lug',
    'Межфланцевое с центрирующими проушинами'=>'??',
    'Фланцевое с укороченной строительной длиной корпуса'=>'???????????',
    'else'=>'Other',

    'ГОСТ'=>'GOST',

    'Запорный'=>'Stop',
    'Регулирующий'=>'Control',
    'Запорно-регулирующий'=>'Stop and Control',

    'Без привода (голый шток)'=>'',
    'Ручной (рукоятка с фиксатором)'=>'Lever',
    'Ручной (редуктор с маховиком)'=>'Gearbox',
    'Электропривод'=>'Electric actuator',
    'Пневмопривод'=>'Pneumatic actuator',
    'Гидропневмопривод'=>'Hydraulic actuator',

    ];

// ##### КОНЕЦ БЛОКА НАСТРОЙКИ #####



require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

$file_general = __DIR__.'/general_template.xlsx';
$file_result = __DIR__.'/result.xlsx';

$data = $_POST;

if ( copy($file_general, $file_result) == false ) {
    echo "Не удалось создать файл ..";
} else {
    //echo "ok";
}



$reader = new Xlsx();
// получаем Excel-книгу
$workbook = $reader->load($file_result);

// формируем штрих-код и вставляем
$bar_code = bar_code($data);
$sheet = $workbook->getSheet(1);
$sheet->setCellValue('B38', $bar_code);
$sheet = $workbook->getSheet(3);
$sheet->setCellValue('C19', $bar_code);
$sheet->setCellValue('E19', $data['quantity']['value']);
$sheet = $workbook->getSheet(4);
$sheet->setCellValue('C19', $bar_code);
$sheet->setCellValue('E19', $data['quantity']['value']);

// заполняем RU
$sheet = $workbook->getSheet(5); // RU
fillCells($sheet, $data);

// заполняем EN
foreach ($data as $key => &$value) 
    $value['value'] = translate($value['value']);
$sheet = $workbook->getSheet(2); // EN
fillCells($sheet, $data);



// Сохраняем результат
$file_upload = "result_upload.xlsx";
$writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($workbook);
$writer->setPreCalculateFormulas(false);
$writer->save($file_upload);


// отправляем на почту
sendMail($file_upload);



exit();

function bar_code($data) {
    
    $bar_code = "AKB BTV XXX ";
    
    switch ($data['valve_purpose']['value']) {
        case "Запорный": $val = "1"; break;
        case "Регулирующий": $val = "2"; break;
        case "Запорно-регулирующий": $val = "3"; break;
        default: $val = "?";
    }
    $bar_code .= $val . "-";

    // диаметр
    $val = sprintf("%'.04d",  $data['dn']['else']);
    $bar_code .= $val . "-";

    // давление
    $val = floatval($data['pressure']['dn']['value']); 
    if ($val > 0) $val = str_repeat("0", 4 - strlen($val)) . $val; 
        else $val = "XXXX-";
    $bar_code .= $val . "-";
    
    // Управление
    switch ($data['operation']['value']) {
        case "Без привода (голый шток)": $val = "0"; break;
        case "Ручной (рукоятка с фиксатором)": $val = "1"; break;
        case "Ручной (редуктор с маховиком)": $val = "2"; break;
        case "Электропривод": $val = "3"; break;
        case "Пневмопривод": $val = "4"; break;
        case "Гидропневмопривод": $val = "5"; break;
        default: $val = "?";
    }
    $bar_code .= $val . "-";
    
    // Присоединение
    switch ($data['connection']['value']) {
        case "Под приварку": $val = "0"; break;
        case "Межфланцевое": $val = "1"; break;
        case "Фланцевое": $val = "2"; break;
        case "Межфланцевое с резьбовыми проушинами": $val = "3"; break;
        case "Межфланцевое с центрирующими проушинами": $val = "4"; break;
        case "Фланцевое с укороченной строительной длиной корпуса": $val = "5"; break;
        case "else": $val = "6"; break;
        default: $val = "?";
    }
    $bar_code .= $val . "-";
    
    // Материал корпуса
    switch ($data['body']['value']) {
        case "GGG40": 
            $val = "ВЧ"; break;
        case "A216 WCB": case "LCB": case "LCC": case "LC1": case "LC2":
            $val = "СТ"; break;
        case "304": case "316": case "316L": case "316Ti": case "321": case "904L": case "2205": 
            $val = "НЖ"; break;
        case "GG25": 
            $val = "СЧ"; break;
        case "AI":
            $val = "А"; break;
        case "PP": 
            $val = "П"; break;
        default: $val = "?";
    }
    $bar_code .= $val . "-";
    
    // Материал седлового уплотнения корпуса
    switch ($data['seat']['value']) {
        case "EPDM": $val = "E"; break;
        case "SBR": $val = "S"; break;
        case "NBR": $val = "HT"; break;
        case "Silicon": $val = "C"; break;
        case "Viton": $val = "B"; break;
        case "PTFE": $val = "Ф"; break;
        case "Metal+Graphite": case "304": case "316": case "2205": case "inconel": 
            $val = "M"; break;
        default: $val = "?";
    }
    $bar_code .= $val . "-";
    
    // Материал диска
    switch ($data['disc']['value']) {
        case "GGG40": $val = "ВЧ"; break;
        case "A216 WCB": $val = "СТ"; break;
        case "304": case "316": case "304L": case "316L": case "316Ti": case "321": case "904L": case "2205": 
            $val = "НЖ"; break;
        default: $val = "?";
    }
    $bar_code .= $val . "-";
    

    $bar_code .= "X-XXX-X-XXXX X";
    return $bar_code;
}

function fillCells($sheet, $data) {

    $sheet->setCellValue('B1', $data['dn']['value']."--".$data['dn']['else']);
    $sheet->setCellValue('B2', $data['pressure']['dn']['value']);
    $sheet->setCellValue('B3', $data['standard']['value']);
    $sheet->setCellValue('B4', $data['design']['value']);
    $sheet->setCellValue('B5', $data['connection']['value']);
    $sheet->setCellValue('B6', $data['connection_standart']['value']);

    $face = $data['face']['value'];
    switch ($face) {
        case 'ГОСТ': $sheet->setCellValue('B7', $face."--".$data['face']['ГОСТ']['type']['value']);
            break;
        case 'EN': $sheet->setCellValue('B7', $face."--".$data['face']['EN']['type']['value']);
            break;
        case 'ASME': $sheet->setCellValue('B7', $face."--".$data['face']['ASME']['type']['value']);
            break;
    }
    $sheet->setCellValue('B8', $data['ftf']['value']);
    $sheet->setCellValue('B9', $data['body']['value']);
    $sheet->setCellValue('B10', $data['disk']['value']);
    $sheet->setCellValue('B11', $data['seat']['value']);
    $sheet->setCellValue('B12', $data['stem']['value']);
    $sheet->setCellValue('B13', $data['tightness']['value']);
    $sheet->setCellValue('B14', $data['working_pressure']['value']);
    $sheet->setCellValue('B15', $data['working_temperature']['value']);
    $sheet->setCellValue('B16', $data['working_ambient']['value']);
    $sheet->setCellValue('B17', $data['valve_purpose']['value']); 
    $sheet->setCellValue('B18', $data['media']['value']);
    $sheet->setCellValue('B19', $data['operation']['value']);
    $sheet->setCellValue('B20', $data['quantity']['value']);
    $sheet->setCellValue('B21', $data['notes']['value']);


}

function sendMail($filename) {
    
	global $from_email,	$recipient_email, $reply_to_email;
    

	//Load POST data from HTML form
	$subject	 = 'Опросный лист на дисковый затвор'; //subject for the email
	$message	 = '.. тут какой-то сопровождающий текст ..'; //body of the email
	
	
	//Get uploaded file data using $_FILES array
	$tmp_name = $filename; // get the temporary file name of the file on the server
	$name = $filename; // get the name of the file
	/*
    $size	 = $_FILES['my_file']['size']; // get size of the file for size validation
	$type	 = $_FILES['my_file']['type']; // get type of the file
	$error	 = $_FILES['my_file']['error']; // get the error (if any)
    */

	//read from the uploaded file & base64_encode content
	$handle = fopen($tmp_name, "r"); // set the file handle only for reading the file
	$content = fread($handle, filesize($filename)); // reading the file
	fclose($handle);				 // close upon completion

	$encoded_content = chunk_split(base64_encode($content));

	$boundary = md5("random"); // define boundary with a md5 hashed value

	//header
	$headers = "MIME-Version: 1.0\r\n"; // Defining the MIME version
	$headers .= "From:".$from_email."\r\n"; // Sender Email
	$headers .= "Reply-To: ".$reply_to_email."\r\n"; // Email address to reach back
	$headers .= "Content-Type: multipart/mixed;"; // Defining Content-Type
	$headers .= "boundary = $boundary\r\n"; //Defining the Boundary
		
	//plain text
	$body = "--$boundary\r\n";
	$body .= "Content-Type: text/plain; charset=UTF-8\r\n";
	$body .= "Content-Transfer-Encoding: base64\r\n\r\n";
	$body .= chunk_split(base64_encode($message));
		
	//attachment
	$body .= "--$boundary\r\n";
	$body .="Content-Type: ; name=".$name."\r\n";
	$body .="Content-Disposition: attachment; filename=".$name."\r\n";
	$body .="Content-Transfer-Encoding: base64\r\n";
	$body .="X-Attachment-Id: ".rand(1000, 99999)."\r\n\r\n";
	$body .= $encoded_content; // Attaching the encoded file with email
	
	$sentMailResult = mail($recipient_email, $subject, $body, $headers);

	if($sentMailResult)	{
	    echo "Отпавил файл на " . $recipient_email;
	    unlink($name); // delete the file after attachment sent.
	} else {
	    die("Чота не так с отправкой на почту!");
	}
}

function translate($text) {

    global $dictionary;

    $flag = false;
    foreach ($dictionary  as $key => $value )
        if ( $text == $key ) {
            $flag = true;
            return $value;
        }
    
    if ( $flag == false )
        return $text;
}







?>