<?php
require_once 'vendor/autoload.php';

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\Exception;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
//use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

$data = $_POST;

//$sender = new Sender();
//$sender->init();


// создаем копию general template
$spreadsheet = new Spreadsheet();
//$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
//$spreadsheet = $reader->load("general template.xlsx");

//$writer = new Xlsx($spreadsheet);
//$writer->save('result.xlsx');

print_r($data['dn']['value']);

//print_r($data);
return;



class Sender{
    public function __construct()
    {
        
    }

    public function init(){
        $data = $_GET;


        $result = [];

        $dn = [];
        foreach ($data as $key => $datum){
            switch ($datum['id']) {
                case 5: //Условный диаметр DN или NPS
                    $value = 'Тип: '. $datum['value'];
                    $diameterType = $datum['value'];
                    if($datum['else']){
                        $value .= '| Значение: '. $datum['else'];
                        $diameterType = $datum['value'];
                    }

                    $result[]  = [
                        'label' => $datum['label'],
                        'value' => $value,
                    ];
                    $dn = $datum;
                    break;
                case 6:
                    if($dn['value'] == 'DN'){
                        $v = $datum['dn']['value'];
                        $diameterValue = $datum['dn']['value'];
                        if($datum['dn']['value']=='else'){
                            $v = $datum['dn']['else'];
                            $diameterValue = $datum['dn']['else'];
                        }
                        $value = $v;
                        $label = $datum['dn']['label'];
                    }

                    if($dn['value'] == 'NPS'){
                        $v = $datum['dn']['value'];
                        $diameterValue = $v;
                        if($datum['dn']['value']=='else'){
                            $v = $datum['dn']['else'];
                            $diameterValue = $v;
                        }
                        $value =  $v;
                        $label = $datum['nps']['label'];
                    }
                    $result[]  = [
                        'label' => $label,
                        'value' => $value,
                    ];
                    break;
                case 11: // face

                    $value = $datum[$datum['value']]['type']['value'];
                    $result[]  = [
                        'label' => $datum['label'],
                        'value' => $value,
                    ];
                    break;
                default:
                    $value = $datum['value'];
                    if($datum['value']=='else'){
                        $value = $datum['value'];
                    }
                    $result[]  = [
                        'label' => $datum['label'],
                        'value' => $value,
                    ];
                    break;
            }
        }



        /*
        $keys = [];
        $values = [];
        
        $doc = [$keys,$values];
        */
        $spreadsheet = new Spreadsheet();
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load("general template.xlsx");
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('F38', $result[19]['value']); // Count of positions
        unset($result[19]);
        /*foreach ($result as $item){
            $keys[] =  $item['label'];
            $values[] =  $item['value'];
        }*/
        $cellNum = 38;
        //$arr = [$keys,$values];
        foreach ($result as $r){
            $sheet->setCellValue('D'.$cellNum, $r['label'].': '.$r['value']);
            $cellNum++;
        }
        $codeOfPars = 'AKB-'.$diameterType.$diameterValue.'-'.$result[8]['value'].'-'.$result[10]['value'];
        $sheet -> setCellValue('B39',$codeOfPars);

        //print_r($result);
        $writer = new Xlsx($spreadsheet);
        $writer->save('result.xlsx');
        
        //print_r($doc);
        //$xlsx = SimpleXLSXGen::fromArray( $doc );
        


        $filename = 'result.xlsx';

       // $xlsx->saveAs($filename);
        //$this->sendMail($filename);
        //$actual_link = (isset($_SERVER['HTTPS']) && $_SERVER['HTTPS'] === 'on' ? "https" : "http") . "://$_SERVER[HTTP_HOST]";
        //header('Location: '.$actual_link.'/emrdev?status=true');

    }


    private function sendMail($filename){

        $email = new PHPMailer();
        $email->setFrom('info@akvalit.by', 'Аквалит');
        $email->Subject   = 'Опросный лист на дисковый затвор';
        $email->Body      = 'Опросный лист на дисковый затвор';
        $email->addAddress( 'dev.emr@yandex.ru');
        $email->addAddress( 'info@akvalit.by');
        
 

        $email->AddAttachment( $filename , 'result.xlsx' );

        return $email->Send();
    }

}


