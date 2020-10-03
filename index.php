<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;

$ruta = 'archivo/usuarios.xlsx';
$mail = new PHPMailer(true);

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
$spreadsheet = $reader->load($ruta);

$sheet = $spreadsheet->getActiveSheet();

$rows =[];

foreach($sheet->getRowIterator() as $row){
	$cellIterator = $row -> getCellIterator();
	$cellIterator->setIterateOnlyExistingCells(false);
	$cells =[];
	foreach ($cellIterator as $cell) {
		$cells[]= $cell->getValue();
	}
	if (isset($keys)) {
    	$rows[] = array_combine($keys, $cells);
	} else {
    		$keys = $cells;
	}

}


$mail->SMTPDebug = SMTP::DEBUG_SERVER;                      // Enable verbose debug output
    $mail->isSMTP();                                            // Send using SMTP
    $mail->Host       = 'smtp.mailtrap.io';                    // Set the SMTP server to send through
    $mail->SMTPAuth   = true;                                   // Enable SMTP authentication
    $mail->Username   = '0e230db69ad6d5';                     // SMTP username
    $mail->Password   = '2d26475acbfd7c';                               // SMTP password
    $mail->SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS;         // Enable TLS encryption; `PHPMailer::ENCRYPTION_SMTPS` encouraged
    $mail->Port       = 587;                                    // TCP port to connect to, use 465 for `PHPMailer::ENCRYPTION_SMTPS` above

foreach ($rows as $key => $value) {
	
    //Recipients
    $mail->setFrom('c3bcc505bb-440891@inbox.mailtrap.io', 'Agrosty');
  //  $mail->addAddress($value['mail'], $value['nombre']);     // Add a recipient
 //   $mail->addAddress('example@test.com');               // Name is optional
 //   $mail->addReplyTo('info@example.com', 'Information');
 //   $mail->addCC('cc@example.com');
    $mail->addBCC($value['mail']);

    //contenido
    $mail->isHTML(true);                                  // Set email format to HTML
    $mail->Subject = 'Usuario y Password';
    $mail->Body    = 'Su usuario es: '.$value['usuario'].' '.'y su password es '.$value['password'];
    $mail->AltBody = 'Su usuario es: '.$value['usuario'].' '.'y su password es '.$value['password'];

    $mail->send();
}