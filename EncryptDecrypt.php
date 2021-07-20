<?php
require_once ('lib/nusoap.php');

$client = new nusoap_client("http://localhost:8000/WebService.asmx?WSDL",true);

  $s = json_encode($a);


  $key = substr(hash('sha256', $key, true), 0, 32);

  $iv = chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0) . chr(0x0);


  $encrypted = base64_encode(openssl_encrypt($s, $method, $key, OPENSSL_RAW_DATA, $iv));



  $result = $client->call("BuscarCedula",array('model'=>$encrypted));



  $decrypted = openssl_decrypt(base64_decode($result['BuscarCedulaResult']), $method, $key, OPENSSL_RAW_DATA, $iv);


  $exit = json_decode($decrypted);

  echo $exit->Persona->CEDULA;
  echo $exit->Persona->PRIMER_NOMBRE;
  echo $exit->Persona->SEGUNDO_NOMBRE;
  echo $exit->Persona->APELLIDO_PATERNO;
  echo $exit->Persona->APELLIDO_MATERNO;
  echo $exit->Persona->SEXO;
  echo $exit->Persona->FECHA_NACIMIENTO;


  ?>
