<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_FILES["input_csv"]) && isset($_FILES["result_data_csv"])) {
  $input_csv = $_FILES["input_csv"]["tmp_name"];
  $data_csv = $_FILES["result_data_csv"]["tmp_name"];

  if ($input_csv=="" || $data_csv=="") {
    die("file upload error");
  }

  $filenames_by_id = [];

  foreach (array_map('str_getcsv', file($input_csv)) as $obj) {
    if (strlen($obj[0])==32) {
      $filenames_by_id[$obj[0]] = $obj[1];
    }
  }

  $rows = [];

  if (($file = fopen($data_csv, "r")) !== FALSE) {
    if (fgets($file, 4) !== "\xef\xbb\xbf") rewind($file); //Skip BOM if present

    while (($data = fgetcsv($file, 0, ";", "\"" , "\\")) !== FALSE) {
      if (strlen($data[0])!=32) continue; //id should be of length 32, so skip first item

      $row = [];
      $row["id"] = $data[0];
      $row["bestandsnaam"] = $filenames_by_id[$row["id"]];

      foreach (json_decode($data[2]) as $item) {
        foreach  ($item[0] as $key => $value) {
          $row[$key] = $value;
        }
      }
      
      $rows[] = $row;
    }
    fclose($file);
  }

  $spreadsheet = new Spreadsheet();
  $sheet = $spreadsheet->getActiveSheet();
  $sheet->fromArray(array_keys($rows[0]),NULL,'A1'); //header
  $sheet->fromArray($rows,NULL,'A2'); //rows
  $writer = new Xlsx($spreadsheet);
  header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  header('Content-Disposition: attachment; filename="result.xlsx"');
  $writer->save('php://output');
  die();

} else {
  ?>
  <style>body { font-family: sans-serif; }</style>
  <h1>Resultaten van HetVolk.org naar Excel</h1>
  <form method="post" enctype="multipart/form-data">
  <p>CSV met bestandsnamen voor HetVolk: <input type="file" name="input_csv" accept=".csv" required></p>
  <p>CSV met resultaten van HetVolk: <input type="file" name="result_data_csv" accept=".csv" required></p>
  <input type="submit" value="Start">
  </form>
<?php
}
?>

