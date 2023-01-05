<?php
// memory_limit = 512M
// upload_max_filesize = 50M
// post_max_size = 50M


require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_FILES["input_csv"]) && isset($_FILES["result_data_csv"])) {
  $input_csv = $_FILES["input_csv"]["tmp_name"];
  $data_csv = $_FILES["result_data_csv"]["tmp_name"];

  if ($input_csv=="" || $data_csv=="") {
    die("Probleem bij het uploaden van de bestanden");
  }

  // lees alle bestandsnamen in
  $filenames_by_id = [];
  foreach (array_map('str_getcsv', file($input_csv)) as $obj) {
    if (strlen($obj[0])==32) {
      $filenames_by_id[$obj[0]] = $obj[1]!="" ? $obj[1] : $obj[5]; # bestandsnaam kan 2e of 5e kolom zijn afhankelijk of het een oorspronkelijk of een 'merged' lijst is.
    }
  }

  $rows = [];
  $fieds = [];

  header("Content-type: text/plain");

  if (($file = fopen($data_csv, "r")) !== FALSE) {
    if (fgets($file, 4) !== "\xef\xbb\xbf") rewind($file); //Skip BOM if present

    while (($data = fgetcsv($file, 0, ";", "\"" , "\\")) !== FALSE) {
      // var_dump($data); //strlen($data[0]);
      // die();
      if (strlen($data[0])!=32) continue; //id should be of length 32, so skip first item
      
      if (!array_key_exists($data[0],$filenames_by_id) || $filenames_by_id[$data[0]]=="") {
        die("Probleem: ID niet gevonden in de lijst met bestandsnamen (of lege bestandsnaam)");
      }

      $json = json_decode($data[2]);

      $containsArray = false;
      foreach ($json as $items) {
        if (is_array($items)) {
          !$containsArray = true;
          break;
        }
      }

      if (!$containsArray) {
        $item = $json;
        $json = new stdClass;
        $json->item = [ $item ] ;
      }

      // var_dump($json);
      // die();

      foreach ($json as $items) { //scans

        foreach ($items as $item) { //records per scan
          $row = [];
          $row["id"] = $data[0]; 
          $row["user"] = $data[3]; //RC 2022-01-05
          $row["tijdstip"] = $data[4]; //RC 2022-01-05
          foreach  ($item as $key => $value) { //columns
            $row[$key] = $value;
            $fields[$key] = $key; //collect al column names
          }
          $rows[] = $row;
        }
        
      }

    }
    fclose($file);
  }

  // var_dump($rows);
  // die();

  $fields[] = "bestandsnaam"; //add column
  $fields[] = "id"; //add column
  $fields[] = "user"; //add column
  $fields[] = "tijdstip"; //add column
  $fields[] = "link"; //add column

  $spreadsheet = new Spreadsheet();
  $sheet = $spreadsheet->getActiveSheet();
  $sheet->fromArray($fields,NULL,'A1'); //header

  $r=2; //start at row 2
  foreach ($rows as $row) {

    $row["bestandsnaam"] = $filenames_by_id[$row["id"]];

    # strip .jpg.cropX uit de betandsnaam
    $row["bestandsnaam"] = preg_replace("/\.jpg\.crop\d{1,2}/", "", $row["bestandsnaam"]);

    $row["link"] = "https://crowd.hetutrechtsarchief.nl/" . $row["id"];

    $c=1; //start at col 1
    foreach ($fields as $field) {
      $value = array_key_exists($field,$row) ? $row[$field] : "";
      $cell = $sheet->getCellByColumnAndRow($c, $r);
      $cell->setValue($value);      

      if ($field=="link") {
        $sheet->getCellByColumnAndRow($c, $r)->getHyperlink()->setUrl($value);
      }
      $c++;
    }
    $r++;
  }

  //add Auto Filter
  $sheet->setAutoFilter($sheet->calculateWorksheetDimension());

  //set column width
  $sheet->getDefaultColumnDimension()->setWidth(100, 'pt');


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
  <small>
    <li>De CSV met bestandsnamen mag de oorspronkelijke CSV zijn maar ook de lijst met 'merged' in de naam.</li>
    <li>In de uiteindelijke spreadsheet wordt '.jpg.cropX' automatisch verwijderd uit de bestandsnaam.</li>
  </small>
<?php
}
?>

