<?php
Use PhpOffice\PhpSpreadsheet\IOFactory;
Use PhpOffice\PhpSpreadsheet\Spreadsheet;
Use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
Use PhpOffice\PhpSpreadsheet\Calculation\TextData\Replace;
require_once('vendor/autoload.php');
$spreadsheet = new Spreadsheet();


if(isset($_POST['sq1'])){
    $searchquery = $_POST['sq1'];
    $data = explode(",", $searchquery);
    $results[] = array();
    $queryString = http_build_query([
        'api_key' => 'demo',
        'q' => $searchquery
      ]);
      
      //https://api.valueserp.com/search?api_key=demo&q=
      $ch = curl_init(sprintf('%s?%s', 'https://api.valueserp.com/search', $queryString));
      
      curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
      curl_setopt($ch, CURLOPT_FOLLOWLOCATION, true);
      curl_setopt($ch, CURLOPT_SSL_VERIFYHOST, false);
      curl_setopt($ch, CURLOPT_SSL_VERIFYPEER, false);
      curl_setopt($ch, CURLOPT_TIMEOUT, 180);
      
      $api_result = curl_exec($ch);
      curl_close($ch);
      
      $results[] = json_decode($api_result, true);
      $request_info = $results[1]['request_info']['success'];
      $organic_results = $results[1]['organic_results'];
      //print_r($organic_results);exit();
      if($request_info){
        
        $spreadsheet->getProperties()->setTitle("excelsheet");
        $spreadsheet->setActiveSheetIndex(0);
        $spreadsheet->getActiveSheet()->SetCellValue('A1', 'Title');
        $spreadsheet->getActiveSheet()->SetCellValue('B1', 'Link');
        $spreadsheet->getActiveSheet()->SetCellValue('C1', 'snippet');
        $spreadsheet->getActiveSheet()
            ->getStyle("A1:F1")
            ->getFont()
            ->setBold(true);
        $rowCount = 2;
        if (! empty($organic_results)) {
            foreach($organic_results as $sub_row){
                $spreadsheet->getActiveSheet()->setCellValue("A" . $rowCount, $sub_row["title"]);
                $spreadsheet->getActiveSheet()->setCellValue("B" . $rowCount, $sub_row["link"]);
                $spreadsheet->getActiveSheet()->setCellValue("C" . $rowCount, $sub_row["snippet"]);
                
                $rowCount ++;
            }
            }
            $spreadsheet->getActiveSheet()
                ->getStyle('A:C')
                ->getAlignment()
                ->setWrapText(true);

            $spreadsheet->getActiveSheet()
                ->getRowDimension($rowCount)
                ->setRowHeight(- 1);
        
        $writer = IOFactory::createWriter($spreadsheet, 'Xls');
        header('Content-Type: text/xls');
        $fileName = 'exported_excel_' . time() . '.xls';
        $headerContent = 'Content-Disposition: attachment;filename="' . $fileName . '"';
        header($headerContent);
        $writer->save('php://output');
          
    }

}

?>