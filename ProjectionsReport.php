<?php

/* 
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
namespace Cis;

use PHPExcel;
use PHPExcel_IOFactory;
use PHPMailer;
/* Include Date Classes  */
use DateTime;
use DatePeriod;
use DateTimeZone;
use DateInterval;

class ProjectionsReport
{
  /**
   * Retrieve report data for entry into Excel report
   * 
   * @param type $year
   * @param int $month
   * @param type $db
   * @param type $msDb
   * @param type $weekNum
   * @param type $multi
   * @return type
   */
  function pullReportData($year,$month,$db,$msDb,$weekNum,$multi = FALSE)
  {
    if($multi){
      $actuals     = $this->getFromGL($year,$month,$msDb);
      $summaries[] = $this->getSummaries($year,$month,$actuals,$db,$weekNum);
      // REPEAT FOR PREVIOUS MONTH
      if($month <= 1){
        $month = 12;
        $year--;
      } else {
        $month--;   
      }
      $wn = (floor(cal_days_in_month(CAL_GREGORIAN, $month, $year)/7)+1);
      $wn = ($wn > 5)? 5: $wn;
      $actuals2    = $this->getFromGL($year,$month,$msDb);
      $summaries[] = $this->getSummaries($year,$month,$actuals2,$db,$wn);
    } else {
      $actuals   = $this->getFromGL($year,$month,$msDb);
      $summaries = $this->getSummaries($year,$month,$actuals,$db,$weekNum);
    }
    return $summaries;
  }

  /**
   * Query from GL for Actuals
   * 
   * @param type $year
   * @param type $month
   * @param type $msDb
   * @return type
   */
  function getFromGL($year, $month, $msDb)
  {   
      $sql = "SELECT 
                  A.DivisionMgr,
                  A.ACCOUNTTYPE,
                  SUM(A.AMOUNT) AS Amount
              FROM GLPL_Detail_Report A
              LEFT OUTER JOIN JC20001 B ON (
                  A.JRNENTRY = B.JRNENTRY 
                  AND A.ACTINDX = B.ACTINDX 
                  AND A.SEQNUMBR = B.LNSEQNBR
                  )
              WHERE year(A.TRXDATE) in (".$year.") 
                  and MONTH(A.TRXDATE) in (".$month.")
              GROUP BY
                  A.DivisionMgr,
                  A.ACCOUNTTYPE
              ORDER BY
                  A.DivisionMgr,
                  A.ACCOUNTTYPE";

      $result = $msDb->query($sql);

      $fromGL = array();

      while($row = $msDb->fetch_array($result)) {
          $fromGL[trim($row["DivisionMgr"])][$row["ACCOUNTTYPE"]] = $row["Amount"];
      }

      return $fromGL; 
  }

  /**
   * Returns the Summaries of projections
   * 
   * @param type $year
   * @param type $month
   * @param type $actualsArray
   * @param type $db
   * @return string
   */
  function getSummaries($year, $month, $actualsArray, $db, $weekNum = 1)
  {
      if(empty($weekNum)){
        $tmpwk = (floor(date('d')/7));
        $weekNum = ($tmpwk < 1)? 1 : $tmpwk;
      }
    
        $sql = "
            SELECT 
                mpbw.id,
                mpbw.divId, 
                IF(
                  INSTR(u.username, '.'),
                  CONCAT(SUBSTRING_INDEX(u.username, '.', 1), ' ', SUBSTRING_INDEX(u.username, '.', -1)),
                  u.username
                ) as username,
                SUM(mpbw.ARForecast) as ARForecast, 
                SUM(mpbw.ARWeek1) as ARWeek1, 
                SUM(mpbw.ARWeek2) as ARWeek2, 
                SUM(mpbw.ARWeek3) as ARWeek3, 
                SUM(mpbw.ARWeek4) as ARWeek4, 
                SUM(mpbw.ARWeek5) as ARWeek5, 
                SUM(mpbw.ARFinal) as RevChanges,
                SUM(mpbw.APForecast) as APForecast, 
                SUM(mpbw.APWeek1) as APWeek1, 
                SUM(mpbw.APWeek2) as APWeek2, 
                SUM(mpbw.APWeek3) as APWeek3, 
                SUM(mpbw.APWeek4) as APWeek4, 
                SUM(mpbw.APWeek5) as APWeek5,
                SUM(mpbw.APFinal) as ExpChanges
            FROM (
                SELECT
                id,
                divId, 
                ARForecast,
                CASE WHEN ARWeek1 >= 999999999
                  THEN 0
                  ELSE ARWeek1 END AS ARWeek1,
                CASE WHEN ARWeek2 >= 999999999
                  THEN 0
                  ELSE ARWeek2 END AS ARWeek2,
                CASE WHEN ARWeek3 >= 999999999
                  THEN 0
                  ELSE ARWeek3 END AS ARWeek3,
                CASE WHEN ARWeek4 >= 999999999
                  THEN 0
                  ELSE ARWeek4 END AS ARWeek4,
                CASE WHEN ARWeek5 >= 999999999
                  THEN 0
                  ELSE ARWeek5 END AS ARWeek5,
                CASE WHEN ISNULL(weekInput) THEN ARForecast
                  WHEN ARWeek".$weekNum." = 999999999 THEN 0
                  ELSE ARWeek".$weekNum." END AS ARFinal,
                APForecast,
                CASE WHEN APWeek1 >= 999999999
                  THEN 0
                  ELSE APWeek1 END AS APWeek1,
                CASE WHEN ARWeek2 >= 999999999
                  THEN 0
                  ELSE APWeek2 END AS APWeek2,
                CASE WHEN APWeek3 >= 999999999
                  THEN 0
                  ELSE APWeek3 END AS APWeek3,
                CASE WHEN APWeek4 >= 999999999
                  THEN 0
                  ELSE APWeek4 END AS APWeek4,
                CASE WHEN APWeek5 >= 999999999
                  THEN 0
                  ELSE APWeek5 END AS APWeek5,
                CASE WHEN ISNULL(weekInput) THEN APForecast
                  WHEN APWeek".$weekNum." = 999999999 THEN 0
                  ELSE APWeek".$weekNum." END AS APFinal 
                FROM MonthlyProjectionsByWeek
                WHERE year = ".$year."
                  AND month = ".$month."
            ) AS mpbw
            LEFT JOIN Users u ON (mpbw.divId = u.id)
            GROUP BY
              mpbw.divId,
              u.username
            ORDER BY 
              u.username ASC; 
            ";
        
        error_log($sql);
        
        if(!$db){
          $db = Db::getConn();
        }
        
        $result = $db->query($sql)->fetchAll(\PDO::FETCH_ASSOC);
        
        $dlRecords = array();
        
        foreach ($result as $row) {
            $tmpItems = [
                'lineId'     => $row['id'],
                'divId'      => $row['divId'],
                'username'   => $row['username'],
                'ARForecast' => $row['ARForecast'],
                'ARWeek1'    => $row['ARWeek1'],
                'ARWeek2'    => $row['ARWeek2'],
                'ARWeek3'    => $row['ARWeek3'],
                'ARWeek4'    => $row['ARWeek4'],
                'ARWeek5'    => $row['ARWeek5'],
                'RevChanges' => $row['RevChanges'],
                'ARActual'   => $actualsArray[trim($row['username'])]['AR'],
                'APForecast' => $row['APForecast'],
                'APWeek1'    => $row['APWeek1'],
                'APWeek2'    => $row['APWeek2'],
                'APWeek3'    => $row['APWeek3'],
                'APWeek4'    => $row['APWeek4'],
                'APWeek5'    => $row['APWeek5'],
                'ExpChanges' => $row['ExpChanges'],
                'APActual'   => $actualsArray[trim($row['username'])]['AP'],
                ];
            
            $dlRecords[] = $tmpItems;
        }
        
        foreach ( $actualsArray as $key => $aa ) {
            
            if ( !$this->inArrayR(trim($key), $dlRecords) && $key != "" && $key != "House") {
                $tmpItems = [
                    'username'   => $key,
                    'ARForecast' => '0',
                    'ARWeek1'    => '0',
                    'ARWeek2'    => '0',
                    'ARWeek3'    => '0',
                    'ARWeek4'    => '0',
                    'ARWeek5'    => '0',
                    'ARActual'   => $aa['AR'],
                    'APForecast' => '0',
                    'APWeek1'    => '0',
                    'APWeek2'    => '0',
                    'APWeek3'    => '0',
                    'APWeek4'    => '0',
                    'APWeek5'    => '0',
                    'APActual'   => $aa['AP'],
                    ];
                
                $dlRecords[] = $tmpItems;
            }
        }
        
        // BLACK MAGIC SORT MULTI_ARRAY FUNCTION === No idea how this is supposed to work
        array_multisort(
            array_map(
                function($element) {
                    return $element['username'];
                }, $dlRecords), SORT_ASC, $dlRecords);
        
        return $dlRecords;
  }

  /**
   * Generate Excel Report for distribution
   * 
   * @param type $templateVars
   */
  function generateExcel($repData, $year, $db, $msDb, $download = FALSE)
  {
      // Create new PHPExcel object
      $objPHPExcel = new PHPExcel;

      // Get report date
      $reportDate = new DateTime(gmdate('dMY')." Friday", new DateTimeZone(date_default_timezone_get()));
      $mnthName = $reportDate->format('F');

      // Set document properties
      $objPHPExcel->getProperties()
              ->setTitle("Operations Projections - Production Report - ".$mnthName." ".$year)
              ->setSubject($mnthName." Summary");

      // Assign Worksheet
      $sheet = $objPHPExcel->getActiveSheet();
      $mnthIdx = 0;
      if(isset($repData[$mnthIdx])){
        // Build Monthlies
        do{
          if($mnthIdx > 0){
            $mnthName = $reportDate->modify('-1 month')->format('F');
            if($mnthName == 'Decemeber'){
              $year--;
            }
            $sheet = $objPHPExcel->createSheet($mnthIdx); //Setting index when creating
          }
          $this->buildMonthlySheet($sheet, $repData[$mnthIdx], $db, $reportDate->format('F'), $year);
          $mnthIdx++;
        } while( $mnthIdx < 2 );
        
        // Build Weeklies
        $mnthIdx = 0;
        $weekNum = floor(date('j')/7)+1;
        //reset Report Date due to modify call
        $reportDate->modify('+1 month');
        do{
          if($mnthIdx > 0){
            $mnthName = $reportDate->modify('-1 month')->format('F');
            if($mnthName == 'Decemeber'){
              $year--;
            }
            $weekNum = 5;
          }
          $sheet = $objPHPExcel->createSheet($mnthIdx+2); //Setting index when creating
          $weeklyData = $this->getProjections('ALL', $year, $reportDate->format('n'), $weekNum,  $db, $msDb);
          $this->buildWeeklySheet($sheet,$weeklyData,$reportDate->format('F'),$year,$reportDate->format('n'), $download);
          $mnthIdx++;
        } while( $mnthIdx < 2 );
        
        $objPHPExcel->setActiveSheetIndex(0);
      } else {
        $this->buildMonthlySheet($sheet, $repData, $db, $mnthName, $year);
      }
      
      if($download){
        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="Operations_Projections_'.gmdate('dMY').'.xls"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        
        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
      } else {
        $xlFile = './Operations_Projections_'.gmdate('dMY').'.xls';

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save($xlFile);

        error_log("Excel report created successfully.");

        return $xlFile;
      }
  }
  
  /**
   * 
   * @param type $sheet
   * @param type $repData
   */
  function buildMonthlySheet($sheet, $repData, $db, $mnthName, $year)
  {
    // Setting Dimensions
    $sheet->getColumnDimension('A')->setWidth(2);
    $sheet->getColumnDimension('B')->setWidth(2);
    $sheet->getColumnDimension('C')->setWidth(20);
    $sheet->getColumnDimension('D')->setWidth(20);
    $sheet->getColumnDimension('E')->setWidth(20);
    $sheet->getColumnDimension('F')->setWidth(20);
    $sheet->getColumnDimension('G')->setWidth(2);
    $sheet->getColumnDimension('H')->setWidth(20);
    $sheet->getColumnDimension('I')->setWidth(2);
    $sheet->getColumnDimension('J')->setWidth(20);

    $sheet->getColumnDimension('K')->setWidth(2);

    $sheet->getColumnDimension('L')->setWidth(20);
    $sheet->getColumnDimension('M')->setWidth(20);
    $sheet->getColumnDimension('N')->setWidth(20);
    $sheet->getColumnDimension('O')->setWidth(2);
    $sheet->getColumnDimension('P')->setWidth(20);
    $sheet->getColumnDimension('Q')->setWidth(2);
    $sheet->getColumnDimension('R')->setWidth(20);
    $sheet->getColumnDimension('S')->setWidth(5);

    $sheet->getColumnDimension('T')->setWidth(10);
    $sheet->getColumnDimension('U')->setWidth(10);
    $sheet->getColumnDimension('V')->setWidth(10);
    
    $sheet->getStyle('A1:V3')->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_CENTER );

    $style1 = array(
            'fill' => array(
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'E6E6E6'),
            )
        );
    $style2 = array(
            'fill' => array(
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => 'C5D9F1'),
            )
        );
    $style3 = array(
            'fill' => array(
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '1F497D'),
            )
        );
    $style4 = array(
            'font'  => array(
              'bold'  => true,
              'size' => '12',
            ),
            'borders' => array(
                'allborders' => array(
                    'style' => \PHPExcel_Style_Border::BORDER_THIN
                )
            )
        );
    $style5 = array(
            'font' => array(
              'bold'  => true,
              'italic'=> true,
              'size'  => '12',
            )
        );
    /***********************************************************************
    $sheet->getStyle('A1')->applyFromArray($styleArray);
    ************************************************************************/

    /**********************************************************************/
    // Projection Labels
    /**********************************************************************/
    
    // Setup Title and Column Headers
    $sheet->mergeCells('D1:J1');
    $sheet->mergeCells('L1:R1');
    
    $sheet->setTitle($mnthName." ".$year);
    $sheet->setCellValueByColumnAndRow(2, 1, $mnthName." ".$year);
    $sheet->setCellValueByColumnAndRow(3, 1, "PROJECT REVENUES");
    $sheet->setCellValueByColumnAndRow(11, 1, "PROJECT EXPENSES");

    $sheet->setCellValueByColumnAndRow(2, 2, " ");
    $sheet->setCellValueByColumnAndRow(3, 2, "Initial Forecast");
    $sheet->setCellValueByColumnAndRow(4, 2, "Adjusted Forecast");
    $sheet->setCellValueByColumnAndRow(5, 2, "Final Forecast");
    $sheet->setCellValueByColumnAndRow(7, 2, "Actual");
    $sheet->setCellValueByColumnAndRow(9, 2, "Remaining");

    $sheet->setCellValueByColumnAndRow(11, 2, "Initial Forecast");
    $sheet->setCellValueByColumnAndRow(12, 2, "Adjusted Forecast");
    $sheet->setCellValueByColumnAndRow(13, 2, "Final Forecast");
    $sheet->setCellValueByColumnAndRow(15, 2, "Actual");
    $sheet->setCellValueByColumnAndRow(17, 2, "Remaining");

    $sheet->setCellValueByColumnAndRow(19, 2, "Rev");
    $sheet->setCellValueByColumnAndRow(20, 2, "Exp");
    $sheet->setCellValueByColumnAndRow(21, 2, "MARGIN");

    /**********************************************************************/
    // Summary rows   
    /**********************************************************************/
    $row = 4;
    foreach( $repData as $prj ) {
      $tmpMonth = new DateTime(gmdate('dMY'), new DateTimeZone(date_default_timezone_get()));
      $dl=($tmpMonth->format('F') === $mnthName)?FALSE:TRUE;
      $this->addSummaryLine($sheet, $row, $prj, $db, $dl);
      $row++;
    }
    $row++;

    /**********************************************************************/
    // Totals row   
    /**********************************************************************/
    $sheet->setCellValueByColumnAndRow(2, $row, "Totals");
    $sheet->setCellValueByColumnAndRow(3, $row, "=SUM(D4:D".($row-1).")");
    $sheet->setCellValueByColumnAndRow(4, $row, "=SUM(E4:E".($row-1).")");
    $sheet->setCellValueByColumnAndRow(5, $row, "=SUM(F4:F".($row-1).")");
    $sheet->setCellValueByColumnAndRow(7, $row, "=SUM(H4:H".($row-1).")");
    $sheet->setCellValueByColumnAndRow(9, $row, "=SUM(J4:J".($row-1).")");

    $sheet->setCellValueByColumnAndRow(11, $row, "=SUM(L4:L".($row-1).")");
    $sheet->setCellValueByColumnAndRow(12, $row, "=SUM(M4:M".($row-1).")");
    $sheet->setCellValueByColumnAndRow(13, $row, "=SUM(N4:N".($row-1).")");
    $sheet->setCellValueByColumnAndRow(15, $row, "=SUM(P4:P".($row-1).")");
    $sheet->setCellValueByColumnAndRow(17, $row, "=SUM(R4:R".($row-1).")");

    $sheet->setCellValueByColumnAndRow(19, $row, "=H".($row)."/F".($row));
    $sheet->setCellValueByColumnAndRow(20, $row, "=P".($row)."/N".($row));
    $sheet->setCellValueByColumnAndRow(21, $row, "=(F".($row)."-N".($row).")/F".($row));


    /**********************************************************************/
    // STYLE ASSIGNMENTS 
    /**********************************************************************/
    $sheet->getStyle('C1:V'.$row)->applyFromArray($style4);
    $sheet->getStyle('C1:C'.$row)->applyFromArray($style5);        

    // Light Blue
    $sheet->getStyle('D1')->applyFromArray($style2);
    $sheet->getStyle('L1')->applyFromArray($style2);
    $sheet->getStyle('T1:V2')->applyFromArray($style2);
    $sheet->getStyle('H2:H'.$row)->applyFromArray($style2);
    $sheet->getStyle('P2:P'.$row)->applyFromArray($style2);
    $sheet->getStyle('D'.$row.':J'.$row)->applyFromArray($style2);
    $sheet->getStyle('L'.$row.':R'.$row)->applyFromArray($style2);

    //Grey
    $sheet->getStyle('D2:F'.$row)->applyFromArray($style1);
    $sheet->getStyle('L2:N'.$row)->applyFromArray($style1);


/**************************************************************************/
    // Finalize the work sheet and send for upload

    $sheet->getStyle('A3:V'.$row)->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT );

    // Setting number formats
    $sheet->getStyle('D3:K'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
    $sheet->getStyle('L3:R'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
    $sheet->getStyle('T4:V'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

    $sheet->getPageSetup()->setPrintArea('B1:V'.$row);
    
    $sheet->getStyle('A1');
  }

  /**
   * 
   * @param type $sheet
   * @param type $lineNum
   * @param type $row
   */
  function addSummaryLine($sheet,$lineNum,$row, $db, $download = FALSE)
  {
    $ARAdjust = $row["RevChanges"];
    $ARFinal = ($ARAdjust == 0) ? $row["ARForecast"] : $ARAdjust;

    $APAdjust = $row["ExpChanges"];
    $APFinal = ($APAdjust == 0) ? $row["APForecast"] : $APAdjust;

    $sheet->setCellValueByColumnAndRow(2, $lineNum, trim(ucwords($row["username"])));

    $sheet->setCellValueByColumnAndRow(3, $lineNum, $row["ARForecast"]);
    $sheet->setCellValueByColumnAndRow(4, $lineNum, $ARAdjust);
    $sheet->setCellValueByColumnAndRow(5, $lineNum, $ARFinal);
    $sheet->setCellValueByColumnAndRow(7, $lineNum, $row["ARActual"]);
    $sheet->setCellValueByColumnAndRow(9, $lineNum, $ARFinal-$row["ARActual"]);

    $sheet->setCellValueByColumnAndRow(11, $lineNum, $row["APForecast"]);
    $sheet->setCellValueByColumnAndRow(12, $lineNum, $APAdjust);
    $sheet->setCellValueByColumnAndRow(13, $lineNum, $APFinal);
    $sheet->setCellValueByColumnAndRow(15, $lineNum, ($row["APActual"]*-1));
    $sheet->setCellValueByColumnAndRow(17, $lineNum, $APFinal-($row["APActual"]*-1)); 

    $sheet->setCellValueByColumnAndRow(19, $lineNum, ($ARFinal == 0 ? 0 : $row["ARActual"]/$ARFinal));
    $sheet->setCellValueByColumnAndRow(20, $lineNum, ($ARFinal == 0 ? 0 : ($row["APActual"]*-1)/$APFinal)); 
    $sheet->setCellValueByColumnAndRow(21, $lineNum, ($ARFinal == 0 ? 0 : ($ARFinal-$APFinal)/$ARFinal));
    
    if(!$download){
      $this->addToOperationsTable($row, $ARFinal, $APFinal, $db);   
    }
  }

  /**
   * Generate Image from Excel File
   * 
   * @param type $filePath
   * @return type
   */
  function generateImage($filePath)
  {
      $endpoint = "https://sandbox.zamzar.com/v1/jobs";
      $apiKey = "59372c1d03d40144eee7c2df13ad0c30816228f9";
      $targetFormat = "png";

      // Since PHP 5.5+ CURLFile is the preferred method for uploading files
      if(function_exists('curl_file_create')) {
        $sourceFile = curl_file_create($filePath);
      } else {
        $sourceFile = '@' . realpath($filePath);
      }

      $postData = array(
        "source_file" => $sourceFile,
        "target_format" => $targetFormat
      );

      $ch = curl_init(); // Init curl
      curl_setopt($ch, CURLOPT_URL, $endpoint); // API endpoint
      curl_setopt($ch, CURLOPT_CUSTOMREQUEST, 'POST');
      curl_setopt($ch, CURLOPT_POSTFIELDS, $postData);
      curl_setopt($ch, CURLOPT_SAFE_UPLOAD, false); // Enable the @ prefix for uploading files
      curl_setopt($ch, CURLOPT_RETURNTRANSFER, true); // Return response as a string
      curl_setopt($ch, CURLOPT_USERPWD, $apiKey . ":"); // Set the API key as the basic auth username
      $body = curl_exec($ch);
      curl_close($ch);

      $response = json_decode($body, true);

      $stat = "";
      //Waiting....
      $waiting = true;
      while($waiting)
      {
          sleep(2);
          $stat = $this->getFileStatus($response["id"],$apiKey);
          $waiting = ($stat["status"]=="successful")?false:true;
          echo "Awaiting response --- Status : ".$stat["status"]." --- ".date("m.d.Y H:i:s")."\r\n";
      }
      $targetFile = $stat["target_files"][0]["name"];
      sleep(5);

      return $this->getImageFile($stat["target_files"][0]["id"],$apiKey,$targetFile);
  }

  /**
   * Get current status of file to image conversion
   * 
   * @param type $jobID
   * @param type $apiKey
   * @return type
   */
  function getFileStatus($jobID,$apiKey)
  {
      $endpoint = "https://sandbox.zamzar.com/v1/jobs/$jobID";

      $ch = curl_init(); // Init curl
      curl_setopt($ch, CURLOPT_URL, $endpoint); // API endpoint
      curl_setopt($ch, CURLOPT_RETURNTRANSFER, true); // Return response as a string
      curl_setopt($ch, CURLOPT_USERPWD, $apiKey . ":"); // Set the API key as the basic auth username
      $body = curl_exec($ch);
      curl_close($ch);

      $job = json_decode($body, true);

      echo print_r($job);
      echo "\r\n";

      return $job;
  }

  /**
   * Pull finished image from server
   * 
   * @param type $fileID
   * @param type $apiKey
   * @return boolean
   */
  function getImageFile($fileID,$apiKey,$localFilename)
  {
      $endpoint = "https://sandbox.zamzar.com/v1/files/$fileID/content";

      $ch = curl_init(); // Init curl
      curl_setopt($ch, CURLOPT_URL, $endpoint); // API endpoint
      curl_setopt($ch, CURLOPT_USERPWD, $apiKey . ":"); // Set the API key as the basic auth username
      curl_setopt($ch, CURLOPT_FOLLOWLOCATION, TRUE);

      $fh = fopen($localFilename, "wb");
      curl_setopt($ch, CURLOPT_FILE, $fh);
      curl_exec($ch);
      curl_close($ch);
      fclose($fh);

      echo "Image downloaded - ".$localFilename." \r\n";

      return $localFilename;
  }

  /**
   * Get DL Addresses to send 
   * 
   * @param type $pdo
   * @return type
   */
  function pullAddresses($pdo,$addEmails)
  {
      $sql = "
      SELECT 
          u.emailAddr 
      FROM UserSecurityLevels usl 
      LEFT JOIN Users u ON( u.id = usl.userId )
      WHERE usl.levelId = 128
      ";  // 128 is the code for Projections permissions

      $addresses = $pdo->query($sql)->fetchAll();

      foreach ($addresses as $addr) {
          $addEmails[] = strtolower($addr["emailAddr"]);
      }

      $cleanEmails = array_unique($addEmails);
      echo print_r($cleanEmails);
      echo "\r\n";
      error_log("Addresses compiled.");

      return $cleanEmails;
  }

  /**
   * Build and send email
   * 
   * @param type $xlFile
   * @param type $addys
   * @param type $subject
   */
  function buildEmail($xlFile,$imageFile,$addys)
  {
      $email = new PHPMailer();
      $email->From     = 'noreply@powerhouseretail.com';
      $email->FromName = 'CMS Projections Notification';
      $email->Subject  = 'Weekly Automated Operations Projections Update';

      foreach ($addys as $addr) {
          $email->AddAddress($addr);
      }

      //$email->AddAddress("scott.richardson@powerhouseretail.com");
      $email->AddAttachment($xlFile,'Operations Projections Report.xls');

      $email->IsHTML(true);
      $email->AddEmbeddedImage($imageFile, 'OpProjReport');
      $email->Body = "<p>Attached is the Automated Weekly Operations Projections Update spreadsheet.<br/>
              For real-time updates, please follow the link below for the current project summary.<br/>
              <a href='https://cms.powerhouseretail.com/cis/public/projections/WeeklySummary?page=summary'>
                  Operations Projections Report
              </a>
              </p>
              <p>
              Thank You,<br />
              <b>
                  Powerhouse Retail Services, LLC
              </b>
              </p>
              <hr/>
              <p><img src='cid:OpProjReport' width='100%' height='100%' /></p>";

      error_log($email->Body);
      $email->Send();

      echo "===================================== \r\n";
      echo "===  SENT MAILING  === \r\n";
  }

  /**
   * is the specified entry inside array?  True/False
   * 
   * @param type $needle
   * @param type $haystack
   * @param type $strict
   * @return boolean
   */
  function inArrayR($needle, $haystack, $strict = false) 
  {
      foreach ($haystack as $item) {
          if (($strict ? $item === $needle : $item == $needle) || (is_array($item) && $this->inArrayR($needle, $item, $strict))) {
              return true;
          }
      }

      return false;
  }
  
  function buildExtraInfo($sheet)
  {
    // Setup
    // loop
    // Totals
  }
  
  /*****************************************************************************
   * OPERATIONS REPORT FUNCTIONS
   ****************************************************************************/
  
  /**
   * Add Monday static data to Operations table for Thursday pull report
   * 
   * @param ResultSet $row
   * @param int $ARFinal
   * @param int $APFinal
   * @param \PDO $db
   * @return bool
   */
  function addToOperationsTable($row, $ARFinal, $APFinal, $db)
  {
    $sql = "
      INSERT INTO WeeklyOperationsReport
      ( divId, divName, revenue, expenses, created )
      VALUES
      (".$row["divId"].",'".$row["username"]."',".$ARFinal.",".$APFinal.",NOW())
      ";
    
    return $db->exec($sql);
  }
  
  /**
  * Retrieve Operations data for entry into Thursday Excel report
  * 
  * @param type $year
  * @param type $month
  * @param type $db
  * @param type $msDb
  * @return type
  */
  function pullOperationsData($year,$month,$db,$msDb)
  {
    $actuals    = $this->getFromGL($year,$month,$msDb);
    $batch      = $this->getBatchedAmounts($msDb);
    $operations = $this->getFromOperationsTable($actuals,$batch,$db);

    return $operations;
  }
  
  /**
   * Pull Monday static data from Operations table for Thursday report
   * 
   * @param type $db
   * @param type $actualsArray
   * @return type
   */   
  function getFromOperationsTable($actualsArray, $batch, $db)
  {
    $sql = "
      SELECT 
        *
      FROM(
        SELECT 
          *
        FROM WeeklyOperationsReport
        WHERE DATEDIFF(NOW(), created) < 5
        ORDER BY
          id,
          created,
          divName
      ) AS wor
      GROUP BY divId
      ORDER BY divName;
      ";
    
    $result = $db->query($sql)->fetchAll(\PDO::FETCH_ASSOC);
        
    $dlRecords = array();
    
    foreach ($result as $row) {
      $tmpItems = [
        'divId'     => $row['divId'], 
        'divName'   => $row['divName'], 
        'revenue'   => $row['revenue'], 
        'expenses'  => $row['expenses'],
        // Actuals
        'ARActual'  => $actualsArray[trim($row['divName'])]['AR'],
        'APActual'  => $actualsArray[trim($row['divName'])]['AP'],
        // Batched
        'ARBatched' => $batch[trim($row['divName'])]['Billing'],
        'APBatched' => $batch[trim($row['divName'])]['Payout'],
      ];

      $dlRecords[] = $tmpItems;
    }
    
    return $dlRecords;
  }
  
  /**
   * Pulls batched amounts from Infinity tables
   * 
   * @param type $msDb
   * @return type
   */
  function getBatchedAmounts($msDb)
  {
    $sql = "  
      SELECT 
        p.[Division Mgr] as Manager,
        SUM(ib.BatchTotal) as BatchTotal,
        ib.BatchType
      FROM dbo.Projectlist p

      JOIN ( 
        SELECT  
          GPProjectID,
          SUM(BatchTotal) AS BatchTotal,
          BatchType

        FROM(
          -- AR Billing Section			
          SELECT	
            Title as Project,
            BillingBatchID,
            PayOutBatchID = null,
            GPProjectID,
            BatchDate,
            DaysTakenToInvoice,
            ISNULL(TaskBilling + InventoryBilling + LaborBilling, 0) as BatchTotal,
            BatchType = 'Billing'
          FROM (
            SELECT	
              p.Title,
              bb.BillingBatchID,
              bbgpid.GPProjectID,
              bb.BatchDate,
              CASE
                WHEN bb.IsManualTax = 1 THEN NULL
                ELSE CAST(DATEDIFF(day, bb.BatchDate, bb.DateSentToGP) as varchar(10))
              END as DaysTakenToInvoice,
              ISNULL(task.TaskBilling, 0) as TaskBilling,
              ISNULL(CASE
                    WHEN p.AreInventoryItemsBatched = 1 THEN inventory.InventoryBilling
                    ELSE 0
                  END, 0) as InventoryBilling,
              ISNULL(labor.LaborBilling, 0) as LaborBilling

            FROM	
              [Powerhouse.Portal].[dbo].[BillingBatch] bb
              JOIN [Powerhouse.Portal].[dbo].[Project] p ON p.ProjectID = bb.ProjectID
              LEFT JOIN [Powerhouse.Portal].[dbo].[BillingBatchToGPProjectID] bbgpid ON (bbgpid.BillingBatchID = bb.BillingBatchID)
              LEFT JOIN (Select BillingBatchID, Sum(ActualBilling * QuantityOfItems) as TaskBilling FROM [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] Group By BillingBatchID) task on task.BillingBatchID = bb.BillingBatchID
              LEFT JOIN (Select BillingBatchID, Sum(CalculatedBilling) as InventoryBilling From [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] pttu join [Powerhouse.Portal].[dbo].[InventoryItemToProjectTaskToUnit] ii on ii.ProjectTaskToUnitID = pttu.ProjectTaskToUnitID Group By BillingBatchID) inventory on inventory.BillingBatchID = bb.BillingBatchID
              LEFT JOIN (Select BillingBatchID, Sum(CalculatedBilling) as LaborBilling From [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] pttu2 join [Powerhouse.Portal].[dbo].[LaborToProjectTaskToUnit] l on l.ProjectTaskToUnitID = pttu2.ProjectTaskToUnitID Group By BillingBatchID) labor on labor.BillingBatchID = bb.BillingBatchID
            WHERE
              YEAR(bb.BatchDate) = YEAR(getdate())
              AND MONTH(bb.BatchDate) = MONTH(getdate())
          ) AS tbl

          UNION

          -- AP Payout Section
          SELECT	
            Title as Project,
            BillingBatchID = null,
            PayOutBatchID,
            GPProjectID,
            BatchDate,
            DaysTakenToInvoice,
            ISNULL(TaskPayout + InventoryPayout + LaborPayout + PayoutTax, 0) as BatchTotal,
            BatchType = 'Payout'
          FROM(
            SELECT	
              p.Title,
              pb.PayOutBatchID,
              pbgpid.GPProjectID,
              pb.BatchDate,
              DATEDIFF(day, pb.BatchDate, pb.DateSentToGP) as DaysTakenToInvoice,
              ISNULL(task.TaskPayout, 0) as TaskPayout,
              tax.PayoutTax,
              ISNULL(CASE
                    WHEN p.AreInventoryItemsBatched = 1 AND p.ShowW2Employees = 0 THEN inventory.InventoryPayout
                    ELSE 0
                  END, 0) as InventoryPayout,
              ISNULL(CASE
                    WHEN p.ShowW2Employees = 0 THEN labor.LaborPayout
                    ELSE 0 
                  END, 0) as LaborPayout

            FROM	
              [Powerhouse.Portal].[dbo].[PayoutBatch] pb
              JOIN [Powerhouse.Portal].[dbo].[Project] p on p.ProjectID = pb.ProjectID
              LEFT JOIN [Powerhouse.Portal].[dbo].[PayoutBatchToGPProjectID] pbgpid ON (pbgpid.PayoutBatchID = pb.PayOutBatchID)
              LEFT JOIN (Select PayoutBatchID, Sum(ActualPayout * QuantityOfItems) as TaskPayout FROM [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] Group By PayoutBatchID) task on task.PayoutBatchID = pb.PayoutBatchID
              LEFT JOIN (Select PayoutBatchID, Sum(PayoutTax) as PayoutTax FROM [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] Group By PayoutBatchID) tax on tax.PayoutBatchID = pb.PayoutBatchID
              LEFT JOIN (Select PayoutBatchID, Sum(CalculatedPayout) as InventoryPayout From [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] pttu join [Powerhouse.Portal].[dbo].[InventoryItemToProjectTaskToUnit] ii on ii.ProjectTaskToUnitID = pttu.ProjectTaskToUnitID Group By PayoutBatchID) inventory on inventory.PayoutBatchID = pb.PayoutBatchID
              LEFT JOIN (Select PayoutBatchID, Sum(CalculatedPayout) as LaborPayout From [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] pttu2 join [Powerhouse.Portal].[dbo].[LaborToProjectTaskToUnit] l on l.ProjectTaskToUnitID = pttu2.ProjectTaskToUnitID Group By PayoutBatchID) labor on labor.PayoutBatchID = pb.PayoutBatchID
            WHERE
              YEAR(pb.BatchDate) = YEAR(getdate())
              AND MONTH(pb.BatchDate) = MONTH(getdate()) 
          ) as tbl
        ) as batchQry
        WHERE DaysTakenToInvoice IS NULL
        GROUP BY 
          GPProjectID,
          BatchType
      ) ib ON (ib.GPProjectID = p.[Proj. No.])
      GROUP BY
        p.[Division Mgr],
        ib.BatchType
        ";
    
    $result = $msDb->query($sql);
        
    $batchRecords = array();

    while($row = $msDb->fetch_array($result)) {
      $batchRecords[trim($row["Manager"])][$row["BatchType"]] = $row["BatchTotal"];
    }
    
    return $batchRecords;
  }
  
  /**
   * Pulls batched amounts from Infinity tables
   * 
   * @param type $msDb
   * @return type
   */
  function getBatchedDetails($msDb)
  {
    $sql = "  
      SELECT 
        p.[Division Mgr] as Manager,
        ib.GPProjectID,
        SUM(ib.BatchTotal) as BatchTotal,
        ib.BatchType
      FROM dbo.Projectlist p

      JOIN ( 
        SELECT  
          GPProjectID,
          SUM(BatchTotal) AS BatchTotal,
          BatchType

        FROM(
          -- AR Billing Section			
          SELECT	
            Title as Project,
            BillingBatchID,
            PayOutBatchID = null,
            GPProjectID,
            BatchDate,
            DaysTakenToInvoice,
            ISNULL(TaskBilling + InventoryBilling + LaborBilling, 0) as BatchTotal,
            BatchType = 'Billing'
          FROM (
            SELECT	
              p.Title,
              bb.BillingBatchID,
              bbgpid.GPProjectID,
              bb.BatchDate,
              CASE
                WHEN bb.IsManualTax = 1 THEN NULL
                ELSE CAST(DATEDIFF(day, bb.BatchDate, bb.DateSentToGP) as varchar(10))
              END as DaysTakenToInvoice,
              ISNULL(task.TaskBilling, 0) as TaskBilling,
              ISNULL(CASE
                    WHEN p.AreInventoryItemsBatched = 1 THEN inventory.InventoryBilling
                    ELSE 0
                  END, 0) as InventoryBilling,
              ISNULL(labor.LaborBilling, 0) as LaborBilling

            FROM	
              [Powerhouse.Portal].[dbo].[BillingBatch] bb
              JOIN [Powerhouse.Portal].[dbo].[Project] p ON p.ProjectID = bb.ProjectID
              LEFT JOIN [Powerhouse.Portal].[dbo].[BillingBatchToGPProjectID] bbgpid ON (bbgpid.BillingBatchID = bb.BillingBatchID)
              LEFT JOIN (Select BillingBatchID, Sum(ActualBilling * QuantityOfItems) as TaskBilling FROM [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] Group By BillingBatchID) task on task.BillingBatchID = bb.BillingBatchID
              LEFT JOIN (Select BillingBatchID, Sum(CalculatedBilling) as InventoryBilling From [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] pttu join [Powerhouse.Portal].[dbo].[InventoryItemToProjectTaskToUnit] ii on ii.ProjectTaskToUnitID = pttu.ProjectTaskToUnitID Group By BillingBatchID) inventory on inventory.BillingBatchID = bb.BillingBatchID
              LEFT JOIN (Select BillingBatchID, Sum(CalculatedBilling) as LaborBilling From [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] pttu2 join [Powerhouse.Portal].[dbo].[LaborToProjectTaskToUnit] l on l.ProjectTaskToUnitID = pttu2.ProjectTaskToUnitID Group By BillingBatchID) labor on labor.BillingBatchID = bb.BillingBatchID
            WHERE
              YEAR(bb.BatchDate) = YEAR(getdate())
              AND MONTH(bb.BatchDate) = MONTH(getdate())
          ) AS tbl

          UNION

          -- AP Payout Section
          SELECT	
            Title as Project,
            BillingBatchID = null,
            PayOutBatchID,
            GPProjectID,
            BatchDate,
            DaysTakenToInvoice,
            ISNULL(TaskPayout + InventoryPayout + LaborPayout + PayoutTax, 0) as BatchTotal,
            BatchType = 'Payout'
          FROM(
            SELECT	
              p.Title,
              pb.PayOutBatchID,
              pbgpid.GPProjectID,
              pb.BatchDate,
              DATEDIFF(day, pb.BatchDate, pb.DateSentToGP) as DaysTakenToInvoice,
              ISNULL(task.TaskPayout, 0) as TaskPayout,
              tax.PayoutTax,
              ISNULL(CASE
                    WHEN p.AreInventoryItemsBatched = 1 AND p.ShowW2Employees = 0 THEN inventory.InventoryPayout
                    ELSE 0
                  END, 0) as InventoryPayout,
              ISNULL(CASE
                    WHEN p.ShowW2Employees = 0 THEN labor.LaborPayout
                    ELSE 0 
                  END, 0) as LaborPayout

            FROM	
              [Powerhouse.Portal].[dbo].[PayoutBatch] pb
              JOIN [Powerhouse.Portal].[dbo].[Project] p on p.ProjectID = pb.ProjectID
              LEFT JOIN [Powerhouse.Portal].[dbo].[PayoutBatchToGPProjectID] pbgpid ON (pbgpid.PayoutBatchID = pb.PayOutBatchID)
              LEFT JOIN (Select PayoutBatchID, Sum(ActualPayout * QuantityOfItems) as TaskPayout FROM [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] Group By PayoutBatchID) task on task.PayoutBatchID = pb.PayoutBatchID
              LEFT JOIN (Select PayoutBatchID, Sum(PayoutTax) as PayoutTax FROM [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] Group By PayoutBatchID) tax on tax.PayoutBatchID = pb.PayoutBatchID
              LEFT JOIN (Select PayoutBatchID, Sum(CalculatedPayout) as InventoryPayout From [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] pttu join [Powerhouse.Portal].[dbo].[InventoryItemToProjectTaskToUnit] ii on ii.ProjectTaskToUnitID = pttu.ProjectTaskToUnitID Group By PayoutBatchID) inventory on inventory.PayoutBatchID = pb.PayoutBatchID
              LEFT JOIN (Select PayoutBatchID, Sum(CalculatedPayout) as LaborPayout From [Powerhouse.Portal].[dbo].[ProjectTaskToUnit] pttu2 join [Powerhouse.Portal].[dbo].[LaborToProjectTaskToUnit] l on l.ProjectTaskToUnitID = pttu2.ProjectTaskToUnitID Group By PayoutBatchID) labor on labor.PayoutBatchID = pb.PayoutBatchID
            WHERE
              YEAR(pb.BatchDate) = YEAR(getdate())
              AND MONTH(pb.BatchDate) = MONTH(getdate()) 
          ) as tbl
        ) as batchQry
        WHERE DaysTakenToInvoice IS NULL
        GROUP BY 
          GPProjectID,
          BatchType
      ) ib ON (ib.GPProjectID = p.[Proj. No.])
      GROUP BY
        p.[Division Mgr],
        ib.GPProjectID,
        ib.BatchType
      ORDER BY
        p.[Division Mgr],
        ib.GPProjectID,
        ib.BatchType
        ";
    
    $result = $msDb->query($sql);
        
    $batchRecords = array();

    while($row = $msDb->fetch_array($result)) {
      $batchRecords[] = array(
        "Manager"     => trim($row["Manager"]),
        "GPProjectID" => $row["GPProjectID"],
        "BatchTotal"  => $row["BatchTotal"],
        "BatchType"   => $row["BatchType"]
      );
    }
    
    return $batchRecords;
  }
  
  /**
   * Generate Excel Report for distribution
   * 
   * @param type $templateVars
   */
  function generateOperationsReport($repData, $year, $month, $db, $msDb, $download = FALSE)
  {
      // Create new PHPExcel object
      $objPHPExcel = new PHPExcel;

      // Get report date
      $reportDate = new DateTime(gmdate('dMY')." Thursday", new DateTimeZone(date_default_timezone_get()));
      $mnthName = $reportDate->format('F');

      // Set document properties
      $objPHPExcel->getProperties()
              ->setTitle("Operations Projections - Production Report - ".$mnthName." ".$year)
              ->setSubject($mnthName." Summary");

      // Assign Worksheet
      $sheet = $objPHPExcel->getActiveSheet();

      // Setting Dimensions
      $sheet->getColumnDimension('A')->setWidth(2);
      $sheet->getColumnDimension('B')->setWidth(2);
      $sheet->getColumnDimension('C')->setWidth(20);
      $sheet->getColumnDimension('D')->setWidth(20);
      $sheet->getColumnDimension('E')->setWidth(2);
      $sheet->getColumnDimension('F')->setWidth(20);
      $sheet->getColumnDimension('G')->setWidth(2);
      $sheet->getColumnDimension('H')->setWidth(20);
      $sheet->getColumnDimension('I')->setWidth(2);
      $sheet->getColumnDimension('J')->setWidth(20);
      $sheet->getColumnDimension('K')->setWidth(2);
      $sheet->getColumnDimension('L')->setWidth(20);
      
      //Separator
      $sheet->getColumnDimension('M')->setWidth(10);
      
      $sheet->getColumnDimension('N')->setWidth(20);
      $sheet->getColumnDimension('O')->setWidth(2);
      $sheet->getColumnDimension('P')->setWidth(20);
      $sheet->getColumnDimension('Q')->setWidth(2);
      $sheet->getColumnDimension('R')->setWidth(20);
      $sheet->getColumnDimension('S')->setWidth(2);
      $sheet->getColumnDimension('T')->setWidth(20);
      $sheet->getColumnDimension('U')->setWidth(2);
      $sheet->getColumnDimension('V')->setWidth(20);
      
      $sheet->getColumnDimension('W')->setWidth(5);
      
      $sheet->getColumnDimension('X')->setWidth(10);
      $sheet->getColumnDimension('Y')->setWidth(10);
      $sheet->getColumnDimension('Z')->setWidth(10);

      // Setup Title and Column Headers
      $sheet->mergeCells('D1:L1');
      $sheet->mergeCells('N1:V1');

      //$sheet->getStyle('A1')->getAlignment()->setWrapText(true);
      $sheet->getStyle('A1:Z3')->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_CENTER );

      $style1 = array(
              'fill' => array(
                  'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                  'color' => array('rgb' => 'E6E6E6'),
              )
          );
      $style2 = array(
              'fill' => array(
                  'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                  'color' => array('rgb' => 'C5D9F1'),
              )
          );
      $style3 = array(
              'fill' => array(
                  'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                  'color' => array('rgb' => '1F497D'),
              )
          );
      $style4 = array(
              'font'  => array(
                'bold'  => true,
                'size' => '12',
              ),
              'borders' => array(
                  'allborders' => array(
                      'style' => \PHPExcel_Style_Border::BORDER_THIN
                  )
              )
          );
      $style5 = array(
              'font' => array(
                'bold'  => true,
                'italic'=> true,
                'size'  => '12',
              )
          );
      /***********************************************************************
      $sheet->getStyle('A1')->applyFromArray($styleArray);
      ************************************************************************/

      /**********************************************************************/
      // Projection Labels
      /**********************************************************************/
      $sheet->setCellValueByColumnAndRow(2, 1, "Thru:");
      $sheet->setCellValueByColumnAndRow(3, 1, "PROJECT REVENUES");
      $sheet->setCellValueByColumnAndRow(13, 1, "PROJECT EXPENSES");

      $sheet->setCellValueByColumnAndRow(2, 2, date("m/d/Y", strtotime("today")));
      $sheet->setCellValueByColumnAndRow(3, 2, "Forecast");
      $sheet->setCellValueByColumnAndRow(5, 2, "Actual");
      $sheet->setCellValueByColumnAndRow(7, 2, "Batched");
      $sheet->setCellValueByColumnAndRow(9, 2, "Actual/Batched Total");
      $sheet->setCellValueByColumnAndRow(11, 2, "Remaining");

      $sheet->setCellValueByColumnAndRow(13, 2, "Forecast");
      $sheet->setCellValueByColumnAndRow(15, 2, "Actual");
      $sheet->setCellValueByColumnAndRow(17, 2, "Batched");
      $sheet->setCellValueByColumnAndRow(19, 2, "Actual/Batched Total");
      $sheet->setCellValueByColumnAndRow(21, 2, "Remaining");

      $sheet->setCellValueByColumnAndRow(23, 2, "Rev");
      $sheet->setCellValueByColumnAndRow(24, 2, "Exp");
      $sheet->setCellValueByColumnAndRow(25, 2, "MARGIN");

      /**********************************************************************/
      // Summary rows   
      /**********************************************************************/
      $row = 4;
      foreach( $repData as $prj ) {
          $this->addOperationsLine($sheet, $row, $prj, $db);
          $row++;
      }
      $row++;

      /**********************************************************************/
      // Totals row   
      /**********************************************************************/
      $sheet->setCellValueByColumnAndRow(2, $row, "Totals");
      $sheet->setCellValueByColumnAndRow(3, $row, "=SUM(D4:D".($row-1).")");
      $sheet->setCellValueByColumnAndRow(5, $row, "=SUM(F4:F".($row-1).")");
      $sheet->setCellValueByColumnAndRow(7, $row, "=SUM(H4:H".($row-1).")");
      $sheet->setCellValueByColumnAndRow(9, $row, "=SUM(J4:J".($row-1).")");
      $sheet->setCellValueByColumnAndRow(11, $row, "=SUM(L4:L".($row-1).")");
      
      $sheet->setCellValueByColumnAndRow(13, $row, "=SUM(N4:N".($row-1).")");
      $sheet->setCellValueByColumnAndRow(15, $row, "=SUM(P4:P".($row-1).")");
      $sheet->setCellValueByColumnAndRow(17, $row, "=SUM(R4:R".($row-1).")");
      $sheet->setCellValueByColumnAndRow(19, $row, "=SUM(T4:T".($row-1).")");
      $sheet->setCellValueByColumnAndRow(21, $row, "=SUM(V4:V".($row-1).")");

      $sheet->setCellValueByColumnAndRow(23, $row, "=F".($row)."/D".($row));
      $sheet->setCellValueByColumnAndRow(24, $row, "=P".($row)."/N".($row));
      $sheet->setCellValueByColumnAndRow(25, $row, "=(D".($row)."-N".($row).")/D".($row));
      
      /**********************************************************************/
      // STYLE ASSIGNMENTS 
      /**********************************************************************/
      $sheet->getStyle('C1:Z'.$row)->applyFromArray($style4);
      $sheet->getStyle('C1:C'.$row)->applyFromArray($style5);        

      // Light Blue
      $sheet->getStyle('D1')->applyFromArray($style2);
      $sheet->getStyle('N1')->applyFromArray($style2);
      $sheet->getStyle('X1:Z2')->applyFromArray($style2);
      $sheet->getStyle('F2:F'.$row)->applyFromArray($style2);
      $sheet->getStyle('P2:P'.$row)->applyFromArray($style2);
      $sheet->getStyle('D'.$row.':L'.$row)->applyFromArray($style2);
      $sheet->getStyle('N'.$row.':V'.$row)->applyFromArray($style2);

      //Grey
      $sheet->getStyle('D2:D'.$row)->applyFromArray($style1);
      $sheet->getStyle('N2:N'.$row)->applyFromArray($style1);

  /**************************************************************************/
      // Finalize the work sheet and send for upload

      $sheet->getStyle('A3:Z'.$row)->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT );

      // Setting number formats
      $sheet->getStyle('D3:L'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
      $sheet->getStyle('N3:V'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
      $sheet->getStyle('X4:Z'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);

      $sheet->getPageSetup()->setPrintArea('B1:Z'.$row);
      
      // Create New Sheet
      $sheet2 = $objPHPExcel->createSheet(1);
      $this->buildOperationDetailSheet( $sheet2, $this->getBatchedDetails($msDb), $mnthName );
      
      $sheet->getStyle('A1');
      
      if($download){
        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="Operations_With_Batch_'.gmdate('dMY').'.xls"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        
        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
      } else {
        $xlFile = './Operations_With_Batch_'.gmdate('dMY').'.xls';

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save($xlFile);

        error_log("Excel report created successfully.");

        return $xlFile;
      }
  }
  
  /**
   * This function handles the formating and assignment of weekly reports in Excel 
   * 
   * @param type $sheet
   * @param type $repData
   */
  function buildOperationDetailSheet($sheet,$repData,$mnthName)
  {
      $sheet->setTitle("Batch Total Data");
      // Setting Dimensions
      $sheet->getColumnDimension('A')->setWidth(1);
      $sheet->getColumnDimension('B')->setWidth(1);
      $sheet->getColumnDimension('C')->setWidth(30);
      $sheet->getColumnDimension('D')->setWidth(15);
      $sheet->getColumnDimension('E')->setWidth(20);
      $sheet->getColumnDimension('F')->setWidth(15);
      
      $sheet->getStyle('A1:F1')->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_CENTER );

      /**********************************************************************/
      // Labels
      /**********************************************************************/
      
      // Setup Title and Column Headers
      $sheet->setCellValueByColumnAndRow(2, 1, "Division Lead");
      $sheet->setCellValueByColumnAndRow(3, 1, "Project ID");
      $sheet->setCellValueByColumnAndRow(4, 1, "Batch Total");
      $sheet->setCellValueByColumnAndRow(5, 1, "Batch Type");
      
      /**********************************************************************/
      // Summary rows   
      /**********************************************************************/
      $row = 2;
      foreach( $repData as $prj ) {
          $this->addDetailLine($sheet, $row, $prj);
          $row++;
      }
      $row++;

      $sheet->getStyle('E2:E'.$row)->getNumberFormat()->setFormatCode( \PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
      $sheet->getStyle('D1:D'.$row)->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_CENTER );
      $sheet->getStyle('F1:F'.$row)->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_CENTER );
  }
  
  /**
   * 
   * @param type $sheet
   * @param type $lineNum
   * @param type $row
   */
  function addOperationsLine($sheet,$lineNum,$row, $db)
  { 
    $sheet->setCellValueByColumnAndRow(2, $lineNum, trim(ucwords($row["divName"])));

    $sheet->setCellValueByColumnAndRow( 3, $lineNum, $row["revenue"]);
    $sheet->setCellValueByColumnAndRow( 5, $lineNum, $row["ARActual"]);
    $sheet->setCellValueByColumnAndRow( 7, $lineNum, $row["ARBatched"]);
    $sheet->setCellValueByColumnAndRow( 9, $lineNum, $row["ARActual"]+$row["ARBatched"]);
    $sheet->setCellValueByColumnAndRow(11, $lineNum, $row["revenue"]-($row["ARActual"]+$row["ARBatched"]));
    
    $sheet->setCellValueByColumnAndRow(13, $lineNum, $row["expenses"]);
    $sheet->setCellValueByColumnAndRow(15, $lineNum, $row["APActual"] * -1 );
    $sheet->setCellValueByColumnAndRow(17, $lineNum, ($row["APBatched"])); 
    $sheet->setCellValueByColumnAndRow(19, $lineNum, ($row["APActual"] * -1) +$row["APBatched"]);
    $sheet->setCellValueByColumnAndRow(21, $lineNum, $row["expenses"]-(($row["APActual"] * -1)+$row["APBatched"]));
    
    $sheet->setCellValueByColumnAndRow(23, $lineNum, ( $row["revenue"] == 0 ? 0 : ($row["ARActual"]/$row["revenue"]) ) );
    $sheet->setCellValueByColumnAndRow(24, $lineNum, ( $row["revenue"] == 0 ? 0 : (($row["APActual"] * -1)/$row["expenses"]) ) );
    $sheet->setCellValueByColumnAndRow(25, $lineNum, ( $row["revenue"] == 0 ? 0 : ($row["revenue"]-$row["expenses"])/$row["revenue"]) );
  }
  
  /**
   * 
   * @param type $sheet
   * @param type $lineNum
   * @param type $row
   */
  function addDetailLine($sheet,$lineNum,$row)
  { 
    $sheet->setCellValueByColumnAndRow( 2, $lineNum, trim(ucwords($row["Manager"])));
    $sheet->setCellValueByColumnAndRow( 3, $lineNum, $row["GPProjectID"]);
    $sheet->setCellValueByColumnAndRow( 4, $lineNum, $row["BatchTotal"]);
    $sheet->setCellValueByColumnAndRow( 5, $lineNum, $row["BatchType"]);
  }
  
  /**
   * Build and send email
   * 
   * @param type $xlFile
   * @param type $addys
   * @param type $subject
   */
  function buildOperationsEmail($xlFile,$imageFile,$addys)
  {
      $email = new PHPMailer();
      $email->From     = 'noreply@powerhouseretail.com';
      $email->FromName = 'Operations Projections Notification';
      $email->Subject  = 'Weekly Automated Operations Projections Update - Thursday with Batched';

      foreach ($addys as $addr) {
          $email->AddAddress($addr);
      }

      $email->AddAttachment($xlFile,'ThursdayOperationsReport.xls');

      $email->IsHTML(true);
      $email->AddEmbeddedImage($imageFile, 'OpProjReport');
      $email->Body = "<p>Attached is the Automated Thursday Operations Projections with Batched data spreadsheet.<br/>
              </p>
              <p>
              Thank You,<br />
              <b>
                  Powerhouse Retail Services, LLC
              </b>
              </p>
              <hr/>
              <p><img src='cid:OpProjReport' width='100%' height='100%' /></p>";

      error_log($email->Body);
      $email->Send();

      echo "===================================== \r\n";
      echo "===  SENT MAILING  === \r\n";
  }
  
  /*****************************************************************************
   * WEEKLY REPORT FUNCTIONS
   *****************************************************************************/
  
  /**
   * Generate Weekly Excel Report for DLs
   * 
   * @param type $templateVars
   */
  function generateWeeklyReport($repData, $year,$month, $db, $download = FALSE)
  {
      // Create new PHPExcel object
      $objPHPExcel = new PHPExcel;

      // Get report date
      $reportDate = new DateTime($month."/01/".$year. " Friday", new DateTimeZone(date_default_timezone_get()));
      $mnthName = $reportDate->format('F');

      // Set document properties
      $objPHPExcel->getProperties()
              ->setTitle("Operations Projections - Production Report - ".$mnthName." ".$year)
              ->setSubject($mnthName." Summary");

      // Assign Worksheet
      $sheet = $objPHPExcel->getActiveSheet();
      
      $this->buildWeeklySheet($sheet,$repData,$mnthName,$year,$month, $download);

      if($download){
        // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="Operations_Line_Items_'.gmdate('dMY').'.xls"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        
        // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
      } else {
        $xlFile = './Operations_Line_Items_'.gmdate('dMY').'.xls';

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save($xlFile);

        error_log("Excel report created successfully.");

        return $xlFile;
      }
  }
  
  /**
   * This function handles the formating and assignment of weekly reports in Excel 
   * 
   * @param type $sheet
   * @param type $repData
   */
  function buildWeeklySheet($sheet,$repData,$mnthName,$year,$month, $download)
  {
      $sheet->setTitle($mnthName." Data");
      // Setting Dimensions
      $sheet->getColumnDimension('A')->setWidth(1);
      $sheet->getColumnDimension('B')->setWidth(1);
      $sheet->getColumnDimension('C')->setWidth(20);
      $sheet->getColumnDimension('D')->setWidth(30);
      $sheet->getColumnDimension('E')->setWidth(35);
      $sheet->getColumnDimension('F')->setWidth(15);
      $sheet->getColumnDimension('G')->setWidth(20);
      $sheet->getColumnDimension('H')->setWidth(20);
      $sheet->getColumnDimension('I')->setWidth(20);
      $sheet->getColumnDimension('J')->setWidth(20);
      $sheet->getColumnDimension('K')->setWidth(20);

      $sheet->getColumnDimension('L')->setWidth(20);
      $sheet->getColumnDimension('M')->setWidth(20);
      $sheet->getColumnDimension('N')->setWidth(20);
      $sheet->getColumnDimension('O')->setWidth(20);
      $sheet->getColumnDimension('P')->setWidth(20);
      $sheet->getColumnDimension('Q')->setWidth(10);
      $sheet->getColumnDimension('R')->setWidth(5);

      $sheet->getColumnDimension('S')->setWidth(20);
      $sheet->getColumnDimension('T')->setWidth(20);
      $sheet->getColumnDimension('U')->setWidth(20);
      $sheet->getColumnDimension('V')->setWidth(20);
      $sheet->getColumnDimension('W')->setWidth(20);
      
      $sheet->getColumnDimension('X')->setWidth(20);
      $sheet->getColumnDimension('Y')->setWidth(20);
      $sheet->getColumnDimension('Z')->setWidth(20);
      $sheet->getColumnDimension('AA')->setWidth(20);
      $sheet->getColumnDimension('AB')->setWidth(20);
      
      $sheet->getStyle('A1:AB2')->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_CENTER );

      /**********************************************************************/
      // Labels
      /**********************************************************************/
      
      // Setup Title and Column Headers
      $sheet->mergeCells('D1:K1');
      $sheet->mergeCells('L1:Q1');
      $sheet->mergeCells('S1:W1');
      $sheet->mergeCells('X1:AB1');
      
      $sheet->setCellValueByColumnAndRow(2, 1, $mnthName." ".$year);
      $sheet->setCellValueByColumnAndRow(3, 1, "PROJECT REVENUES");
      $sheet->setCellValueByColumnAndRow(11, 1, "PROJECT EXPENSES");
      $sheet->setCellValueByColumnAndRow(19, 1, "AR ADJUSTMENTS");
      $sheet->setCellValueByColumnAndRow(24, 1, "AP ADJUSTMENTS");
     
      $sheet->setCellValueByColumnAndRow(2, 2, "Division Lead");

      $sheet->setCellValueByColumnAndRow( 3, 2, "Customer");
      $sheet->setCellValueByColumnAndRow( 4, 2, "Project");
      $sheet->setCellValueByColumnAndRow( 5, 2, "ProjectCode");

      $sheet->setCellValueByColumnAndRow( 6, 2, "Initial Rev Proj");
      $sheet->setCellValueByColumnAndRow( 7, 2, "Rev Proj Change");
      $sheet->setCellValueByColumnAndRow( 8, 2, "Final Rev Proj");
      $sheet->setCellValueByColumnAndRow( 9, 2, "Rev Actual");
      $sheet->setCellValueByColumnAndRow(10, 2, "Rev Variance");

      $sheet->setCellValueByColumnAndRow(11, 2, "Initial Exp Proj");
      $sheet->setCellValueByColumnAndRow(12, 2, "Exp Proj Change");
      $sheet->setCellValueByColumnAndRow(13, 2, "Final Exp Proj");
      $sheet->setCellValueByColumnAndRow(14, 2, "Exp Actual");
      $sheet->setCellValueByColumnAndRow(15, 2, "Exp Variance");

      $sheet->setCellValueByColumnAndRow(16, 2, "Margin" );
      
      $weekEndings = $this->getDates($year,$month);
      
      $sheet->setCellValueByColumnAndRow(18, 2, $weekEndings["Week1"]);
      $sheet->setCellValueByColumnAndRow(19, 2, $weekEndings["Week2"]);
      $sheet->setCellValueByColumnAndRow(20, 2, $weekEndings["Week3"]);
      $sheet->setCellValueByColumnAndRow(21, 2, $weekEndings["Week4"]);
      $sheet->setCellValueByColumnAndRow(22, 2, $weekEndings["Week5"]);
      
      $sheet->setCellValueByColumnAndRow(23, 2, $weekEndings["Week1"]);
      $sheet->setCellValueByColumnAndRow(24, 2, $weekEndings["Week2"]);
      $sheet->setCellValueByColumnAndRow(25, 2, $weekEndings["Week3"]);
      $sheet->setCellValueByColumnAndRow(26, 2, $weekEndings["Week4"]);
      $sheet->setCellValueByColumnAndRow(27, 2, $weekEndings["Week5"]);

      /**********************************************************************/
      // Summary rows   
      /**********************************************************************/
      $row = 3;
      foreach( $repData as $prj ) {
          $this->addWeeklyLine($sheet, $row, $prj);
          $row++;
      }
      $row++;

      /**********************************************************************/
      // Totals row   
      /**********************************************************************/
      
      if($download){
        $sheet->setCellValueByColumnAndRow(2, $row, "Totals");

        $sheet->setCellValueByColumnAndRow(6, $row, "=SUM(G3:G".($row-1).")");
        $sheet->setCellValueByColumnAndRow(7, $row, "=SUM(H3:H".($row-1).")");
        $sheet->setCellValueByColumnAndRow(8, $row, "=SUM(I3:I".($row-1).")");
        $sheet->setCellValueByColumnAndRow(9, $row, "=SUM(J3:J".($row-1).")");
        $sheet->setCellValueByColumnAndRow(10, $row, "=SUM(K3:K".($row-1).")");

        $sheet->setCellValueByColumnAndRow(11, $row, "=SUM(L3:L".($row-1).")");
        $sheet->setCellValueByColumnAndRow(12, $row, "=SUM(M3:M".($row-1).")");
        $sheet->setCellValueByColumnAndRow(13, $row, "=SUM(N3:N".($row-1).")");
        $sheet->setCellValueByColumnAndRow(14, $row, "=SUM(O3:O".($row-1).")");
        $sheet->setCellValueByColumnAndRow(15, $row, "=SUM(P3:P".($row-1).")");

        $sheet->setCellValueByColumnAndRow(16, $row, "=SUM(Q3:Q".($row-1).")");
      }
   
    /**************************************************************************/
     
//      $sheet->getStyle('A3:V'.$row)->getAlignment()->setHorizontal( \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT );

      // Setting number formats
      $sheet->getStyle('G3:P'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
      $sheet->getStyle('Q3:Q'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_PERCENTAGE);
      $sheet->getStyle('S3:AB'.$row)->getNumberFormat()->setFormatCode(\PHPExcel_Style_NumberFormat::FORMAT_CURRENCY_USD_SIMPLE);
      
      $sheet->getPageSetup()->setPrintArea('C1:AB'.$row);
      
      $sheet->getStyle('A1');
  }
  
  /**
   * 
   * @param type $sheet
   * @param type $lineNum
   * @param type $row
   */
  function addWeeklyLine($sheet,$lineNum,$row)
  { 
    $sheet->setCellValueByColumnAndRow( 2, $lineNum, trim(ucwords($row["divLead"])));

    $sheet->setCellValueByColumnAndRow( 3, $lineNum, $row["customer"]);
    $sheet->setCellValueByColumnAndRow( 4, $lineNum, $row["Project"]);
    $sheet->setCellValueByColumnAndRow( 5, $lineNum, $row["ProjectCode"]);
    
    $sheet->setCellValueByColumnAndRow( 6, $lineNum, $row["ARForecast"]);
    $sheet->setCellValueByColumnAndRow( 7, $lineNum, $row["RevChanges"]);
    $sheet->setCellValueByColumnAndRow( 8, $lineNum, $row["RevChanges"]);
    $sheet->setCellValueByColumnAndRow( 9, $lineNum, $row["ARActual"]);
    $sheet->setCellValueByColumnAndRow(10, $lineNum, $row["RevChanges"]-$row["ARActual"]);
    
    $sheet->setCellValueByColumnAndRow(11, $lineNum, $row["APForecast"]);
    $sheet->setCellValueByColumnAndRow(12, $lineNum, $row["ExpChanges"]);
    $sheet->setCellValueByColumnAndRow(13, $lineNum, $row["ExpChanges"]);
    $sheet->setCellValueByColumnAndRow(14, $lineNum, ($row["APActual"]*-1));
    $sheet->setCellValueByColumnAndRow(15, $lineNum, $row["ExpChanges"]-($row["APActual"]*-1));
    
    $sheet->setCellValueByColumnAndRow(16, $lineNum, ( $row["RevChanges"] == 0 ? 0 : (($row["RevChanges"] - $row["ExpChanges"])/$row["RevChanges"]) ) );
    
    $sheet->setCellValueByColumnAndRow(18, $lineNum, $row["ARWeek1"]);
    $sheet->setCellValueByColumnAndRow(19, $lineNum, $row["ARWeek2"]);
    $sheet->setCellValueByColumnAndRow(20, $lineNum, $row["ARWeek3"]);
    $sheet->setCellValueByColumnAndRow(21, $lineNum, $row["ARWeek4"]);
    $sheet->setCellValueByColumnAndRow(22, $lineNum, $row["ARWeek5"]);
    
    $sheet->setCellValueByColumnAndRow(23, $lineNum, $row["APWeek1"]);
    $sheet->setCellValueByColumnAndRow(24, $lineNum, $row["APWeek2"]);
    $sheet->setCellValueByColumnAndRow(25, $lineNum, $row["APWeek3"]);
    $sheet->setCellValueByColumnAndRow(26, $lineNum, $row["APWeek4"]);
    $sheet->setCellValueByColumnAndRow(27, $lineNum, $row["APWeek5"]);
  }
    
  /**
   * Returns the list of projections
   */
  private function getProjections($userId, $year, $month, $weekNum, $db, $msDb)
  {
    $sqlUser = "";
    if($userId != "ALL"){
      $sqlUser = "divId = ".$userId."
              AND";
    }
    
      $sql = "
          SELECT 
              id, 
              divId, 
              customer, 
              Project, 
              ProjectCode, 
              ARForecast,
              CASE WHEN ARWeek1 >= 999999999
                THEN 0
                ELSE ARWeek1 END AS ARWeek1,
              CASE WHEN ARWeek2 >= 999999999
                THEN 0
                ELSE ARWeek2 END AS ARWeek2,
              CASE WHEN ARWeek3 >= 999999999
                THEN 0
                ELSE ARWeek3 END AS ARWeek3,
              CASE WHEN ARWeek4 >= 999999999
                THEN 0
                ELSE ARWeek4 END AS ARWeek4,
              CASE WHEN ARWeek5 >= 999999999
                THEN 0
                ELSE ARWeek5 END AS ARWeek5,
              CASE WHEN ISNULL(weekInput) THEN ARForecast
                WHEN ARWeek".$weekNum." = 999999999 THEN 0
                ELSE ARWeek".$weekNum." END AS RevChanges,
              ARActual, 
              APForecast, 
              CASE WHEN APWeek1 >= 999999999
                THEN 0
                ELSE APWeek1 END AS APWeek1,
              CASE WHEN ARWeek2 >= 999999999
                THEN 0
                ELSE APWeek2 END AS APWeek2,
              CASE WHEN APWeek3 >= 999999999
                THEN 0
                ELSE APWeek3 END AS APWeek3,
              CASE WHEN APWeek4 >= 999999999
                THEN 0
                ELSE APWeek4 END AS APWeek4,
              CASE WHEN APWeek5 >= 999999999
                THEN 0
                ELSE APWeek5 END AS APWeek5,
              CASE WHEN ISNULL(weekInput) THEN APForecast
                WHEN APWeek".$weekNum." = 999999999 THEN 0
                ELSE APWeek".$weekNum." END AS  ExpChanges,
              APActual,
              weekInput
          FROM MonthlyProjectionsByWeek
          WHERE 
              ".$sqlUser." year = ".$year."
              AND month = ".$month."
          ORDER BY 
              divLead ASC,
              customer ASC,
              Project ASC
          ";
      
      $result =  $db->query($sql)->fetchAll(\PDO::FETCH_ASSOC);

      $dlRecords = array();

      $prjact = $this->getProjectsFromGL($year, $month, $msDb);

      foreach ($result as $row) {
          $tmpItems = [
              'lineId'     => $row['id'],
              'divId'      => $row['divId'],
              'divLead'    => $row['divLead'],
              'customer'   => $row['customer'],
              'Project'    => $row['Project'],
              'ProjectCode'=> $row['ProjectCode'],
              'ARForecast' => $row['ARForecast'],
              'ARWeek1'    => $row['ARWeek1'],
              'ARWeek2'    => $row['ARWeek2'],
              'ARWeek3'    => $row['ARWeek3'],
              'ARWeek4'    => $row['ARWeek4'],
              'ARWeek5'    => $row['ARWeek5'],
              'ARActual'   => $prjact[trim($row['ProjectCode'])]['AR'],
              'RevChanges' => $row['RevChanges'],
              'APForecast' => $row['APForecast'],
              'APWeek1'    => $row['APWeek1'],
              'APWeek2'    => $row['APWeek2'],
              'APWeek3'    => $row['APWeek3'],
              'APWeek4'    => $row['APWeek4'],
              'APWeek5'    => $row['APWeek5'],
              'APActual'   => $prjact[trim($row['ProjectCode'])]['AP'],
              'ExpChanges' => $row['ExpChanges'],
              'weekInput'  => $row['weekInput'],
              ];

          $dlRecords[] = $tmpItems;
      }

      return $dlRecords; 
  }

  /**
   * Query from GL for Actuals by Project
   * 
   * @return type
   */
  private function getProjectsFromGL($year, $month, $msDb)
  {   
      $sql = "
          SELECT 
              CUSTOMER,
              PROJECT,
              PROJECTCODE,
              ACCOUNTTYPE,
              SUM(AMOUNT) AS Amount
          FROM GLPL_Detail_Report
          WHERE year(TRXDATE) in (".$year.") 
              AND MONTH(TRXDATE) in (".$month.")
          GROUP BY
              PROJECTCODE,
              PROJECT,
              CUSTOMER,
              ACCOUNTTYPE
          ORDER BY
              CUSTOMER,
              PROJECT,
              ACCOUNTTYPE";

      $result = $msDb->query($sql);

      $fromGL = array();

      while ($row = $msDb->fetch_array($result)) {
          $fromGL[trim($row["PROJECTCODE"])][$row["ACCOUNTTYPE"]] = $row["Amount"];
      }

      return $fromGL; 
  }
      
  /**
     * 
     */
  private function getDates($year, $month) 
  {
      $periodArray = array();
      $timezone = 'America/Chicago';
      $format = 'n/d';

      if($month == 12) {
          $nextmonth = 1;
          $nextyear = $year + 1;
      }else{
          $nextmonth = $month + 1;
          $nextyear = $year;
      }

      // Add First of month
      $fom = new DateTime($month."/01/".$year, new DateTimeZone($timezone));
      $periodArray["Start"] = $fom->format($format);

      // Add Fridays
      $startDate = new DateTime($month."/01/".$year. " Friday", new DateTimeZone($timezone));
      $endDate = new DateTime($nextmonth."/01/".$nextyear, new DateTimeZone($timezone));
      $int = new DateInterval('P7D');
      $loopNum = 1;
      foreach(new DatePeriod($startDate, $int, $endDate) as $d) {
          $periodArray["Week".$loopNum] = $d->format($format);
          $loopNum++;
      }

      // Add End of month
      $dim = cal_days_in_month(CAL_GREGORIAN, $month, $year);
      $eom = new DateTime($month."/".$dim."/".$year, new DateTimeZone($timezone));
      $periodArray["End"] = $eom->format($format);

      return $periodArray;
  }
  
  /*****************************************************************************/
}
