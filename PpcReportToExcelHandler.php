<?php
namespace ProspectRelationshipManager\AppBundle\FormatConverter;

use PHPExcel;
use PHPExcel_Worksheet;
use PHPExcel_Style_Alignment;
use JWT\ProspectSdk\Model\GroupsAbstract;

/**
 * Class PpcReportToExcelHandler handler
 *
 * @author Scott Richardson ~~~ <scott.richardson@jwt.com>
 */
  class PpcReportToExcelHandler
  {
    /**
     * @var PHPExcel $workbook
     */
    protected $workbook;

    /**
     * @var int $sheetNum
     */
    protected $sheetNum;

    /**
     * @var int $rowNum
     */
    protected $rowNum;

    /**
     * @var PHPExcel_Worksheet $sheet
     */
    protected $sheet;

    /**
     * Method handle
     *
     * @param PpcReportToExcel $command
     */
    public function handle(PpcReportToExcel $command)
    {
        return $this->createWorkbook($command);
    }

    /**
     * Method createWorkbook
     *
     * Creates the excel Workbook.
     *
     * @param \Traversable|array $command
     * @return \ExcelAnt\Adapter\PhpExcel\Workbook\Workbook
     * @author Scott Richardson ~~~ <scott.richardson@jwt.com>
     */
    private function createWorkbook($command)
    {
      $params = [];
      $params['reportData'] = $command->reportData;
      $params['reportTitle'] = "ePPC Monthly Report - ";
      $params['reportDate'] = " - ".$command->year."-".$command->month;
      $params['header'] = array(
          'NAME',
          'DUE',
          'RET on TIME',
          'PDRL',
          'TOT RET',
          '% RET',
          'GL',
          '% GL',
          'WKBL',
          '% WKBL',
          'REF TO EPPC',
          'REF FROM EPPC',
          'PAST DUE',
          'TOTAL APP',
          'OQL APP',
          'WKBL:APP'
          );

      $this->workbook = new PHPExcel();
      $this->workbook->getDefaultStyle()->getFont()->setName('Arial');
      $this->workbook->getDefaultStyle()->getFont()->setSize(10);

      if ((string) $command->userContext->getGroup() === GroupsAbstract::DISTRICT) {
        $this->buildDistrictReport($params);
      } else {
        $this->buildStationReport($params);
      }

      return $this->workbook;
    }

    /**
     * @param type $params
     */
    private function buildDistrictReport($params)
    {
      $reportData = $params['reportData'];
      $this->sheetNum = 0;
      /**********************
       * Start Loop here - Each Station is a new sheet
       **********************/
      foreach ($reportData->getMonthlyStations() as $report) {
        $this->initSheet($report->stationName, $params);
        //Creating Rows
        foreach ($reportData->getMonthlySubStationsForStation($report->stationId) as $subData) {
          if($subData->subStationName != ""){
            $row = $this->createRow($subData);
            $this->rowNum++;
            $this->sheet->fromArray($row, null, 'A'.$this->rowNum, true);
            if($subData->active === 0){
              $styleArray = array(
                      'font'  => array(
                          'bold'  => true,
                          'color' => array('rgb' => 'FF0000'),
                      ));
              $this->sheet->getStyle('A'.$this->rowNum.':P'.$this->rowNum)->applyFromArray($styleArray);
            }
          }
        }
        // Station Monthly TOTALS
        $row = $this->createRow($report, true);
        $this->rowNum++;
        $this->sheet->fromArray($row, null, 'A'.$this->rowNum, true);
        $this->sheet->getStyle('A'.$this->rowNum)->getFont()->setBold(true);

        $this->middleSheet($params);
        foreach ($reportData->getYearToDateSubStationsForStation($report->stationId) as $subData) {
            if($subData->subStationName != ""){
              $row = $this->createRow($subData);
              $this->rowNum++;
              $this->sheet->fromArray($row, null, 'A'.$this->rowNum, true);
              if($subData->active === 0){
                $styleArray = array(
                      'font'  => array(
                          'bold'  => true,
                          'color' => array('rgb' => 'FF0000'),
                      ));
                $this->sheet->getStyle('A'.$this->rowNum.':P'.$this->rowNum)->applyFromArray($styleArray);
              }
            }
        }
        
        // Yearly Report TOTALS
        foreach ($reportData->getYearToDateStations() as $yearly) {
          if($report->stationName == $yearly->stationName){
            $row = $this->createRow($yearly, true);
            $this->rowNum++;
            $this->sheet->fromArray($row, null, 'A'.$this->rowNum, true);
            $this->sheet->getStyle('A'.$this->rowNum)->getFont()->setBold(true);
          }
        }

        $this->closeSheet();
      }
    }

    /**
     * 
     * @param type $this->workbook
     * @param type $params
     */
    private function buildStationReport($params)
    {
      $reportData = $params['reportData'];
      $this->sheetNum = 0;
      foreach ($reportData->getMonthlySubStations() as $subData) {
        if( $this->sheetNum == 0 ){
          $this->initSheet($subData->stationName, $params);
        }

        //Creating Rows
        if($subData->subStationName != ""){
          $row = $this->createRow($subData);
          $this->rowNum++;
          $this->sheet->fromArray($row, null, 'A'.$this->rowNum, true);
          if($subData->active === 0){
            $styleArray = array(
                  'font'  => array(
                      'bold'  => true,
                      'color' => array('rgb' => 'FF0000'),
                  ));
            $this->sheet->getStyle('A'.$this->rowNum.':P'.$this->rowNum)->applyFromArray($styleArray);
          }
        }
      }
      foreach ($reportData->getMonthlyStations() as $report) {
        // Station Monthly TOTALS
        $row = $this->createRow($report, true);
        $this->rowNum++;
        $this->sheet->fromArray($row, null, 'A'.$this->rowNum, true);
        $this->sheet->getStyle('A'.$this->rowNum)->getFont()->setBold(true);
      }

      $this->middleSheet($params);
      foreach ($reportData->getYearToDateSubStations() as $subData) {
        if($subData->subStationName != ""){
          $row = $this->createRow($subData);
          $this->rowNum++;
          $this->sheet->fromArray($row, null, 'A'.$this->rowNum, true);
          if($subData->active === 0){
            $styleArray = array(
                  'font'  => array(
                      'bold'  => true,
                      'color' => array('rgb' => 'FF0000'),
                  ));
            $this->sheet->getStyle('A'.$this->rowNum.':P'.$this->rowNum)->applyFromArray($styleArray);
          }
        }
      }
      foreach ($reportData->getYearToDateStations() as $report) {
        // Station Monthly TOTALS
        $row = $this->createRow($report, true);
        $this->rowNum++;
        $this->sheet->fromArray($row, null, 'A'.$this->rowNum, true);
        $this->sheet->getStyle('A'.$this->rowNum)->getFont()->setBold(true);
      }

      $this->closeSheet();
    }

    /**
     *
     * @param type $data
     * @return type
     */
    private function createRow($data, $totalRow = false)
    {
      if($totalRow){
        $name = $data->stationName;
      } elseif($data->active == 0) {
        $name = $data->subStationName . " ( INACTIVE )";
      } else {
        $name = $data->subStationName;
      }

      $row = array(
        $name,
        $data->due,
        $data->returnedOnTime,
        $data->pastDueReturned,
        $data->returned,
        $data->returnedPercentage."%",
        $data->goodLead,
        $data->goodLeadPercentage."%",
        $data->workable,
        $data->workablePercentage."%",
        $data->transferOut,
        $data->transferIn,
        $data->pastDue,
        $data->contract,
        $data->ppcContract,
        $data->workablePpcContractsRatio,
      );

      return $row;
    }

    /**
     *
     * @param type $reportName
     * @param type $params
     */
    private function initSheet($reportName, $params)
    {
      $tabTitle = "RS ".$reportName;
      $sheetTitle = $params['reportTitle'] . $tabTitle . $params['reportDate'];
      $this->sheet = new PHPExcel_Worksheet($this->workbook, $tabTitle);
      $this->workbook->addSheet($this->sheet, $this->sheetNum);
      $this->sheetNum++;

      //Creating Headers
      $this->sheet->fromArray(array($sheetTitle), null, 'A1', true);     // Report Title
      //Format Sheet Title
      $this->sheet->mergeCells('A1:P1');
      $this->sheet->getStyle('A1:P1')->getFont()->setBold(true);
      $this->sheet->getStyle('A1:P1')->getFont()->setSize(14);
      $this->sheet->getStyle('A1:P1')->getAlignment()->setHorizontal('center');

      //Blank section
      $this->sheet->mergeCells('A2:P2');

      //Format Section Title
      $this->sheet->fromArray(array("Current Month"), null, 'A3', true); // Section Title
      $this->sheet->mergeCells('A3:P3');
      $this->sheet->getStyle('A3:P3')->getFont()->setBold(true);
      $this->sheet->getStyle('A3:P3')->getFont()->setSize(14);
      $this->sheet->getStyle('A3:P3')->getAlignment()->setHorizontal('center');

      //Format column headers
      $this->sheet->fromArray($params['header'], null, 'A4', true);
      $this->sheet->getStyle('A4:P4')->getFont()->setBold(true);

      $this->rowNum = 4;
    }

    /**
     * 
     * @param type $params
     */
    private function middleSheet($params)
    {
      // Yearly Report Section
      $this->rowNum++;  // Blank Line
      $this->sheet->mergeCells('A'.$this->rowNum.':P'.$this->rowNum, true); // Blank Line
      $this->rowNum++;
      $this->sheet->fromArray(array("Fiscal Year"), null, 'A'.$this->rowNum, true);  // Section Title
      //Format Section Title
      $this->sheet->mergeCells('A'.$this->rowNum.':P'.$this->rowNum);
      $this->sheet->getStyle('A'.$this->rowNum.':P'.$this->rowNum)->getFont()->setBold(true);
      $this->sheet->getStyle('A'.$this->rowNum.':P'.$this->rowNum)->getFont()->setSize(14);
      $this->sheet->getStyle('A'.$this->rowNum.':P'.$this->rowNum)->getAlignment()->setHorizontal('center');

      $this->rowNum++;
      $this->sheet->fromArray($params['header'], null, 'A'.$this->rowNum, true);  // Headers
      //Format column headers
      $this->sheet->getStyle('A'.$this->rowNum.':P'.$this->rowNum)->getFont()->setBold(true);
    }

    /**
     *
     */
    private function closeSheet()
    {
      $this->sheet->getColumnDimension('A')->setAutoSize(true);
      $this->sheet->getColumnDimension('C')->setAutoSize(true);
      $this->sheet->getColumnDimension('K')->setAutoSize(true);
      $this->sheet->getColumnDimension('L')->setAutoSize(true);
      $this->sheet->getColumnDimension('N')->setAutoSize(true);
      $this->sheet->getColumnDimension('P')->setAutoSize(true);
    }
 }
