<?php
/**
 * Weekly Summary Pages
 *
 * @created September 13, 2017
 * @author Scott Richardson {*.*} <scott.richardson@powerhouseretail.com> 
 */
namespace Controllers\Projections;

use Cis\BaseController;
use Cis\TwigHelper;
use Cis\Db;
use Cis\MSSQLDb;
use Cis\Model\User;
use Cis\ProjectionsReport;
/* Include Date Classes  */
use DateTime;
use DatePeriod;
use DateTimeZone;
use DateInterval;

class WeeklySummary extends BaseController 
{
    /**
     * MySQL Db connection
     *  
     * @var type 
     */
    private $Db;
    
    /**
     * Main report display function 
     */
    public function get() 
    {
        $this->post();
    }
    
    /**
     * Main report display function 
     */
    public function post() 
    {
        if(isset($_REQUEST['month'])) {
            $month = $_REQUEST['month'];  // Month from form
        } else {
            $month = date('n');  // Current Month
        }
        
        if(isset($_REQUEST['year'])) {
            $year = $_REQUEST['year'];  // Month from form
        } else {
            $year = date('Y');  // Current Month
        }
        
        $this->Db = Db::getConn();
        $this->msDb = new MSSQLDb();
        $twigHelper = new TwigHelper();
        
        $pr = new ProjectionsReport();
        
        // User info
        $user = User::getCurrent();
        
        $count = 0;
        
        // Setup the initial template variables.
        $templateVars = $twigHelper->getDefaultTemplateVariables();
        $templateVars["info"] = $this->getInfo($year, $month, $user); 
                
        if(!is_null($_REQUEST['DivId'])){   
            $divId = $_REQUEST['DivId'];
            $week = $_REQUEST['weekNum'];
        
            foreach ($_REQUEST['LineIds'] as $cust) {
                $idVal = $cust['value'];
                $arVal = ($_REQUEST['RevChange'][$count]['value'] == 0) ? 999999999 :$_REQUEST['RevChange'][$count]['value'];
                $apVal = ($_REQUEST['ExpChange'][$count]['value'] == 0) ? 999999999 :$_REQUEST['ExpChange'][$count]['value'];
                $this->updateProjections($divId, $week, $idVal, $arVal, $apVal);
                $count++;
            }       
        }
        
        if(!is_null($_REQUEST['download'])){
          
          if(isset($_REQUEST['type'])){
            // WEEKLY REPORT EXCEL BY USER
            // Pull Data
            $repData = $this->getProjections($user->id, $year, $month, $templateVars["info"]["weekNum"]); 
            // Generate Report
            $pr->generateWeeklyReport($repData,$year,$month,$this->Db,TRUE);
          } elseif(isset($_REQUEST['batched'])) {
            // WEEKLY REPORT W/ BATCH INFO EXCEL BY USER
            // Pull Data
            $repData = $pr->pullOperationsData($year,$month,$this->Db,$this->msDb,$templateVars["info"]["weekNum"]);
            // Generate Report
            $pr->generateOperationsReport($repData,$year,$month,$this->Db,$this->msDb,TRUE);
          } else {
            // MONTHLY REPORT EXCEL
            // Pull Data
            $repData = $pr->pullReportData($year,$month,$this->Db,$this->msDb,$templateVars["info"]["weekNum"],TRUE);
            // Generate Report
            $pr->generateExcel($repData,$year,$this->Db,$this->msDb,TRUE);
          }
        }
        
        // Setup the remaining template variables.
        $templateVars["formValues"] = $_REQUEST; 
        $templateVars["dates"] = $this->getDates($year, $month); 
        $templateVars["actuals"] = $this->getFromGL($year,$month);

        if(isset($_REQUEST['page'])) {
            $templateVars["projections"] = $pr->getSummaries($year, $month, $templateVars["actuals"], $this->Db, $templateVars["info"]["weekNum"]); 
            echo $twigHelper->getTwig()->render('projections/new_projections_summary.twig', $templateVars);
        } elseif(isset($_REQUEST['print'])) {
            $templateVars["projections"] = $this->getProjections($user->id, $year, $month, $templateVars["info"]["weekNum"]); 
            $templateVars["noForecast"] = $this->getUnforecastedProjects($user->id, $year, $month);
            echo $twigHelper->getTwig()->render('projections/new_projections_print.twig', $templateVars);
        } else {
            $templateVars["projections"] = $this->getProjections($user->id, $year, $month, $templateVars["info"]["weekNum"]); 
            $templateVars["noForecast"] = $this->getUnforecastedProjects($user->id, $year, $month);
            $templateVars["divActuals"] = $this->getProjectsFromGL($year, $month, $user->username);
            echo $twigHelper->getTwig()->render('projections/new_projections.twig', $templateVars);
        }
    }
  
    /**
     * Returns the list of projections
     */
    private function getProjections($userId, $year, $month, $weekNum)
    {
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
                divId = ".$userId."
                AND year = ".$year."
                AND month = ".$month."
            ORDER BY 
                customer ASC
            ";
        
        $result = $this->Db->query($sql)->fetchAll(\PDO::FETCH_ASSOC);
        
        $dlRecords = array();
        
        $prjact = $this->getProjectsFromGL($year, $month);
        
        foreach ($result as $row) {
            $tmpItems = [
                'lineId'     => $row['id'],
                'divId'      => $row['divId'],
                'divLead'    => $row['divLead'],
                'customer'   => $row['customer'],
                'Project'    => $row['Project'],
                'ProjectCode'=> $row['ProjectCode'],
                'ARForecast' => (float)$row['ARForecast'],
                'ARWeek1'    => (float)$row['ARWeek1'],
                'ARWeek2'    => (float)$row['ARWeek2'],
                'ARWeek3'    => (float)$row['ARWeek3'],
                'ARWeek4'    => (float)$row['ARWeek4'],
                'ARWeek5'    => (float)$row['ARWeek5'],
                'ARActual'   => (float)$prjact[trim($row['ProjectCode'])]['AR'],
                'RevChanges' => (float)$row['RevChanges'],
                'APForecast' => (float)$row['APForecast'],
                'APWeek1'    => (float)$row['APWeek1'],
                'APWeek2'    => (float)$row['APWeek2'],
                'APWeek3'    => (float)$row['APWeek3'],
                'APWeek4'    => (float)$row['APWeek4'],
                'APWeek5'    => (float)$row['APWeek5'],
                'APActual'   => (float)$prjact[trim($row['ProjectCode'])]['AP'],
                'ExpChanges' => (float)$row['ExpChanges'],
                'weekInput'  => (int)$row['weekInput'],
                ];
            
            $dlRecords[] = $tmpItems;
        }
        
        return $dlRecords; 
    }
  
    /**
     * 
     * @return string
     */
    private function getInfo($year, $month, $user) 
    {
        $timezone = 'America/Chicago';
        $dayAdjust = 0;
        switch(date('N')){
            case 7:
                $dayAdjust++;
            case 6:
                $dayAdjust++;
            default:
                $tmpDay = date('d')-$dayAdjust;
                if($tmpDay < 1){
                    $month--;  
                }
                if($month<1){
                    $year--;
                }
                $tmpDate = new DateTime($month."/".$tmpDay."/".$year. " Friday", new DateTimeZone($timezone));  
        }
        $tmpMnth = new DateTime($month."/01/".$year, new DateTimeZone($timezone));
        
        $dim = cal_days_in_month(CAL_GREGORIAN, $month, $year);
        $dom = date('d');
        
        if($year<date('Y')){
            $pom = 100;
        }elseif($year>date('Y')){
            $pom = 0;
        }else{
            if($month<date('n')) {
                $pom = 100;
            }elseif($month>date('n')){
                $pom = 0;
            }else{
                $pom = floor(($dom/$dim)*100);
            }
        }
                
        $infoVals = [
            'Month'  => $tmpMnth->format('F'),
            'mthNum' => $month,
            'curMth' => date('n'),
            'year'   => $year,
            'curYear'=> date('Y'),
            'pom'    => $pom,
            'weekEnd'=> $tmpDate->format('n/d/Y'),
            'weekNum'=> (floor($dom/7)+1),
            'User'   => ucwords(str_replace(".", " ", $user->username)),
            'divId'  => $user->id,
            'dom'    => $dom,
        ];
        return $infoVals;
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
    
    /**
     * Returns the Summaries of projections
     */
    private function getSummaries($year, $month, $actualsArray)
    {
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
                SUM(COALESCE(NULLIF(mpbw.ARWeek5,0),NULLIF(mpbw.ARWeek4,0),NULLIF(mpbw.ARWeek3,0),NULLIF(mpbw.ARWeek2,0),NULLIF(mpbw.ARWeek1,0),mpbw.ARForecast)) as RevChanges, 
                SUM(mpbw.APForecast) as APForecast, 
                SUM(mpbw.APWeek1) as APWeek1, 
                SUM(mpbw.APWeek2) as APWeek2, 
                SUM(mpbw.APWeek3) as APWeek3, 
                SUM(mpbw.APWeek4) as APWeek4, 
                SUM(mpbw.APWeek5) as APWeek5,
                SUM(COALESCE(NULLIF(mpbw.APWeek5,0),NULLIF(mpbw.APWeek4,0),NULLIF(mpbw.APWeek3,0),NULLIF(mpbw.APWeek2,0),NULLIF(mpbw.APWeek1,0),mpbw.APForecast)) as ExpChanges 
            FROM MonthlyProjectionsByWeek mpbw
            LEFT JOIN Users u ON (mpbw.divId = u.id)
            WHERE year = ".$year."
                AND month = ".$month."
            GROUP BY
                    mpbw.divId,
                    u.username
            ORDER BY 
               u.username ASC; 
            ";

        $result = $this->Db->query($sql)->fetchAll(\PDO::FETCH_ASSOC);
        
        $dlRecords = array();
        
        foreach ($result as $row) {
            $tmpItems = [
                'lineId'     => $row['id'],
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
     * Query from GL for Actuals
     * 
     * @return type
     */
    private function getFromGL($year, $month)
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
        
        $result = $this->msDb->query($sql);
        
        error_log($sql);
        
        $fromGL = array();
        
        while ($row = $this->msDb->fetch_array($result)) {
            $fromGL[trim($row["DivisionMgr"])][$row["ACCOUNTTYPE"]] = $row["Amount"];
        }

        return $fromGL; 
    }
    
    /**
     * Query from GL for Actuals by Project
     * 
     * @return type
     */
    private function getProjectsFromGL($year, $month, $divLead = NULL)
    {   
      if(is_null($divLead)){
        $divSql = $divLead;
      } else {
        $divSql = "AND DivisionMgr LIKE '".ucwords(str_replace('.',' ',$divLead))."%'";
      }
      
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
              ".$divSql."  
          GROUP BY
              PROJECTCODE,
              PROJECT,
              CUSTOMER,
              ACCOUNTTYPE
          ORDER BY
              CUSTOMER,
              PROJECT,
              ACCOUNTTYPE";

        $result = $this->msDb->query($sql);
        
        error_log($sql);
        
        $fromGL = array();
        
        if(is_null($divLead)){
          while ($row = $this->msDb->fetch_array($result)) {
              $fromGL[trim($row["PROJECTCODE"])][$row["ACCOUNTTYPE"]] = $row["Amount"];
          }
        } else {
          while ($row = $this->msDb->fetch_array($result)) {
            $tmpItems = [
                'Customer'   => $row['CUSTOMER'],
                'Project'    => $row['PROJECT'],
                'ProjectCode'=> $row['PROJECTCODE'],
                'AccountType'=> $row['ACCOUNTTYPE'],
                'Amount'     => $row['Amount'],
                ];
            
            $fromGL[] = $tmpItems;
          }        
        }
        return $fromGL; 
    }
    
    /**
     * Returns the list of projections
     */
    private function updateProjections($divId, $week, $idVal, $arVal, $apVal)
    {
        $sql = "
            UPDATE MonthlyProjectionsByWeek 
            SET 
                ARWeek".$week." = ".$arVal.",
                APWeek".$week." = ".$apVal.",
                weekInput = ".$week."
            WHERE
                id = ".$idVal." 
                AND divId = ".$divId."
            ";
        
        error_log($sql);

        return $this->Db->exec($sql);
    }
  
    /**
     * 
     * @param type $needle
     * @param type $haystack
     * @param type $strict
     * @return boolean
     */
    private function inArrayR($needle, $haystack, $strict = false) 
    {
        foreach ($haystack as $item) {
            if (($strict ? $item === $needle : $item == $needle) || (is_array($item) && $this->inArrayR($needle, $item, $strict))) {
                return true;
            }
        }

        return false;
    }
    
    /**
     * Returns a list of projects with no forecasts against them for the specified 
     * time frame
     * 
     * @param type $userId
     * @param type $year
     * @param type $month
     * @return type
     */
    private function getUnforecastedProjects($userId, $year, $month)
    {
        $sql = "
            SELECT 
                Customer, 
                Project, 
                ProjectCode, 
                active
            FROM ActiveProjectsByDL
            WHERE 
                ProjectCode NOT IN(
                    SELECT 
                        ProjectCode 
                    FROM MonthlyProjectionsByWeek 
                    WHERE 
                        divId = ".$userId."
                        AND year = ".$year."
                        AND month = ".$month."
                )
                AND DivisionLeadId = ".$userId."
            ORDER BY 
                Customer ASC
            ";

        $result = $this->Db->query($sql)->fetchAll(\PDO::FETCH_ASSOC);
        
        $dlRecords = array();
        
        foreach ($result as $row) {
            $tmpItems = [
                'divId'      => $row['DivisionLeadId'],
                'customer'   => $row['Customer'],
                'Project'    => $row['Project'],
                'ProjectCode'=> $row['ProjectCode'],
                'active'     => $row['active'],
                'combined'   => $row['Customer'].'|'.$row['Project'].'|'.$row['ProjectCode'],
                ];
            
            $dlRecords[] = $tmpItems;
        }
        
        return $dlRecords; 
    }  
}
