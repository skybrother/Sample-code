<?php
/******************************************************************************* 
 * This script will schedule checks and notifications for projection information
 * 
 * Sent messages will be kept in a history table
 * and will need to think about table clean-up later.
 *
 * @created Oct 02, 2017
 * @author Scott Richardson {*.*} <scott.richardson@powerhouseretail.com> 
 *******************************************************************************/

// =======================================================
// ==  Please configure crontab as follows:
// ==   */30 * * * * /usr/bin/php /path/to/this_file.php -t *type* > /path/to/log/LOG.txt 2> /dev/null
// ==  ( this will run every 30 minutes )
// ======================================================= 

namespace ProjectionSummary;

require_once __DIR__ . '/../vendor/autoload.php';
require_once __DIR__ . '/../config/constants.inc.php';

use Cis\Config;
use Cis\Db;
use Cis\MSSQLDb;
use Cis\ProjectionsReport;

date_default_timezone_set('America/Chicago');

/*******************************************************************************/
echo "================================================= \r\n";
echo "STARTING Projection Weekly Summary @ ".date("m-d-Y H:i:s")."\r\n";
echo "================================================= \r\n";
/*******************************************************************************/
echo php_uname('s')."\r\n";
/*******************************************************************************
// Configuration Section
/*******************************************************************************/
$cfg     = Config::getInstance();
$month   = date('n');  // Current Month
$year    = date('Y');  // Current Year
$db      = Db::getConn();
$msDb    = new MSSQLDb();

$pr      = new ProjectionsReport();

$addEmails = array(
    /* Adam Marbut */       strtolower("Adam.Marbut@powerhouseretail.com"), 
    /* Amber Alvarez */     strtolower("Amber.Alvarez@powerhouseretail.com"), 
    // /* Brent Teeter */      strtolower("Brent.Teeter@powerhouseretail.com"), 
    /* Charlie Setchell */  strtolower("Charlie.Setchell@powerhouseretail.com"), 
    /* David Frederick */   strtolower("David.Frederick@powerhouseretail.com"), 
    // /* David Hargrave */    strtolower("David.Hargrave@powerhouseretail.com"), 
    /* Frank Daniels */     strtolower("Frank.Daniels@powerhouseretail.com"), 
    /* Garry Smith */       strtolower("Garry.Smith@powerhouseretail.com"), 
    /* Gary Chatha */       strtolower("Gary.Chatha@powerhouseretail.com"), 
    /* Jack Lewis */        strtolower("Jack.Lewis@powerhouseretail.com"), 
    /* Max Draper */        strtolower("Max.Draper@powerhouseretail.com"), 
    /* Megan Varano */      strtolower("Megan.Varano@powerhouseretail.com"), 
    // /* Michael Wroughton */ strtolower("Michael.Wroughton@powerhouseretail.com"), 
    // /* Mike Murphy */       strtolower("Mike.Murphy@powerhouseretail.com"), 
    // /* Robert Blake-Ward */ strtolower("Robert.Blake-Ward@powerhouseretail.com"), 
    /* Ryan Jackson */      strtolower("Ryan.Jackson@powerhouseretail.com"),
    /* Cortney Hernandez */ strtolower("Cortney.Hernandez@powerhouseretail.com"),
    /* Shelby Forster */    strtolower("shelby.forster@powerhouseretail.com"), 
    /* Lindsy Jaetzold */   strtolower("Lindsy.Jaetzold@powerhouseretail.com"),
    /* Michele Jaso */      strtolower("Michele.Jaso@powerhouseretail.com"),
    /* Scott Richardson */  strtolower("Scott.Richardson@powerhouseretail.com"), 
    /* Sue McCarty */       strtolower("Sue.McCarty@powerhouseretail.com")
);

//Get Week Number -- Will be the PREVIOUS WEEK
$dim = cal_days_in_month(CAL_GREGORIAN, $month-1, $year);
$weekNum = floor(date('d')/7);
if($weekNum < 1){ $month--; }

// Pull Data
$repData = $pr->pullReportData($year,$month,$db,$msDb,$weekNum,TRUE);
// Generate Report
$xlFile = $pr->generateExcel($repData,$year,$db,$msDb);
echo "===================================== \r\n";
$imageFile = $pr->generateImage($xlFile);

// Pull Address List
$addys = $pr->pullAddresses($db,$addEmails);

// Build and Send email
$pr->buildEmail($xlFile,$imageFile,$addys);

echo "===================================== \r\n";
echo "Finished processing. \r\n";
echo "===================================== \r\n";
echo "\r\n";

/*******************************************************************************
 *  END OF LINE 
 *******************************************************************************/

