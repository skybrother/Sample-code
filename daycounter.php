<?php
/******************************************************************************
 * Simple Day Counter I wanted to make
 * 
 * @Created July 14, 2017
 * @Author - Scott Richardson
 ******************************************************************************/

// Babies
$jonathan = 38*7;  // Jonathan was 38 weeks
$jbBirth = new \DateTimeImmutable("2015-07-21");
$hannah = (36*7)+6; // Hannah was 36 weeks 6 days
$hlBirth = new \DateTimeImmutable("2016-07-02");
$caleb = 38*7;  // Caleb was 38 weeks
$cmBirth = new \DateTimeImmutable("2018-02-19");
//$kcstart = new \DateTime("2017-06-01");

// Met Jeri
$meetdate = new \DateTimeImmutable("2014-03-25");
// Our First Date
$firstdate = new \DateTimeImmutable("2014-04-21");
// Married Date
$marrieddate = new \DateTimeImmutable("2014-09-26");
// Today's Date
$todaysdate = new \DateTimeImmutable();

//Days I've known Jeri
$daysKnown = $todaysdate->diff($meetdate, true);
//Days since first Date
$sinceFirst = $todaysdate->diff($firstdate, true);
//Days we've been married
$daysMarried = $todaysdate->diff($marrieddate, true);
//Days she's been pregnant
$daysPrego = $jonathan + $hannah + $caleb; //($todaysdate->diff($kcstart, true)->days);
//Days she hasn't been with child
$daysNotPrego = ($daysKnown->days) - $daysPrego;

$jbConceive = $jbBirth->modify("-".$jonathan." days");
$hlConceive = $hlBirth->modify("-".$hannah." days");
$cmConceive = $cmBirth->modify("-".$caleb." days");

// Pregnant Days
$pregDays[] = [ (string) $meetdate->getTimestamp()*1000 , "0"]; // Met Jeri
$pregDays[] = [ (string) $firstdate->getTimestamp()*1000, "0"]; // First Date
$pregDays[] = [ (string) $marrieddate->getTimestamp()*1000, "0"]; // Got Married!!!!
$pregDays[] = [ (string) $jbConceive->getTimestamp()*1000, "1"]; // Expecting Jonathan!!!!
$pregDays[] = [ (string) $jbBirth->getTimestamp()*1000, (string) $jonathan ]; // Jonathan Born!!!!
$pregDays[] = [ (string) $hlConceive->getTimestamp()*1000, (string) $jonathan ]; // Expecting Hannah!!!!
$pregDays[] = [ (string) $hlBirth->getTimestamp()*1000, (string) ($jonathan+$hannah) ]; // Hannah Born!!!!
$pregDays[] = [ (string) $cmConceive->getTimestamp()*1000, (string) ($jonathan+$hannah) ]; // Expecting Caleb!!!!
$pregDays[] = [ (string) $cmBirth->getTimestamp()*1000, (string) $daysPrego]; // Caleb Born!!!!
$pregDays[] = [ (string) $todaysdate->getTimestamp()*1000, (string) $daysPrego]; // Today

// Not pregnant days from first meeting
$notPregDays[] = [ (string) $meetdate->getTimestamp()*1000, "1"]; // Met Jeri
$notPregDays[] = [ (string) $firstdate->getTimestamp()*1000, (string) $firstdate->diff($meetdate, true)->days ]; // First Date
$notPregDays[] = [ (string) $marrieddate->getTimestamp()*1000, (string) $marrieddate->diff($meetdate, true)->days]; // Got Married!!!!
$notPregDays[] = [ 
  (string) $jbConceive->getTimestamp()*1000, 
  (string) $jbConceive->diff($meetdate, true)->days 
  ]; // Expecting Jonathan!!!!
$notPregDays[] = [ 
  (string) $jbBirth->getTimestamp()*1000, 
  (string) ($jbBirth->diff($meetdate, true)->days - $jonathan )
  ]; // Jonathan Born!!!!
$notPregDays[] = [ 
  (string) $hlConceive->getTimestamp()*1000, 
  (string) ($hlConceive->diff($meetdate, true)->days - $jonathan )
  ]; // Expecting Hannah!!!!
$notPregDays[] = [ 
  (string) $hlBirth->getTimestamp()*1000, 
  (string) ( $hlBirth->diff($meetdate, true)->days - ($jonathan+$hannah) )
  ]; // Hannah Born!!!!
$notPregDays[] = [ 
  (string) $cmConceive->getTimestamp()*1000, 
  (string) ( $cmConceive->diff($meetdate, true)->days - ($jonathan+$hannah) )
  ]; // Expecting Caleb!!!!
$notPregDays[] = [ 
  (string) $cmBirth->getTimestamp()*1000, 
  (string) ( $cmBirth->diff($meetdate, true)->days - $daysPrego )
  ]; // Caleb Born!!!!
$notPregDays[] = [ 
  (string) $todaysdate->getTimestamp()*1000, 
  (string) ( $todaysdate->diff($meetdate, true)->days - $daysPrego )
  ]; // Today

// Not pregnant days from first DATE
$notPregFirstDate[] = [ (string) $meetdate->getTimestamp()*1000, "0" ]; // Met Jeri
$notPregFirstDate[] = [ (string) $firstdate->getTimestamp()*1000, "1" ]; // First Date
$notPregFirstDate[] = [ (string) $marrieddate->getTimestamp()*1000, (string) $marrieddate->diff($firstdate, true)->days]; // Got Married!!!!
$notPregFirstDate[] = [ 
  (string) $jbConceive->getTimestamp()*1000, 
  (string) $jbConceive->diff($firstdate, true)->days 
  ]; // Expecting Jonathan!!!!
$notPregFirstDate[] = [ 
  (string) $jbBirth->getTimestamp()*1000, 
  (string) ( $jbBirth->diff($firstdate, true)->days - $jonathan )
  ]; // Jonathan Born!!!!
$notPregFirstDate[] = [ 
  (string) $hlConceive->getTimestamp()*1000, 
  (string) ( $hlConceive->diff($firstdate, true)->days - $jonathan )
  ]; // Expecting Hannah!!!!
$notPregFirstDate[] = [ 
  (string) $hlBirth->getTimestamp()*1000, 
  (string) ( $hlBirth->diff($firstdate, true)->days - ($jonathan+$hannah) )
  ]; // Hannah Born!!!!
$notPregFirstDate[] = [ 
  (string) $cmConceive->getTimestamp()*1000, 
  (string) ( $cmConceive->diff($firstdate, true)->days - ($jonathan+$hannah) )
  ]; // Expecting Caleb!!!!
$notPregFirstDate[] = [ 
  (string) $cmBirth->getTimestamp()*1000, 
  (string) ( $cmBirth->diff($firstdate, true)->days - $daysPrego )
  ]; // Caleb Born!!!!
$notPregFirstDate[] = [ 
  (string) $todaysdate->getTimestamp()*1000, 
  (string) ( $todaysdate->diff($firstdate, true)->days - $daysPrego )
  ]; // Today
?>
<!DOCTYPE html>
<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html;charset=utf-8">
        <title>Jeri &amp; Scott! - CountDown Timer!</title>
        <link rel="icon" href="favicon.ico" type="image/x-icon" /> 
        <link rel="shortcut icon" href="favicon.ico" type="image/x-icon" /> 
        <link rel="stylesheet" href="css/jquery.countdown.css">
        <link rel="stylesheet" href="css/main.css">
        <script language="Javascript" type="text/javascript" src="js/jquery-1.4.1.js"></script>
        <script language="Javascript" type="text/javascript" src="js/jquery.lwtCountdown-1.0.js"></script>
        <script language="Javascript" type="text/javascript" src="js/jquery.flot.js"></script>
        <script language="javascript" type="text/javascript" src="js/jquery.flot.time.js"></script>
        <script language="Javascript" type="text/javascript" src="js/misc.js"></script>
        <link rel="Stylesheet" type="text/css" href="style/dark.css"></link>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />

    </head>
    <body>
      <div id="header">
        <div id="left_link"><a href="weddingPics.php">Our Wedding</a></div>
        <div id="middle_link"><a href="jonathan.php">Our Jonathan</a></div>
        <div id="right_link"><a href="hannah.php">Our Hannah</a></div>
      </div>
        <a name="topofthepage"></a>
        <div id="main_div">
            <div id="copy-text">
                <div id="pagetop"></div>
                <div id="pagetopmid"></div>
                <div id="pageback">
                    <div id="pagebody"> 
                        We started talking<br>
                        <em><?php echo($daysKnown->days); ?> days ago</em>
                        <br><br>
                        
                        It has been <br>
                        <em><?php echo($sinceFirst->days); ?> days</em><br>
                        since our first date.
                        <br><br>
                        
                        <hr id="hr3" class="ir">
                        <!-- Countdown dashboard start -->
                        Happily Married For:<br>
                        <em><?php echo($daysMarried->days); ?> days</em>
                        <br><br>
                        <hr id="hr3" class="ir">
                        She has been pregnant <br>
                        <em><?php echo($daysPrego); ?> days</em><br>
                        since we have been together.
                        <br><br>
                        She has NOT been pregnant<br>
                        <em><?php echo($daysNotPrego); ?> days</em><br>
                        since we have been together.
                        <br><br>
                        <div id="placeholder" style="width:600px;height:300px"></div>
                    </div>
                </div>
                <div id="pagebottom">
                    <a id="backtotop" class="ir sprite" href="#topofthepage">
                            Click here to scroll back to the top
                    </a>
                </div>
            </div>
        </div>
        <div id="footer">
          <div id="paginate">
            &nbsp;&nbsp;&nbsp;
          </div>
        </div>
        
        <script type="text/javascript">
          var pregs             = <?php echo(json_encode($pregDays)); ?>;
          var notPregsMeet      = <?php echo(json_encode($notPregDays)); ?>;
          var notPregsFirstDate = <?php echo(json_encode($notPregFirstDate)); ?>;
          
          $.plot("#placeholder", [{
                label: "Days Pregnant",
                data: pregs,
                lines: { show: true, fill: true },
                points: { show: true }
              }, {
                label: "Days NOT Preg - Since met",
                data: notPregsMeet,
                lines: { show: true, fill: true },
                points: { show: true }
              }, {
                label: "Days NOT Preg - Since First Date",
                data: notPregsFirstDate,
                lines: { show: true, fill: true },
                points: { show: true }
              }],
            {
              xaxis: {
                label: "Days",
                mode: "time",
              },
              yaxis: {
                label: "Date",
              },
              legend: {
                position: "nw",
                backgroundOpacity: 1
              }
            });
        </script>
    </body>
</html>

