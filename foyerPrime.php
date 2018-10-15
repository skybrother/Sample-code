<?php
/* 
 * Custom Member Foyer Display for the Northwest church of Christ
 *
 * @author Scott Richardson - Sept 2016 - <scottric@tx.rr.com>
 */

/** Include PHPExcel_IOFactory */
require_once dirname(__FILE__) . '/Classes/PHPExcel/IOFactory.php';

$date_today = date('l jS').'<br/>'.date('F Y');
$date_month = "Assignments".date('Y');
$date_DOW   = date('D');
$date_ref   = "";

$date_add = 0;

switch ($date_DOW) {
  case "Mon":
    $date_add++;
  case "Tue":
    $date_add++;
  case "Wed":
    $date_ref = (date('j') + $date_add).date('-M');
    break;
  case "Thu":
    $date_add++;
  case "Fri":
    $date_add++;
  case "Sat":
    $date_add++;
  case "Sun":
    $date_ref = (date('j') + $date_add).date('-M');
    break;
}

$assignments_file = "Assignments/$date_month.xlsx";

if (!file_exists($assignments_file)) {
	exit("Please provide assignments spreadsheet. == " . $assignments_file);
}

$inputFileType = PHPExcel_IOFactory::identify($assignments_file);
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
$objPHPExcel = $objReader->load($assignments_file);

$sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);

$assign_output = "";
$spcr = '&nbsp;&nbsp;&nbsp;&nbsp;';
$cntr = 1;

foreach($sheetData as $row) {
  if($row['B'] == $date_ref) {
    // var_dump($row);
    switch($row['C'])
    {
      case "AM":
        $assign_output .= sundayAM($sheetData, $row, $spcr);
        $assign_output .= writeLine('', $spcr, true);
        break;
      case "PM":
        $assign_output .= sundayPM($sheetData, $row, $spcr);
        $assign_output .= writeLine('', $spcr, true);
        $assign_output .= wednesday($sheetData, $sheetData[$cntr+1], $spcr);
        $assign_output .= writeLine('', $spcr, true);
        $assign_output .= ushers($sheetData, $row, $spcr);
        break;
      default:
        $assign_output .= wednesday($sheetData, $row, $spcr);
        $assign_output .= writeLine('', $spcr, true);
        $assign_output .= sundayAM($sheetData, $sheetData[$cntr+1], $spcr);
        $assign_output .= writeLine('', $spcr, true);
        $assign_output .= sundayPM($sheetData, $sheetData[$cntr+2], $spcr);
        $assign_output .= writeLine('', $spcr, true);
        $assign_output .= ushers($sheetData, $row, $spcr);
        break;
    }
  }
  $cntr++;
}

function writeLine($label, $value, $blank = false) {
  if($blank) {
    return '<tr><td>&nbsp;</td><td>'.$value.'</td></tr>';
  } else {
    return '<tr><td>'.$label.' </td><td>'.$value.'</td></tr>';
  }
}

function writeCore($sheet, $row, $spcr) {
  $output  = writeLine($spcr.$sheet['1']['I'], $row['I']);
  $output .= writeLine($spcr.$sheet['1']['D'], $row['D']);
  $output .= writeLine($spcr.$sheet['1']['E'], $row['E']);
  $output .= writeLine($spcr.$sheet['1']['F'], $row['F']);
  $output .= writeLine($spcr.$sheet['1']['G'], $row['G']);
  $output .= writeLine($spcr.$sheet['1']['H'], $row['H']);
  return $output;
}

function sundayAM($sheetData, $row, $spcr) {
  $output  = writeLine('<b>Sunday Morning</b>', $spcr);
  $output .= writeCore($sheetData, $row, $spcr);
  $output .= writeLine($spcr.$sheetData['1']['J'], $row['J']);
  $output .= writeLine($spcr.$sheetData['1']['K'], $row['K']);
  $output .= writeLine($spcr, $row['L'], true);
  $output .= writeLine($spcr, $row['M'], true);
  $output .= writeLine($spcr, $row['N'], true);
  return $output;
}

function sundayPM($sheetData, $row, $spcr) {
  $output  = writeLine('<b>Sunday Evening</b>', $spcr);
  $output .= writeCore($sheetData, $row, $spcr);
  return $output;
}

function wednesday($sheetData, $row, $spcr) {
  $output  = writeLine('<b>Wednesday Night</b>', $spcr);
  $output .= writeLine($spcr.$sheetData['1']['D'], $row['D']);
  $output .= writeLine($spcr.$sheetData['1']['G'], $row['G']);
  $output .= writeLine($spcr.$sheetData['1']['H'], $row['H']);
  return $output;
}

function ushers($sheetData, $row, $spcr) {
  $output  = writeLine('<b>'.$sheetData['1']['O'].'</b>', $spcr);
  $output .= writeLine($spcr, $row['O']);
  $output .= writeLine($spcr, $row['P']);
  $output .= writeLine($spcr, $row['Q']);
  return $output;
}

// Column Names....
  //  'D' => string 'Song Leader' (length=11)
  //  'E' => string 'Opening Prayer' (length=14)
  //  'F' => string 'Second Scripture ' (length=17)
  //  'G' => string 'Sermon' (length=6)
  //  'H' => string 'Closing Prayer' (length=14)
  //  'I' => string 'First Scripture Reading' (length=23)
  //  'J' => string 'Serve Communion' (length=15)
  //  'L' => string 'Serve Communion' (length=15)
  //  'M' => string 'Serve Communion' (length=15)
  //  'N' => string 'Serve Communion' (length=15)
  //  'O' => string 'Ushers'
  //  'P' => string 'Ushers'
  //  'Q' => string 'Ushers'

?>
<html>
  <head>
    <title>Northwest Members</title>
    <meta http-equiv="refresh" content="3600">
    <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.1/jquery.min.js" type="text/javascript"></script>
    <script src="js/jquery.carouFredSel.js" type="text/javascript" ></script>

    <script type="text/javascript">
      $(document).ready( function() {

//        $('#notice-slider').carouFredSel({
//          padding: 5,
//          width: "100%",
//          height: "100%",
//          items: {
//            visible: 1,
//          },
//          scroll: {
//              items: 1,
//              fx: "crossfade",
//              duration: 1000
//          }
//        });

      });

      
    </script>

    <style>
    body {
      font-family: Helvetica,Arial,sans-serif;
      color: #333;
      background: #dddddd none repeat scroll 0 0;
    }

    #info_section {
      float: left;
      width: 25%;
      overflow: hidden;
    }

    #directory {
      float: left;
      width: 74%;
      padding-left: 10px;
      margin: 0;
      height: 106%;
    }
    
    #notice-slider img {
      height: 90%;
      float: left;
    }

    #assignments {
      border: 2px solid black;
      border-radius: 10px;
      /* float: left; */
      font-family: sans-serif;
      font-size: 30px;
      font-weight: bold;
      margin: 20px auto;
      padding: 15px;
      text-align: center;
      width: 90%;
      height: 75%;
    }
    
    #notice-slider {
      overflow: hidden;
    }

    #today_date {
      font-family: sans-serif;
      font-size: 30px;
      font-weight: bold;
      margin: 0 auto;
      padding: 15px;
      width: 90%;
      border: solid 2px black;
      border-radius: 10px;
      text-align: center;
    }
    </style>

  </head>

  <body>
    <div id='info_section'>
      <img src='http://www.pursuingthepath.com/media/uploads/images/gray-logo.jpg' />
      <div id="today_date">
        <?php echo $date_today; ?>
      </div>
      <!--
      <div id='notice-slider'>
        <img src='http://images.sharefaith.com/images/3/1254166493740_42/page-02.jpg' />
        <img src='http://images.sharefaith.com/images/3/1292892403800_436/page-02.jpg' />
        <img src='http://images.sharefaith.com/images/3/1266443805955_103/page-02.jpg' />  
      </div>
      -->
      <div id="assignments">
          <table>
            <?php echo $assign_output; ?>
          </table>
        </div>
    </div>
    <div id='directory'>
      <iframe src="http://www.teamjett.com/foyer.php" width="100%" height="100%" />
    </div>
  </body>
</html>

