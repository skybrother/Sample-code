<?php
/* 
 * Custom Member Foyer Display for the Northwest church of Christ
 *
 * @author Scott Richardson - April 2016 - <scottric@tx.rr.com>
 */

  // Defining the basic cURL function
  function curl($url) {
      // Assigning cURL options to an array
      $options = Array(
          CURLOPT_RETURNTRANSFER => TRUE,  // Setting cURL's option to return the webpage data
          CURLOPT_FOLLOWLOCATION => TRUE,  // Setting cURL to follow 'location' HTTP headers
          CURLOPT_AUTOREFERER => TRUE, // Automatically set the referer where following 'location' HTTP headers
          CURLOPT_CONNECTTIMEOUT => 120,   // Setting the amount of time (in seconds) before the request times out
          CURLOPT_TIMEOUT => 120,  // Setting the maximum amount of time for cURL to execute queries
          CURLOPT_MAXREDIRS => 10, // Setting the maximum number of redirections to follow
          CURLOPT_USERAGENT => "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.1a2pre) Gecko/2008073000 Shredder/3.0a2pre ThunderBrowse/3.2.1.8",  // Setting the useragent
          CURLOPT_URL => $url, // Setting cURL's URL option with the $url variable passed into the function
      );

      $ch = curl_init();  // Initialising cURL 
      curl_setopt_array($ch, $options);   // Setting cURL's options using the previously assigned array data in $options
      $data = curl_exec($ch); // Executing the cURL request and assigning the returned data to the $data variable
      curl_close($ch);    // Closing cURL 
      return $data;   // Returning the data from the function 
  }

// Defining the basic scraping function
  function scrape_between($data, $start, $end){
      $data = stristr($data, $start); // Stripping all data from before $start
      $data = substr($data, strlen($start));  // Stripping $start
      $stop = stripos($data, $end);   // Getting the position of the $end of the data to scrape
      $data = substr($data, 0, $stop);    // Stripping all data from after and including the $end of the data to scrape
      return $data;   // Returning the scraped data from the function
  }

  $scraped_page = curl("http://www.pursuingthepath.com/memberboard");    
  $scraped_data = scrape_between($scraped_page, "<!-- Additional required wrapper -->", "<!-- If we need pagination -->");

  $parsed_data = str_replace('src="/media', 'src="http://www.pursuingthepath.com/media', $scraped_data);
  $parsed_data = str_replace('<div class="swiper-wrapper">', '', $parsed_data);

?>
<html>
  <head>
    <title>Northwest Members</title>
    <meta http-equiv="refresh" content="600">
    <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.10.1/jquery.min.js" type="text/javascript"></script>
    <script src="js/jquery.carouFredSel.js" type="text/javascript"></script>

    <script type="text/javascript">
      $(document).ready( function() {

        var slide_deck = $("#carousel1");
        slide_deck.clone().attr("id","carousel2").appendTo($("#middle_slider"));
        slide_deck.clone().attr("id","carousel3").appendTo($("#bottom_slider"));

        var slide_count = $('.swiper-slide').length;
        var tmp_pos = Math.ceil(slide_count / 9);
        var slide_speed = 0.07;

          $("#carousel1").carouFredSel({
              width: "100%",
              height: "100%",
              direction: 'left',
              padding: [10, 10],
              items: {
                  start: 0,
              },
              scroll: {
                items: slide_count / 3,
                fx: "scroll"
              },
              auto: {
                  timeoutDuration: 0,
                  duration: slide_speed,
                  easing: "linear",
              }
          });
          $("#carousel2").carouFredSel({
              width: "100%",
              height: "100%",
              direction: 'left',
              padding: [10, 10],
              items: {
                  start: tmp_pos,
              },
              scroll: {
                fx: "scroll",
                items: slide_count / 3,
              },
              auto: {
                  timeoutDuration: 0,
                  duration: slide_speed,
                  easing: "linear",
              }
          });
          $("#carousel3").carouFredSel({
              width: "100%",
              height: "100%",
              direction: 'left',
              padding: [10, 10],
              items: {
                  start: tmp_pos * 2,
              },
              scroll: {
                fx: "scroll",
                items: slide_count / 3,
              },
              auto: {
                  timeoutDuration: 0,
                  duration: slide_speed,
                  easing: "linear",
              }
          });
      });
    </script>

    <style>
    body {
      font-family: Helvetica,Arial,sans-serif;
      color: #333;
      background: #358 none repeat scroll 0 0;
    }

    .swiper-slide {
      float: left;
      display: block;
      padding: 0 10px;
    }

    .swiper-slide img {
      height: 325px;
      padding: 0;
    }

    figcaption {
      background:rgba(210, 210, 210, 0.6);
      display: block;
      font-size: 1.125em;
      font-weight: bold;
      height: 50px;
      margin: auto;
      padding: 5px 0 0;
      position: relative;
      top: -50px;
      width: 100%;
    }

    figure {
      text-align: center;
      text-wrap: normal;
      padding: 0;
      margin: 0;
    }

    #directory {
      padding: 0 5px;
    }
    #top_slider,
    #middle_slider,
    #bottom_slider {
      height: 325px;
      width: 100%;
      overflow: hidden;
    }
    </style>

  </head>

  <body>
    <div id='directory'>
      <div id='top_slider'>
        <div id="carousel1">
        <?php echo $parsed_data; ?>
      </div>
      <div id='middle_slider'>
      </div>
      <div id='bottom_slider'>
      </div>
    </div>
  </body>
</html>

