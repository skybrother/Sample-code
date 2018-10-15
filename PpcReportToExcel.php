<?php
namespace ProspectRelationshipManager\AppBundle\FormatConverter;

use JWT\ProspectSdk\Model\UserContextInterface;

/**
 * Class PpcReportToExcel command
 *
 * @author Scott Richardson ~~~ <scott.richardson@jwt.com>
 */
 class PpcReportToExcel
 {
   /**
     * LeadRecords
     * @var array $reportData
     */
    public $reportData;

    public $year;

    public $month;

    /**
     * @var UserContextInterface $userContext
     */
    public $userContext;

    /**
     * Method __construct
     *
     * @param $reportData
     */
    public function __construct( $reportData, $year, $month, $userContext )
    {
        $this->reportData  = $reportData;
        $this->year        = $year;
        $this->month       = $month;
        $this->userContext = $userContext;
    }
 }
