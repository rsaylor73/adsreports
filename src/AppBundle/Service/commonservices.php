<?php
/* This is a service class */
namespace AppBundle\Service;

use Doctrine\ORM\EntityManagerInterface;
use Symfony\Component\DependencyInjection\ContainerInterface;
use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\DependencyInjection\ContainerAwareInterface;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

class commonservices extends Controller
{
    
    protected $em;
    protected $container;

    public function __construct(EntityManagerInterface $entityManager, ContainerInterface $container)
    {
        $this->em = $entityManager;
        $this->container = $container;

    }

    public function states($state) {
        $territory = "";
        $dealerID = "";
        $data = "";

        switch ($state) {
            case "tn":
            $territory = "107";
            break;

            case "az":
            $territory = "17";
            break;

            case "md":
            $territory = "3";
            break;

            case "mo":
            $territory = "30,129,130";
            break;

            case "il":
            $territory = "21,23,90,91,123,127";
            $dealerID = "553";
            break;

            case "ks":
            $territory = "58,142,139";
            break;

            case "or":
            $territory = "95,121,116";
            break;

            case "nv":
            $territory = "122,64,117,128,62";
            break;

            default:
            $territory = "0";
            break;            
        }
        $data = array(
                "territory" => $territory,
                "dealerID" => $dealerID
        );
        return($data);
    }

    public function getDealerNames($territory,$start,$end,$dealerID='') {
        /*
        There are 3 query, each could have dealers and each could be different. So we
        will need to run all 3 query and build the dealers into a unique list and
        store that into a value. We will then break the 3 queries out and use the dealer
        as part of the query.
        */

        $em = $this->em;

        $sql_dealer = "";
        if ($dealerID != "") {
            $sql_dealer = "AND DealerID IN (".$dealerID.")";
        }

        $data = array();
        $i = "0";

        // installs
        $sql = "
        SELECT 
            de.CompanyName,
            DATE(MIN(Imported)) AS InstallDate
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE TerritoryID IN ($territory)
        $sql_dealer
        GROUP BY DriverID
        HAVING InstallDate BETWEEN '$start' AND '$end'
        ORDER BY CompanyName
        ";

        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
            $data[$i] = $row['CompanyName'];
            $i++;
        }

        // removals
        $sql = "
        SELECT 
            de.CompanyName,
            DATE(MAX(Imported)) AS RemovalDate
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE NOT EXISTS (
          SELECT NULL
          FROM Items
          WHERE ProductID = 1
            AND Items.DriverID = dr.DriverID
        ) AND TerritoryID IN ($territory)
        $sql_dealer
        GROUP BY DriverID
        HAVING RemovalDate BETWEEN '$start' AND '$end'
        ORDER BY CompanyName
        ";    
           
        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
            $data[$i] = $row['CompanyName'];
            $i++;
        }

        // downloads
        $sql = "
        SELECT 
            de.CompanyName,
            DATE(Imported) DownloadDate
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE Type = 'Details'
          AND NOT Comment LIKE '%Server side removal detected%'
          AND DATE(Imported) BETWEEN '$start' AND '$end'
          AND TerritoryID IN ($territory)
          $sql_dealer
        ORDER BY CompanyName
        ";

        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
            $data[$i] = $row['CompanyName'];
            $i++;
        }
        return($data);        

    }

    public function installs_v2($dealername,$territory,$start,$end,$dealerID='') {
        $em = $this->em;

        $sql_dealer = "";
        if ($dealerID != "") {
            $sql_dealer = "AND DealerID IN (".$dealerID.")";
        }

        $sql = "
        SELECT 
            de.CompanyName,
            dr.EmployerName AS 'Name', 
            dr.FullName, 
            dr.LicenseNumber, 
            DATE(MIN(Imported)) AS InstallDate
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE TerritoryID IN ($territory)
        $sql_dealer
        GROUP BY DriverID
        HAVING InstallDate BETWEEN '$start' AND '$end' AND CompanyName = ?
        ORDER BY CompanyName, InstallDate, FullName
        ";

        $data = array();
        $i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->bindValue(1, $dealername);
        $result->execute();  
        while ($row = $result->fetch()) {          
            $data[$i]['FullName'] = $row['FullName'];
            $data[$i]['LicenseNumber'] = $row['LicenseNumber'];
            $data[$i]['InstallDate'] = $row['InstallDate'];
            $data[$i]['Name'] = $row['Name'];
            $i++;
        }
        return($data);      
    }

    public function installs_v2_oldstyle($territory,$start,$end,$dealerID='') {
        $em = $this->em;

        $sql_dealer = "";
        if ($dealerID != "") {
            $sql_dealer = "AND DealerID IN (".$dealerID.")";
        }

        $sql = "
        SELECT 
            de.CompanyName,
            dr.EmployerName AS 'Name', 
            dr.FullName, 
            dr.LicenseNumber, 
            DATE(MIN(Imported)) AS InstallDate
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE TerritoryID IN ($territory)
        $sql_dealer
        GROUP BY DriverID
        HAVING InstallDate BETWEEN '$start' AND '$end'
        ORDER BY CompanyName, InstallDate, FullName
        ";

        $data = array();
        $i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
            $data[$i]['CompanyName'] = $row['CompanyName'];
            $data[$i]['FullName'] = $row['FullName'];
            $data[$i]['LicenseNumber'] = $row['LicenseNumber'];
            $data[$i]['InstallDate'] = $row['InstallDate'];
            $data[$i]['Name'] = $row['Name'];
            $i++;
        }
        return($data);      
    }

    public function removals_v2($dealername,$territory,$start,$end,$dealerID='') {
        $em = $this->em;

        $sql_dealer = "";
        if ($dealerID != "") {
            $sql_dealer = "AND DealerID IN (".$dealerID.")";
        }

        $sql = "
        SELECT 
            de.CompanyName,
            dr.EmployerName AS 'Name', 
            dr.FullName, 
            dr.LicenseNumber, 
            DATE(MAX(Imported)) AS RemovalDate
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE NOT EXISTS (
          SELECT NULL
          FROM Items
          WHERE ProductID = 1
            AND Items.DriverID = dr.DriverID
        ) AND TerritoryID IN ($territory)
        $sql_dealer
        GROUP BY DriverID
        HAVING RemovalDate BETWEEN '$start' AND '$end' AND CompanyName = ?
        ORDER BY CompanyName, RemovalDate, FullName
        ";
        $data = array();
        $i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->bindValue(1, $dealername);
        $result->execute();  
        while ($row = $result->fetch()) {          
            $data[$i]['FullName'] = $row['FullName'];
            $data[$i]['LicenseNumber'] = $row['LicenseNumber'];
            $data[$i]['RemovalDate'] = $row['RemovalDate'];
            $data[$i]['Name'] = $row['Name'];
            $i++;
        }
        return($data);          
    }


    public function removals_v2_oldstyle($territory,$start,$end,$dealerID='') {
        $em = $this->em;

        $sql_dealer = "";
        if ($dealerID != "") {
            $sql_dealer = "AND DealerID IN (".$dealerID.")";
        }

        $sql = "
        SELECT 
            de.CompanyName,
            dr.EmployerName AS 'Name', 
            dr.FullName, 
            dr.LicenseNumber, 
            DATE(MAX(Imported)) AS RemovalDate
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE NOT EXISTS (
          SELECT NULL
          FROM Items
          WHERE ProductID = 1
            AND Items.DriverID = dr.DriverID
        ) AND TerritoryID IN ($territory)
        $sql_dealer
        GROUP BY DriverID
        HAVING RemovalDate BETWEEN '$start' AND '$end'
        ORDER BY CompanyName, RemovalDate, FullName
        ";
        $data = array();
        $i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {          
            $data[$i]['CompanyName'] = $row['CompanyName'];
            $data[$i]['FullName'] = $row['FullName'];
            $data[$i]['LicenseNumber'] = $row['LicenseNumber'];
            $data[$i]['RemovalDate'] = $row['RemovalDate'];
            $data[$i]['Name'] = $row['Name'];
            $i++;
        }
        return($data);          
    }


    public function downloads_v2($dealername,$territory,$start,$end,$dealerID='') {
        $em = $this->em;

        $sql_dealer = "";
        if ($dealerID != "") {
            $sql_dealer = "AND DealerID IN (".$dealerID.")";
        }

        $sql = "
        SELECT
            BaiidReportID, 
            de.CompanyName,
            dr.EmployerName AS 'Name', 
            dr.FullName, 
            dr.LicenseNumber, 
            DATE(Imported) DownloadDate, 
            REPLACE(Comment, '\n', ' ') AS Comment
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE Type = 'Details'
          AND NOT Comment LIKE '%Server side removal detected%'
          AND DATE(Imported) BETWEEN '$start' AND '$end'
          AND TerritoryID IN ($territory)
          $sql_dealer
        HAVING CompanyName = ?
        ORDER BY CompanyName, DownloadDate, FullName
        ";

        $data = array();
        $i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->bindValue(1, $dealername);
        $result->execute();  
        while ($row = $result->fetch()) {
            $data[$i]['FullName'] = $row['FullName'];
            $data[$i]['LicenseNumber'] = $row['LicenseNumber'];
            $data[$i]['DownloadDate'] = $row['DownloadDate'];
            $data[$i]['Name'] = $row['Name'];
            $data[$i]['Comment'] = $row['Comment'];
            $data[$i]['unlockcode'] = $this->getlockoutcodes($row['BaiidReportID']);
            $i++;
        }
        return($data);                  
    }

    public function downloads_v2_oldstyle($territory,$start,$end,$dealerID='') {
        $em = $this->em;

        $sql_dealer = "";
        if ($dealerID != "") {
            $sql_dealer = "AND DealerID IN (".$dealerID.")";
        }

        $sql = "
        SELECT
            BaiidReportID, 
            de.CompanyName,
            dr.EmployerName AS 'Name', 
            dr.FullName, 
            dr.LicenseNumber, 
            DATE(Imported) DownloadDate, 
            REPLACE(Comment, '\n', ' ') AS Comment
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE Type = 'Details'
          AND NOT Comment LIKE '%Server side removal detected%'
          AND DATE(Imported) BETWEEN '$start' AND '$end'
          AND TerritoryID IN ($territory)
          $sql_dealer
        ORDER BY CompanyName, DownloadDate, FullName
        ";

        $data = array();
        $i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
            $data[$i]['CompanyName'] = $row['CompanyName'];
            $data[$i]['FullName'] = $row['FullName'];
            $data[$i]['LicenseNumber'] = $row['LicenseNumber'];
            $data[$i]['DownloadDate'] = $row['DownloadDate'];
            $data[$i]['Name'] = $row['Name'];
            $data[$i]['Comment'] = $row['Comment'];
            $data[$i]['unlockcode'] = $this->getlockoutcodes($row['BaiidReportID']);
            $i++;
        }
        return($data);                  
    }

    public function getlockoutcodes($BaiidReportID) {
        $em = $this->em;
        $string = "";

        $sql = "
        SELECT 
            `r`.`RawReport`
        
        FROM `BaiidReports` r
            
        WHERE 
            `r`.`BaiidReportID` = '$BaiidReportID'
        ";
        $csv = "";

        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
            $csv = $row['RawReport'];
        }        

        $lines = explode(PHP_EOL, $csv);
        $array = array();
        foreach ($lines as $line) {
            $array[] = str_getcsv($line);
        }

        $code = "";
        $event = "";
        $date_stamp = "";
        $time_stamp = "";
        $event_code = "";
        $text1 = "";
        $text2 = "";
        $text3 = "";

        foreach ($array as $key=>$value) {
            if(is_array($value)) {
                $code = str_replace(' ','',$value[0]);
                if ($code == "2D") {
                    $event = $value[1];
                    $date_stamp = $value[2];
                    $time_stamp = $value[3];
                    $event_code = $value[4];
                    $text1 = $value[5];
                    $text2 = $value[6];
                    $text3 = $value[7];
                    $string = $text1 . " " . $text2 . " " . $text3;
                }
            }
        }
        return($string);
    }

    public function create_file_v2($dealers,$dealer_install_data,$dealer_removal_data,$dealer_download_data,$filename,$site_path) {

        $count1 = "0";
        $count2 = "0";
        $count3 = "0";

        // Style
        // changed startColor from FFA0A0A0 : FFFFFFFF
        // changed argb to rgb

        /*
        Notes:
        ffd7e1e8 = made a nice green
        ffb3c6d3 = lighter shade of gray
        ffcccccc = gray
        */
        $styleHeadArray = array(
            'font' => array(
                'bold' => true,
            ),            
            'alignment' => array(
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ),            
            'borders' => array(
                'top' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'bottom' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'left' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'right' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
            ),
            'fill' => array(
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startColor' => array(
                    'argb' => 'ffb3c6d3',
                ),
                'endColor' => array(
                    'argb' => 'ffb3c6d3',
                ),
            ),            
        );        

        $styleTitleArray = array(
            'font' => array(
                'bold' => true,
            ),            
            'alignment' => array(
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ),            
            'borders' => array(
                'top' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'bottom' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'left' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'right' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
            ),           
        ); 

        $styleBodyArray = array(
            'borders' => array(
                'allBorders' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ),
            ),          
        );        

        $spreadsheet = new Spreadsheet();

        // Header
        $spreadsheet->getProperties()
        ->setCreator('ADS')
        ->setLastModifiedBy('ADS')
        ->setTitle('ADS Report')
        ->setSubject('ADS Report')
        ->setDescription('ADS Report')
        ->setKeywords('ADS Report')
        ->setCategory('ADS Report');

        $counter = "0";

        if(is_array($dealers)) {
            foreach($dealers as $key=>$dealer) {

                $count1 = count($dealer_install_data[$dealer]);
                $count2 = count($dealer_removal_data[$dealer]);
                $count3 = count($dealer_download_data[$dealer]);

                if (($count1 > 0) or ($count2 > 0) or ($count3 > 0)) {
                    // Maximum 31 characters allowed in sheet title.
                    if (strlen($dealer) > 31) {
                        $dealer_title = substr($dealer,0,31);
                    } else {
                        $dealer_title = $dealer;
                    }
                    $myWorkSheet1 = new Worksheet($spreadsheet, $dealer_title);
                    $spreadsheet->addSheet($myWorkSheet1, $counter);
                    $spreadsheet->setActiveSheetIndex($counter);                    

                    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
                    $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
                    $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
                    $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(40);
                    $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(100);
                    $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(40);

                    // Title
                    $spreadsheet->getActiveSheet()->mergeCells('A1:D1');
                    $spreadsheet->getActiveSheet()->setCellValue('A1',$dealer);
                    $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
                    $spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setSize(16);

                    // Installs

                    $spreadsheet->getActiveSheet()->mergeCells('A2:D2');
                    $spreadsheet->getActiveSheet()->setCellValue('A2','INSTALLS');

                    $spreadsheet->getActiveSheet()->getStyle('A2:D2')->applyFromArray($styleHeadArray);

                    $spreadsheet->getActiveSheet()
                    ->setCellValue('A3', 'CUSTOMER')
                    ->setCellValue('B3', 'LN')
                    ->setCellValue('C3', 'INSTALL DATE')
                    ->setCellValue('D3', 'SUB-DEALER');

                    $spreadsheet->getActiveSheet()->getStyle('A3:D3')->applyFromArray($styleTitleArray);

                    $data = $dealer_install_data[$dealer];
                    $part2 = "";
                    $next_cell = "";
                    if(is_array($data)) {
                        $part2 = count($data);
                        $spreadsheet->getActiveSheet()->fromArray($data, null, 'A4');

                        $num = $part2 + 3;    
                        $cells = "A4:D" . $num;

                        $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray);

                    }
                  
                    // Removals

                    $next_cell = 4 + $part2 + 1;
                    $cells = "A" . $next_cell . ":D" . $next_cell;

                    $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);


                    $cell = "A".$next_cell.":D".$next_cell;
                    $spreadsheet->getActiveSheet()->mergeCells($cell);
                    $cell = "A".$next_cell;
                    $spreadsheet->getActiveSheet()->setCellValue($cell,'REMOVALS');

                    $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleHeadArray);


                    $next_cell++;

                    $cell1 = "A".$next_cell;
                    $cell2 = "B".$next_cell;
                    $cell3 = "C".$next_cell;
                    $cell4 = "D".$next_cell;
                    $spreadsheet->getActiveSheet()
                    ->setCellValue($cell1, 'CUSTOMER')
                    ->setCellValue($cell2, 'LN')
                    ->setCellValue($cell3, 'REMOVAL DATE')                    
                    ->setCellValue($cell4, 'SUB-DEALER');                    

                    $cells = "A" . $next_cell . ":D" . $next_cell;

                    $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

                    $cell = "A".$next_cell.":E".$next_cell;
                    $next_cell++;
                    $cell1 = "A".$next_cell;

                    $part3 = "";
                    $data = $dealer_removal_data[$dealer];
                    if(is_array($data)) {
                        $part3 = count($data);
                        $spreadsheet->getActiveSheet()->fromArray($data, null, $cell1);

                        $num = ($part3 + $next_cell) - 1;    
                        $cells = "A".$next_cell.":D" . $num;
                        $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray);                        
                    }
                 
                    // Downloads

                    $next_cell = $next_cell + $part3 + 1;

                    $cells = "A" . $next_cell . ":F" . $next_cell;

                    $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

                    $cell = "A".$next_cell.":F".$next_cell;
                    $spreadsheet->getActiveSheet()->mergeCells($cell);
                    $cell = "A".$next_cell;
                    $spreadsheet->getActiveSheet()->setCellValue($cell,'DOWNLOADS');

                    $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleHeadArray);

                    $next_cell++;

                    $cell1 = "A".$next_cell;
                    $cell2 = "B".$next_cell;
                    $cell3 = "C".$next_cell;
                    $cell4 = "D".$next_cell;
                    $cell5 = "E".$next_cell;
                    $cell6 = "F".$next_cell;

                    $spreadsheet->getActiveSheet()
                    ->setCellValue($cell1, 'CUSTOMER')
                    ->setCellValue($cell2, 'LN')
                    ->setCellValue($cell3, 'DOWNLOAD DATE')
                    ->setCellValue($cell4, 'SUB-DEALER')       
                    ->setCellValue($cell5, 'COMMENT')         
                    ->setCellValue($cell6, 'LOCK CODES');         


                    $cells = "A" . $next_cell . ":F" . $next_cell;

                    $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

                    $cell = "A".$next_cell.":F".$next_cell;
                    $next_cell++;
                    $cell1 = "A".$next_cell;

                    $part4 = "";
                    $temp_next_cell = $next_cell;
                    $lock_cell = "";
                    $data = $dealer_download_data[$dealer];
                    if(is_array($data)) {
                        $part4 = count($data);
                        $spreadsheet->getActiveSheet()->fromArray($data, null, $cell1);

                        foreach($data as $key=>$value) {
                            if(is_array($value)) {
                                foreach ($value as $key2=>$value2) {
                                    if ($key2 == "unlockcode") {
                                        if ($value2 != "") {
                                            $lock_cell = "F" . $temp_next_cell;
                                            //print "$key2 : $value2 : $lock_cell\n";
                                            // text color
                                            $spreadsheet->getActiveSheet()
                                            ->getStyle($lock_cell)
                                            ->getFont()
                                            ->getColor()
                                            ->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK);

                                            // background color
                                            $spreadsheet->getActiveSheet()
                                            ->getStyle($lock_cell)
                                            ->getFill()
                                            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                            ->getStartColor()
                                            ->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);

                                            $num = ($part4 + $next_cell) - 1;    
                                            $cells = "A".$next_cell.":F" . $num;
                                            $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray); 
                                        }
                                    }
                                }
                                $temp_next_cell++;
                            }
                        }                    
                    }

                    $counter++;
                }

            }

            // Clean Up
            $sheetIndex = $spreadsheet->getIndex(
                $spreadsheet->getSheetByName('Worksheet')
            );
            $spreadsheet->removeSheetByIndex($sheetIndex);        
            $spreadsheet->setActiveSheetIndex(0);

            // Save
            $writer = new Xlsx($spreadsheet);
            $newfile = $site_path . "/" . $filename;
            $writer->save($newfile);

        } else {
            // error - dealers is not an array
        }
    }

    public function create_file_v2_oldstyle($install_data,$removal_data,$download_data,$filename,$site_path) {

        $count1 = "0";
        $count2 = "0";
        $count3 = "0";

        $styleHeadArray = array(
            'font' => array(
                'bold' => true,
            ),            
            'alignment' => array(
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ),            
            'borders' => array(
                'top' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'bottom' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'left' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'right' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
            ),
            'fill' => array(
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startColor' => array(
                    'argb' => 'ffb3c6d3',
                ),
                'endColor' => array(
                    'argb' => 'ffb3c6d3',
                ),
            ),            
        );        

        $styleTitleArray = array(
            'font' => array(
                'bold' => true,
            ),            
            'alignment' => array(
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
            ),            
            'borders' => array(
                'top' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'bottom' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'left' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
                'right' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ),
            ),           
        ); 

        $styleBodyArray = array(
            'borders' => array(
                'allBorders' => array(
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ),
            ),          
        );        

        $spreadsheet = new Spreadsheet();

        // Header
        $spreadsheet->getProperties()
        ->setCreator('ADS')
        ->setLastModifiedBy('ADS')
        ->setTitle('ADS Report')
        ->setSubject('ADS Report')
        ->setDescription('ADS Report')
        ->setKeywords('ADS Report')
        ->setCategory('ADS Report');

        $count1 = count($install_data);
        $count2 = count($removal_data);
        $count3 = count($download_data);
        $counter = "0"; // does not count only static

        if (($count1 > 0) or ($count2 > 0) or ($count3 > 0)) {

            $myWorkSheet1 = new Worksheet($spreadsheet, 'Report');
            $spreadsheet->addSheet($myWorkSheet1, $counter);
            $spreadsheet->setActiveSheetIndex($counter);                    

            $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(40);
            $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(100);
            $spreadsheet->getActiveSheet()->getColumnDimension('G')->setWidth(40);

            // Installs

            $spreadsheet->getActiveSheet()->mergeCells('A1:E1');
            $spreadsheet->getActiveSheet()->setCellValue('A1','INSTALLS');

            $spreadsheet->getActiveSheet()->getStyle('A1:D1')->applyFromArray($styleHeadArray);

            $spreadsheet->getActiveSheet()
            ->setCellValue('A2', 'DEALER')
            ->setCellValue('B2', 'CUSTOMER')
            ->setCellValue('C2', 'LN')
            ->setCellValue('D2', 'INSTALL DATE')
            ->setCellValue('E2', 'SUB-DEALER');

            $spreadsheet->getActiveSheet()->getStyle('A2:E2')->applyFromArray($styleTitleArray);

            $data = $install_data;
            $part2 = "";
            $next_cell = "";
            if(is_array($data)) {
                $part2 = count($data);
                $spreadsheet->getActiveSheet()->fromArray($data, null, 'A3');

                $num = $part2 + 3;    
                $cells = "A4:E" . $num;

                $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray);

            }
          
            // Removals

            $next_cell = 4 + $part2 + 1;
            $cells = "A" . $next_cell . ":E" . $next_cell;

            $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);


            $cell = "A".$next_cell.":E".$next_cell;
            $spreadsheet->getActiveSheet()->mergeCells($cell);
            $cell = "A".$next_cell;
            $spreadsheet->getActiveSheet()->setCellValue($cell,'REMOVALS');

            $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleHeadArray);


            $next_cell++;

            $cell1 = "A".$next_cell;
            $cell2 = "B".$next_cell;
            $cell3 = "C".$next_cell;
            $cell4 = "D".$next_cell;
            $cell5 = "E".$next_cell;

            $spreadsheet->getActiveSheet()
            ->setCellValue($cell1, 'DEALER')
            ->setCellValue($cell2, 'CUSTOMER')
            ->setCellValue($cell3, 'LN')
            ->setCellValue($cell4, 'REMOVAL DATE')                    
            ->setCellValue($cell5, 'SUB-DEALER');                    

            $cells = "A" . $next_cell . ":E" . $next_cell;

            $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

            $cell = "A".$next_cell.":E".$next_cell;
            $next_cell++;
            $cell1 = "A".$next_cell;

            $part3 = "";
            $data = $removal_data;
            if(is_array($data)) {
                $part3 = count($data);
                $spreadsheet->getActiveSheet()->fromArray($data, null, $cell1);

                $num = ($part3 + $next_cell) - 1;    
                $cells = "A".$next_cell.":E" . $num;
                $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray);                        
            }
         
            // Downloads

            $next_cell = $next_cell + $part3 + 1;

            $cells = "A" . $next_cell . ":G" . $next_cell;

            $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

            $cell = "A".$next_cell.":G".$next_cell;
            $spreadsheet->getActiveSheet()->mergeCells($cell);
            $cell = "A".$next_cell;
            $spreadsheet->getActiveSheet()->setCellValue($cell,'DOWNLOADS');

            $spreadsheet->getActiveSheet()->getStyle($cell)->applyFromArray($styleHeadArray);

            $next_cell++;

            $cell1 = "A".$next_cell;
            $cell2 = "B".$next_cell;
            $cell3 = "C".$next_cell;
            $cell4 = "D".$next_cell;
            $cell5 = "E".$next_cell;
            $cell6 = "F".$next_cell;
            $cell7 = "G".$next_cell;

            $spreadsheet->getActiveSheet()
            ->setCellValue($cell1, 'DEALER')
            ->setCellValue($cell2, 'CUSTOMER')
            ->setCellValue($cell3, 'LN')
            ->setCellValue($cell4, 'DOWNLOAD DATE')
            ->setCellValue($cell5, 'SUB-DEALER')       
            ->setCellValue($cell6, 'COMMENT')         
            ->setCellValue($cell7, 'LOCK CODES');         


            $cells = "A" . $next_cell . ":G" . $next_cell;

            $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

            $cell = "A".$next_cell.":G".$next_cell;
            $next_cell++;
            $cell1 = "A".$next_cell;

            $part4 = "";
            $temp_next_cell = $next_cell;
            $lock_cell = "";
            $data = $download_data;
            if(is_array($data)) {
                $part4 = count($data);
                $spreadsheet->getActiveSheet()->fromArray($data, null, $cell1);

                foreach($data as $key=>$value) {
                    if(is_array($value)) {
                        foreach ($value as $key2=>$value2) {
                            if ($key2 == "unlockcode") {
                                if ($value2 != "") {
                                    $lock_cell = "G" . $temp_next_cell;
                                    //print "$key2 : $value2 : $lock_cell\n";
                                    // text color
                                    $spreadsheet->getActiveSheet()
                                    ->getStyle($lock_cell)
                                    ->getFont()
                                    ->getColor()
                                    ->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK);

                                    // background color
                                    $spreadsheet->getActiveSheet()
                                    ->getStyle($lock_cell)
                                    ->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()
                                    ->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);

                                    $num = ($part4 + $next_cell) - 1;    
                                    $cells = "A".$next_cell.":G" . $num;
                                    $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray); 
                                }
                            }
                        }
                        $temp_next_cell++;
                    }
                }                    
            }
        }

        // Clean Up
        $sheetIndex = $spreadsheet->getIndex(
            $spreadsheet->getSheetByName('Worksheet')
        );
        $spreadsheet->removeSheetByIndex($sheetIndex);        
        $spreadsheet->setActiveSheetIndex(0);

        // Save
        $writer = new Xlsx($spreadsheet);
        $newfile = $site_path . "/" . $filename;
        $writer->save($newfile);

    }    





} // end class
