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

class oldstyleservices extends Controller
{
    
    protected $em;
    protected $container;

    public function __construct(EntityManagerInterface $entityManager, ContainerInterface $container)
    {
        $this->em = $entityManager;
        $this->container = $container;

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
            //$data[$i]['unlockcode'] = $this->getlockoutcodes($row['BaiidReportID']);
            //$data[$i]['unlockcode'] = $this->get('commonservices')->getlockoutcodes($row['BaiidReportID']);
            $i++;
        }
        return($data);                  
    }

    public function lockcodes_v2_oldstyle($territory,$start,$end,$dealerID='') {
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
        $unlockcode = "";  
        while ($row = $result->fetch()) {
            $unlockcode = $this->getlockoutcodes($row['BaiidReportID']);
            if ($unlockcode != "") {
                $data[$i]['CompanyName'] = $row['CompanyName'];
                $data[$i]['FullName'] = $row['FullName'];
                $data[$i]['LicenseNumber'] = $row['LicenseNumber'];
                $data[$i]['DownloadDate'] = $row['DownloadDate'];
                //$data[$i]['Name'] = $row['Name'];
                //$data[$i]['Comment'] = $row['Comment'];
                $data[$i]['unlockcode'] = $this->getlockoutcodes($row['BaiidReportID']);
            }
            $unlockcode = "";
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
                    $string = $event . " " . $date_stamp . " " . $time_stamp . " " . $event_code . " " . $text1 . " " . $text2 . " " . $text3;
                }
            }
        }
        return($string);
    }

    public function create_file_v2_oldstyle($install_data,$removal_data,$download_data,$unlock_data,$filename,$site_path) {

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
                    'argb' => 'ff6a68dd',
                ),
                'endColor' => array(
                    'argb' => 'ff6a68dd',
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
        $counter = "0"; 

        if (($count1 > 0) or ($count2 > 0) or ($count3 > 0)) {

            $myWorkSheet1 = new Worksheet($spreadsheet, 'Installs');
            $spreadsheet->addSheet($myWorkSheet1, $counter);
            $spreadsheet->setActiveSheetIndex($counter);                    

            $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(40);

            // Installs

            $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'DEALER')
            ->setCellValue('B1', 'CUSTOMER')
            ->setCellValue('C1', 'LN')
            ->setCellValue('D1', 'INSTALL DATE')
            ->setCellValue('E1', 'SUB-DEALER');

            $spreadsheet->getActiveSheet()->getStyle('A1:E1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE);

            //$spreadsheet->getActiveSheet()->getStyle('A1:E1')->applyFromArray($styleTitleArray);

            $data = $install_data;
            $part2 = "";
            $next_cell = "";
            if(is_array($data)) {
                $part2 = count($data);
                $spreadsheet->getActiveSheet()->fromArray($data, null, 'A2');

                //$num = $part2 + 3;
                $num = $part2 + 1;    
                $cells = "A2:E" . $num;

                $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray);

            }
          
            // Removals

	        $counter++;
	        print "Counter $counter\n";
	        $myWorkSheet1 = new Worksheet($spreadsheet, 'Removals');
	        $spreadsheet->addSheet($myWorkSheet1, $counter);
	        $spreadsheet->setActiveSheetIndex($counter);                    

	        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
	        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
	        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
	        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(40);
	        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(40);

	        $next_cell = "1";
            //$next_cell = 4 + $part2 + 1;
            $cells = "A" . $next_cell . ":E" . $next_cell;

            $spreadsheet->getActiveSheet()->getStyle('A1:E1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            
            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE);

            //$spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);



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

            $spreadsheet->getActiveSheet()->getStyle('A1:E1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            
            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE);

            //$spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

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

	        $counter++;
	        print "Counter $counter\n";
	        $myWorkSheet1 = new Worksheet($spreadsheet, 'Downloads');
	        $spreadsheet->addSheet($myWorkSheet1, $counter);
	        $spreadsheet->setActiveSheetIndex($counter);                    

	        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
	        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
	        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
	        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
	        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(40);
            $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(80);

	        $next_cell = "1";

            //$next_cell = $next_cell + $part3 + 1;

            $cells = "A" . $next_cell . ":F" . $next_cell;

            $spreadsheet->getActiveSheet()->getStyle('A1:F1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
            
            $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE);

            //$spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

            $cell1 = "A".$next_cell;
            $cell2 = "B".$next_cell;
            $cell3 = "C".$next_cell;
            $cell4 = "D".$next_cell;
            $cell5 = "E".$next_cell;
            $cell6 = "F".$next_cell;

            $spreadsheet->getActiveSheet()
            ->setCellValue($cell1, 'DEALER')
            ->setCellValue($cell2, 'CUSTOMER')
            ->setCellValue($cell3, 'LN')
            ->setCellValue($cell4, 'DOWNLOAD DATE')
            ->setCellValue($cell5, 'SUB-DEALER')       
            ->setCellValue($cell6, 'COMMENT');        


            $cells = "A" . $next_cell . ":F" . $next_cell;

            $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

            $cell = "A".$next_cell.":F".$next_cell;
            $next_cell++;
            $cell1 = "A".$next_cell;

            $part4 = "";
            $temp_next_cell = $next_cell;
            $lock_cell = "";
            $data = $download_data;
            if(is_array($data)) {
                $part4 = count($data);
                $spreadsheet->getActiveSheet()->fromArray($data, null, $cell1);
            }
        }

        // lock codes
        $counter++;
        $next_cell = "1";
        print "Counter $counter\n";
        $myWorkSheet1 = new Worksheet($spreadsheet, 'Lock Codes');
        $spreadsheet->addSheet($myWorkSheet1, $counter);
        $spreadsheet->setActiveSheetIndex($counter);                    

        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(80);


        $spreadsheet->getActiveSheet()
        ->setCellValue('A1', 'DEALER')
        ->setCellValue('B1', 'CUSTOMER')
        ->setCellValue('C1', 'LN')
        ->setCellValue('D1', 'SERVICE DATE')
        ->setCellValue('E1', 'COMMENT');

        $spreadsheet->getActiveSheet()->getStyle('A1:E1')
        ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
        
        $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()
        ->setFillType(Fill::FILL_SOLID)
        ->getStartColor()->setARGB(Color::COLOR_BLUE);

        $cells = "A" . $next_cell . ":E" . $next_cell;

        $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

        $cell = "A".$next_cell.":E".$next_cell;
        $next_cell++;
        $cell1 = "A".$next_cell;

        $part4 = "";
        $temp_next_cell = $next_cell;
        $lock_cell = "";
        $data = $unlock_data;
        if(is_array($data)) {
            $part4 = count($data);
            $spreadsheet->getActiveSheet()->fromArray($data, null, $cell1);
        }

        
        // end lock codes

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
