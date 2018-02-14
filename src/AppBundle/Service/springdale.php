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

class springdale extends Controller
{

    protected $em;
    protected $container;

    public function __construct(EntityManagerInterface $entityManager, ContainerInterface $container)
    {
        $this->em = $entityManager;
        $this->container = $container;

    }

    public function distro($name) {
        $distro = "";

        switch ($name) {
            case "ads60":
            $distro = "106";
            break;

            case "ads30":
            $distro = "117";
            break;

        }

        return($distro);
    }

    public function getDealerNames_v2($distro) {
        $em = $this->em;
        $sql = "SELECT `CompanyName` FROM `Dealers` WHERE `DistributorID` = '$distro' ORDER BY `CompanyName` ASC";
        $data = array();
        $i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
            $data[$i] = $row['CompanyName'];
            $i++;
        }
        return($data);        

    }

    public function getDealerNames($distro,$start,$end) {
        /*
        There are 3 query, each could have dealers and each could be different. So we
        will need to run all 3 query and build the dealers into a unique list and
        store that into a value. We will then break the 3 queries out and use the dealer
        as part of the query.
        */

        $em = $this->em;

        // install
        $sql = "
        SELECT 
            de.CompanyName,
            DATE(MIN(Imported)) AS InstallDate
        FROM BaiidReports
          INNER JOIN Drivers dr USING(DriverID)
          INNER JOIN Dealers de USING(DealerID)
          INNER JOIN Distributors USING(DistributorID)
        WHERE DistributorID IN ($distro)
        GROUP BY DriverID
        HAVING InstallDate BETWEEN '$start' AND '$end'
        ORDER BY CompanyName
        ";

        $data = array();
        $i = "0";
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
        ) AND DistributorID IN ($distro)
        GROUP BY DriverID
        HAVING DATE(MAX(Imported)) BETWEEN '$start' AND '$end'
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
          AND DistributorID IN ($distro)
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

    public function installs_v2($dealername,$distro,$start,$end) {
        $em = $this->em;

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
        WHERE DistributorID IN ($distro)
        GROUP BY DriverID
        HAVING DATE(MIN(Imported)) BETWEEN '$start' AND '$end' AND CompanyName = ?
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

    public function removals_v2($dealername,$distro,$start,$end) {
        $em = $this->em;

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
        ) AND DistributorID IN ($distro)
        GROUP BY DriverID
        HAVING DATE(MAX(Imported)) BETWEEN '$start' AND '$end' AND CompanyName = ?
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

    public function downloads_v2($dealername,$distro,$start,$end) {
        $em = $this->em;

        $sql = "
        SELECT 
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
          AND DistributorID IN ($distro)
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
            $i++;
        }
        return($data);                  
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

                //if (($count1 > 0) or ($count2 > 0) or ($count3 > 0)) {
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
                    $temp = array();
                    if(empty($data)) {
                        $temp[0]['FullName'] = 'No data';
                        $temp[0]['LicenseNumber'] = '';
                        $temp[0]['InstallDate'] = '';
                        $temp[0]['Name'] = '';
                        $dealer_install_data[$dealer] = $temp;
                        $data = $temp;
                    }

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
                    $temp = array();
                    if(empty($data)) {
                        $temp[0]['FullName'] = 'No data';
                        $temp[0]['LicenseNumber'] = '';
                        $temp[0]['RemovalDate'] = '';
                        $temp[0]['Name'] = '';
                        $dealer_removal_data[$dealer] = $temp;
                        $data = $temp;
                    }

                    if(is_array($data)) {
                        $part3 = count($data);
                        $spreadsheet->getActiveSheet()->fromArray($data, null, $cell1);

                        $num = ($part3 + $next_cell) - 1;    
                        $cells = "A".$next_cell.":D" . $num;
                        $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray);                        
                    }
                 
                    // Downloads

                    $next_cell = $next_cell + $part3 + 1;

                    $cells = "A" . $next_cell . ":E" . $next_cell;

                    $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

                    $cell = "A".$next_cell.":E".$next_cell;
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

                    $spreadsheet->getActiveSheet()
                    ->setCellValue($cell1, 'CUSTOMER')
                    ->setCellValue($cell2, 'LN')
                    ->setCellValue($cell3, 'DOWNLOAD DATE')
                    ->setCellValue($cell4, 'SUB-DEALER')       
                    ->setCellValue($cell5, 'COMMENT');         


                    $cells = "A" . $next_cell . ":E" . $next_cell;

                    $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleTitleArray);

                    $cell = "A".$next_cell.":E".$next_cell;
                    $next_cell++;
                    $cell1 = "A".$next_cell;

                    $part4 = "";
                    $data = $dealer_download_data[$dealer];

                    $temp = array();
                    if(empty($data)) {
                        $temp[0]['FullName'] = 'No data';
                        $temp[0]['LicenseNumber'] = '';
                        $temp[0]['DownloadDate'] = '';
                        $temp[0]['Name'] = '';
                        $temp[0]['Comment'] = '';
                        $dealer_download_data[$dealer] = $temp;
                        $data = $temp;
                    }

                    if(is_array($data)) {
                        $part4 = count($data);
                        $spreadsheet->getActiveSheet()->fromArray($data, null, $cell1);

                        $num = ($part4 + $next_cell) - 1;    
                        $cells = "A".$next_cell.":E" . $num;
                        $spreadsheet->getActiveSheet()->getStyle($cells)->applyFromArray($styleBodyArray);                         
                    }

                    $counter++;
                //}

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



}