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
        $subdealer = "No";
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
            $subdealer = "Yes";
            break;

            case "il":
            $territory = "21,23,90,91,123,127";
            $subdealer = "Yes";
            $dealerID = "553";
            break;

            case "ks":
            $territory = "58,142,139";
            $subdealer = "Yes";
            break;

            case "or":
            $territory = "95,121,116";
            $subdealer = "Yes";
            break;

            case "nv":
            $territory = "122,64,117,128,62";
            $subdealer = "Yes";
            break;

            default:
            $territory = "0";
            break;            
        }
        $data['territory'] = $territory;
        $data['subdealer'] = $subdealer;
        $data['dealerID'] = $dealerID;
        return($data);
    }

    public function installs($territory,$start,$end,$subdealer,$dealerID='') {
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
            if ($subdealer == "Yes") {
                $data[$i]['Name'] = $row['Name'];
            }            
        	$data[$i]['FullName'] = $row['FullName'];
        	$data[$i]['LicenseNumber'] = $row['LicenseNumber'];
        	$data[$i]['InstallDate'] = $row['InstallDate'];
        	$i++;
        }
        return($data);    	
    }

 	public function removals($territory,$start,$end,$subdealer,$dealerID='') {
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
            if ($subdealer == "Yes") {
                $data[$i]['Name'] = $row['Name'];
            }            
        	$data[$i]['FullName'] = $row['FullName'];
        	$data[$i]['LicenseNumber'] = $row['LicenseNumber'];
        	$data[$i]['RemovalDate'] = $row['RemovalDate'];
        	$i++;
        }
        return($data);   		
 	}

 	public function downloads($territory,$start,$end,$subdealer,$dealerID='') {
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
            if ($subdealer == "Yes") {
                $data[$i]['Name'] = $row['Name'];
            }
        	$data[$i]['FullName'] = $row['FullName'];
        	$data[$i]['LicenseNumber'] = $row['LicenseNumber'];
        	$data[$i]['DownloadDate'] = $row['DownloadDate'];
        	$data[$i]['Comment'] = $row['Comment'];
        	$i++;
        }
        return($data);   		 		
 	}

    public function createfile($tab1,$tab2,$tab3,$filename,$site_path,$subdealer)
    {

        $spreadsheet = new Spreadsheet();

        $myWorkSheet1 = new Worksheet($spreadsheet, 'Installs');
        $spreadsheet->addSheet($myWorkSheet1, 0);

        // Header
        $spreadsheet->getProperties()->setCreator('ADS')
        ->setLastModifiedBy('ADS')
        ->setTitle('ADS Monthly Report')
        ->setSubject('ADS Monthly Report')
        ->setDescription('ADS Monthly Report')
        ->setKeywords('ADS Monthly Report')
        ->setCategory('ADS Monthly Report');


        // page 1
        $spreadsheet->setActiveSheetIndex(0);

        if ($subdealer == "Yes") {
            $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);

            $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'DEALER')
            ->setCellValue('B1', 'SUB-DEALER')
            ->setCellValue('C1', 'CUSTOMER')
            ->setCellValue('D1', 'LN')
            ->setCellValue('E1', 'INSTALL DATE');

            // style
            $spreadsheet->getActiveSheet()->getStyle('A1:E1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE);  

            $dataArray = $tab1;
            $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A2');
            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFont()->setBold(true);
            $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        } else {
            $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);

            $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'DEALER')
            ->setCellValue('B1', 'CUSTOMER')
            ->setCellValue('C1', 'LN')
            ->setCellValue('D1', 'INSTALL DATE');

            // style
            $spreadsheet->getActiveSheet()->getStyle('A1:D1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

            $spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE);   

            $dataArray = $tab1;
            $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A2');
            $spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setBold(true);
            $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        }

        // page 2
        $myWorkSheet2 = new Worksheet($spreadsheet, 'Removals');
        $spreadsheet->addSheet($myWorkSheet2, 1);

        $dataArray = $tab2;
        //$spreadsheet->createSheet();

        $spreadsheet->setActiveSheetIndex(1);

        if ($subdealer == "Yes") {
            $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);

            $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'DEALER')
            ->setCellValue('B1', 'SUB-DEALER')
            ->setCellValue('C1', 'CUSTOMER')
            ->setCellValue('D1', 'LN')
            ->setCellValue('E1', 'REMOVAL DATE');

            // style
            $spreadsheet->getActiveSheet()->getStyle('A1:E1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE); 

            $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A2');
            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFont()->setBold(true);
            $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        } else {

            $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);

            $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'DEALER')
            ->setCellValue('B1', 'CUSTOMER')
            ->setCellValue('C1', 'LN')
            ->setCellValue('D1', 'REMOVAL DATE');

            // style
            $spreadsheet->getActiveSheet()->getStyle('A1:D1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

            $spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE); 

            $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A2');
            $spreadsheet->getActiveSheet()->getStyle('A1:D1')->getFont()->setBold(true);
            $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        }

        // page 3
        $myWorkSheet3 = new Worksheet($spreadsheet, 'Downloads');
        $spreadsheet->addSheet($myWorkSheet3, 2);

        $dataArray = $tab3;
        //$spreadsheet->createSheet();

        $spreadsheet->setActiveSheetIndex(2);

        if ($subdealer == "Yes") {
            $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('F')->setWidth(60);

            $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'DEALER')
            ->setCellValue('B1', 'SUB-DEALER')
            ->setCellValue('C1', 'CUSTOMER')
            ->setCellValue('D1', 'LN')
            ->setCellValue('E1', 'SERVICE DATE')
            ->setCellValue('F1', 'COMMENT');

            // style
            $spreadsheet->getActiveSheet()->getStyle('A1:F1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

            $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE); 

            $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A2');
            $spreadsheet->getActiveSheet()->getStyle('A1:F1')->getFont()->setBold(true);
            $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        } else {

            $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('C')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(20);
            $spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(60);

            $spreadsheet->getActiveSheet()
            ->setCellValue('A1', 'DEALER')
            ->setCellValue('B1', 'CUSTOMER')
            ->setCellValue('C1', 'LN')
            ->setCellValue('D1', 'SERVICE DATE')
            ->setCellValue('E1', 'COMMENT');

            // style
            $spreadsheet->getActiveSheet()->getStyle('A1:E1')
            ->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()->setARGB(Color::COLOR_BLUE); 

            $spreadsheet->getActiveSheet()->fromArray($dataArray, null, 'A2');
            $spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFont()->setBold(true);
            $spreadsheet->getActiveSheet()->setAutoFilter($spreadsheet->getActiveSheet()->calculateWorksheetDimension());

        }

        // Clean Up
        $sheetIndex = $spreadsheet->getIndex(
            $spreadsheet->getSheetByName('Worksheet')
        );
        $spreadsheet->removeSheetByIndex($sheetIndex);        
        $spreadsheet->setActiveSheetIndex(0);

        // Save
        $writer = new Xlsx($spreadsheet);
        $writer->save('helloworld.xlsx');


        $writer = new Xlsx($spreadsheet);
        $newfile = $site_path . "/" . $filename;
        $writer->save($newfile);

        return null;
    }    

}