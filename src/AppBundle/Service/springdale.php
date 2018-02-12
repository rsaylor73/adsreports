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
















    // old


    public function getInstalls($date1,$date2) {
    	$em = $this->em;
    	$dealerID = "572";

    	$sql = "
	    SELECT 
	    	DriverID, 
	    	FullName, 
	    	LicenseNumber,
	      	DATE(MIN(Imported)) AS InstallDate
	    FROM BaiidReports
	      INNER JOIN Drivers USING(DriverID)
	    WHERE DealerID IN ($dealerID)
	    GROUP BY DriverID
	    HAVING DATE(MIN(Imported)) BETWEEN '$date1' AND '$date2'
    	";

    	$data = array();
    	$i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
        	foreach ($row as $key=>$value) {
        		$data[$i][$key] = $value;
        	}
        	$i++;
        }
        return($data);    	
    }

    public function getRemovals($date1,$date2) {
    	$em = $this->em;
    	$dealerID = "572";

    	$sql = "
	    SELECT 
	    	DriverID,
	        FullName,
	        LicenseNumber,
	        DATE(MAX(Imported)) AS RemovalDate
	    FROM BaiidReports
	      INNER JOIN Drivers USING(DriverID)
	    WHERE NOT EXISTS (
	      SELECT NULL
	      FROM Items
	      WHERE ProductID = 1
	        AND Items.DriverID = Drivers.DriverID
	    ) AND DealerID IN ($dealerID)
	    GROUP BY DriverID
	    HAVING DATE(MAX(Imported)) BETWEEN '$date1' AND '$date2'
    	";
    	$data = array();
    	$i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
        	foreach ($row as $key=>$value) {
        		$data[$i][$key] = $value;
        	}
        	$i++;
        }
        return($data);
    }

    public function getDownloads($date1,$date2) {
    	$em = $this->em;
    	$dealerID = "572";

    	$sql = "
	    SELECT DriverID
	      , FullName
	      , LicenseNumber
	      , DATE(Imported) AS DownloadDate
	      , REPLACE(Comment, '\n', ' ') AS Comment
	    FROM BaiidReports
	      INNER JOIN Drivers USING(DriverID)
	    WHERE Type = 'Details'
	      AND NOT Comment LIKE '%Server side removal detected%'
	      AND DealerID IN ($dealerID)
	      AND DATE(Imported) BETWEEN '$date1' AND '$date2'
    	";
    	$data = array();
    	$i = "0";
        $result = $em->getConnection()->prepare($sql);
        $result->execute();  
        while ($row = $result->fetch()) {
        	foreach ($row as $key=>$value) {
        		$data[$i][$key] = $value;
        	}
        	$i++;
        }
        return($data);
    }



}