<?php

namespace AppBundle\Command;

use Symfony\Bundle\FrameworkBundle\Command\ContainerAwareCommand;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use Mailgun\Mailgun;

class LogsCommand extends ContainerAwareCommand
{
    protected function configure()
    {
        $this
            ->setName('app:logs')
            ->setDescription('test')
        ;
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {

        $doctrine = $this->getContainer()->get('doctrine');
        $em = $doctrine->getManager();

        $serviceID = "1300134";

        /*
        $sql0 = "
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
          AND DATE(Imported) BETWEEN '2018-02-01' AND '2018-02-22'
          AND TerritoryID IN (107)
        ORDER BY CompanyName, DownloadDate, FullName 
        ";

        $result0 = $em->getConnection()->prepare($sql0);
        $result0->execute();  
        while ($row0 = $result0->fetch()) {
        */
            //$serviceID = $row0['BaiidReportID'];
            $sql = "
            SELECT 
                `r`.`RawReport`
            
            FROM `BaiidReports` r
                
            WHERE 
                `r`.`BaiidReportID` = '$serviceID'
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

                        print "$serviceID : $code : $text1 $text2 $text3\n";
                    }

                }
            }
        //}

    }
}
