<?php

namespace AppBundle\Command;

use Symfony\Bundle\FrameworkBundle\Command\ContainerAwareCommand;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use Mailgun\Mailgun;

class SpringdaleCommand extends ContainerAwareCommand
{
    protected function configure()
    {
        $this
            ->setName('app:springdale')
            ->setDescription('This will email the reports daily. Usage php bin/console app:springdale daily ads60 ads30')
            ->addArgument(
                'type',
                InputArgument::REQUIRED,
                'daily, weekly or monthly'
            )
            ->addArgument(
                'distro',
                InputArgument::IS_ARRAY | InputArgument::REQUIRED,
                'List each distro seperated by a space'
            )
        ;
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {

        $mg_domain = $this->getContainer()->getParameter('mg_domain');
        $mg_from = $this->getContainer()->getParameter('mg_from');
        $mg_api_key = $this->getContainer()->getParameter('mg_api_key');

        $site_path = $this->getContainer()->getParameter('site_path');
        $email = $this->getContainer()->getParameter('reportemails');

        $doctrine = $this->getContainer()->get('doctrine');
        $em = $doctrine->getManager();
        $type = $input->getArgument('type');
        $distro_array = $input->getArgument('distro');

        if(is_array($distro_array)) {
            foreach ($distro_array as $key=>$value) {

                $title = "";
                if ($value == "ads60") {
                    $title = "Discount Interlock 60 Daily Activity Report";
                } elseif ($value == "ads30") {
                    $title = "Discount Interlock 30 Daily Activity Report";
                }

                // init vars
                $date = "";
                $date1 = "";
                $date2 = "";

                switch($type) {
                    case "daily":
                        $date = date("Y-m-d");
                        $date1 = date("Y-m-d", strtotime($date . "-1 day"));
                        $date2 = date("Y-m-d", strtotime($date . "-1 day"));

                        //$today = date("Y-m-d");
                        //$start = date("Y-m-d", strtotime($today . "-1 DAY"));
                        //$end = $start;
                        $year = date("Y", strtotime($date1));
                        $month = date("F", strtotime($date1));
                        $day = date("d", strtotime($date1));
                        $today = date("m/d/Y", strtotime($date1));
                        $subj = $title;
                        $body = "Report Type: " . $title . "<br>Date: $today<br><br>";
                        $body_text = "Report Type: ". $title . " | Date: $today";
                        $filename = $day . "_" . $month."_".$year."_".$value.".xlsx";

                    break;

                    case "weekly":
                        print "Only supported daily currently.\n";
                        die;
                    break;

                    case "monthly":
                        print "Only supported daily currently.\n";
                        die;
                    break;

                    default:
                        print "Only supported daily currently.\n";
                        die;
                    break;
                }

                $distro = $this->getContainer()->get('springdale')->distro($value);

                $dealers = $this->getContainer()->get('springdale')
                ->getDealerNames_v2($distro);
                $dealers = array_unique($dealers);

                // init
                $dealer_install_data = array();
                $dealer_removal_data = array();
                $dealer_download_data = array();

                if (is_array($dealers)) {
                    foreach ($dealers as $key=>$dealer) {
                        
                        // Installs
                        $dealer_install_data[$dealer] = $this->getContainer()->get('springdale')
                        ->installs_v2($dealer,$distro,$date1,$date2);

                        // Removals
                        $dealer_removal_data[$dealer] = $this->getContainer()->get('springdale')
                        ->removals_v2($dealer,$distro,$date1,$date2);

                        // Downloads
                        $dealer_download_data[$dealer] = $this->getContainer()->get('springdale')
                        ->downloads_v2($dealer,$distro,$date1,$date2);

                    }
                }


                $counter = "0";
                $counter = count($dealer_install_data) + count($dealer_removal_data) + count($dealer_download_data);
                if ($counter > 0) {
                    $this->getContainer()->get('springdale')->create_file_v2($dealers,$dealer_install_data,$dealer_removal_data,$dealer_download_data,$filename,$site_path);

                    //$mg = new Mailgun($mg_api_key, new \Http\Adapter\Guzzle6\Client());
                    $mg = new Mailgun($mg_api_key);
                    $msg = $mg->MessageBuilder();
                    $msg->setFromAddress($mg_from);
                    $msg->addToRecipient(implode(',',$email));
                    $msg->setSubject($subj);
                    $msg->setTextBody($body_text);
                    $msg->setHTMLBody("<html><body><p>".$body."</p><br><br></body></html>");

                    $attach_file = $site_path . "/" . $filename;
                    $files['attachment'] = array();
                    $files['attachment'][] = $attach_file;

                    $mg->post($mg_domain."/messages", $msg->getMessage(), $files);
                } else {
                    $body .= "<br><font color=\"red\">Sorry but there was not any data available for this report.</font><br>";
                    $body_text .= "\n\nSorry but there was not any data available for this report.\n\n";
                    //$mg = new Mailgun($mg_api_key, new \Http\Adapter\Guzzle6\Client());
                    $mg = new Mailgun($mg_api_key);
                    $msg = $mg->MessageBuilder();
                    $msg->setFromAddress($mg_from);
                    $msg->addToRecipient(implode(',',$email));
                    $msg->setSubject($subj);
                    $msg->setTextBody($body_text);
                    $msg->setHTMLBody("<html><body><p>".$body."</p><br><br></body></html>");

                    $mg->post($mg_domain."/messages", $msg->getMessage(), null);                    
                }


            }
        }

    }

}