<?php

namespace AppBundle\Command;

use Symfony\Bundle\FrameworkBundle\Command\ContainerAwareCommand;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use Mailgun\Mailgun;

class ReportsCommand extends ContainerAwareCommand
{
    protected function configure()
    {
        $this
            ->setName('app:reports')
            ->setDescription('This will email the reports daily. Usage php bin/console app:daily single|multiple state1 state2 state3 etc')
            ->addArgument(
                'type',
                InputArgument::REQUIRED,
                'daily, weekly or monthly'
            )
            ->addArgument(
                'state',
                InputArgument::IS_ARRAY | InputArgument::REQUIRED,
                'List each state seperated by a space'
            )
        ;
    }

    protected function execute(InputInterface $input, OutputInterface $output)
    {

        $doctrine = $this->getContainer()->get('doctrine');
        $em = $doctrine->getManager();
        $type = $input->getArgument('type');
        $state_array = $input->getArgument('state');

        if(is_array($state_array)) {
            foreach ($state_array as $key=>$state) {
                $mg_domain = $this->getContainer()->getParameter('mg_domain');
                $mg_from = $this->getContainer()->getParameter('mg_from');
                $mg_api_key = $this->getContainer()->getParameter('mg_api_key');

                $site_path = $this->getContainer()->getParameter('site_path');
                $email = $this->getContainer()->getParameter('reportemails');


                switch ($type) {
                    case "daily":
                        $today = date("Y-m-d");
                        $start = date("Y-m-d", strtotime($today . "-1 DAY"));
                        $end = $start;
                        $year = date("Y", strtotime($start));
                        $month = date("F", strtotime($start));
                        $day = date("d", strtotime($start));
                        $today = date("m/d/Y", strtotime($start));
                        $subj = "ADS daily " . strtoupper($state) . " report";
                        $body = "Report Type: Daily<br>Date: $today<br>State: ".strtoupper($state)."<br><br>";
                        $body_text = "Report Type: Daily | Date: $today | State: ".strtoupper($state);
                        $filename = $day . "_" . $month."_".$year."_".$state.".xlsx";
                    break;

                    case "weekly":

                        $previous_week = strtotime("-1 week +1 day");
                        $start = date("Y-m-d", strtotime("last monday midnight",$previous_week));
                        $end = date("Y-m-d", strtotime($start . "+ 6 DAY"));

                        $start_p = date("m/d/Y", strtotime($start));
                        $end_p = date("m/d/Y", strtotime($end));

                        $year = date("Y", strtotime($start));
                        $week = date("W", strtotime($start));
                        $subj = "ADS weekly " . strtoupper($state) . " report";
                        $body = "Report Type: Weekly<br>Date: $start_p to $end_p<br>Week: $week<br>Year: $year<br>State: ".strtoupper($state)."<br><br>";
                        $body_text = "Report Type: Weekly | Date: $start_p to $end_p | Week: $week | Year: $year | State: ".strtoupper($state);

                        $filename = $week."_".$year."_".$state.".xlsx";
                    break;

                    case "monthly":
                        $start = date("Y-m-d", strtotime("first day of previous month"));
                        $end = date("Y-m-d", strtotime("last day of previous month"));
                        $start_p = date("m/d/Y", strtotime($start));
                        $end_p = date("m/d/Y", strtotime($end));                
                        $year = date("Y", strtotime("first day of previous month"));
                        $month = date("F", strtotime("first day of previous month"));
                        $subj = "ADS monthly " . strtoupper($state) . " report";
                        $body = "Report Type: Monthly<br>Date: $start_p to $end_p<br>Month: $month<br>Year: $year<br>State: ".strtoupper($state)."<br><br>";
                        $body_text = "Report Type: Monthly | Date: $start_p to $end_p | Month: $month | Year: $year | State: ".strtoupper($state);
                        $filename = $month.$year."_".$state.".xlsx";
                    break;

                    default:
                        print "Error! Invalid type\n";
                        die;
                    break;                    
                }

                // test dates
                //$start = "2018-01-28";
                //$end = "2018-01-28";

                $data = $this->getContainer()->get('commonservices')->states($state);

                $territory = $data['territory'];
                $dealerID = $data['dealerID'];

                if ($territory == "0") {
                    $output->writeln("State is missing or unsupported state was passed. IE: php bin/console app:daily tn");
                    die;            
                }

                $dealers = $this->getContainer()->get('commonservices')
                ->getDealerNames($territory,$start,$end,$dealerID);
                $dealers = array_unique($dealers);

                // init
                $dealer_install_data = array();
                $dealer_removal_data = array();
                $dealer_download_data = array();
                $dealer_lockcodes_data = array();

                if (is_array($dealers)) {
                    foreach ($dealers as $key=>$dealer) {
                        // Installs
                        $dealer_install_data[$dealer] = $this->getContainer()->get('commonservices')
                        ->installs_v2($dealer,$territory,$start,$end,$dealerID);

                        // Removals
                        $dealer_removal_data[$dealer] = $this->getContainer()->get('commonservices')
                        ->removals_v2($dealer,$territory,$start,$end,$dealerID);

                        // Downloads
                        $dealer_download_data[$dealer] = $this->getContainer()->get('commonservices')
                        ->downloads_v2($dealer,$territory,$start,$end,$dealerID);

                        // Lock Codes
                        $dealer_lockcodes_data[$dealer] = $this->getContainer()->get('commonservices')
                        ->lockcodes_v2($dealer,$territory,$start,$end,$dealerID);
                    }
                }

                $counter = "0";
                $counter = count($dealer_install_data) + count($dealer_removal_data) + count($dealer_download_data);
                if ($counter > 0) {
                    $this->getContainer()->get('commonservices')->create_file_v2($dealers,$dealer_install_data,$dealer_removal_data,$dealer_download_data,$dealer_lockcodes_data,$filename,$site_path);

                    $mg = new Mailgun($mg_api_key, new \Http\Adapter\Guzzle6\Client());
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
                    $mg = new Mailgun($mg_api_key, new \Http\Adapter\Guzzle6\Client());
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