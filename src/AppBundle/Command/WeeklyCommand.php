<?php

namespace AppBundle\Command;

use Symfony\Bundle\FrameworkBundle\Command\ContainerAwareCommand;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use Mailgun\Mailgun;

class WeeklyCommand extends ContainerAwareCommand
{
    protected function configure()
    {
        $this
            ->setName('app:weekly')
            ->setDescription('This will email the reports weekly. Usage php bin/console app:weekly single|multiple state1 state2 state3 etc')
            ->addArgument(
                'format',
                InputArgument::REQUIRED,
                'single or multiple emails'
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
        $format = $input->getArgument('format');
        $state_array = $input->getArgument('state');

        if(is_array($state_array)) {
            foreach ($state_array as $key=>$state) {
                $mg_domain = $this->getContainer()->getParameter('mg_domain');
                $mg_from = $this->getContainer()->getParameter('mg_from');
                $mg_api_key = $this->getContainer()->getParameter('mg_api_key');

                $site_path = $this->getContainer()->getParameter('site_path');
                $email = $this->getContainer()->getParameter('reportemails');

                $previous_week = strtotime("-1 week +1 day");
                $start = date("Y-m-d", strtotime("last monday midnight",$previous_week));
                $end = date("Y-m-d", strtotime($start . "+ 6 DAY"));

                $start_p = date("m/d/Y", strtotime($start));
                $end_p = date("m/d/Y", strtotime($end));

                $year = date("Y", strtotime($start));
                $week = date("W", strtotime($start));

                $data = $this->getContainer()->get('commonservices')->states($state);

                $territory = $data['territory'];
                $subdealer = $data['subdealer'];
                $dealerID = $data['dealerID'];

                if ($territory == "0") {
                    $output->writeln("State is missing or unsupported state was passed. IE: php bin/console app:daily tn");
                    die;            
                }

                $report_type = "Weekly";

                $installs = $this->getContainer()
                    ->get('commonservices')
                    ->installs($territory,$start,$end,$subdealer,$dealerID);
                $removals = $this->getContainer()
                    ->get('commonservices')
                    ->removals($territory,$start,$end,$subdealer,$dealerID);
                $downloads = $this->getContainer()
                    ->get('commonservices')
                    ->downloads($territory,$start,$end,$subdealer,$dealerID);

                $filename = $week."_".$year."_".$state.".xlsx";

                $this->getContainer()
                    ->get('commonservices')
                    ->createfile($installs,$removals,$downloads,$filename,$site_path,$subdealer);

                if ($format == "single") {
                    $attach[] = $site_path . "/" . $filename;
                }
                if ($format == "multiple") {
                    $attach_file = $site_path . "/" . $filename;
                   
                    $subj = "ADS weekly " . strtoupper($state) . " report";
                    $today = date("m/d/Y");

                    $mg = new Mailgun($mg_api_key, new \Http\Adapter\Guzzle6\Client());
                    $msg = $mg->MessageBuilder();
                    $msg->setFromAddress($mg_from);
                    $msg->addToRecipient(implode(',',$email));
                    $msg->setSubject($subj);
                    $msg->setTextBody(strtoupper($state).' '. $report_type . ' report for week : '. $week . ' year : ' . $year);
                    $msg->setHTMLBody("<html><body><p>Date: <b>$start_p to $end_p</b><br>Week: <b>$week</b><br><br>".strtoupper($state).' '. $report_type . ' report for week : '. $week . ' year : ' . $year."</p><br><br></body></html>");
                    $files['attachment'] = array();
                    $files['attachment'][] = $attach_file;

                    $mg->post($mg_domain."/messages", $msg->getMessage(), $files);
                }  
            }
        }
        if ($format == "single") {
            $subj = "ADS weekly report";

            $today = date("m/d/Y");

            $mg = new Mailgun($mg_api_key, new \Http\Adapter\Guzzle6\Client());
            $msg = $mg->MessageBuilder();
            $msg->setFromAddress($mg_from);
            $msg->addToRecipient(implode(',',$email));
            $msg->setSubject($subj);
            $msg->setTextBody('The weekly ADS reports are attached to this email.');
            $msg->setHTMLBody("<html><body><p>Date: <b>$start_p to $end_p</b><br>Week: <b>$week</b><br><br>The weekly ADS reports are attached to this email.</p><br><br></body></html>");
            $files['attachment'] = array();
            if(is_array($attach)) {
                foreach ($attach as $key=>$value) {
                    $files['attachment'][] = $value;
                }
            }        
            $mg->post($mg_domain."/messages", $msg->getMessage(), $files); 
        }  

    }

}