<?php

namespace AppBundle\Command;

use Symfony\Bundle\FrameworkBundle\Command\ContainerAwareCommand;
use Symfony\Component\Console\Input\InputArgument;
use Symfony\Component\Console\Input\InputInterface;
use Symfony\Component\Console\Input\InputOption;
use Symfony\Component\Console\Output\OutputInterface;
use Mailgun\Mailgun;

class DailyTestCommand extends ContainerAwareCommand
{
    protected function configure()
    {
        $this
            ->setName('app2:test')
            ->setDescription('This is a test command for testing new featured for the daily, weekly and monthly reports.')
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

                $today = date("Y-m-d");
                $start = date("Y-m-d", strtotime($today . "-1 DAY"));
                $end = $start;
                $year = date("Y", strtotime($start));
                $month = date("F", strtotime($start));
                $day = date("d", strtotime($start));

                $data = $this->getContainer()->get('commonservices')->states($state);

                $territory = $data['territory'];
                $subdealer = $data['subdealer'];
                $dealerID = $data['dealerID'];

                if ($territory == "0") {
                    $output->writeln("State is missing or unsupported state was passed. IE: php bin/console app:daily tn");
                    die;            
                }

                $report_type = "Daily";

                $installs = $this->getContainer()
                    ->get('commonservices')
                    ->installs($territory,$start,$end,$subdealer,$dealerID);
                $removals = $this->getContainer()
                    ->get('commonservices')
                    ->removals($territory,$start,$end,$subdealer,$dealerID);
                $downloads = $this->getContainer()
                    ->get('commonservices')
                    ->downloads($territory,$start,$end,$subdealer,$dealerID);

                $filename = $day . "_" . $month."_".$year."_".$state.".xlsx";

                $this->getContainer()
                    ->get('commonservices')
                    ->createfile($installs,$removals,$downloads,$filename,$site_path,$subdealer);

                if ($format == "single") {
                    $attach[] = $site_path . "/" . $filename;
                }
                if ($format == "multiple") {
                // MailGun  
                    $subj = "ADS daily " . strtoupper($state) . " report";
                    $attach_file = $site_path . "/" . $filename;
                    $today = date("m/d/Y");

                    $mg = new Mailgun($mg_api_key, new \Http\Adapter\Guzzle6\Client());
                    $msg = $mg->MessageBuilder();
                    $msg->setFromAddress($mg_from);
                    $msg->addToRecipient(implode(',',$email));
                    $msg->setSubject($subj);
                    $msg->setTextBody(strtoupper($state).' '. $report_type . ' report for '. $day . ' ' . $month . ' ' . $year);
                    $msg->setHTMLBody("<html><body><p>Date: <b>$today</b><br><br>".strtoupper($state).' '. $report_type . ' report for '. $day . ' ' . $month . ' ' . $year."</p><br><br></body></html>");                    
                    $files['attachment'] = array();
                    $files['attachment'][] = $attach_file;

                    $mg->post($mg_domain."/messages", $msg->getMessage(), $files);                       
                }
            }
        }

        if ($format == "single") {
            $subj = "ADS daily report";

            $today = date("m/d/Y");

            $mg = new Mailgun($mg_api_key, new \Http\Adapter\Guzzle6\Client());
            $msg = $mg->MessageBuilder();
            $msg->setFromAddress($mg_from);
            $msg->addToRecipient(implode(',',$email));
            $msg->setSubject($subj);
            $msg->setTextBody('The daily reports are attached to this email.');
            $msg->setHTMLBody("<html><body><p>Date: <b>$today</b><br><br>The daily ADS reports are attached to this email.</p><br><br></body></html>");
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