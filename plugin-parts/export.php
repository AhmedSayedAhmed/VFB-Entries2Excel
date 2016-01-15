<?php
/**
 * Created by PhpStorm.
 * User: mm
 * Date: 04.09.2015
 * Time: 11:34
 */
require_once(ABSPATH . "wp-content/plugins/vfb-pro/inc/class-db-data.php");

class export
{
    private $headers = array();
    private $data = array();

    public function viewEntries($form_id)
    {

        // the object that reads the entry data
        $vfbdb = new VFB_Pro_Data();
        // getting all post ids for entries for that form id
        $entriesIDs = $vfbdb->get_entries_meta_by_form_id($form_id, "");

        // looping through entries to get individual data
        foreach ($entriesIDs as $entryID) {
            $entry_id = $entryID['ID'];
            $fields = $vfbdb->get_fields($form_id, "AND field_type NOT IN ('page-break','captcha','submit') ORDER BY field_order ASC");


            // looping through each field for a sing entry
            if (is_array($fields) && !empty($fields)) {

                foreach ($fields as $field) {
                    $label = isset($field['data']['label']) ? $field['data']['label'] : '';
                    $meta_key = '_vfb_field-' . $field['id'];
                    $value = $vfbdb->get_entry_meta_by_id($entry_id, $meta_key);

                    // skipping the file-upload field type
                    if (in_array($field['field_type'], array('url', 'file-upload'))) {
                        continue;
                    }
                    if (empty($value)) {
                        continue;
                    }
                    $value = strip_tags($value);

                    // adding the header label to the headers variable
                    if (!in_array($label, $this->headers)) {
                        array_push($this->headers, $label);
                    }

                    // exporting the data to the array
                    $data_pair = array($label, $value);
                    array_push($this->data, $data_pair);
                }
            }
        }
// hardcoded for now :S
        for ($i = 0; $i < sizeof($this->headers); $i++) {
            if ($this->headers[$i] === 'bearbeitet') {
                $temp = $this->headers[11];
                $this->headers[11] = $this->headers[$i];
                $this->headers[$i] = $temp;
            }
        }
    }

    public function flush2Xls()
    {
        /** Error reporting */
        error_reporting(E_ALL);
        ini_set('display_errors', TRUE);
        ini_set('display_startup_errors', TRUE);
        date_default_timezone_set('Europe/Berlin');

        if (PHP_SAPI == 'cli')
            die('This example should only be run from a Web Browser');

        /** Include PHPExcel */
        require_once dirname(__FILE__) . '/PHPExcel/PHPExcel.php';

// Create new PHPExcel object
        $objPHPExcel = new PHPExcel();

// Set document properties
        $objPHPExcel->getProperties()->setCreator("Entries2Excel")
            ->setLastModifiedBy("Ahmed Sayed")
            ->setTitle("Form Export " . date("Y-m-d H:i:s"))
            ->setSubject("Visual forms entries export")
            ->setDescription("Generated using PHPExcel.");


// Adding the headers-row
        $objPHPExcel->setActiveSheetIndex(0);
        // setting the constant for the sheet rows
        $alphabet = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K',
            'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');

        // main loop that prints everything
        for ($i = 0; $i < sizeof($this->headers); $i++) {
            if ($i < 26) {
                $cell = $alphabet[$i];
            } elseif ($i < 52) {
                $cell = 'A' . $alphabet[$i - 26];
            }
            // printing the headers
            $objPHPExcel->getActiveSheet()->setCellValue($cell . '1', $this->headers[$i]);

            //styling the headers
            $objPHPExcel->getActiveSheet()->getStyle($cell . '1')->getFont()->setBold(true);
            $objPHPExcel->getActiveSheet()->getStyle($cell . '1')->getAlignment()->setWrapText(TRUE);
            $objPHPExcel->getActiveSheet()->getColumnDimension($cell)->setWidth(17);


            // variable that decides which cell to use for the data
            $order = 2;

            // internal loop that prints the sheet data
            for ($j = 0; $j < sizeof($this->data); $j++) {
                if ($this->data[$j][0] === $this->headers[$i]) {
                    $objPHPExcel->getActiveSheet()->setCellValue($cell . $order, $this->data[$j][1]);
                    $objPHPExcel->getActiveSheet()->getStyle($cell . $order)->getAlignment()->setWrapText(TRUE);
                    $order++;
                }
            }
        }

// Set styling
        $objPHPExcel->getActiveSheet()->freezePane('A2');

// Save Excel 2007 file
        $objPHPExcel->setActiveSheetIndex(0);
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save(str_replace('.php', '.xlsx', __FILE__));
    }
}