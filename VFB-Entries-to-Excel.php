<?php
/*
Plugin Name:    Entries2Excel
Plugin URI:     #
Description:    This plugin allows you to download the entries submitted by
				Visual Form Builder as an Excel sheet export
Version:        1.0
Author:         Ahmed Sayed
Author URI:     https://eg.linkedin.com/in/ahmeddabdallah
License:        GPL2
Note:			this plugin was built using code from:
				https://dipakgajjar.com/export-mysql-to-excel-xls-using-php/
*/


// for security reasons
defined('ABSPATH') or die('Plugin file cannot be accessed directly.');
require_once(ABSPATH . "wp-content/plugins/VFB-E2Excel/plugin-parts/listForms.php");
require_once(ABSPATH . "wp-content/plugins/VFB-E2Excel/plugin-parts/export.php");

function admin_menu_function()
{
    //variable that holds the path to the excel sheet
    $xlsPath = site_url() . "/wp-content/plugins/VFB-E2Excel/plugin-parts/export.xlsx";


    if (!isset($_GET['form_id'])) {
        echo "<br>";
        echo "<h1>Please select a form:</h1>";
        //if user didn't select a form display the forms to choose from
        new listForms();
    } else {
        //construct the excel sheet for the chosen form
        $sheet = new export();
        $sheet->viewEntries($_GET['form_id']);
        $sheet->flush2Xls();

        echo "<br>";
        echo "<h1>Click below to download your excel sheet</h1>";
        echo '<a href="' . $xlsPath . ' "><button id="download_all" href="' . $xlsPath . '" >Download all Entries</button></a>';


    }
}

function add_admin_menu()
{
    add_menu_page('Entries2Excel Page', 'Entries2Excel', 'manage_options', 'entries2excel', 'admin_menu_function');
}

add_action('admin_menu', 'add_admin_menu');