<?php
/**
 * class that builds the excel sheet file
 * original code taken from:
 * https://dipakgajjar.com/export-mysql-to-excel-xls-using-php/
 * and I made some adjustments
 *
 */
/*
 *


*/

class listForms
{

    function __construct()
    {
        $form_id = '';
        $vfbdb = new VFB_Pro_Data();
        $forms = $vfbdb->get_all_forms();
            ?>
        <select id="vfb-entry-form-ids" name="form-id" onchange="this.options[this.selectedIndex].value != 0 && (window.location = window.location+'&form_id='+this.options[this.selectedIndex].value);">
            <option value="0"><?php _e('Select a Form', 'vfb-pro'); ?></option>
            <?php
            if (is_array($forms) && !empty($forms)) {
                foreach ($forms as $form) {
                    $entry_count = $vfbdb->get_entries_count($form['id']);

                    echo sprintf(
                        '<option value="%1$d"%3$s>%1$d - %2$s (%4$d)</option>',
                        $form['id'],
                        $form['title'],
                        selected($form['id'], $form_id, false),
                        $entry_count
                    );
                }
            }
            ?>
        </select>
    <?php
    }

}