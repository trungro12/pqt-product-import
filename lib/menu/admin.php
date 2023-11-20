<?php
class PQTProductImport_Menu_Admin
{
    static function init()
    {
        self::addMenu();
    }


    static function addMenu()
    {
        add_action('admin_menu', function(){
            add_options_page(PQT_PRODUCT_IMPORT_NAME, PQT_PRODUCT_IMPORT_NAME, 'administrator', 'pqt-product-import', 'PQTProductImport_Menu_Admin::__htmlMenu', 1);
        });
    }

    static function __htmlMenu()
    {
?>
        <div class="wrap">
            <form action="" method="post" action="options.php">
                <table class="form-table" role="presentation">
                    <tbody>
                        <tr>
                            <th scope="row"><label for="blogname">Import Product (Excel File)</label></th>
                            <td><input name="importFile"  accept=".xlsx" type="file" id="importFile" value="test" class="regular-text"></td>
                        </tr>
                    </tbody>
                </table>
                <?php submit_button('Save') ?>
            </form>
        </div>
<?php
    }
}
