<?php

use \PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

class PQTProductImport_Menu_Admin
{
    static function init()
    {
        require_once(PQT_PRODUCT__PLUGIN_DIR . "/vendor/autoload.php");
        self::addMenu();
    }


    static function addMenu()
    {
        add_action('admin_menu', function () {
            add_options_page(PQT_PRODUCT_IMPORT_NAME, PQT_PRODUCT_IMPORT_NAME, 'administrator', 'pqt-product-import', 'PQTProductImport_Menu_Admin::__htmlMenu', 1);
        });
    }

    static function __htmlMenu()
    {
        $excelMauURL = esc_url(PQT_PRODUCT__PLUGIN_URL . '/download/mau.xlsx');
        $arrImported = [];
        if (
            isset($_FILES['importFile']) &&
            !empty($_POST['action']) &&
            !empty($_POST['_wpnonce']) &&
            wp_verify_nonce($_POST['_wpnonce'], $_POST['action'])
        ) {
            $fileData = $_FILES['importFile'];
            $importNumber = (int) $_POST['importNumber'];
            $importUnique = (!empty($_POST['importUnique'])) && (int) $_POST['importUnique'] === 1 ? true : false;
            $arrDataImport = self::readExcelFile($fileData, $importNumber);
            $arrImported = self::insertProduct($arrDataImport, $importUnique);
        }
?>
        <h1><?php echo PQT_PRODUCT_IMPORT_NAME; ?></h1>
        <p style="font-weight:bold">Import sản phẩm Woocommerce theo Format đã có sẵn, tải <a href="<?php echo $excelMauURL; ?>"> mẫu ở đây</a></p>
        <div class="wrap">
            <?php if (!empty($arrImported)) : ?>
                <div class="notice notice-success">
                    <p>Đã Import <?php echo count($arrImported); ?> sản phẩm thành công.</p>
                </div>
            <?php endif; ?>

            <form action="" method="post" action="options.php" enctype="multipart/form-data">
                <table class="form-table" role="presentation">
                    <tbody>
                        <tr>
                            <th scope="row"><label for="blogname">Import Product (Excel File)</label></th>
                            <td><input name="importFile" accept=".xlsx" type="file" id="importFile" value="test" class="regular-text">
                            </td>
                        </tr>

                        <tr>
                            <th scope="row"><label for="blogname">Số lượng (0 là toàn bộ): </label></th>
                            <td><input name="importNumber" type="number" id="importNumber" value="0" class="regular-text"></td>
                        </tr>

                        <tr>
                            <th scope="row"><label for="blogname">Check trùng lặp: </label></th>
                            <td>
                                <input name="importUnique" checked type="checkbox" id="importUnique" value="1" class="regular-text"><br>
                                <label for="">Nếu có sản phẩm trùng, chỉ cập nhật các thuộc tính, không tạo sản phẩm mới.</label>
                            </td>
                        </tr>

                    </tbody>

                    <input type="hidden" name="action" value="importFile">
                    <?php wp_nonce_field('importFile');  ?>

                </table>
                <?php submit_button('Import') ?>
            </form>
        </div>
<?php
    }



    static function readExcelFile($fileData, $max = 0)
    {
        $arrData = [];
        $inputFile = $fileData['tmp_name'];
        $extension = strtoupper(end(explode(".", $fileData['name'])));
        if ($extension == 'XLSX' || $extension == 'ODS') {

            $reader = new Xlsx();
            $spreadsheet = $reader->load($inputFile);
            $worksheet = $spreadsheet->getActiveSheet();
            $worksheet_arr = $worksheet->toArray();

            // Remove header row 
            unset($worksheet_arr[0]);
            $count = 0;
            foreach ($worksheet_arr as $row) {
                $count++;
                $arrData[] = $row;
                if ($max > 0 && $max < $count) break;
            }
        } else {
            echo "Please upload an XLSX or ODS file";
        }
        return $arrData;
    }

    static function insertProduct($arrDataImport, $unique = false)
    {
        $postData = [];
        foreach ($arrDataImport as $value) {
            $sku = sanitize_text_field($value[0]); // SKU
            $productCatName = sanitize_text_field($value[1]); // Product Category
            $attrItemType = sanitize_text_field($value[2]); // Item Type
            $productName = sanitize_text_field($value[3]); // Product Name
            $attrDesigner = sanitize_text_field($value[4]); // Designer
            $attrBrand = sanitize_text_field($value[5]); // Brand
            $productDescription = sanitize_text_field($value[6]); // Product Description
            $attrGender = sanitize_text_field($value[7]); // Gender
            $attrFragranceNotes = sanitize_text_field($value[8]); // Fragrance Notes
            $attrYearIntroduced = sanitize_text_field($value[9]); // Year Introduced
            $attrRecommendedUse = sanitize_text_field($value[10]); // Recommended Use
            $salePrice = sanitize_text_field($value[11]); // MSRP => sale price
            $price = sanitize_text_field($value[12]); // FNET Wholesale Price => price
            $urlImageLarge = sanitize_text_field($value[13]); // Image Large URL
            $urlImageSmall = sanitize_text_field($value[14]); // Image Small URL
            $url = sanitize_text_field($value[15]); // url product

            // create new product 
            $postId = 0;
            $post = $unique ? self::getProductByName($productName) : null;
            if (empty($post)) {
                $postId = wp_insert_post(array(
                    //'post_title' => 'Adams Product',
                    'post_title' => $productName,
                    'post_content' => $productDescription,
                    'post_status' => 'publish',
                    'post_type' => "product",
                ));
            } else $postId = $post->ID;

            if ($postId) {

                // insert product category 
                $category = get_term_by('name', $productCatName, 'product_cat');
                $categoryId = 0;
                if (empty($category)) {
                    $category = wp_insert_term(
                        $productCatName, // the term 
                        'product_cat', // the Woocommerce product category taxonomy
                        array( // (optional)
                            // 'description'=> 'This is a red apple.', // (optional)
                            // 'slug' => 'apple', // optional
                            // 'parent'=> $parent_term['term_id']  // (Optional) The parent numeric term id
                        )
                    );
                }
                if ($category) $categoryId = $category->term_id;
                // set cat to product 
                wp_set_post_terms($postId, array($categoryId), 'product_cat', true);


                // add product attr 
                $attrs = [
                    "Item Type" => $attrItemType,
                    "Designer" => $attrDesigner,
                    "Brand"   => $attrBrand,
                    "Gender" => $attrGender,
                    "Year Introduced" => $attrYearIntroduced,
                    "Recommended Use" => $attrRecommendedUse,
                ];
                $productAttr = [];
                foreach ($attrs as $name => $value) {
                    $productAttr[] =  [
                        "name" => $name,
                        "value" => $value,
                        "is_visible" => 1,
                        "variation" => 0,
                    ];
                }

                update_post_meta($postId, '_product_attributes', $productAttr);

                update_post_meta($postId, '_visibility', 'visible');
                update_post_meta($postId, '_stock_status', 'instock');
                update_post_meta($postId, 'total_sales', '0');
                update_post_meta($postId, '_downloadable', 'no');
                update_post_meta($postId, '_virtual', 'no');
                update_post_meta($postId, '_regular_price', $price);
                update_post_meta($postId, '_sale_price', $salePrice);
                update_post_meta($postId, '_purchase_note', $attrFragranceNotes);
                update_post_meta($postId, '_featured', 'no');
                // update_post_meta($postId, '_weight', '');
                // update_post_meta($postId, '_length', '');
                // update_post_meta($postId, '_width', '');
                // update_post_meta($postId, '_height', '');
                update_post_meta($postId, '_sku', $sku);
                // update_post_meta($postId, '_sale_price_dates_from', '');
                // update_post_meta($postId, '_sale_price_dates_to', '');
                update_post_meta($postId, '_price', $price);
                update_post_meta($postId, '_sold_individually', '');
                update_post_meta($postId, '_manage_stock', 'no');
                update_post_meta($postId, '_backorders', 'no');
                update_post_meta($postId, '_stock', '');

                // upload image 
                self::uploadImageToPost($urlImageLarge, $postId);

                $postData[] = $postId;
            }
        }
        return $postData;
    }






    static function uploadImageToPost($imageUrl, $postId)
    {
        $filename = self::downloadImageFromUrl($imageUrl);

        // The ID of the post this attachment is for.
        $parent_post_id = $postId;

        // Check the type of file. We'll use this as the 'post_mime_type'.
        $filetype = wp_check_filetype(basename($filename), null);

        // Get the path to the upload directory.
        $wp_upload_dir = wp_upload_dir();

        // Prepare an array of post data for the attachment.
        $attachment = array(
            'guid'           => $wp_upload_dir['url'] . '/' . basename($filename),
            'post_mime_type' => $filetype['type'],
            'post_title'     => preg_replace('/\.[^.]+$/', '', basename($filename)),
            'post_content'   => '',
            'post_status'    => 'inherit'
        );

        // Insert the attachment.
        $attach_id = wp_insert_attachment($attachment, $filename, $parent_post_id);

        // Make sure that this file is included, as wp_generate_attachment_metadata() depends on it.
        require_once(ABSPATH . 'wp-admin/includes/image.php');

        // Generate the metadata for the attachment, and update the database record.
        $attach_data = wp_generate_attachment_metadata($attach_id, $filename);
        wp_update_attachment_metadata($attach_id, $attach_data);

        set_post_thumbnail($parent_post_id, $attach_id);
    }


    static function downloadImageFromUrl($imageUrl)
    {

        $imageurl = $imageUrl;
        $imagetype = end(explode('/', getimagesize($imageurl)['mime']));
        $uniq_name = date('dmY') . '' . (int) microtime(true);
        $filename = $uniq_name . '.' . $imagetype;

        $uploaddir = wp_upload_dir();
        $uploadfile = $uploaddir['path'] . '/' . $filename;
        $contents = file_get_contents($imageurl);
        $savefile = fopen($uploadfile, 'w');
        $isOK = fwrite($savefile, $contents);
        fclose($savefile);
        return $isOK ? $uploadfile : null;
    }

    static function getProductByName(string $name, string $post_type = "product")
    {
        $query = new WP_Query([
            "post_type" => $post_type,
            "name" => $name
        ]);

        return $query->have_posts() ? reset($query->posts) : null;
    }
}
