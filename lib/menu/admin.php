<?php

use \PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

class PQTProductImport_Menu_Admin
{
    const IMPORT_UNIQUE_UPDATE = 1; // cập nhật sản phẩm trùng
    const IMPORT_UNIQUE_SKIP = 2; // skip các sản phẩm trùng

    static function init()
    {
        require_once(PQT_PRODUCT__PLUGIN_DIR . "/vendor/autoload.php");
        self::addAjaxInsertProduct();
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
        $excelMauURL = esc_url(PQT_PRODUCT__PLUGIN_URL . 'download/mau.xlsx');
        $arrImported = [];
        if (
            isset($_FILES['importFile']) &&
            !empty($_POST['action']) &&
            !empty($_POST['_wpnonce']) &&
            wp_verify_nonce($_POST['_wpnonce'], $_POST['action'])
        ) {
            $fileData = $_FILES['importFile'];
            $importNumber = (int) $_POST['importNumber'];
            $importStep = (int) $_POST['importStep'];
            $importUnique = (!empty($_POST['importUnique'])) ? (int) $_POST['importUnique'] : 2;
            $arrDataImport = self::readExcelFile($fileData, $importNumber);
            // $arrImported = self::insertProduct($arrDataImport, $importUnique);

            if ($arrDataImport === null) {
                echo '<script>alert("Không thể Upload File");</script>';
                $arrDataImport = [];
            }
?>

            <script>
                (function($) {
                    $(function() {
                        const importMax = <?php echo $importNumber; ?>;
                        const arrData = JSON.parse(atob('<?php echo base64_encode(json_encode($arrDataImport)); ?>'));
                        const productTotal = arrData.length;
                        var productOK = 0;
                        var productErr = 0;
                        const stepImport = <?php echo $importStep; ?>;
                        const showInfo = $('#showInfo');
                        showInfo.show();

                        function importProduct(arrProduct, cb = function() {}) {
                            $.ajax({
                                type: "post",
                                url: "<?php echo admin_url('admin-ajax.php'); ?>",
                                data: {
                                    action: "ajaxInsertProductImport",
                                    arrProduct: arrProduct,
                                    importUnique: <?php echo $importUnique; ?>,
                                },
                                success: function(msg) {
                                    if (msg.success) {
                                        productOK += msg.data.count;

                                    } else {}
                                    cb();
                                },
                                complete: function() {

                                }
                            });
                        }

                        var index = 0;
                        var isFinish = false;
                        // const count = arrData.length;
                        const callBackImport = function() {
                            $('#productOk').text(productOK);

                            if (index >= arrData.length) {
                                $('.waitForImport').remove();
                                $('.loadingscreen').remove();
                                if (arrData.length < 1) return;
                                
                                productErr = productTotal - productOK;
                                showInfo.find('#productError').text(productErr);
                                isFinish = true;
                                if (!isFinish) alert('Import hoàn tất!');
                                return;
                            };
                            const arrDataProduct = [];
                            for (let i = index; i < index + stepImport; i++) {
                                index++;
                                if (typeof arrData[i] !== 'object' || arrData.length <= i) break;
                                // console.log(arrData[i]);
                                arrDataProduct.push(arrData[i]);

                            }
                            if (arrDataProduct.length > 0) importProduct(arrDataProduct, callBackImport);
                        }
                        callBackImport();
                    });
                })(jQuery);
            </script>

        <?php
            self::loadingScreen();
        }
        ?>
        <h1><?php echo PQT_PRODUCT_IMPORT_NAME; ?></h1>
        <p style="font-weight:bold">Import sản phẩm Woocommerce theo Format đã có sẵn, tải <a href="<?php echo $excelMauURL; ?>"> mẫu ở đây</a></p>
        <div class="wrap">
            <div id="showInfo" style="display: none;">
                <div class="notice notice-success">
                    <p> <b style="color:red" class="waitForImport">Đang thực hiện...</b> Đã Import <b style="color: red;"><span id="productOk">0</span></b> sản phẩm thành công.</p>
                </div>
                <div class="notice notice-error">
                    <p><b style="color:red" class="waitForImport">Đang thực hiện...</b><b style="color: red;"><span id="productError">0</span></b> sản phẩm thất bại hoặc đã bỏ qua.</p>
                </div>
            </div>

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
                            <td><input name="importNumber" type="number" id="importNumber" value="<?php echo (int)(!empty($_POST['importNumber']) ? $_POST['importNumber'] : 0); ?>" class="regular-text"></td>
                        </tr>

                        <tr>
                            <th scope="row"><label for="blogname">Import mỗi lần: </label></th>
                            <td><input name="importStep" type="number" id="importStep" value="<?php echo (int) (!empty($_POST['importStep']) && $_POST['importStep'] ? $_POST['importStep'] : 20); ?>" class="regular-text"><br>
                                <label for="">Đối với file có nhiều sản phẩm, sẽ chia ra thành nhiều lần import để tránh bị gián đoạn.</label>
                            </td>
                        </tr>

                        <tr>
                            <th scope="row"><label for="blogname">Kiểm tra trùng lặp: </label></th>
                            <td>
                                <input name="importUnique" <?php echo empty($_POST['importUnique']) || (int) $_POST['importUnique'] == self::IMPORT_UNIQUE_UPDATE ? 'checked' : ''; ?> type="radio" value="<?php echo self::IMPORT_UNIQUE_UPDATE; ?>" class="regular-text">
                                <label for="">Cập nhật sản phẩm trùng lặp.</label><br>
                                <input name="importUnique" <?php echo !empty($_POST['importUnique']) && (int) $_POST['importUnique'] == self::IMPORT_UNIQUE_SKIP ? 'checked' : ''; ?> type="radio" value="<?php echo self::IMPORT_UNIQUE_SKIP; ?>" class="regular-text">
                                <label for="">Bỏ qua nếu có sản phẩm trùng lặp.</label>
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

        $inputFile = self::copyFileToUploadDir($inputFile);
        if (empty($inputFile)) return null;

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
                if ($max > 0 && $count >= $max) break;
            }
        } else {
            echo "Please upload an XLSX or ODS file";
        }
        unlink($inputFile);
        return $arrData;
    }

    static function insertProduct($arrDataImport, $unique = self::IMPORT_UNIQUE_UPDATE)
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
            // $post = self::getProductByName($productName);
            $post = self::getProductBySku($sku);

            if (!empty($post)) {

                // check unique 
                if ($unique == self::IMPORT_UNIQUE_SKIP) continue;

                $postId = $post->ID;
            }

            if (empty($post)) {
                $postId = wp_insert_post(array(
                    //'post_title' => 'Adams Product',
                    'post_title' => $productName,
                    'post_content' => $productDescription,
                    'post_status' => 'publish',
                    'post_type' => "product",
                ));
            }

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
                if (!has_post_thumbnail($postId)) {
                    self::uploadImageToPost($urlImageLarge, $postId);
                }

                $postData[] = $postId;
            }
        }
        return $postData;
    }






    static function uploadImageToPost($imageUrl, $postId)
    {
        $filename = self::downloadImageFromUrl($imageUrl);
        if (empty($filename)) return;

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

    static function copyFileToUploadDir($fileName)
    {
        $uploaddir = wp_upload_dir();
        if (!file_exists($uploaddir['path'])) mkdir($uploaddir['path'], 0755);
        $uploadFile = $uploaddir['path'] . '/' . basename($fileName);

        if (@move_uploaded_file($fileName, $uploadFile)) {
            return $uploadFile;
        }

        return null;
    }


    static function downloadImageFromUrl($imageUrl)
    {

        $imageurl = $imageUrl;
        $imagetype = end(explode('/', getimagesize($imageurl)['mime']));
        $uniq_name = date('dmY') . '' . (int) microtime(true);
        $filename = $uniq_name . '.' . $imagetype;

        $uploaddir = wp_upload_dir();

        if (!file_exists($uploaddir['path'])) mkdir($uploaddir['path'], 0755);

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

    static function getProductBySku(string $sku)
    {
        $query = [
            // 'lang'       => $lang,
            'post_type'  => 'product',
            'meta_query' => [
                [
                    'key'     => '_sku',
                    'value'   => $sku,
                    'compare' => '='
                ]
            ],
            // 'tax_query'  => [
            //     [
            //         'taxonomy' => 'product_type',
            //         'terms'    => [ 'grouped' ],
            //         'field'    => 'name',
            //     ]
            // ]
        ];
        $posts = (new WP_Query())->query($query);

        return count($posts) > 0 ? $posts[0] : null;
    }



    static function addAjaxInsertProduct()
    {
        function ajaxInsertProductImport()
        {
            $arrProduct = empty($_POST['arrProduct']) ? [] : $_POST['arrProduct'];
            $importUnique = empty($_POST['importUnique']) ? PQTProductImport_Menu_Admin::IMPORT_UNIQUE_UPDATE : (int) $_POST['importUnique'];
            if ($count = count(PQTProductImport_Menu_Admin::insertProduct($arrProduct, $importUnique))) {
                wp_send_json_success(['count' => $count]);
            } else wp_send_json_error();
            exit;
        }
        add_action('wp_ajax_ajaxInsertProductImport', 'ajaxInsertProductImport'); // executed when logged in
        // add_action('wp_ajax_ajaxInsertProductImport', 'ajaxInsertProductImport' ); // executed when logged out
    }


    static function loadingScreen()
    {
    ?>
        <style>
            .screen {
                opacity: .3;
                position: fixed;
                top: 0;
                left: 0;
                right: 0;
                bottom: 0;
                background: black;
                display: flex;
                align-items: center;
            }

            .loader {
                width: 100%;
                height: 15px;
                text-align: center;
            }

            .dot {
                position: relative;
                width: 15px;
                height: 15px;
                margin: 0 2px;
                display: inline-block;
            }

            .dot:first-child:before {
                animation-delay: 0ms;
            }

            .dot:first-child:after {
                animation-delay: 0ms;
            }

            .dot:last-child:before {
                animation-delay: 200ms;
            }

            .dot:last-child:after {
                animation-delay: 200ms;
            }

            .dot:before {
                content: "";
                position: absolute;
                left: 0;
                width: 15px;
                height: 15px;
                background-color: blue;
                animation-name: dotHover;
                animation-duration: 900ms;
                animation-timing-function: cubic-bezier(.82, 0, .26, 1);
                animation-iteration-count: infinite;
                animation-delay: 100ms;
                background: white;
                border-radius: 100%;
            }

            .dot:after {
                content: "";
                position: absolute;
                z-index: -1;
                background: black;
                box-shadow: 0px 0px 1px black;
                opacity: .20;
                width: 100%;
                height: 3px;
                left: 0;
                bottom: -2px;
                border-radius: 100%;
                animation-name: dotShadow;
                animation-duration: 900ms;
                animation-timing-function: cubic-bezier(.82, 0, .26, 1);
                animation-iteration-count: infinite;
                animation-delay: 100ms;
            }

            @keyframes dotShadow {
                0% {
                    transform: scaleX(1);
                }

                50% {
                    opacity: 0;
                    transform: scaleX(.6);
                }

                100% {
                    transform: scaleX(1);
                }
            }

            @keyframes dotHover {
                0% {
                    top: 0px;
                }

                50% {
                    top: -50px;
                    transform: scale(1.1);
                }

                100% {
                    top: 0;
                }
            }
        </style>
        <div class="loadingscreen">
            <div class="screen">

                <div class="loader">
                    <div class="dot"></div>
                    <div class="dot"></div>
                    <div class="dot"></div>
                </div>

            </div>
        </div>
<?php
    }
}
