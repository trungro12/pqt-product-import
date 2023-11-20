<?php
class PQTProductImport{
    public static function init(){

        if(!is_admin()) return;
        require_once(PQT_PRODUCT__PLUGIN_DIR . '/lib/menu/init.php');
    }
    
    public static function pluginActivation(){

    }
    public static function pluginDeactivation(){
        
    }
}