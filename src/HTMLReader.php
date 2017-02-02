<?php
/**
 * Created by PhpStorm.
 * User: gnat
 * Date: 13/04/16
 * Time: 4:59 PM
 */

namespace NS\TableExportBundle;

use \PHPExcel_Reader_Exception;
use \DOMDocument;
use \PHPExcel;

class HTMLReader extends \PHPExcel_Reader_HTML
{
    public function loadFromString($xmlString, PHPExcel $objPHPExcel)
    {
        while ($objPHPExcel->getSheetCount() <= $this->_sheetIndex) {
            $objPHPExcel->createSheet();
        }

        $objPHPExcel->setActiveSheetIndex($this->_sheetIndex);

        $dom = new DOMDocument();

        $loaded = $dom->loadHTML($this->securityScan($xmlString));
        if ($loaded === FALSE) {
            throw new PHPExcel_Reader_Exception('Failed to load input as a DOM Document');
        }

        $dom->preserveWhiteSpace = false;

        $row = 0;
        $column = 'A';
        $content = '';
        $this->_processDomElement($dom, $objPHPExcel->getActiveSheet(), $row, $column, $content);

        return $objPHPExcel;
    }
}
