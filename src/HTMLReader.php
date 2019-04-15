<?php

namespace NS\TableExportBundle;

use \PHPExcel_Reader_Exception;
use \DOMDocument;
use \PHPExcel;
use \PHPExcel_Reader_HTML;

class HTMLReader extends PHPExcel_Reader_HTML
{
    public function loadFromString($xmlString, PHPExcel $objPHPExcel)
    {
        while ($objPHPExcel->getSheetCount() <= $this->sheetIndex) {
            $objPHPExcel->createSheet();
        }

        $objPHPExcel->setActiveSheetIndex($this->sheetIndex);

        $dom = new DOMDocument();

        $loaded = $dom->loadHTML($this->securityScan($xmlString));
        if ($loaded === FALSE) {
            throw new PHPExcel_Reader_Exception('Failed to load input as a DOM Document');
        }

        $dom->preserveWhiteSpace = false;

        $row = 0;
        $column = 'A';
        $content = '';
        $this->processDomElement($dom, $objPHPExcel->getActiveSheet(), $row, $column, $content);

        return $objPHPExcel;
    }
}
