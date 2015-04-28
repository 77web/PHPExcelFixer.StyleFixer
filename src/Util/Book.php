<?php


namespace PHPExcelFixer\StyleFixer\Util;


class Book
{
    /**
     * 人間可読なシート名 => シートのリレーションファイルのパス の連想配列を作る
     *
     * @param \ZipArchive $zip
     * @return array
     */
    public function makeSheetRelationMap(\ZipArchive $zip)
    {
        $sheetMap = $this->makeSheetMap($zip);

        array_walk($sheetMap, function(&$xmlPath){
            $xmlPath = str_replace('worksheets/', 'worksheets/_rels/', $xmlPath).'.rels';
        });

        return $sheetMap;
    }

    /**
     * 人間可読なシート名 => シートファイルのパス の連想配列を作る
     *
     * @param \ZipArchive $zip
     * @return array
     */
    public function makeSheetMap(\ZipArchive $zip)
    {
        $relXml = $zip->getFromName('xl/_rels/workbook.xml.rels');

        $relDom = new \DOMDocument;
        $relDom->loadXml($relXml);
        $relXPath = new \DOMXPath($relDom);
        $relXPath->registerNamespace('r', XmlNamespace::RELATIONSHIPS_NS_URL);

        $sheetRelationIdMaps = $this->makeSheetRelationIdMap($zip);

        $map = [];
        foreach ($sheetRelationIdMaps as $sheetName => $relId) {
            /** @var null|\DOMElement $relEntry */
            $relEntry = $relXPath->query('//r:Relationships/r:Relationship[@Id="'.$relId.'"]')->item(0);
            if ($relEntry) {
                $map[$sheetName] = 'xl/'.$relEntry->getAttribute('Target');
            }
        }

        return $map;
    }

    /**
     * 人間可読なシート名 => relationship id の連想配列を作る
     * @param \ZipArchive $zip
     * @return array
     */
    private function makeSheetRelationIdMap(\ZipArchive $zip)
    {
        $xml = $zip->getFromName('xl/workbook.xml');

        $dom = new \DOMDocument;
        $dom->loadXML($xml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

        $map = [];
        foreach ($xpath->query('//s:workbook/s:sheets/s:sheet') as $sheet) {
            /** @var \DOMElement $sheet */
            $map[$sheet->getAttribute('name')] = $sheet->getAttribute('r:id');
        }

        return $map;
    }

    /**
     * 人間可読なシート名 => 印刷範囲 の連想配列を作る
     *
     * @param \ZipArchive $zip
     * @return array
     */
    public function makePrintAreaMap(\ZipArchive $zip)
    {
        $xml = $zip->getFromName('xl/workbook.xml');

        $dom = new \DOMDocument;
        $dom->loadXML($xml);
        $xpath = new \DOMXPath($dom);
        $xpath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

        $sheets = $xpath->query('//s:workbook/s:sheets/s:sheet');
        $printAreas = $xpath->query('//s:workbook/s:definedNames/s:definedName[@name="_xlnm.Print_Area"]');
        $map = [];
        $localSheetId = 0;
        foreach ($sheets as $sheet) {
            /** @var \DOMElement $definedName */
            $definedName = $printAreas->item($localSheetId);
            if ($definedName) {
                /** @var \DOMElement $sheet */
                $map[$sheet->getAttribute('name')] = $definedName->nodeValue;
            }
            $localSheetId++;
        }

        return $map;
    }

    /**
     * 人間可読なシート名 => drawing.xmlファイルの配列　の連想配列を作る
     *
     * @param \ZipArchive $zip
     * @return array
     */
    public function makeDrawingMap(\ZipArchive $zip)
    {
        $map = [];

        $sheetRelationMap = $this->makeSheetRelationMap($zip);
        foreach ($sheetRelationMap as $sheetName => $sheetRelFile) {
            $drawings = [];

            $relXml = $zip->getFromName($sheetRelFile);
            $dom = new \DOMDocument();
            $dom->loadXML($relXml);
            $xpath = new \DOMXPath($dom);
            $xpath->registerNamespace('r', XmlNamespace::RELATIONSHIPS_NS_URL);

            $drawingFiles = $xpath->query('//r:Relationships/r:Relationship[@Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"]');
            foreach ($drawingFiles as $drawingFile) {
                /** @var \DOMElement $drawingFile */
                $drawings[] = str_replace('../', 'xl/', $drawingFile->getAttribute('Target'));
            }

            $xpath = null;
            $dom = null;

            if (0 !== count($drawings)) {
                $map[$sheetName] = $drawings;
            }
        }

        return $map;
    }
}
