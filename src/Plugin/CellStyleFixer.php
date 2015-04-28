<?php


namespace PHPExcelFixer\StyleFixer\Plugin;


use PHPExcelFixer\StyleFixer\Util\Book as BookUtil;
use PHPExcelFixer\StyleFixer\Util\Sheet as SheetUtil;
use PHPExcelFixer\StyleFixer\Util\XmlNamespace;

class CellStyleFixer implements Plugin
{
    /**
     * @var \PHPExcelFixer\StyleFixer\Util\Book
     */
    private $bookUtil;

    /**
     * @var \PHPExcelFixer\StyleFixer\Util\Sheet
     */
    private $sheetUtil;

    /**
     * @param BookUtil $bookUtil
     * @param SheetUtil $sheetUtil
     */
    public function __construct(BookUtil $bookUtil, SheetUtil $sheetUtil)
    {
        $this->bookUtil = $bookUtil;
        $this->sheetUtil = $sheetUtil;
    }

    /**
     * 各ワークシート内の各セルに付属している書式設定を修復する
     *
     * @param \ZipArchive $output
     * @param \ZipArchive $template
     */
    public function execute(\ZipArchive $output, \ZipArchive $template)
    {
        // テンプレート側のシート名のマッピング
        $srcSheetMap = $this->bookUtil->makeSheetMap($template);
        // 出力ファイル側のシート名のマッピング
        $distSheetMap = $this->bookUtil->makeSheetMap($output);
        // 出力ファイル側のシートの印刷範囲マッピング
        $distPrintAreas = $this->bookUtil->makePrintAreaMap($output);

        foreach ($distSheetMap as $sheetName => $distSheetPath) {
            $distXml = $output->getFromName($distSheetPath);
            $distDom = new \DOMDocument;
            $distDom->loadXML($distXml);
            $distXPath = new \DOMXPath($distDom);
            $distXPath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

            $srcXml = $template->getFromName($srcSheetMap[$sheetName]);
            $srcDom = new \DOMDocument;
            $srcDom->loadXML($srcXml);
            $srcXPath = new \DOMXPath($srcDom);
            $srcXPath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);

            $printArea = null;
            if (isset($distPrintAreas[$sheetName])) {
                list(,$printArea) = explode('!', $distPrintAreas[$sheetName]);
            }

            // セル番地 => スタイル番号
            $styleMap = [];
            foreach ($srcXPath->query('//s:worksheet/s:sheetData/s:row') as $srcRow) {
                /** @var \DOMElement $srcRow */
                if ($srcRow->hasAttribute('s')) {
                    /** @var \DOMElement $distRow */
                    $distRow = $distXPath->query('//s:worksheet/s:sheetData/s:row[@r="'.$srcRow->getAttribute('r').'"]')->item(0);
                    if ($distRow) {
                        $distRow->setAttribute('s', $srcRow->getAttribute('s'));
                    }
                }

                if ($srcRow->hasChildNodes()) {
                    foreach ($srcRow->childNodes as $srcCell) {
                        if ($srcCell instanceOf \DOMElement && $srcCell->tagName == 'c' && $srcCell->hasAttribute('s')) {
                            $styleMap[$srcCell->getAttribute('r')] = $srcCell->getAttribute('s');
                        }
                    }
                }
            }

            foreach ($distXPath->query('//s:worksheet/s:sheetData/s:row/s:c') as $distCell) {
                /** @var \DOMElement $distCell */
                $coordinate = $distCell->getAttribute('r');
                if (isset($styleMap[$coordinate]) && ($this->cellHasValue($distCell) || !$printArea || $this->sheetUtil->inRange($coordinate, $printArea))) {
                    $distCell->setAttribute('s', $styleMap[$coordinate]);
                } else {
                    $distCell->removeAttribute('s');
                }
            }

            $output->addFromString($distSheetPath, $distDom->saveXML());
        }
    }

    /**
     * @param \DOMElement $cell
     * @return bool
     */
    private function cellHasValue(\DOMElement $cell)
    {
        $hasValue = false;

        if ($cell->hasChildNodes()) {
            foreach ($cell->childNodes as $child) {
                if ($child instanceOf \DOMElement && $child->tagName == 'v') {
                    $hasValue = true;
                    break;
                }
            }
        }

        return $hasValue;
    }
}
