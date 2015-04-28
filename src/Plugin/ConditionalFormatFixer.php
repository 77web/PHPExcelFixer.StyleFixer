<?php

namespace PHPExcel\StyleFixer\Plugin;

use PHPExcel\StyleFixer\Util\Book as BookUtil;
use PHPExcel\StyleFixer\Util\XmlNamespace;

/**
 * Class ConditionalFormatFixer
 * PHPExcelが壊した条件付き書式を修復する
 */
class ConditionalFormatFixer implements Plugin
{
    /**
     * @var \PHPExcel\StyleFixer\Util\Book
     */
    private $bookUtil;

    public function __construct(BookUtil $bookUtil)
    {
        $this->bookUtil = $bookUtil;
    }

    /**
     * 各ワークシート内の条件付き書式設定を修復する
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

        foreach ($distSheetMap as $sheetName => $distSheetPath) {
            $distXml = $output->getFromName($distSheetPath);
            if (false !== strpos($distXml, 'conditionalFormatting')) {
                $distDom = new \DOMDocument;
                $distDom->loadXML($distXml);
                $distXPath = new \DOMXPath($distDom);
                $distXPath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);
                $distRoot = $distXPath->query('//s:worksheet')->item(0);
                $distConditionalFormattings = $distXPath->query('//s:worksheet/s:conditionalFormatting');
                $elementAfterConditionalFormatting = null;
                foreach ($distConditionalFormattings as $distConditionalFormatting) {
                    $elementAfterConditionalFormatting = $distConditionalFormatting->nextSibling;
                    $distRoot->removeChild($distConditionalFormatting);
                }

                while ($elementAfterConditionalFormatting instanceOf \DOMNode && !$elementAfterConditionalFormatting instanceOf \DOMElement) {
                    $elementAfterConditionalFormatting = $elementAfterConditionalFormatting->nextSibling;
                }
                if (!$elementAfterConditionalFormatting instanceOf \DOMElement) {
                    break;
                }

                $srcSheetPath = $srcSheetMap[$sheetName];
                $srcXml = $template->getFromName($srcSheetPath);
                $srcDom = new \DOMDocument;
                $srcDom->loadXML($srcXml);
                $srcXPath = new \DOMXPath($srcDom);
                $srcXPath->registerNamespace('s', XmlNamespace::SPREADSHEETML_NS_URL);
                
                $conditionalFormattings = $srcXPath->query('//s:worksheet/s:conditionalFormatting');
                foreach ($conditionalFormattings as $conditionalFormatting) {
                    /** @var \DOMElement $conditionalFormatting */
                    $newDistConditionalFormatting = $distDom->importNode($conditionalFormatting, true);
                    $distRoot->insertBefore($newDistConditionalFormatting, $elementAfterConditionalFormatting);
                }

                $output->addFromString($distSheetPath, $distDom->saveXML());
            }
        }
    }
}
