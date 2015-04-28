<?php


namespace PHPExcel\StyleFixer\Plugin;

use PHPExcel\StyleFixer\Util\Sheet;

class CellStyleFixerTest extends BasePluginTest
{
    public function test_execute()
    {
        $output = $this->getMock('\ZipArchive');
        $template = $this->getMock('\ZipArchive');
        $bookUtil = $this->createBookUtil(2, ['Sheet1' => 'xl/worksheets/sheet1.xml', 'Sheet2' => 'xl/worksheets/sheet2.xml']);
        $bookUtil
            ->expects($this->once())
            ->method('makePrintAreaMap')
            ->with($this->isInstanceOf('\ZipArchive'))
            ->will($this->returnValue(['Sheet1' => 'Sheet1!$A$18:$A$18', 'Sheet2' => 'Sheet2!$A$18:$A$18']))
        ;
        $sheetUtil = new Sheet();

        $output
            ->expects($this->exactly(2))
            ->method('getFromName')
            ->with($this->logicalOr('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'))
            ->will($this->returnCallback([$this, 'getOutputXml']))
        ;
        $template
            ->expects($this->exactly(2))
            ->method('getFromName')
            ->with($this->logicalOr('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'))
            ->will($this->returnCallback([$this, 'getTemplateXml']))
        ;
        $output
            ->expects($this->exactly(2))
            ->method('addFromString')
            ->with($this->logicalOr('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'), $this->callback([$this, 'assertXmlHasExpectedContent']))
        ;

        $fixer = new CellStyleFixer($bookUtil, $sheetUtil);
        $fixer->execute($output, $template);
    }

    public function assertXmlHasExpectedContent($xml)
    {
        // 行スタイルが修復されている
        $this->assertContains('<row r="18" s="99">', $xml);
        $this->assertNotContains('<row r="18" s="11">', $xml);
        // セルスタイルが修復されている
        $this->assertContains('<c r="A18" s="33"', $xml);
        $this->assertNotContains('<c r="A18" s="11"', $xml);
        if (strpos($xml, 'B18')) {
            // 印刷範囲外でも値があるセルは修復する
            $this->assertContains('<c r="B18" s="33"', $xml);
            $this->assertNotContains('<c r="B18" s="11"', $xml);
        }
        if (strpos($xml, 'C18')) {
            // 印刷範囲外で値もないセルは修復せずスタイル情報を破棄
            $this->assertContains('<c r="C18"/>', $xml);
            $this->assertNotContains('<c r="C18" s="33"', $xml);
            $this->assertNotContains('<c r="C18" s="11"', $xml);
        }

        return true;
    }
}
